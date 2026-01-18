VERSION 5.00
Begin VB.Form OPESTADI 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Estadística de Movimientos"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   105
      Top             =   105
   End
   Begin VB.CommandButton Command1 
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
      Left            =   4305
      Picture         =   "OPESTADI.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1050
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080FF80&
      Caption         =   "Selección de Presentación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   210
      TabIndex        =   10
      Top             =   3465
      Width           =   3480
      Begin VB.OptionButton Option6 
         Caption         =   "Option1"
         Height          =   225
         Left            =   525
         TabIndex        =   12
         Top             =   840
         Width           =   225
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Option2"
         Height          =   195
         Left            =   525
         TabIndex        =   11
         Top             =   420
         Width           =   225
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Importes"
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
         TabIndex        =   14
         Top             =   840
         Width           =   2430
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidades "
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
         TabIndex        =   13
         Top             =   420
         Width           =   2010
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Opción de Trabajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   210
      TabIndex        =   2
      Top             =   1050
      Width           =   3480
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   225
         Left            =   525
         TabIndex        =   8
         Top             =   420
         Width           =   225
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   195
         Left            =   525
         TabIndex        =   4
         Top             =   840
         Width           =   225
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option1"
         Height          =   225
         Left            =   525
         TabIndex        =   3
         Top             =   1260
         Width           =   225
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Listado por Impresora"
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
         TabIndex        =   9
         Top             =   1260
         Width           =   2430
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Generar Archivo"
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
         TabIndex        =   6
         Top             =   420
         Width           =   2010
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Consulta por Pantalla"
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
         TabIndex        =   5
         Top             =   840
         Width           =   2430
      End
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
      Left            =   1575
      TabIndex        =   1
      Text            =   " "
      Top             =   315
      Width           =   3585
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
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
      Left            =   4305
      Picture         =   "OPESTADI.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3990
      Width           =   855
   End
   Begin VB.Label Desde 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Corresponde a:"
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
      Left            =   0
      TabIndex        =   7
      Top             =   420
      Width           =   1485
   End
End
Attribute VB_Name = "OPESTADI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    OPESTADI.Hide
End Sub

Private Sub Command2_Click()
    If OPESTADI.Option1.Value = True Then
        ' OPESTADI.Hide
        Call CONECRUN("#ESTASTO", "")
        Exit Sub
      ElseIf OPESTADI.Option2.Value = True Then
        If OPESTADI.Option5.Value = True Then
            Call CONECRUN("*ESTASTO", "")
          Else
            Call CONECRUN("*ESTAVA", "")
        End If
        Exit Sub
      ElseIf OPESTADI.Option3.Value = True Then
        If OPESTADI.Option5.Value = True Then
            Call CONECRUN("*LIESTOK", "")
          Else
            Call CONECRUN("*LIESVAL", "")
        End If
        Exit Sub
    End If
End Sub

Private Sub Form_Activate()
    Text1.Text = CONTROL$("ESTASTO", "ESTASTO")
    Option2.Value = True
    Option5.Value = True
    If TRIM$(Text1.Text) = "" Then
       Option1.Value = True
       Option5.Value = False
    End If
    '
End Sub

Private Sub Timer1_Timer()
     TTXS$ = CONTROL$("ESTASTO", "ESTASTO")
     If TTXS$ <> Text1.Text Then
        Text1.Text = TTXS$
        OPESTADI.Refresh
     End If
End Sub

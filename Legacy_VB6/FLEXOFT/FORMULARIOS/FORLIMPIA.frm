VERSION 5.00
Begin VB.Form FORLIMPIA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1065
      Left            =   1155
      TabIndex        =   8
      Top             =   3570
      Width           =   4425
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   3675
         Picture         =   "FORLIMPIA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   210
         Width           =   645
      End
      Begin VB.TextBox Text4 
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
         Left            =   1995
         TabIndex        =   13
         Top             =   630
         Width           =   1590
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Marcar Contacto"
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
         TabIndex        =   11
         Top             =   525
         Width           =   3270
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Borrar Mail de Rechazo"
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
         TabIndex        =   10
         Top             =   210
         Width           =   3270
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   15
         Left            =   2835
         TabIndex        =   15
         Top             =   735
         Width           =   15
      End
      Begin VB.Label Label5 
         Caption         =   "con Motivo ......"
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
         Left            =   525
         TabIndex        =   12
         Top             =   735
         Width           =   1380
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Siguiente Mail Rechazo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   5775
      TabIndex        =   4
      Top             =   3675
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Finalizar Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   5775
      TabIndex        =   3
      Top             =   315
      Width           =   1065
   End
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   1155
      TabIndex        =   2
      Top             =   3150
      Width           =   4425
   End
   Begin VB.TextBox Text2 
      Height          =   2115
      Left            =   1155
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "FORLIMPIA.frx":030A
      Top             =   840
      Width           =   4425
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1155
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   315
      Width           =   4425
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   210
      TabIndex        =   16
      Top             =   1260
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Acciones:"
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
      Left            =   -315
      TabIndex        =   9
      Top             =   3780
      Width           =   1380
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Dir-El-Mai:"
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
      Left            =   -315
      TabIndex        =   7
      Top             =   3255
      Width           =   1380
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Mensaje:"
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
      Left            =   -315
      TabIndex        =   6
      Top             =   860
      Width           =   1380
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Origen:"
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
      Left            =   -315
      TabIndex        =   5
      Top             =   370
      Width           =   1380
   End
End
Attribute VB_Name = "FORLIMPIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Command1_Click()
  '
  If TRIM$(Text3) = "" Then
    If Text2.SelLength > 0 Then
       Text3 = Mid$(Text2, 1 + Text2.SelStart, Text2.SelLength)
    End If
  End If
  ATENTI% = 0                          ' No borra Mail de rechazo
  If Check1.Value = 1 Then ATENTI% = 1 ' borra Mail de Rechazo
  '
  If Check2.Value = 1 Then             ' MARCA EN LISTA DE ENVIOS
    MOTIBA$ = TRIM$(Text4)
    If MOTIBA$ = "" Then
      Call MENSERR(24, "Falta Motivo de Baja")
      Exit Sub
    End If
    '
    MARCABA% = 0
    HOII% = HOY(HOS$)
    If TRIM$(Text3) <> "" Then
      Screen.MousePointer = 11
      For U& = 1 To ULTREG&("LIDIRMAI")
        XXX$ = REGLEIDO$("LIDIRMAI", U&)
        XX1$ = TRIM$(LCase$(Mid$(XXX$, 105, 60)))
        If XX1$ = TRIM$(LCase$(Text3)) Then
          Call REPLA(XXX$, MKI$(HOII%), 255, 2)
          Call REPLA(XXX$, MOTIBA$, 257, 8)
          Call GRAREG("LIDIRMAI", XXX$, U&)
          MARCABA% = 1
        End If
      Next U&
      Screen.MousePointer = 1
      If MARCABA% = 0 Then
        Call COMUNI("No encontrado en Bases de Datos")
      End If
    End If
  End If
  Text3 = ""
  On Error Resume Next
  Hide
  On Error GoTo 0
  '
End Sub

Private Sub Command2_Click()
   ATENTI% = 2
   Hide
End Sub

Private Sub Command3_Click()
  ATENTI% = 1
  Text3 = ""
  Hide
End Sub

Private Sub Command4_Click()
  ATENTI% = 0
  Text3 = ""
  Hide
End Sub

Private Sub Form_Activate()
   Text2.SetFocus
End Sub


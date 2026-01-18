VERSION 5.00
Begin VB.Form CLASIMAS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0FFD0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clasificacion Maestro de Items"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6075
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   5020
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   18
      Top             =   2940
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   19
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.CommandButton command1 
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
      Left            =   5020
      Picture         =   "CLASIMAS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Cierra - Abandona la Aplicación"
      Top             =   340
      Width           =   855
   End
   Begin VB.CommandButton Command11 
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
      Left            =   5020
      Picture         =   "CLASIMAS.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1260
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0FFD0&
      Caption         =   "Clasificación del Maestro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3610
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4635
      Begin VB.Frame Frame6 
         BackColor       =   &H00D0FFD0&
         Caption         =   "Criterios y Tablas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1420
         Left            =   315
         TabIndex        =   10
         Top             =   525
         Width           =   4005
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0FFD0&
            Caption         =   "Grupos Contables"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   315
            TabIndex        =   15
            Top             =   1065
            Width           =   2640
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0FFD0&
            Caption         =   "Tablas de Clasificación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   315
            TabIndex        =   14
            Top             =   815
            Width           =   2745
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0FFD0&
            Caption         =   "Criterios Técnicos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   315
            TabIndex        =   13
            Top             =   565
            Width           =   2640
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Abrir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   700
            Left            =   3045
            Picture         =   "CLASIMAS.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   525
            Width           =   750
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0FFD0&
            Caption         =   "Tabla de Tipos de Material"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   315
            TabIndex        =   11
            Top             =   315
            Width           =   2640
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00D0FFD0&
         Caption         =   "Operación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   315
         TabIndex        =   7
         Top             =   2835
         Width           =   4005
         Begin VB.OptionButton Option2 
            BackColor       =   &H00D0FFD0&
            Caption         =   "Por Rango"
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
            Left            =   2415
            TabIndex        =   9
            Top             =   250
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00D0FFD0&
            Caption         =   "Interactiva"
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
            Left            =   525
            TabIndex        =   8
            Top             =   250
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   1845
         TabIndex        =   4
         Top             =   2100
         Width           =   1905
         Begin VB.CommandButton Command7 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "CLASIMAS.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Clasificación del Maestro por Criterios Técnicos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Clasificación Técnica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   630
            TabIndex        =   6
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture10 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2415
         ScaleHeight     =   585
         ScaleWidth      =   1845
         TabIndex        =   1
         Top             =   2100
         Width           =   1905
         Begin VB.CommandButton Command14 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "CLASIMAS.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Clasificación por Grupos Contables"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Clasificación Contable"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   630
            TabIndex        =   3
            Top             =   105
            Width           =   1590
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4480
      Picture         =   "CLASIMAS.frx":0D72
      Top             =   3300
      Width           =   2010
   End
End
Attribute VB_Name = "CLASIMAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call OPNOIN("Clasificación Contable")
   Command14.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   If Option1(0).Value = True Then
        Call CLASINTER
      ElseIf Option2.Value = True Then
        Call CLASIRANGO
   End If
   Command7.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    If Option1(1) = True Then
        Call CARTATIMA
      ElseIf Option1(5).Value = True Then
        Call CARGRUCOVE
      ElseIf Option1(6).Value = True Then
        Call CARTACLAS
      ElseIf Option1(7).Value = True Then
        Call CARCRITEC
      Else
        Call OPNOIN("")
    End If
    Command6.Enabled = True
End Sub

Private Sub Option1_DblClick(Index As Integer)
    '
    If Index = 1 Then
        Call CARTATIMA
     ElseIf Index = 5 Then
        Call CARGRUCOVE
     ElseIf Index = 6 Then
        Call CARTACLAS
     ElseIf Index = 7 Then
        Call CARCRITEC
    End If
    '
End Sub


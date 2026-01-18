VERSION 5.00
Begin VB.Form FORDOCIN03 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresion de Documentos de Ingenieria"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10710
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   2415
      Left            =   9825
      ScaleHeight     =   2355
      ScaleWidth      =   1035
      TabIndex        =   46
      Top             =   0
      Width           =   1095
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
         Height          =   690
         Left            =   120
         Picture         =   "FORDOCIN03.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   240
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Left            =   120
         Picture         =   "FORDOCIN03.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   930
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -300
         Picture         =   "FORDOCIN03.frx":0454
         Top             =   1920
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Archivo Imagen Isologo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   4410
      TabIndex        =   43
      Top             =   210
      Width           =   5055
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2310
         MaxLength       =   20
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   420
         Width           =   2550
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del Archivo:"
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
         Left            =   -1260
         TabIndex        =   45
         Top             =   525
         Width           =   3480
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Formularios de Impresion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      Left            =   210
      TabIndex        =   6
      Top             =   2310
      Width           =   10350
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Index           =   5
         Left            =   5250
         TabIndex        =   37
         Top             =   2205
         Width           =   4950
         Begin VB.CommandButton BOTOPENFORM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   650
            Index           =   5
            Left            =   4200
            Picture         =   "FORDOCIN03.frx":2376
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   105
            Width           =   540
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Height          =   285
            Index           =   5
            Left            =   3885
            TabIndex        =   39
            Top             =   105
            Width           =   175
         End
         Begin VB.TextBox FORMUNAME 
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
            Left            =   2625
            TabIndex        =   38
            Text            =   " "
            Top             =   105
            Width           =   1275
         End
         Begin VB.Label FORMUECO 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   210
            TabIndex        =   42
            Top             =   430
            Width           =   3855
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Informe de No-Conformidad:"
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
            Index           =   5
            Left            =   105
            TabIndex        =   41
            Top             =   165
            Width           =   2430
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Index           =   4
         Left            =   5250
         TabIndex        =   31
         Top             =   1260
         Width           =   4950
         Begin VB.TextBox FORMUNAME 
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
            Index           =   4
            Left            =   2625
            TabIndex        =   34
            Text            =   " "
            Top             =   105
            Width           =   1275
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Height          =   285
            Index           =   4
            Left            =   3885
            TabIndex        =   33
            Top             =   105
            Width           =   175
         End
         Begin VB.CommandButton BOTOPENFORM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   650
            Index           =   4
            Left            =   4200
            Picture         =   "FORDOCIN03.frx":2680
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   105
            Width           =   540
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hoja de Actualización:"
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
            Index           =   4
            Left            =   210
            TabIndex        =   36
            Top             =   160
            Width           =   2325
         End
         Begin VB.Label FORMUECO 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Index           =   4
            Left            =   210
            TabIndex        =   35
            Top             =   430
            Width           =   3855
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Index           =   3
         Left            =   5250
         TabIndex        =   25
         Top             =   315
         Width           =   4950
         Begin VB.TextBox FORMUNAME 
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
            Index           =   3
            Left            =   2625
            TabIndex        =   28
            Text            =   " "
            Top             =   105
            Width           =   1275
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Height          =   285
            Index           =   3
            Left            =   3885
            TabIndex        =   27
            Top             =   105
            Width           =   175
         End
         Begin VB.CommandButton BOTOPENFORM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   650
            Index           =   3
            Left            =   4200
            Picture         =   "FORDOCIN03.frx":298A
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   105
            Width           =   540
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Plan de Inspeccion (PIE):"
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
            Index           =   3
            Left            =   210
            TabIndex        =   30
            Top             =   160
            Width           =   2325
         End
         Begin VB.Label FORMUECO 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Index           =   3
            Left            =   210
            TabIndex        =   29
            Top             =   430
            Width           =   3855
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Index           =   2
         Left            =   105
         TabIndex        =   19
         Top             =   2205
         Width           =   4950
         Begin VB.TextBox FORMUNAME 
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
            Index           =   2
            Left            =   2625
            TabIndex        =   22
            Text            =   " "
            Top             =   105
            Width           =   1275
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Height          =   285
            Index           =   2
            Left            =   3885
            TabIndex        =   21
            Top             =   105
            Width           =   175
         End
         Begin VB.CommandButton BOTOPENFORM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   650
            Index           =   2
            Left            =   4200
            Picture         =   "FORDOCIN03.frx":2C94
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   105
            Width           =   540
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Instrucc. / Procedimientos:"
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
            Index           =   2
            Left            =   105
            TabIndex        =   24
            Top             =   165
            Width           =   2430
         End
         Begin VB.Label FORMUECO 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Index           =   2
            Left            =   210
            TabIndex        =   23
            Top             =   430
            Width           =   3855
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Index           =   1
         Left            =   105
         TabIndex        =   13
         Top             =   1260
         Width           =   4950
         Begin VB.TextBox FORMUNAME 
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
            Left            =   2625
            TabIndex        =   16
            Text            =   " "
            Top             =   105
            Width           =   1275
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Height          =   285
            Index           =   1
            Left            =   3885
            TabIndex        =   15
            Top             =   105
            Width           =   175
         End
         Begin VB.CommandButton BOTOPENFORM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   650
            Index           =   1
            Left            =   4200
            Picture         =   "FORDOCIN03.frx":2F9E
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   105
            Width           =   540
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hoja de Especificaciones:"
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
            TabIndex        =   18
            Top             =   160
            Width           =   2325
         End
         Begin VB.Label FORMUECO 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   210
            TabIndex        =   17
            Top             =   430
            Width           =   3855
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   315
         Width           =   4950
         Begin VB.CommandButton BOTOPENFORM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   650
            Index           =   0
            Left            =   4200
            Picture         =   "FORDOCIN03.frx":32A8
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   105
            Width           =   540
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Height          =   285
            Index           =   0
            Left            =   3885
            TabIndex        =   9
            Top             =   105
            Width           =   175
         End
         Begin VB.TextBox FORMUNAME 
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
            Index           =   0
            Left            =   2625
            TabIndex        =   8
            Text            =   " "
            Top             =   105
            Width           =   1275
         End
         Begin VB.Label FORMUECO 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Index           =   0
            Left            =   210
            TabIndex        =   12
            Top             =   430
            Width           =   3855
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Plano de Pieza / Conjunto:"
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
            TabIndex        =   11
            Top             =   160
            Width           =   2325
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   5145
         X2              =   5145
         Y1              =   105
         Y2              =   3255
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ubicación de Archivos Gráficos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   3
      Top             =   1260
      Width           =   9255
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   735
         MaxLength       =   40
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   420
         Width           =   8325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta:"
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
         TabIndex        =   5
         Top             =   525
         Width           =   645
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Tipo de Archivo (JPG - GIF - etc)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4005
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
         Height          =   360
         Left            =   2835
         MaxLength       =   8
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   " *.*"
         Top             =   420
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código Tipo de Archivo:"
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
         Left            =   -735
         TabIndex        =   2
         Top             =   525
         Width           =   3480
      End
   End
   Begin VB.Line Line2 
      X1              =   210
      X2              =   10500
      Y1              =   3570
      Y2              =   3465
   End
End
Attribute VB_Name = "FORDOCIN03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTOPENFORM_Click(Index As Integer)
    Call CONECRUN("FLFORMS/" + TRIM$(FORMUNAME(Index).TEXT), "Formularios de Impresion")
    If EXISTE%("FLFORMS.EXE") < 1 Then
        Call FINAL("")
    End If
End Sub

Private Sub BOTSELFORM_Click(Index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        FORMUNAME(Index).TEXT = VDEVU$
        FORMUECO(Index).Caption = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    TIARGRA$ = TRIM$(Text1.TEXT)
    Call GRACONTROL("", "TIARGRA", Text1.TEXT)
    '
    RUTARGRA$ = TRIM$(Text2.TEXT)
    Call GRACONTROL("", "RUTARGRA", Text2.TEXT)
    '
    ISOLOGO$ = TRIM$(Text3.TEXT)
    Call GRACONTROL("", "ISOLOGO", Text3.TEXT)
    '
    For KK% = 0 To 5
       On Error Resume Next
       Call GRACONTROL("FORGRAFI", "DOCUGRA" + TRIM$(Str$(KK%)), FORMUNAME(KK%).TEXT)
    Next KK%
    '
    On Error GoTo 0
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Form_Load()
    '
    TIARGRA$ = TRIM$(CONTROL$("", "TIARGRA"))
    If TIARGRA$ <> "" Then
       Text1.TEXT = TIARGRA$
    End If
    '
    RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
    If RUTARGRA$ <> "" Then
       Text2.TEXT = RUTARGRA$
    End If
    '
    ISOLOGO$ = TRIM$(CONTROL$("", "ISOLOGO"))
    If ISOLOGO$ <> "" Then
       Text3.TEXT = ISOLOGO$
    End If
    '
    For KK% = 0 To 5
       On Error Resume Next
       FORMUNAME(KK%).TEXT = CONTROL$("FORGRAFI", "DOCUGRA" + TRIM$(Str$(KK%)))
       FORMUECO(KK%).Caption = ECOARCH$("INDIFRM", FORMUNAME(KK%).TEXT)
    Next KK%
    On Error GoTo 0
    '
End Sub

Private Sub FORMUNAME_DblClick(Index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        FORMUNAME(Index).TEXT = VDEVU$
        FORMUECO(Index).Caption = VALACT2$("INDIFRM")
    End If
End Sub

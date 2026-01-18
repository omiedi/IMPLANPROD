VERSION 5.00
Begin VB.Form FORMALTNU 
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Accesos Directos"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6510
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox SORTLIST 
      Height          =   450
      ItemData        =   "FORMALTNU.frx":0000
      Left            =   4830
      List            =   "FORMALTNU.frx":0007
      Sorted          =   -1  'True
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   2730
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.ListBox LISORT 
      Height          =   450
      ItemData        =   "FORMALTNU.frx":0015
      Left            =   4830
      List            =   "FORMALTNU.frx":001C
      TabIndex        =   71
      Top             =   2205
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   4830
      TabIndex        =   70
      Top             =   105
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   210
      TabIndex        =   69
      Top             =   3675
      Visible         =   0   'False
      Width           =   6105
   End
   Begin VB.ListBox LINOSORT 
      Height          =   450
      ItemData        =   "FORMALTNU.frx":0028
      Left            =   4830
      List            =   "FORMALTNU.frx":002F
      TabIndex        =   68
      Top             =   1680
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.ListBox PSPOOL 
      Height          =   1425
      Left            =   210
      TabIndex        =   67
      Top             =   3360
      Visible         =   0   'False
      Width           =   6105
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFE0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3060
      Left            =   1155
      TabIndex        =   33
      Top             =   105
      Width           =   3480
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   0
         Left            =   0
         TabIndex        =   0
         Top             =   210
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   1
         Left            =   0
         TabIndex        =   1
         Top             =   105
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   2
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   3
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   4
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   5
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   6
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   7
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   8
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   9
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   10
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   11
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   12
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   13
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   14
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   15
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   16
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   17
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   18
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   19
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   20
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   21
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   22
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   23
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   24
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   25
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   26
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   27
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   28
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   29
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   30
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CommandButton BOTALTER 
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
         Height          =   330
         Index           =   31
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   65
         Top             =   1050
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   64
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   0
         TabIndex        =   63
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   62
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   0
         TabIndex        =   61
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   0
         TabIndex        =   59
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   0
         TabIndex        =   58
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   0
         TabIndex        =   57
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   0
         TabIndex        =   54
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   0
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   0
         TabIndex        =   52
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   14
         Left            =   0
         TabIndex        =   51
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   0
         TabIndex        =   50
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   16
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   17
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   18
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   19
         Left            =   0
         TabIndex        =   46
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   20
         Left            =   0
         TabIndex        =   45
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   21
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   22
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   23
         Left            =   0
         TabIndex        =   42
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   24
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   25
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   26
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   27
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   28
         Left            =   0
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   29
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   30
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label LEYEN 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   31
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Visible         =   0   'False
         Width           =   75
      End
   End
   Begin VB.CommandButton Command1 
      Height          =   330
      Left            =   220
      Picture         =   "FORMALTNU.frx":003D
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Imprime Mensaje por Impresora"
      Top             =   2730
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label COMALT 
      Caption         =   " COMALT"
      Height          =   225
      Left            =   210
      TabIndex        =   73
      Top             =   5250
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   1065
      Left            =   210
      Stretch         =   -1  'True
      Top             =   1365
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   210
      Picture         =   "FORMALTNU.frx":013F
      Top             =   315
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   210
      Picture         =   "FORMALTNU.frx":0449
      Top             =   315
      Width           =   480
   End
End
Attribute VB_Name = "FORMALTNU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REMENSA$(32)
'
Private Sub BOTALTER_Click(Index As Integer)
   '
   If MODALTER% = 1 Then
     Hide
     Call FRESHALL
     If MODACDIR$(Index + 1) <> "" Then
         Call CONECRUN(MODACDIR$(Index + 1), "")
       Else
         Call FRESHALL
     End If
   ElseIf MODALTER% = 2 Then
     ALTERX% = Index% + 1
     On Error Resume Next
     FORMALTER.BOTALTER(0).SetFocus
     On Error GoTo 0
     Hide
     Call FRESHALL
   End If
   '
End Sub

Private Sub COMALT_Change()
   '
   If COMALT <> "" Then
     '
     TEXT$ = LTrim$(COMALT)
     POSI$ = ""
     TEXT$ = LTrim$(TEXT$)
     If Left$(TEXT$, 1) = "&" Then
       TEXT$ = Mid$(TEXT$, 2)
       POSK% = InStr(TEXT$, "&")
       If POSK% > 0 Then
         POSI$ = Left$(TEXT$, POSK% - 1)
         TEXT$ = Mid$(TEXT$, POSK% + 1)
       End If
     End If
     '
     RGQ% = Int(Val(POSI$))
     If RGQ% < 2 Or RGQ% > 23 Then RGQ% = 13
     '
     CLQ% = 100 * (Val(POSI$) - RGQ%)
     If CLQ% < 1 Or CLQ% > 80 Then CLQ% = 40
     '
     NOCIE% = 0: If InStr(POSI$, "/NC") > 0 Then NOCIE% = 1
     '
     Call SUPRACENTO(TEXT$)
     POSEP% = InStr(TEXT$, "\\")
     If POSEP% < 1 Then
       MsgBox "Opciones Indefinidas en Rutina 'COMALTER'"
       TIMENSA% = 0
       Hide
     End If
     '
     OPCI$ = Mid$(TEXT$, POSEP% + 2)
     TEXT$ = Left$(TEXT$, POSEP% - 1)
     '
     If Len(OPCI$) < 1 Then
       MsgBox "No hay Opciones en Rutina 'COMALTER'"
       TIMENSA% = 0
       Hide
     End If
     '
     Frame1.Top = 0
     Frame1.Left = 0
     For I% = 0 To 31
       BOTALTER(I%).Visible = False
       LEYEN(I%).Visible = False
     Next I%
     '
     TEXTX$ = TEXT$
     CADAT% = 0: AFORMU = 0
     BOSUP = 120
     For I% = 1 To 16
       PX% = InStr(TEXTX$, "\")
       If PX% < 1 Then PX% = 1 + Len(TEXTX$)
       Alter$ = Left$(TEXTX$, PX% - 1)
       TEXTX$ = Mid$(TEXTX$, PX% + 1)
       If Len(Alter$) > 0 Then
         CADAT% = CADAT% + 1
         LEYEN(CADAT% - 1) = "   " + Alter$ + "   "
         REMENSA$(CADAT%) = Alter$
         If LEYEN(CADAT% - 1).Width > AFORMU Then
           AFORMU = LEYEN(CADAT% - 1).Width
         End If
         LEYEN(CADAT% - 1).Top = BOSUP
         LEYEN(CADAT% - 1).Left = 0
         LEYEN(CADAT% - 1).Visible = True
         BOSUP = BOSUP + LEYEN(CADAT% - 1).Height
       End If
       If Len(TEXTX$) < 1 Then GoTo 105
    Next I%
    '
105 TEXTX$ = OPCI$
    REMENSA$(CADAT% + 1) = OPCI$
    CAOP% = 0
    ABOTO = 0
    For I% = 1 To 10
      PX% = InStr(TEXTX$, "\")
      If PX% < 1 Then PX% = 1 + Len(TEXTX$)
      OPX$ = Left$(TEXTX$, PX% - 1)
      TEXTX$ = Mid$(TEXTX$, PX% + 1)
      If Len(OPX$) > 0 Then
        CAOP% = CAOP% + 1
        BOTALTER(CAOP% - 1).Caption = OPX$
        TEPRUEBA.Caption = "  " + OPX$ + "  "
        BOTALTER(CAOP% - 1).Width = TEPRUEBA.Width
        ABOTO = ABOTO + BOTALTER(CAOP% - 1).Width
      End If
      If Len(TEXTX$) < 1 Then GoTo 107
106 Next I%
    '
107 If ABOTO + 150 * (CAOP% + 1) > AFORMU Then
      AFORMU = ABOTO + 150 * (CAOP% + 1)
    End If
    INTERBOT = (AFORMU - ABOTO) / (CAOP% + 1)
    '
    BOSUP = BOSUP + 240
    BIZQ = INTERBOT
    For I% = 1 To CAOP%
      BOTALTER(I% - 1).Left = BIZQ
      BOTALTER(I% - 1).Top = BOSUP
      BOTALTER(I% - 1).Visible = True
      BIZQ = BIZQ + BOTALTER(I% - 1).Width + INTERBOT
    Next I%
    '
    command1.Top = BOSUP
    '
    For I% = 1 To CADAT%
      While LEYEN(I% - 1).Width < AFORMU
        LEYEN(I% - 1).Caption = " " + LEYEN(I% - 1).Caption + " "
      Wend
    Next I%
    '
    ALTUTIT = Height - ScaleHeight
    Width = AFORMU
    Height = BOSUP + BOTALTER(0).Height + 120 + ALTUTIT
    Frame1.Width = Width
    Frame1.Height = ScaleHeight
    '
    Caption = "Opciones Alternativas"
    Image1.Visible = True
    Image2.Visible = False
    If CAOP% < 2 Or TIMENSA% = 1 Then
      Caption = "Información Crítica"
      Width = 900 + Width + command1.Width + 100
      Frame1.Left = 900
    End If
    If TIMENSA% = 2 Then
      Caption = "Operación Incorrecta"
      Image1.Visible = False
      Image2.Visible = True
    End If
    '
    Top = Screen.Height * RGQ% / 25 - Height / 2
    Left = Screen.Width * CLQ% / 80 - Width / 2
    MODALTER% = 2
    ALTERX% = 0
    Refresh
    SMP = Screen.MousePointer
    Screen.MousePointer = 1
Exit Sub
    TII% = HORANUM%(Time$)
On Error Resume Next
    'FORMALTER.Command1.Visible = True
    'FORMALTER.Show 1
On Error GoTo 0
    Screen.MousePointer = SMP
    Call FRESHALL
    'COMALTER% = ALTERX%
    '
    If VERCD% = 1 Then
      If TRIM$(LUDAT$) = "" Then
        'Exit Function
      End If
    End If
    '
    If UCase$(App.EXEName) = "FLEXOFT" Then
       If EXISTE%(LUDAT$ + "MENSAJES.RFS") < 1 Then
          'Exit Function
       End If
    End If
    '
    HOI% = HOY(HO$)
    TIF% = HORANUM%(Time$)
    TIGRA = Timer
    MNX$ = REGBLAN$("MENSAJES")
    Call REPLA(MNX$, MKS$(TIGRA), 1, 4)
    Call REPLA(MNX$, MKI$(HOI%), 5, 2)
    Call REPLA(MNX$, MKI$(TII%), 7, 2)
    Call REPLA(MNX$, MKI$(TIF%), 9, 2)
    Call REPLA(MNX$, MKI$(USNBR% Mod 100), 11, 2)
    Call REPLA(MNX$, NOMTER$, 13, 30)
    Call REPLA(MNX$, UCase$(App.EXEName), 43, 16)
    Call REPLA(MNX$, PROCON$, 59, 16)
    Call REPLA(MNX$, MKI$(ALTERX%), 157, 2)
    '
    For I% = 1 To CADAT%
       If TRIM$(REMENSA$(I%)) <> "" Then
          Call REPLA(MNX$, "." + AJUCEN$(REMENSA$(I%), 81), 75, 82)
          Call REGAPP("MENSAJES", MNX$)
          FORMALTER.LEYEN(I% - 1).Caption = ""
       End If
    Next I%
    '
    If CAOP% > 1 Then
       Call REPLA(MNX$, "." + AJUCEN$(REMENSA$(CADAT% + 1), 81), 75, 82)
       Call REGAPP("MENSAJES", MNX$)
    End If
    '
    Call CIERRARCH("MENSAJES")
    TIMENSA% = 0
    '
   
   End If
End Sub

Private Sub Command1_Click()
   PRINTFORM
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   FORMALTER.BOTALTER(0).SetFocus
   On Error GoTo 0
End Sub

Private Sub Form_DblClick()
   PRINTFORM
End Sub

Private Sub Form_Load()
   '
   Call TRANSLABELS(Name)
   '
End Sub

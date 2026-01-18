VERSION 5.00
Begin VB.Form XFORTAB 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00DDFFCC&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editor Tabulado"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11010
   FillStyle       =   0  'Solid
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
   ScaleHeight     =   7620
   ScaleWidth      =   11010
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox FRAMETOT 
      BackColor       =   &H00C0C0C0&
      Height          =   750
      Left            =   420
      ScaleHeight     =   690
      ScaleWidth      =   5835
      TabIndex        =   19
      Top             =   4200
      Width           =   5895
      Begin VB.Label TOTALI 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   20
         Top             =   105
         Width           =   120
      End
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   31
      Left            =   0
      TabIndex        =   183
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   30
      Left            =   0
      TabIndex        =   182
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   29
      Left            =   0
      TabIndex        =   181
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   28
      Left            =   0
      TabIndex        =   180
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   27
      Left            =   0
      TabIndex        =   179
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   26
      Left            =   0
      TabIndex        =   178
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   25
      Left            =   0
      TabIndex        =   177
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   24
      Left            =   0
      TabIndex        =   176
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   23
      Left            =   0
      TabIndex        =   175
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   22
      Left            =   0
      TabIndex        =   174
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   21
      Left            =   0
      TabIndex        =   173
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   20
      Left            =   0
      TabIndex        =   172
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   19
      Left            =   0
      TabIndex        =   171
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   18
      Left            =   0
      TabIndex        =   170
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   17
      Left            =   0
      TabIndex        =   169
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   16
      Left            =   0
      TabIndex        =   168
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   15
      Left            =   0
      TabIndex        =   167
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   14
      Left            =   0
      TabIndex        =   166
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   13
      Left            =   0
      TabIndex        =   165
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   12
      Left            =   0
      TabIndex        =   164
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   11
      Left            =   0
      TabIndex        =   163
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   10
      Left            =   0
      TabIndex        =   162
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   9
      Left            =   0
      TabIndex        =   161
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox FRAMENCA 
      BackColor       =   &H00C0C0C0&
      Height          =   540
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   5835
      TabIndex        =   17
      Top             =   840
      Width           =   5895
      Begin VB.Label ENCACOL 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   18
         Top             =   105
         Width           =   720
      End
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   8
      Left            =   0
      TabIndex        =   160
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   7
      Left            =   0
      TabIndex        =   159
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   6
      Left            =   0
      TabIndex        =   158
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   5
      Left            =   0
      TabIndex        =   157
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   4
      Left            =   0
      TabIndex        =   156
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   3
      Left            =   0
      TabIndex        =   155
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   2
      Left            =   0
      TabIndex        =   154
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   1
      Left            =   0
      TabIndex        =   153
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Frame LINVER 
      Caption         =   "Frame1"
      Height          =   4425
      Index           =   0
      Left            =   525
      TabIndex        =   152
      Top             =   735
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   9135
      ScaleHeight     =   75
      ScaleWidth      =   900
      TabIndex        =   150
      Top             =   3255
      Visible         =   0   'False
      Width           =   960
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   151
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.ListBox LITAB 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   2310
      Left            =   0
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   1785
      Width           =   7575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   4425
      Left            =   0
      ScaleHeight     =   4365
      ScaleWidth      =   8355
      TabIndex        =   21
      Top             =   0
      Width           =   8415
      Begin VB.TextBox FIECO 
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
         Index           =   9
         Left            =   0
         TabIndex        =   76
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   8
         Left            =   0
         TabIndex        =   77
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Left            =   0
         TabIndex        =   78
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Left            =   0
         TabIndex        =   79
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Left            =   0
         TabIndex        =   80
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Left            =   0
         TabIndex        =   81
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   31
         Left            =   0
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   30
         Left            =   0
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Left            =   0
         TabIndex        =   82
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   29
         Left            =   0
         TabIndex        =   56
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   28
         Left            =   0
         TabIndex        =   57
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   27
         Left            =   0
         TabIndex        =   58
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   26
         Left            =   0
         TabIndex        =   59
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   25
         Left            =   0
         TabIndex        =   60
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   24
         Left            =   0
         TabIndex        =   61
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   23
         Left            =   0
         TabIndex        =   62
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   22
         Left            =   0
         TabIndex        =   63
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   21
         Left            =   0
         TabIndex        =   64
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   20
         Left            =   0
         TabIndex        =   65
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Left            =   0
         TabIndex        =   83
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   19
         Left            =   0
         TabIndex        =   66
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   18
         Left            =   0
         TabIndex        =   67
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   17
         Left            =   0
         TabIndex        =   68
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   16
         Left            =   0
         TabIndex        =   69
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   15
         Left            =   0
         TabIndex        =   70
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   14
         Left            =   0
         TabIndex        =   71
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   13
         Left            =   0
         TabIndex        =   72
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   12
         Left            =   0
         TabIndex        =   73
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   11
         Left            =   0
         TabIndex        =   74
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Index           =   10
         Left            =   0
         TabIndex        =   75
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Left            =   0
         TabIndex        =   84
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   0
         TabIndex        =   116
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIECO 
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
         Left            =   5880
         TabIndex        =   85
         TabStop         =   0   'False
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
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
         Left            =   3465
         TabIndex        =   117
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   0
         TabIndex        =   115
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   0
         TabIndex        =   114
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   0
         TabIndex        =   113
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   0
         TabIndex        =   112
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   6
         Left            =   0
         TabIndex        =   111
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   7
         Left            =   0
         TabIndex        =   110
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   8
         Left            =   0
         TabIndex        =   109
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   9
         Left            =   0
         TabIndex        =   108
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   10
         Left            =   0
         TabIndex        =   107
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   11
         Left            =   0
         TabIndex        =   106
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   12
         Left            =   0
         TabIndex        =   105
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   13
         Left            =   0
         TabIndex        =   104
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   14
         Left            =   0
         TabIndex        =   103
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   15
         Left            =   0
         TabIndex        =   102
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   16
         Left            =   0
         TabIndex        =   101
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   17
         Left            =   0
         TabIndex        =   100
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   18
         Left            =   0
         TabIndex        =   99
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   19
         Left            =   0
         TabIndex        =   98
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   20
         Left            =   0
         TabIndex        =   97
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   21
         Left            =   0
         TabIndex        =   96
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   22
         Left            =   0
         TabIndex        =   95
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   23
         Left            =   0
         TabIndex        =   94
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   24
         Left            =   0
         TabIndex        =   93
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   25
         Left            =   0
         TabIndex        =   92
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   26
         Left            =   0
         TabIndex        =   91
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   27
         Left            =   0
         TabIndex        =   90
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   28
         Left            =   0
         TabIndex        =   89
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   29
         Left            =   0
         TabIndex        =   88
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   30
         Left            =   0
         TabIndex        =   87
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox FIENTRA 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   31
         Left            =   0
         TabIndex        =   86
         Text            =   " "
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   5355
         Picture         =   "XFORTAB.frx":0000
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   0
         Picture         =   "XFORTAB.frx":030A
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   0
         Picture         =   "XFORTAB.frx":0614
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   0
         Picture         =   "XFORTAB.frx":091E
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   0
         Picture         =   "XFORTAB.frx":0C28
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   0
         Picture         =   "XFORTAB.frx":0F32
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   0
         Picture         =   "XFORTAB.frx":123C
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   0
         Picture         =   "XFORTAB.frx":1546
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   0
         Picture         =   "XFORTAB.frx":1850
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   0
         Picture         =   "XFORTAB.frx":1B5A
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   0
         Picture         =   "XFORTAB.frx":1E64
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   0
         Picture         =   "XFORTAB.frx":216E
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         Left            =   0
         Picture         =   "XFORTAB.frx":2478
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   13
         Left            =   0
         Picture         =   "XFORTAB.frx":2782
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   14
         Left            =   0
         Picture         =   "XFORTAB.frx":2A8C
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   15
         Left            =   0
         Picture         =   "XFORTAB.frx":2D96
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   16
         Left            =   0
         Picture         =   "XFORTAB.frx":30A0
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   17
         Left            =   0
         Picture         =   "XFORTAB.frx":33AA
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   18
         Left            =   0
         Picture         =   "XFORTAB.frx":36B4
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   19
         Left            =   0
         Picture         =   "XFORTAB.frx":39BE
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   20
         Left            =   0
         Picture         =   "XFORTAB.frx":3CC8
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   21
         Left            =   0
         Picture         =   "XFORTAB.frx":3FD2
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   22
         Left            =   0
         Picture         =   "XFORTAB.frx":42DC
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   23
         Left            =   0
         Picture         =   "XFORTAB.frx":45E6
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   24
         Left            =   0
         Picture         =   "XFORTAB.frx":48F0
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   25
         Left            =   0
         Picture         =   "XFORTAB.frx":4BFA
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   26
         Left            =   0
         Picture         =   "XFORTAB.frx":4F04
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   27
         Left            =   0
         Picture         =   "XFORTAB.frx":520E
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   28
         Left            =   0
         Picture         =   "XFORTAB.frx":5518
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   29
         Left            =   0
         Picture         =   "XFORTAB.frx":5822
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   30
         Left            =   0
         Picture         =   "XFORTAB.frx":5B2C
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.CommandButton FIBOTSE 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   31
         Left            =   0
         Picture         =   "XFORTAB.frx":5E36
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label FILABEL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         Height          =   240
         Index           =   0
         Left            =   2415
         TabIndex        =   149
         Top             =   0
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   0
         TabIndex        =   148
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   0
         TabIndex        =   147
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   0
         TabIndex        =   146
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   0
         TabIndex        =   145
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   0
         TabIndex        =   144
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   0
         TabIndex        =   143
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   0
         TabIndex        =   142
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   0
         TabIndex        =   141
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   0
         TabIndex        =   140
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   0
         TabIndex        =   139
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   0
         TabIndex        =   138
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   12
         Left            =   0
         TabIndex        =   137
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   13
         Left            =   0
         TabIndex        =   136
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   14
         Left            =   0
         TabIndex        =   135
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   15
         Left            =   0
         TabIndex        =   134
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   16
         Left            =   0
         TabIndex        =   133
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   17
         Left            =   0
         TabIndex        =   132
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   18
         Left            =   0
         TabIndex        =   131
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   19
         Left            =   0
         TabIndex        =   130
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   20
         Left            =   0
         TabIndex        =   129
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   21
         Left            =   0
         TabIndex        =   128
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   22
         Left            =   0
         TabIndex        =   127
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   23
         Left            =   0
         TabIndex        =   126
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   24
         Left            =   0
         TabIndex        =   125
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   25
         Left            =   0
         TabIndex        =   124
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   26
         Left            =   0
         TabIndex        =   123
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   27
         Left            =   0
         TabIndex        =   122
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   28
         Left            =   0
         TabIndex        =   121
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   29
         Left            =   0
         TabIndex        =   120
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   30
         Left            =   0
         TabIndex        =   119
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label FILABEL 
         BackStyle       =   0  'Transparent
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   31
         Left            =   0
         TabIndex        =   118
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.ListBox LIREGSEL 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   4620
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5355
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2430
      Left            =   7350
      Max             =   1
      Min             =   1
      TabIndex        =   14
      Top             =   1785
      Value           =   1
      Width           =   240
   End
   Begin VB.Frame FRAMEBOT2 
      BackColor       =   &H00DDFFCC&
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2350
      Left            =   9135
      TabIndex        =   8
      Top             =   3465
      Width           =   960
      Begin VB.CommandButton BOTEDIT 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   105
         TabIndex        =   13
         ToolTipText     =   "Editar Registro para Actualización"
         Top             =   1470
         Width           =   700
      End
      Begin VB.CommandButton BOTAGRE 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   105
         TabIndex        =   12
         ToolTipText     =   "Agregar Registro al Final"
         Top             =   315
         Width           =   700
      End
      Begin VB.CommandButton BOTINSE 
         Caption         =   "Ins."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   105
         TabIndex        =   11
         ToolTipText     =   "Intercalar Registro Antes del Activo"
         Top             =   700
         Width           =   700
      End
      Begin VB.CommandButton BOTBUSCA 
         Caption         =   "Look"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   105
         TabIndex        =   10
         ToolTipText     =   "Buscar Texto por Campo"
         Top             =   1860
         Width           =   700
      End
      Begin VB.CommandButton BOTERASE 
         Caption         =   "Del."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   105
         TabIndex        =   9
         ToolTipText     =   "Borrar Registro Activo"
         Top             =   1085
         Width           =   700
      End
   End
   Begin VB.Frame FRAMEBOT1 
      BackColor       =   &H00DDFFCC&
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3060
      Left            =   9135
      TabIndex        =   3
      Top             =   105
      Width           =   960
      Begin VB.CommandButton BOTEXIT 
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
         Left            =   105
         Picture         =   "XFORTAB.frx":6140
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Salir de la Aplicación"
         Top             =   280
         Width           =   700
      End
      Begin VB.CommandButton BOTHELP 
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
         Left            =   105
         Picture         =   "XFORTAB.frx":628A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1610
         Width           =   700
      End
      Begin VB.CommandButton BOTPRINT 
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
         Left            =   105
         Picture         =   "XFORTAB.frx":6594
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprimir Formulario o Texto Seleccionado"
         Top             =   2275
         Width           =   700
      End
      Begin VB.CommandButton BOTREC 
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
         Left            =   105
         Picture         =   "XFORTAB.frx":6BFE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Grabar - Guardar Modificaciones"
         Top             =   945
         Width           =   700
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   105
      Top             =   6195
   End
   Begin VB.ListBox SORTLIST 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   735
      Sorted          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5355
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label1 
      Height          =   330
      Left            =   630
      TabIndex        =   184
      Top             =   6825
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   420
      X2              =   6300
      Y1              =   4950
      Y2              =   4950
   End
   Begin VB.Line Line2 
      X1              =   7665
      X2              =   7665
      Y1              =   1155
      Y2              =   5145
   End
   Begin VB.Image Image3 
      Height          =   510
      Left            =   3990
      Picture         =   "XFORTAB.frx":6F08
      Top             =   4725
      Width           =   2010
   End
   Begin VB.Label TEPRU 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   240
      Left            =   2520
      TabIndex        =   16
      Top             =   525
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3465
      TabIndex        =   1
      Top             =   6405
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "XFORTAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RGFJ%(32), TFF%(32)
Dim VFMIN$(32), VFMAX$(32)
'
Dim VDS$(32, 16), VDX#(32, 16), VD%(32, 16), VFS$(32), VFSA$(32), VFX#(32), VF%(32), FI$(32)
Dim TVAL%(20), CHE$(20), BAN$(20), IMPCHE#(20), FECHE%(20)
Dim VDP%(32), VDPX#(32), VDPS$(32), VDA$(32), VDAA%(32), VDAB#(32)
Dim SUMA#(32, 2), SUPAR#(32)
Dim LEYE$(16)
'
Dim CALIMI As Integer
Dim CALIMS As Integer
Dim COEFH As Single
'
Dim MIZQ$
'
Dim PPZZ As String * 128
Dim FUFIJA$(128)
'
'
Private Sub BOTAGRE_Click()
    Call ADDLISTA
End Sub

Private Sub BOTBUSCA_Click()
    Call BULISTA
End Sub

Private Sub BOTEDIT_Click()
    Call EDLISTA
End Sub

Private Sub BOTERASE_Click()
    Call BORLISTA
    MODIFI% = 1
End Sub

Private Sub BOTHELP_Click()
    Call HELPONLINE("WINTABU")
End Sub

Private Sub BOTINSE_Click()
    Call INSLISTA
End Sub

Private Sub BOTPRINT_Click()
   Enabled = False
   If LITAB.ListCount > 0 Then
      Call PRINTABU
   End If
   Enabled = True
End Sub

Private Sub BOTREC_Click()
    If GRABACT% = 1 Then
        Call GRABALI
    End If
    MODIFI% = 0
End Sub

Private Sub FIBOTSE_Click(Index As Integer)
    '
    If ARCHITRA$ = "COKARDEX" Then
      Screen.MousePointer = 11
      Call CIERRARCH("*.*")
      Call LIBARCH("*.*")
      Call CLEARFORM
      Screen.MousePointer = 1
      Hide
      Exit Sub
    End If
    '
    LITAB.Clear
    Call VERISCROLL
    If TVL%(CFF%(Index + 1)) = 5 Then
         FIENTRA(Index).Text = "."
         Call SELECHO("SELFECHA")
         FIENTRA(Index).Text = VALACT1$("SELFECHA")
         IANT% = Index
         Exit Sub
    End If
    '
    If TRIM$(AEF$(Index + 1)) <> "" Then
         FIENTRA(Index).Text = "."
         Call SELECHO(AEF$(Index + 1))
         FIENTRA(Index).Text = VALACT1$(AEF$(Index + 1))
         Call ENTRAFIL
       Else
         Call MENSERR(24, "No hay Tabla de Validación\para el Campo Seleccionado.")
    End If
    IANT% = Index
    '
End Sub
'

Private Sub FIECO_GotFocus(Index As Integer)
    '
    FIENTRA(Index).SetFocus
    '
End Sub
'

Private Sub FIENTRA_Change(Index As Integer)
    '
    YAFIL% = 0
    LOMACA% = LKF%(Index + 1)
    If LOMACA% < 1 Then LOMACA% = LCF%(Index + 1)
    '
    If Len(FIENTRA(Index).Text) >= LOMACA% Then
        FIENTRA(Index).Text = Left$(FIENTRA(Index).Text, LOMACA%)
        FIENTRA(Index).SelStart = 9999
    End If
    '
    If AEF$(Index + 1) <> "" Then
       If CLF%(Index + 1) > 0 Then
          If LEF%(Index + 1) > 0 Then
             VDEC$ = ""
             If TRIM$(UCase$(AEF$(Index + 1))) <> "MASTER" Then
                XXZ$ = CBIN$(FIENTRA(Index).Text, TVL%(CFF%(Index + 1)), LF%(CFF%(Index + 1)))
                VDEC$ = ECOARCH$(AEF$(Index + 1), XXZ$)
               Else
                CI% = CODINT%(FIENTRA(Index).Text)
                VDEC$ = DESCRIT$(CI%)
             End If
             '
             FIECO(Index).Text = VDEC$
             '
          End If
       End If
    End If
    '
End Sub
'

Private Sub FIENTRA_Click(Index As Integer)
    '
    If ARCHITRA$ = "COKARDEX" Then
       Call CIERRARCH("*.*")
       Call FINAL("")
    End If
    '
End Sub

Private Sub FIENTRA_DBLCLICK(Index As Integer)
    '
    If ARCHITRA$ = "COKARDEX" Then
       Call CIERRARCH("*.*")
       Call FINAL("")
    End If
    '
    If TVL%(CFF%(Index + 1)) = 5 Then
         FIENTRA(Index).Text = "."
         Call SELECHO("SELFECHA")
         FIENTRA(Index).Text = VALACT1$("SELFECHA")
         IANT% = Index
         Exit Sub
    End If
    '
    If TRIM$(AEF$(Index + 1)) <> "" Then
         FIENTRA(Index).Text = "."
         Call SELECHO(AEF$(Index + 1))
         FIENTRA(Index).Text = VALACT1$(AEF$(Index + 1))
         Call ENTRAFIL
       Else
         Call MENSERR(24, "No hay Tabla de Validación\para el Campo Seleccionado.")
    End If
    IANT% = Index
    '
End Sub
'

Private Sub FIENTRA_KeyPress(Index As Integer, KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
        YAFIL% = 0
        Call ENTRAFIL
    End If
    '
End Sub
'

Private Sub BOTEXIT_Click()
    '
    Screen.MousePointer = 11
    Call CIERRARCH("*.*")
    Call LIBARCH("*.*")
    Call CLEARFORM
    Screen.MousePointer = 1
    Hide
    '
End Sub
'

Private Sub Form_Activate()
    '
    Refresh
    DoEvents
    Call ENTRAFIL
    If MODEDI% = 1 Then
      If YAPRE% = 0 Then
        Call ADDLISTA
      End If
    End If
    '
End Sub
'


Private Sub Form_Load()
    '
    YAFIL% = 0
    Enabled = True
    TIFORTAB% = 1
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Screen.MousePointer = 11
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Call CLEARFORM
   Screen.MousePointer = 1
   Hide
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If USNBR% > 100 Then
     If Button = 2 Then
       If Shift = 1 Then
         If EXISTE%("PARAMED.EXE") > 0 Then
           Call CIERRARCH("*.*")
           Call CONECRUN("#PARAMED/" + PROPRIN$, "Definición Formulario Consulta Tabulada")
           Call FINAL("")
         End If
       End If
     End If
   End If

End Sub

Private Sub Label1_Change()
   Visible = False
   If Label1 <> "" Then
     Call CARDATFORM
     Call DIBUFORM
   End If
End Sub

Private Sub LITAB_DblClick()
    Call EDLISTA
End Sub

Sub CARDATFORM()
     '
     TIFORMED% = 2
     For II% = 0 To 31
       FILABEL(II%).Visible = False
       FIENTRA(II%).Visible = False
       FIECO(II%).Visible = False
       FIBOTSE(II%).Visible = False
       LINVER(II%).Visible = False
       If II% > 0 Then
         FILABEL(II%).AutoSize = FILABEL(0).AutoSize
         FILABEL(II%).Alignment = FILABEL(0).Alignment
         FILABEL(II%).Font = FILABEL(0).Font
         FILABEL(II%).FontName = FILABEL(0).FontName
         FILABEL(II%).FontSize = FILABEL(0).FontSize
         FILABEL(II%).FontBold = FILABEL(0).FontBold
       End If
       RTT%(II% + 1) = 0
     Next II%
     '
     FORMA$ = Label1
5800 GRABACT% = 0
     Dim YYY As String * 128
     If N15% > 0 Then Close #N15%
     N15% = FILEOPEN%(FORMA$, 9, 128)
     Get #N15%, 1, YYY$
     '
     If LOF(N15%) <= 128 Then
         On Error Resume Next
         YYY$ = String$(128, 0)
         Put #N15%, 1, YYY$
         Close #N15%
         Kill FORMA$
         On Error GoTo 0
         FORMB$ = UCase$(REPLACAR$(FORMA$, "\", "/"))
         Call MENSERR(24, "Proceso Interrumpido.\Falta Controlador de Procesos '" + FORMB$ + "'.")
         Close: End
     End If
     '
     Archi$ = Mid$(YYY$, 57, 8)
     ARCHITRA$ = TRIM$(Archi$)
     ARCHILOC$ = "!" + ARCHITRA$
     PROCE$ = TRIM$(Mid$(YYY$, 65, 8))
     LOR1% = Asc(Mid$(YYY$, 106, 1)): If LOR1% > 16 Then LOR1% = 0
     LOR% = 256 * LOR1% + Asc(Mid$(YYY$, 107))
     CAMAX& = CVI(Mid$(YYY$, 108))
     PEOF% = 1 + Asc(Mid$(YYY$, 110))
     EOFS$ = Mid$(YYY$, 111, 16)
5803 S1$ = Left$(YYY$, 48)
     Sali$ = Mid$(YYY$, 49, 8)
     COD$ = Mid$(YYY$, 81)
     RGI& = Asc(Mid$(COD$, 1))
     RGF& = Asc(Mid$(COD$, 2))
     OFS% = Asc(Mid$(COD$, 3))
5804 BLP% = Asc(Mid$(COD$, 4))
     VIV% = Asc(Mid$(COD$, 5))
     APQ% = Asc(Mid$(COD$, 6))
        BOTAGRE.Enabled = True
        If APQ% < 1 Then
            BOTAGRE.Enabled = False
        End If
     IDL% = Asc(Mid$(COD$, 7))
        BOTINSE.Enabled = True
        BOTERASE.Enabled = True
        If IDL% < 1 Then
            BOTINSE.Enabled = False
            BOTERASE.Enabled = False
        End If
     BOTEDIT.Enabled = False
     HDR& = Asc(Mid$(COD$, 8))
     RFT% = Asc(Mid$(COD$, 9))
     If IDL% > 0 Or APQ% > 0 Then GRABACT% = 1
     PRX% = Asc(Mid$(COD$, 10))
     NCL% = 0
     NFT% = 0
     NCLF% = 0
     TCL% = 0
     HAYTOTA% = 0
     RG% = 1
     '
5805 RG% = RG% + 1
     If RG% > LOF(N15%) / 128 Or RG% > 33 Then GoTo 5820
5806 Get #N15%, RG%, YYY$
     If YYY$ = String$(128, 0) Then GoTo 5820
     '
5807 TCL% = TCL% + 1
     TVL%(TCL%) = Asc(Mid$(YYY$, 33))
     LDL%(TCL%) = Asc(Mid$(YYY$, 34))
         If RG% = 2 Then TVL%(TCL%) = 3: LDL%(TCL%) = 0
5808 PR%(TCL%) = Asc(Mid$(YYY$, 35))
     LF%(TCL%) = Asc(Mid$(YYY$, 36))
         If RG% = 2 Then
           PR%(TCL%) = 0
           LF%(TCL%) = 2
         End If
5809 TT%(TCL%) = Asc(Mid$(YYY$, 88))
     If TT%(TCL%) > 0 Then
        HAYTOTA% = 1
        RTT%(TT%(TCL%) Mod 100) = 0
        If TT%(TCL%) <> TCL% Then
            RTT%(TT%(TCL%) Mod 100) = 1
        End If
     End If
     '
5811 If Asc(Mid$(YYY$, 69)) < 1 Then GoTo 5816 ' SUPRIME COLUMNAS LONGITUD 0
     '
5812 NCL% = NCL% + 1
     CLT%(TCL%) = NCL%
     CFT%(NCL%) = RG% - 1
     CL0%(NCL%) = RG% - 1
     ECL$(NCL%) = TRIM$(Mid$(YYY$, 37, 32))
     If Left$(ECL$(NCL%), 1) = "@" Then
          NCLF% = NCLF% + 1
     End If
5813 ACL%(NCL%) = Asc(Mid$(YYY$, 69))
     LK%(NCL%) = Asc(Mid$(YYY$, 70))
5814 MG%(NCL%) = Asc(Mid$(YYY$, 71))
     PE%(NCL%) = Asc(Mid$(YYY$, 73))
          If RG% = 2 Then PE%(NCL%) = 2
     If PE%(NCL%) = 1 Then
        GRABACT% = 1
        BOTEDIT.Enabled = True
     End If
5815 JT%(NCL%) = Asc(Mid$(YYY$, 72))
     AE$(NCL%) = Mid$(YYY$, 74, 8)
     CE%(NCL%) = Asc(Mid$(YYY$, 82))
5816 GoTo 5805
5817 '
5820 ATOT% = NCL% - 1 - NCLF%
     CLII% = OFS% + 2
     For KCL% = 1 To NCL%
        LC%(KCL%) = ACL%(KCL%)
        If Left$(ECL$(KCL%), 1) <> "@" Then
            LC%(KCL%) = ACL%(KCL%)
            ATOT% = ATOT% + ACL%(KCL%)
            CL%(KCL%) = CLII%
            CLII% = CLII% + ACL%(KCL%) + 1
        End If
     Next KCL%
     '
     If ATOT% > 85 Then
        FORMAA$ = REPLACAR$(FORMA$, "\", "/")
        MENSA$ = "Error: Imposible Presentar Formulario Tabulado.\Ancho de Pantalla Tabulada " + FORMAA$ + " =" + Str$(ATOT% + 2) + " Pos."
        If USNBR% < 100 Then
          MENSA$ = MENSA$ + "\  \Informe al Supervisor del Sistema."
        End If
        Call MENSERR(24, MENSA$)
        Call CIERRARCH("*.*")
        If USNBR% > 100 Then
          If EXISTE%("PARAMED.EXE") > 0 Then
            Call CONECRUN("#PARAMED/" + PROPRIN$, "Definición Formulario Consulta Tabulada")
          End If
        End If
        Call FINAL("")
     End If
     '
5821 For KCL% = 1 To NCL%
        If CE%(KCL%) <= 0 Then GoTo 5824
5822    For KCL1% = 1 To NCL%
           If CL0%(KCL1%) = CE%(KCL%) Then
               CE%(KCL%) = KCL1%: PE%(KCL1%) = 3
               GoTo 5824
           End If
5823    Next KCL1%
5824 Next KCL%
     '
5825 If BLP% > 0 Then
'         If OFS% = 0 Then
'             OFS% = (78 - ATOT%) / 2
'             For KCL% = 1 To NCL%
'                 CL%(KCL%) = CL%(KCL%) + OFS%
'             Next KCL%
'         End If
     End If
     '
5830 NFT% = 0
     RG% = 1
     RG0% = RGI&
     RGFI% = 25: RGFS% = 1
     '
5835 RG% = RG% + 1
     If RG% > LOF(N15%) / 128 Or RG% > 33 Then GoTo 5850
5840 Get #N15%, RG%, YYY$
     If YYY$ = String$(128, 0) Then GoTo 5850
5841 If Asc(Mid$(YYY$, 83)) < 1 Then GoTo 5847 ' SUPRIME RG-FILTRO=0
5842 NFT% = NFT% + 1
     CFF%(NFT%) = RG% - 1
     LCF%(NFT%) = Asc(Mid$(YYY$, 83))
     RGFL&(NFT%) = Asc(Mid$(YYY$, 84))
     CLF%(NFT%) = Asc(Mid$(YYY$, 85))
5843 If RGFL&(NFT%) - 1 < RG0% Then
         RG0% = RGFL&(NFT%) - 1
     End If
     If RGFL&(NFT%) < RGFI% Then RGFI% = RGFL&(NFT%)
     If RGFL&(NFT%) > RGFS% Then RGFS% = RGFL&(NFT%)
     '
5844 ECLF$(NFT%) = RTrim$(Mid$(YYY$, 37, 32))
     If Len(ECLF$(NFT%)) > 0 Then
         ECLF$(NFT%) = ECLF$(NFT%) + ":"
     End If
5845 LKF%(NFT%) = Asc(Mid$(YYY$, 70))
     LEF%(NFT%) = Asc(Mid$(YYY$, 86))
     CEF%(NFT%) = Asc(Mid$(YYY$, 87))
5846 AEF$(NFT%) = Mid$(YYY$, 74, 8)
5847 GoTo 5835
 '
5850 Close #N15%: N15% = 0
     '
'     Call CLEARFORM
     '
End Sub
'

Sub DIBUFORM()     ' Rutina de dibujo pantalla tabulada
'
HREN% = ScaleHeight / 25
Caption = TRIM$(S1$)
'
EPAC$ = TRIM$(EMPREAC$)
If EPAC$ <> "" Then
    Caption = Caption + "  -  " + EPAC$
End If
'
MAIZQ = 0
MAINT = 170
MADER = 220
'
ENCA$ = " "
'
CALINVE% = 0
For KCL% = 1 To NCL%
      '
      If Left$(ECL$(KCL%), 1) <> "@" Then
          If Left$(ECL$(KCL%), 1) = "<" Then
              ENCACO$ = AJUSTI$(Mid$(ECL$(KCL%), 2), LC%(KCL%))
              ENCACO$ = Space$(MG%(KCL%)) + ENCACO$
              ENCACO$ = Left$(ENCACO$, LC%(KCL%))
            ElseIf Left$(ECL$(KCL%), 1) = ">" Then
              ENCACO$ = AJUSTD$(Mid$(ECL$(KCL%), 2), LC%(KCL%))
              ENCACO$ = Mid$(ENCACO$, MG%(KCL%) + 1) + Space$(MG%(KCL%))
            Else
              ENCACO$ = TRIM$(ECL$(KCL%))
              While Len(ENCACO$) < LC%(KCL%)
                 ENCACO$ = " " + ENCACO$ + " "
              Wend
              ENCACO$ = Left$(ENCACO$, LC%(KCL%))
          End If
          ENCA$ = ENCA$ + ENCACO$
          ENCA$ = ENCA$ + " "
          TEPRUEBA.Caption = ENCA$
          CALINVE% = CALINVE% + 1
          LINVER(CALINVE% - 1).Left = TEPRUEBA.Width
        Else
          ' CL%(KCL%) = 4 + Len(ECL$(KCL%))
      End If
      '
Next KCL%
'
ENCA$ = ENCA$ + "   "
TOTALI.Caption = ""
TEPRUEBA.Caption = ENCA$
LITAB.Width = TEPRUEBA.Width - 50
FRAMENCA.Width = TEPRUEBA.Width - 50
ENCACOL.Caption = ENCA$
Picture1.Width = TEPRUEBA.Width - 50
ENCACOL.Width = TEPRUEBA.Width - 50
FRAMETOT.Width = TEPRUEBA.Width - 50
TOTALI.Width = TEPRUEBA.Width - 50
LITAB.Left = MAIZQ
FRAMENCA.Left = MAIZQ
ENCACOL.Left = 50
FRAMETOT.Left = MAIZQ
TOTALI.Left = 50
Width = MAIZQ + LITAB.Width + MAINT + FRAMEBOT1.Width + MADER
COEFH = LITAB.Width / Len(ENCA$)
'
For II% = 1 To NFT%
   '
   If TRIM$(ECLF$(II%)) <> "" Then
      CLL% = CLF%(II%) - 2
      FILABEL(II% - 1).Caption = ECLF$(II%) + " "
      FILABEL(II% - 1).Left = MAIZQ + COEFH * (CLF%(II%) - Len(ECLF$(II%)) - 3)
      FILABEL(II% - 1).Top = 150 + (0.5 + (RGFL&(II%) - RGFI%) * 1.5) * FIENTRA(II% - 1).Height
      FILABEL(II% - 1).Width = COEFH * (2 + Len(ECLF$(II%)))
      FILABEL(II% - 1).Visible = True
   End If
   '
   FIENTRA(II% - 1).Left = COEFH * (CLF%(II%) - 1)
   FIENTRA(II% - 1).Top = 70 + (0.5 + (RGFL&(II%) - RGFI%) * 1.5) * FIENTRA(II% - 1).Height
   '
   LCFF% = LCF%(II%)
   TEPRU.Font = FIENTRA(II% - 1).Font
   TEPRU.FontSize = FIENTRA(II% - 1).FontSize
   TEPRU.FontBold = FIENTRA(II% - 1).FontBold
   TEPRU.Caption = String$(LCFF%, "M") + "."
   FIENTRA(II% - 1).Width = TEPRU.Width
   If LCFF% >= 6 Then
        FIENTRA(II% - 1).Width = COEFH * (1 + LCFF%)
   End If
   '
   FIENTRA(II% - 1).Visible = True
   '
   If TRIM$(AEF$(II%)) <> "" Then
      If LEF%(II%) > 0 Then
         FIECO(II% - 1).Left = COEFH * (CEF%(II%) - 1)
         FIECO(II% - 1).Top = 70 + (0.5 + (RGFL&(II%) - RGFI%) * 1.5) * FIENTRA(II% - 1).Height
         FIECO(II% - 1).Width = COEFH * (LEF%(II%) + 1)
         FIECO(II% - 1).Visible = True
      End If
   End If
   '
   If TRIM$(AEF$(II%)) <> "" Or TVL%(CFF%(II%)) = 5 Then
      If Left$(UCase$(AEF$(II%)), 4) <> "AUTO" Then
         FIBOTSE(II% - 1).Top = FIENTRA(II% - 1).Top
         FIBOTSE(II% - 1).Height = FIENTRA(II% - 1).Height
         FIBOTSE(II% - 1).Width = FIBOTSE(II% - 1).Height * 0.6
         FIBOTSE(II% - 1).Left = FIENTRA(II% - 1).Left + FIENTRA(II% - 1).Width
         FIBOTSE(II% - 1).Visible = True
      End If
   End If
   '          '
Next II%
'
If NFT% > 0 Then
     Picture1.Visible = True
     FRAMENCA.Top = (2.5 + (RGFS% - RGFI%) * 1.5) * FIENTRA(0).Height
   Else
     Picture1.Visible = False
     FRAMENCA.Top = 0.5 * FIENTRA(0).Height
End If
FRAMENCA.Height = 2 * ENCACOL.Height
ENCACOL.Top = (FRAMENCA.Height - ENCACOL.Height) / 2
ENCACOL.Caption = ENCA$
LITAB.Top = FRAMENCA.Top + FRAMENCA.Height
'
FRAMETOT.Height = 2.2 * TOTALI.Height
LITAB.Height = Height - 200 - LITAB.Top - 0.2 * FIENTRA(0).Height - FRAMETOT.Height
FRAMETOT.Top = LITAB.Top + LITAB.Height
TOTALI.Top = (FRAMETOT.Height - TOTALI.Height) / 2
FRAMETOT.Visible = True
'
Picture1.Height = FRAMETOT.Top + FRAMETOT.Height
HTOPE = Height - ScaleHeight
If Picture1.Height > ScaleHeight Then
  Height = Picture1.Height + HTOPE
End If
'
Image3.Top = FRAMETOT.Top '- 60
Image3.Left = FRAMETOT.Width - 370
'
For KKI% = 1 To CALINVE%
  LINVER(KKI% - 1).Top = FRAMENCA.Top
  LINVER(KKI% - 1).Height = FRAMETOT.Top - FRAMENCA.Top + FRAMETOT.Height
  LINVER(KKI% - 1).Visible = True
Next KKI%
'
Line1.X1 = FRAMETOT.Left
Line1.X2 = 12000 ' FRAMETOT.Left + FRAMETOT.Width
Line1.Y1 = FRAMETOT.Top + FRAMETOT.Height
Line1.Y2 = FRAMETOT.Top + FRAMETOT.Height
'
Line2.X1 = FRAMENCA.Left + FRAMENCA.Width
Line2.X2 = FRAMENCA.Left + FRAMENCA.Width
Line2.Y1 = 0  ' FRAMENCA.Top
Line2.Y2 = FRAMETOT.Top + FRAMETOT.Height
'
FRAMEBOT1.Top = FIENTRA(0).Height * 0.5
FRAMEBOT1.Left = MAIZQ + LITAB.Width + MAINT
FRAMEBOT2.Top = FRAMETOT.Top - 2560
FRAMEBOT2.Left = MAIZQ + LITAB.Width + MAINT
'
MARCOBARRA.Top = FRAMEBOT2.Top - 300
MARCOBARRA.Left = FRAMEBOT2.Left
MARCOBARRA.Width = FRAMEBOT2.Width
'
LITAB.Clear
While LITAB.TopIndex <= 0
   LITAB.AddItem "   "
   'Call VERISCROLL
   LITAB.ListIndex = LITAB.ListCount - 1
Wend
CARELI% = LITAB.ListCount - 1
' AQUI ESTA BIEN
LITAB.Clear
'Call VERISCROLL
'
VScroll1.Top = LITAB.Top
VScroll1.Height = LITAB.Height
VScroll1.Left = LITAB.Left + LITAB.Width - VScroll1.Width - 20
VScroll1.Visible = True
'
Refresh
DoEvents
'
End Sub

Sub CLEARFORM()
    '
    YAFIL% = 0
    LITAB.Clear
    Call VERISCROLL
    For FT% = 0 To 31
        FIENTRA(FT%).Text = ""
        FIENTRA(FT%).Enabled = True
    Next FT%
    '
    For K2% = 1 To 32
       TTX#(K2%) = 0
       TXX#(K2%) = 0
    Next K2%
    '
End Sub
'
Sub ENTRAFIL()
    '
    On Error Resume Next
    SetFocus
    On Error GoTo 0
    If YAFIL% < 1 Then
       LITAB.Clear
       Call VERISCROLL
       For K2% = 1 To 32
          TTX#(K2%) = 0
          TXX#(K2%) = 0
       Next K2%
       '
       For FT% = 1 To NFT%
          FIENTRA(FT% - 1).Visible = True
          FIENTRA(FT% - 1).Enabled = True
          ARCHECO$ = UCase$(TRIM$(AEF$(FT%)))
          If ARCHECO$ = "AUTO" Then
             If Val(FIENTRA(FT% - 1).Text) < 1 Then
                REG& = ULTREG&(ARCHITRA$)
                TF1% = TIPFIELD%(ARCHITRA$, CFF%(FT%) - 1)
                LF1% = LENFIELD%(ARCHITRA$, CFF%(FT%) - 1)
                PF1% = POSFIELD%(ARCHITRA$, CFF%(FT%) - 1)
                XXZ$ = Mid$(REGLEIDO$(ARCHITRA$, REG&), PF1%, LF1%)
                NRO = 1 + Val(CSTRING$(XXZ$, TF1%, 12, 0, 1))
                FIENTRA(FT% - 1).Text = TRIM$(Str$(NRO))
             End If
          End If
       Next FT%
       '
       For FT% = 1 To NFT%
          If FT% = 1 Then
            If ATRIBU$ <> "" Then
               FIENTRA(0).Text = ATRIBU$
               ATRIBU$ = ""
            End If
          End If
          '
          If LKF%(FT%) > 0 Then
             If TRIM$(FIENTRA(FT% - 1).Text) = "" Then
                 On Error Resume Next
                 FIENTRA(FT% - 1).Enabled = True
                 FIENTRA(FT% - 1).SetFocus
                 On Error GoTo 0
                 Exit Sub
             End If
             '
             If TRIM$(FIENTRA(FT% - 1).Text) = "." Then
                 Exit Sub
             End If
             '
           End If
       Next FT%
       '
       Call CARLITAB(FINOVA%)
       If FINOVA% > 0 Then
          'YAFIL% = 1
          Call CLEARFORM
          Exit Sub
       End If
       '
       If LITAB.ListCount > 0 Then
           YAFIL% = 2
           LITAB.ListIndex = 0
           '
           If TRIM$(UCase$(PROPRIN$)) = "KARDEX" Then
              If LITAB.ListCount > 20 Then
                 LITAB.TopIndex = LITAB.ListCount - 20
              End If
              LITAB.ListIndex = LITAB.ListCount - 1
           End If
           '
           Exit Sub
         Else
           '
           If APQ% < 1 Then
               If NFT% < 1 Then
                    YAFIL% = (-1)
                    Call MENSERR(24, "Archivo de Trabajo Vacío")
                    Exit Sub
                  Else
                    Call MENSERR(24, "Ningun Elemento Cumple\Condición de Filtro")
                    If NFT% = 1 Then
                       If LKF%(1) < 1 Then
                          Unload WEDITAB
                       End If
                    End If
                    
               End If
               Call CLEARFORM
               Exit Sub
           End If
           '
           For FT% = 1 To NFT%
              If TRIM$(FIENTRA(FT% - 1).Text) = "" Then
                 FIENTRA(FT% - 1).Enabled = True
                 FIENTRA(FT% - 1).SetFocus
                 Exit Sub
              End If
           Next FT%
           YAFIL% = 1
           If APQ% > 0 Then
               Call ADDLISTA
               Exit Sub
             Else
               If NFT% < 1 Then
                    YAFIL% = (-1)
                    Call MENSERR(24, "Archivo de Trabajo Vacío")
                    Exit Sub
                  Else
                    Call MENSERR(24, "Ningun Elemento Cumple\Condición de Filtro")
               End If
               Call CLEARFORM
               Exit Sub
           End If
       End If
       '
    End If
    '
    If YAFIL% < 2 Then
       For FT% = 1 To NFT%
          If LKF%(FT%) < 1 Then
              If TRIM$(FIENTRA(FT% - 1).Text) = "" Then
                 FIENTRA(FT% - 1).SetFocus
                 Exit Sub
              End If
          End If
       Next FT%
       YAFIL% = 2
    End If
    '
End Sub
'
Sub ADDLISTA()
    '
    If NFT% > 0 Then
       If YAFIL% < 1 Then
          Call ENTRAFIL
          Exit Sub
       End If
    End If
    '
    LITAB.AddItem Space$(ATOT%)
    Call VERISCROLL
    LITAB.ListIndex = LITAB.ListCount - 1
    If LITAB.TopIndex < LITAB.ListIndex - CARELI% Then
        LITAB.TopIndex = LITAB.ListCount - CARELI% - 3
    End If
    MODEDI% = 1
    NUREGIS& = LITAB.ListCount + 1
    REACTIVO$ = REGBLAN$(ARCHILOC$)
    Call PRENTRA
    '
End Sub
'

Sub INSLISTA()
    '
    If NFT% > 0 Then
       If YAFIL% < 1 Then
          Exit Sub
       End If
    End If
    '
    ILISTA& = LITAB.ListIndex
    If ILISTA& < 0 Or ILISTA& > LITAB.ListCount - 1 Then
        ILISTA& = 0
    End If
    '
    LITAB.AddItem Space$(ATOT%), ILISTA&
    Call VERISCROLL
    LITAB.ListIndex = ILISTA&
    If LITAB.TopIndex < LITAB.ListIndex - CARELI% Then
        LITAB.TopIndex = LITAB.ListIndex - CARELI% / 2
    End If
    '
    For u& = ULTREG&(ARCHILOC$) To ILISTA& + 1 Step -1
       REACTIVO$ = REGLEIDO$(ARCHILOC$, u&)
       Call GRAREG(ARCHILOC$, REACTIVO$, u& + 1)
    Next u&
    REACTIVO$ = REGBLAN$(ARCHILOC$)
    Call GRAREG(ARCHILOC$, REACTIVO$, ILISTA& + 1)
    '
    MODEDI% = 2
    NUREGIS& = ILISTA& + 1
    REACTIVO$ = REGBLAN$(ARCHILOC$)
    Call PRENTRA
    '
End Sub
'
Sub BORLISTA()
    '
    If NFT% > 0 Then
       If YAFIL% < 1 Then
          Exit Sub
       End If
    End If
    '
    ILISTA& = LITAB.ListIndex
    If ILISTA& < 0 Or ILISTA& > LITAB.ListCount - 1 Then
        Exit Sub
    End If
    '
    LITAB.RemoveItem ILISTA&
    Call VERISCROLL
    For u& = ILISTA& + 2 To ULTREG&(ARCHILOC$)
       REACTIVO$ = REGLEIDO$(ARCHILOC$, u&)
       Call GRAREG(ARCHILOC$, REACTIVO$, u& - 1)
    Next u&
    REACTIVO$ = String$(RECLEN%(ARCHILOC$), 0)
    Call GRAREG(ARCHILOC$, REACTIVO$, ULTREG&(ARCHILOC$))
    Call CIERRARCH(ARCHILOC$)
    '
    If ILISTA& > LITAB.ListCount - 1 Then
       ILISTA& = LITAB.ListCount - 1
    End If
    LITAB.ListIndex = ILISTA&
    '
    NUREGIS& = ILISTA& + 1
    REACTIVO$ = REGLEIDO$(ARCHILOC$, NUREGIS&)
    '
End Sub
'

Sub EDLISTA()
    '
    If LITAB.ListIndex < 0 Then
        Exit Sub
    End If
    '
    If LITAB.ListIndex >= LITAB.ListCount Then
        Exit Sub
    End If
    '
    MODEDI% = 4
    NUREGIS& = LITAB.ListIndex + 1
    REACTIVO$ = REGLEIDO$(ARCHILOC$, NUREGIS&)
    '
    Call PRENTRA
    '
End Sub
'
Sub BULISTA()
    '
    Call OPNOIN("")
    Exit Sub
    '
    If LITAB.ListCount < 1 Then
        Exit Sub
    End If
    MODEDI% = 5
    Call PRENTRA
    '
End Sub
'
Sub CARLITAB(FINOVA%)
    '
    Static N18%
    Static IXT As String * 1024
    '
    Screen.MousePointer = 11
    FINOVA% = 0
    '
    SORTLIST.Clear
    LITAB.Visible = False
    LITAB.Clear
    Call VERISCROLL
    'Call BLANARCH(ARCHILOC$)
    FINDEXLOC% = FILEINDEX%(ARCHILOC$)
    FINUMELOC% = FILENBR%(ARCHILOC$)
    LOREIILOC% = LOREGIS%(FINDEXLOC%)
    URELOC& = 0
    '
    For K2% = 1 To 32
       TTX#(K2%) = 0
       TXX#(K2%) = 0
    Next K2%
    TOTALI.Caption = ""
    '
    If ULTREG&(ARCHITRA$) > 0 Then
       MARCOBARRA.Top = FRAMEBOT2.Top - 300
       MARCOBARRA.Left = FRAMEBOT2.Left
       MARCOBARRA.Width = FRAMEBOT2.Width
    End If
    '
    ' FINUII% = FILENBR%(ARCHITRA$)
    ' LOREII% = RECLEN%(ARCHITRA$)
    FINDEX% = FILEINDEX%(ARCHITRA$)
    FINUME% = FILENBR%(ARCHITRA$)
    LOREII% = LOREGIS%(FINDEX%)
    '
    If NFT% > 0 Then
       '
       CONTA& = 0
       IDX% = 0
       CARESEL& = 0
       ORDE% = 0: If TVL%(CFT%(1)) = 5 Then ORDE% = 1
       '
       FT1% = 0
       For FT% = 1 To NFT%
          FI$(FT%) = ""
          If LKF%(FT%) > 0 Then
             FT1% = FT%
             FI$(FT%) = CBIN$(FIENTRA(FT% - 1).Text, TVL%(CFF%(FT%)), LF%(CFF%(FT%)))
             If TRIM$(AEF$(FT%)) <> "" Then
                If Left$(AEF$(FT%), 4) <> "AUTO" Then
                   If ECOARCH$(AEF$(FT%), FI$(FT%)) = "" Then
                      On Error GoTo 999
                      FIENTRA(FT% - 1).SetFocus
                      On Error GoTo 0
                      Call MENSERR(24, "Filtro no Válido")
                      FINOVA% = 1
                      GoTo 999
                   End If
                End If
             End If
          End If
       Next FT%
       '
       If FT1% = 1 Then
          ' intenta seleccion indexada
          'COLUX$ = TRIM$(Str$(CFF%(1) - 1))
          'While Len(COLUX$) < 2: COLUX$ = "0" + COLUX$: Wend
          'ARIX$ = ARCHITRA$ + ".X" + COLUX$
          '
          'If N18% > 0 Then Close #N18%
          'N18% = FILEOPEN%(LU$ + ARIX$, 0, 1024)
          '
          'If LOF(N18%) < 1 Then
          '      IXT$ = String$(1024, 0)
          '      Put #N18%, 1, IXT$
          '      Close #N18%: N18% = 0
          '      Kill LU$ + ARIX$
          '   Else             ' leer en forma secuencial
                IDX% = 1
                ARGU$ = FI$(1)
                LARGU& = Len(ARGU$)
                RFKX$ = RECFILT$(ARCHITRA$, CFF%(1) - 1, ARGU$)
                For IX& = 1 To Len(RFKX$) Step 4
                  u& = CVS(Mid$(RFKX$, IX&, 4))
                  GoSub 701
                Next IX&
                '
                'For IX& = 1 To LOF(N18%) / 1024 + 1
                '   Get #N18%, IX&, IXT$
                '   PINI& = 1
650             '   POSI& = InStr(PINI&, IXT$, ARGU$)
                '   If POSI& >= PINI& Then
                '      If (POSI& Mod LARGU&) = 1 Then
                '            U& = (1024& * (IX& - 1) + (POSI& + LARGU& - 1)) / LARGU&
                '            GoSub 701                ' EXTRACTAR
                '            PINI& = POSI& + LARGU&
                '         Else
                '            PINI& = POSI& + 1
                '      End If
                '      GoTo 650
                '   End If
                'Next IX&
                GoTo 990
          'End If
       End If
    End If
    '
    TINI = Timer
    Screen.MousePointer = 11
    '
    FIN& = ULTREG&(ARCHITRA$)
    For u& = 1 To FIN&
        '
        If (Timer - TINI) > 1 Then
            MARCOBARRA.Visible = True
        End If
        '
        Call TRACE(20, u&, FIN&)
701     'XXX$ = LEEREG$(FINUII%, LOREII%, U&)
        'XXX$ = REGLEIDO$(ARCHITRA$, U&)
        XXX$ = LEEREGIS$(FINDEX%, FINUME%, LOREII%, u&)
        RELIB$ = Space$(ATOT% + 2)
        '
        For FT% = 1 To NFT%
            If LKF%(FT%) > 0 Then
               If Mid$(XXX$, PR%(CFF%(FT%)), LF%(CFF%(FT%))) <> FI$(FT%) Then GoTo 979
            End If
        Next FT%
        '
        CARESEL& = CARESEL& + 1
        If CARESEL& = 1 Then
            For FT% = 1 To NFT%
               If LKF%(FT%) < 1 Then
                  XXZ$ = Mid$(XXX$, PR%(CFF%(FT%)), LF%(CFF%(FT%)))
                  TXT$ = CSTRING$(XXZ$, TVL%(CFF%(FT%)), LCF%(FT%), LDL%(CFF%(FT%)), 1)
                  FIENTRA(FT% - 1).Text = TXT$
'                  FIENTRA(FT% - 1).Enabled = False
               End If
            Next FT%
        
' SELECCIONAR Y PRESENTAR FILTROS
        End If
        '
        For K2% = 1 To NCL%
           If CFT%(K2%) <> 1 Then
             If RTT%(CFT%(K2%)) < 1 Then
               If LF%(CFT%(K2%)) > 0 Then
                   XXZ$ = Mid$(XXX$, PR%(CFT%(K2%)), LF%(CFT%(K2%)))
                   ACLX% = ACL%(K2%) - MG%(K2%)
                   TXT$ = CSTRING$(XXZ$, TVL%(CFT%(K2%)), ACLX%, LDL%(CFT%(K2%)), JT%(K2%))
                   '
                   If JT%(K2%) = 2 Then
                       TXT$ = TXT$ + Space$(MG%(K2%))
                     Else
                       TXT$ = Space$(MG%(K2%)) + TXT$
                   End If
                   If CL%(K2%) > 0 Then Mid$(RELIB$, CL%(K2%), LC%(K2%)) = TXT$
                   '
                   If ORDE% = 1 Then
                      If K2% = 1 Then
                         SORTAR$ = CSTRING$(XXZ$, 1, 6, 0, 2) + CSTRING$(MKS$(u&), 3, 9, 0, 2)
                      End If
                   End If
                   '
                   If TRIM$(AE$(K2%)) <> "" Then
                      XXE$ = ECOARCH$(AE$(K2%), XXZ$)
                      COLOG% = K2%
                      CLE% = CL%(CE%(COLOG%))
                      LCE% = LC%(CE%(COLOG%))
                      JTE% = JT%(CE%(COLOG%))
                      TXT$ = CSTRING$(XXE$, 6, LCE%, 0, JTE%)
                      If CLE% > 0 Then Mid$(RELIB$, CLE%, LCE%) = TXT$
                   End If
                   '
                   TTT% = TT%(CFT%(K2%))
                   If TTT% > 0 Then
                      If (TTT% Mod 100) <= TCL% Then
                         TTX#(K2%) = TTX#(K2%) + Val(TXT$)
                         If TTT% < 100 Then TXX#(TTT%) = TXX#(TTT%) + Val(TXT$) Else TXX#(TTT% Mod 100) = TXX#(TTT% Mod 100) - Val(TXT$)
                         If (TTT% Mod 100) <> CFT%(K2%) Then
                             COLOG% = CLT%(TTT% Mod 100)   ' CFT%(TTT% Mod 100)
                             CLE% = CL%(COLOG%)
                             LCE% = LC%(COLOG%)
                             JTE% = JT%(COLOG%)
                             TXT$ = CSTRING$(MKD$(TXX#(TTT% Mod 100)), TVL%(CFT%(COLOG%)), LCE%, LDL%(CFT%(COLOG%)), JTE%)
                             If CLE% > 0 Then Mid$(RELIB$, CLE%, LCE%) = TXT$
                         End If
                      End If
                   End If
                   '
                End If
             End If
           End If
        Next K2%
        '
        On Error GoTo 980
        If ORDE% = 1 Then
              SORTLIST.AddItem SORTAR$ + RELIB$ + "   " + Str$(u&)
            Else
              LITAB.AddItem RELIB$ + "   " + Str$(u&)
              'Call VERISCROLL
        End If
        On Error GoTo 0
        '
        If IDX% = 1 Then
            Return
        End If
        '
979 Next u&
    GoTo 990
    '
980 Resume 981
981 On Error GoTo 0
    Call MENSERR(24, "Imposible Presentar Consulta Completa.\La Base de Datos es Demasiado Grande.")
    '
990 LITAB.Visible = False
    If ORDE% = 1 Then
        For K2% = 1 To 32: TXX#(K2%) = 0: Next K2%
        FIN& = SORTLIST.ListCount
        For u& = 1 To SORTLIST.ListCount
            Call TRACE(1, u&, FIN&)
            SORTLIST.ListIndex = u& - 1
            RELIB$ = Mid$(SORTLIST.Text, 16)
            '
            For K2% = 1 To NCL%
               TTT% = TT%(CFT%(K2%))
               If TTT% > 0 Then
                   If (TTT% Mod 100) <= TCL% Then
                      VALO# = Val(Mid$(RELIB$, CL%(K2%), LC%(K2%)))
                      If TTT% < 100 Then TXX#(TTT%) = TXX#(TTT%) + VALO# Else TXX#(TTT% Mod 100) = TXX#(TTT% Mod 100) - VALO
                      If (TTT% Mod 100) <> CFT%(K2%) Then
                          COLOG% = CLT%(TTT% Mod 100)   ' CFT%(TTT% Mod 100)
                          CLE% = CL%(COLOG%)
                          LCE% = LC%(COLOG%)
                          JTE% = JT%(COLOG%)
                          TXT$ = CSTRING$(MKD$(TXX#(TTT% Mod 100)), TVL%(CFT%(COLOG%)), LCE%, LDL%(CFT%(COLOG%)), JTE%)
                          If CLE% > 0 Then Mid$(RELIB$, CLE%, LCE%) = TXT$
                       End If
                   End If
               End If                   '
            Next K2%
            '
            LITAB.AddItem RELIB$
            'Call VERISCROLL
        Next u&
    End If
    SORTLIST.Clear
    '
    LITAB.Visible = False
    LIREGSEL.Clear
    FIN& = LITAB.ListCount
    For u& = 1 To LITAB.ListCount
        Call TRACE(1, u&, FIN&)
        LITAB.ListIndex = u& - 1
        RELIB$ = Mid$(LITAB.Text, ATOT% + 3)
        LIREGSEL.AddItem RELIB$
        '
        REGPRI& = Val(RELIB$)
        'XXX$ = REGLEIDO$(ARCHITRA$, REGPRI&)
        XXX$ = LEEREGIS$(FINDEX%, FINUME%, LOREII%, REGPRI&)
        URELOC& = URELOC& + 1
        'Call REGAPP(ARCHILOC$, XXX$)
        POSGRABA& = (URELOC& - 1) * LOREIILOC% + 1
        Put #FINUMELOC%, POSGRABA&, XXX$
        '
    Next u&
    Call SETULTREG(ARCHILOC$, URELOC&)
    LITAB.Visible = True
    '
    RELIB$ = Space$(ATOT% + 2)
    For K2% = 1 To NCL%
       If TT%(CFT%(K2%)) > 0 Then
           TVLL% = 4: XXZ$ = MKD$(TTX#(K2%))
           ACO% = ACL%(K2%)
           If JT%(K2%) = 2 Then ACO% = ACO% - MG%(K2%)
           TXT$ = CSTRING$(XXZ$, TVLL%, ACO%, LDL%(CFT%(K2%)), JT%(K2%))
           If CL%(K2%) > 0 Then Mid$(RELIB$, CL%(K2%), LC%(K2%)) = TXT$
       End If
    Next K2%
    TOTALI.Caption = RELIB$
    '
999 On Error GoTo 0
    LITAB.Visible = True
    MARCOBARRA.Visible = False
    Screen.MousePointer = 1
    TOTALI.Visible = True
    Call VERISCROLL
    MODIFI% = 0
    '
End Sub
'
Sub VERISCROLL()
    '
    Static LIANT&
    If LITAB.ListCount <> LIANT& Then
        LIANT& = LITAB.ListCount
        If LITAB.ListCount > CARELI% Then
            VScroll1.Visible = False
          Else
            VScroll1.Visible = True
        End If
    End If
    '
End Sub
'
Sub PRINTABU()
   '
   NOMIMPRE$ = TRIM$(Printer.DeviceName)
   Printer.Orientation = 1
   Printer.ScaleMode = 1
   Printer.FontBold = False
   Printer.FontItalic = False
   Printer.ColorMode = 1     ' blanco y negro
   On Error Resume Next
   Printer.ForeColor = QBColor(0)
   Printer.DrawMode = 13     ' linea negra
   Printer.FillStyle = 1
   Printer.FillColor = QBColor(15)
   Printer.FontTransparent = True
   On Error GoTo 0
   '
   MARSEGUR = Int((Printer.Width - Printer.ScaleWidth) / 56.7 + 1)
   ANCHOPAP = Int(Printer.ScaleWidth / 56.7) - MARSEGUR
   If ANCHOPAP > 196 Then ANCHOPAP = 196
   LARGOPAP = Int(Printer.ScaleHeight / 56.7)
   MARGENI = 0: MARGENS = 0
   '
   ENCA$ = ""
   For KCL% = 1 To NCL%
      If Left$(ECL$(KCL%), 1) <> "@" Then
         If Left$(ECL$(KCL%), 1) = "<" Then
              ENCACO$ = AJUSTI$(Mid$(ECL$(KCL%), 2), LC%(KCL%))
              ENCACO$ = Space$(MG%(KCL%)) + ENCACO$
              ENCACO$ = Left$(ENCACO$, LC%(KCL%))
           ElseIf Left$(ECL$(KCL%), 1) = ">" Then
              ENCACO$ = AJUSTD$(Mid$(ECL$(KCL%), 2), LC%(KCL%))
              ENCACO$ = Mid$(ENCACO$, MG%(KCL%) + 1) + Space$(MG%(KCL%))
           Else
              ENCACO$ = TRIM$(ECL$(KCL%))
              While Len(ENCACO$) < LC%(KCL%)
                 ENCACO$ = " " + ENCACO$ + " "
              Wend
              ENCACO$ = Left$(ENCACO$, LC%(KCL%))
         End If
         ENCA$ = ENCA$ + ENCACO$
         If KCL% < NCL% Then
            ENCA$ = ENCA$ + " "
         End If
      End If
   Next KCL%
   ATOT% = 2 + Len(ENCA$)
   '
10 YAESTA% = 0: CAFUFIJ% = 0
   NX% = FILEOPEN("C:\FLEXOFT\PRINTERS.DRV", 0, 128)
   For IPRIN& = 1 To LOF(NX%) / 128
      Get #NX%, IPRIN&, PPZZ$
      If TRIM$(Left$(PPZZ$, 48)) = NOMIMPRE$ Then
         YAESTA% = 1
         If UCase$(Mid$(PPZZ$, 125, 2)) = "SI" Then
            CAFUFIJ% = CAFUFIJ% + 1
            FUFIJA$(CAFUFIJ%) = PPZZ$
         End If
      End If
   Next IPRIN&
   '
   Close #NX%
   NOMIMP$ = REPLACAR$(NOMIMPRE$, "\", "/")
   If YAESTA% < 1 Then
      OPX% = COMALTER%("Se ha Detectado un Modelo de Impresora\Desconocido para el Sistema Aplicativo\  \'" + NOMIMP$ + "'\  \\Configurar\Cancelar")
      If OPX% <> 1 Then Exit Sub
      Call CIERRARCH("*.*")
      Hide
      Call SETUPRINTER
      GoTo 10
   End If
   '
   If CAFUFIJ% < 1 Then
      OPX% = COMALTER%("No hay Fuentes Utilizables para Impresora\  \'" + NOMIMP$ + "'\  \\Configurar\Cancelar")
      If OPX% <> 1 Then Exit Sub
      Call CIERRARCH("*.*")
      Hide
      Call SETUPRINTER
      GoTo 10
   End If
   '
   ANCHOJA = 56.7 * (ANCHOPAP - 16)
   ATEXMAX = 0
   FUENTESEL$ = ""
   TAFUESEL = 0
   TALEBASI = 24
   '
20 TEPRUX$ = String$(ATOT%, "*")
   For I% = 1 To CAFUFIJ%
      Printer.FontName = TRIM$(Mid$(FUFIJA$(I%), 65, 48))
      If Printer.FontName <> TRIM$(Mid$(FUFIJA$(I%), 65, 48)) Then
         Call COMUNI("Se ha Detectado un Error\en la Configuración Local\de la Impresora Predeterminada.\  \Imposible Emitir el Listado Solicitado.")
         Call CIERRARCH("*.*")
         Close
         Call FILEKILL("C:\FLEXOFT\PRINTERS.DRV")
         Call FINAL("")
      End If
      '
      FUENTEMIN = CVS(Mid$(FUFIJA$(I%), 113, 4))
      FUENTEMAX = CVS(Mid$(FUFIJA$(I%), 117, 4))
      '
      TALEBA = FUENTEMAX + 0.4
      While TALEBA > FUENTEMIN
         TALEBA = TALEBA - 0.4
         Printer.FontSize = TALEBA
         If Printer.FontSize > TALEBA Then GoTo 25
         ATEXTO = 0
         On Error Resume Next
         ATEXTO = Printer.TextWidth(TEPRUX$)
         On Error GoTo 0
         If ATEXTO < 1 Then GoTo 25
         If ATEXTO <= ANCHOJA Then
            If ATEXTO > ATEXMAX Then
               FUENTESEL$ = Printer.FontName
               TAFUESEL = Printer.FontSize ' TALEBASI
               ATEXMAX = ATEXTO
            End If
            GoTo 29
         End If
         TALEBA = Printer.FontSize
25    Wend
      '
29 Next I%
   '
   ATEXTO = ATEXMAX
   If TRIM$(FUENTESEL$) = "" Or TAFUESEL < 1 Then
      Call MENSERR(24, "No hay Fuentes Adecuadas\para la Impresion de Este Listado")
      Exit Sub
   End If
   '
   Printer.FontName = FUENTESEL$
   Printer.FontSize = TAFUESEL
   HREN = Printer.TextHeight(TEPRUX$)
   MARGENI = 16 * 56.7 + ANCHOJA - ATEXTO: If MARGENI < 0 Then MARGENI = 0
   '
   MODOSEL% = 0
   CARESEL& = LITAB.SelCount
   If CARESEL& > 0 Then
      MODOSEL% = COMALTER%("Seleccione Rango de Impresión\\Todo\Texto Seleccionado\Cancelar") - 1
      If MODOSEL% < 0 Or MODOSEL% > 1 Then Exit Sub
   End If
   If MODOSEL% = 0 Then CARESEL& = LITAB.ListCount
   '
   NUHOJ% = 0: CAHO% = 0
   Printer.Orientation = 1
   GoSub IMPREFORM
   '
   On Error Resume Next
   Printer.ForeColor = QBColor(0)
   Printer.FillStyle = 1
   Printer.FillColor = QBColor(15)
   Printer.FontTransparent = True
   On Error GoTo 0
   '
   For u& = 1 To LITAB.ListCount
      If MODOSEL% = 0 Or LITAB.Selected(u& - 1) = True Then
35       POVERT = POVERT + HREN
         Printer.FontName = FUENTESEL$
         Printer.FontSize = TAFUESEL
         If POVERT > (LARGOPAP - 25) * 56.7 - 3.2 * HREN Then
            GoSub LEPIEPA
            Printer.NewPage
            '
            GoSub IMPREFORM
            GoTo 35
         End If
         '
         Printer.CurrentX = MARGENI + 100
         Printer.CurrentY = POVERT
         Printer.Print RTrim$(Mid$(LITAB.List(u& - 1), 2, ATOT% - 1))
      End If
   Next u&
   '
   If MODOSEL% = 0 Then
      Printer.CurrentX = MARGENI + 100
      Printer.CurrentY = (LARGOPAP - 25) * 56.7 - 1.5 * HREN
      Printer.Print RTrim$(Mid$(TOTALI.Caption, 2, ATOT% - 1))
   End If
   '
   GoSub LEPIEPA
   Printer.EndDoc
   Exit Sub

IMPREFORM:
   '
   Printer.ScaleMode = 1
   Printer.FontBold = False
   Printer.FontItalic = False
   '
   Printer.FontName = FUENTESEL$
   Printer.FontSize = TAFUESEL
   '
   On Error Resume Next
   Printer.ForeColor = QBColor(0)
   Printer.DrawMode = 13    ' linea negra
   Printer.DrawWidth = 2
   Printer.DrawStyle = 0
   Printer.FillStyle = 1
   Printer.FillColor = QBColor(15)
   Printer.FontTransparent = True
   On Error GoTo 0
   '
   Printer.CurrentX = 0
   Printer.CurrentY = 0
   Printer.Print "    "       ' para inicializar impresora
   '
   H1 = MARGENI: V1 = HREN
   H2 = H1 + ATEXTO: V2 = 3 * HREN
   Printer.Line (H1, V1)-(H2, V2), QBColor(0), B
   '
   Printer.CurrentX = MARGENI + 100
   Printer.CurrentY = 1.7 * HREN
   Printer.Print S1$
   '
   VFILMAX = 0
   COEF = ATEXTO / (LITAB.Width - 240)
   For II% = 1 To NFT%
      If FILABEL(II% - 1).Visible = True Then
         V2 = 4 * HREN + FIENTRA(II% - 1).Top + HREN * 1.3
         If V2 > VFILMAX Then VFILMAX = V2
      End If
   Next II%
   '
   Printer.DrawWidth = 2
   TOPETAB = 4 * HREN
   If NFT% > 0 Then
      H1 = MARGENI: V1 = 4 * HREN
      H2 = H1 + ATEXTO: V2 = VFILMAX + 0.5 * HREN
      TOPETAB = V2 + HREN
      Printer.Line (H1, V1)-(H2, V2), QBColor(0), B
   End If
   '
   Printer.DrawWidth = 1
   For II% = 1 To NFT%
      '
      If FILABEL(II% - 1).Visible = True Then
         '
         TTXX$ = FILABEL(II% - 1).Caption
         H1 = COEF * FILABEL(II% - 1).Left + MARGENI
         V1 = 4 * HREN + FILABEL(II% - 1).Top
         Printer.CurrentX = H1
         Printer.CurrentY = V1
         Printer.Print TTXX$
         '
         H1 = COEF * FIENTRA(II% - 1).Left + MARGENI
         H2 = H1 + COEF * FIENTRA(II% - 1).Width
         If H2 > MARGENI + ATEXTO - 150 Then H2 = MARGENI + ATEXTO - 150
         ANCHOCUA = H2 - H1 - 100
         V1 = 4 * HREN + FIENTRA(II% - 1).Top
         V2 = V1 + HREN * 1.3
         If V2 > VFILMAX Then VFILMAX = V2
         Printer.Line (H1, V1)-(H2, V2), QBColor(0), B
         '
         TTXX$ = FIENTRA(II% - 1).Text
34       If Printer.TextWidth(TTXX$) > ANCHOCUA Then
            If Len(TTXX$) > 0 Then
               TTXX$ = Left$(TTXX$, Len(TTXX$) - 1)
               GoTo 34
            End If
         End If
         Printer.CurrentX = H1 + 50
         Printer.CurrentY = V1 + 0.2 * HREN
         Printer.Print TTXX$
         '
         If FIECO(II% - 1).Visible = True Then
            H1 = COEF * FIECO(II% - 1).Left + MARGENI
            H2 = H1 + COEF * FIECO(II% - 1).Width
            If H2 > MARGENI + ATEXTO - 150 Then H2 = MARGENI + ATEXTO - 150
            ANCHOCUA = H2 - H1 - 100
            V1 = 4 * HREN + FIECO(II% - 1).Top
            V2 = V1 + HREN * 1.3
            If V2 > VFILMAX Then VFILMAX = V2
            Printer.Line (H1, V1)-(H2, V2), QBColor(0), B
            '
            TTXX$ = FIECO(II% - 1).Text
36          If Printer.TextWidth(TTXX$) > ANCHOCUA Then
               If Len(TTXX$) > 0 Then
                  TTXX$ = Left$(TTXX$, Len(TTXX$) - 1)
                  GoTo 36
               End If
            End If
            Printer.CurrentX = H1 + 50
            Printer.CurrentY = V1 + 0.2 * HREN
            Printer.Print TTXX$
         End If
         '
      End If
   Next II%
   '
   H1 = MARGENI: V1 = TOPETAB
   H2 = H1 + ATEXTO: V2 = (LARGOPAP - 25) * 56.7
   Printer.Line (H1, V1)-(H2, V2), QBColor(0), B
   '
   H1 = MARGENI: V1 = TOPETAB + 2 * HREN
   H2 = H1 + ATEXTO: V2 = V1
   Printer.Line (H1, V1)-(H2, V2), QBColor(0)
   POVERT = V2 - 0.5 * HREN
   '
   Printer.CurrentX = MARGENI + 100
   Printer.CurrentY = TOPETAB + 0.7 * HREN
   TTXX$ = ENCA$
   Printer.Print TTXX$
   '
   H1 = MARGENI: V1 = (LARGOPAP - 25) * 56.7 - 2 * HREN
   H2 = H1 + ATEXTO: V2 = V1
   Printer.Line (H1, V1)-(H2, V2), QBColor(0)
   '
   If CAHO% < 1 Then
      CAREHO% = (V1 - TOPETAB) / HREN
      CAHOJAS = CARESEL& / CAREHO%
      If CAHOJAS > Int(CAHOJAS) Then CAHOJAS = CAHOJAS + 1
      CAHO% = Int(CAHOJAS)
   End If
   '
   NUHOJ% = NUHOJ% + 1
   TTXX$ = "Hoja No. " + TRIM$(Str$(NUHOJ%)) + "/" + TRIM$(Str$(CAHO%))
   Printer.CurrentX = MARGENI + ATEXTO - 150 - Printer.TextWidth(TTXX$)
   Printer.CurrentY = 1.7 * HREN
   Printer.Print TTXX$
   H1 = MARGENI + ATEXTO - 300 - Printer.TextWidth(TTXX$)
   H2 = H1
   V1 = HREN: V2 = 3 * HREN
   Printer.DrawWidth = 2
   Printer.Line (H1, V1)-(H2, V2), QBColor(0)
   '
   H3 = H1
   HOII% = HOY(HOS$)
   TTXX$ = HOS$
   Printer.CurrentX = H3 - 150 - Printer.TextWidth(TTXX$)
   Printer.CurrentY = 1.7 * HREN
   Printer.Print TTXX$
   H1 = H3 - 300 - Printer.TextWidth(TTXX$)
   H2 = H1
   V1 = HREN: V2 = 3 * HREN
   Printer.DrawWidth = 2
   Printer.Line (H1, V1)-(H2, V2), QBColor(0)
   '
Return
'
LEPIEPA:
   '
   Printer.FontSize = 6
   TXT$ = "Sistema FLEXOFT s/n " + SERINUM$
   TXT$ = TXT$ + " licenciado a " + EMPRELI$
   TXT$ = TXT$ + " - Derechos Reservados L.11723"
   TXT$ = MIZQ$ + Space$(OFS%) + TXT$
   '
   While Printer.TextWidth(TXT$) < ATEXTO
      TXT$ = " " + TXT$ + " "
   Wend
   Printer.CurrentX = MARGENI
   Printer.CurrentY = (LARGOPAP - 25) * 56.7 + 100
   Printer.Print TXT$
   Printer.FontSize = TAFUESEL
   '
Return
'
End Sub


Sub GRABALI()
    '
    If BOTAGRE.Enabled = True Then GoTo 5
    If BOTINSE.Enabled = True Then GoTo 5
    If BOTERASE.Enabled = True Then GoTo 5
    If BOTEDIT.Enabled = True Then GoTo 5
    Exit Sub
    '
5   UU& = 0
    For u& = 1 To LITAB.ListCount
        '
        XXX$ = REGLEIDO$(ARCHILOC$, u&)
        For FT% = 1 To NFT%
            XXZ$ = FIENTRA(FT% - 1).Text
            XXY$ = CBIN$(XXZ$, TVL%(CFF%(FT%)), LF%(CFF%(FT%)))
            Call REPLA$(XXX$, XXY$, PR%(CFF%(FT%)), LF%(CFF%(FT%)))
        Next FT%
        '
        REGGRA& = 0
        UU& = UU& + 1
        If UU& <= LIREGSEL.ListCount Then
            LIREGSEL.ListIndex = UU& - 1
            REGGRA& = Val(LIREGSEL.Text)
        End If
        '
        If REGGRA& < 1 Or REGGRA& > ULTREG&(ARCHITRA$) Then
              Call REGAPP(ARCHITRA$, XXX$)
              REGGRA& = ULTREG&(ARCHITRA$)
              LIREGSEL.AddItem TRIM$(Str$(REGGRA&))
            Else
              Call GRAREG(ARCHITRA$, XXX$, REGGRA&)
        End If
        '
    Next u&
    '
    For u& = UU& + 1 To LIREGSEL.ListCount
        LIREGSEL.ListIndex = u& - 1
        REGGRA& = Val(LIREGSEL.Text)
        If REGGRA& > 0 Then
           If REGGRA& <= ULTREG&(ARCHITRA$) Then
              LORE% = RECLEN%(ARCHITRA$)
              XXX$ = String$(LORE%, 0)
              Call GRAREG(ARCHITRA$, XXX$, REGGRA&)
           End If
        End If
    Next u&
    '
10  Call CIERRARCH(ARCHITRA$)
    '
End Sub
'
Sub PRENTRA()
    '
    'If YAPRE% = 1 Then
    '    Exit Sub
    'End If
    ACAREN = LITAB.Height / CARELI%
    RENGLOAC = LITAB.ListIndex - LITAB.TopIndex + 1
    If RENGLOAC < 1 Or RENGLOAC > CARELI% Then
        Exit Sub
    End If
    YAPRE% = 1
    BORDESUPE = Top + LITAB.Top + RENGLOAC * ACAREN
    KPLA% = 0: PRICA% = 0
    '
    BLANCO = &H80000005
    For II% = 0 To 31
        '
        FORMEDIT.ENTRA(II%).FontSize = 7
        FORMEDIT.ENTRA(II%).FontBold = True
        FORMEDIT.ENTRA(II%).Text = ""
        FORMEDIT.ENTRA(II%).Height = 285
        FORMEDIT.ENTRA(II%).Top = 0
        FORMEDIT.ENTRA(II%).BackColor = BLANCO
        FORMEDIT.ENTRA(II%).Visible = False
        FORMEDIT.BOTSEL(II%).Visible = False
        '
    Next II%
    '
    ALTUCON% = 285
    FORMEDIT.Top = BORDESUPE
    FORMEDIT.Left = Left
    FORMEDIT.Width = Width
    FORMEDIT.Height = ALTUCON% * 1.2
    '
    If MODEDI% = 5 Then
        ' FORMEDIT.Top = Top + FRAMETOT.Top
        FORMEDIT.Top = Top + LITAB.Top + (CARELI% + 1) * ACAREN + (FRAMETOT.Height - FORMEDIT.Height) / 2
    End If
    '
    FORMEDIT.Command1.Left = FRAMEBOT2.Left
    FORMEDIT.Command1.Width = BOTAGRE.Width * 0.6
    FORMEDIT.Command2.Left = FORMEDIT.Command1.Left + FORMEDIT.Command1.Width
    FORMEDIT.Command2.Width = BOTAGRE.Width * 0.4
    FORMEDIT.Command1.Height = ALTUCON%
    FORMEDIT.Command2.Height = ALTUCON%
    '
    POSEDI% = 0
    For II% = 1 To NCL%
       '
       If CL%(II%) > 0 Then
          FORMEDIT.ENTRA(II% - 1).Left = COEFH * (CL%(II%) - 0.9)
          FORMEDIT.ENTRA(II% - 1).Top = 0
          FORMEDIT.ENTRA(II% - 1).Width = COEFH * (ACL%(II%) + 0.9)
          FORMEDIT.ENTRA(II% - 1).Visible = True
          FORMEDIT.ENTRA(II% - 1).FontName = LITAB.FontName
          FORMEDIT.ENTRA(II% - 1).FontSize = LITAB.FontSize
          FORMEDIT.ENTRA(II% - 1).FontBold = LITAB.FontBold
          If PE%(II%) <> 1 Then
            If PE%(II%) = 3 Or MODEDI% = 4 Then
              FORMEDIT.ENTRA(II% - 1).BackColor = &H80000013
            End If
          End If
       End If
       '
       If FORMEDIT.ENTRA(II% - 1).BackColor = BLANCO Then
           POSEDI% = 1
       End If
       '
    Next II%
    '
    If MODEDI% = 4 Then                     ' editar registro
        RELIB$ = LITAB.Text
        XXX$ = REGLEIDO(ARCHILOC$, LITAB.ListIndex + 1)
        For II% = 1 To NCL%
           If CL%(II%) > 0 Then
              FORMEDIT.ENTRA(II% - 1).Text = Mid$(RELIB$, CL%(II%), ACL%(II%))
              If TVL%(CFT%(II%)) <= 4 Then
                  If Left$(TRIM$(FORMEDIT.ENTRA(II% - 1)), 1) = "*" Then
                      XXZ$ = Mid$(XXX$, PR%(CFT%(II%)), LF%(CFT%(II%)))
                      FORMEDIT.ENTRA(II% - 1).Text = TRIM$(CSTRING$(XXZ$, TVL%(CFT%(II%)), 15, LDL%(CFT%(II%)), 1))
                  End If
              End If
           End If
        Next II%
    End If
    '
    If POSEDI% = 1 Or MODEDI% < 4 Then
        FORMEDIT.Refresh
        On Error Resume Next
        COLUAC% = 0
        FORMEDIT.Show 1
        On Error GoTo 0
      Else
        BOTEDIT.Enabled = False
    End If
    '
End Sub
'
Sub SETUPRINTER()
    '
    ' WLISTAB.Hide
    '
    NOMIMPRE$ = TRIM$(Printer.DeviceName)
    If NOMIMPRE$ = "" Then
        Call MENSERR(24, "Falta Seleccionar Impresora\en la Configuración de Windows")
        Call FINAL("")
    End If
    '
    YAESTA% = 0
    Call BLANARCH("!PRINTERS")
    NOMI$ = AJUSTI$(NOMIMPRE$, 48)
    NX% = FILEOPEN("C:\FLEXOFT\PRINTERS.DRV", 0, 128)
    For IPRIN& = 1 To LOF(NX%) / 128
        Get #NX%, IPRIN&, PPZZ$
        PPXX$ = PPZZ$
        If TRIM$(Left$(PPXX$, 48)) = NOMIMPRE$ Then
            Call REGAPP("!PRINTERS", PPXX$)
            YAESTA% = 1
            GoTo 10
        End If
    Next IPRIN&
    '
10  Close #NX%
    If YAESTA% = 1 Then
        PPXX$ = Left$(PPXX$, 52)
        'GoTo 50
    End If
    '
    Printer.ScaleMode = 1
    ANCHOJMM = Int(Printer.ScaleWidth / 56.7)
    Call REPLA(PPXX$, MKS$(ANCHOJMM), 49, 4)
    '
    CAFUFIJ% = 0
    For I% = 1 To Printer.FontCount
        Printer.FontName = Printer.Fonts(I% - 1)
        A1 = 0: A2 = 0
        On Error Resume Next
        A1 = Printer.TextWidth(String$(80, "l"))
        A2 = Printer.TextWidth(String$(80, "M"))
        On Error GoTo 0
        If A1 > 0 And A2 > 0 Then
          If A1 = A2 Then
            If CAFUFIJ% < 128 Then
              CAFUFIJ% = CAFUFIJ% + 1
              FUFIJA$(CAFUFIJ%) = Printer.Fonts(I% - 1)
            End If
          End If
        End If
    Next I%
    '
    If CAFUFIJ% < 1 Then
       Call MENSERR(24, "No se Encontraron Fuentes de Ancho Fijo\para la Impresora Seleccionada.\  \No es Posible Emitir Listado por Impresora.\  \Consulte a su Soporte de Sistemas.")
       Call FINAL("")
    End If
    '
    Call BLANARCH("!PRINTERS")
    TALEBASI = 24
    'TEPRUEBA$ = String$(128, "*")
    Printer.FontBold = False
    Printer.FontItalic = False
    For I% = 1 To CAFUFIJ%
        Printer.FontName = FUFIJA$(I%)
        If Printer.FontName <> FUFIJA$(I%) Then GoTo 19
        '
        Printer.FontSize = 24
        FUENTEMAX = Printer.FontSize
        Printer.FontSize = 2
        FUENTEMIN = Printer.FontSize
        '
14      If FUENTEMIN <= FUENTEMAX Then
            XX$ = REGBLAN$("!PRINTERS")
            Call REPLA(XX$, NOMIMPRE$, 1, 48)
            Call REPLA(XX$, MKS$(ANCHOJMM), 49, 4)
            Call REPLA(XX$, FUFIJA$(I%), 65, 48)
            Call REPLA(XX$, MKS$(FUENTEMIN), 113, 4)
            Call REPLA(XX$, MKS$(FUENTEMAX), 117, 4)
            Call REPLA(XX$, "SI", 125, 2)
            Call REGAPP("!PRINTERS", XX$)
        End If
        '
19  Next I%
    Call CIERRARCH("!PRINTERS")
    '
    PPXX$ = AJUSTI$(NOMIMPRE$, 48) + MKS$(ANCHOJMM) + String$(12, 0)
    '
50  PPYY$ = OBJPLA$("DATIMPRE", PPXX$)
    If PPYY$ = "" Then
        Call CIERRARCH("PRINTERS")
        GoTo 99
    End If
    '
    ANCHOJMM = CVS(Mid$(PPYY$, 49, 8))
    If ANCHOJMM > Int(Printer.ScaleWidth / 56.7) Or ANCHOJMM < 48 Then
        PPXX$ = PPYY$
        GoTo 50
    End If
    '
    VTB% = VENTABU%("PRINSET/" + AJUSTI$(NOMIMPRE$, 48))
    If VTB% >= 0 Then
        NX% = FILEOPEN%("C:\FLEXOFT\PRINTERS.DRV", 0, 128)
        NOMI$ = AJUSTI$(NOMIMPRE$, 48)
        JJ& = 0
        For IPRIN& = 1 To LOF(NX%) / 128
            Get #NX%, IPRIN&, PPZZ$
            If Left$(PPZZ$, 48) <> NOMI$ Then
               JJ& = JJ& + 1
               Put #NX%, JJ&, PPZZ$
            End If
        Next IPRIN&
        For IPRIN& = 1 To ULTREG&("!PRINTERS")
            PPXX$ = REGLEIDO$("!PRINTERS", IPRIN&)
            Call REPLA(PPXX$, NOMI$, 1, 48)
            Call REPLA(PPXX$, MKS$(ANCHOJMM), 49, 4)
            If UCase$(Mid$(PPXX$, 125, 2)) <> "SI" Then
                Call REPLA(PPXX$, "  ", 125, 2)
            End If
            PPZZ$ = PPXX$
            JJ& = JJ& + 1
            Put #NX%, JJ&, PPZZ$
        Next IPRIN&
        For IPRIN& = JJ& + 1 To LOF(NX%) / 128
            PPZZ$ = String$(128, 0)
            Put #NX%, IPRIN&, PPZZ$
        Next IPRIN&
        Close #NX%
    End If
    Call CIERRARCH("!PRINTERS")
    '
99  FORMA$ = FORMALI$
    If TRIM$(FORMA$) <> "" Then
        Call CARDATLIS
        Call CARDAIMP
      Else
        Call COMUNI("Para Efectivizar la nueva Configuración\Debe Volver a pedir el Listado.")
        Call FINAL("")
    End If
    '
End Sub
'
Sub CARDATLIS()
     '
5800 GRABACT% = 0
     TOTALIX% = 0
     POSITOT% = 0
     '
     FORMA$ = Label13
     Dim YYY As String * 128
     If N15% > 0 Then Close #N15%
     N15% = FILEOPEN%(FORMA$, 9, 128)
     Get #N15%, 1, YYY$
     '
     If LOF(N15%) <= 128 Then
         On Error Resume Next
         YYY$ = String$(128, 0)
         Put #N15%, 1, YYY$
         Close #N15%
         Kill FORMA$
         On Error GoTo 0
         FORMB$ = UCase$(REPLACAR$(FORMA$, "\", "/"))
         Call MENSERR(24, "Proceso Interrumpido.\Falta Controlador de Procesos '" + FORMB$ + "'.")
         Call FINAL("")
     End If
     '
     Archi$ = Mid$(YYY$, 57, 8)
     ARCHITRA$ = TRIM$(Archi$)
     Label5 = ARCHITRA$
     Label6 = TITUARCH$(ARCHITRA$)
     If Text2.Enabled Then Text2 = ARCHITRA$
5801 Get #N15%, 1, YYY$
     LOR1% = Asc(Mid$(YYY$, 106, 1)): If LOR1% > 16 Then LOR1% = 0
     LOR% = 256 * LOR1% + Asc(Mid$(YYY$, 107))
     CAMAX& = CVI(Mid$(YYY$, 108))
     PEOF% = 1 + Asc(Mid$(YYY$, 110))
     EOFS$ = Mid$(YYY$, 111, 16)
     Label8 = PROPRIN$
5803 S1$ = TRIM$(Left$(YYY$, 48))
     Label9 = TRIM$(S1$)
     '
     EPAC$ = TRIM$(EMPREAC$)
     If EPAC$ <> "" Then
         Caption = Caption + "  -  " + EPAC$
     End If
    '
     Sali$ = Mid$(YYY$, 49, 8)
     ARCHISAL$ = Mid$(YYY$, 65, 12)
     If Left$(ARCHISAL$, 3) = "LPT" Then ARCHISAL$ = ""
     COD$ = Mid$(YYY$, 81)
     RGI& = Asc(Mid$(COD$, 1))
     RGF& = Asc(Mid$(COD$, 2))
     OFS% = Asc(Mid$(COD$, 3))
5804 BLP% = Asc(Mid$(COD$, 4))
        CXC$ = "": If BLP% = 1 Then CXC$ = "¦"
     VIV% = Asc(Mid$(COD$, 5))
     APQ% = Asc(Mid$(COD$, 6))
     IDL% = REHOJ%: REHOJ1% = REHOJ%
     If Asc(Mid$(COD$, 7)) < REHOJ% Then IDL% = Asc(Mid$(COD$, 7))
     '
     If VIV% = 1 Then IDL% = 8 * IDL% / 6
     REHOJX% = 0
     If Asc(Mid$(COD$, 7)) > 96 Then
       REHOJX% = Asc(Mid$(COD$, 7))
     End If
     SUPRIFE% = 0
     HDR& = Asc(Mid$(COD$, 8))
     If HDR& = 99 Then
        HDR& = 0
        SUPRIFE% = 1
     End If
     RFT% = Asc(Mid$(COD$, 9))
     NCL% = 0: TCL% = 0: TNC% = 0
     RG% = 1: HST% = 0
     HOJINI% = CVI(Mid$(COD$, 10, 2))
5805 RG% = RG% + 1
     If RG% > LOF(N15%) / 128 Or RG% > 33 Then GoTo 5820
5806 Get #N15%, RG%, YYY$
     If YYY$ = String$(128, 0) Then GoTo 5820
     '
5807 TCL% = TCL% + 1
     TVL%(TCL%) = Asc(Mid$(YYY$, 33))
         If RG% = 2 Then TVL%(TCL%) = 1
     LDL%(TCL%) = Asc(Mid$(YYY$, 34))
5808 PR%(TCL%) = Asc(Mid$(YYY$, 35))
     While PR%(TCL%) < PR%(TCL% - 1)
       PR%(TCL%) = PR%(TCL%) + 256
     Wend
     LF%(TCL%) = Asc(Mid$(YYY$, 36))
         If RG% = 2 Then PR%(TCL%) = 0: LF%(TCL%) = 2
         '
     If Asc(Mid$(YYY$, 72)) = 99 Then
        TNC% = TNC% + 1
        CLNC%(TNC%) = TCL%
     End If
     '
5811 If Asc(Mid$(YYY$, 69)) < 1 Then GoTo 5817   ' SUPRIME COLUMNAS LONGITUD 0
5812 NCL% = NCL% + 1
     CFT%(NCL%) = RG% - 1
     CL0%(NCL%) = RG% - 1
     ECL$(NCL%) = TRIM$(Mid$(YYY$, 37, 32))
5813 ACL%(NCL%) = Asc(Mid$(YYY$, 69))
     LK%(NCL%) = Asc(Mid$(YYY$, 70))
5814 MG%(NCL%) = Asc(Mid$(YYY$, 71))
     TT%(NCL%) = Asc(Mid$(YYY$, 73))
        If TT%(NCL%) = 1 Then TT%(NCL%) = TCL%
        If TT%(NCL%) = 255 Then HST% = 1
5815 PE%(NCL%) = 1
        If RG% = 2 Then PE%(NCL%) = 2
5816 JT%(NCL%) = Asc(Mid$(YYY$, 72))
     AE$(NCL%) = Mid$(YYY$, 74, 8)
     CE%(NCL%) = Asc(Mid$(YYY$, 82))
     If TT%(NCL%) > 0 Then
        NTT% = NTT% + 1
        TOTALIX% = 1
     End If
     '
     If POSITOT% = 0 Then
        If TT%(NCL%) < 1 Then
           If TVL%(TCL%) = 6 Then
              If ACL%(NCL%) >= 10 Then
                 POSITOT% = NCL%
              End If
           End If
        End If
     End If
     
5817 GoTo 5805
     '
5820 If POSITOT% = 0 Then POSITOT% = 1
     ATOT% = BLP% * (NCL% + 1)
     CLII% = OFS% + 2
     For KCL% = 1 To NCL%
          LC%(KCL%) = ACL%(KCL%)
          ATOT% = ATOT% + ACL%(KCL%) + MG%(KCL%)
          CL%(KCL%) = CLII%
          CLII% = CLII% + ACL%(KCL%) + 1
     Next KCL%
5821 For KCL% = 1 To NCL%
          If CE%(KCL%) <= 0 Then GoTo 5824
5822      For KCL1% = 1 To NCL%
              If CL0%(KCL1%) = CE%(KCL%) Then CE%(KCL%) = KCL1%: PE%(KCL1%) = 3: GoTo 5824
5823      Next KCL1%
5824 Next KCL%
5825 For KCL% = 1 To NCL%
          If TT%(KCL%) <= 0 Then GoTo 5828
5826      For KCL1% = 1 To NCL%
             If CL0%(KCL1%) = TT%(KCL%) Then TT%(KCL%) = KCL1%: GoTo 5828
             If 100 + CL0%(KCL1%) = TT%(KCL%) Then TT%(KCL%) = 100 + KCL1%: GoTo 5828
5827      Next KCL1%
5828 Next KCL%
     '
5830 NFT% = 0    ' NO HAY FILTROS
5850 If N15% > 0 Then
         Close #N15%
         N15% = 0
     End If
     '
End Sub
'
Sub CARDAIMP()
    '
    NOMIMPRE$ = TRIM$(Printer.DeviceName)
    Printer.Orientation = 1
    Printer.ScaleMode = 1
    ANCHOPAP = Int(Printer.ScaleWidth / 56.7)
    LARGOPAP = Int(Printer.ScaleHeight / 56.7)
    MARGENI = 0
    MARGENS = 0
    ORIENTA% = 1
    CACOPIAS% = 1
    '
    Dim PPZZ As String * 128
    NX% = FILEOPEN("C:\FLEXOFT\LISTSET.DRV", 0, 128)
    FIN& = LOF(NX%) / 128
    For IPRIN& = 1 To FIN&
         Get #NX%, IPRIN&, PPZZ$
         PPXX$ = PPZZ$
         If TRIM$(Left$(PPXX$, 48)) = NOMIMPRE$ Then
           If TRIM$(UCase$(Mid$(PPXX$, 57, 8))) = TRIM$(UCase$(PROPRIN$)) Then
             ANCHOPAP1 = CVS(Mid$(PPXX$, 49, 4))
             LARGOPAP1 = CVS(Mid$(PPXX$, 53, 4))
             '
             If ANCHOPAP1 > 0 Then
                If ANCHOPAP1 < ANCHOPAP Then
                   ANCHOPAP = ANCHOPAP1
                End If
             End If
             '
             If LARGOPAP1 > 0 Then
                If LARGOPAP1 < LARGOPAP Then
                   LARGOPAP = LARGOPAP1
                End If
             End If
             '
             GoTo 10
             '
           End If
         End If
    Next IPRIN&
    '
10  Close #NX%
    NX% = FILEOPEN%("C:\FLEXOFT\LISTABU.DRV", 0, 128)
    FIN& = LOF(NX%) / 128
    For IPRIN& = 1 To FIN&
         Get #NX%, IPRIN&, PPZZ$
         PPXX$ = PPZZ$
         If TRIM$(Left$(PPXX$, 8)) = TRIM$(PROPRIN$) Then
             MARGENI1 = CVS(Mid$(PPXX$, 9, 4))
             MARGENS1 = CVS(Mid$(PPXX$, 13, 4))
             ORIENTA1% = Asc(Mid$(PPXX$, 17, 1))
             CACOPIAS1% = Asc(Mid$(PPXX$, 18, 1))
             '
             If MARGENI1 > MARGENI Then
                 MARGENI = MARGENI1
             End If
             '
             If MARGENS1 > MARGENS Then
                   MARGENS = MARGENS1
             End If
             '
             If ORIENTA1% > 0 Then
                If ORIENTA1% <= 2 Then
                    ORIENTA% = ORIENTA1%
                End If
             End If
             '
             If CACOPIAS1% > 0 Then
                If CACOPIAS1% <= 32 Then
                    CACOPIAS% = CACOPIAS1%
                End If
             End If
             '
             GoTo 20
             '
         End If
    Next IPRIN&
    '
20  End Sub


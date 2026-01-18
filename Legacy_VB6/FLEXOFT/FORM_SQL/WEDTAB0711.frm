VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form WEDTAB07 
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
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   184
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   105
      Top             =   5565
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8505
      Top             =   210
   End
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
         Picture         =   "WEDTAB0711.frx":0000
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
         Picture         =   "WEDTAB0711.frx":030A
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
         Picture         =   "WEDTAB0711.frx":0614
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
         Picture         =   "WEDTAB0711.frx":091E
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
         Picture         =   "WEDTAB0711.frx":0C28
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
         Picture         =   "WEDTAB0711.frx":0F32
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
         Picture         =   "WEDTAB0711.frx":123C
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
         Picture         =   "WEDTAB0711.frx":1546
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
         Picture         =   "WEDTAB0711.frx":1850
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
         Picture         =   "WEDTAB0711.frx":1B5A
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
         Picture         =   "WEDTAB0711.frx":1E64
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
         Picture         =   "WEDTAB0711.frx":216E
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
         Picture         =   "WEDTAB0711.frx":2478
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
         Picture         =   "WEDTAB0711.frx":2782
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
         Picture         =   "WEDTAB0711.frx":2A8C
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
         Picture         =   "WEDTAB0711.frx":2D96
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
         Picture         =   "WEDTAB0711.frx":30A0
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
         Picture         =   "WEDTAB0711.frx":33AA
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
         Picture         =   "WEDTAB0711.frx":36B4
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
         Picture         =   "WEDTAB0711.frx":39BE
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
         Picture         =   "WEDTAB0711.frx":3CC8
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
         Picture         =   "WEDTAB0711.frx":3FD2
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
         Picture         =   "WEDTAB0711.frx":42DC
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
         Picture         =   "WEDTAB0711.frx":45E6
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
         Picture         =   "WEDTAB0711.frx":48F0
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
         Picture         =   "WEDTAB0711.frx":4BFA
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
         Picture         =   "WEDTAB0711.frx":4F04
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
         Picture         =   "WEDTAB0711.frx":520E
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
         Picture         =   "WEDTAB0711.frx":5518
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
         Picture         =   "WEDTAB0711.frx":5822
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
         Picture         =   "WEDTAB0711.frx":5B2C
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
         Picture         =   "WEDTAB0711.frx":5E36
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
         Picture         =   "WEDTAB0711.frx":6140
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
         Picture         =   "WEDTAB0711.frx":628A
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
         Picture         =   "WEDTAB0711.frx":6594
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
         Picture         =   "WEDTAB0711.frx":6BFE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Grabar - Guardar Modificaciones"
         Top             =   945
         Width           =   700
      End
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
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "WEDTAB0711.frx":6F08
      TabIndex        =   185
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "WEDTAB0711.frx":6F82
      Top             =   1200
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   630
      Index           =   1
      Left            =   7680
      OleObjectBlob   =   "WEDTAB0711.frx":71B6
      TabIndex        =   186
      Top             =   6720
      Visible         =   0   'False
      Width           =   1350
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
      Picture         =   "WEDTAB0711.frx":7230
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
Attribute VB_Name = "WEDTAB07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
       Call CIERRARCH("*.*")
       Call FINAL("")
    End If
    '
    WEDTAB07.LITAB.Clear
    Call VERISCROLL
    If TVL%(CFF%(Index + 1)) = 5 Then
         FIENTRA(Index).TEXT = "."
         Call SELECHO("SELFECHA")
         FIENTRA(Index).TEXT = VALACT1$("SELFECHA")
         IANT% = Index
         Exit Sub
    End If
    '
    If TRIM$(AEF$(Index + 1)) <> "" Then
         FIENTRA(Index).TEXT = "."
         Call SELECHO(AEF$(Index + 1))
         FIENTRA(Index).TEXT = VALACT1$(AEF$(Index + 1))
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
    If Len(FIENTRA(Index).TEXT) >= LOMACA% Then
        FIENTRA(Index).TEXT = Left$(FIENTRA(Index).TEXT, LOMACA%)
        FIENTRA(Index).SelStart = 9999
    End If
    '
    If AEF$(Index + 1) <> "" Then
       If CLF%(Index + 1) > 0 Then
          If LEF%(Index + 1) > 0 Then
             VDEC$ = ""
             If TRIM$(UCase$(AEF$(Index + 1))) <> "MASTER" Then
                XXZ$ = CBIN$(FIENTRA(Index).TEXT, TVL%(CFF%(Index + 1)), LF%(CFF%(Index + 1)))
                VDEC$ = ECOARCH$(AEF$(Index + 1), XXZ$)
               Else
                CI% = CODINT%(FIENTRA(Index).TEXT)
                VDEC$ = DESCRIT$(CI%)
             End If
             '
             WEDTAB07.FIECO(Index).TEXT = VDEC$
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
         FIENTRA(Index).TEXT = "."
         Call SELECHO("SELFECHA")
         FIENTRA(Index).TEXT = VALACT1$("SELFECHA")
         IANT% = Index
         Exit Sub
    End If
    '
    If TRIM$(AEF$(Index + 1)) <> "" Then
         FIENTRA(Index).TEXT = "."
         Call SELECHO(AEF$(Index + 1))
         FIENTRA(Index).TEXT = VALACT1$(AEF$(Index + 1))
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
    Call CIERRARCH("*.*")
    If ARCHITRA$ = "COKARDEX" Then
       Call CIERRARCH("*.*")
       Call FINAL("")
    End If
    '
    If MODIFI% > 0 Then
        RESPU = MsgBox("Abandona sin Guardar los Cambios ?", 4371, "Operacion Critica")
        If RESPU <> 6 Then
            BOTREC.SetFocus
            Exit Sub
        End If
        MODIFI% = 0
    End If
    '
    If NFT% < 1 Then
          Call CIERRARCH("*.*")
          If COMMA$ <> "" Then
             Call BLOQEXE("*")
          End If
          WEDTAB07.Hide
          Exit Sub
          '
        Else    ' hay filtros
          '
          If NFT% = 1 Then
             If LKF%(1) < 1 Then
                Unload WEDTAB04
             End If
          End If
          '
          If WEDTAB07.LITAB.ListCount < 1 Then
              '
              Call CIERRARCH("*.*")
              If COMMA$ <> "" Then
                 Call BLOQEXE("*")
              End If
              WEDTAB07.Hide
              Exit Sub
              '
            Else
              WEDTAB07.LITAB.Clear
              Call VERISCROLL
              For K2% = 1 To 32
                 TTX#(K2%) = 0
                 TXX#(K2%) = 0
              Next K2%
              '
              Call CLEARFORM
              YAFIL% = 0
              Call ENTRAFIL
              Exit Sub
              '
          End If
          '
    End If
    '
End Sub
'

Private Sub Form_Activate()
    '
    On Error Resume Next
    SetFocus
    Refresh
    On Error GoTo 0
    '
    If YAFIL% = (-1) Then
       Call FINAL("")
    End If
    Call ENTRAFIL
    If MODEDI% = 1 Then
        If YAPRE% = 0 Then
            Call ADDLISTA
        End If
    End If
    '
End Sub
'


Private Sub Form_Click()
   DoEvents
End Sub

Private Sub Form_Load()
    '
    YAFIL% = 0
    Enabled = True
    'Call APLICACIONSKINS(Me)
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Screen.MousePointer = 11
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Screen.MousePointer = 1
   Hide
   'Call FINAL("")
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub LITAB_DblClick()
    Call EDLISTA
End Sub

Private Sub Timer2_Timer()
'   On Error Resume Next
'   DoEvents
'   AppActivate Caption, True
'   Call Form_Click
'   On Error GoTo 0
End Sub

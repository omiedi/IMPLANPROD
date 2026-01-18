VERSION 5.00
Begin VB.Form INFODEM 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Información"
   ClientHeight    =   2085
   ClientLeft      =   2550
   ClientTop       =   300
   ClientWidth     =   3075
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   3075
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2310
      Top             =   1575
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFC0&
      Height          =   1380
      Left            =   0
      ScaleHeight     =   1320
      ScaleWidth      =   3315
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3375
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
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
         Left            =   100
         TabIndex        =   1
         Top             =   30
         Visible         =   0   'False
         Width           =   3060
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   195
      Left            =   735
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   870
   End
End
Attribute VB_Name = "INFODEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

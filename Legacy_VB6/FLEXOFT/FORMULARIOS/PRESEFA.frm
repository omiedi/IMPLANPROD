VERSION 5.00
Begin VB.Form PRESEFA 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10485
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10515
      Begin VB.Frame Frame4 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   540
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   10515
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   $"PRESEFA.frx":0000
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   330
            Left            =   0
            TabIndex        =   3
            Top             =   105
            Width           =   10410
         End
         Begin VB.Line Line1 
            X1              =   840
            X2              =   840
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Line Line4 
            X1              =   2310
            X2              =   2310
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Line Line5 
            X1              =   7560
            X2              =   7560
            Y1              =   0
            Y2              =   1050
         End
         Begin VB.Line Line6 
            X1              =   8820
            X2              =   8820
            Y1              =   0
            Y2              =   1050
         End
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4560
         Left            =   0
         TabIndex        =   1
         Top             =   525
         Width           =   10515
      End
      Begin VB.Line Line9 
         X1              =   0
         X2              =   10500
         Y1              =   1590
         Y2              =   1590
      End
      Begin VB.Line Line2 
         Index           =   0
         X1              =   840
         X2              =   840
         Y1              =   525
         Y2              =   5565
      End
      Begin VB.Line Line2 
         Index           =   1
         X1              =   2310
         X2              =   2310
         Y1              =   525
         Y2              =   5565
      End
      Begin VB.Line Line2 
         Index           =   2
         X1              =   7560
         X2              =   7560
         Y1              =   525
         Y2              =   5565
      End
      Begin VB.Line Line2 
         Index           =   3
         X1              =   8820
         X2              =   8820
         Y1              =   525
         Y2              =   5565
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
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
      Left            =   840
      TabIndex        =   7
      Top             =   5280
      Width           =   4950
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "O/D's:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   6
      Top             =   5250
      Width           =   645
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Importe Total Neto:"
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
      Left            =   5355
      TabIndex        =   5
      Top             =   5460
      Width           =   2745
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8295
      TabIndex        =   4
      Top             =   5355
      Width           =   1905
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   10500
      Y1              =   5060
      Y2              =   5060
   End
End
Attribute VB_Name = "PRESEFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

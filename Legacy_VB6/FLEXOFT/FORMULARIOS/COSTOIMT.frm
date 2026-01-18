VERSION 5.00
Begin VB.Form COSTOIMT 
   BackColor       =   &H00F0F0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "S.I.T. - Complementos, Consultas, Listados"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00F0F0E0&
      Caption         =   "Control de Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   150
      TabIndex        =   3
      Top             =   210
      Width           =   4215
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1050
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   10
         Top             =   1470
         Width           =   2640
         Begin VB.CommandButton Command4 
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
            Picture         =   "COSTOIMT.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Listado de Control Importacion Temporaria"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de Control por Factura Exportacion"
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
            Left            =   720
            TabIndex        =   12
            Top             =   100
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   4
         Top             =   960
         Width           =   2640
         Begin VB.CommandButton Command2 
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
            Picture         =   "COSTOIMT.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Listado de Saldos de Importacion Temporaria"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de Saldos de Importacion Tempor."
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
            Left            =   735
            TabIndex        =   6
            Top             =   45
            Width           =   1905
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   7
         Top             =   420
         Width           =   2640
         Begin VB.CommandButton Command3 
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
            Picture         =   "COSTOIMT.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Listado de Control Importacion Temporaria"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Listado de Control  por Permiso ITemp."
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
            Left            =   735
            TabIndex        =   9
            Top             =   50
            Width           =   1905
         End
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
      Height          =   720
      Left            =   4620
      Picture         =   "COSTOIMT.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cierra - Abandona la Aplicación"
      Top             =   315
      Width           =   720
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4620
      ScaleHeight     =   75
      ScaleWidth      =   660
      TabIndex        =   1
      Top             =   1155
      Width           =   720
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   2
         Top             =   105
         Width           =   12000
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   4095
      Picture         =   "COSTOIMT.frx":0A68
      Top             =   1995
      Width           =   2010
   End
End
Attribute VB_Name = "COSTOIMT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call FINAL("*CONSIMPT")
   Command3.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call FINAL("*SALIMPT")
   Command2.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   JJ& = 0
   REBLA$ = REGBLAN$("FACIMPT")
   For U& = 1 To ULTREG&("CONSIMPT")
     XX$ = REGLEIDO$("CONSIMPT", U&)
     YY$ = REBLA$
     Call REPLA(YY$, Mid$(XX$, 129, 4), 1, 4)
     Call REPLA(YY$, Mid$(XX$, 157, 16), 29, 16)
     Call REPLA(YY$, MKS$(0), 45, 4)
     Call REPLA(YY$, Mid$(XX$, 183, 2), 55, 2)
     Call REPLA(YY$, Mid$(XX$, 193, 6), 65, 6)
     Call REPLA(YY$, Mid$(XX$, 33, 42), 71, 42)
     Call REPLA(YY$, Mid$(XX$, 185, 8), 129, 8)
     Call REPLA(YY$, Mid$(XX$, 215, 16), 145, 16)
     JJ& = JJ& + 1
     Call GRAREG("FACIMPT", YY$, JJ&)
   Next U&
   Call SETULTREG("FACIMPT", JJ&)
   Call FILESORT("FACIMPT", "", 1, 9)
   '
   Call FINAL("*FACIMPT")
   Command4.Enabled = True

End Sub


VERSION 5.00
Begin VB.Form ELECIMPOR 
   Caption         =   "Importacion"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione la Importación"
      Height          =   5160
      Left            =   210
      TabIndex        =   0
      Top             =   105
      Width           =   4950
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFFF&
         Height          =   750
         Left            =   3255
         TabIndex        =   19
         Top             =   4095
         Width           =   1275
         Begin VB.CommandButton Command10 
            Caption         =   "Movimientos Contables"
            Height          =   435
            Left            =   105
            MaskColor       =   &H8000000F&
            TabIndex        =   20
            Top             =   210
            Width           =   1065
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0C0FF&
         Height          =   2115
         Left            =   1785
         TabIndex        =   17
         Top             =   1890
         Width           =   1275
         Begin VB.CommandButton Command9 
            Caption         =   "Recuperar Prefijo de Facturas"
            Enabled         =   0   'False
            Height          =   1800
            Left            =   150
            TabIndex        =   18
            Top             =   210
            Width           =   960
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FF80FF&
         Height          =   960
         Left            =   1785
         TabIndex        =   15
         Top             =   4095
         Width           =   1275
         Begin VB.CommandButton Command8 
            Caption         =   "Cuenta Corriente Proveedores"
            Height          =   645
            Left            =   105
            TabIndex        =   16
            Top             =   210
            Width           =   1065
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C00000&
         Height          =   1170
         Left            =   3255
         TabIndex        =   13
         Top             =   2835
         Width           =   1275
         Begin VB.CommandButton Command7 
            Caption         =   "Composicion de Deuda"
            Enabled         =   0   'False
            Height          =   855
            Left            =   105
            MaskColor       =   &H8000000F&
            TabIndex        =   14
            Top             =   210
            Width           =   1065
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00800080&
         Height          =   1170
         Left            =   210
         TabIndex        =   11
         Top             =   2835
         Width           =   1275
         Begin VB.CommandButton Command6 
            Caption         =   "Cuenta Corriente Clientes"
            Height          =   855
            Left            =   105
            MaskColor       =   &H8000000F&
            TabIndex        =   12
            Top             =   210
            Width           =   1065
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00000080&
         Height          =   750
         Left            =   1785
         TabIndex        =   9
         Top             =   1050
         Width           =   1275
         Begin VB.CommandButton Command5 
            Caption         =   "Plan de Cuentas"
            Enabled         =   0   'False
            Height          =   435
            Left            =   105
            TabIndex        =   10
            Top             =   210
            Width           =   1065
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00008080&
         Height          =   750
         Left            =   3255
         TabIndex        =   7
         Top             =   1890
         Width           =   1275
         Begin VB.CommandButton Command4 
            Caption         =   "Movimientos de Stock"
            Height          =   435
            Left            =   105
            MaskColor       =   &H8000000F&
            TabIndex        =   8
            Top             =   210
            Width           =   1065
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00808000&
         Height          =   750
         Left            =   210
         TabIndex        =   5
         Top             =   1890
         Width           =   1275
         Begin VB.CommandButton Command3 
            Caption         =   "Clientes"
            Height          =   435
            Left            =   105
            MaskColor       =   &H8000000F&
            TabIndex        =   6
            Top             =   210
            Width           =   1065
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         Height          =   750
         Left            =   210
         TabIndex        =   3
         Top             =   210
         Width           =   1275
         Begin VB.CommandButton Command2 
            Caption         =   "Proveedores"
            Height          =   435
            Left            =   105
            TabIndex        =   4
            Top             =   210
            Width           =   1065
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFE0&
         Height          =   750
         Left            =   3255
         TabIndex        =   1
         Top             =   210
         Width           =   1275
         Begin VB.CommandButton Command1 
            Caption         =   "Maestro"
            Enabled         =   0   'False
            Height          =   435
            Left            =   105
            MaskColor       =   &H8000000F&
            TabIndex        =   2
            Top             =   210
            Width           =   1065
         End
      End
   End
End
Attribute VB_Name = "ELECIMPOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Enabled = False
   IMPORMASTER.Show
Command1.Enabled = True
End Sub

Private Sub Command10_Click()
Command10.Enabled = False
    IMPMOCON.Show
Command10.Enabled = True
End Sub

Private Sub Command2_Click()
Command2.Enabled = False
   IMPROVE.Show
Command2.Enabled = True
End Sub

Private Sub Command3_Click()
Command3.Enabled = False
IMPORCLIE.Show
Command3.Enabled = True
End Sub

Private Sub Command4_Click()
Command4.Enabled = False
IMPOMOSTK.Show
Command4.Enabled = True
End Sub

Private Sub Command5_Click()
Command5.Enabled = False
  IMPORCUE.Show
Command5.Enabled = True
End Sub

Private Sub Command6_Click()  
Command6.Enabled = False
    IMPCUECLI.Show 1
Command6.Enabled = True
End Sub

Private Sub Command7_Click()
Command7.Enabled = False
    IMPOPECO.Show
Command7.Enabled = True
End Sub

Private Sub Command8_Click()
Command8.Enabled = False
     IMPCUEPRO.Show
Command8.Enabled = True
End Sub

Private Sub Command9_Click()
Command9.Enabled = False
IMPREFAC.Show
Command9.Enabled = True
End Sub

Private Sub Form_Load()
    '//////////********///////////
    'LOS COMAND BUTON  1,4,5,7 ESTAN CON ENABLED=FALSE XQ
    ' NO LOS NECESITAMOS PARA LAS IMPORTACIONES A REALIZAR EN
    'GUIDI EL 01/11/04
    '///////////*******////////////
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    Call CENTRAFORM(Me)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

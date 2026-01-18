VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Items de Stock"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10605
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
   ScaleHeight     =   7260
   ScaleWidth      =   10605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFF80&
      Caption         =   "Modificacion de Datos "
      Enabled         =   0   'False
      Height          =   4530
      Left            =   5850
      TabIndex        =   23
      Top             =   2415
      Visible         =   0   'False
      Width           =   4305
      Begin VB.CommandButton Command20 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1470
         TabIndex        =   28
         Top             =   3465
         Width           =   1380
      End
      Begin VB.CommandButton Command19 
         Caption         =   "      Secuencia de Operaciones                        de Fabricación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   585
         TabIndex        =   27
         Top             =   2730
         Width           =   3135
      End
      Begin VB.CommandButton Command18 
         Caption         =   "           Estructura de Conjunto                    Composición y Fórmulas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   585
         TabIndex        =   26
         Top             =   1995
         Width           =   3135
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Modificar Descripción y Atributos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   585
         TabIndex        =   25
         Top             =   1260
         Width           =   3135
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Sustituir Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   585
         TabIndex        =   24
         Top             =   525
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Actualizacion del Archivo"
      ForeColor       =   &H00000040&
      Height          =   3810
      Left            =   540
      TabIndex        =   18
      Top             =   315
      Width           =   4560
      Begin VB.CommandButton Command1 
         Caption         =   "Alta Nuevo Item de Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   585
         TabIndex        =   22
         Top             =   420
         Width           =   3480
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Modificacion de Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   585
         TabIndex        =   21
         Top             =   1155
         Width           =   3525
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Baja (logica) de Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   585
         TabIndex        =   20
         Top             =   1890
         Width           =   3525
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Consultas y Listados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   585
         TabIndex        =   19
         Top             =   2730
         Width           =   3480
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Alta Nuevo Item de Stock"
      ForeColor       =   &H00000080&
      Height          =   4530
      Left            =   5850
      TabIndex        =   7
      Top             =   2415
      Visible         =   0   'False
      Width           =   4305
      Begin VB.Frame Frame5 
         BackColor       =   &H0080FF80&
         Caption         =   "Alta Directa de Item de Stock"
         ForeColor       =   &H00000040&
         Height          =   4530
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
         Width           =   4305
         Begin VB.CommandButton Command15 
            Caption         =   "Cancelar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2535
            TabIndex        =   17
            Top             =   1785
            Width           =   1575
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Aceptar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   195
            TabIndex        =   16
            Top             =   1785
            Width           =   1770
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   13
            Top             =   735
            Width           =   2550
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo"
            Height          =   330
            Left            =   585
            TabIndex        =   15
            Top             =   840
            Width           =   795
         End
         Begin VB.Label Label1 
            BackColor       =   &H0080FF80&
            Caption         =   $"ABM-ITEM.frx":0000
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1800
            Left            =   195
            TabIndex        =   14
            Top             =   2625
            Width           =   3915
         End
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   1560
         TabIndex        =   11
         Top             =   3570
         Width           =   1380
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Importacion de Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   780
         TabIndex        =   10
         Top             =   2520
         Width           =   2940
      End
      Begin VB.CommandButton Command11 
         Caption         =   "        Alta por Copia de Item                                 Existente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   780
         TabIndex        =   9
         Top             =   1575
         Width           =   2940
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Alta Directa (por Teclado)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   780
         TabIndex        =   8
         Top             =   630
         Width           =   2940
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Clasificaciones"
      ForeColor       =   &H00000040&
      Height          =   1905
      Left            =   5850
      TabIndex        =   4
      Top             =   315
      Width           =   4305
      Begin VB.CommandButton Command9 
         Caption         =   "Clasificacion Contable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   390
         TabIndex        =   6
         Top             =   1155
         Width           =   3525
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Clasificacion de Stock y Ventas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   390
         TabIndex        =   5
         Top             =   420
         Width           =   3525
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Controles de Consistencia"
      ForeColor       =   &H00000040&
      Height          =   2850
      Left            =   540
      TabIndex        =   0
      Top             =   4095
      Width           =   4635
      Begin VB.CommandButton Command7 
         Caption         =   "Repeticion en Despieces"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   630
         TabIndex        =   3
         Top             =   2100
         Width           =   3480
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Recursividad de Estructuras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   630
         TabIndex        =   2
         Top             =   1260
         Width           =   3480
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Maestro de Items de Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   630
         TabIndex        =   1
         Top             =   420
         Width           =   3480
      End
   End
   Begin VB.Image Image1 
      Height          =   1995
      Left            =   7020
      Picture         =   "ABM-ITEM.frx":011E
      Top             =   3885
      Width           =   1980
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public COD$

Private Sub Command1_Click()
    Frame4.Visible = True
    Frame4.Enabled = True
End Sub

Private Sub Command10_Click()
    Frame5.Visible = True
    Text1.Visible = True
    Label1.Visible = True
    Command14.Visible = True
    Command15.Visible = True
End Sub

Private Sub Command12_Click()
    Call COMUNI("Opcion no Disponible\en la Presente Version")
End Sub

Private Sub Command13_Click()
    Frame4.Visible = False
End Sub

Private Sub Command14_Click()
    COD$ = Text1.Text
    Call ALTAITEM(COD$)
End Sub

Private Sub Command15_Click()
    Frame5.Visible = False
End Sub

Private Sub Command16_Click()
    Form3.Show
End Sub

Private Sub Command2_Click()
    Frame6.Visible = True
    Frame6.Enabled = True
End Sub

Private Sub Command20_Click()
    Frame6.Enabled = False
    Frame6.Visible = False
End Sub

Private Sub Form_Load()
    cr$ = Chr$(10) + Chr$(13)
End Sub

Private Sub Image1_Click()
   End
End Sub

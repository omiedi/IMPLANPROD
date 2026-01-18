VERSION 5.00
Begin VB.Form FORTIPIF 
   BackColor       =   &H00E0F0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fórmulas Tipificadas - A-B-M-L Estructuras Productos a Re-Exportar"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   Icon            =   "FORTIPIF.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   1360
      Left            =   9135
      ScaleHeight     =   1305
      ScaleWidth      =   60
      TabIndex        =   53
      Top             =   540
      Width           =   120
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   5000
         Left            =   0
         TabIndex        =   54
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0F0F0&
      Caption         =   "Varios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1850
      Left            =   8130
      TabIndex        =   49
      Top             =   210
      Width           =   930
      Begin VB.CommandButton Command39 
         Caption         =   "Print"
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
         Left            =   105
         Picture         =   "FORTIPIF.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Imprime Ficha del Artículo"
         Top             =   1020
         Width           =   700
      End
      Begin VB.CommandButton Command37 
         Caption         =   "Setup"
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
         Left            =   105
         Picture         =   "FORTIPIF.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   315
         Width           =   700
      End
   End
   Begin VB.TextBox Text2 
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
      Left            =   525
      MaxLength       =   22
      TabIndex        =   32
      Top             =   1710
      Width           =   2250
   End
   Begin VB.Frame MARCOFICHA 
      BackColor       =   &H00E0F0F0&
      BorderStyle     =   0  'None
      Caption         =   "2"
      Height          =   6315
      Index           =   2
      Left            =   105
      TabIndex        =   14
      Top             =   2100
      Width           =   11775
      Begin VB.Frame Frame13 
         BackColor       =   &H00E0F0F0&
         Caption         =   "Varios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   10650
         TabIndex        =   40
         Top             =   4410
         Width           =   1040
         Begin VB.CommandButton Command5 
            Caption         =   "Master"
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
            Left            =   125
            Picture         =   "FORTIPIF.frx":0DB6
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Acceso a Maestro de Items de Stock"
            Top             =   315
            Width           =   765
         End
         Begin VB.CommandButton Command38 
            Caption         =   "List"
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
            Left            =   125
            Picture         =   "FORTIPIF.frx":0F00
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Emisión de Listados"
            Top             =   960
            Width           =   765
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         Left            =   105
         TabIndex        =   31
         Top             =   630
         Width           =   11565
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   105
         ScaleHeight     =   375
         ScaleWidth      =   11490
         TabIndex        =   29
         Top             =   180
         Width           =   11550
         Begin VB.Label Label27 
            BackColor       =   &H00E0E0E0&
            Caption         =   $"FORTIPIF.frx":120A
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   30
            TabIndex        =   30
            Top             =   105
            Width           =   11355
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0F0F0&
         Caption         =   "Actualizar Fórmula Tipificada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   105
         TabIndex        =   15
         Top             =   4410
         Width           =   10460
         Begin VB.CommandButton Command7 
            Caption         =   "Importar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   265
            Left            =   105
            TabIndex        =   55
            Top             =   1210
            Width           =   1065
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Importar- Viejo Proc."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   265
            Left            =   1470
            TabIndex        =   52
            Top             =   1365
            Visible         =   0   'False
            Width           =   2325
         End
         Begin VB.TextBox Text32 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   2
            Left            =   7875
            TabIndex        =   44
            Text            =   " "
            Top             =   820
            Width           =   1380
         End
         Begin VB.TextBox Text32 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   7875
            TabIndex        =   42
            Text            =   " "
            Top             =   420
            Width           =   1380
         End
         Begin VB.CommandButton Command18 
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
            Height          =   275
            Left            =   9450
            Picture         =   "FORTIPIF.frx":1293
            TabIndex        =   22
            ToolTipText     =   "Eldición del Elemento Activo de la Fórmula"
            Top             =   1260
            Width           =   855
         End
         Begin VB.CommandButton Command31 
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
            Height          =   300
            Left            =   3045
            TabIndex        =   21
            Top             =   420
            Width           =   175
         End
         Begin VB.CommandButton Command19 
            Caption         =   "Delete"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   275
            Left            =   9450
            Picture         =   "FORTIPIF.frx":159D
            TabIndex        =   20
            ToolTipText     =   "Elimina el Elemento Activo de la Fórmula"
            Top             =   970
            Width           =   855
         End
         Begin VB.CommandButton Command26 
            Caption         =   "Insert"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   275
            Left            =   9450
            Picture         =   "FORTIPIF.frx":18A7
            TabIndex        =   19
            ToolTipText     =   "Intercala el Elemento en la Fórmula Antes del Activo"
            Top             =   695
            Width           =   855
         End
         Begin VB.TextBox Text31 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1260
            TabIndex        =   18
            Top             =   420
            Width           =   1800
         End
         Begin VB.TextBox Text32 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   5460
            TabIndex        =   17
            Text            =   " "
            Top             =   420
            Width           =   1380
         End
         Begin VB.CommandButton Command27 
            Caption         =   "Append"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   275
            Left            =   9450
            Picture         =   "FORTIPIF.frx":1BB1
            TabIndex        =   16
            ToolTipText     =   "Agrega el Elemento al Final de la Fórmula"
            Top             =   410
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7875
            TabIndex        =   47
            Top             =   1220
            Width           =   1380
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "C.Bruto:"
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
            Left            =   6930
            TabIndex        =   46
            Top             =   1300
            Width           =   855
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Pérdidas:"
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
            Left            =   6930
            TabIndex        =   45
            Top             =   900
            Width           =   855
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Mermas:"
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
            Left            =   6930
            TabIndex        =   43
            Top             =   500
            Width           =   855
         End
         Begin VB.Label Label41 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4095
            TabIndex        =   28
            Top             =   420
            Width           =   435
         End
         Begin VB.Label Label40 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   670
            Left            =   1260
            TabIndex        =   27
            Top             =   820
            Width           =   5580
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción:"
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
            Left            =   105
            TabIndex        =   26
            Top             =   945
            Width           =   1065
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "U-M:"
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
            Left            =   3150
            TabIndex        =   25
            Top             =   525
            Width           =   855
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
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
            Left            =   210
            TabIndex        =   24
            Top             =   525
            Width           =   960
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "C.Neto:"
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
            Left            =   4515
            TabIndex        =   23
            Top             =   525
            Width           =   855
         End
      End
      Begin VB.Line Line2 
         X1              =   11655
         X2              =   11655
         Y1              =   180
         Y2              =   630
      End
      Begin VB.Line Line1 
         X1              =   8295
         X2              =   8295
         Y1              =   380
         Y2              =   840
      End
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1575
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   735
      Width           =   6360
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1575
      MaxLength       =   15
      TabIndex        =   5
      Top             =   315
      Width           =   2040
   End
   Begin VB.CommandButton Command9 
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
      Height          =   330
      Left            =   3600
      Picture         =   "FORTIPIF.frx":1EBB
      TabIndex        =   4
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0F0F0&
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1850
      Left            =   9345
      TabIndex        =   0
      Top             =   210
      Width           =   2430
      Begin VB.CommandButton Command21 
         Caption         =   "New"
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
         Left            =   105
         Picture         =   "FORTIPIF.frx":21C5
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Alta de Nueva Materia Prima"
         Top             =   315
         Width           =   700
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Check"
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
         Left            =   105
         Picture         =   "FORTIPIF.frx":230F
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Control y Ordenamiento Creciente Archivo Materias Primas"
         Top             =   1020
         Width           =   700
      End
      Begin VB.CommandButton Command3 
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
         Height          =   700
         Index           =   4
         Left            =   1575
         Picture         =   "FORTIPIF.frx":2619
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   315
         Width           =   700
      End
      Begin VB.CommandButton Command4 
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
         Height          =   700
         Left            =   1575
         Picture         =   "FORTIPIF.frx":2763
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   1020
         Width           =   700
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Clear"
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
         Left            =   840
         Picture         =   "FORTIPIF.frx":2A6D
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Restablece los Valores Guardados"
         Top             =   1020
         Width           =   700
      End
      Begin VB.CommandButton Command1 
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
         Height          =   700
         Left            =   840
         Picture         =   "FORTIPIF.frx":2D77
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   315
         Width           =   700
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   3255
      TabIndex        =   35
      Top             =   1710
      Width           =   1170
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4935
      TabIndex        =   37
      Top             =   1710
      Width           =   2985
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Responsable.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4410
      TabIndex        =   36
      Top             =   1455
      Width           =   1320
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "F.U.A.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2625
      TabIndex        =   34
      Top             =   1455
      Width           =   1005
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "N.C.M.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -105
      TabIndex        =   33
      Top             =   1460
      Width           =   1065
   End
   Begin VB.Label Label38 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "U-M:"
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
      Left            =   3780
      TabIndex        =   12
      Top             =   420
      Width           =   1005
   End
   Begin VB.Label Label37 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4875
      TabIndex        =   11
      Top             =   315
      Width           =   540
   End
   Begin VB.Label Label36 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Interno:"
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
      Left            =   6090
      TabIndex        =   10
      Top             =   420
      Width           =   1005
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7140
      TabIndex        =   9
      Top             =   315
      Width           =   780
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
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
      Left            =   105
      TabIndex        =   8
      Top             =   735
      Width           =   1380
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
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
      Left            =   420
      TabIndex        =   7
      Top             =   420
      Width           =   1060
   End
End
Attribute VB_Name = "FORTIPIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TOPEINI(3), IZQUINI(3), ANCHINI(3), ALTUINI(3)
'
Dim CIACTI%(15)
Dim ARCHEX$(15)
Dim YACAR%, MODIFIC%
Dim GRABATODO%
Dim AGREITEM%
Dim MODOEDIT%
'
Dim RESECO$(128), DESOPE$(128)
Dim MOREPO%, AREPO$
'
Private Sub Command1_Click()
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   GRABATODO% = 1
   Call GRADATITEM
   GRABATODO% = 0
End Sub

Private Sub Command18_Click()
    '
    If VALCODART% < 1 Then Exit Sub
    '
    If List1.ListIndex < 0 Then
       Call MENSERR(24, "Opción Edición no Válida")
       Exit Sub
    End If
    '
    Text31.TEXT = Left$(List1.TEXT, 16)
    CDYY% = CODINT%(Text31.TEXT)
    Label40.Caption = DESCRIT0$(CDYY%)
    Label41.Caption = UNIMED$(CDYY%)
    Text32(0).TEXT = Mid$(List1.TEXT, 52, 11)
    CONBRU# = Val(Mid$(List1.TEXT, 83, 11))
    Text32(1).TEXT = FORMATNUM$(Val(Mid$(List1.TEXT, 64, 8)) * CONBRU# / 100, "F11.6")
    Text32(2).TEXT = FORMATNUM$(Val(Mid$(List1.TEXT, 73, 8)) * CONBRU# / 100, "F11.6")
    '
    ILISTA& = List1.ListIndex
    List1.RemoveItem ILISTA&
    Command18.Enabled = False
    Command19.Enabled = False
    '
    If ILISTA& < List1.ListCount Then
        List1.ListIndex = ILISTA&
      Else
        List1.ListIndex = List1.ListCount - 1
    End If
    Text32(0).SetFocus
    MODOEDIT% = 1
    MODIFIC% = 1
    '
End Sub

Private Sub Command19_Click()
    '
    If VALCODART% < 1 Then Exit Sub
    '
    If List1.ListIndex < 0 Then
       Call MENSERR(24, "Opción Borrado no Válida")
       Exit Sub
    End If
    '
    ILISTA& = List1.ListIndex
    List1.RemoveItem ILISTA&
    '
    If ILISTA& < List1.ListCount Then
        List1.ListIndex = ILISTA&
      Else
        List1.ListIndex = List1.ListCount - 1
    End If
    MODIFIC% = 1
    '
End Sub

Private Sub Command2_Click()
   '
   If COMALTER%("Realmente Desea Descartar los Cambios ?\\No, Continuar\Si, Descartar Cambios") <> 2 Then
      Exit Sub
   End If
   '
   YACAR% = 0
   MODIFIC% = 0
   '
   Command18.Enabled = True
   Command19.Enabled = True
   MODOEDIT% = 0
   Call BLANFORMU
   If CODINT%(Text1.TEXT) > 0 Then
       Call CARDATITEM
   End If
   '
End Sub

Private Sub COMMAND21_CLICK()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command21.Enabled = False
   Call GRADATITEM
   '
   ALTX% = ALTERNA%("M.Prima de Codigo Existente\Alta de Nuevo Codigo")
   If ALTX% < 1 Or ALTX% > 2 Then GoTo 99
   '
   If ALTX% = 1 Then
10   Call SELECHO("MASTER")
     CODD$ = VALACT1$("MASTER")
     If CODD$ <> "" Then
       CIXI% = CODINT%(CODD$)
       RFF$ = RECFILT$("MATPRIM", 1, MKI$(CIXI%))
       If Len(RFF$) < 4 Then
           DEXX$ = DESCRIT0$(CIXI%)
           REBLA$ = REGBLAN$("MATPRIM")
           Call REPLA(REBLA$, MKI$(CIXI%), 1, 2)
           Call REPLA(REBLA$, DEXX$, 3, 62)
           Call REGAPP("MATPRIM", REBLA$)
           Call CIERRARCH("MATPRIM")
         Else
           Call MENSERR(24, "Ya Registrado como M.Prima")
           GoTo 10
       End If
       Text1 = CODD$
     End If
     GoTo 99
   End If
   '
   VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
   Call REPLA(VDEF$, "U.", 109, 2)
   Call REPLA(VDEF$, Chr$(1), 93, 1)
   '
   OKX% = 0
   Call ALTACODIGO(VDEF$, OKX%)
   If OKX% = 1 Then
      CODINUE$ = Mid$(VDEF$, 77, 16)
      Text1.TEXT = CODINUE$
      '
      CIXI% = Val(Label28)
      DEXX$ = DESCRIT0$(CIXI%)
      REBLA$ = REGBLAN$("MATPRIM")
      Call REPLA(REBLA$, MKI$(CIXI%), 1, 2)
      Call REPLA(REBLA$, Text3, 3, 62)
      Call REGAPP("MATPRIM", REBLA$)
      Call CIERRARCH("MATPRIM")
   End If
   '
99 Command21.Enabled = True
   '
End Sub

Private Sub Command22_Click()
   '
   Command22.Enabled = False
   '
   Screen.MousePointer = 1
   JJ& = 0
   FIN& = ULTREG&("MATPRIM")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("MATPRIM", U&)
     CIXI% = CVI(Left$(XX$, 2))
     If STATUIT%(CIXI%) > 0 Then
       DEDD$ = DESCRIT0$(CIXI%)
       Call REPLA(XX$, DEDD$, 3, 62)
       JJ& = JJ& + 1
       Call GRAREG("MATPRIM", XX$, JJ&)
     End If
   Next U&
   Call SETULTREG("MATPRIM", JJ&)
   Call CIERRARCH("MATPRIM")
   '
   Call FILESORT("MATPRIM", "", 1, 1)
   Screen.MousePointer = 1
   '
99 Command22.Enabled = True
   '
End Sub

Private Sub Command26_Click()
   '
   CODXX$ = TRIM$(Text31.TEXT)
   If CODXX$ = "" Or CODINT%(CODXX$) < 1 Then
      Call MENSERR(24, "Código no Válido")
      Text31.SetFocus
      Exit Sub
   End If
   '
   CIY% = CODINT%(CODXX$)
   DEXX$ = DESCRIT$(CIY%)
   UNXX$ = UNIMED$(CIY%)
   USO = Val(Text32(0).TEXT)
   If Abs(USO) < 0.000001 Then
      Call MENSERR(24, "Uso no Válido")
      Text32(0).SetFocus
      Exit Sub
   End If
   '
   MODOEDIT% = 0
   CICO% = CIY%
   If CICO% > 0 Then
     If CICO% <= NUMAS% Then
       TTXX$ = Space$(128)
       Call REPLA(TTXX$, CODEXT$(CICO%), 1, 15)
       Call REPLA(TTXX$, COTIPIF$(CICO%), 17, 12)
       Call REPLA(TTXX$, DESCRIT0$(CICO%), 31, 16)
       Call REPLA(TTXX$, UNIMED$(CICO%), 48, 4)
       USONE# = Val(Text32(0))
       Call REPLA(TTXX$, FORMATNUM$(USONE#, "F11.6"), 52, 11)
       USOTO# = Val(Label2)
       Call REPLA(TTXX$, FORMATNUM$(100 * (Val(Text32(1))) / USOTO#, "F8.2"), 64, 8)
       Call REPLA(TTXX$, FORMATNUM$(100 * (Val(Text32(2))) / USOTO#, "F8.2"), 73, 8)
       Call REPLA(TTXX$, FORMATNUM$(USOTO#, "F11.6"), 83, 11)
       ILISTA& = List1.ListIndex
       If ILISTA& < 0 Then ILISTA& = 0
       '
       List1.AddItem TTXX$, ILISTA&
       List1.ListIndex = ILISTA&
       List1.Refresh
       '
     End If
   End If
   '
   Text31.TEXT = ""
   CODE$ = ""
   Label40.Caption = ""
   Label41.Caption = ""
   For KKI% = 0 To 2: Text32(KKI%) = "": Next KKI%
   Text31.SetFocus
   MODIFIC% = 1
   '
   Command18.Enabled = True
   Command19.Enabled = True
   '
End Sub

Private Sub Command27_Click()
   '
   If VALCODART% < 1 Then Exit Sub
   '
   CODXX$ = TRIM$(Text31.TEXT)
   If CODXX$ = "" Or CODINT%(CODXX$) < 1 Then
      Call MENSERR(24, "Código no Válido")
      Text31.SetFocus
      Exit Sub
   End If
   '
   CIY% = CODINT%(CODXX$)
   DEXX$ = DESCRIT$(CIY%)
   UNXX$ = UNIMED$(CIY%)
   USO = Val(Text32(0).TEXT)
   If Abs(USO) < 0.000001 Then
      Call MENSERR(24, "Uso no Válido")
      Text32(0).SetFocus
      Exit Sub
   End If
   '
   MODOEDIT% = 0
   CICO% = CIY%
   If CICO% > 0 Then
     If CICO% <= NUMAS% Then
       TTXX$ = Space$(128)
       Call REPLA(TTXX$, CODEXT$(CICO%), 1, 15)
       Call REPLA(TTXX$, COTIPIF$(CICO%), 17, 12)
       Call REPLA(TTXX$, DESCRIT0$(CICO%), 31, 16)
       Call REPLA(TTXX$, UNIMED$(CICO%), 48, 4)
       USONE# = Val(Text32(0))
       Call REPLA(TTXX$, FORMATNUM$(USONE#, "F11.6"), 52, 11)
       USOTO# = Val(Label2)
       Call REPLA(TTXX$, FORMATNUM$(100 * (Val(Text32(1))) / USOTO#, "F8.2"), 64, 8)
       Call REPLA(TTXX$, FORMATNUM$(100 * (Val(Text32(2))) / USOTO#, "F8.2"), 73, 8)
       Call REPLA(TTXX$, FORMATNUM$(USOTO#, "F11.6"), 83, 11)
       JJ& = JJ& + 1
       List1.AddItem TTXX$
       List1.ListIndex = List1.ListCount - 1
     End If
   End If
   '
   Text31.TEXT = ""
   CODE$ = ""
   Label40.Caption = ""
   Label41.Caption = ""
   For KKI% = 0 To 2: Text32(KKI%) = "": Next KKI%
   Text31.SetFocus
   MODIFIC% = 1
   '
   Command18.Enabled = True
   Command19.Enabled = True
   '
End Sub

Private Sub Command3_Click(Index As Integer)
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub COMMAND31_CLICK()
   '
   Call SELECHO("MATPRIM") 'MASTER")
   COCOMPO$ = TRIM$(VALACT1$("MATPRIM"))   'MASTER"))
   If COCOMPO$ <> "" Then
     Text31.TEXT = COCOMPO$
     CICOMPO% = CODINT%(COCOMPO$)
     Label40.Caption = DESCRIT0$(CICOMPO%)
     Label41.Caption = UNIMED$(CICOMPO%)
     If TRIM$(Text32(0).TEXT) = "" Then Text32(0).TEXT = Str$(CABAS)
     Text32(0).SetFocus
   End If
    '
End Sub

Private Sub Command37_Click()
   Command37.Enabled = False
   'OPMAST00.Show 1
   Call CONECRUN("FLSETUP/OPMAST00", "Configuración de Funcionamiento")
   Command37.Enabled = True
End Sub

Private Sub Command38_Click()
   Command38.Enabled = False
   Call LIESTRU
   Command38.Enabled = True
End Sub

Private Sub Command39_Click()
   Command39.Enabled = False
   If VALCODART% > 0 Then
     PRINTFORM
   End If
   Command39.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call HELPONLINE("SIMPTEMP")
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   ACONEC$ = "AMASTO02"
   If EXISTE%("AMASTO02.EXE") < 1 Then
     ACONEC$ = "AMASTO01"
     If EXISTE%("AMASTO01.EXE") < 1 Then
       ACONEC$ = "AMASTO"
     End If
   End If
   Call CONECRUN(ACONEC$, "Maestro de Items de Stock")
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   '
   Dim CIMATI%(256)
   Dim PORPER(256)
   Dim CONTOT#(256)
   Dim CONNET#(256)
   Dim CONPER#(256)
   '
   TTXX$ = "Para Importar Fórmulas Tipificadas\Debe Colocar Archivo 'TEXTIPIF.TXT'\del 13/10/03 en carpeta 'C:/FLEXOFT'"
   If COMALTER%(TTXX$ + "\\Cancelar\Continuar Importación") <> 2 Then Exit Sub
   '
   Call COMUNI("Este Proceso Puede Demorar Varios Minutos")
   '
   Screen.MousePointer = 11
   N1% = FreeFile
   Dim BUFLE As String * 256
   '
   Open "C:\FLEXOFT\TEXTIPIF.TXT" For Random Access Read As #N1% Len = 256
   FIN& = LOF(N1%) / 256
   For U& = 1 To LOF(N1%) / 256
     Screen.MousePointer = 11
     Call TRACE(20, U&, FIN&)
     Get #N1%, U&, BUFLE$
     CODCON$ = TRIM$(Left$(BUFLE$, 16))
     CICON% = CODINT%(CODCON$)
     DECON$ = TRIM$(Mid$(BUFLE$, 17, 48))
     If CICON% < 1 Then
       Close #N1%
       RENUE$ = REGBLAN$("MASTER")
       NMS% = NUMAS% + 1
       Call REPLA(RENUE$, MKI$(NMS%), 1, 2)
       Call REPLA(RENUE$, CODCON$, 77, 16)
       Call REPLA(RENUE$, DECON$, 5, 64)
       Call REPLA(RENUE$, Chr$(1), 93, 1)
       Call REPLA(RENUE$, "U.", 109, 2)
       Call ALTACODIGO(RENUE$, (-1))
       Screen.MousePointer = 11
       N1% = FreeFile
       Open "C:\FLEXOFT\TEXTIPIF.TXT" For Random Access Read As #N1% Len = 256
     End If
     '
     CODMAT$ = TRIM$(Mid$(BUFLE$, 81, 16))
     If CODMAT$ <> "" Then
       CIMAT% = CODINT%(CODMAT$)
       DEMAT$ = TRIM$(Mid$(BUFLE$, 97, 48))
       If CIMAT% < 1 Then
         Close #N1%
         RENUE$ = REGBLAN$("MASTER")
         NMS% = NUMAS% + 1
         Call REPLA(RENUE$, MKI$(NMS%), 1, 2)
         Call REPLA(RENUE$, CODMAT$, 77, 16)
         Call REPLA(RENUE$, DEMAT$, 5, 64)
         Call REPLA(RENUE$, Chr$(1), 93, 1)
         Call REPLA(RENUE$, "U.", 109, 2)
         Call ALTACODIGO(RENUE$, OKX%)
         N1% = FreeFile
         Open "C:\FLEXOFT\TEXTIPIF.TXT" For Random Access Read As #N1% Len = 256
       End If
     End If
   Next U&
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
15 PUCOIN$ = ""
   Screen.MousePointer = 11
   N1% = FreeFile
   Open "C:\FLEXOFT\TEXTIPIF.TXT" For Random Access Read As #N1% Len = 256
   For U& = 1 To LOF(N1%) / 256
     Get #N1%, U&, BUFLE$
     CODCON$ = TRIM$(Left$(BUFLE$, 16))
     CICON% = CODINT%(CODCON$)
     For TU% = 1 To Len(PUCOIN$) Step 2
       If CVI(Mid$(PUCOIN$, TU%, 2)) = CICON% Then GoTo 25
     Next TU%
     PUCOIN$ = PUCOIN$ + MKI$(CICON%)
25 Next U&
   '
   HOII% = HOY(HOS$)
   UREFO& = LOF(N1%) / 256
   JJ& = 0
   '
   Screen.MousePointer = 11
   For TU% = 1 To Len(PUCOIN$) Step 2
     CICON% = CVI(Mid$(PUCOIN$, TU%, 2))
     COCON$ = AJUSTI$(CODEXT$(CICON%), 15)
     CAMATI% = 0: NOARAN$ = ""
     For U& = 1 To UREFO&
       Get #N1%, U&, BUFLE$
       If Left$(BUFLE$, 15) = COCON$ Then
         CAMATI% = CAMATI% + 1
         If NOARAN$ = "" Then NOARAN$ = TRIM$(Mid$(BUFLE$, 65, 16))
         CIMATI%(CAMATI%) = CODINT%(Mid$(BUFLE$, 81, 16))
         PORPER(CAMATI%) = Val(Mid$(BUFLE$, 145, 12))
         CONTOT#(CAMATI%) = Val(Mid$(BUFLE$, 157, 12))
         CONNET#(CAMATI%) = CONTOT#(CAMATI%) * (100 - PORPER(CAMATI%)) / 100
         CONPER#(CAMATI%) = CONTOT#(CAMATI%) - CONNET#(CAMATI%)
       End If
     Next U&
     '
     REBLA$ = REGBLAN$("TIPIFIC")
     Call REPLA(REBLA$, MKI$(CICON%), 1, 2)
     Call REPLA(REBLA$, NOARAN$, 3, 22)
     Call REPLA(REBLA$, MKI$(HOII%), 25, 2)
     Call REPLA(REBLA$, USERNAME$, 27, 22)
     '
     For CAMAI% = 1 To CAMATI%
       TTXX$ = REBLA$
       Call REPLA(TTXX$, MKI$(CIMATI%(CAMAI%)), 49, 2)
       Call REPLA(TTXX$, UNIMED$(CIMATI%(CAMAI%)), 51, 4)
         VALOD1# = (Int(1000000 * CONNET#(CAMAI%) + 0.5)) / 1000000 ' 6 DECIMALES
       Call REPLA(TTXX$, MKD$(VALOD1#), 55, 8)
       Call REPLA(TTXX$, MKD$(0), 63, 8)
       Call REPLA(TTXX$, MKS$(0), 71, 4)
         VALOD2# = (Int(1000000 * CONPER#(CAMAI%) + 0.5)) / 1000000 ' 6 DECIMALES
       Call REPLA(TTXX$, MKD$(VALOD2#), 75, 8)
         VALOD3 = (Int(1000000 * PORPER(CAMAI%) + 0.5)) / 1000000 ' 6 DECIMALES
       Call REPLA(TTXX$, MKS$(PORPER(CAMAI%)), 83, 4)
       Call REPLA(TTXX$, MKD$(VALOD1# + VALOD2#), 87, 8)
       '
       JJ& = JJ& + 1
       Call GRAREG("TIPIFIC", TTXX$, JJ&)
     Next CAMAI%
     '
     For JJI% = 1 To 256
       EXPLOLIN$(JJI%) = ""
       If JJI% <= CAMATI% Then
         TTXY$ = Space$(64)
         Call REPLA(TTXY$, CODEXT$(CIMATI%(JJI%)), 1, 16)
         Call REPLA(TTXY$, DESCRIT0$(CIMATI%(JJI%)), 17, 28)
         Call REPLA(TTXY$, UNIMED$(CIMATI%(JJI%)), 47, 2)
         Call REPLA(TTXY$, FORMATNUM$(VALOD1# + VALOD2#, "F12.4"), 49, 12)
         EXPLOLIN$(JJI%) = TTXY$
       End If
     Next JJI%
     CAIT% = CAMATI%
     '
     CODD$ = CODEXT$(CICON%)
     CABAS = 1
     Call EXPLOGR(CODD$, CABAS)
     '
   Next TU%
   '
   Call SETULTREG("TIPIFIC", JJ&)
   Call CIERRARCH("TIPIFIC")
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   MsgBox "Importacion Completa"
   '
End Sub

Private Sub Command7_Click()
   '
   Command7.Enabled = False
   '
   Dim CIMATI%(256)
   Dim PORPER(256)
   Dim CONTOT#(256)
   Dim CONNET#(256)
   Dim CONPER#(256)
   '
   Dim CODIVIE$(1024), CODINUE$(1024)
   '
   JMP& = 0
   TTXX$ = "Para Importar Fórmulas Tipificadas\Debe Colocar Archivo 'TEXTIPIF_NUE.PRN'\del 30/12/03 y 'TOMAPRI.PRN' del 30/12/03\en carpeta 'C:/FLEXOFT'"
   If COMALTER%(TTXX$ + "\\Cancelar\Continuar Importación") <> 2 Then Exit Sub
   '
   Call COMUNI("Este Proceso Puede Demorar Varios Minutos")
   '
   Screen.MousePointer = 11
   N1% = FILEOPEN%("C:\FLEXOFT\TEXTIPIF_NUE.PRN", 1, 128) ' abre para lectura secuencial
   Dim BUFLE As String * 256
   '
   FIN& = LOF(N1%) / 10
   LACU& = 0
   ATX% = 0: NMS% = NUMAS%
   While Not EOF(N1%)
     Screen.MousePointer = 11
     Line Input #N1%, BUFLE$
     LACU& = LACU& + Len(BUFLE$)
     Call TRACE(20, LACU& / 10, FIN&)
     '
     CODCON$ = TRIM$(Left$(BUFLE$, 16))
     CICON% = CODINT%(CODCON$)
     DECON$ = TRIM$(Mid$(BUFLE$, 17, 48))
     If CICON% < 1 Then
       ATX% = 1
       RENUE$ = REGBLAN$("MASTER")
       NMS% = NMS% + 1
       Call REPLA(RENUE$, MKI$(NMS%), 1, 2)
       Call REPLA(RENUE$, CODCON$, 77, 16)
       Call REPLA(RENUE$, DECON$, 5, 64)
       Call REPLA(RENUE$, Chr$(1), 93, 1)
       Call REPLA(RENUE$, "U.", 109, 2)
       Call REGAPP("MASTER", RENUE$)
     End If
   Wend
   '
   Close #N1%
   If ATX% > 0 Then
     Screen.MousePointer = 1
     Call COMUNI("Se Incorporaron Códigos de Articulo.\Ejecute Control de Maestro.")
     GoTo 99
   End If
   '
   CACOREP% = 0
   N1% = FILEOPEN%("C:\FLEXOFT\TEXTIPIF_NUE.PRN", 1, 128) ' abre para lectura secuencial
   '
   FIN& = LOF(N1%) / 10
   LACU& = 0
   ATX% = 0: NMS% = NUMAS%
   While Not EOF(N1%)
     Screen.MousePointer = 11
     Line Input #N1%, BUFLE$
     LACU& = LACU& + Len(BUFLE$)
     Call TRACE(20, LACU& / 10, FIN&)
     '
     CODVIE$ = TRIM$(Mid$(BUFLE$, 81, 16))
     For KK% = 1 To CACOREP%
       If CODIVIE$(KK%) = CODVIE$ Then GoTo 12
     Next KK%
     CACOREP% = CACOREP% + 1
     CODIVIE$(CACOREP%) = CODVIE$
     CODINUE$(CACOREP%) = TRIM$(Mid$(BUFLE$, 97, 16))
     '
12 Wend
   '
   Close #N1%
   If CACOREP% > 0 Then
     ATX% = 0
     For KK% = 1 To CACOREP%
       CODD$ = CODIVIE$(KK%)
       CIXI% = CODINT%(CODD$)
       If CIXI% > 0 Then
         If CODINUE$(KK%) <> CODIVIE$(KK%) Then
           REMA& = CIXI% + 1
           RENUE$ = REGLEIDO$("MASTER", REMA&)
           Call REPLA(RENUE$, CODINUE$(KK%), 77, 16)
           Call GRAREG("MASTER", RENUE$, REMA&)
           ATX% = 1
         End If
       End If
     Next KK%
     Call CIERRARCH("MASTER")
     '
     Screen.MousePointer = 1
     If ATX% > 0 Then
       Call COMUNI("Se Sustituyeron Códigos de Componentes.\Ejecute Control de Maestro.")
       GoTo 99
     End If
   End If
   '
   N1% = FILEOPEN%("C:\FLEXOFT\TOMAPRI.PRN", 1, 128) ' abre para lectura secuencial
   '
   CACONUE% = 0
   FIN& = LOF(N1%) / 10
   LACU& = 0: CACONUE% = 0
   ATX% = 0: NMS% = NUMAS%
   While Not EOF(N1%)
     Screen.MousePointer = 11
     Line Input #N1%, BUFLE$
     LACU& = LACU& + Len(BUFLE$)
     Call TRACE(20, LACU& / 10, FIN&)
     '
     CODCON$ = TRIM$(Left$(BUFLE$, 16))
     If CODCON$ <> "" Then
       CICON% = CODINT%(CODCON$)
       DECON$ = TRIM$(Mid$(BUFLE$, 17))
       If CICON% < 1 Then
         For KK% = 1 To CACONUE%
           If CODINUE$(KK%) = CODCON$ Then GoTo 17
         Next KK%
         ATX% = 1
         RENUE$ = REGBLAN$("MASTER")
         NMS% = NMS% + 1
         Call REPLA(RENUE$, MKI$(NMS%), 1, 2)
         Call REPLA(RENUE$, CODCON$, 77, 16)
         Call REPLA(RENUE$, DECON$, 5, 64)
         Call REPLA(RENUE$, Chr$(1), 93, 1)
         Call REPLA(RENUE$, "U.", 109, 2)
         Call REGAPP("MASTER", RENUE$)
         '
         CACONUE% = CACONUE% + 1
         CODINUE$(CACONUE%) = CODCON$
       End If
       Call AGREMAPRIM(CICON%)
     End If
17 Wend
   '
   Close #N1%
   If ATX% > 0 Then
     Screen.MousePointer = 1
     Call COMUNI("Se Incorporaron Códigos de Materias Primas.\Ejecute Control de Maestro.")
     GoTo 99
   End If
   '
   N1% = FILEOPEN%("C:\FLEXOFT\TEXTIPIF_NUE.PRN", 1, 128) ' abre para lectura secuencial
   '
   FIN& = LOF(N1%) / 10
   LACU& = 0: CACONUE% = 0
   ATX% = 0: NMS% = NUMAS%
   While Not EOF(N1%)
     Screen.MousePointer = 11
     Line Input #N1%, BUFLE$
     LACU& = LACU& + Len(BUFLE$)
     Call TRACE(20, LACU& / 10, FIN&)
     '
     CODCON$ = TRIM$(Mid$(BUFLE$, 97, 16))
     If CODCON$ <> "" Then
       CICON% = CODINT%(CODCON$)
       DECON$ = TRIM$(Mid$(BUFLE$, 113, 24))
       If CICON% < 1 Then
         For KK% = 1 To CACONUE%
           If CODINUE$(KK%) = CODCON$ Then GoTo 18
         Next KK%
         ATX% = 1
         RENUE$ = REGBLAN$("MASTER")
         NMS% = NMS% + 1
         Call REPLA(RENUE$, MKI$(NMS%), 1, 2)
         Call REPLA(RENUE$, CODCON$, 77, 16)
         Call REPLA(RENUE$, DECON$, 5, 64)
         Call REPLA(RENUE$, Chr$(1), 93, 1)
         Call REPLA(RENUE$, "U.", 109, 2)
         Call REGAPP("MASTER", RENUE$)
         '
         CACONUE% = CACONUE% + 1
         CODINUE$(CACONUE%) = CODCON$
       End If
       Call AGREMAPRIM(CICON%)
     End If
18 Wend
   '
   Close #N1%
   If ATX% > 0 Then
     Screen.MousePointer = 1
     Call COMUNI("Se Incorporaron Códigos de Componentes.\Ejecute Control de Maestro.")
     GoTo 99
   End If
   '
   Dim COEQUI$(32767)
   N1% = FILEOPEN%("C:\FLEXOFT\TEXTIPIF_NUE.PRN", 1, 128) ' abre para lectura secuencial
   '
   FIN& = LOF(N1%) / 10
   LACU& = 0: CACONUE% = 0
   While Not EOF(N1%)
     Screen.MousePointer = 11
     Line Input #N1%, BUFLE$
     LACU& = LACU& + Len(BUFLE$)
     Call TRACE(20, LACU& / 10, FIN&)
     '
     CODTIP$ = TRIM$(Mid$(BUFLE$, 81, 16))
     CODCON$ = TRIM$(Mid$(BUFLE$, 97, 16))
     If CODCON$ <> "" Then
       CICON% = CODINT%(CODCON$)
       If CICON% > 0 Then
         If COEQUI$(CICON%) = "" Then
             COEQUI$(CICON%) = CODTIP$
           ElseIf COEQUI$(CICON%) <> CODTIP$ Then
             Call COMUNI("Codigo " + CODCON$ + " con 2 Tipificaciones " + COEQUI$(CICON%) + ", " + CODTIP$)
         End If
       End If
     End If
   Wend
   '
   Close #N1%
   '
   JJ& = 0
   For KK% = 1 To 32000
     If COEQUI$(KK%) <> "" Then
       JJ& = JJ& + 1
       XX$ = REGBLAN$("COEQTIP")
       Call REPLA(XX$, MKI$(KK%), 1, 2)
       Call REPLA(XX$, COEQUI$(KK%), 9, 24)
       Call GRAREG("COEQTIP", XX$, JJ&)
     End If
   Next KK%
   Call SETULTREG("COEQTIP", JJ&)
   Call CIERRARCH("COEQTIP")
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
15 PUCOIN$ = ""
   Screen.MousePointer = 11
   N1% = FILEOPEN%("C:\FLEXOFT\TEXTIPIF_NUE.PRN", 1, 128) ' abre para lectura secuencial
   '
   FIN& = LOF(N1%) / 10
   LACU& = 0
   While Not EOF(N1%)
     Screen.MousePointer = 11
     Line Input #N1%, BUFLE$
     LACU& = LACU& + Len(BUFLE$)
     Call TRACE(20, LACU& / 10, FIN&)
     CODCON$ = TRIM$(Left$(BUFLE$, 16))
     CICON% = CODINT%(CODCON$)
     For TU% = 1 To Len(PUCOIN$) Step 2
       If CVI(Mid$(PUCOIN$, TU%, 2)) = CICON% Then GoTo 25
     Next TU%
     PUCOIN$ = PUCOIN$ + MKI$(CICON%)
25 Wend
   Close #N1%
   '
   HOII% = HOY(HOS$)
   JJ& = 0
   '
   Screen.MousePointer = 11
   For TU% = 1 To Len(PUCOIN$) Step 2
     CICON% = CVI(Mid$(PUCOIN$, TU%, 2))
     COCON$ = AJUSTI$(CODEXT$(CICON%), 15)
     CAMATI% = 0: NOARAN$ = ""
     N1% = FILEOPEN%("C:\FLEXOFT\TEXTIPIF_NUE.PRN", 1, 128) ' abre para lectura secuencial
     While Not EOF(N1%)
       Line Input #N1%, BUFLE$
       If TRIM$(Left$(BUFLE$, 16)) = TRIM$(COCON$) Then
         CAMATI% = CAMATI% + 1
         If NOARAN$ = "" Then NOARAN$ = TRIM$(Mid$(BUFLE$, 65, 16))
         CIMATI%(CAMATI%) = CODINT%(TRIM$(Mid$(BUFLE$, 97, 16)))
         PORPER(CAMATI%) = Val(Mid$(BUFLE$, 137, 12))
         CONTOT#(CAMATI%) = Val(Mid$(BUFLE$, 149, 16))
         CONNET#(CAMATI%) = CONTOT#(CAMATI%) * (100 - PORPER(CAMATI%)) / 100
         CONPER#(CAMATI%) = CONTOT#(CAMATI%) - CONNET#(CAMATI%)
       End If
     Wend
     Close #N1%
     '
     REBLA$ = REGBLAN$("TIPIFIC")
     Call REPLA(REBLA$, MKI$(CICON%), 1, 2)
     Call REPLA(REBLA$, NOARAN$, 3, 22)
     Call REPLA(REBLA$, MKI$(HOII%), 25, 2)
     Call REPLA(REBLA$, USERNAME$, 27, 22)
     '
     For CAMAI% = 1 To CAMATI%
       TTXX$ = REBLA$
       Call REPLA(TTXX$, MKI$(CIMATI%(CAMAI%)), 49, 2)
       Call REPLA(TTXX$, UNIMED$(CIMATI%(CAMAI%)), 51, 4)
         VALOD1# = (Int(1000000 * CONNET#(CAMAI%) + 0.5)) / 1000000 ' 6 DECIMALES
       Call REPLA(TTXX$, MKD$(VALOD1#), 55, 8)
       Call REPLA(TTXX$, MKD$(0), 63, 8)
       Call REPLA(TTXX$, MKS$(0), 71, 4)
         VALOD2# = (Int(1000000 * CONPER#(CAMAI%) + 0.5)) / 1000000 ' 6 DECIMALES
       Call REPLA(TTXX$, MKD$(VALOD2#), 75, 8)
         VALOD3 = (Int(1000000 * PORPER(CAMAI%) + 0.5)) / 1000000 ' 6 DECIMALES
       Call REPLA(TTXX$, MKS$(PORPER(CAMAI%)), 83, 4)
       Call REPLA(TTXX$, MKD$(VALOD1# + VALOD2#), 87, 8)
       '
       JJ& = JJ& + 1
       Call GRAREG("TIPIFIC", TTXX$, JJ&)
     Next CAMAI%
     '
     For JJI% = 1 To 256
       EXPLOLIN$(JJI%) = ""
       If JJI% <= CAMATI% Then
         TTXY$ = Space$(64)
         Call REPLA(TTXY$, CODEXT$(CIMATI%(JJI%)), 1, 16)
         Call REPLA(TTXY$, DESCRIT0$(CIMATI%(JJI%)), 17, 28)
         Call REPLA(TTXY$, UNIMED$(CIMATI%(JJI%)), 47, 2)
         Call REPLA(TTXY$, FORMATNUM$(VALOD1# + VALOD2#, "F12.4"), 49, 12)
         EXPLOLIN$(JJI%) = TTXY$
       End If
     Next JJI%
     CAIT% = CAMATI%
     '
     CODD$ = CODEXT$(CICON%)
     CABAS = 1
     Call EXPLOGR(CODD$, CABAS)
     '
   Next TU%
   '
   Call SETULTREG("TIPIFIC", JJ&)
   Call CIERRARCH("TIPIFIC")
   Call AGREMAPRIM(-1)    ' PARA MARCAR EOF
   Call SAVEFILE(LUDAT$ + "IMATPRI.EXC", String$(128, 0))
   Call CIERRARCH("*.*")
95 Screen.MousePointer = 1
   MsgBox "Importacion Completa"
   '
99 Call CIERRARCH("MASTER")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Command7.Enabled = True
   '
End Sub

Sub AGREMAPRIM(CIIX%)
   Static JMP&, CIAGRE%(8192)
   '
   If CIIX% < 0 Then
     Call SETULTREG("MATPRIM", JMP&)
     Call CIERRARCH("MATPRIM")
     Exit Sub
   End If
   '
   If CIIX% < 1 Then Exit Sub
   '
   For KMP& = 1 To JMP&
     If CIAGRE%(KMP&) = CIIX% Then Exit Sub
   Next KMP&
   '
   REBLA$ = REGBLAN$("MATPRIM")
   Call REPLA(REBLA$, MKI$(CIIX%), 1, 2)
   Call REPLA(REBLA$, DESCRIT0$(CIIX%), 3, 62)
   JMP& = JMP& + 1
   Call GRAREG("MATPRIM", REBLA$, JMP&)
   CIAGRE%(JMP&) = CIIX%
   '
End Sub

Private Sub Command9_Click()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     Text1.TEXT = COD$
   End If
   '
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    If EXISTE%(LUDAT$ + "IMATPRI.EXC") > 0 Then
      Command7.Visible = False
    End If
    '
    Call CENTRAFORM(Me)
    'SSTab1.Tab = 0
    'Call TRANSLABELS(Name)
    'Show
    'Refresh
    Screen.MousePointer = 11
    '
   
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    'Call GRATITFOR(Caption)
    '
    Refresh
    '
    Call CARTAVALISTO
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
    Screen.MousePointer = 1
    '
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   IMPLOSI.Hide
   On Error GoTo 0
   DoEvents
   If WindowState = 1 Then
       Call CIERRARCH("*.*")
       Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If MODOEDIT% = 1 Then
     Cancel = 1
     Call MENSAEDIT
     Exit Sub
   End If
   Call GRADATITEM
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Cancel = 0
   Hide
End Sub


Private Sub List1_DblClick()
   '
   If Command18.Enabled = True Then
     Call Command18_Click
   End If
   '
End Sub

Function VALCODART%()
   VALCODART% = 1
   If CODINT%(Text1.TEXT) < 1 Then
      Call MENSERR(24, "Falta Ingresar Código de Artículo")
      VALCODART% = 0
      Text1.SetFocus
   End If
End Function
'
Sub BLANFORMU()
   '
   Text3.TEXT = ""
   Label37.Caption = ""
   Label28.Caption = ""
   '
   '
   List1.Clear
   Text31 = ""
   Text32(0) = ""
   Text27 = ""
   Label40 = ""
   Label41 = ""
   '
   Label43 = ""
   Text11 = ""
   '
   Text2.TEXT = ""
   '
End Sub

Sub CARDATITEM()
    '
    CODD$ = Text1.TEXT
    CI% = CODINT%(CODD$)
    If CI% < 1 Then
       If TRIM$(CODD$) <> "" Then
          Call MENSERR(24, "Código Inexistente o no Válido")
       End If
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Text3.TEXT = DESCRIT0$(CI%)
    Label37.Caption = UNIMED$(CI%)
    Label28.Caption = Str$(CI%)
    '
    Label6 = FECHATEX$(FEREFOR%(CI%))
    Label8 = USREFOR$(CI%)
    '
    MODIFIC% = 0
    '
    Call CARFORMULA
    '
    Screen.MousePointer = 1
    '
End Sub

Sub CARFORMULA()
   '
   If YACAR% < 1 Then
      '
      CODD$ = Text1.TEXT
      CI% = CODINT%(CODD$)
      If CI% < 1 Then Exit Sub
      '
      Screen.MousePointer = 11
      RFF$ = RECFILT$("TIPIFIC", 1, MKI$(CI%))
      List1.Clear: JJ& = 0
      Text2 = "": Label6 = "": Label8 = ""
      '
      For U& = 1 To Len(RFF$) Step 4
        RETI& = CVS(Mid$(RFF$, U&, 4))
        XX$ = REGLEIDO$("TIPIFIC", RETI&)
        CICO% = CVI(Mid$(XX$, 49, 2))
        If CICO% > 0 Then
          If CICO% <= NUMAS% Then
            TTXX$ = Space$(128)
            Call REPLA(TTXX$, CODEXT$(CICO%), 1, 15)
            Call REPLA(TTXX$, COTIPIF$(CICO%), 17, 12)
            Call REPLA(TTXX$, DESCRIT0$(CICO%), 31, 16)
            Call REPLA(TTXX$, UNIMED$(CICO%), 48, 4)
            USONE# = CVD(Mid$(XX$, 55, 8))
            Call REPLA(TTXX$, FORMATNUM$(USONE#, "F11.6"), 52, 11)
            MERMAA# = CVD(Mid$(XX$, 63, 8))
            PERDID# = CVD(Mid$(XX$, 75, 8))
            USOTO# = CVD(Mid$(XX$, 87, 8))
            Call REPLA(TTXX$, FORMATNUM$(100 * MERMAA# / USOTO#, "F8.2"), 64, 8)
            Call REPLA(TTXX$, FORMATNUM$(100 * PERDID# / USOTO#, "F8.2"), 73, 8)
            Call REPLA(TTXX$, FORMATNUM$(USOTO#, "F11.6"), 83, 11)
            JJ& = JJ& + 1
            List1.AddItem TTXX$
            '
            If TRIM$(Text2) = "" Then Text2 = TRIM$(Mid$(XX$, 3, 22))
            If TRIM$(Label6) = "" Then Label6 = FECHATEX$(CVI(Mid$(XX$, 25, 2)))
            If TRIM$(Label8) = "" Then Label8 = TRIM$(Mid$(XX$, 27, 22))
            '
          End If
        End If
      Next U&
      '
      If TRIM$(Label6) = "" Then Label6 = FECHATEX$(HOY(HOS$))
      If TRIM$(Label8) = "" Then Label8 = USERNAME$
      '
      YACAR% = 1
      MODIFIC% = 0
      Screen.MousePointer = 1
      '
   End If
   '
End Sub
'
Sub GRADATITEM()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Then Exit Sub
   '
   If MODIFIC% > 0 Then
       Command9.SetFocus
       Call GRAGENITEM
       Call GRAFORMULA
   End If
   Screen.MousePointer = 1
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
End Sub
'
Sub GRAGENITEM()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Then Exit Sub
   If MODIFIC% < 1 Then Exit Sub
   '
   If DERACCE%("MODITMAS", "Modificación Datos Maestro de Stock") < 2 Then
      Exit Sub
   End If
   '
   TTXX$ = MKI$(CI%) + MKI$(HOII%) + USERNAME$
   Call FILTERPUTRECORD("REVIFORM", 1, MKI$(CI%), TTXX$)
   '
   If YACAR% < 1 Then    ' SOLO CAMBIO LA DESCRIPCION
      Screen.MousePointer = 11
      REG& = CI% + 1
      REMAS$ = REGLEIDO$("MASTER", REG&)
      '
      MODESCRI% = 0
      DE$ = TRIM$(Text3.TEXT)
      If DE$ = "" Then DE$ = TRIM$(Mid$(REMAS$, 5, 64))
      MODESCRI% = 1
      Call REPLA(REMAS$, DE$, 5, 64)
   End If
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se han Modificado Datos del Artículo\'" + TRIM$(CODEXT$(CI%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR% = 0
       MODIFIC% = 0
       Call CARDATITEM
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   '
   MODIFIC% = 0
   MODESCRI% = 0
   '
End Sub
'
Sub GRAFORMULA()
   '
   CIKKI% = CODINT%(Text1.TEXT)
   If CIKKI% < 1 Or CIKKI% > NUMAS% Then Exit Sub
   If YACAR% < 1 Then Exit Sub
   If MODIFIC% < 1 And GRABATODO% < 1 Then Exit Sub
   '
   If DERACCE%("MOFORTIP", "Modificación Fórmula Tipificada") < 2 Then
      Exit Sub
   End If
   '
   If GRABATODO% <> 1 Then
     RESPU% = COMALTER%("Se ha Modificado la Fórmula Tipificada del Artículo\'" + TRIM$(CODEXT$(CI%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar")
     If RESPU% <> 1 Then
       YACAR% = 0
       MODIFIC% = 0
       Call CARFORMULA
       GoTo 99
     End If
   End If
   '
   Screen.MousePointer = 11
   CAIT% = List1.ListCount
   '
   Dim EXPLOLI1$(256)
   CAIT1% = CAIT%
   For JJI% = 1 To 256
      EXPLOLI1$(JJI%) = ""
      If JJI% <= List1.ListCount Then
         EXPLOLI1$(JJI%) = Left$(List1.List(JJI% - 1), 44) + "     " + Mid$(List1.List(JJI% - 1), 83, 11)
      End If
   Next JJI%
   '
   For I& = 1 To CAIT%
     EXPLOLIN(I&) = List1.List(I& - 1)
     CICOMPO% = CODINT%(Left$(EXPLOLIN(I&), 16))
     Call EXPLOMUL(CICOMPO%)
     For JJI% = 1 To CONTAMUL%
       If CIK%(JJI%) = CI% Then
         Call MENSERR(24, "Imposible Registrar Esta Fórmula.\Error de Recursividad en Código " + TRIM$(CODEXT$(CI%)) + "\" + DESCRIT0$(CI%) + ".")
         YACAR% = 0
         MODIFIC% = 0
         SSTab1.Tab = 2
         Call CARFORMULA
         GoTo 99
       End If
     Next JJI%
   Next I&
   '
   CAIT% = CAIT1%
   For JJI% = 0 To 256
      EXPLOLIN$(JJI%) = EXPLOLI1$(JJI%)
   Next JJI%
   '
   CODD$ = AJUSTI$(Text1.TEXT, 16)
   CABAS = 1
   CAIT% = List1.ListCount
   '
   If CABAS > 0.01 Then
      Call EXPLOGR(CODD$, CABAS)
      MODIFIC% = 0
   End If
   '
   CODD$ = AJUSTI$(Text1.TEXT, 16)
   CABAS = 1
   CAIT% = List1.ListCount
   '
   CI% = CODINT%(Text1.TEXT)
   URN& = ULTREG&("TIPIFIC")
   RFF$ = RECFILT$("TIPIFIC", 1, MKI$(CI%))
   KK& = (-3)
   '
   For ITI% = 1 To CAIT%
     REFOR$ = REGBLAN$("TIPIFIC")
     TTXX$ = List1.List(ITI% - 1)
     Call REPLA(REFOR$, MKI$(CI%), 1, 2)
     Call REPLA(REFOR$, TRIM$(Text2), 3, 22)
     Call REPLA(REFOR$, MKI$(Int(FECHANUM(Label6))), 25, 2)
     Call REPLA(REFOR$, USERNAME$, 27, 22)
     '
     CICO% = CODINT%(Left$(TTXX$, 15))
     Call REPLA(REFOR$, MKI$(CICO%), 49, 2)
     Call REPLA(REFOR$, UNIMED$(CICO%), 51, 4)
     USONE# = Val(Mid$(TTXX$, 52, 11))
     Call REPLA(REFOR$, MKD$(USONE#), 55, 8)
     USOTO# = Val(Mid$(TTXX$, 83, 11))
     PORME = Val(Mid$(TTXX$, 64, 8))
     Call REPLA(REFOR$, MKD$(PORME * USOTO# / 100), 63, 8)
     Call REPLA(REFOR$, MKS$(PORME), 71, 4)
     PORPE = Val(Mid$(TTXX$, 73, 8))
     Call REPLA(REFOR$, MKD$(PORPE * USOTO# / 100), 75, 8)
     Call REPLA(REFOR$, MKS$(PORPE), 83, 4)
     Call REPLA(REFOR$, MKD$(USOTO#), 87, 8)
     '
     KK& = KK& + 4
     If KK& <= Len(RFF$) - 3 Then
         RENOTA& = CVS(Mid$(RFF$, KK&, 4))
         XY$ = REGLEIDO$("TIPIFIC", RENOTA&)
         If CVI(Left$(XY$, 2)) <> CI% Then
           Call MENSERR(24, "Error de Grabación de Registro de Tipificación.")
           URN& = URN& + 1
           RENOTA& = URN&
         End If
       Else
         URN& = URN& + 1
         RENOTA& = URN&
     End If
     Call GRAREG("TIPIFIC", REFOR$, RENOTA&)
   Next ITI%
   '
25 KK& = KK& + 4
   If KK& <= Len(RFF$) - 3 Then
     RENOTA& = CVS(Mid$(RFF$, KK&, 4))
     XY$ = REGLEIDO$("TIPIFIC", RENOTA&)
     If CVI(Left$(XY$, 2)) = CI% Then
       Call GRAREG("TIPIFIC", String$(256, 0), RENOTA&)
     End If
     GoTo 25
   End If
   '
   Call CIERRARCH("TIPIFIC")
   '
   If CABAS > 0.01 Then
      Call EXPLOGR(CODD$, CABAS)
      MODIFIC% = 0
   End If
   '
   Call BAJALDISCO
   '
99 Text31 = ""
   For UKI% = 0 To 2: Text32(UKI%) = "": Next UKI%
   Text27 = ""
   Label40 = ""
   Label41 = ""
   Command18.Enabled = True
   Command19.Enabled = True
   Screen.MousePointer = 1
   '
End Sub
'
Private Sub Text1_Change()
   '
   YACAR% = 0
   '
   Call BLANFORMU
   If CODINT%(Text1.TEXT) > 0 Then
       Call CARDATITEM
   End If
   '
End Sub

Private Sub Text1_DblClick()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     Text1.TEXT = COD$
   End If
   '
End Sub

Private Sub Text1_GotFocus()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
End Sub

Private Sub Text2_Change()
   MODIFIC% = 1
End Sub

Private Sub Text3_Change()
   MODIFIC% = 1
End Sub

Private Sub Text31_DblClick()
   '
   Call SELECHO("MASTER")
   COCOMPO$ = TRIM$(VALACT1$("MASTER"))
   If COCOMPO$ <> "" Then
     Text31.TEXT = COCOMPO$
     CICOMPO% = CODINT%(COCOMPO$)
     Label40.Caption = DESCRIT0$(CICOMPO%)
     Label41.Caption = UNIMED$(CICOMPO%)
     If TRIM$(Text32(0).TEXT) = "" Then Text32(0).TEXT = Str$(CABAS)
     Text32(0).SetFocus
   End If
   '
End Sub

Private Sub Text31_KeyPress(KeyAscii As Integer)
   '
   If KeyAscii = 13 Then
      CODY$ = TRIM$(Text31.TEXT)
      CIY% = CODINT%(CODY$)
      If CIY% > 0 Then
         CODE$ = CODY$: CIE% = CIY%
         Text31.TEXT = CODY$
         CODE$ = CODY$: CIE% = CIY%
         Label40.Caption = DESCRIT0$(CIE%)
         Label41.Caption = UNIMED$(CIE%)
         Text32(0).SetFocus
        Else
         Call MENSERR(24, "Código no Válido")
         Text31.TEXT = CODE$
      End If
   End If
   '
End Sub

Private Sub Text32_Change(Index As Integer)
   SUMATO# = 0
   For KKI% = 0 To 2
     SUMATO# = SUMATO# + Val(Text32(KKI%))
   Next KKI%
   Label2 = FORMATNUM$(SUMATO#, "F11.6")
End Sub

Private Sub Text32_GotFocus(Index As Integer)
   Text32(0).SelStart = 0
   Text32(0).SelLength = Len(Text32(0).TEXT)
End Sub

Sub ALTACODIGO(VDEF$, OKX%)
   '
   If OKX% < 0 Then
     OBX$ = VDEF$
     GoTo 11
   End If
   '
10 If OKX% < 0 Then GoTo 99
   OBX$ = OBJPLA$("ALTACODIGO", VDEF$)
11 If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   CODINUE$ = TRIM$(Mid$(OBX$, 77, 16))
   '
   If Len(CODINUE$) < 3 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Menor que 3 Caracteres.")
      GoTo 10
   End If
   '
   If Len(CODINUE$) > 15 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Mayor que 15 Caracteres.")
      GoTo 10
   End If
   '
   If CODINT%(CODINUE$) <> 0 Then
      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
      GoTo 10
   End If
   '
   If TRIM$(Mid$(OBX$, 5, 64)) = "" Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Descripcion")
      GoTo 10
   End If
   '
   TIMA% = Asc(Mid$(OBX$, 93, 1))
   If ECOARCH$("TATIMAT", Chr$(TIMA%)) = "" Then
      Call MENSERR(24, "Tipo de Material no Válido")
      GoTo 10
   End If
   '
   UNID$ = Mid$(OBX$, 109, 2)
   If ECOARCH$("UNIMED", UNID$) = "" Then
      Call MENSERR(24, "Unidad de Medida no Válida")
      GoTo 10
   End If
   '
   Screen.MousePointer = 11
   PUNCODIG$ = PUNCODIG$ + AJUSTI$(CODINUE$, 16)
   PPCC$ = PUNCODIG$
   JI% = 1
   While Len(PPCC$) > 2048
      PPCC$ = Mid$(PPCC$, 2049)
      JI% = JI% + 1
   Wend
   While Len(PPCC$) < 2048
      PPCC$ = PPCC$ + Chr$(0)
   Wend
   '
   NX% = FILEOPEN%(LUDAT$ + "PUNCODIG.DAT", 0, 2048)
   BUFERCOD$ = PPCC$
   Put #NX%, JI%, BUFERCOD$
   Close #NX%
   '
   REMAS& = NUMAS% + 2
   Call GRAREG("MASTER", OBX$, REMAS&)
   NMSI% = NUMAS% + 1
   X$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(X$, MKI$(NMSI%), 1, 2)
   Call GRAREG("MASTER", X$, 1)
   Call CIERRARCH("MASTER")
   '
   EL1$ = REGBLAN$("ECOMAST")
   Call REPLA(EL1$, CODINUE$, 1, 16)
   Call REPLA(EL1$, Mid$(OBX$, 5, 64), 17, 46)
   Call REPLA(EL1$, MKI$(NMSI%), 63, 2)
   '
   URECO& = ULTREG&("ECOMAST")
   While URECO& > 0
     XXX$ = REGLEIDO$("ECOMAST", URECO&)
     If TRIM$(XXX$) <> "" Then GoTo 25
     URECO& = URECO& - 1
   Wend
   '
25 LOTOTE& = 64 * URECO&
   Call CIERRARCH("ECOMAST")
   FIECO% = FILEOPEN%(LUDAT$ + "ECOMAST.DAT", 0, 2048)
   Dim EL As String * 2048
   TTXX$ = String$(LOTOTE&, 0)
   '
   For U& = 1 To 1 + LOF(FIECO%) / 2048
      Get #FIECO%, U&, EL$
      On Error Resume Next
      Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
      On Error GoTo 0
   Next U&
   TTXX$ = Left$(TTXX$, LOTOTE&)
   '
   For U& = 1 To Len(TTXX$) Step 64
      If Mid$(TTXX$, U&, 16) > CODINUE$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + EL1$ + Mid$(TTXX$, U&)
         GoTo 60
      End If
   Next U&
   TTXX$ = TTXX$ + EL1$
   '
60 JJ& = 0: LOTEGRA& = Len(TTXX$)
   TTXX$ = TTXX$ + String$(2048, 0)
   For U& = 1 To LOTEGRA& Step 2048
      EL$ = Mid$(TTXX$ + String$(2048, 0), U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   '
   NVII& = 1 + NUINV%
   EL1$ = REGBLAN$("INVERT")
   Call REPLA(EL1$, CODINUE$, 1, 16)
   Call REPLA(EL1$, MKI$(NMSI%), 17, 2)
   LOTOTE& = 18 * ULTREG&("INVERT")
   If LOTOTE& < 18 * NVII& Then LOTOTE& = 18 * NVII&
   FIECO% = FILEOPEN%(LUDAT$ + "INVERT.DAT", 0, 2048)
   TTXX$ = String$(LOTOTE&, 0)
   '
   For U& = 1 To 1 + LOF(FIECO%) / 2048
      Get #FIECO%, U&, EL$
      On Error Resume Next
      Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
      On Error GoTo 0
   Next U&
   TTXX$ = Left$(TTXX$, LOTOTE&)
   While Len(TTXX$) < LOTOTE&: TTXX$ = TTXX$ + Chr$(0): Wend
   '
   NVIJ% = NVII&
   Mid$(TTXX$, 17, 2) = MKI$(Int(NVIJ%))
   For U& = 19 To Len(TTXX$) Step 18
      If Mid$(TTXX$, U&, 16) > CODINUE$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + EL1$ + Mid$(TTXX$, U&)
         GoTo 70
      End If
   Next U&
   TTXX$ = TTXX$ + EL1$
   '
70 TTXX$ = TTXX$
   JJ& = 0
   For U& = 1 To Len(TTXX$) Step 2048
      EL$ = Mid$(TTXX$ + String$(2048, 0), U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   '
   Call CIERRARCH("ECOMAST")
   Call FRESHECHO("MASTER")
   Call BAJALDISCO
   Call REPUNCODIG
   OKX% = 1
   '
99 Screen.MousePointer = 1

End Sub
'
Sub COPYFORMU(CODORI%, CODEST%)
   '
   CODOR$ = TRIM$(CODEXT$(CODORI%))
   CODES$ = TRIM$(CODEXT$(CODEST%))
   '
   Screen.MousePointer = 11
   Call CIERRARCH("*.*")
   '
   J& = 0
   PUNTCI$ = ""
   ESP1$ = "": ESP2$ = "": ESP3$ = ""
   For I& = 1 To ULTREG&("ESTRUNUE")
      EST1$ = REGLEIDO$("ESTRUNUE", I&)
      CIV% = CVI(Left$(EST1$, 2))
      EST2$ = REGLEIDO$("ESTRUDOS", I&)
      If CIV% > 0 Then
        If CIV% <= NUMAS% Then
          If CIV% <> CODEST% Then
            J& = J& + 1
            Call GRAREG("ESTRUNUE", EST1$, J&)
            Call GRAREG("ESTRUDOS", EST2$, J&)
            PUNTCI$ = PUNTCI$ + MKI$(CIV%)
          End If
        End If
      End If
      '
      If CIV% = CODORI% Then
        If ESP1$ = "" Then
          ESP1$ = EST1$
          ESP2$ = EST2$
        End If
      End If
   Next I&
   '
   Call REPLA(ESP1$, MKI$(CODEST%), 1, 2)
   EST1$ = ESP1$
   EST2$ = ESP2$
   J& = J& + 1
   Call GRAREG("ESTRUNUE", EST1$, J&)
   Call GRAREG("ESTRUDOS", EST2$, J&)
   PUNTCI$ = PUNTCI$ + MKI$(CODEST%)
   '
   For I& = J& + 1 To ULTREG&("ESTRUNUE")
     Call GRAREG("ESTRUNUE", String$(128, 0), I&)
     Call GRAREG("ESTRUDOS", String$(128, 0), I&)
   Next I&
   '
   Call CIERRARCH("*.*")
   '
   J& = 0
   For I& = 1 To ULTREG&("ESTRULAR")
      EST3$ = REGLEIDO$("ESTRULAR", I&)
      CIV% = CVI(Left$(EST3$, 2))
      If CIV% > 0 Then
        If CIV% <= NUMAS% Then
          If CIV% <> CODEST% Then
            J& = J& + 1
            Call GRAREG("ESTRULAR", EST3$, J&)
          End If
        End If
      End If
      '
      If CIV% = CODORI% Then
        If ESP3$ = "" Then
          ESP3$ = EST3$
        End If
      End If
   Next I&
   '
   If ESP3$ <> "" Then
     Call REPLA(ESP3$, MKI$(CODEST%), 1, 2)
     EST3$ = ESP3$
     J& = J& + 1
     Call GRAREG("ESTRULAR", EST3$, J&)
   End If
   Call CIERRARCH("ESTRULAR")
   '
   Call BAJALDISCO
   '
   Screen.MousePointer = 1
   '
End Sub

Sub COPYTIPIF(CODORI%, CODEST%)
   '
   Screen.MousePointer = 11
   RFFY$ = RECFILT$("TIPIFIC", 1, MKI$(CODEST%))
   '
   If RFFY$ <> "" Then
     Call MENSERR(24, "ATENCION: Código Destino con Fórmula Tipificada.\   \Borre Fórmula de Destino\y vuelva a Intentar Copia.")
   End If
   '
   RFFX$ = RECFILT$("TIPIFIC", 1, MKI$(CODORI%))
   K1% = 0
   URE& = ULTREG&("TIPIFIC")
   For U& = 1 To Len(RFFX$) Step 4
     RESE& = CVS(Mid$(RFFX$, U&, 4))
     If RESE& > 0 Then
       If RESE& <= URE& Then
         X$ = REGLEIDO$("TIPIFIC", RESE&)
         Call REPLA(X$, MKI$(CODEST%), 1, 2)
         Call REGAPP("TIPIFIC", X$)
       End If
     End If
   Next U&
   '
   Call CIERRARCH("TIPIFIC")
   '
90 Screen.MousePointer = 1
   '
End Sub
'
Sub COPYRUTA(CODORI%, CODEST%)
   '
   Dim REGTX$(256)
   CODOR$ = TRIM$(CODEXT$(CODORI%))
   CODES$ = TRIM$(CODEXT$(CODEST%))
   '
   Screen.MousePointer = 11
   PRIREORI& = 0
   PRIREDEST& = 0
   '
   RFFX$ = RECFILT$("SECOPER", 1, MKI$(CODORI%))
   RFFY$ = RECFILT$("SECOPER", 1, MKI$(CODEST%))
   '
   K1% = 0
   URE& = ULTREG&("SECOPER")
   For U& = 1 To Len(RFFX$) Step 4
     RESE& = CVS(Mid$(RFFX$, U&, 4))
     If RESE& > 0 Then
       If RESE& <= URE& Then
         X$ = REGLEIDO$("SECOPER", RESE&)
         Call REPLA(X$, MKI$(CODEST%), 1, 2)
         K1% = K1% + 1
         REGTX$(K1%) = X$
       End If
     End If
   Next U&
   '
   For K3% = 1 To K1%
     If 4 * K3% <= Len(RFFY$) Then
         RESE& = CVS(Mid$(RFFY$, 4 * K3% - 3, 4))
       Else
         RESE& = 1 + ULTREG&("SECOPER")
     End If
     VALANT$ = REGLEIDO$("SECOPER", RESE&)
     If REGTX$(K3%) <> VALANT$ Then
       Call GRAREG("SECOPER", REGTX$(K3%), RESE&)
     End If
   Next K3%
   For K3% = 4 * (K1% + 1) - 3 To Len(RFFY$) Step 4
     RESE& = CVS(Mid$(RFFY$, K3%, 4))
     RGX$ = REGBLAN$("SECOPER")
     Call GRAREG("SECOPER", RGX$, RESE&)
   Next K3%
   '
   Call CIERRARCH("SECOPER")
   '
   RFFX$ = RECFILT$("OBSEOPER", 1, MKI$(CODORI%))
   RFFY$ = RECFILT$("OBSEOPER", 1, MKI$(CODEST%))
   '
   K1% = 0
   URE& = ULTREG&("OBSEOPER")
   For U& = 1 To Len(RFFX$) Step 4
     RESE& = CVS(Mid$(RFFX$, U&, 4))
     If RESE& > 0 Then
       If RESE& <= URE& Then
         X$ = REGLEIDO$("OBSEOPER", RESE&)
         Call REPLA(X$, MKI$(CODEST%), 1, 2)
         K1% = K1% + 1
         REGTX$(K1%) = X$
       End If
     End If
   Next U&
   '
   For K3% = 1 To K1%
     If 4 * K3% <= Len(RFFY$) Then
         RESE& = CVS(Mid$(RFFY$, 4 * K3% - 3, 4))
       Else
         RESE& = 1 + ULTREG&("OBSEOPER")
     End If
     VALANT$ = REGLEIDO$("OBSEOPER", RESE&)
     If REGTX$(K3%) <> VALANT$ Then
       Call GRAREG("OBSEOPER", REGTX$(K3%), RESE&)
     End If
   Next K3%
   For K3% = 4 * (K1% + 1) - 3 To Len(RFFY$) Step 4
     RESE& = CVS(Mid$(RFFY$, K3%, 4))
     RGX$ = REGBLAN$("OBSEOPER")
     Call GRAREG("OBSEOPER", RGX$, RESE&)
   Next K3%
   '
   Call CIERRARCH("OBSEOPER")
   '
   RFFX$ = RECFILT$("PLACONTR", 1, MKI$(CODORI%))
   RFFY$ = RECFILT$("PLACONTR", 1, MKI$(CODEST%))
   '
   K1% = 0
   URE& = ULTREG&("PLACONTR")
   For U& = 1 To Len(RFFX$) Step 4
     RESE& = CVS(Mid$(RFFX$, U&, 4))
     If RESE& > 0 Then
       If RESE& <= URE& Then
         X$ = REGLEIDO$("PLACONTR", RESE&)
         Call REPLA(X$, MKI$(CODEST%), 1, 2)
         K1% = K1% + 1
         REGTX$(K1%) = X$
       End If
     End If
   Next U&
   '
   For K3% = 1 To K1%
     If 4 * K3% <= Len(RFFY$) Then
         RESE& = CVS(Mid$(RFFY$, 4 * K3% - 3, 4))
       Else
         RESE& = 1 + ULTREG&("PLACONTR")
     End If
     VALANT$ = REGLEIDO$("PLACONTR", RESE&)
     If REGTX$(K3%) <> VALANT$ Then
       Call GRAREG("PLACONTR", REGTX$(K3%), RESE&)
     End If
   Next K3%
   For K3% = 4 * (K1% + 1) - 3 To Len(RFFY$) Step 4
     RESE& = CVS(Mid$(RFFY$, K3%, 4))
     RGX$ = REGBLAN$("PLACONTR")
     Call GRAREG("PLACONTR", RGX$, RESE&)
   Next K3%
   '
   Call CIERRARCH("PLACONTR")
   '
   RFFX$ = RECFILT$("PUEPUNTO", 1, MKI$(CODORI%))
   RFFY$ = RECFILT$("PUEPUNTO", 1, MKI$(CODEST%))
   '
   K1% = 0
   URE& = ULTREG&("PUEPUNTO")
   For U& = 1 To Len(RFFX$) Step 4
     RESE& = CVS(Mid$(RFFX$, U&, 4))
     If RESE& > 0 Then
       If RESE& <= URE& Then
         X$ = REGLEIDO$("PUEPUNTO", RESE&)
         Call REPLA(X$, MKI$(CODEST%), 1, 2)
         K1% = K1% + 1
         REGTX$(K1%) = X$
       End If
     End If
   Next U&
   '
   For K3% = 1 To K1%
     If 4 * K3% <= Len(RFFY$) Then
         RESE& = CVS(Mid$(RFFY$, 4 * K3% - 3, 4))
       Else
         RESE& = 1 + ULTREG&("PUEPUNTO")
     End If
     VALANT$ = REGLEIDO$("PUEPUNTO", RESE&)
     If REGTX$(K3%) <> VALANT$ Then
       Call GRAREG("PUEPUNTO", REGTX$(K3%), RESE&)
     End If
   Next K3%
   For K3% = 4 * (K1% + 1) - 3 To Len(RFFY$) Step 4
     RESE& = CVS(Mid$(RFFY$, K3%, 4))
     RGX$ = REGBLAN$("PUEPUNTO")
     Call GRAREG("PUEPUNTO", RGX$, RESE&)
   Next K3%
   '
   Call CIERRARCH("PUEPUNTO")
   '
   Screen.MousePointer = 1
   '
End Sub
'
Sub COPYSPEC(CODORI%, CODEST%)
   '
   Dim REGTX$(1024)
   Screen.MousePointer = 11
   RFFX$ = RECFILT$("ESPECIFI", 1, MKI$(CODORI%))
   RFFY$ = RECFILT$("ESPECIFI", 1, MKI$(CODEST%))
   '
   K1% = 0
   URE& = ULTREG&("ESPECIFI")
   For U& = 1 To Len(RFFX$) Step 4
     RESE& = CVS(Mid$(RFFX$, U&, 4))
     If RESE& > 0 Then
       If RESE& <= URE& Then
         X$ = REGLEIDO$("ESPECIFI", RESE&)
         Call REPLA(X$, MKI$(CODEST%), 1, 2)
         K1% = K1% + 1
         REGTX$(K1%) = X$
       End If
     End If
   Next U&
   '
   For K3% = 1 To K1%
     If 4 * K3% <= Len(RFFY$) Then
         RESE& = CVS(Mid$(RFFY$, 4 * K3% - 3, 4))
       Else
         RESE& = 1 + ULTREG&("ESPECIFI")
     End If
     VALANT$ = REGLEIDO$("ESPECIFI", RESE&)
     If REGTX$(K3%) <> VALANT$ Then
       Call GRAREG("ESPECIFI", REGTX$(K3%), RESE&)
     End If
   Next K3%
   For K3% = 4 * (K1% + 1) - 3 To Len(RFFY$) Step 4
     RESE& = CVS(Mid$(RFFY$, K3%, 4))
     RGX$ = REGBLAN$("ESPECIFI")
     Call GRAREG("ESPECIFI", RGX$, RESE&)
   Next K3%
   '
   Call CIERRARCH("ESPECIFI")
   Screen.MousePointer = 1
   '
End Sub
'
Sub REVIMASTER()                         ' revision del maestro
        '
        Screen.MousePointer = 11
        HAYINCO% = 0
        '
        Call BLOQARCH("MASTER")
        Call BLOQARCH("INVERT")
        Call BLOQARCH("INVENUE")
        '
        Dim CODISEL$(32767), CISEL%(32767)
        '
        If Left$(UCase$(EMPREAC$), 5) = "BOREN" Then
           ' ejecuta conversion de codigos
           FIN& = NUMAS%
           For I& = 1 To FIN&
              Call TRACE(20, I&, FIN&)
              X$ = REGLEIDO$("MASTER", I& + 1)
              CODES$ = Mid$(X$, 77, 16)
              BAJA% = 0
              If Mid$(CODES$, 16, 1) = Chr$(96) Then
                 CODES$ = Left$(CODES$, 15)
                 BAJA% = 1
              End If
              If Mid$(CODES$, 5, 7) = Space$(7) Then
                 CODIS$ = AJUSTI$(Left$(CODES$, 4) + "-" + TRIM$(Mid$(CODES$, 5)), 16)
                 If BAJA% = 1 Then Call REPLA(CODIS$, Chr$(96), 16, 1)
                 If TRIM$(Mid$(CODIS$, 6, 10)) = "" Then
                    Call REPLA(CODIS$, " ", 5, 1)
                 End If
                 Call REPLA(X$, CODIS$, 77, 16)
                 Call GRAREG("MASTER", X$, I& + 1)
              End If
           Next I&
           Call CIERRARCH("MASTER")
           '
           FIN& = NUINV%
           For I& = 1 To FIN&
              Call TRACE(20, I&, FIN&)
              X$ = REGLEIDO$("INVERT", I& + 1)
              CODES$ = Mid$(X$, 1, 16)
              BAJA% = 0
              If Mid$(CODES$, 16, 1) = Chr$(96) Then
                 CODES$ = Left$(CODES$, 15)
                 BAJA% = 1
              End If
              If Mid$(CODES$, 5, 7) = Space$(7) Then
                 CODIS$ = AJUSTI$(Left$(CODES$, 4) + "-" + TRIM$(Mid$(CODES$, 5)), 16)
                 If BAJA% = 1 Then Call REPLA(CODIS$, Chr$(96), 16, 1)
                 If TRIM$(Mid$(CODIS$, 6, 10)) = "" Then
                    Call REPLA(CODIS$, " ", 5, 1)
                 End If
                 Call REPLA(X$, CODIS$, 1, 16)
                 Call GRAREG("INVERT", X$, I& + 1)
              End If
           Next I&
           Call CIERRARCH("INVERT")
           '
        End If
        '
        NM% = ULTREG&("MASTER") - 1
        If CVI(Left$(REGLEIDO$("MASTER", 1&), 2)) <> NM% Then
              X$ = REGLEIDO$("MASTER", 1&)
              Call REPLA(X$, MKI$(NM%), 1, 2)
              Call GRAREG("MASTER", X$, 1&)
              Call COMUNI("Número de Registros\de Archivo Maestro:\  \Indice Restaurado.")
              HAYINCO% = 1
          Else
              Call COMUNI("Número de Registros\de Archivo Maestro:\  \Control O.K.")
        End If
        '
        Screen.MousePointer = 11
        NVL& = ULTREG&("INVERT") - 1
        FIN& = NVL&
        For II& = 1 To NVL&
            Call TRACE(20, II&, FIN&)
            X$ = REGLEIDO$("INVERT", II& + 1)
            COD$ = Left$(X$, 16): CII% = CVI(Mid$(X$, 17, 2))
            If Mid$(COD$, 16, 1) = Chr$(96) Then
               If CII% > 0 Then
                  If CII% <= NUMAS% Then
                     REGMA& = CII% + 1
                     Y$ = REGLEIDO$("MASTER", REGMA&)
                     If Mid$(Y$, 77, 15) = Left$(COD$, 15) Then
                        Call REPLA(Y$, Chr$(96), 92, 1)
                        Call GRAREG("MASTER", Y$, REGMA&)
                     End If
                  End If
               End If
            End If
        Next II&
        '
        Screen.MousePointer = 11
        FORMALTER.LINOSORT.Clear
        FORMALTER.LISORT.Clear
        FIN& = NM%
        RECUMAX& = 0
        For II& = 1 To NM%
            Call TRACE(20, II&, FIN&)
            X$ = REGLEIDO$("MASTER", II& + 1)
            If CVI(Left$(X$, 2)) <> II& Then
                Call COMUNI("Error de Indice en Maestro de Artículos\ha sido Recuperado Automáticamente.\  \Oportunamente Consulte al Soporte de Sistemas.")
                Screen.MousePointer = 11
                cix% = II&
                Call REPLA(X$, MKI$(cix%), 1, 2)
                Call GRAREG("MASTER", X$, II& + 1)
            End If
            COD$ = AJUSTI$(TRIM$(Mid$(X$, 77, 15)), 15) + Mid$(X$, 92, 1)
            If Left$(COD$, 4) = "zzz-" Then
              If Val(Mid$(COD$, 5, 6)) > RECUMAX& Then
                RECUMAX& = Val(Mid$(COD$, 5, 6))
              End If
            End If
            '
            ZZ$ = "A" + COD$ + CSTRING$(MKS$(II&), 3, 7, 0, 2)
            FORMALTER.LISORT.AddItem ZZ$
        Next II&
        '
        FIN& = FORMALTER.LISORT.ListCount
        For II& = 1 To FORMALTER.LISORT.ListCount
           Call TRACE(20, II&, FIN&)
           ZZ$ = FORMALTER.LISORT.List(II& - 1)
           Call AGRELISCOD(ZZ$)
        Next II&
        '
        NVL& = ULTREG&("INVERT") - 1
        If NVL& < 1 Then
             Call REGAPP("INVERT", String$(16, 0) + MKI$(0))
             Call CIERRARCH("INVERT")
        End If
        '
        NV% = FORMALTER.LINOSORT.ListCount
        X$ = REGBLAN$("INVERT")
        Call REPLA(X$, MKI$(NV%), 17, 2)
        Call GRAREG("INVERT", X$, 1)
        '
        J& = 1
        FIN& = NV% - 1
        For II& = 0 To FORMALTER.LINOSORT.ListCount - 1
            Call TRACE(20, II&, FIN&)
            FORMALTER.LINOSORT.ListIndex = II&
            CODISEL$(II&) = Mid$(FORMALTER.LINOSORT.TEXT, 2, 16)
            CISEL%(II&) = Val(Mid$(FORMALTER.LINOSORT.TEXT, 18))
            X$ = REGBLAN$("INVERT")
            Call REPLA(X$, CODISEL$(II&), 1, 16)
            Call REPLA(X$, MKI$(CISEL%(II&)), 17, 2)
            Call GRAREG("INVERT", X$, II& + 2)
        Next II&
        Call SETULTREG("INVERT", 1& + NV%)
        '
        Call CIERRARCH("*.*")
        Call BAJALDISCO
        '
        CODA$ = ""
        FIN& = NV%
        For II& = 1 To NV%
           Call TRACE(20, II&, FIN&)
           IK$ = REGLEIDO$("INVERT", II& + 1)
           CODL$ = Left$(IK$, 15)
           If TRIM$(CODL$) = "" Or CODL$ = CODA$ Then
                'If TRIM$(CODL$) = "" Then GoTo 19
                REPLAI% = REPLAI% + 1
                'CODNUE$ = TRIM$(CODL$) + "-" + Chr$(64 + REPLAI%)
                CODNUE$ = "zzz-" + TRIM$(Str$(RECUMAX& + 1))
                RECUMAX& = RECUMAX& + 1
                If Left$(CODNUE$, 1) = "-" Then CODNUE$ = "0" + CODNUE$
                RGM& = 1 + CVI(Mid$(IK$, 17, 2))
                M$ = REGLEIDO$("MASTER", RGM&)
                Call REPLA(M$, CODNUE$, 77, 16)
                Call GRAREG("MASTER", M$, RGM&)
                Call REPLA(IK$, CODNUE$, 1, 15)
                Call REPLA(IK$, Chr$(96), 16, 1)
                Call GRAREG("INVERT", IK$, II& + 1)
                Call COMUNI("Codigo '" + TRIM$(CODL$) + "' Duplicado o no Válido.\  \Reemplazado por '" + TRIM$(CODNUE$) + "'.")
                Screen.MousePointer = 11
                HAYINCO% = 1
              Else
                REPLAI% = 0
           End If
           '
           If REPLAI% = 0 Then CODA$ = CODL$
19      Next II&
        '
        Call CIERRARCH("*.*")
        Call BLANARCH("INVENUE")
        '
17000   'Call BLIMESS("Actualizando Archivos de Validación ...")
        '
        'Call BLOQARCH("ECOMAST")
        FIECO% = FILEOPEN%(LUDAT$ + "ECOMAST.DAT", 0, 64)
        Dim ECM As String * 64
        '
        'Call BLOQARCH("CIACTI")
        FICIA% = FILEOPEN%(LUDAT$ + "CIACTI.DAT", 0, 2)
        Dim CCT As String * 2
        '
        NV% = ULTREG&("INVERT") - 1
        CCT$ = MKI$(NV%)
        Put #FICIA%, 1, CCT$
        '
        K% = 0
        FIN& = NV%
        For XL& = 1 To NV%
             Call TRACE(20, XL&, FIN&)
             IK$ = REGLEIDO$("INVERT", XL& + 1)
             If Mid$(IK$, 16, 1) <> Chr$(96) Then
               If TRIM$(Left$(IK$, 16)) <> "" Then
                 '
                 CI% = CVI(Mid$(IK$, 17, 2))
                 DE$ = DESCRIT$(CI%)
                 '
                 EC$ = Space$(62) + MKI$(CI%)
                 Call REPLA(EC$, Left$(IK$, 16), 1, 16)
                 Call REPLA(EC$, DE$, 17, 46)
                 '
                 K% = K% + 1: ECM$ = EC$: Put #FIECO%, K%, ECM$
                              CCT$ = MKI$(CI%): Put #FICIA%, K% + 1, CCT$
                 '
               End If
             End If
        Next XL&
        '
        CCT$ = MKI$(K%): Put #FICIA%, 1, CCT$
        '
        For K1% = K% + 1 To LOF(FIECO%) / 64
           ECM$ = String$(64, 0): Put #FIECO%, K1%, ECM$
           CCT$ = String$(2, 0): Put #FICIA%, K1% + 1, CCT$
        Next K1%
        '
        Close #FIECO%: FIECO% = 0 'Call CIERRARCH("ECOMAST")
        Close #FICIA%: FICIA% = 0 'Call CIERRARCH("CIACTI")
        FIECO% = 0: FICIA% = 0
        '
        Call GRACONTROL("", "AMASTO", "NORMAL")
        '
        Call CIERRARCH("*.*")
        Call LIBARCH("MASTER")
        '
        Call FRESHECHO("MASTER")
        Call FRESHECHO("ECOMAST")
        '
        'Call BLIMESS("-")
        '
        Screen.MousePointer = 1
        If HAYINCO% = 0 Then
            Call COMUNI("No se han Detectado Inconsistencias\en la Revisión del Maestro de Items.")
        End If
        Exit Sub
        '
        '
12500   If ULTREG&("INVERT") < 1 Then
             Call REGAPP("INVERT", String$(15, 0) + Chr$(1) + MKI$(0))
             Call CIERRARCH("INVERT")
        End If
        '
        NV% = ULTREG&("INVERT") - 1
        If CVI(Mid$(REGLEIDO$("INVERT", 1&), 17, 2)) <> NV% Then
            Call COMUNI("Posible Inconsistencia (1) !!!\Ejecute Opción de Control.")
            Call FINAL("")    ' GoTo 2000
        End If
        '
        ABRE$ = AJUSTI$(COD$, 15)
        LI& = 1: LS& = NV% + 2
        ENCU% = 0
        '
12510   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 12580
        L& = LI& + E&: If L& < 2 Or L& > NV% + 1 Then GoTo 12580
        IK$ = REGLEIDO$("INVERT", L&): I1$ = Left$(IK$, 15)
                                  If I1$ < ABRE$ Then LI& = L&: GoTo 12510
                                  If I1$ > ABRE$ Then LS& = L&: GoTo 12510
        ENCU% = 1: CI% = CVI(Mid$(IK$, 17, 2)): RGI& = L&
                If Mid$(IK$, 16, 1) = Chr$(96) Then
                    ENCU% = (-1)
                End If
        '
12580   Return
        '
        '
End Sub

Sub AGRELISCOD(RAGRE$)
   '
   ULIS& = FORMALTER.LINOSORT.ListCount
   If ULIS& < 1 Or Left$(RAGRE$, 17) > FORMALTER.LINOSORT.List(ULIS& - 1) Then
       FORMALTER.LINOSORT.AddItem RAGRE$
       GoTo 10
     Else
       For J& = ULIS& - 1 To 0 Step -1
         If Left$(FORMALTER.LINOSORT.List(J&), 17) <= Left$(RAGRE$, 17) Then
           FORMALTER.LINOSORT.AddItem RAGRE$, J& + 1
           GoTo 10
         End If
       Next J&
       FORMALTER.LINOSORT.AddItem RAGRE$, 0
10 End If
   '
End Sub

Sub CARTAVALISTO()
    '
10  If ULTREG&("SUCURSAL") < 1 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Falta Definir Depósitos")
        'OPMAST00.Show 1
        Call CONECRUN("FLSETUP/OPMAST00", "Configuración de Funcionamiento")
        GoTo 10
    End If
    '
    If ULTREG&("UMEDIDA") < 2 Then
        Call BLANARCH("UMEDIDA")
        Call REGAPP("UMEDIDA", "U.Unidades")
        Call REGAPP("UMEDIDA", "KgKilogramos")
        Call REGAPP("UMEDIDA", "grGramos")
        Call REGAPP("UMEDIDA", "gKGramos - Kilos")
        Call REGAPP("UMEDIDA", "M.Metros")
        Call REGAPP("UMEDIDA", "M2M.Cuadrados")
        Call REGAPP("UMEDIDA", "M3M.Cúbicos")
        Call REGAPP("UMEDIDA", "mmMilímetros")
        Call REGAPP("UMEDIDA", "cmCentímetros")
        Call REGAPP("UMEDIDA", "LtLitros")
        Call CIERRARCH("UMEDIDA")
    End If
    XX$ = REGLEIDO$("UMEDIDA", 1)
    '
    If ULTREG&("UNIMED") < 2 Then
        Call BLANARCH("UNIMED")
        Call REGAPP("UNIMED", "U.  Unidades" + String$(256, 0))
        Call REGAPP("UNIMED", "Kg  Kilogramos" + String$(256, 0))
        Call REGAPP("UNIMED", "gr  Gramos" + String$(256, 0))
        Call REGAPP("UNIMED", "gK  Gramos - Kilos" + String$(256, 0))
        Call REGAPP("UNIMED", "M.  Metros" + String$(256, 0))
        Call REGAPP("UNIMED", "M2  M.Cuadrados" + String$(256, 0))
        Call REGAPP("UNIMED", "M3  M.Cúbicos" + String$(256, 0))
        Call REGAPP("UNIMED", "mm  Milímetros" + String$(256, 0))
        Call REGAPP("UNIMED", "cm  Centímetros" + String$(256, 0))
        Call REGAPP("UNIMED", "Lt  Litros" + String$(256, 0))
        Call CIERRARCH("UNIMED")
    End If
    XX$ = REGLEIDO$("UNIMED", 1)
    '
    If ULTREG&("TATIMAT") < 1 Then
        Call BLANARCH("TATIMAT")
        Call REGAPP("TATIMAT", Chr$(1) + "Materia Prima")
        Call REGAPP("TATIMAT", Chr$(2) + "Pieza Comercial")
        Call REGAPP("TATIMAT", Chr$(3) + "Pieza Elaborada")
        Call REGAPP("TATIMAT", Chr$(4) + "Semielaborado")
        Call REGAPP("TATIMAT", Chr$(5) + "Prod.Terminado")
        Call REGAPP("TATIMAT", Chr$(8) + "Servicios")
        Call REGAPP("TATIMAT", Chr$(9) + "Improductivo")
        Call CIERRARCH("TATIMAT")
    End If
    XX$ = REGLEIDO$("TATIMAT", 1)
    '
    If ULTREG&("TABTIMA") < 1 Then
       DPPRE% = Asc(Left$(REGLEIDO$("ECODEP", 1), 1))
       For U& = 1 To ULTREG&("TATIMAT")
          X$ = REGLEIDO$("TATIMAT", U&)
          Y$ = REGBLAN$("TABTIMA")
          Call REPLA(Y$, X$, 1, 16)
          Call REPLA(Y$, Chr$(DPPRE%), 17, 1)
          Call REGAPP("TABTIMA", Y$)
       Next U&
       Call CIERRARCH("TABTIMA")
    End If
    XX$ = REGLEIDO$("TABTIMA", 1)
    '
    If ULTREG&("TACRIRE") < 4 Then
        Call BLANARCH("TACRIRE")
        Call REGAPP("TACRIRE", Chr$(1) + "Por Pedido")
        Call REGAPP("TACRIRE", Chr$(2) + "Por Mínimos")
        Call REGAPP("TACRIRE", Chr$(3) + "Cons.Estadístico")
        Call REGAPP("TACRIRE", Chr$(4) + "Programación")
        Call CIERRARCH("TACRIRE")
    End If
    XX$ = REGLEIDO$("TACRIRE", 1)
    '
    If ULTREG&("TAMONED") < 2 Then
        Call BLANARCH("TAMONED")
        Call REGAPP("TAMONED", Chr$(1) + AJUSTI$("Pesos", 27) + MKS$(1))
        Call REGAPP("TAMONED", Chr$(2) + AJUSTI$("Dolares U$S", 27) + MKS$(1))
        Call CIERRARCH("TAMONED")
    End If
    XX$ = REGLEIDO$("TAMONED", 1)
    '
    If ULTREG&("TACOMPO") < 2 Then
        Call BLANARCH("TACOMPO")
        Call REGAPP("TACOMPO", Chr$(0) + "Manual")
        Call REGAPP("TACOMPO", Chr$(1) + "Automática")
        Call CIERRARCH("TACOMPO")
    End If
    XX$ = REGLEIDO$("TACOMPO", 1)
    '
    If ULTREG&("GRUPIVA") < 1 Then
        Call SETULTREG("GRUPIVA", 0)
        Call REGAPP("GRUPIVA", Chr$(0) + "IVA 21 %   " + MKS$(21))
        Call REGAPP("GRUPIVA", Chr$(1) + "IVA 10,5 % " + MKS$(10.5))
        Call REGAPP("GRUPIVA", Chr$(2) + "IVA 27 %   " + MKS$(27))
        Call CIERRARCH("GRUPIVA")
    End If
    XX$ = REGLEIDO$("GRUPIVA", 1)
    '
    Screen.MousePointer = 1
    '
End Sub
'


Sub COPYSTRUC()
   '
10 CICJ$ = OBJPLA$("DEACODIGO", CICJ0$)
   If Len(CICJ$) < 4 Then
     GoTo 99
   End If
   '
   CODORI% = CVI(Left$(CICJ$, 2))
   CODEST% = CVI(Mid$(CICJ$, 3, 2))
   CICJ0$ = CICJ$
   '
   CODOR$ = TRIM$(CODEXT$(CODORI%))
   CODES$ = TRIM$(CODEXT$(CODEST%))
   '
   If CODOR$ = "" Then
     Call MENSERR(24, "Codigo Origen Inexistente")
     GoTo 10
   End If
   '
   If CODES$ = "" Then
     Call MENSERR(24, "Codigo Destino Inexistente")
     GoTo 10
   End If
   '
   Screen.MousePointer = 11
   Call LEEEXPLO(CODOR$, 1)
   If CAIT% < 1 Then
     Screen.MousePointer = 1
     OPX% = COMALTER%("Copia Imposible.\Código Origen " + TRIM$(CODOR$) + " sin Estructura.\\Repetir\Abandonar")
     If OPX% = 1 Then GoTo 10
     GoTo 99
   End If
   '
   Call LEEEXPLO(CODES$, 1)
   If CAIT% > 0 Then
     Screen.MousePointer = 1
     OPX% = COMALTER%("ATENCION: Código Destino " + TRIM$(CODES$) + " con Estructura.\\Cancelar\Sobre-escribir")
     If OPX% <> 2 Then GoTo 10
   End If
   '
   Call COPYFORMU(CODORI%, CODEST%)
   '
   Call CIERRARCH("ESTRUNUE")
   Call CIERRARCH("ESTRUDOS")
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
   Text1.TEXT = CODES$
   Text1.Refresh
   '
99 End Sub

Sub COPYRUTAS()
   '
Dim REG&(256), REGTX$(256)
   '
10 Screen.MousePointer = 1
   CICJ$ = OBJPLA$("DEACODIGO", CICJ0$)
   If Len(CICJ$) < 4 Then
     Call CIERRARCH("*.*")
     Exit Sub
   End If
   '
   CODORI% = CVI(Left$(CICJ$, 2))
   CODEST% = CVI(Mid$(CICJ$, 3, 2))
   CICJ0$ = CICJ$
   '
   CODOR$ = TRIM$(CODEXT$(CODORI%))
   CODES$ = TRIM$(CODEXT$(CODEST%))
   '
   If CODOR$ = "" Then
     Call MENSERR(24, "Codigo Origen Inexistente")
     GoTo 10
   End If
   '
   If CODES$ = "" Then
     Call MENSERR(24, "Codigo Destino Inexistente")
     GoTo 10
   End If
   '
   Call BLOQARCH("SECOPER")
   Screen.MousePointer = 11
   PRIREORI& = 0
   PRIREDEST& = 0
   '
   RFFX$ = RECFILT$("SECOPER", 1, MKI$(CODORI%))
   If Len(RFFX$) < 4 Then
     Screen.MousePointer = 1
     Call COMUNI("Copia Imposible.\Código Origen sin Operaciones.")
     GoTo 10
   End If
   '
   RFFY$ = RECFILT$("SECOPER", 1, MKI$(CODEST%))
   If Len(RFFY$) >= 4 Then
     Screen.MousePointer = 1
     OPX% = COMALTER%("ATENCION: Código Destino con Operaciones !!!\\Cancelar\Sobre-escribir")
     If OPX% <> 2 Then GoTo 10
   End If
   '
   Call COPYRUTA(CODORI%, CODEST%)
   '
   Call CIERRARCH("SECOPER")
   '
   Text1.TEXT = CODES$
   Text1.Refresh
   Screen.MousePointer = 1
   '
End Sub
'
Sub LISTRUTAS()
   '
   XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
   XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
   VDEF$ = MKI$(CI1%) + MKI$(CI2%)
   X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
   If Len(X$) < 4 Then
     Exit Sub
   End If
   '
   Screen.MousePointer = 11
   Dim PUNCOD$(32767)
   '
   CI1% = CVI(Left$(X$, 2))
   cod1$ = CODEXT$(CI1%): If TRIM$(cod1$) = "" Then Exit Sub
   CI2% = CVI(Mid$(X$, 3, 2))
   cod2$ = CODEXT$(CI2%): If TRIM$(cod2$) = "" Then Exit Sub
   '
   V& = 0
   FIN& = ULTREG&("SECOPER")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("SECOPER", U&)
     CI0% = CVI(Left$(X$, 2))
     If CI0% > 0 Then
       If CI0% <= NUMAS% Then
         PUNCOD$(CI0%) = PUNCOD$(CI0%) + MKS$(U&)
       End If
     End If
   Next U&
   '
   FIN& = NUINV%
   For IU& = 1 To NUINV%
     Call TRACE(20, IU&, FIN&)
     Z$ = REGLEIDO$("INVERT", IU& + 1)
     CODEX$ = Left$(Z$, 16)
     If Mid$(Z$, 16, 1) <> "`" Then
       CODEX$ = TRIM$(CODEX$)
       If CODEX$ >= cod1$ Then
         If CODEX$ <= cod2$ Then
           CI0% = CVI(Mid$(Z$, 17, 2))
           If Len(PUNCOD$(CI0%)) >= 4 Then
              '
              LOSTAND = LOTESTAN(CI0%)
              DESCRI$ = TRIM$("Lt.Std." + TRIM$(Str$(LOSTAND)) + " " + TRIM$(UNIMED$(CI0%)))
              DESCRI$ = " (" + DESCRI$ + ")"
                LOMADE% = 60 - Len(DESCRI$)
                If LOMADE% > 0 Then
                  DESCRI$ = Left$(TRIM$(DESCRIT0$(CI0%)), LOMADE%) + DESCRI$
                End If
                '
              Y$ = REGBLAN$("SECOPSEL")
              Call REPLA(Y$, MKI$(CI0%), 1, 2)
              Call REPLA(Y$, DESCRI$, 3, 60)
              V& = V& + 1
              Call GRAREG("SECOPSEL", Y$, V&)
              '
              Call REPLA(Y$, Space$(60), 3, 60)
              V& = V& + 1
              Call GRAREG("SECOPSEL", Y$, V&)
              '
              For IV% = 1 To Len(PUNCOD$(CI0%)) Step 4
                U& = CVS(Mid$(PUNCOD$(CI0%), IV%, 4))
                If U& > 0 Then
                  If U& <= ULTREG&("SECOPER") Then
                    Y1$ = REGLEIDO$("SECOPER", U&)
                    NOPES$ = TRIM$(CSTRING$(Mid$(Y1$, 3, 2), 1, 6, 0, 2))
                    While Len(NOPES$) < 4: NOPES$ = "." + NOPES$: Wend
                    DOPES$ = Mid$(Y1$, 5, 62)
                    TVAR = CVS(Mid$(Y1$, 67, 4)) * LOSTAND
                    Y$ = Y1$
                    Call REPLA(Y$, NOPES$, 3, 6)
                    Call REPLA(Y$, DOPES$, 9, 62)
                    Call REPLA(Y$, MKS$(TVAR), 71, 4)
                    V& = V& + 1
                    Call GRAREG("SECOPSEL", Y$, V&)
                  End If
                End If
              Next IV%
              '
              Call REPLA(Y$, String$(126, 0), 3, 126)
              V& = V& + 1
              Call GRAREG("SECOPSEL", Y$, V&)
              '
              Y$ = REGBLAN$("SECOPSEL")
              V& = V& + 1
              Call GRAREG("SECOPSEL", Y$, V&)
              '
           End If
         End If
       End If
     End If
   Next IU&
   '
   Call SETULTREG("SECOPSEL", V&)
   Call CIERRARCH("SECOPER")
   Call CIERRARCH("SECOPSEL")
   Screen.MousePointer = 1
   Call FINAL("*LISECOP")
   '
End Sub

Sub REDESPIE()
    '
    CAERR% = 0
    FIN& = ULTREG&("ESTRUNUE")
    FINDOS& = ULTREG&("ESTRUDOS")
    '
    Screen.MousePointer = 11
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      R0$ = REGLEIDO$("ESTRUNUE", I&)
      R1$ = Left$(R0$, 2)
      DPS$ = Mid$(R0$, 3, 120)
      If I& <= FINDOS& Then
        DPS$ = DPS$ + REGLEIDO$("ESTRUDOS", I&)
        CI1% = CVI(R1$)
        If CI1% > 0 Then
          If CI1% <= NUMAS% Then
            A$ = CODEXT$(CI1%) + " " + DESCRIT$(CI1%)
            '
            J% = 0
            For K% = 1 To Len(DPS$) Step 8
              DPSX$ = Mid$(DPS$, K%, 8)
              CI2% = CVI(Left$(DPSX$, 2))
              If CI2% > 0 Then
                If CI2% <= NUMAS% Then
                  J% = J% + 1
                  CII1%(J%) = CI2%
                End If
              End If
            Next K%

            For K% = 1 To J%
              If CII1%(K%) = CI1% Then
                MENSA$ = "Recursividad en Código " + TRIM$(CODEXT$(CI1%)) + "\" + DESCRIT0$(CI1%)
                If COMALTER%(MENSA$ + "\\Continuar\Cancelar") = 2 Then GoTo 99
              End If
              '
              For K1% = K% + 1 To J%
                If CII1%(K1%) = CII1%(K%) Then
                  CAERR% = CAERR% + 1
                  cod1$ = TRIM$(CODEXT$(CI1%))
                  cod2$ = TRIM$(CODEXT$(CII1%(K%)))
                  If CODINT%(cod1$) = CI1% Then ' SUPRIME ITEMS DE BAJA
                    MENSA$ = "Repetición del Código " + cod2$ + "\en Fórmula de Codigo " + cod1$
                    If COMALTER%(MENSA$ + "\\Continuar\Cancelar") = 2 Then GoTo 99
                  End If
                  GoTo 89
                End If
              Next K1%
            Next K%
          End If
        End If
      End If
89  Next I&
    '
    If CAERR% < 1 Then
        Call COMUNI("Control O.K. - No se Detectaron Repeticiones")
      Else
        Call MENSERR(20, "Se Encontraron" + Str$(CAERR%) + " Repeticiones en Despieces")
    End If
    '
99  Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
End Sub


Sub LIESTRU()      ' listado estructuras y enlaces por codigo
    '
    CODACT$ = Text1.TEXT
    CIACT% = CODINT%(Text1.TEXT)
    If CIACT% > 0 Then
         CI1% = CIACT%: CI2% = CIACT%
       Else
         XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
         XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    End If
    X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
    If Len(X$) < 4 Then
        Exit Sub
    End If
    CI1% = CVI(Left$(X$, 2)): cod1$ = CODEXT$(CI1%): If TRIM$(cod1$) = "" Then Exit Sub
    CI2% = CVI(Mid$(X$, 3, 2)): cod2$ = CODEXT$(CI2%): If TRIM$(cod2$) = "" Then Exit Sub
    '
    Call HEADERS("LITIPIF", "")
    Call HEADERS("LITIPIF", "Desde Codigo: " + TRIM$(cod1$) + " - " + DESCRIT$(CI1%))
    Call HEADERS("LITIPIF", "Hasta Codigo: " + TRIM$(cod2$) + " - " + DESCRIT$(CI2%))
    '
    Screen.MousePointer = 11
    NESTRU& = ULTREG&("TIPIFIC")
    J& = 0
    '
    For I& = 1 To NESTRU&
      Call TRACE(20, I&, NESTRU&)
      XX$ = REGLEIDO$("TIPIFIC", I&)
      CI% = CVI(Left$(XX$, 2))
      ACTIX% = 0
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          CODYY$ = CODEXT$(CI%)
          If Mid$(CODYY$, 16, 1) <> Chr$(96) Then ' SUPRIME ITEMS DE BAJA
            If CODYY$ >= cod1$ Then
              If CODYY$ <= cod2$ Then
                CICO% = CVI(Mid$(XX$, 49, 2))
                Call REPLA(XX$, COTIPIF$(CICO%), 27, 22)
                J& = J& + 1
                Call GRAREG("LITIPIF", XX$, J&)
              End If
            End If
          End If
        End If
      End If
    Next I&
    '
    Call SETULTREG("LITIPIF", J&)
    Call CIERRARCH("LITIPIF")
    '
    Call FILESORT("LITIPIF", "", 1, 1)
    Screen.MousePointer = 1
    Call FINAL("*LITIPIF")
    '
End Sub
'

Sub LIMULTI()
    '
    CODACT$ = Text1.TEXT
    CIACT% = CODINT%(Text1.TEXT)
    If CIACT% > 0 Then
         CI1% = CIACT%: CI2% = CIACT%
       Else
         XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
         XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    End If
    X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
    If Len(X$) < 4 Then
        Exit Sub
    End If
    '
    CI1% = CVI(Left$(X$, 2)): cod1$ = CODEXT$(CI1%): If TRIM$(cod1$) = "" Then Exit Sub
    CI2% = CVI(Mid$(X$, 3, 2)): cod2$ = CODEXT$(CI2%): If TRIM$(cod2$) = "" Then Exit Sub
    '
    Screen.MousePointer = 11
    Dim ACTI%(16384)
    '
    NM% = NUMAS%
    FIN& = ULTREG&("ESTRUNUE")
    For REG& = 1 To FIN&
        '
        Call TRACE(20, REG&, FIN&)
        X$ = REGLEIDO$("ESTRUNUE", REG&)
        CI% = CVI(Left$(X$, 2) + String$(2, 0))
        If CI% > 0 Then
            If CI% <= NUMAS% Then
               CODYY$ = CODEXT$(CI%)
               If Mid$(CODYY$, 16, 1) <> Chr$(96) Then
                  If ACTI%(CI%) = 0 Then
                     ACTI%(CI%) = 1
                  End If
               End If
            End If
        End If
        '
        DPS$ = Mid$(X$, 3, 120)
        If ULTREG&("ESTRUDOS") >= REG& Then
             DPS$ = DPS$ + REGLEIDO$("ESTRUDOS", REG&)
        End If
        '
        For K% = 1 To Len(DPS$) Step 8
            DPSX$ = Mid$(DPS$, K%, 8)
            CI% = CVI(Left$(DPSX$, 2))
            If CI% <= 0 Then GoTo 99
            If CI% <= NUMAS% Then
                ACTI%(CI%) = (-1)
            End If
        Next K%
        '
99  Next REG&
    '
    REGQ& = 0
'    FORMULA.List2.Clear
    '
    FINV& = NUINV%
    For REG& = 1 To NUINV%
        '
        Call TRACE(20, REG&, FINV&)
        X$ = REGLEIDO$("INVERT", REG&)
        CODYY$ = Left$(X$, 16)
        If CODYY$ >= cod1$ Then
          If CODYY$ <= cod2$ Then
            If Mid$(CODYY$, 16, 1) <> Chr$(96) Then
              CI% = CVI(Mid$(X$, 17, 2))
              If ACTI%(CI%) = 1 Or CI1% = CI2% Then
                Call EXPLOMUL(CI%)
                If CONTAMUL% > 1 Then
                  '
                  If REGQ& > 0 Then
                      REGQ& = REGQ& + 1
                      Call GRAREG("LIMULTI", String$(128, "="), REGQ&)
                  End If
                  '
                  For I% = 1 To CONTAMUL%
                    '
                    If I% = 1 Then
                       XX$ = REGBLAN$("LIMULTI")
                       REGQ& = REGQ& + 1
                       Call GRAREG("LIMULTI", XX$, REGQ&)
                    End If
                    '
                    If CANT(I%) > 0.00005 Then  ' SUPRIME DESARMADO
                       XX$ = REGBLAN$("LIMULTI")
                       Call REPLA(XX$, MKI$(I% - 1), 1, 2)
                       Call REPLA(XX$, CODEXT$(CIK%(I%)), 3, 16)
                       Call REPLA(XX$, DESCRIT$(CIK%(I%)), 17 + 2 * NIVE%(I%), 48)
                       Call REPLA(XX$, CSTRING$(MKS$(CANT(I%)), 4, 10, 5, 2), 75, 10)
                       Call REPLA(XX$, CSTRING$(MKI$(PAD%(I%)), 1, 6, 0, 2), 85, 6)
                       REGQ& = REGQ& + 1
                       Call GRAREG("LIMULTI", XX$, REGQ&)
                    End If
                    '
                    If I% = 1 Then
                       XX$ = REGBLAN$("LIMULTI")
                       REGQ& = REGQ& + 1
                       Call GRAREG("LIMULTI", XX$, REGQ&)
                    End If
                    '
                 Next I%
                  '
                End If
              End If
            End If
          End If
        End If
        '
    Next REG&
    '
    Call SETULTREG("LIMULTI", REGQ&)
    Call CIERRARCH("LIMULTI")
    '
    Call HEADERS("LIMULTI", "")
    Screen.MousePointer = 1
    Call FINAL("*LIMULTI")
    '
End Sub
'

Sub LIMULIN()
    '
    CODACT$ = Text1.TEXT
    CIACT% = CODINT%(Text1.TEXT)
    If CIACT% > 0 Then
         CI1% = CIACT%: CI2% = CIACT%
       Else
         XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
         XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    End If
    X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
    If Len(X$) < 4 Then
        Exit Sub
    End If
    '
    CI1% = CVI(Left$(X$, 2)): cod1$ = CODEXT$(CI1%): If TRIM$(cod1$) = "" Then Exit Sub
    CI2% = CVI(Mid$(X$, 3, 2)): cod2$ = CODEXT$(CI2%): If TRIM$(cod2$) = "" Then Exit Sub
    '
    Screen.MousePointer = 11
    Dim ACTI%(16384)
    '
    NM% = NUMAS%
    FIN& = ULTREG&("ESTRUNUE")
    For REG& = 1 To FIN&
        '
        Call TRACE(20, REG&, FIN&)
        X$ = REGLEIDO$("ESTRUNUE", REG&)
        CI% = CVI(Left$(X$, 2) + String$(2, 0))
        If CI% > 0 Then
            If CI% <= NUMAS% Then
               CODYY$ = CODEXT$(CI%)
               If Mid$(CODYY$, 16, 1) <> Chr$(96) Then
                  If ACTI%(CI%) = 0 Then
                     ACTI%(CI%) = 1
                  End If
               End If
            End If
        End If
        '
        DPS$ = Mid$(X$, 3, 120)
        If ULTREG&("ESTRUDOS") >= REG& Then
             DPS$ = DPS$ + REGLEIDO$("ESTRUDOS", REG&)
        End If
        '
        For K% = 1 To Len(DPS$) Step 8
            DPSX$ = Mid$(DPS$, K%, 8)
            CI% = CVI(Left$(DPSX$, 2))
            If CI% <= 0 Then GoTo 99
            If CI% <= NUMAS% Then
                ACTI%(CI%) = (-1)
            End If
        Next K%
        '
99  Next REG&
    '
    REGQ& = 0
    '
    J& = 0
    For REG& = 1 To NUINV%
        '
        X$ = REGLEIDO$("INVERT", REG&)
        CODYY$ = Left$(X$, 16)
        If CODYY$ >= cod1$ Then
          If CODYY$ <= cod2$ Then
            If Mid$(CODYY$, 16, 1) <> Chr$(96) Then
              CI% = CVI(Mid$(X$, 17, 2))
              If ACTI%(CI%) = 1 Or CI1% = CI2% Then
                Call EXPLOCOM(CI%, 1)
                If CONTAMUL% > 1 Then
                  '
                  For K% = 1 To CONTAMUL%
                    CI1% = CIK%(K%)
                    CAN = CANT(K%)
                    PD% = 0
                    If CI1% > 0 Then
                      If CI1% <= NUMAS% Then
                        If Abs(CAN) > 0.0005 Then
                          UM$ = UNIMED$(CI1%)
                          Z$ = REGBLAN$("LIESTRU")
                          Call REPLA(Z$, MKI$(CI1%), 3, 2)
                          Call REPLA(Z$, UM$, 5, 2)
                          Call REPLA(Z$, MKI$(CI%), 1, 2)
                          Call REPLA(Z$, MKS$(CAN), 7, 4)
                          Call REPLA(Z$, Chr$(PD%), 11, 1)
                          '
                          J& = J& + 1
                          Call GRAREG("LIESTRU", Z$, J&)
                          ACTIX% = 1
                          '
                        End If
                      End If
                    End If
                  Next K%
                End If
              End If
            End If
          End If
        End If
        '
    Next REG&
    '
    Call SETULTREG("LIESTRU", J&)
    Call CIERRARCH("LIESTRU")
    '
    Call HEADERS("LIESTRU", "")
    Screen.MousePointer = 1
    Call FINAL("*LIESTRU")
    '
End Sub
'
Sub MENSAEDIT()
   Call MENSERR(24, "Operación no Válida.\  \Complete Previamente la Edición de la Fórmula\Haciendo Click sobre 'Append' o 'Insert'.\  \Luego Vuelva a Intentarlo.")
   Command26.SetFocus
End Sub

Sub LIMPLOSI()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Then Exit Sub
   '
20 Screen.MousePointer = 11
   IMPLOSI.List1.Clear
   IMPLOSI.Caption = "Implosión " + DESCRIT$(CI%)
   '
   UREIMP& = ULTREG&("IMPLOSI")
   RFF$ = RECFILT$("IMPLOSI", 2, MKI$(CI%))
   For IK& = 1 To Len(RFF$) Step 4
     RIMPLO& = CVS(Mid$(RFF$, IK&, 4))
     If RIMPLO& > 0 Then
       If RIMPLO& <= UREIMP& Then
         IMPLOX$ = REGLEIDO$("IMPLOSI", RIMPLO&)
         CICONJU% = CVI(Left$(IMPLOX$, 2))
         If CICONJU% > 0 Then
           If CICONJU% <= NUMAS% Then
             Z$ = Space$(64)
             Call REPLA(Z$, CODEXT$(CICONJU%), 1, 16)
             Call REPLA(Z$, DESCRIT0$(CICONJU%), 17, 32)
             Call REPLA(Z$, UNIMED$(CICONJU%), 51, 2)
             Call REPLA(Z$, CSTRING$(Mid$(IMPLOX$, 5, 4), 3, 10, 4, 2), 55, 10)
             IMPLOSI.List1.AddItem Z$
             On Error Resume Next
             IMPLOSI.List1.TopIndex = IMPLOSI.List1.ListCount - 10
             On Error GoTo 0
             DoEvents
           End If
         End If
       End If
     End If
   Next IK&
   '
99 Screen.MousePointer = 1
   IMPLOSI.Show 1
   '
End Sub
'
Sub COMULTI()
   '
   CODD$ = Text1.TEXT
   CI% = CODINT%(CODD$)
   If CI% < 1 Then Exit Sub
   CICON% = CI%
   '
   Screen.MousePointer = 11
   Call EXPLOMUL(CI%)
   If CONTAMUL% < 2 Then Exit Sub
   '
   IMPLOSI.List1.Clear
   IMPLOSI.Caption = "Explosión Multinivel " + DESCRIT$(CI%)
   '
   FIN& = CONTAMUL
   For KKI% = 2 To CONTAMUL
     Z$ = Space$(64)
     SP$ = Space$(2 * (NIVE%(KKI%) - 1))
     Call REPLA(Z$, CODEXT$(CIK%(KKI%)), 1, 15)
     Call REPLA(Z$, SP$ + DESCRIT0$(CIK%(KKI%)), 17, 32)
     Call REPLA(Z$, UNIMED$(CIK%(KKI%)), 51, 2)
     Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%)), 3, 10, 4, 2), 55, 10)
     IMPLOSI.List1.AddItem Z$
   Next KKI%
   Screen.MousePointer = 1
   '
   IMPLOSI.Show 1
   '
End Sub

Private Sub Text32_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii% = 13 Then
      Text32(Index) = FORMATNUM$(Val(Text32(Index)), "F11.6")
   End If
End Sub

Private Sub Text32_LostFocus(Index As Integer)
   Text32(Index) = FORMATNUM$(Val(Text32(Index)), "F11.6")
End Sub

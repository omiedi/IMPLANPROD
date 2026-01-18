VERSION 5.00
Begin VB.Form LISMAT04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lista de Materiales - A-B-M-L Materiales por Grupo de O/T"
   ClientHeight    =   8010
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   Icon            =   "LISMAT04.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text12 
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
      Left            =   10320
      MaxLength       =   5
      TabIndex        =   67
      Top             =   870
      Width           =   645
   End
   Begin VB.TextBox Text8 
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
      Left            =   9120
      MaxLength       =   5
      TabIndex        =   65
      Top             =   870
      Width           =   645
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00C0E0F0&
      Caption         =   "Ctl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   11025
      TabIndex        =   63
      Top             =   5880
      Width           =   855
      Begin VB.CommandButton Command7 
         Caption         =   "App"
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
         Left            =   105
         Picture         =   "LISMAT04.frx":0442
         TabIndex        =   51
         ToolTipText     =   "Agrega Registro al Final de la Lista"
         Top             =   945
         Width           =   650
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Ins"
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
         Left            =   105
         Picture         =   "LISMAT04.frx":074C
         TabIndex        =   52
         ToolTipText     =   "Intercala Registro Encima del Activo"
         Top             =   1220
         Width           =   650
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Del"
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
         Left            =   105
         Picture         =   "LISMAT04.frx":0A56
         TabIndex        =   53
         ToolTipText     =   "Elimina el Registro Activo de la Lista"
         Top             =   1600
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Left            =   105
         Picture         =   "LISMAT04.frx":0D60
         TabIndex        =   50
         ToolTipText     =   "Edición del Registro Activo para su Modificación"
         Top             =   590
         Width           =   650
      End
      Begin VB.CommandButton Command6 
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
         Height          =   275
         Left            =   105
         TabIndex        =   49
         ToolTipText     =   "Crea un Nuevo Registro"
         Top             =   315
         Width           =   650
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0E0F0&
      Caption         =   "Actualizacion de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   105
      TabIndex        =   39
      Top             =   5775
      Width           =   10830
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0E0F0&
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   7875
         TabIndex        =   72
         Top             =   230
         Width           =   2850
         Begin VB.Label Label18 
            BackColor       =   &H00C0E0F0&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   225
            Left            =   1470
            TabIndex        =   74
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label17 
            BackColor       =   &H00C0E0F0&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   225
            Left            =   105
            TabIndex        =   73
            Top             =   315
            Width           =   1275
         End
      End
      Begin VB.CommandButton Command10 
         Height          =   265
         Left            =   8085
         Picture         =   "LISMAT04.frx":106A
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Acceso a Planera Electronica"
         Top             =   1365
         Width           =   420
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Pl-Pos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   190
         Left            =   9765
         TabIndex        =   70
         ToolTipText     =   "Alta de Item por Plano-Posición"
         Top             =   1120
         Width           =   750
      End
      Begin VB.CommandButton Command17 
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
         Height          =   190
         Left            =   9220
         TabIndex        =   69
         ToolTipText     =   "Alta de Nuevo Item de Stock"
         Top             =   1120
         Width           =   540
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1155
         TabIndex        =   41
         Top             =   315
         Width           =   3885
      End
      Begin VB.CommandButton Command1 
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
         Left            =   10500
         Picture         =   "LISMAT04.frx":1374
         TabIndex        =   48
         Top             =   1365
         Width           =   175
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8610
         TabIndex        =   47
         Top             =   1365
         Width           =   1890
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5775
         TabIndex        =   46
         Top             =   1365
         Width           =   1890
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7140
         TabIndex        =   45
         Top             =   840
         Width           =   525
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5775
         TabIndex        =   44
         Top             =   840
         Width           =   840
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5775
         TabIndex        =   43
         Top             =   315
         Width           =   1890
      End
      Begin VB.TextBox Text19 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   40
         Top             =   1680
         Width           =   840
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   1155
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         Top             =   690
         Width           =   4150
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   0
         TabIndex        =   64
         Top             =   375
         Width           =   1110
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   5775
         TabIndex        =   62
         Top             =   1680
         Width           =   4740
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   8190
         TabIndex        =   61
         Top             =   1140
         Width           =   1110
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Material:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   5460
         TabIndex        =   60
         Top             =   1140
         Width           =   1110
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Llave:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   6825
         TabIndex        =   59
         Top             =   630
         Width           =   1110
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Posición:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   5460
         TabIndex        =   58
         Top             =   630
         Width           =   1110
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Plano:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   5200
         TabIndex        =   57
         Top             =   375
         Width           =   1110
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "N-Ord:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   210
         TabIndex        =   56
         Top             =   780
         Width           =   1110
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Cant:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   55
         Top             =   1440
         Width           =   1110
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   525
         TabIndex        =   54
         Top             =   1050
         Width           =   540
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   4280
      Left            =   10605
      TabIndex        =   38
      Top             =   1365
      Width           =   15
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame7"
      Height          =   4280
      Left            =   10290
      TabIndex        =   37
      Top             =   1365
      Width           =   15
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame6"
      Height          =   4280
      Left            =   8610
      TabIndex        =   36
      Top             =   1365
      Width           =   15
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   4280
      Left            =   7245
      TabIndex        =   35
      Top             =   1365
      Width           =   15
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   4280
      Left            =   6615
      TabIndex        =   34
      Top             =   1365
      Width           =   15
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   4280
      Left            =   5040
      TabIndex        =   33
      Top             =   1365
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4280
      Left            =   1155
      TabIndex        =   32
      Top             =   1365
      Width           =   15
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4280
      Left            =   525
      TabIndex        =   31
      Top             =   1365
      Width           =   15
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00E0E0E0&
      Height          =   4185
      Left            =   105
      ScaleHeight     =   4125
      ScaleWidth      =   10770
      TabIndex        =   27
      Top             =   1500
      Width           =   10830
      Begin VB.Frame Frame16 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   330
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   10515
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   $"LISMAT04.frx":167E
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
            Left            =   0
            TabIndex        =   30
            Top             =   50
            Width           =   10530
         End
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3840
         ItemData        =   "LISMAT04.frx":1727
         Left            =   -30
         List            =   "LISMAT04.frx":1729
         TabIndex        =   28
         Top             =   330
         Width           =   10830
      End
   End
   Begin VB.TextBox Text9 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9820
      TabIndex        =   13
      Top             =   210
      Width           =   1125
   End
   Begin VB.TextBox Text11 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9975
      TabIndex        =   12
      Top             =   8595
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.TextBox Text10 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9820
      TabIndex        =   11
      Top             =   525
      Width           =   1125
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   5685
      Left            =   11025
      ScaleHeight     =   5625
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command15 
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
         Left            =   105
         Picture         =   "LISMAT04.frx":172B
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Guardar "
         Top             =   1530
         Width           =   640
      End
      Begin VB.CommandButton Command25 
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
         Picture         =   "LISMAT04.frx":1A35
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprime Ficha Lista de Materiales"
         Top             =   2200
         Width           =   650
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Copy"
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
         Picture         =   "LISMAT04.frx":209F
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Copia de Lista de Materiales"
         Top             =   2880
         Width           =   650
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Shop"
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
         Picture         =   "LISMAT04.frx":21E9
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Accesos a Gestión de Compras por O.Trabajo"
         Top             =   3540
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   9
         Top             =   5010
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   10
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command14 
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
         Height          =   690
         Left            =   105
         Picture         =   "LISMAT04.frx":24F3
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   4210
         Width           =   650
      End
      Begin VB.CommandButton Command99 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "LISMAT04.frx":2935
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   8100
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "LISMAT04.frx":2C3F
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exit"
         Top             =   105
         Width           =   650
      End
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
         Height          =   690
         Left            =   105
         Picture         =   "LISMAT04.frx":2D89
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Ayuda en Línea"
         Top             =   735
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -210
         Picture         =   "LISMAT04.frx":3093
         Stretch         =   -1  'True
         Top             =   5250
         Width           =   1410
      End
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "PL-7:"
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
      Left            =   9480
      TabIndex        =   68
      Top             =   960
      Width           =   810
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "PL-5:"
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
      Left            =   8280
      TabIndex        =   66
      Top             =   960
      Width           =   810
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11025
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label Label8 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1155
      TabIndex        =   26
      Top             =   840
      Width           =   7350
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1155
      TabIndex        =   25
      Top             =   525
      Width           =   7350
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5565
      TabIndex        =   24
      Top             =   210
      Width           =   2940
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3675
      TabIndex        =   23
      Top             =   210
      Width           =   945
   End
   Begin VB.Label Label31 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1155
      TabIndex        =   22
      Top             =   210
      Width           =   1680
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro de E.V.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   21
      Top             =   315
      Width           =   1140
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Denom.Obra:"
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
      Left            =   -210
      TabIndex        =   20
      Top             =   630
      Width           =   1350
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Den.Grupo:"
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
      Left            =   0
      TabIndex        =   19
      Top             =   945
      Width           =   1110
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo:"
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
      Left            =   2730
      TabIndex        =   18
      Top             =   315
      Width           =   900
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
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
      Left            =   9045
      TabIndex        =   17
      Top             =   315
      Width           =   690
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Entrega:"
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
      Left            =   8925
      TabIndex        =   16
      Top             =   630
      Width           =   810
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Resp.:"
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
      Left            =   9450
      TabIndex        =   15
      Top             =   8700
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
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
      Left            =   4410
      TabIndex        =   14
      Top             =   315
      Width           =   1110
   End
   Begin VB.Line Line1 
      X1              =   -420
      X2              =   11580
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu Archiv 
      Caption         =   "Archivos"
      Begin VB.Menu NuMat 
         Caption         =   "Nuevo Material"
      End
      Begin VB.Menu plh1 
         Caption         =   "-"
      End
      Begin VB.Menu GuaCli 
         Caption         =   "Guardar"
      End
      Begin VB.Menu php1a 
         Caption         =   "-"
      End
      Begin VB.Menu Print 
         Caption         =   "Imprimir Ficha"
      End
      Begin VB.Menu plh3 
         Caption         =   "-"
      End
      Begin VB.Menu List 
         Caption         =   "Consultas y Listados"
      End
      Begin VB.Menu plh4 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu EdiCli 
      Caption         =   "Edición"
      Begin VB.Menu mnuEdit 
         Caption         =   "Editar"
      End
      Begin VB.Menu mnuagrefi 
         Caption         =   "Agregar al Final"
      End
      Begin VB.Menu mnuIns 
         Caption         =   "Insertar"
      End
      Begin VB.Menu php4c 
         Caption         =   "-"
      End
      Begin VB.Menu CoInLiMa 
         Caption         =   "Copia Interactiva de Lista de Materiales"
      End
      Begin VB.Menu php4b 
         Caption         =   "-"
      End
      Begin VB.Menu mnudel 
         Caption         =   "Borrar"
      End
      Begin VB.Menu php5a 
         Caption         =   "-"
      End
      Begin VB.Menu BusCli 
         Caption         =   "Buscar"
      End
   End
   Begin VB.Menu ConLiCli 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu LiGeClaCli 
         Caption         =   "Listado General y Clasificado"
      End
   End
   Begin VB.Menu ConfCli 
      Caption         =   "Configuración"
      Begin VB.Menu ConGenCli 
         Caption         =   "Configuración General"
      End
   End
   Begin VB.Menu AyuCli 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "LISMAT04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHEX$(16)
Dim YACAR%, MODIFIC%
Dim GRABATODO%
Dim AGRECLI%
Dim RECONTANT$
Dim MODOEDIT%

Private Sub AyuCli_Click()
  Call Command4_Click
End Sub

Private Sub Baj_Cli_Click()
End Sub

Private Sub BusCli_Click()
'Call Command20_Click
End Sub

Private Sub CoInBDCli_Click()
Call Command99_Click
End Sub

Private Sub CoInLiMa_Click()
    Call Command22_Click
End Sub

Private Sub Command1_Click()
    '
    If Text6.BackColor = 15132390 Then
        Exit Sub
    End If
    '
    Call SELECHO("MASTER")
    CODD$ = VALACT1$("MASTER")
    If CODD$ = "" Then
        Exit Sub
    End If
    COD$ = CODD$
    Text6.TEXT = COD$
    '
End Sub

Private Sub Command10_Click()
    '
    Command10.Enabled = False
    If TRIM$(Text6) <> "" Then
        Call CONECRUN("PLANEL04/" + TRIM$(Text6), "Planera Electrónica")
    End If
    Command10.Enabled = True
    '
End Sub

Private Sub Command11_Click()
'    'Aca viene el lista de Materiales pendientes de compra
'    Call ABREORTRA
'    SQLX$ = "SELECT * FROM LISTAMAT "
'    SQLX$ = SQLX$ + " WHERE (Nro_Ocom = '' or isnull(Nro_Ocom) = true OR IsEmpty(Nro_Ocom) = true)" 'CONDICION PENDIENTE DE COMPRA
'    SQLX$ = SQLX$ + " AND NRO_OTRAB = '" & TRIM$(Label31) & "' "
'    SQLX$ = SQLX$ + " AND NRO_GRUP = '" & TRIM$(Label2) & "' "
'    SQLX$ = SQLX$ + " ORDER BY Nro_Otrab, Nro_Grup "
'    Set LMATPEN = OTrabajo.OpenRecordset(SQLX$, 4)
'    With LMATPEN
'        '
'     On Error Resume Next
'     .MoveLast
'     'Valida que la tabla no este vacia
'     If Err Then
'       Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
'       On Error GoTo 0
'       GoTo 90
'     End If
'     On Error GoTo 0
'     FIN& = .RecordCount
'     .MoveFirst
'     JJ& = 0
'     Do While (.EOF = False)
'      XXX$ = REGBLAN$("LIMATPEN")
'         '
'         Call REPLA(XXY$, !Nro_Otrab, 1, 16)
'         Call REPLA(XXY$, !Nro_Grup, 18, 10)
'         ' ACA HAY QUE VALIDAR SI ESTA CODIFICADO !CODEXTE
'         'SINO SE USA EL CAMPO DESCRI_COR (IGUAL VALIDAR CON CODINTE)
'         Call REPLA(XXY$, !CODEXTE, 29, 48)
'         Call REPLA(XXY$, MKS(!CANTID), 78, 4)
'         '
'         JJ& = JJ& + 1
'         Call GRAREG("LIMATPEN", XXY$, JJ&)
'       .MoveNext
'     Loop
'     End With
'     '
'     Call SETULTREG("LIMATPEN", JJ&)
'     Call CIERRARCH("LIMATPEN")
'90  Screen.MousePointer = 1
'    Call FINAL("*LIMATPEN")
'    '
End Sub

Private Sub Command12_Click()
   '
   If TRIM$(Text7) = "" Then Exit Sub
   '
   U& = List3.ListIndex
   If U& < 0 Then U& = 0
   '
   TTXX$ = Space$(128)
   NORDE% = List3.ListCount + 1
   Call REPLA(TTXX$, FORMATNUM$(NORDE%, "I2"), 1, 2)
   Call REPLA(TTXX$, FORMATNUM$(Val(Text19), "F6.1"), 4, 6)
   Call REPLA(TTXX$, AJUSTI$(Text7, 36), 11, 36)
   Call REPLA(TTXX$, AJUSTI$(Text2, 14), 48, 14)
   Call REPLA(TTXX$, AJUSTI$(Text3, 5), 63, 5)
   Call REPLA(TTXX$, AJUSTI$(Text5, 12), 69, 12)
   Call REPLA(TTXX$, AJUSTI$(Text6, 16), 82, 16)
   Call REPLA(TTXX$, AJUSTI$(Text4, 2), 98, 2)
   TTXX$ = TTXX$ + Text1
   List3.AddItem TTXX$, U&
   '
   Call RENUMERA
   List3.ListIndex = U&
   '
   MODOEDIT% = 0
   Call CAMGRIS
   Command6.Enabled = True
   Command2.Enabled = True
   Command5.Enabled = True
   '
End Sub

Private Sub Command13_Click()
   '
   Command13.Enabled = False
   Call CONECRUN("COMVOI04", "Gestión de Compras por Lista de Materiales")
99 Command13.Enabled = True
   '
End Sub

Private Sub Command17_Click()
   Command17.Enabled = False
   Call CONECRUN("AMASTO04/NEW", "Maestro de Items de Stock")
   Command17.Enabled = True
End Sub

Private Sub Command18_Click()
    'validaciones
    If MODOEDIT% <> 1 Then Exit Sub
    '
    If TRIM$(Text2) = "" Or TRIM(Text3) = "" Then
        Call COMUNI("El Material Debe Tener un Plano y una Posición. Imposible dar el Alta.")
        Exit Sub
    End If
    If TRIM$(Text7) = "" Then
        Call COMUNI("El Material debe Tener una Referencia. Imposible dar el Alta.")
        Exit Sub
    End If
    CODI2$ = TRIM$(Text2) + "-" + TRIM$(Text3)
    If Len(CODI2$) > 15 Then
        Call COMUNI("Codigo de Material No Válido. El Codigo Excede la Cantidad \ de Caracteres Permitidos. Imposible dar el Alta.")
        Exit Sub
    End If
    '
' MUY IMPORTANTE VER LO QUE SIGUE CON WALTER
    Call ABREMASTER
    SQLX$ = "SELECT * FROM MASTER "
    SQLX$ = SQLX$ + " WHERE TRIM(CODIGO) = '" & CODI2$ & "' "
    Set COCODTMP = Articulos.OpenRecordset(SQLX$, 4)
    With COCODTMP
     On Error Resume Next
     .MoveFirst
     If Err Then 'el codigo no es pre-existente LO aGREGO
       On Error GoTo 0
       With Master
          On Error GoTo 0
          GoTo 15
       End With
     End If
     On Error GoTo 0
     Call MENSERR(24, "Código de Articulo no válido o Pre-Existente. No se ha Podido dar de Alta el Articulo.")
     Exit Sub
    End With
    '
15  ALPORPLA04.Text1 = CODI2$
    ALPORPLA04.Text2 = TRIM$(Text7)
    ALPORPLA04.Show 1
    '
End Sub

Private Sub Command22_Click()
  Command22.Enabled = False
    LITMAOP.Show 1
  Command22.Enabled = True
End Sub

Private Sub Command4_Click()
   Call HELPONLINE("AMACLI")
End Sub

Private Sub Command7_Click()
   '
   If TRIM$(Text7) = "" Then Exit Sub
   '
   TTXX$ = Space$(128)
   NORDE% = List3.ListCount + 1
   Call REPLA(TTXX$, FORMATNUM$(NORDE%, "I2"), 1, 2)
   Call REPLA(TTXX$, FORMATNUM$(Val(Text19), "F6.1"), 4, 6)
   Call REPLA(TTXX$, AJUSTI$(Text7, 36), 11, 36)
   Call REPLA(TTXX$, AJUSTI$(Text2, 14), 48, 14)
   Call REPLA(TTXX$, AJUSTI$(Text3, 5), 63, 5)
   Call REPLA(TTXX$, AJUSTI$(Text5, 12), 69, 12)
   Call REPLA(TTXX$, AJUSTI$(Text6, 16), 82, 16)
   Call REPLA(TTXX$, AJUSTI$(Text4, 2), 98, 2)
   TTXX$ = TTXX$ + Text1
   List3.AddItem TTXX$
   '
   Call RENUMERA
   List3.ListIndex = List3.ListCount - 1
   '
   MODOEDIT% = 0
   Call CAMGRIS
   Command6.Enabled = True
   Command2.Enabled = True
   Command5.Enabled = True
   '
End Sub

Sub RENUMERA()
   For U& = 1 To List3.ListCount
     TTXX$ = List3.List(U& - 1)
     List3.RemoveItem U& - 1
     NORDE% = U&
     Call REPLA(TTXX$, FORMATNUM$(NORDE%, "I2"), 1, 2)
     List3.AddItem TTXX$, U& - 1
   Next U&
End Sub



Private Sub Command9_Click()
'    Command9.Enabled = False
'    '21/02/2005
'    RESP2% = COMALTER%("Formato de Alta de Articulo\\Por Código de Familia\Material Codificado\Combinación Plano-Posición")
'    If RESP2% = 1 Then ALPORFAM04.Show 1
'    If RESP2% = 2 Then Call CONECRUN("ARCHIG04/AMASTO", "Archivos Maestros")
'    If RESP2% = 3 Then ALPORPLA04.Show 1
'    '
'    Command9.Enabled = True
End Sub

Private Sub Command99_Click()
   Command99.Enabled = False
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("DEUDOR1")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     R0$ = REGLEIDO$("DEUDOR1", U&)
     NCLIE% = CVI(Left$(R0$, 2))
     R1$ = Mid$(R0$, 3, 30)
     If Left$(R1$, 1) = Chr$(64) Then
       R1$ = AJUSTI$(TRIM$(Mid$(R1$, 2)), 30)
     End If
     R1$ = REPLACAR$(R1$, "@", " ")
     R2$ = AJUSTI$(RASOCLI$(NCLIE%), 30)
     If R1$ <> R2$ Then
       HAYINCO% = HAYINCO% + 1
       Call MENSERR(24, "Revisar Cuenta " + TRIM$(Str$(NCLIE%)))
     End If
   Next U&
   '
   If HAYINCO% < 1 Then
     Call COMUNI("Control Completo.\No se Detectaron Errores.")
   End If
   '
   Screen.MousePointer = 1
   Command99.Enabled = True
End Sub

Private Sub ConGenCli_Click()
Call Command14_Click
End Sub

Private Sub ConPagCli_Click()
With OPMASCL04
  .Option1(6).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub CuCoCli_Click()
With OPMASCL04
  .Option1(8).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub DeHaCaCli_Click()
Call Command2_Click
End Sub

Private Sub DeLiCreCli_Click()
    Call AMACLIS04.Command3_Click
End Sub

Private Sub EditCli_Click()
   'Call Command9_Click
End Sub

Private Sub Exit_Click()
   Call Command3_Click
End Sub

'
Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
End Sub
'
Private Sub Form_Activate()
   WindowState = 0
   OPENFORMS = DoEvents
End Sub
'
Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   OPORTRA04.Show 1
   Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
    '
    If DERACCE%("GRALISMA", "Grabar Lista de Materiales") >= 2 Then
      Call GRALISMAT
      '
      TTXX$ = Space$(129)
      Label13 = Left$(TTXX$, 2)
      Text19 = Mid$(TTXX$, 4, 6)
      Text7 = Mid$(TTXX$, 11, 36)
      Text2 = Mid$(TTXX$, 48, 14)
      Text3 = Mid$(TTXX$, 63, 5)
      Text5 = Mid$(TTXX$, 69, 12)
      Text6 = Mid$(TTXX$, 82, 16)
      Text4 = Mid$(TTXX$, 98, 2)
      Label17 = Mid$(TTXX$, 105, 12)
      Label18 = Mid$(TTXX$, 117, 12)
      Text1 = Mid$(TTXX$, 129)
      '
      List3.ListIndex = (-1)
    End If
End Sub

Private Sub Command2_Click()
   'Para que no se pueda editar las que tiene O/C o B/R o ambas
   If TRIM$(Label17) = "" And TRIM$(Label18) = "" Then
    If List3.ListIndex >= 0 Then
      LINDEX = List3.ListIndex
      MODOEDIT% = 1
      Call CAMBLAN
      Command6.Enabled = False
      Command2.Enabled = False
      Command5.Enabled = False
      '
      U& = List3.ListIndex
      List3.RemoveItem U&
      If LINDEX < List3.ListCount Then
          List3.ListIndex = LINDEX
        Else
          List3.ListIndex = -1
      End If
    End If
   End If
End Sub

Sub CAMGRIS()
   '
   COLOBAK = RGB(230, 230, 230)
   Text1.BackColor = COLOBAK
   Text2.BackColor = COLOBAK
   Text3.BackColor = COLOBAK
   Text4.BackColor = COLOBAK
   Text5.BackColor = COLOBAK
   Text6.BackColor = COLOBAK
   Text7.BackColor = COLOBAK
   Text19.BackColor = COLOBAK
   '
End Sub

Sub CAMBLAN()
   '
   COLOBAK = RGB(255, 255, 255)
   Text1.BackColor = COLOBAK
   Text2.BackColor = COLOBAK
   Text3.BackColor = COLOBAK
   Text4.BackColor = COLOBAK
   Text5.BackColor = COLOBAK
   Text6.BackColor = COLOBAK
   Text7.BackColor = COLOBAK
   Text19.BackColor = COLOBAK
   '
End Sub

Private Sub Command25_Click()
   Call PRILISMAT
  ' PRINTFORM
End Sub

Private Sub Command3_Click()
   '
   If MODOEDIT% = 1 Then Exit Sub
   List3.Clear
   '
   TTXX$ = Space$(129)
   Label13 = Left$(TTXX$, 2)
   Text19 = Mid$(TTXX$, 4, 6)
   Text7 = Mid$(TTXX$, 11, 36)
   Text2 = Mid$(TTXX$, 48, 14)
   Text3 = Mid$(TTXX$, 63, 5)
   Text5 = Mid$(TTXX$, 69, 12)
   Text6 = Mid$(TTXX$, 82, 16)
   Text4 = Mid$(TTXX$, 98, 2)
   Text1 = Mid$(TTXX$, 129)
   Label17 = Mid$(TTXX$, 105, 12)
   Label18 = Mid$(TTXX$, 117, 12)
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
   '
End Sub

Private Sub Command5_Click()
   'Para que no se pueda borrar las que tiene O/C o B/R o ambas
   If TRIM$(Label17) = "" And TRIM$(Label18) = "" Then
    If List3.ListIndex >= 0 Then
      U& = List3.ListIndex
      List3.RemoveItem U&
      If U& < List3.ListCount Then
          List3.ListIndex = U&
        Else
          List3.ListIndex = List3.ListCount - 1
      End If
    End If
   End If
End Sub

Private Sub Command6_Click()
   '
   TTXX$ = Space$(128)
   Label13 = Left$(TTXX$, 2)
   Text19 = Mid$(TTXX$, 4, 6)
   Text7 = Mid$(TTXX$, 11, 36)
   Text2 = Mid$(TTXX$, 48, 14)
   Text3 = Mid$(TTXX$, 63, 5)
   Text5 = Mid$(TTXX$, 69, 12)
   Text6 = Mid$(TTXX$, 82, 16)
   Text4 = Mid$(TTXX$, 98, 2)
   Text1 = Mid$(TTXX$, 129)
   '
   MODOEDIT% = 1
   Call CAMBLAN
   Command6.Enabled = False
   Command2.Enabled = False
   Command5.Enabled = False
   '
End Sub


Private Sub LiCuInCli_Click()
    Call AMACLIS04.Command5_Click
End Sub

Private Sub Frame3_DBLCLICK()
'    Call List1_DblClick
End Sub

Private Sub Label17_dblClick()
    'MUESTRA AL O/C
     NPX& = Val(Mid$(Label17, 4, 7))
     If NPX& > 0 Then
        Call SHOWOCOM(NPX&)
     End If
    '
End Sub
Sub SHOWOCOM(NPX&)
    '
    TIDOCUM$ = "Nota de Pedido"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU$ = " " + NUDOCU$ + " "
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    Exit Sub
    '
5   For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    Exit Sub
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    FIN& = Len(TPSP$)
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7070
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7070
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    UI& = UI& + 1
    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
    If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.HScroll1.Min = 0
    EDITFORM.HScroll1.Value = 0
    EDITFORM.HScroll1.Max = 1
    If ANCHPAGINA > EDITFORM.Frame1.Width Then
      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command1.Enabled = False
    EDITFORM.Show 1
    '
End Sub

Private Sub Label18_dblClick()
    '
    NPX& = Val(Mid$(Label18, 4, 7))
    If NPX& > 0 Then
     Call SHOWBORE(NPX&)
    End If
    '
End Sub
Sub SHOWBORE(NPX&)
    '
    TIDOCUM$ = "Boleta de Recepción"
    NUDOCU$ = TRIM$(Str$(NPX&))
    NUDOCU$ = " " + NUDOCU$ + " -"
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        If InStr(XX$, NUDOCU$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    Exit Sub
    '
5  For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    Exit Sub
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    FIN& = Len(TPSP$)
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7070
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7070
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    UI& = UI& + 1
    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
    If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.HScroll1.Min = 0
    EDITFORM.HScroll1.Value = 0
    EDITFORM.HScroll1.Max = 1
    If ANCHPAGINA > EDITFORM.Frame1.Width Then
      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command1.Enabled = False
    EDITFORM.Show 1
    '
End Sub



Private Sub Label2_dblClick()
    '
    'TAMONED VA A ESTAR DEFINIDA COMO:
    '1=PESOS
    '2=DÓLARES
    '3=EUROS
    '4=REALES
     Select Case UCase(Label2)
       Case "PESOS"
        Label2.ForeColor = RGB(255, 0, 0)
        Label2 = "Dólares"
        MONEMI% = 2
        Label2.ToolTipText = "Doble-Click para 'Euros'"
        
      Case "DÓLARES"
        Label2.ForeColor = vbYellow
        Label2 = "Euros"
        MONEMI% = 3
        Label2.ToolTipText = "Doble-Click para 'Reales'"
        
      Case "EUROS"
        Label2.ForeColor = vbMagenta
        Label2 = "Reales"
        Label2.ToolTipText = "Doble-Click para 'Pesos'"
        
        MONEMI% = 4
      Case "REALES"
        Label2.ForeColor = RGB(0, 0, 255)
        Label2.Caption = "Pesos"
        Label2.ToolTipText = "Doble-Click para 'Dólares'"
        MONEMI% = 1
     End Select
     '
End Sub





Private Sub LiGeClaCli_Click()
 Call FORTRA04.Command13_Click
End Sub

Private Sub SuspOper_Click()
    Call Command3_Click
End Sub

Sub GENEPLACO() 'GENERAR CAPLACO.DAT
    '
    NUOTRAB$ = TRIM$(Text1.TEXT)
    Call ABREORTRA
    SQLX$ = "SELECT * FROM PlanCob "
    SQLX$ = SQLX$ + " WHERE Nro_Otrab= '" & NUOTRAB$ & "' "
    Set PLACOTMP = OTrabajo.OpenRecordset(SQLX$, 4)
    
    J& = 0
    With PLACOTMP
      '
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do Until .EOF = False
          Y$ = REGBLAN("CAPLACO")
          Call REPLA(Y$, !Nro_Otrab, 1, 16)
          Call REPLA(Y$, !NuOrItem, 17, 2)
          Call REPLA(Y$, !Refe_Pcob, 19, 64)
          Call REPLA(Y$, !Impo_Pcob, 83, 8)
          Call REPLA(Y$, !FeVen_Pcob, 91, 2)
          Call REPLA(Y$, !NroFac_Pcob, 93, 16)
          Call REPLA(Y$, !MONEDA, 109, 2)
          Call REPLA(Y$, !Coti_Pcob, 111, 8)
          J& = J& + 1
          Call GRAREG("!CAPLACO", Y$, J&)
        .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    '
    Call SETULTREG("!CAPLACO", J&)
    Call CIERRARCH("!CAPLACO")
    
End Sub

Private Sub List3_Click()
   '
   If MODOEDIT% = 1 Then
     Exit Sub
   End If
   '
   Call CAMGRIS
   '
   TTXX$ = List3.TEXT
   Label13 = Left$(TTXX$, 2)
   Text19 = Mid$(TTXX$, 4, 6)
   Text7 = Mid$(TTXX$, 11, 36)
   Text2 = Mid$(TTXX$, 48, 14)
   Text3 = Mid$(TTXX$, 63, 5)
   Text5 = Mid$(TTXX$, 69, 12)
   Text6 = Mid$(TTXX$, 82, 16)
   Text4 = Mid$(TTXX$, 98, 2)
   'CARGO O/C Y BR
   NOC1$ = Mid$(TTXX$, 105, 12)
   NBR1$ = Mid$(TTXX$, 117, 12)
   If TRIM$(NOC1$) <> "" Then
     Label17 = NOC1$
   Else
     Label17 = ""
   End If
   '
   If TRIM$(NBR1$) <> "" Then
      Label18 = NBR1$
    Else
      Label18 = ""
   End If
   '
   Text1 = Mid$(TTXX$, 129)
   '
End Sub



Private Sub mnuagrefi_Click()
    Call Command7_Click
End Sub

Private Sub mnudel_Click()
    Call Command5_Click
End Sub

Private Sub mnuIns_Click()
    Call Command12_Click
End Sub

Private Sub NuMat_Click()
    Call Command6_Click
End Sub

Private Sub Text1_GotFocus()
   If MODOEDIT% <> 1 Then List3.SetFocus
End Sub

Private Sub Text10_GotFocus()
    Text8.SetFocus
End Sub

Private Sub Text2_GotFocus()
   If MODOEDIT% <> 1 Then List3.SetFocus
End Sub

Private Sub Text3_GotFocus()
   If MODOEDIT% <> 1 Then List3.SetFocus
End Sub

Private Sub Text4_GotFocus()
   If MODOEDIT% <> 1 Then List3.SetFocus
End Sub

Private Sub Text5_GotFocus()
   If MODOEDIT% <> 1 Then List3.SetFocus
End Sub

Private Sub Text6_Change()
   Label14 = ""
   If CODINT%(Text6) > 0 Then
     CI% = CODINT%(COD$)
     Label14 = DESCRIT$(CI%)
     If TRIM$(Text7) = "" Then
       Text7 = Label14
     End If
   End If
End Sub

Private Sub Text6_GotFocus()
   If MODOEDIT% <> 1 Then List3.SetFocus
End Sub

Private Sub Text7_GotFocus()
   If MODOEDIT% <> 1 Then List3.SetFocus
End Sub

Private Sub Text19_GotFocus()
   If MODOEDIT% <> 1 Then List3.SetFocus
End Sub

Private Sub Text11_GotFocus()
    Text25.SetFocus
End Sub

Private Sub Text2_LostFocus()
   Text2 = AJUSTI$(Text2, 14)
End Sub

Private Sub Text3_LostFocus()
   Text3 = AJUSTI$(Text3, 5)
End Sub

Private Sub Text4_LOSTFOCUS()
    Text4 = AJUSTI$(Text4, 2)
End Sub

Sub BLANFORMU()
   '
   Text2 = ""
   Text3 = ""
   Text24 = ""
   'Text25.TEXT = ""
   'Text36.TEXT = ""
   'Text38.TEXT = ""
   
   'Text26.TEXT = ""
'   Text27.TEXT = ""
'   Text20.TEXT = ""
  Text4.TEXT = ""
  'Text14.TEXT = ""
  Text19.TEXT = ""
   'Text28.TEXT = ""
   'Text29.TEXT = ""
   'Text30.TEXT = ""
'  'Text32.TEXT = ""
   'Text19.TEXT = ""
'   Text33.TEXT = ""
'   Text34.TEXT = ""
   'Text36.TEXT = ""
   'Text21.TEXT = ""
   '
'   For KK% = 0 To 7
'     COTEXT(KK%).TEXT = ""
'   Next KK%
   '
'   Text28.TEXT = ""
   '
     Text23 = ""
     Text35 = ""
     Text36 = ""
     Text37 = ""
     'Text1 = ""
     Text3 = ""
     '
     Text18 = ""
     Text5 = ""
     Text7 = ""
     Text8 = ""
     '
     Text9 = ""
     Text10 = ""
     Text25 = ""
     Text24 = ""
     Text11 = ""
     Text13 = ""
     Text12 = ""
     Text17 = ""
'   List2.Clear
   List3.Clear
   'List1.Clear
   '
   YACAR% = 0
   '
End Sub

'
Private Sub Form_Unload(Cancel As Integer)
   '
   Unload Me
   '
End Sub

Private Sub GruCli_Click()
With OPMASCL04
  .Option1(9).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub GuaCli_Click()
    Call COMMAND15_CLICK
End Sub

Private Sub LiPreCli_Click()
With OPMASCL04
  .Option1(7).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub List_Click()
   'Call Command13_Click
End Sub

Private Sub List3_DblClick()
   'Para que no se pueda editar las que tiene O/C o B/R o ambas
   If TRIM$(Label17) = "" And TRIM$(Label18) = "" Then
    If MODOEDIT% = 0 Then
         Call Command2_Click    ' EDITAR
    End If
   End If
End Sub

Private Sub NueCli_Click()
'   Call Command21_Click
End Sub

Private Sub OrDBCli_Click()
Call COMMAND15_CLICK
End Sub

Private Sub Print_Click()
   Call Command25_Click
End Sub

Private Sub PrZoVeCli_Click()
With OPMASCL04
  .Option1(5).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub Susp_Cli_Click()
   'Call Command23_Click
End Sub

Private Sub Text1_Change()
   'For Index = 0 To 8: YACAR%(Index) = 0: Next Index
End Sub
'
Sub text19_lostfocus()
   Text19 = FORMATNUM(Val(Text19), "F6.1")
End Sub

Private Sub Text5_LostFocus()
   Text5 = AJUSTI$(Text5, 12)
End Sub

Private Sub Text6_LostFocus()
   Text6 = AJUSTI$(Text6, 16)
End Sub

Private Sub Text7_LostFocus()
   Text7 = AJUSTI$(Text7, 36)
End Sub

Private Sub VenCli_Click()
With OPMASCL04
  .Option1(4).Value = True
  Call .Command16_Click
End With
End Sub
'
Sub PRILISMAT()
    '
    'ACA VA CODIGO PARA IMPRIMIR OBSERVACIONES
    'Call SETULTREG("!OBSERVOF", 0)
    '
    ' IMPRESION DE LA LISTA DE MATERIALES
    '
    FORIPRE$ = CONTROL$("", "LISTAMAT")
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
' VER ESTO CON WALTER
' TODA LA PARTE DE LOS PARAMETROS TENES QUE HACERLA VOS, FLAVIO
      FECHDOC$ = Text9
      NUMEDOC$ = TRIM$(Label31) + " / " + TRIM$(Label2)
      TIPODOC$ = "Lista de Materiales "
      '
      DESTIDOC% = Val(Label22) 'CLIENTE
      '
      CODPARAM$(1) = "DOCUORIG" 'NUMERO DE OBRA
      CODPARAM$(2) = "REF-MAT1" 'CODIGO DE GRUPO
      CODPARAM$(3) = "RASO-CLI" 'CLIENTE
      CODPARAM$(4) = "FECH-EMI" 'FECHA DE EMISION
      CODPARAM$(5) = "REF-MAT2" 'DENOMINACION DE LA OBRA
      CODPARAM$(6) = "REF-MAT3" 'DENOMINACION DEL GRUPO
      CODPARAM$(7) = "TELE-CLI" 'RESPONSABLE
      
      CAPARDOC% = 7
      
      DOCPARAM(1) = Label31
      DOCPARAM(2) = Label2
      DOCPARAM(3) = Label5
      DOCPARAM(4) = Text9
      DOCPARAM(5) = Label7
      DOCPARAM(6) = Label8
      DOCPARAM(7) = Text11
      '
 'A PARTIR DE AQUI LO MODIFIQUE YO - EPB
      CAITAB1% = 0
      CAITCRU% = 0
      '
 'AQUI PONER LOS CODIGOS DE COLUMNAS A LISTAR
      CODTABU1$(1) = "ORD-ITEM"
      CODTABU1$(2) = "CANTIDAD"
      CODTABU1$(3) = "DES-MAT"
      CODTABU1$(4) = "N-PLANO"
      CODTABU1$(5) = "MEDIDA-1"
      CODTABU1$(6) = "MEDIDA-2"
      CODTABU1$(7) = "COD-MAT"
      CODTABU1$(8) = "MEDIDA-3"
      
      CACOLTAB1% = 8
      '
      For U& = 1 To List3.ListCount
        '
        TTXX$ = List3.List(U& - 1)
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = Left$(TTXX$, 2)
        TETABU1$(2, CAITAB1%) = Mid$(TTXX$, 4, 6)
        TETABU1$(3, CAITAB1%) = Mid$(TTXX$, 11, 36)
        TETABU1$(4, CAITAB1%) = Mid$(TTXX$, 48, 14)
        TETABU1$(5, CAITAB1%) = Mid$(TTXX$, 63, 5)
        TETABU1$(6, CAITAB1%) = Mid$(TTXX$, 69, 12)
        TETABU1$(7, CAITAB1%) = Mid$(TTXX$, 82, 16)
        TETABU1$(8, CAITAB1%) = Mid$(TTXX$, 98, 2)
        '
        'AQUI AGREGA LA DESCRIPCION LARGA
        DELARGA$ = Mid$(TTXX$, 129)
        Call FRATEXTO$(DELARGA$, 36)
        For UI% = 1 To CARETEX%
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = ""
          TETABU1$(2, CAITAB1%) = ""
          TETABU1$(3, CAITAB1%) = RETEX$(UI%)
          TETABU1$(4, CAITAB1%) = ""
          TETABU1$(5, CAITAB1%) = ""
          TETABU1$(6, CAITAB1%) = ""
          TETABU1$(7, CAITAB1%) = ""
          TETABU1$(8, CAITAB1%) = ""
        Next UI%
        '
        'ESTO ES PARA QUE EL PROXIMO EMPIECE EN UN NUEVO RENGLON TRIPLE
        While (CAITAB1% Mod 3) <> 0
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = ""
          TETABU1$(2, CAITAB1%) = ""
          TETABU1$(3, CAITAB1%) = ""
          TETABU1$(4, CAITAB1%) = ""
          TETABU1$(5, CAITAB1%) = ""
          TETABU1$(6, CAITAB1%) = ""
          TETABU1$(7, CAITAB1%) = ""
          TETABU1$(8, CAITAB1%) = ""
        Wend
        '
      Next U&
      '
      NOCONFIRM% = 1
      Call PRINTDOC("LISTAMAT")                   ' LISTA DE MATERIALES
      '
      'LIMPIA LA LISTA
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
    End If
    '
199 Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1

End Sub

Sub CARGOPLACO() 'CARGO PLAN DE COBRA. DE UNA OT DETERMINADA
   '
   NUOTR$ = TRIM$(Text1)
   If NUOTR$ = "" Or ECOARCH("INDIOTRA", Text1) = "" Then Exit Sub
   '
   Call ABREORTRA
   SQLX$ = "SELECT * FROM PLACOB "
   SQLX$ = SQLX$ + "WHERE Nro_Otrab = '" & NUOTR$ & "' "
   SQLX$ = SQLX$ + "ORDER BY FeVen_Pcob ASC "
   Set OTTMP = OTrabajo.OpenRecordset(SQLX$, 4)
   With OTTMP
      '
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      '
      Do Until .EOF = True
         NOITOT1$ = CStr(!NuOrItem)
         DESCRIOT1$ = CStr(!Refe_Pcob)
         IMPOT1$ = FORMATNUM(!Impo_Pcob, "F10.2")
         FEOT1$ = FECHATEX(CVINT(!FeVen_Pcob))
         Y$ = Space$(64)
         Call REPLA(Y$, AJUSTD$(NOITOT1$, 3), 1, 3)
         Call REPLA(Y$, AJUSTI$(DESCRIOT1$, 15), 6, 15)
         Call REPLA(Y$, IMPOT1$, 21, 10)
         Call REPLA(Y$, FEOT1$, 33, 8)
         List1.AddItem Y$
        .MoveNext
      Loop
      
      '
   End With
End Sub

Sub CARGOLISMAT() 'CARGO LISTA DE MATERIALES DE UNA OT Y GRUPO DETERMINADA
   '
   NUOTR$ = TRIM$(Label31)
   NUGRU$ = TRIM$(Label2)
   '
   If NUOTR$ = "" Or ECOARCH("INDIOTRA", NUOTR$) = "" Then Exit Sub
   If NUGRU$ = "" Then Exit Sub
   '
   Call ABREORTRA
   SQLX$ = "SELECT * FROM LISTAMAT "
   SQLX$ = SQLX$ + "WHERE Nro_Otrab = '" & NUOTR$ & "' "
   SQLX$ = SQLX$ + "AND Nro_Grup = '" & NUGRU$ & "' "
   SQLX$ = SQLX$ + "ORDER BY NuOrden ASC; "
   Set LIMAGRU = OTrabajo.OpenRecordset(SQLX$, 4)
   '
   List3.Clear
   '
   With LIMAGRU
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do Until .EOF = True
         TTXX$ = Space$(128)
         NORDE% = List3.ListCount + 1
         Call REPLA(TTXX$, FORMATNUM$(NORDE%, "I2"), 1, 2)
         Call REPLA(TTXX$, FORMATNUM$(!CANTID, "F6.1"), 4, 6)
         Call REPLA(TTXX$, !DESCRI_COR, 11, 36)
         Call REPLA(TTXX$, !NPlano, 48, 14)
         Call REPLA(TTXX$, !Posic, 63, 5)
         Call REPLA(TTXX$, !Material, 69, 12)
         Call REPLA(TTXX$, !CODEXTE, 82, 16)
         Call REPLA(TTXX$, !CLAVE, 98, 2)
         If IsNull(!Nro_Ocom) Then
            NROCOM$ = ""
          Else
            NROCOM$ = !Nro_Ocom
         End If
         Call REPLA(TTXX$, NROCOM$, 105, 12)
         If IsNull(!NRO_BOLRE) Then
            NROBORE$ = ""
           Else
            NROBORE$ = !NRO_BOLRE
         End If
         Call REPLA(TTXX$, NROBORE$, 117, 12)
         TTXX$ = TTXX$ + !Descri_Lar
         List3.AddItem TTXX$
       .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   '
   'CARGO LOS PLAZOS DEL GRUPO
    SQLX$ = "SELECT * FROM GRUPOTRA "
    SQLX$ = SQLX$ + "WHERE Nro_Otrab = '" & NUOTR$ & "' "
    SQLX$ = SQLX$ + "AND Nro_Grup = '" & NUGRU$ & "' "
    Set PLATMP = OTrabajo.OpenRecordset(SQLX$, 4)
    With PLATMP
' VER ESTO CON WALTER
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       If IsNull(!PLA05_GRUP) Then
         Text8 = ""
       Else
         Text8 = !PLA05_GRUP
       End If
       If IsNull(!PLA07_GRUP) Then
         Text12 = ""
       Else
         Text12 = !PLA07_GRUP
       End If
     Else
        On Error GoTo 0
        Text8 = ""
        Text12 = ""
     End If
    End With
   '
End Sub

Sub GRALISMAT()
   '
   Call ABREORTRA
   '
   NUOTR$ = TRIM$(Label31)
   NUGRU$ = TRIM$(Label2)
   '
   'Aca grabo los Plazos de Entrega del Grupo en el Grupo
   With GrupOtra
     .FindFirst "Nro_otrab= '" & NUOTR$ & "' AND NRO_GRUP = '" & NUGRU$ & "'"
     If .NoMatch = False Then
        .Edit
           !PLA05_GRUP = TRIM$(Text8)
           !PLA07_GRUP = TRIM$(Text12)
        .Update
     End If
   End With
   '
   With ListaMat
20   .FindFirst "Nro_Otrab = '" & NUOTR$ & "' AND Nro_Grup = '" & NUGRU$ & "'"
     If .NoMatch = False Then
       .Delete
       GoTo 20
     End If
     '
     For U& = 1 To List3.ListCount
       Y$ = List3.List(U& - 1)
       .AddNew
       !Nro_Otrab = NUOTR$
       !Nro_Grup = NUGRU$
         NUOR% = U&
       !NuOrden = NUOR%
       !CANTID = Val(Mid$(Y$, 4, 8))
       !DESCRI_COR = Mid$(Y$, 11, 36)
       !NPlano = Mid$(Y$, 48, 14)
       !Posic = Mid$(Y$, 63, 5)
       !Material = Mid$(Y$, 69, 12)
       !CODEXTE = Mid$(Y$, 82, 16)
       !CLAVE = Mid$(Y$, 98, 2)
       !Nro_Ocom = Mid$(Y$, 105, 12)
       !NRO_BOLRE = Mid$(Y$, 117, 12)
       !Descri_Lar = Mid$(Y$, 129)
       '!Nro_Ocom = ""
       .Update
     Next U&
     '
   End With
   '
End Sub

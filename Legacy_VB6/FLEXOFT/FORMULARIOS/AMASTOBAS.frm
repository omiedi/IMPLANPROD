VERSION 5.00
Begin VB.Form AMASTOBAS 
   BackColor       =   &H00D0F0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Artículos - A-B-M-L Maestro de Items de Stock"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   Icon            =   "AMASTOBAS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame MARCOFICHA 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Atributos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6525
      Index           =   0
      Left            =   210
      TabIndex        =   27
      Top             =   1575
      Width           =   10320
      Begin VB.Frame Frame7 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Características"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   210
         TabIndex        =   47
         Top             =   420
         Width           =   4740
         Begin VB.TextBox Text6 
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
            Left            =   1155
            MaxLength       =   3
            TabIndex        =   54
            Text            =   " "
            Top             =   525
            Width           =   570
         End
         Begin VB.CommandButton Command11 
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
            Left            =   1700
            TabIndex        =   53
            Top             =   525
            Width           =   175
         End
         Begin VB.TextBox Text4 
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
            Left            =   1155
            MaxLength       =   2
            TabIndex        =   52
            Text            =   " "
            Top             =   975
            Width           =   570
         End
         Begin VB.CommandButton Command10 
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
            Left            =   1700
            TabIndex        =   51
            Top             =   975
            Width           =   175
         End
         Begin VB.TextBox Text24 
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
            Left            =   1155
            TabIndex        =   50
            Top             =   1425
            Width           =   1200
         End
         Begin VB.CommandButton Command20 
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
            Left            =   4380
            TabIndex        =   49
            Top             =   1425
            Width           =   175
         End
         Begin VB.TextBox Text25 
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
            Left            =   3420
            TabIndex        =   48
            Text            =   " "
            Top             =   1425
            Width           =   945
         End
         Begin VB.Label Label29 
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
            Left            =   1995
            TabIndex        =   60
            Top             =   525
            Width           =   2535
         End
         Begin VB.Label Label30 
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
            Left            =   1995
            TabIndex        =   59
            Top             =   975
            Width           =   2535
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Kilos por Unidad:"
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
            Left            =   -210
            TabIndex        =   58
            Top             =   1365
            Width           =   1290
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Unidad de Medida:"
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
            Left            =   -105
            TabIndex        =   57
            Top             =   910
            Width           =   1185
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Material:"
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
            Left            =   -315
            TabIndex        =   56
            Top             =   420
            Width           =   1380
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha de Revisión:"
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
            Left            =   2205
            TabIndex        =   55
            Top             =   1410
            Width           =   1170
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Reposición"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   5145
         TabIndex        =   36
         Top             =   420
         Width           =   4980
         Begin VB.CommandButton Command12 
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
            Left            =   1600
            TabIndex        =   41
            Top             =   525
            Width           =   175
         End
         Begin VB.TextBox Text20 
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
            Left            =   3465
            TabIndex        =   40
            Top             =   1425
            Width           =   1305
         End
         Begin VB.TextBox Text8 
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
            Left            =   1260
            MaxLength       =   2
            TabIndex        =   39
            Text            =   " "
            Top             =   525
            Width           =   360
         End
         Begin VB.TextBox Text18 
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
            Left            =   315
            TabIndex        =   38
            Top             =   1425
            Width           =   1305
         End
         Begin VB.TextBox Text19 
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
            Left            =   1890
            TabIndex        =   37
            Top             =   1425
            Width           =   1305
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Lote Máximo:"
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
            Left            =   3255
            TabIndex        =   46
            Top             =   1155
            Width           =   1245
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Minimo:"
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
            Left            =   0
            TabIndex        =   45
            Top             =   1155
            Width           =   1395
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Económico:"
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
            Left            =   1575
            TabIndex        =   44
            Top             =   1155
            Width           =   1245
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Criterio de Reposicion:"
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
            Left            =   -210
            TabIndex        =   43
            Top             =   420
            Width           =   1380
         End
         Begin VB.Label Label32 
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
            Left            =   1890
            TabIndex        =   42
            Top             =   525
            Width           =   2900
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   2310
         ScaleHeight     =   375
         ScaleWidth      =   7725
         TabIndex        =   34
         Top             =   2835
         Width           =   7780
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Descripción General - Especificaciones"
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
            Left            =   0
            TabIndex        =   35
            Top             =   105
            Width           =   7680
         End
      End
      Begin VB.CommandButton Command7 
         Height          =   260
         Left            =   2310
         TabIndex        =   33
         Top             =   5970
         Width           =   2825
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Left            =   8760
         TabIndex        =   32
         Top             =   5970
         Width           =   1320
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Pegar"
         Height          =   260
         Left            =   7480
         TabIndex        =   31
         Top             =   5970
         Width           =   1275
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Copiar"
         Height          =   260
         Left            =   6310
         TabIndex        =   30
         Top             =   5970
         Width           =   1170
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Cortar"
         Height          =   260
         Left            =   5130
         TabIndex        =   29
         Top             =   5970
         Width           =   1170
      End
      Begin VB.TextBox Text28 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2745
         Left            =   2310
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   3255
         Width           =   7785
      End
      Begin VB.Line Line1 
         X1              =   10090
         X2              =   10090
         Y1              =   2835
         Y2              =   6220
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción y Especificaciones:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   315
         TabIndex        =   61
         Top             =   3255
         Width           =   1905
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00D0F0F0&
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
      Height          =   2800
      Left            =   10740
      TabIndex        =   21
      Top             =   5300
      Width           =   1040
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
         Left            =   125
         Picture         =   "AMASTOBAS.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Imprime Ficha del Artículo"
         Top             =   1920
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
         Height          =   690
         Left            =   125
         Picture         =   "AMASTOBAS.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Emisión de Listados"
         Top             =   1170
         Width           =   765
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
         Height          =   690
         Left            =   125
         Picture         =   "AMASTOBAS.frx":0C7E
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   420
         Width           =   765
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   1110
      Left            =   7140
      ScaleHeight     =   1050
      ScaleWidth      =   60
      TabIndex        =   19
      Top             =   315
      Width           =   120
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   5000
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D0F0F0&
      Caption         =   "A-B-M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Left            =   10740
      TabIndex        =   9
      Top             =   1600
      Width           =   1040
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
         Left            =   125
         Picture         =   "AMASTOBAS.frx":10C0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Alta de Nuevo Item por Copia Completa de Datos de Artículo Existente"
         Top             =   1170
         Width           =   765
      End
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
         Height          =   690
         Left            =   125
         Picture         =   "AMASTOBAS.frx":120A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Alta de Nuevo Item de Stock"
         Top             =   420
         Width           =   750
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Subst"
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
         Left            =   125
         Picture         =   "AMASTOBAS.frx":1354
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Sustitución de Código de Artículo"
         Top             =   1920
         Width           =   765
      End
      Begin VB.CommandButton Command24 
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
         Left            =   125
         Picture         =   "AMASTOBAS.frx":149E
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Baja Lógica Individual"
         Top             =   2670
         Width           =   765
      End
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
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
      Top             =   840
      Width           =   5415
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
      Picture         =   "AMASTOBAS.frx":17A8
      TabIndex        =   4
      Top             =   315
      Width           =   175
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00D0F0F0&
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   7455
      TabIndex        =   0
      Top             =   210
      Width           =   4320
      Begin VB.CommandButton Command25 
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
         Height          =   750
         Left            =   105
         Picture         =   "AMASTOBAS.frx":1AB2
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   315
         Width           =   750
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
         Height          =   750
         Index           =   4
         Left            =   3400
         Picture         =   "AMASTOBAS.frx":1DBC
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   315
         Width           =   750
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
         Height          =   750
         Left            =   2650
         Picture         =   "AMASTOBAS.frx":1F06
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Ayuda en Línea"
         Top             =   315
         Width           =   750
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Restore"
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
         Left            =   1755
         Picture         =   "AMASTOBAS.frx":2210
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Restablece los Valores Anteriores"
         Top             =   315
         Width           =   750
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
         Height          =   750
         Left            =   1020
         Picture         =   "AMASTOBAS.frx":251A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   315
         Width           =   750
      End
   End
   Begin VB.Label Label58 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   210
      TabIndex        =   25
      Top             =   1155
      Visible         =   0   'False
      Width           =   960
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
      Left            =   3465
      TabIndex        =   17
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
      Left            =   4560
      TabIndex        =   16
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
      Left            =   5040
      TabIndex        =   15
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
      Left            =   6195
      TabIndex        =   14
      Top             =   315
      Width           =   750
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
      Top             =   840
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
      Left            =   375
      TabIndex        =   7
      Top             =   420
      Width           =   1005
   End
End
Attribute VB_Name = "AMASTOBAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TOPEINI(3), IZQUINI(3), ANCHINI(3), ALTUINI(3)
'
Dim CIACTI%(15)
Dim ARCHEX$(15)
Dim YACAR%(15), MODIFIC%(15)
Dim GRABATODO%
Dim AGREITEM%
Dim MODOEDIT%
'
Private Sub Command1_Click()
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   GRABATODO% = 1
   Call GRADATITEM
   GRABATODO% = 0
End Sub

Private Sub Command10_Click()
   Call SELECHO("UMEDIDA")
   VALONUE$ = VALACT1$("UMEDIDA")
   If VALONUE$ <> "" Then
      Text4.TEXT = VALONUE$
   End If
End Sub

Private Sub Command11_Click()
   Call SELECHO("TATIMAT")
   VALONUE$ = VALACT1$("TATIMAT")
   If VALONUE$ <> "" Then
      Text6.TEXT = VALONUE$
   End If
End Sub

Private Sub Command12_Click()
   Call SELECHO("TACRIRE")
   VALONUE$ = VALACT1$("TACRIRE")
   If VALONUE$ <> "" Then
      Text8.TEXT = VALONUE$
   End If
End Sub

Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command2_Click()
   '
   If COMALTER%("Realmente Desea Descartar los Cambios ?\\No, Continuar\Si, Descartar Cambios") <> 2 Then
      Exit Sub
   End If
   '
   For Index = 0 To 15
     YACAR%(Index) = 0
     MODIFIC%(Index) = 0
   Next Index
   '
   MODOEDIT% = 0
   Call BLANFORMU
   If CODINT%(Text1.TEXT) > 0 Then
       Call CARDATITEM
   End If
   '
End Sub

Private Sub Command20_Click()
   Call SELECHO("SELFECHA")
   VALONUE$ = VALACT1$("SELFECHA")
   If VALONUE$ <> "" Then
      Text25.TEXT = VALONUE$
   End If
End Sub

Private Sub Command21_Click()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command21.Enabled = False
   Call GRADATITEM
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
   End If
   '
99 Command21.Enabled = True
   '
End Sub

Private Sub Command22_Click()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Command22.Enabled = False
   Call GRADATITEM
   If VALCODART% < 1 Then GoTo 99
   '
   CODIVIE$ = AJUSTI$(Text1.TEXT, 16)
   CII% = CODINT%(Text1.TEXT)
   If CII% < 1 Or CII% > NUMAS% Then
      Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
      Command25.SetFocus
      GoTo 99
   End If
   '
   REMAS& = 1 + CII%
   If REMAS& < 2 Or REMAS& > ULTREG&("MASTER") Then
      Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
      Command25.SetFocus
      GoTo 99
   End If
   '
   VDEF$ = REGLEIDO$("MASTER", REMAS&)
   CINUE% = 1 + NUMAS%
   Call REPLA(VDEF$, MKI$(CINUE%), 1, 2)
   Call REPLA(VDEF$, Space$(16), 77, 16)
   '
   OKX% = 0
   Call ALTACODIGO(VDEF$, OKX%)
   If OKX% = 1 Then
      Call COPYFORMU(CII%, CINUE%)
      Call COPYRUTA(CII%, CINUE%)
      Call COPYSPEC(CII%, CINUE%)
      CODINUE$ = Mid$(VDEF$, 77, 16)
      Text1.TEXT = CODINUE$
   End If
   '
99 Command22.Enabled = True
   '
End Sub

Private Sub Command23_Click()
   '
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   '
   Command23.Enabled = False
   If VALCODART% < 1 Then GoTo 99
   '
   VDEF$ = ""
10 OBX$ = OBJPLA$("ENTRACODNUE", VDEF$)
   If OBX$ = "" Then GoTo 99
   '
   VDEF$ = OBX$
   CODINUE$ = TRIM$(OBX$)
   CODIVIE$ = AJUSTI$(Text1.TEXT, 16)
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
   Screen.MousePointer = 11
   For U& = 1 To Len(PUNCODIG$) Step 16
      If Mid$(PUNCODIG$, U&, 16) = CODIVIE$ Then
         Mid$(PUNCODIG$, U&, 16) = AJUSTI$(CODINUE$, 16)
         GoTo 20
      End If
   Next U&
   Call COMUNI("Operación Cancelada (1) - Reintente")
   GoTo 99
   '
20 REMAS& = 1 + (U& + 15) / 16
   NMSI% = REMAS& - 1
   XX$ = REGLEIDO$("MASTER", REMAS&)
   If Mid$(XX$, 77, 15) <> Left$(CODIVIE$, 15) Then
      Call COMUNI("Operación Cancelada (2) - Reintente")
      GoTo 99
   End If
   Call REPLA(XX$, CODINUE$, 77, 16)
   Call GRAREG("MASTER", XX$, REMAS&)
   Call CIERRARCH("MASTER")
   '
   PPCC$ = PUNCODIG$
   NX% = FILEOPEN%(LUDAT$ + "PUNCODIG.DAT", 0, 2048)
   '
   JI% = 0
   While Len(PPCC$) > 0
      BUFERCOD$ = Left$(PPCC$ + String$(2048, 0), 2048)
      JI% = JI% + 1
      Put #NX%, JI%, BUFERCOD$
      PPCC$ = Mid$(PPCC$, 2049)
   Wend
   Close #NX%
   '
   EL1$ = REGBLAN$("ECOMAST")
   Call REPLA(EL1$, CODINUE$, 1, 16)
   Call REPLA(EL1$, Text3.TEXT, 17, 46)
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
      If Mid$(TTXX$, U&, 16) = CODIVIE Then
         TTXX$ = Left$(TTXX$, U& - 1) + Mid$(TTXX$, U& + 64)
         GoTo 30
      End If
   Next U&
   '
30 For U& = 1 To Len(TTXX$) Step 64
      If Mid$(TTXX$, U&, 16) > CODINUE$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + EL1$ + Mid$(TTXX$, U&)
         GoTo 40
      End If
   Next U&
   TTXX$ = TTXX$ + EL1$
   '
40 JJ& = 0: LOTEGRA& = Len(TTXX$)
   TTXX$ = TTXX$ + String$(2048, 0)
   For U& = 1 To LOTEGRA& Step 2048
      EL$ = Mid$(TTXX$, U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   '
50 NVII& = 1 + NUINV%
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
   For U& = 19 To Len(TTXX$) Step 18
      If Mid$(TTXX$, U&, 16) = CODIVIE$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + Mid$(TTXX$, U& + 18)
         GoTo 60
      End If
   Next U&
   '
60 For U& = 19 To Len(TTXX$) Step 18
      If Mid$(TTXX$, U&, 16) > CODINUE$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + EL1$ + Mid$(TTXX$, U&)
         GoTo 70
      End If
   Next U&
   TTXX$ = TTXX$ + EL1$
   '
70 JJ& = 0
   For U& = 1 To Len(TTXX$) Step 2048
      EL$ = Mid$(TTXX$, U&, 2048)
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
   Screen.MousePointer = 1
   '
   Text1.TEXT = CODINUE$
   '
99 Screen.MousePointer = 1
   Command23.Enabled = True
   '
End Sub

Private Sub Command24_Click()
   '
   Command24.Enabled = False
   Dim EL As String * 2048
   '
   If VALCODART% < 1 Then
      Text1.SetFocus
      GoTo 99
   End If
   '
   CII% = CODINT%(Text1.TEXT)
   COD$ = AJUSTI$(CODEXT$(CII%), 15)
   '
   If COMALTER%(" Baja Lógica del Codigo '" + TRIM$(COD$) + "' \" + TRIM$(DESCRIT$(CI%)) + "\\Cancelar\Confirmar") <> 2 Then GoTo 99
   '
   Screen.MousePointer = 11
   Call BLANFORMU
   Screen.MousePointer = 11
   Text1.TEXT = ""
   If CII% < 1 Or CII% > NUMAS% Then
      Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
      Command25.SetFocus
      GoTo 99
   End If
   '
   Screen.MousePointer = 11
   REGMA& = CII% + 1
   X$ = REGLEIDO$("MASTER", REGMA&)
   If Mid$(X$, 77, 15) <> AJUSTI$(COD$, 15) Then
      Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
      Command25.SetFocus
      GoTo 99
   End If
   '
   Call REPLA(X$, Chr$(96), 92, 1)
   Call GRAREG("MASTER", X$, REGMA&)
   Call CIERRARCH("MASTER")
   '
10 NVII& = 1 + NUINV%
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
   For U& = 19 To Len(TTXX$) Step 18
      If Mid$(TTXX$, U&, 15) = COD$ Then
         Mid$(TTXX$, U&, 16) = COD$ + Chr$(96)
         GoTo 20
      End If
   Next U&
   Close #FIECO%
   Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
   Command25.SetFocus
   GoTo 99
   '
20 JJ& = 0
   For U& = 1 To Len(TTXX$) Step 2048
      EL$ = Mid$(TTXX$, U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   '
   For U& = 1 To Len(PUNCODIG$) Step 16
      If Mid$(PUNCODIG$, U&, 15) = COD$ Then
         Mid$(PUNCODIG$, U&, 16) = COD$ + Chr$(96)
         GoTo 25
      End If
   Next U&
   Call MENSERR(24, "Posible Inconsistencia de la Base de Datos !!!\Ejecute Control de Integridad.")
   Command25.SetFocus
   GoTo 99
   '
25 PPCC$ = PUNCODIG$
   NX% = FILEOPEN%(LUDAT$ + "PUNCODIG.DAT", 0, 2048)
   JI% = 0
   While Len(PPCC$) > 0
      BUFERCOD$ = Left$(PPCC$ + String$(2048, 0), 2048)
      JI% = JI% + 1
      Put #NX%, JI%, BUFERCOD$
      PPCC$ = Mid$(PPCC$, 2049)
   Wend
   Close #NX%
   '
   LOTOTE& = 64 * ULTREG&("ECOMAST")
   FIECO% = FILEOPEN%(LUDAT$ + "ECOMAST.DAT", 0, 2048)
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
      If Mid$(TTXX$, U&, 15) = COD$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + Mid$(TTXX$, U& + 64)
         GoTo 30
      End If
   Next U&
   '
30 While (Len(TTXX$) Mod 2048) <> 0: TTXX$ = TTXX$ + Chr$(0): Wend
   JJ& = 0
   
   For U& = 1 To Len(TTXX$) Step 2048
      EL$ = Mid$(TTXX$, U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   Call FRESHECHO("MASTER")
   '
99 Screen.MousePointer = 1
   Command24.Enabled = True
   '
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    CONTROLINTEG% = 1
    If NUMAS% < 1 Then GoTo 99
    '
    If COMALTER%("Realmente Desea Controlar el Maestro\\No, Cancelar\Si, Controlar") = 2 Then
      Call BLOQARCH("MASTER")
      Call REVIMASTER
      Call REPUNCODIG
    End If
99  Command25.Enabled = True
    CONTROLINTEG% = 0
End Sub

Private Sub Command3_Click(Index As Integer)
   If MODOEDIT% = 1 Then Call MENSAEDIT: Exit Sub
   Call GRADATITEM
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub Command37_Click()
   OPMAST00.Show 1
End Sub

Private Sub Command38_Click()
   Command38.Enabled = False
   OPLMAS00.Show 1
   Command38.Enabled = True
End Sub

Private Sub Command39_Click()
   Command39.Enabled = False
   PRINTFORM
   Command39.Enabled = True
End Sub


Private Sub Command5_Click()
   Text28.SelStart = 0
   Text28.SelLength = Len(Text28.TEXT)
   Text28.SetFocus
End Sub

Private Sub Command6_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Clipboard.GetText + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
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
     openforms = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    'Call TRANSLABELS(Name)
    'Show
    'Refresh
    Call CENTRAFORM(Me)
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
   Text4.TEXT = ""
   Label30.Caption = ""
   Text6.TEXT = ""
   Label29.Caption = ""
   Text8.TEXT = ""
   Label32.Caption = ""
   Text18.TEXT = ""
   Text19.TEXT = ""
   Text20.TEXT = ""
   Text24.TEXT = ""
   Text25.TEXT = ""
   '
   Text28.TEXT = ""
   '
   Text31 = ""
   Text32 = ""
   Text27 = ""
   Label40 = ""
   Label41 = ""
   '
   Label43 = ""
   Text11 = ""
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
    MODIFIC(0) = 0
    '
    Call CARGENITEM
    MODIFIC%(0) = 0
    Call CARESPECI
    MODIFIC%(1) = 0
    Screen.MousePointer = 1
    '
End Sub

Sub CARGENITEM()
   '
   If YACAR%(0) < 1 Then
      '
      CODD$ = Text1.TEXT
      CI% = CODINT%(CODD$)
      If CI% < 1 Then Exit Sub
      '
      Screen.MousePointer = 11
      REG& = CI% + 1
      REMAS$ = REGLEIDO$("MASTER", REG&)
      '
        UNID$ = UNIMED$(CI%)
      Text4.TEXT = UNID$
      Label30.Caption = ECOARCH$("UMEDIDA", UNID$)
         TIMA% = Asc(Mid$(REMAS$, 93, 1))
      Text6.TEXT = TRIM$(Str$(TIMA%))
      Label29.Caption = ECOARCH$("TATIMAT", Chr$(TIMA%))
         CRIT% = Asc(Mid$(REMAS$, 94, 1))
      Text8.TEXT = TRIM$(Str$(CRIT%))
      Label32.Caption = ECOARCH$("TACRIRE", Chr$(CRIT%))
      Text18.TEXT = CSTRING$(MKS$(STOMIN(CI%)), 3, 10, 1, 2)
      Text19.TEXT = CSTRING$(MKS$(LOTESTAN(CI%)), 3, 10, 1, 2)
      Text20.TEXT = CSTRING$(MKS$(LOREPOMA(CI%)), 3, 10, 1, 2)
         FEX = FEREVI%(CI%)
      Text24 = FORMATNUM$(PESUNI(CI%), "F9.5")
      Text25.TEXT = FECHATEX$(FEX)
         'CCOMPO% = Asc(Mid$(REMAS$, 98, 1))
         'Text25.TEXT = TRIM$(Str$(CCOMPO%))
         'Label39.Caption = TRIM$(ECOARCH$("TACOMPO", Chr$(CCOMPO%)))
      Text36 = REVICOD$(CI%)
      '
      Screen.MousePointer = 1
      '
   End If
   '
End Sub

Sub CARESPECI()
   '
   If YACAR%(0) < 1 Then
      '
      CODD$ = Text1.TEXT
      CI% = CODINT%(CODD$)
      If CI% < 1 Then Exit Sub
      '
      Text28.TEXT = "": TTYZ$ = ""
      If CI% > 0 Then
         If CI% <= NUMAS% Then
            Screen.MousePointer = 11
            RNC$ = RECFILT$("ESPECIFI", 1, MKI$(CI%))
            URN& = ULTREG&("ESPECIFI")
            '
            For U& = 1 To Len(RNC$) Step 4
               RENOTA& = CVS(Mid$(RNC$, U&, 4))
               If RENOTA& > 0 Then
                  If RENOTA& <= URN& Then
                     TTYY$ = RTrim$(Mid$(REGLEIDO$("ESPECIFI", RENOTA&), 3))
                     If TTYZ$ <> "" Then
                        TTYZ$ = TTYZ$ + Chr$(13) + Chr$(10)
                     End If
                     TTYZ$ = TTYZ$ + RTrim$(TTYY$)
                  End If
               End If
            Next U&
            Screen.MousePointer = 1
         End If
      End If
      LOCADE& = Len(TTYZ$)
      For U& = LOCADE& To 1 Step -1
        If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 30
      Next U&
      U& = 0
30    Text28.TEXT = Left$(TTYZ$, U&)
      Text28.Refresh
      '
      YACAR%(0) = 1
      MODIFIC%(1) = 0
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
   If MODIFIC%(0) > 0 Or MODIFIC%(1) > 0 Then
       On Error Resume Next
       Command9.SetFocus
       On Error GoTo 0
       Call GRAGENITEM
       Call GRAESPECI
   End If
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
   If MODIFIC%(0) < 1 Then Exit Sub
   '
   If DERACCE%("MODITMAS", "Modificación Datos Maestro de Stock") < 2 Then
      Exit Sub
   End If
   '
   If YACAR%(0) < 1 Then    ' SOLO CAMBIO LA DESCRIPCION
      Screen.MousePointer = 11
      REG& = CI% + 1
      REMAS$ = REGLEIDO$("MASTER", REG&)
      '
      MODESCRI% = 0
      DE$ = TRIM$(Text3.TEXT)
      If DE$ = "" Then DE$ = TRIM$(Mid$(REMAS$, 5, 64))
      MODESCRI% = 1
      Call REPLA(REMAS$, DE$, 5, 64)
      GoTo 30
   End If
   '
   If YACAR%(0) < 1 Then Exit Sub
   If GRABATODO% <> 1 Then
     If COMALTER%("Se han Modificado Datos del Artículo\'" + TRIM$(CODEXT$(CI%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(0) = 0
       MODIFIC(0) = 0
       Call CARGENITEM
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   '
   TTITPE$ = TRIM$(Text24.TEXT)
   PESU = Val(TTITPE$)
   If InStr(TTITPE$, "/") > 0 Then
      TTQP$ = TTITPE$
      TTQP$ = REPLACAR$(TTQP$, "/", ".")
      PESU = (-1) * Abs(Val(TTQP$))
   End If
   '
   REG& = CI% + 1
   REMAS$ = REGLEIDO$("MASTER", REG&)
   '
   MODESCRI% = 0
   DE$ = TRIM$(Text3.TEXT)
   If DE$ = "" Then DE$ = TRIM$(Mid$(REMAS$, 5, 64))
   If AJUSTI$(DE$, 64) <> Mid$(REMAS$, 5, 64) Then
      MODESCRI% = 1
   End If
   Call REPLA(REMAS$, DE$, 5, 64)
   '
   TIMA% = Val(Text6.TEXT)
   Call REPLA(REMAS$, Chr$(TIMA%), 93, 1)
   '
   CRIT% = Val(Text8.TEXT)
   Call REPLA(REMAS$, Chr$(CRIT%), 94, 1)
   '
   CCOMPO% = Val(Text25.TEXT)
   Call REPLA(REMAS$, Chr$(CCOMPO%), 98, 1)
   '
   Call REPLA(REMAS$, MKS$(PESU), 99, 4)
   '
   STMII% = (-1) * Val(Text18) / 100: If Val(Text18) < 32000 Then STMII% = Val(Text18)
   LOREI% = (-1) * Val(Text19) / 100: If Val(Text19) < 32000 Then LOREI% = Val(Text19)
   LOTEMAI% = (-1) * Val(Text20) / 100: If Val(Text20) < 32000 Then LOTEMAI% = Val(Text20)
   If Abs(LOTEMAI%) < Abs(LOREI%) Then LOTEMAI% = LOREI%
   '
   Call REPLA(REMAS$, MKI$(LOTEMAI%), 103, 2)
   Call REPLA(REMAS$, MKI$(STMII%), 105, 2)
   Call REPLA(REMAS$, MKI$(LOREI%), 107, 2)
   '
   UNID$ = TRIM$(Text4.TEXT)
   Call REPLA(REMAS$, UNID$, 109, 2)
   '
   FERE% = FECHANUM(Text25.TEXT)
   Call REPLA(REMAS$, MKI$(FERE%), 121, 2)
   '
30 Call GRAREG("MASTER", REMAS$, REG&)
   Call CIERRARCH("MASTER")
   '
   If MODESCRI% > 0 Then
      CODARTIC$ = AJUSTI$(Text1.TEXT, 16)
      LOTOTE& = 64 * ULTREG&("ECOMAST")
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
        If Mid$(TTXX$, U&, 16) = CODARTIC$ Then
          Mid$(TTXX$, U& + 16, 46) = AJUSTI$(DE$, 46)
          GoTo 40
        End If
      Next U&
      '
40    JJ& = 0
      For U& = 1 To Len(TTXX$) Step 2048
        EL$ = Mid$(TTXX$ + String$(2048, 0), U&, 2048)
        If Len(EL$) < 2048 Then
          EL$ = Left$(EL$ + String$(2048, 0), 2048)
        End If
        JJ& = JJ& + 1
        Put #FIECO%, JJ&, EL$
      Next U&
      Close #FIECO%
      Call FRESHECHO("ECOMAST")
      '
   End If
   '
   If CI% > 1 Then     ' fuerza actualizacion de descripcion
      DDD = DESCRIT$(CI% - 1)
   End If
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Screen.MousePointer = 1
   MODESCRI% = 0
   MODIFIC%(0) = 0
   '
End Sub
'
Sub GRAESPECI()
   '
   CI% = CODINT%(Text1.TEXT)
   If CI% < 1 Or CI% > NUMAS% Then Exit Sub
   If YACAR%(0) < 1 Then Exit Sub
   If MODIFIC%(1) < 1 Then Exit Sub
   '
   If DERACCE%("MODISPEC", "Modificación de Especificaciones") < 2 Then
      Exit Sub
   End If
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se ha Modificado la Especificación del Artículo\'" + TRIM$(CODEXT$(CI%)) + " - " + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR%(1) = 0
       MODIFIC(0) = 0
       Call CARESPECI
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   URN& = ULTREG&("ESPECIFI")
   RNC$ = RECFILT$("ESPECIFI", 1, MKI$(CI%))
   KK& = (-3)
   '
   TTYZ$ = Text28.TEXT
   LOCADE& = Len(TTYZ$)
   For U& = LOCADE& To 1 Step -1
     If Asc(Mid$(TTYZ$, U&, 1)) > 32 Then GoTo 10
   Next U&
   U& = 0
   '
10 TEXTO$ = Left$(TTYZ$, U&) + Chr$(13) + Chr$(10)
   Call FRATEXTO(TEXTO$, 62)
   For LL% = 1 To CARETEX%
      TTYY$ = MKI$(CI%) + RETEX$(LL%) + Space$(62)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG("ESPECIFI", TTYY$, RENOTA&)
   Next LL%
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("ESPECIFI", MKI$(0) + Space$(62), RENOTA&)
   Wend
   '
   Call CIERRARCH("ESPECIFI")
   MODIFIC%(1) = 0
   Screen.MousePointer = 1
   '
End Sub
'
Private Sub Text1_Change()
   '
   For Index = 0 To 15: YACAR%(Index) = 0: Next Index
   '
   'If IUCOD% < 128 Then
   '  If TRIM$(Text1) <> "" Then
   '    If Text1 <> UCOD$(IUCOD%) Then
   '      IUCOD% = IUCOD% + 1: UCOD$(IUCOD%) = COD$
   '                           ICOMPO%(IUCOD%) = List1.ListIndex
   '    End If
   '  End If
   'End If
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

Private Sub Text18_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text18_LostFocus()
   Text18.TEXT = CSTRING$(MKS$(Val(Text18.TEXT)), 3, 10, 1, 2)
End Sub

Private Sub Text19_Change()
   MODIFIC%(0) = 1
   If Val(Text20.TEXT) < Val(Text19.TEXT) Then
     Text20.TEXT = CSTRING$(MKS$(Val(Text19.TEXT)), 3, 10, 1, 2)
   End If
End Sub

Private Sub Text19_LostFocus()
   Text19.TEXT = CSTRING$(MKS$(Val(Text19.TEXT)), 3, 10, 1, 2)
End Sub

Private Sub Text20_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text20_LostFocus()
   Text20.TEXT = CSTRING$(MKS$(Val(Text20.TEXT)), 3, 10, 1, 2)
End Sub

Private Sub Text24_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text24_LostFocus()
   Text24.TEXT = CSTRING$(MKS$(Val(Text24.TEXT)), 3, 10, 5, 2)
End Sub

Private Sub Text25_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text25_DblClick()
   Call SELECHO("SELFECHA")
   VALONUE$ = VALACT1$("SELFECHA")
   If VALONUE$ <> "" Then
      Text25.TEXT = VALONUE$
   End If
End Sub


Private Sub Text28_Change()
   MODIFIC%(1) = 1
End Sub

Private Sub Text3_Change()
   MODIFIC%(0) = 1
End Sub

'
Private Sub Text4_Change()
   Label37.Caption = Text4.TEXT
   Label30.Caption = ECOARCH$("UMEDIDA", Text4.TEXT)
   MODIFIC%(0) = 1
End Sub

Private Sub Text4_DblClick()
   Call SELECHO("UMEDIDA")
   VALONUE$ = VALACT1$("UMEDIDA")
   If VALONUE$ <> "" Then
      Text4.TEXT = VALONUE$
   End If
End Sub


Private Sub Text6_Change()
   VALONU = Val(Text6.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text6.TEXT = "": Exit Sub
   End If
   Label29.Caption = ECOARCH$("TATIMAT", Chr$(VALONU))
   MODIFIC%(0) = 1
End Sub

Private Sub Text6_DblClick()
   Call SELECHO("TATIMAT")
   VALONUE$ = VALACT1$("TATIMAT")
   If VALONUE$ <> "" Then
      Text6.TEXT = VALONUE$
   End If
End Sub


Private Sub Text8_Change()
   VALONU = Val(Text8.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Text8.TEXT = "": Exit Sub
   End If
   Label32.Caption = ECOARCH$("TACRIRE", Chr$(VALONU))
   MODIFIC%(0) = 1
End Sub

Private Sub Text8_DblClick()
   Call SELECHO("TACRIRE")
   VALONUE$ = VALACT1$("TACRIRE")
   If VALONUE$ <> "" Then
      Text8.TEXT = VALONUE$
   End If
End Sub

Sub ALTACODIGO(VDEF$, OKX%)
   '
10 OBX$ = OBJPLA$("ALTACODIGO", VDEF$)
   If OBX$ = "" Then GoTo 99
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
   '
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
        CONTROLINTEG% = 1
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
        CONTROLINTEG% = 0
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
      TTXX$ = Space$(63)
      Call REPLA(TTXX$, "DEPOSITO CENTRAL", 1, 20)
      Call REPLA(TTXX$, "DEPOSITO", 21, 8)
      Call GRAREG("TADEPOSI", Chr$(1) + TTXX$, 1): Call CIERRARCH("TADEPOSI")
      Call GRAREG("ECODEP", Chr$(1) + Left$(TTXX$, 20), 1): Call CIERRARCH("ECODEP")
      Call GRAREG("SUCURSAL", TTXX$, 1): Call CIERRARCH("SUCURSAL")
      '
      '  Screen.MousePointer = 1
      '  Call MENSERR(24, "Falta Definir Depósitos")
      '  OPMAST00.Show 1
      '  GoTo 10
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
    FIECO% = FILENBR%("LIESTRU")
    Dim ECM As String * 16
    '
    Call HEADERS("LIESTRU", "")
    Call HEADERS("LIESTRU", "Desde Codigo: " + TRIM$(cod1$) + " - " + DESCRIT$(CI1%))
    Call HEADERS("LIESTRU", "Hasta Codigo: " + TRIM$(cod2$) + " - " + DESCRIT$(CI2%))
    '
    Screen.MousePointer = 11
    NESTRU& = ULTREG&("ESTRUNUE")
    J& = 0
    '
    For I& = 1 To NESTRU&
      Call TRACE(20, I&, NESTRU&)
      CI% = CVI(Left$(REGLEIDO$("ESTRUNUE", I&), 2))
      ACTIX% = 0
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          CODYY$ = CODEXT$(CI%)
          If Mid$(CODYY$, 16, 1) <> Chr$(96) Then ' SUPRIME ITEMS DE BAJA
            If CODYY$ >= cod1$ Then
              If CODYY$ <= cod2$ Then
                Call LEEEXPLO(CODYY$, 1)
                If CAIT% > 0 Then
                  For K% = 1 To CAIT%
                    CI1% = CIJ%(K%)
                    CAN = USOI(K%)
                    PD% = PDI%(K%)
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
      End If
      '
      If ACTIX% > 0 Then
        Z$ = REGBLAN$("LIESTRU")
        Call REPLA(Z$, MKI$(CI%), 1, 2)
        J& = J& + 1
        Call GRAREG("LIESTRU", Z$, J&)
      End If
      '
    Next I&
    '
    Call SETULTREG("LIESTRU", J&)
    Call CIERRARCH("LIESTRU")
    '
    Screen.MousePointer = 1
    Call FINAL("*SOLIEST")
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


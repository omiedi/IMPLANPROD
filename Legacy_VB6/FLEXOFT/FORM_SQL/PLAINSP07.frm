VERSION 5.00
Begin VB.Form PLAINSP07 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Plan de Inspección y Ensayo"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11835
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00006060&
      ForeColor       =   &H00FFFFFF&
      Height          =   8655
      Left            =   10920
      ScaleHeight     =   8595
      ScaleWidth      =   1185
      TabIndex        =   43
      Top             =   -240
      Width           =   1245
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
         Left            =   120
         Picture         =   "PLAINSP07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Imprime Plan de Inspección"
         Top             =   5370
         Width           =   650
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
         Left            =   120
         Picture         =   "PLAINSP07.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   4680
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
         Left            =   120
         Picture         =   "PLAINSP07.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Genera Plan de Inspección COPIANDO uno Existente"
         Top             =   3450
         Width           =   650
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
         Left            =   120
         Picture         =   "PLAINSP07.frx":0BF6
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Plan de Otro Item de Stock"
         Top             =   2760
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
         Left            =   120
         Picture         =   "PLAINSP07.frx":0D40
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Abandona Ficha Plan de Inspección"
         Top             =   495
         Width           =   650
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
         Height          =   690
         Left            =   120
         Picture         =   "PLAINSP07.frx":0E8A
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Ayuda en Línea"
         Top             =   1185
         Width           =   650
      End
      Begin VB.Label COMMLAB 
         Height          =   225
         Left            =   120
         TabIndex        =   51
         Top             =   2040
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label DOCUORI 
         Height          =   225
         Left            =   120
         TabIndex        =   50
         Top             =   2355
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "PLAINSP07.frx":1194
         Top             =   8040
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   105
      TabIndex        =   14
      Top             =   6825
      Width           =   10620
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
         Height          =   1005
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   315
         Width           =   10420
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Aprobación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Index           =   0
      Left            =   8085
      TabIndex        =   17
      Top             =   1365
      Width           =   2640
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1400
         Left            =   180
         TabIndex        =   29
         Top             =   360
         Width           =   2380
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   1
            Left            =   0
            TabIndex        =   40
            Top             =   1050
            Width           =   2250
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
            Height          =   280
            Index           =   1
            Left            =   2100
            Picture         =   "PLAINSP07.frx":30B6
            TabIndex        =   39
            Top             =   735
            Width           =   175
         End
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   1
            Left            =   1155
            TabIndex        =   38
            Top             =   735
            Width           =   990
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
            Height          =   280
            Index           =   0
            Left            =   2100
            Picture         =   "PLAINSP07.frx":33C0
            TabIndex        =   32
            Top             =   0
            Width           =   175
         End
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   0
            Left            =   0
            TabIndex        =   31
            Top             =   315
            Width           =   2250
         End
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Index           =   0
            Left            =   1155
            TabIndex        =   30
            Top             =   0
            Width           =   990
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "Preparó:"
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
            Left            =   105
            TabIndex        =   34
            Top             =   45
            Width           =   1005
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "Aprobó:"
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
            Left            =   105
            TabIndex        =   33
            Top             =   780
            Width           =   1005
         End
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Plan de Inspección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   105
      TabIndex        =   19
      Top             =   1365
      Width           =   7890
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "AQL"
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
         Index           =   4
         Left            =   210
         TabIndex        =   25
         Top             =   315
         Width           =   1065
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
            Height          =   300
            Left            =   420
            MaxLength       =   15
            TabIndex        =   26
            Top             =   260
            Width           =   435
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Nivel"
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
         Index           =   1
         Left            =   1365
         TabIndex        =   23
         Top             =   315
         Width           =   1275
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   315
            MaxLength       =   15
            TabIndex        =   27
            Top             =   260
            Width           =   540
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "%"
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
            Left            =   945
            TabIndex        =   24
            Top             =   315
            Width           =   225
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Frecuencia de Inspección"
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
         Index           =   2
         Left            =   2730
         TabIndex        =   21
         Top             =   315
         Width           =   4950
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Frecuencia de Inspección"
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
            Index           =   5
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   4950
            Begin VB.TextBox Text6 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   315
               TabIndex        =   36
               Top             =   260
               Width           =   4425
            End
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
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
            Index           =   2
            Left            =   315
            TabIndex        =   22
            Top             =   260
            Width           =   4425
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Puntos de Detención Obligatoria"
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
         Index           =   3
         Left            =   210
         TabIndex        =   20
         Top             =   1050
         Width           =   7470
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   420
            TabIndex        =   37
            Top             =   260
            Width           =   6840
         End
      End
   End
   Begin VB.Frame MARCOFICHA 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "1"
      Height          =   3900
      Index           =   1
      Left            =   -1050
      TabIndex        =   11
      Top             =   3045
      Width           =   11880
      Begin VB.Frame Frame8 
         Caption         =   "Frame8"
         Height          =   3375
         Left            =   8600
         TabIndex        =   42
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   3375
         Left            =   5020
         TabIndex        =   41
         Top             =   315
         Width           =   15
      End
      Begin VB.ListBox List1 
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
         Height          =   2985
         Left            =   1155
         TabIndex        =   35
         Top             =   735
         Width           =   10605
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   1155
         ScaleHeight     =   375
         ScaleWidth      =   10545
         TabIndex        =   12
         Top             =   315
         Width           =   10600
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Dimensiones Según Plano                   Instrumento                                        Observaciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   260
            Left            =   25
            TabIndex        =   13
            Top             =   80
            Width           =   10515
         End
      End
      Begin VB.Line Line1 
         X1              =   1155
         X2              =   11760
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line3 
         X1              =   11760
         X2              =   11760
         Y1              =   315
         Y2              =   3720
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   1110
      Left            =   8190
      ScaleHeight     =   1050
      ScaleWidth      =   60
      TabIndex        =   8
      Top             =   210
      Width           =   120
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   5000
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
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
      Left            =   8520
      TabIndex        =   0
      Top             =   120
      Width           =   2220
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
         Height          =   690
         Left            =   360
         Picture         =   "PLAINSP07.frx":36CA
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   315
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Undo"
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
         Left            =   1320
         Picture         =   "PLAINSP07.frx":39D4
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Restablece los Valores Anteriores"
         Top             =   315
         Width           =   650
      End
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
      Left            =   1575
      TabIndex        =   18
      Top             =   210
      Width           =   2115
   End
   Begin VB.Label Label6 
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
      Height          =   550
      Left            =   1575
      TabIndex        =   10
      Top             =   735
      Width           =   6420
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
      Left            =   4095
      TabIndex        =   7
      Top             =   315
      Width           =   1005
   End
   Begin VB.Label Label39 
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
      Left            =   5250
      TabIndex        =   6
      Top             =   210
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
      TabIndex        =   5
      Top             =   315
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
      Left            =   7245
      TabIndex        =   4
      Top             =   210
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
      TabIndex        =   3
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
      Left            =   375
      TabIndex        =   2
      Top             =   315
      Width           =   1005
   End
End
Attribute VB_Name = "PLAINSP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODIPLA%
Dim GRABATODO%
'
Private Sub Command1_Click()
   Command1.Enabled = False
   GRABATODO% = 1
   Call GRAPLAINSP
   GRABATODO% = 0
   Command1.Enabled = True
End Sub


Private Sub Command2_Click()
   Call GRAPLAINSP
End Sub

Private Sub COMMAND21_CLICK()
   '
   Call GRAPLAINSP
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     On Error Resume Next
     DOCUINGE.Text1.TEXT = COD$
     DOCUMI07.Text1.TEXT = COD$
     On Error GoTo 0
     Label28 = Str$(CODINT%(COD$))
   End If
   '
End Sub

Private Sub Command22_Click()
   '
   Call GRAPLAINSP
   Call SELECHO("MASTER/Seleccione ORIGEN del Plan de Inspección")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     CODIORIG% = CODINT%(COD$)
     If CODIORIG% > 0 Then
       Call COPYPLA(CODIORIG%)
     End If
   End If
   '
End Sub

Private Sub Command3_Click()
    Call GRAPLAINSP
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command37_Click()
   Command37.Enabled = False
   FORDOCIN07.Show 1
   Command37.Enabled = True
End Sub

Private Sub Command39_Click()
   Command39.Enabled = False
   DOCUORI.Caption = ""
   COMMLAB.Caption = "PRINT"
   Command39.Enabled = True
End Sub

Private Sub Command9_Click(Index As Integer)
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text5(Index) = VDEV$
   End If
End Sub

Private Sub COMMLAB_Change()
   If TRIM$(UCase$(COMMLAB.Caption)) = "PRINT" Then
      DOCUOR$ = TRIM$(DOCUORI.Caption)
      Call PRIPLAINSP(DOCUOR$)
      COMMLAB.Caption = ""
      DOCUORI.Caption = ""
   End If
End Sub

Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    Refresh
End Sub

Private Sub Form_Resize()
   DoEvents
   If WindowState = 1 Then
      Call BAJALDISCO
   End If
End Sub

Private Sub Label22_DblClick()
   Call EDIPLAINSP
End Sub

Private Sub Label28_Change()
   '
   CI% = Val(Label28)
   '
   Call BLANFORMU
   '
   Label37 = CODEXT$(CI%)
   Label6 = DESCRIT0$(CI%)
   Label39 = UNIMED$(CI%)
   '
   Call CARPLAINSP
   MODIPLA% = 0
   'Call CARHOJAC
   '
End Sub

Private Sub List1_DblClick()
   Call EDIPLAINSP
End Sub

Private Sub Text1_Change()
   MODIPLA% = 1
End Sub

Private Sub Text2_Change()
   MODIPLA% = 1
End Sub

Private Sub Text3_Change()
   MODIPLA% = 1
End Sub

Private Sub Text4_Change(Index As Integer)
   MODIPLA% = 1
End Sub

'
Private Sub Text4_CLICK(Index As Integer)
   If TRIM$(Text4(Index)) = "" Then
      Text4(Index) = USERNAME$
   End If
End Sub

Private Sub Text5_Change(Index As Integer)
   MODIPLA% = 1
End Sub

Private Sub Text5_DblClick(Index As Integer)
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text5(Index) = VDEV$
   End If
End Sub

Sub CARENCAPLA()
   '
   CI% = Val(Label28)
   '
   Screen.MousePointer = 11
   RFF$ = RECFILT$("CODIPROC", 1, MKI$(CI%))
   If Len(RFF$) >= 4 Then
     RECODIPRO& = CVS(Left$(RFF$, 4))
     If RECODIPRO& > 0 Then
       If RECODIPRO& <= ULTREG&("CODIPROC") Then
         YY$ = REGLEIDO$("CODIPROC", RECODIPRO&)
         Text1 = TRIM$(Str$(CVI(Mid$(YY$, 17, 2))))
         Text3 = TRIM$(Str$(CVI(Mid$(YY$, 19, 2))))
         Text6 = Mid$(YY$, 21, 32)
         Text7 = Mid$(YY$, 53, 48)
         '
         Text4(0) = Mid$(YY$, 101, 24)
         FEX = CVI(Mid$(YY$, 125, 2))
         Text5(0) = FECHATEX$(FEX)
         Text4(1) = Mid$(YY$, 127, 24)
         FEX = CVI(Mid$(YY$, 151, 2))
         Text5(1) = FECHATEX$(FEX)
         '
         Text2 = RTrim$(Mid$(YY$, 161, 352))
       End If
     End If
   End If
   Screen.MousePointer = 1
   '
End Sub
'
Sub GRAENCAPLA()
   '
   CI% = Val(Label28)
   '
   Call BLOQARCH("CODIPROC")
   URECODI& = ULTREG&("CODIPROC")
   RECODIPRO& = 0
   '
   RFF$ = RECFILT$("CODIPROC", 1, MKI$(CI%))
   If Len(RFF$) >= 4 Then
       RECODIPRO& = CVS(Left$(RFF$, 4))
   End If
   If RECODIPRO& < 1 Or RECODIPRO& > URECODI& Then
         RECODIPRO& = 1 + URECODI&
         YY$ = String$(512, 0)
      Else
         YY$ = REGLEIDO$("CODIPROC", RECODIPRO&)
   End If
   '
   Call REPLA(YY$, MKI$(CI%), 1, 2)
   Call REPLA(YY$, MKI$(Val(Text1)), 17, 2)
   Call REPLA(YY$, MKI$(Val(Text3)), 19, 2)
   Call REPLA(YY$, Text6, 21, 32)
   Call REPLA(YY$, Text7, 53, 48)
   '
   Call REPLA(YY$, Text4(0), 101, 24)
   FEXI% = FECHANUM(Text5(0))
   Call REPLA(YY$, MKI$(FEXI%), 125, 2)
   Call REPLA(YY$, Text4(1), 127, 24)
   FEXI% = FECHANUM(Text5(1))
   Call REPLA(YY$, MKI$(FEXI%), 151, 2)
   '
   Call REPLA(YY$, Text2, 161, 362)
   Call GRAREG("CODIPROC", YY$, RECODIPRO&)
   Call CIERRARCH("CODIPROC")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
End Sub

Sub EDIPLAINSP()
   '
   CI% = Val(Label28)
   Call BLANARCH("!PLAINSP")
   For U& = 1 To List1.ListCount
     Z$ = List1.List(U& - 1)
     XX$ = MKI$(CI%) + Space$(128)
     Call REPLA(XX$, Mid$(Z$, 1, 32), 3, 32)
     Call REPLA(XX$, Mid$(Z$, 33, 30), 35, 30)
     Call REPLA(XX$, Mid$(Z$, 63, 64), 65, 64)
     Call REGAPP("!PLAINSP", XX$)
   Next U&
   Call CIERRARCH("!PLAINSP")
   '
20 ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=Carga de Plan de Inspeccion y Ensayo"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(PLAINSP07.Top + MARCOFICHA(1).Top + Picture3.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(PLAINSP07.Left + MARCOFICHA(1).Left + Picture3.Left))
   '
   VTB% = VENTABU%("PLAINSP" + ATRI$)
   '
   VACON% = 1
   If VTB% >= 0 Then
     MODIPLA% = 1
     List1.Clear
     For U& = 1 To ULTREG&("!PLAINSP")
       Y$ = REGLEIDO$("!PLAINSP", U&)
       Z$ = Space$(128)
       Call REPLA(Z$, Mid$(Y$, 3, 32), 1, 32)
       Call REPLA(Z$, Mid$(Y$, 35, 30), 33, 30)
       Call REPLA(Z$, Mid$(Y$, 65, 64), 63, 64)
       List1.AddItem Z$
     Next U&
   End If
   OPENFORMS = DoEvents
   '
   VACON% = 1
   '
End Sub

Sub CARPLAINSP()
   '
   URN& = ULTREG&("PLAINSP")
   CI% = Val(Label28)
   If CI% < 1 Then Exit Sub
   '
   Call CARENCAPLA
   '
   List1.Clear
   Screen.MousePointer = 11
   RNC$ = RECFILT$("PLAINSP", 1, MKI$(CI%))
   '
   For U& = 1 To Len(RNC$) Step 4
     RENOTA& = CVS(Mid$(RNC$, U&, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         Y$ = REGLEIDO$("PLAINSP", RENOTA&)
         Z$ = Space$(128)
         Call REPLA(Z$, Mid$(Y$, 3, 32), 1, 32)
         Call REPLA(Z$, Mid$(Y$, 35, 30), 33, 30)
         Call REPLA(Z$, Mid$(Y$, 65, 64), 63, 64)
         List1.AddItem Z$
       End If
     End If
   Next U&
   MODIPLA% = 0
   OPENFORMS = DoEvents
   Screen.MousePointer = 1
   '
End Sub
'

Sub GRAPLAINSP()
   '
   URN& = ULTREG&("PLAINSP")
   CI% = Val(Label28)
   If CI% < 1 Then Exit Sub
   If MODIPLA% <> 1 Then Exit Sub
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se ha Modificado el Plan de Inspeccion del Artículo\'" + TRIM$(Label37) + " - " + TRIM$(Label6) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       Call CARPLAINSP
       MODIPLA% = 0
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   Call BLOQARCH("PLAINSP")
   URN& = ULTREG&("PLAINSP")
   Call GRAENCAPLA
   RNC$ = RECFILT$("PLAINSP", 1, MKI$(CI%))
   KK& = (-3)
   '
   For LL% = 1 To List1.ListCount
      Z$ = List1.List(LL% - 1)
      XX$ = MKI$(CI%) + Space$(128)
      Call REPLA(XX$, Mid$(Z$, 1, 32), 3, 32)
      Call REPLA(XX$, Mid$(Z$, 33, 30), 35, 30)
      Call REPLA(XX$, Mid$(Z$, 63, 64), 65, 64)
      '
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG("PLAINSP", XX$, RENOTA&)
   Next LL%
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("PLAINSP", MKI$(0) + Space$(126), RENOTA&)
   Wend
   '
   Call CIERRARCH("PLAINSP")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   MODIPLA% = 0
   Screen.MousePointer = 1
   '
End Sub

Sub BLANFORMU()
  '
  Text1 = ""
  Text2 = ""
  Text3 = ""
  Text4(0) = "": Text4(1) = ""
  Text5(0) = "": Text5(1) = ""
  Text6 = ""
  Text7 = ""
  List1.Clear
  Label6 = ""
  Label37 = ""
  Label39 = ""
  '
End Sub

Sub PRIPLAINSP(DOCUORIGEN$)
  '
  FORIPRE$ = TRIM$(CONTROL$("FORGRAFI", "DOCUGRA3"))
  
  If FORIPRE$ = "" Then
       PRINTFORM
     Else
       Screen.MousePointer = 11
       FEX = HOY(HOS$)
       FECHDOC$ = HOS$
       NUMEDOC$ = ""
       TIPODOC$ = "Plan de Inspección " + TRIM$(Label37)
       CI1% = XVALO(Label28)
       LETR$ = REVICOD$(CI1%)
       '
       CODPARAM$(1) = "COD-CONJ"
       CODPARAM$(2) = "DES-CONJ"
       CODPARAM$(3) = "" ' "DOCUORIG"
       CODPARAM$(4) = "N-PLANO"
       CODPARAM$(5) = "R-PLANO"
       CODPARAM$(6) = "F-PLANO"
       CODPARAM$(7) = "Z-PLANO"
       CODPARAM$(8) = "FECH-VEN"
       CODPARAM$(9) = "VENDEDOR"
       CODPARAM$(10) = "COND-PAG"
       CODPARAM$(11) = "REF-PLCO"
       CAPARDOC% = 11
       '
       DOCPARAM$(1) = TRIM$(Label37)
       DOCPARAM$(2) = TRIM$(Label6)
       DOCPARAM$(3) = "" 'TRIM$(DOCUORIGEN)
       DOCPARAM$(4) = TRIM$(Text1)
       DOCPARAM$(5) = TRIM$(Text3)
       DOCPARAM$(6) = TRIM$(Text6)
       DOCPARAM$(7) = TRIM$(Text7)
       DOCPARAM$(8) = TRIM$(Text5(1))
       DOCPARAM$(9) = TRIM$(Text4(0))
       DOCPARAM$(10) = TRIM$(Text4(1))
       DOCPARAM$(11) = TRIM$(LETR$)
       '
       Call FRATEXTO(Text2.TEXT, 72)
       If CARETEX% > 6 Then CARETEX% = 6
       For KKL& = 1 To CARETEX%
         Call GRAREG("!OBSERVOF", RETEX$(KKL&), KKL&)
       Next KKL&
       KKL& = CARETEX%
       Call SETULTREG("!OBSERVOF", KKL&)
       '
       CODTABU1$(1) = "N-HINSP"
       CODTABU1$(2) = "N-HPROF"
       CODTABU1$(3) = "MEDIDA-3"
       CACOLTAB1% = 3
       '
       CAITAB1% = 0
       For U& = 1 To List1.ListCount
         Z$ = List1.List(U& - 1)
         CAITAB1% = CAITAB1% + 1
         TETABU1$(1, CAITAB1%) = Left$(Z$, 32)
         TETABU1$(2, CAITAB1%) = Mid$(Z$, 33, 30)
         TETABU1$(3, CAITAB1%) = Mid$(Z$, 63, 30)
       Next U&
       '
       Call PRINTDOC("@" + FORIPRE$)
       Screen.MousePointer = 1
       '
  End If
  '
End Sub

Private Sub Text6_Change()
   MODIPLA% = 1
End Sub

Private Sub Text7_Change()
   MODIPLA% = 1
End Sub

Sub COPYPLA(CODORI%)
   '
   URN& = ULTREG&("PLAINSP")
   If CODORI% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   RFF$ = RECFILT$("CODIPROC", 1, MKI$(CODORI%))
   If Len(RFF$) >= 4 Then
     RECODIPRO& = CVS(Left$(RFF$, 4))
     If RECODIPRO& > 0 Then
       If RECODIPRO& <= ULTREG&("CODIPROC") Then
         YY$ = REGLEIDO$("CODIPROC", RECODIPRO&)
         Text1 = TRIM$(Str$(CVI(Mid$(YY$, 17, 2))))
         Text3 = TRIM$(Str$(CVI(Mid$(YY$, 19, 2))))
         Text6 = Mid$(YY$, 21, 32)
         Text7 = Mid$(YY$, 53, 48)
         '
         Text4(0) = Mid$(YY$, 101, 24)
         FEX = CVI(Mid$(YY$, 125, 2))
         Text5(0) = FECHATEX$(FEX)
         Text4(1) = Mid$(YY$, 127, 24)
         FEX = CVI(Mid$(YY$, 151, 2))
         Text5(1) = FECHATEX$(FEX)
         '
         Text2 = RTrim$(Mid$(YY$, 161, 352))
       End If
     End If
   End If
   '
   List1.Clear
   RNC$ = RECFILT$("PLAINSP", 1, MKI$(CODORI%))
   '
   For U& = 1 To Len(RNC$) Step 4
     RENOTA& = CVS(Mid$(RNC$, U&, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         Y$ = REGLEIDO$("PLAINSP", RENOTA&)
         Z$ = Space$(128)
         Call REPLA(Z$, Mid$(Y$, 3, 32), 1, 32)
         Call REPLA(Z$, Mid$(Y$, 35, 30), 33, 30)
         Call REPLA(Z$, Mid$(Y$, 65, 64), 63, 64)
         List1.AddItem Z$
       End If
     End If
   Next U&
   MODIPLA% = 1
   OPENFORMS = DoEvents
   Screen.MousePointer = 1
   '
End Sub

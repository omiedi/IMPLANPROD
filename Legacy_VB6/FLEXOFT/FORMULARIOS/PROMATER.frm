VERSION 5.00
Begin VB.Form PROMATER 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MRP-II - Programación de Materiales"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11880
   Icon            =   "PROMATER.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11880
   Begin VB.Frame Frame6 
      Caption         =   "Cálculo de Necesidades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1640
      Left            =   3570
      TabIndex        =   10
      Top             =   105
      Width           =   3585
      Begin VB.CommandButton Command10 
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
         Left            =   2730
         Picture         =   "PROMATER.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Configuración - Valores por Omisión"
         Top             =   420
         Width           =   660
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   210
         ScaleHeight     =   585
         ScaleWidth      =   2055
         TabIndex        =   11
         Top             =   420
         Width           =   2115
         Begin VB.CommandButton Command7 
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
            Picture         =   "PROMATER.frx":074C
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Cálculo de Requerimientos de Materiales y Componentes"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Re-Cálculo de Necesidades"
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
            TabIndex        =   13
            Top             =   105
            Width           =   1380
         End
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1080
         TabIndex        =   55
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label Label95 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   2310
         TabIndex        =   49
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label Label121 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "hs"
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
         Left            =   3045
         TabIndex        =   48
         Top             =   1275
         Width           =   330
      End
      Begin VB.Label Label122 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2175
         TabIndex        =   47
         Top             =   1155
         Width           =   120
      End
      Begin VB.Label Label123 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ult.Cálc:"
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
         Left            =   0
         TabIndex        =   46
         Top             =   1275
         Width           =   960
      End
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   3570
      TabIndex        =   63
      Top             =   0
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   330
      Left            =   315
      TabIndex        =   62
      Top             =   105
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Frame Frame7 
      Caption         =   "Programa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1640
      Left            =   2100
      TabIndex        =   50
      Top             =   105
      Width           =   1380
      Begin VB.CommandButton Command25 
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
         Left            =   630
         Picture         =   "PROMATER.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Listado General del Programa de Entregas"
         Top             =   880
         Width           =   600
      End
      Begin VB.CommandButton Command6 
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
         Left            =   210
         Picture         =   "PROMATER.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Consulta Programa de Entregas por Artículo (Pantalla)"
         Top             =   420
         Width           =   600
      End
      Begin VB.Label Label9 
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   435
         Left            =   210
         TabIndex        =   61
         ToolTipText     =   "Doble-Click para Cambiar Modalidad de Programación"
         Top             =   1050
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label8 
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   945
         TabIndex        =   60
         ToolTipText     =   "Doble-Click para Cambiar Modalidad de Programación"
         Top             =   350
         Width           =   380
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   10815
      ScaleHeight     =   75
      ScaleWidth      =   915
      TabIndex        =   44
      Top             =   7350
      Width           =   975
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   45
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command5 
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
      Left            =   11025
      Picture         =   "PROMATER.frx":106A
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Inicia Proceso de Cálculo"
      Top             =   1050
      Width           =   750
   End
   Begin VB.Frame Frame3 
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1640
      Left            =   7245
      TabIndex        =   4
      Top             =   105
      Width           =   3585
      Begin VB.CommandButton Command11 
         Caption         =   "AE's"
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
         Left            =   2835
         TabIndex        =   59
         Top             =   1155
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   330
         Left            =   630
         TabIndex        =   56
         Top             =   1260
         Width           =   2745
         Begin VB.OptionButton Option2 
            Caption         =   "Impresión"
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
            Left            =   1260
            TabIndex        =   58
            Top             =   0
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Pantalla"
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
            Left            =   0
            TabIndex        =   57
            Top             =   0
            Value           =   -1  'True
            Width           =   1170
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   210
         ScaleHeight     =   585
         ScaleWidth      =   3105
         TabIndex        =   5
         Top             =   420
         Width           =   3165
         Begin VB.CommandButton Command9 
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
            Left            =   1995
            Picture         =   "PROMATER.frx":1374
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Necesidades de Compras"
            Top             =   0
            Width           =   600
         End
         Begin VB.CommandButton Command8 
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
            Left            =   2520
            Picture         =   "PROMATER.frx":17B6
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Necesidades de Fabricación"
            Top             =   0
            Width           =   600
         End
         Begin VB.CommandButton Command15 
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
            Picture         =   "PROMATER.frx":1AC0
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Análisis de Cobertura"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Necesidades Programadas"
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
            Top             =   105
            Width           =   1170
         End
      End
   End
   Begin VB.CommandButton Command1 
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
      Left            =   11025
      Picture         =   "PROMATER.frx":1DCA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   210
      Width           =   750
   End
   Begin VB.Frame Frame2 
      Caption         =   "Actualización de Programa de Entregas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   105
      TabIndex        =   1
      ToolTipText     =   "Doble-Click para Cambiar Modalidad de Programación"
      Top             =   7035
      Width           =   10590
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
         Height          =   330
         Index           =   5
         Left            =   9135
         TabIndex        =   33
         Top             =   630
         Width           =   700
      End
      Begin VB.CommandButton Command4 
         Height          =   670
         Left            =   9975
         Picture         =   "PROMATER.frx":1F14
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   315
         Width           =   500
      End
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   225
         Left            =   5565
         TabIndex        =   35
         Top             =   420
         Width           =   4425
         Begin VB.Label INISEM 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "25-02"
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
            Index           =   11
            Left            =   3600
            TabIndex        =   41
            Top             =   0
            Width           =   645
         End
         Begin VB.Label INISEM 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "25-02"
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
            Index           =   10
            Left            =   2880
            TabIndex        =   40
            Top             =   0
            Width           =   645
         End
         Begin VB.Label INISEM 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "25-02"
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
            Index           =   9
            Left            =   2160
            TabIndex        =   39
            Top             =   0
            Width           =   645
         End
         Begin VB.Label INISEM 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "25-02"
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
            Index           =   8
            Left            =   1420
            TabIndex        =   38
            Top             =   0
            Width           =   645
         End
         Begin VB.Label INISEM 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "25-02"
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
            Index           =   7
            Left            =   705
            TabIndex        =   37
            Top             =   0
            Width           =   645
         End
         Begin VB.Label INISEM 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "25-02"
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
            Index           =   6
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   645
         End
      End
      Begin VB.CommandButton BOTSEL 
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
         Left            =   1890
         Picture         =   "PROMATER.frx":221E
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   630
         Width           =   175
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
         Height          =   330
         Index           =   0
         Left            =   5525
         TabIndex        =   28
         Top             =   630
         Width           =   700
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
         Height          =   330
         Index           =   4
         Left            =   8410
         TabIndex        =   32
         Top             =   630
         Width           =   700
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
         Height          =   330
         Index           =   3
         Left            =   7685
         TabIndex        =   31
         Top             =   630
         Width           =   700
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
         Height          =   330
         Index           =   2
         Left            =   6960
         TabIndex        =   30
         Top             =   630
         Width           =   700
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
         Height          =   330
         Index           =   1
         Left            =   6235
         TabIndex        =   29
         Top             =   630
         Width           =   700
      End
      Begin VB.TextBox Text1 
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
         Left            =   105
         TabIndex        =   26
         Top             =   630
         Width           =   1800
      End
      Begin VB.Label Label5 
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
         Height          =   560
         Left            =   2150
         TabIndex        =   42
         Top             =   400
         Width           =   3270
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Programa de Entregas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5160
      Left            =   105
      TabIndex        =   0
      ToolTipText     =   "Doble-Click para Cambiar Modalidad de Programación"
      Top             =   1785
      Width           =   11670
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
         Height          =   3885
         Left            =   105
         TabIndex        =   2
         Top             =   1140
         Width           =   11460
      End
      Begin VB.PictureBox Picture1 
         Height          =   850
         Left            =   105
         ScaleHeight     =   795
         ScaleWidth      =   11400
         TabIndex        =   14
         Top             =   315
         Width           =   11460
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   225
            Left            =   5490
            TabIndex        =   19
            Top             =   550
            Width           =   4425
            Begin VB.Label INISEM 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "d/m-2"
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
               Left            =   705
               TabIndex        =   25
               Top             =   0
               Width           =   645
            End
            Begin VB.Label INISEM 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "d/m-1"
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
               TabIndex        =   24
               Top             =   0
               Width           =   645
            End
            Begin VB.Label INISEM 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "d/m-3"
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
               Left            =   1420
               TabIndex        =   23
               Top             =   0
               Width           =   645
            End
            Begin VB.Label INISEM 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "d/m-4"
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
               Left            =   2160
               TabIndex        =   22
               Top             =   0
               Width           =   645
            End
            Begin VB.Label INISEM 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "d/m-5"
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
               Index           =   4
               Left            =   2880
               TabIndex        =   21
               Top             =   0
               Width           =   645
            End
            Begin VB.Label INISEM 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "d/m-6"
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
               Index           =   5
               Left            =   3600
               TabIndex        =   20
               Top             =   0
               Width           =   645
            End
         End
         Begin VB.CommandButton Command3 
            Height          =   330
            Left            =   9135
            Picture         =   "PROMATER.frx":2528
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton Command2 
            Height          =   330
            Left            =   5670
            Picture         =   "PROMATER.frx":2832
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   0
            Width           =   540
         End
         Begin VB.Line Line1 
            X1              =   6090
            X2              =   9135
            Y1              =   315
            Y2              =   315
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   16
            Top             =   0
            Width           =   2955
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo           Descripción              Ant.  S-1   S-2   S-3   S-4   S-5   S-6  Post Total"
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
            Left            =   0
            TabIndex        =   15
            Top             =   315
            Width           =   11460
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   840
      Picture         =   "PROMATER.frx":2B3C
      Top             =   315
      Width           =   480
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "MRP-II"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   750
      Left            =   210
      TabIndex        =   53
      Top             =   945
      Width           =   1800
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      FillStyle       =   7  'Diagonal Cross
      Height          =   960
      Left            =   210
      Shape           =   2  'Oval
      Top             =   480
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10320
      Picture         =   "PROMATER.frx":2E46
      Top             =   7680
      Width           =   2010
   End
End
Attribute VB_Name = "PROMATER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PROGENT(0 To 32767) As String * 212
Dim CODIPRO%(32767)
Dim CACODIP%
Dim LINVE$(0 To 32767)
Dim SEMI%(-1 To 7)
Dim ISEMI%(-1 To 7)
Dim CAPRO(-1 To 7)

Private Sub BOTSEL_Click(Index As Integer)
   Call SELECHO("MASTER")
   Text1 = VALACT1$("MASTER")
End Sub

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Call FINAL("")
   '
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   OPPROMRP.Show 1
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   If UCase$(Command11.Caption) = "MAIL" Then
       Call CONECRUN("PLANIPRO/TRANECOMAI", "Transferir Necesidades de Compras por Mail")
       Call FINAL("")
     Else
       Call CONECRUN("AENTRE02", "Autorizaciones de Entregas")
   End If
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    Command12.Enabled = False
    Dim COSU%(62), SUCU$(62)
    NUSU% = 1
    For I& = 1 To ULTREG&("SUCURSAL")
      X$ = REGLEIDO$("SUCURSAL", I&)
      If I& <= 62 Then
        If Asc(Left$(X$, 1)) > 32 Then
          NUSU% = I&
          COSU%(I&) = Asc(Mid$(X$, 64, 1))
          SUCU$(I&) = Left$(X$, 20)
        End If
      End If
    Next I&
    Call CIERRARCH("SUCURSAL")
    '
    Call COPYSTRU("SALDEPOS", "SALDEPO")
    '
    FIN& = NUMAS% + 1
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      SSQ$ = String$(255, 0) + Chr$(1)
      Call GRAREG("SALDEPO", SSQ$, I&)
    Next I&
    Call SETULTREG("SALDEPO", FIN&)
    '
    FIN& = ULTREG&("MOVISTO")
    NMS% = NUMAS%
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      MVS$ = REGLEIDO$("MOVISTO", I&)
      CI% = CVI(Mid$(MVS$, 3, 2))
      DEP% = Asc(Mid$(MVS$, 83, 1))
105   SUCX% = 1
      For U% = 1 To NUSU%
        If COSU%(U%) = DEP% Then SUCX% = U%: GoTo 110
      Next U%
      '
110   CANTI = CVD(Mid$(MVS$, 88, 8)) - CVD(Mid$(MVS$, 96, 8))
      '
      If CI% > 0 Then
        If CI% <= NMS% Then
          RESAL& = CI% + 1
          SSQ$ = REGLEIDO$("SALDEPO", RESAL&)
          CANUE = CANTI + (Int(10 * CVS(Mid$(SSQ$, 4 * SUCX% - 3, 4)) + 0.5)) / 10
          Mid$(SSQ$, 4 * SUCX% - 3, 4) = MKS$(CANUE)
          Call GRAREG("SALDEPO", SSQ$, RESAL&)
        End If
      End If
119 Next I&
    '
    Call CIERRARCH("MOVISTO")
    Call CIERRARCH("SALDEPO")
    '
    FIN& = NUMAS%
    For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      XX$ = Left$(REGLEIDO$("SALDEPOS", I& + 1), 248)
      YY$ = Left$(REGLEIDO$("SALDEPO", I& + 1), 248)
      If XX$ <> YY$ Then
        CAX = 0: CAY = 0
        For JI% = 1 To 248 Step 4
          CAX = CAX + CVS(Mid$(XX$, JI%, 4))
          CAY = CAY + CVS(Mid$(YY$, JI%, 4))
        Next JI%
          If (Abs(CAX - CAY)) > 0.05 Then
            CIXI% = I&
            MsgBox I& & "   " & CODEXT$(CIXI%)
            GoTo 129
          End If
        'Next JI%
      End If
129 Next I&
    Call CIERRARCH("*.*")
    MsgBox "Control Completo"
    '
End Sub

Private Sub Command15_Click()
    Dim CADISP(32000) As Single
    '
    Command15.Enabled = False
    OP% = ALTERNA%("Cobertura Completa\Cobertura Inmediata")
    Select Case OP%
      Case 1
        If Option1.Value = True Then
            Call CONECRUN("*COBERFA", "")
          Else
            OPPP% = COMALTER%("Elija Formato:\\Detalle\Resumen")
            If OPPP% = 2 Then
                Call RESUCOBE
                GoTo 99
              ElseIf OPPP% = 1 Then
                GoTo 15
              Else
                GoTo 99
            End If
            '
15          Call FILESORT("COBERFA", "COBERFB", 2, 2)
            'Call CONECRUN("*SOCOB01", "Listado de Cobertura de Stocks")
            Call CONECRUN("*LICOB01", "Listado de Cobertura de Stocks")
        End If
      Case 2
        ' SELECCIONA COBERTURAS INSUFICIENTE HASTA FLIM%
        Call COPYSTRU("NEGASTO", "NEGASTX")
        Call SELECHO("FECHA")
        VDEVU$ = VALACT1$("FECHA")
        FLIM% = FECHANUM(VDEVU$)
        If FLIM% > HOY(HOS$) Then
          FIN& = ULTREG&("COBERFA")
          For U& = 1 To FIN&
            XX$ = REGLEIDO$("COBERFA", U&)
            If CVI(Left$(XX$, 2)) <= FLIM% Then
              CI% = CVI(Mid$(XX$, 3, 2))
              CADISP(CI%) = CADISP(CI%) + CVS(Mid$(XX$, 53, 4)) - CVS(Mid$(XX$, 57, 4))
            End If
          Next U&
          '
          JJ& = 0
          For U& = 1 To NUMAS%
            If CADISP(U&) < (-0.05) Then
              CI% = U&
              YY$ = REGBLAN$("NEGASTX")
              Call REPLA(YY$, MKI$(CI%), 1, 2)
              Call REPLA(YY$, MKS$(CADISP(CI%)), 3, 4)
              JJ& = JJ& + 1
              Call GRAREG("NEGASTX", YY$, JJ&)
            End If
          Next U&
          '
          Call SETULTREG("NEGASTX", JJ&)
          Call CIERRARCH("NEGASTX")
          '
          Call HEADERS("NEGASTX", "")
          FEX = FLIM%
          Call HEADERS("NEGASTX", "Hasta: " + FECHATEX$(FEX))
          Call FINAL("*NEGASTX")
          '
        End If
    End Select
99  Command15.Enabled = True
End Sub
'
Sub RESUCOBE()
    '
    Dim CANCOL#()
    NMS% = NUMAS%
    ReDim CANCOL#(NMS%, 12)
    FIN& = ULTREG("COBERFA")
    '
    For III& = 1 To FIN&
        '
        Call TRACE(20, III&, FIN&)
        X$ = REGLEIDO$("COBERFA", III&)
        CI% = CVI(Mid$(X$, 3, 2))
        FF% = CVI(Mid$(X$, 1, 2))
        REFE$ = UCase$(TRIM$(Mid$(X$, 21, 30)))
        CANTI# = CVS(Mid$(X$, 53, 4))
        CANTE# = CVS(Mid$(X$, 57, 4))
        '
        If Left$(REFE$, 8) = "STOCK AL" Then
            COLUI% = 1
          ElseIf Left$(REFE$, 12) = "STOCK MINIMO" Then
            COLUI% = 7
          ElseIf Left$(REFE$, 3) = "OC." Then
            COLUI% = 2
          ElseIf Left$(REFE$, 12) = "O/F A EMITIR" Then
            COLUI% = 5
          ElseIf Left$(REFE$, 7) = "RESERVA" Then
            COLUI% = 10
          ElseIf Left$(REFE$, 6) = "ATRASO" Then
            COLUI% = 9
          ElseIf Left$(REFE$, 8) = "PROGRAMA" Then
            COLUI% = 8
          ElseIf Left$(REFE$, 20) = "TENTATIVO DE COMPRAS" Then
            COLUI% = 3
          Else
            COLUI% = 4
        End If
        '
        If CANTI# < 0.05 Then
          If CANTE# < 0.05 Then
            GoTo 19
          End If
        End If
        '
        If CI% > 0 Then
          If CI% <= NMS% Then
            CANCOL#(CI%, 0) = 1
            CANCOL#(CI%, COLUI%) = CANCOL#(CI%, COLUI%) + CANTI# + CANTE#
            If COLUI% < 6 Then
                CANCOL#(CI%, 6) = CANCOL#(CI%, 6) + CANTI# + CANTE#
              Else
                CANCOL#(CI%, 11) = CANCOL#(CI%, 11) + CANTI# + CANTE#
            End If
            If CANCOL#(CI%, 6) > CANCOL#(CI%, 11) Then
                 CANCOL#(CI%, 12) = CANCOL#(CI%, 6) - CANCOL#(CI%, 11)
               Else
                 CANCOL#(CI%, 12) = 0
            End If
          End If
        End If
        '
19  Next III&
    '
    KK& = 0
    For KIII% = 1 To NUMAS%
       If CANCOL#(KIII%, 0) > 0.5 Then
          XXX$ = MKI$(KIII%)
          For J% = 1 To 12
             XXX$ = XXX$ + MKD$(CANCOL#(KIII%, J%))
          Next J%
          XXX$ = Left$(XXX$ + String$(128, 0), 128)
          KK& = KK& + 1
          Call GRAREG("RESUCOB", XXX$, KK&)
       End If
    Next KIII%
    Call SETULTREG("RESUCOB", KK&)
    Call CIERRARCH("RESUCOB")
    '
    Call FILESORT("RESUCOB", "", 1, 1)
    '
    Call CONECRUN("*RESUCOB", "")
    '
End Sub
'
Private Sub Command2_Click()
   MESPAN$ = MESANO$(Label2, DES%, HAS%)
   FEX = DES%
   FEPANUE% = DIANTE(FEX)
   If FEPANUE% < COMESACT% Then Exit Sub
   Label2 = METEXTO$(FEPANUE%)
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    '\\\OT-06-0236-WAR-08/06/06///
    '
    OPX% = COMALTER%("Formato de Salida:\\Tabla\Listado")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    If OPX% = 1 Then
      Call LISTABPLAN
      GoTo 99
    End If
    '
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If TRIM$(PERIO$) = "" Then
       APROGRA$ = "*LPROGEN"
       If MODOPLAN% = 1 Then APROGRA$ = "*LPROGFA"
       Call CONECRUN(APROGRA$, "")
       GoTo 99
    End If
    '
    APROGRA$ = "PROGENT": TILIS$ = "Período de Entregas: "
    If MODOPLAN% = 1 Then
      APROGRA$ = "PROGFAB"
      TILIS$ = "Período de Fabricación: "
    End If
    Call COPYSTRU("PROGENT", "PROGRAM")
    '
    XBLA$ = REGBLAN$("PROGRAM")
    JJ& = 0
    For XJ& = 1 To ULTREG(APROGRA$)
      Y$ = XBLA$
      X$ = REGLEIDO$(APROGRA$, XJ&)
      FEP% = CVI(Mid$(X$, 3, 2))
      If FEP% >= DES% Then
        If FEP% <= HAS% Then
          Y$ = X$
          JJ& = JJ& + 1
          Call GRAREG("PROGRAM", Y$, JJ&)
        End If
      End If
    Next XJ&
    Call SETULTREG("PROGRAM", JJ&)
    '
    Call HEADERS("PROGRAM", "")
    Call HEADERS("PROGRAM", TILIS$ + PERIO$)
    '
    Call CONECRUN("*LPROGRAM", "")
    '\\\OT-06-0236-WAR-FIN///
99  Command25.Enabled = True
End Sub

Sub LISTABPLAN()
   Dim INISE(16)
   CMACT = COMESACT%
   SMXII% = SEMANU%(CMACT, DELX, ALX)
   INISE(1) = DELX
   CMACT = DELX
   JII% = 1
   '
10 If SEMANU%(CMACT, DELX, ALX) = SMXII% Then
      CMACT = DIPOST(CMACT)
      GoTo 10
   End If
   SMXII% = SEMANU%(CMACT, DELX, ALX)
   If DELX = INISE(JII%) Then GoTo 10
   JII% = JII% + 1
   INISE(JII%) = DELX
   If JII% < 15 Then GoTo 10
   '
   TTXX$ = "Codigo            Descripción" + Space$(30) + "Vencido"
   For JII% = 1 To 14
     TTXX$ = TTXX$ + "  " + FECHATEX$(INISE(JII%))
   Next JII%
   TTXX$ = TTXX$ + "     Total"
   TTXX$ = Left$(TTXX$ + Space$(256), 256)
   Call SAVEFILE(LUCOM$ + "TAPROGEN.HDP", TTXX$)
   '
   JJ& = 0
   For U& = 1 To CACODIP%
     TTXX$ = REGBLAN$("TAPROGEN")
     Call REPLA(TTXX$, MKI$(CODIPRO%(U&)), 1, 2)
     Call REPLA(TTXX$, Left$(PROGENT$(U&), 60), 3, 60)
     SUMA = 0: For UI% = 1 To 15: SUMA = SUMA + CVS(Mid$(PROGENT$(U&), 4 * UI% - 3, 4)): Next UI%
     Call REPLA(TTXX$, MKS$(SUMA), 63, 4)
     JJ& = JJ& + 1
     Call GRAREG("TAPROGEN", TTXX$, JJ&)
   Next U&
   Call SETULTREG("TAPROGEN", JJ&)
   Call CIERRARCH("TAPROGEN")
   '
   Call FINAL("*TAPROGE")
   '
End Sub
'
Private Sub Command3_Click()
   MESPAN$ = MESANO$(Label2, DES%, HAS%)
   FEX = HAS%
   FEPANUE% = DIPOST(FEX)
   If DIENTRE(COMESACT%, FEPANUE%) > 360 Then Exit Sub
   Label2 = METEXTO$(FEPANUE%)
End Sub

Private Sub Command4_Click()
   If CODINT%(Text1.TEXT) > 0 Then
     CI% = CODINT(Text1.TEXT)
     Label5 = DESCRIT0$(CI%)
     Call GRAPROCAR
   End If
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   If OPCIPLAN% = 1 Then
     MODOP$ = "ENTREGA": If MODOPLAN% = 1 Then MODOP$ = "FABRICA"
     Call GRACONTROL("PROMRP", "MODOPLAN", MODOP$)
   End If
   Call GRAPROGENT
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Static TANTU
    If Timer < TANTU + 10 Then Exit Sub
    TANTU = Timer
    Command6.Enabled = False
    If EXISTE%("PLANIPRO.EXE") > 0 Then
        Call CONECRUN("PLANIPRO", "Planilla de Producción de Máquinas")
        Call FINAL("")
      Else
        Call FINAL("*CPROGEN")
    End If
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   FORCANEC.Show 1
   Refresh
   DoEvents
   If App.EXEName <> "FLEXO00" Then
     Call CONECRUN(App.EXEName, "")
     Call FINAL("")
   End If
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    If Option1.Value = True Then
          Call FINAL("*CPLAFAB")
        Else
          Call FINAL("*LPLAFAB")
    End If
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    AXXX$ = "": ARPLA$ = "PLACOMP"
    If InStr(EPAC$, "FERVI") > 0 Then
       AXXX$ = "Q"
       ARPLA$ = "PLACOMQ"
    End If
    MODOLIS% = 0
    If EXISTE%(LUDAT$ + "PLACOMC.RFS") > 0 Then MODOLIS% = 1
    If Option1.Value = True Then
        Call FINAL("*CPLACOM" + AXXX$)
      Else
        OPX1% = COMALTER%("Elija Secuencia / Agrupamiento\\Por Proveedor\Por Artículo")
        If OPX1% = 1 Then
            Call FILESORT(ARPLA$, "", 8, 8)
          ElseIf OPX1% = 2 Then
            Call FILESORT(ARPLA$, "", 1, 3)
        End If
        '
        TNT% = 0
        If AXXX$ = "" Then
10        TMT% = Asc(Left$(REGSEL$("TATIMAT"), 1))
          If TMT% > 0 Then
            Screen.MousePointer = 11
            Call COPYSTRU("PLACOMP", "PLACOMF")
            JJ& = 0: FIN& = ULTREG&("PLACOMP")
            For U& = 1 To FIN&
              Call TRACE(20, U&, FIN&)
              X$ = REGLEIDO$("PLACOMP", U&)
              CIXI% = CVI(Left$(X$, 2))
              If TIMATE%(CIXI%) = TMT% Then
                JJ& = JJ& + 1
                Call GRAREG("PLACOMF", X$, JJ&)
              End If
            Next U&
            Screen.MousePointer = 1
            If JJ& < 1 Then
              Call MENSERR(24, "Sin Necesidades\para ese Tipo de Material.")
              GoTo 10
            End If
            Call SETULTREG("PLACOMF", JJ&)
            ARENTRA$ = "PLACOMF"
            If MODOLIS% = 1 Then GoTo 60
            '
            Call HEADERS("PLACOMF", "")
            Call HEADERS("PLACOMF", "Tipo de Material: " + ECOARCH$("TATIMAT", Chr$(TMT%)))
            Call FINAL("*LPLACOF")
            GoTo 99
          End If
        End If
        ARENTRA$ = "PLACOMP"
        If MODOLIS% = 1 Then GoTo 60
        Call FINAL("*LPLACOM" + AXXX$)
        GoTo 99
    End If
    '
60  Dim CANCO(32767), CANRE(32767)
    NMS% = NUMAS%
    Screen.MousePointer = 11
    FIN& = ULTREG&("OCOMDETA")
    JJ& = 0
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XXX$ = REGLEIDO$("OCOMDETA", U&)
      If Asc(Mid$(XXX$, 124, 1)) < 8 Then ' ordenes abiertas Y NO CUMPLIDAS
        CAPEN = CVS(Mid$(XXX$, 55, 4)) - CVS(Mid$(XXX$, 99, 4))
        If CAPEN < 0.05 Then CAPEN = 0
        CIXI% = CVI(Mid$(XXX$, 9, 2))
        If CIXI% > 0 Then
          If CIXI% <= NMS% Then
            CANCO(CIXI%) = CANCO(CIXI%) + CAPEN
          End If
        End If
      End If
    Next U&
    '
    FIN& = ULTREG&("BOLREPEN")
    JJ& = 0
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XXX$ = REGLEIDO$("BOLREPEN", U&)
      CAPEN = CVD(Mid$(XXX$, 89, 8))
      If CAPEN < 0.05 Then CAPEN = 0
      CIXI% = CVI(Mid$(XXX$, 83, 2))
      If CIXI% > 0 Then
        If CIXI% <= NMS% Then
          CANRE(CIXI%) = CANRE(CIXI%) + CAPEN
        End If
      End If
    Next U&
    '
    REBLA$ = REGBLAN$("PLACOMC")
    CIA% = 0
    JJ& = 0
    FIN& = ULTREG&(ARENTRA$)
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$(ARENTRA$, U&)
      CIXI% = CVI(Mid$(XX$, 1, 2))
      YY$ = REBLA$
      Call REPLA(YY$, Left$(XX$, 32), 1, 32)
      Call REPLA(YY$, Mid$(XX$, 33), 45, 32)
      If CIXI% <> CIA% Then
        CIA% = CIXI%
        If InStr(EPAC$, "FERVI") > 0 Then
          Call REPLA(YY$, MKS$(STOCKACT(CIXI%)), 29, 4)
        End If
        Call REPLA(YY$, MKS$(CANCO(CIXI%)), 33, 4)
        Call REPLA(YY$, MKS$(CANRE(CIXI%)), 37, 4)
        Call REPLA(YY$, MKS$(CVS(Mid$(XX$, 29, 4)) + CANCO(CIXI%) + CANRE(CIXI%)), 41, 4)
      End If
      JJ& = JJ& + 1
      Call GRAREG("PLACOMC", YY$, JJ&)
    Next U&
    Call SETULTREG("PLACOMC", JJ&)
    Call CIERRARCH("PLACOMC")
    '
    If OPX1% = 1 Then
         Call FILESORT("PLACOMC", "", 11, 11)
       ElseIf OPX1% = 2 Then
         Call FILESORT("PLACOMC", "", 1, 3)
    End If
    Screen.MousePointer = 1
    Call FINAL("*PLACOMC")
    GoTo 99
    '
99  Command9.Enabled = True
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Refresh
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
   '
   Show
   Refresh
   VERANTER$ = "PROMRP01"
   Call VERJOBID(OKEY%)

   If OKEY% < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
     '
10   Caption = Caption + "  -  " + EPAC$
     Call GRATITFOR(Caption)
     '
   End If
   '
   OPCIPLAN% = 0
   If InStr(UCase$(EPAC$), "DOSIVA") > 0 Then OPCIPLAN% = 1
   Call GRATITFOR(Caption)
   '
   If EXISTE%("AENTRE02.EXE") > 0 Then
     Frame8.Left = 210
     Command11.Visible = True
   End If
   '
   'BAJMAICOM% = 0
   'If InStr(UCase$(EPAC$), "ASCENSORES DEL SUR") > 0 Or InStr(UCase$(EPAC$), "ADELSUR") > 0 Then
    If EXISTE%("PLANIPRO.EXE") > 0 Then
     Frame8.Left = 210
     Command11.Caption = "Mail"
     Command11.ToolTipText = "Transferir Necesidades Compras por Mail"
     Command11.Visible = True
   '  BAJMAICOM% = 1
   End If
   '
   OPENFORMS = DoEvents
   '
   MODOPLAN% = 0
   If OPCIPLAN% = 1 Then
     If CONTROL("PROMRP", "MODOPLAN") = "FABRICA" Then
       MODOPLAN% = 1
     End If
     If MODOPLAN% = 1 Then
         Frame1.ForeColor = RGB(80, 0, 0)
         Frame1.Caption = "Plan Maestro de Fabricación"
         Label8.Visible = False: Label9.Visible = True
         Command6.ToolTipText = "Consulta Plan de Fabricación por Artículo (Pantalla)"
         Command25.ToolTipText = "Listado General del Plan Maestro de Fabricación"
       Else
         Frame1.ForeColor = RGB(0, 0, 170)
         Frame1.Caption = "Programa de Entregas"
         Label9.Visible = False: Label8.Visible = True
         Command6.ToolTipText = "Consulta Programa de Entregas por Artículo (Pantalla)"
         Command25.ToolTipText = "Listado General del Programa de Entregas"
     End If
     Frame2.ForeColor = Frame1.ForeColor
     Frame2.Caption = "Actualización del " + Frame1.Caption
   End If
   '
   FECHOR$ = CONTROL$("CANECOM", "FECHOCAL")
   Label7.Caption = Left$(FECHOR$, 8): Label7.Refresh
   Label95.Caption = Mid$(FECHOR$, 11, 5): Label95.Refresh
   '
   PRISEMIN% = 53
   If Val(CONTROL$("", "PRISEMIN")) = 1 Then PRISEMIN% = 1
   '
   Call CARPROGENT
   '
   FECHAC% = HOY(HO$)
   Label2.Caption = METEXTO$(FECHAC%)
   '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
      Call CIERRARCH("*.*")
      Call BAJALDISCO
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '
   Call CIERRARCH("*.*")
   Call FINAL("")
   '
End Sub

Private Sub Frame1_DblClick()
   If OPCIPLAN% = 1 Then
     If MODOPLAN% = 0 Then
         MODOPLAN% = 1
         Frame1.ForeColor = RGB(80, 0, 0)
         Frame1.Caption = "Plan Maestro de Fabricación"
         Label8.Visible = False: Label9.Visible = True
         Command6.ToolTipText = "Consulta Plan de Fabricación por Artículo (Pantalla)"
         Command25.ToolTipText = "Listado General del Plan Maestro de Fabricación"
       Else
         MODOPLAN% = 0
         Frame1.ForeColor = RGB(0, 0, 170)
         Frame1.Caption = "Programa de Entregas"
         Label9.Visible = False: Label8.Visible = True
         Command6.ToolTipText = "Consulta Programa de Entregas por Artículo (Pantalla)"
         Command25.ToolTipText = "Listado General del Programa de Entregas"
     End If
     Frame2.ForeColor = Frame1.ForeColor
     Frame2.Caption = "Actualización del " + Frame1.Caption
     Call CARPROGENT
     Call MUEPROGENT
   End If
End Sub

Private Sub Frame2_DblClick()
   If OPCIPLAN% = 1 Then
     If MODOPLAN% = 0 Then
         MODOPLAN% = 1
         Frame1.ForeColor = RGB(80, 0, 0)
         Frame1.Caption = "Plan Maestro de Fabricación"
         Label8.Visible = False: Label9.Visible = True
         Command6.ToolTipText = "Consulta Plan de Fabricación por Artículo (Pantalla)"
         Command25.ToolTipText = "Listado General del Plan Maestro de Fabricación"
       Else
         MODOPLAN% = 0
         Frame1.ForeColor = RGB(0, 0, 170)
         Frame1.Caption = "Programa de Entregas"
         Label9.Visible = False: Label8.Visible = True
         Command6.ToolTipText = "Consulta Programa de Entregas por Artículo (Pantalla)"
         Command25.ToolTipText = "Listado General del Programa de Entregas"
     End If
     Frame2.ForeColor = Frame1.ForeColor
     Frame2.Caption = "Actualización del " + Frame1.Caption
     Call CARPROGENT
     Call MUEPROGENT
   End If
End Sub

Private Sub Image1_DblClick()
   ' calculo de capacidad
   '
   Screen.MousePointer = 11
   Dim CANPROG(32767)
   APROGRA$ = "PROGENT"
   If MODOPLAN% = 1 Then APROGRA$ = "PROGFAB"
   '
   REBLA$ = REGBLAN$("CAPANECE")
   JJ& = 0
   FIN& = ULTREG&(APROGRA$)
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     P00$ = REGLEIDO$(APROGRA$, I&)
     P1$ = Left$(P00$, 2)
     CIXI% = CVI(P1$)
     '
     If CIXI% > 0 Then
       If CIXI% <= NUMAS% Then
         If STATUIT%(CI%) >= 0 Then
           CANPROG(CIXI%) = CANPROG(CIXI%) + CVS(Mid$(P00$, 7, 4))
         End If
       End If
     End If
   Next I&
   '
   FIT& = NUMAS%
   For U& = 1 To FIT&
     Call TRACE(20, U&, FIT&)
     CIKI% = U&
     If CANPROG(CIKI%) >= 0.05 Then
        Call EXPLOMUL(CIKI%)
        '
        FIN& = CONTAMUL%
        URE& = ULTREG&("SECOPER")
        SUTIE = 0: SUMOD = 0: SUMAQ = 0
        ICX% = 0
        '
        For UI& = 1 To FIN&
          '
          CIXI% = CIK%(UI&)
          CODEX$ = CODEXT$(CIXI%)
          CACOM = CANT(UI&)
          '
          RFF$ = RECFILT$("SECOPER", 1, MKI$(CIXI%))
          IOP% = 0
          For UJ& = 1 To Len(RFF$) Step 4
            I& = CVS(Mid$(RFF$, UJ&, 4))
            If I& > 0 Then
              If I& <= URE& Then
                SCP$ = REGLEIDO$("SECOPER", I&)
                NOPY% = CVI(Mid$(SCP$, 3, 2))
                DOPE$ = Mid$(SCP$, 5, 48)
                LOTESTA = LOTESTAN(CIXI%): If LOTESTA < 1 Then LOTESTA = 1
                CICLOHE = CVS(Mid$(SCP$, 99, 4)): If CICLOHE < 1 Then CICLOHE = 1
                MINPREP = 60 * CVS(Mid$(SCP$, 59, 4))
                MINFIJO = 60 * CVS(Mid$(SCP$, 63, 4))
                MINVARI = LOTESTA * CVS(Mid$(SCP$, 67, 4))
                MINHERR = LOTESTA * 60 * CVS(Mid$(SCP$, 95, 4)) / CICLOHE
                MINDEMO = (MINFIJO + MINVARI + MINHERR) * CVS(Mid$(SCP$, 75, 4)) / 100
                MINLOTE = MINPREP + MINFIJO + MINVARI + MINHERR + MINDEMO
                HORAPROD = CANPROG(CIKI%) * CACOM * MINLOTE / 60 / LOTESTA
                '
                CAOPS = CVS(Mid$(SCP$, 83, 4))
                '
                If HORAPROD > 0.01 Then
                  EQUIPO$ = Mid$(SCP$, 53, 6) + Mid$(SCP$, 103, 18)
                  For KU% = 1 To 24 Step 6
                    EQUI$ = TRIM$(Mid$(EQUIPO$, KU%, 6))
                    If EQUI$ <> "" Then
                      RECAP$ = REBLA$
                      Call REPLA(RECAP$, EQUI$, 1, 6)
                      Call REPLA(RECAP$, MKI$(CIXI%), 7, 2)
                      Call REPLA(RECAP$, MKI$(NOPY%), 9, 2)
                      Call REPLA(RECAP$, DOPE$, 11, 48)
                      Call REPLA(RECAP$, MKS$(CANPROG(CIKI%) * CACOM), 59, 4)
                      Call REPLA(RECAP$, MKI$(CIKI%), 67, 2)
                      Call REPLA(RECAP$, MKS$(HORAPROD), 69, 4)
                      Call REPLA(RECAP$, MKS$(HORAPROD), 77, 4)
                      JJ& = JJ& + 1
                      Call GRAREG("CAPANECE", RECAP$, JJ&)
                    End If
                  Next KU%
                End If
              End If
            End If
          Next UJ&
        Next UI&
     End If
   Next U&
   '
   Call SETULTREG("CAPANECE", JJ&)
   Call FILESORT("CAPANECE", "", 1, 3)
   Call FINAL("*LICAPAC")
   '
   Screen.MousePointer = 1
   '
End Sub

Private Sub INISEM_Change(Index As Integer)
   If Index < 6 Then
     INISEM(Index + 6) = INISEM(Index)
   End If
End Sub

Private Sub Label2_Change()
   '
   FEX = COMESACT
   SMX0% = SEMANU%(FEX, DELX, ALX)
   '
   MESPAN$ = MESANO$(Label2, DES%, HAS%)
   FEX = DES%
   FEX1 = DIANTE(FEX)
   While DIASEM%(FEX1) <> 1
     FEX1 = DIANTE(FEX1)
   Wend
   ABC$ = FECHATEX$(FEX1)
   INISEM(0).Caption = Left$(FECHATEX$(FEX1), 5)
   ISEMI%(KKI%) = FEX1
   For KKI% = 0 To 5
     If KKI% > 0 Then
10     FEX1 = DIPOST(FEX1)
       If DIASEM%(FEX1) <> 1 Then GoTo 10
       INISEM(KKI%).Caption = Left$(FECHATEX$(FEX1), 5)
       ISEMI%(KKI%) = FEX1
     End If
     FEX = FEX1
     SMX% = SEMANU(FEX, DELX, ALX)
     '
     Text3(KKI%).Enabled = True
     SMXA% = SMX%: If SMXA% < SMX0% Then SMXA% = SMX% + 53
     If SMXA% < SMX0% Then Text3(KKI%).Enabled = False
     '
     SEMI%(KKI%) = SMX%
     If SMX% < SMX0% - 1 Then
        SEMI%(KKI%) = SEMI%(KKI%) + 53
     End If
     '
     If KKI% > 0 Then
       If SEMI%(KKI%) < SEMI%(KKI% - 1) Then
         SEMI%(KKI%) = SEMI%(KKI%) + 53
       End If
     End If
     '
     SMTX$ = Right$("      S." + TRIM$(Str$(SMX%)), 6)
     TTXX$ = Label1.Caption
     Mid$(TTXX$, 46 + 6 * KKI%, 6) = SMTX$
     Label1.Caption = TTXX$
   Next KKI%
   '
   FEX1 = ISEMI%(5)
   FEX2 = DIPOST(FEX1)
   While DIENTRE%(FEX1, FEX2) < 7
     FEX2 = DIPOST(FEX2)
   Wend
   ISEMI%(6) = FEX2
   '
   DoEvents
   '
   Call MUEPROGENT
   Call MUEPROCAR
   '
End Sub

Sub CARPROGENT()
     '
     FEX = COMESACT%
     FECOMACT = COMESACT%
     SMX0% = SEMANU%(FEX, DELX, ALX)
     CACODIP% = 0
     '
     Screen.MousePointer = 11
     APROGRA$ = "PROGENT"
     If MODOPLAN% = 1 Then APROGRA$ = "PROGFAB"
     '
     For KKI% = 0 To NUMAS%
       PROGENT$(KKI%) = String$(212, 0)
     Next KKI%
     '
     FEX = COMESACT
     SMX0% = SEMANU%(FEX, DELX, ALX)
     '
     FEX = DELX
     FEX1 = DIANTE(FEX)
     While DIASEM%(FEX1) <> 1
       FEX1 = DIANTE(FEX1)
     Wend
     FEX2% = FEX1: ABC$ = MKI$(FEX2%) + MKI$(0)          'FECHATEX$(FEX1)
20   SEMANANT% = SEMANU(FEX1, DELX, ALX)
     While SEMANU%(FEX1, DELX, ALX) = SEMANANT%
       FEX1 = DIPOST(FEX1)
     Wend
     While DIASEM%(FEX1) <> 1
       FEX1 = DIPOST(FEX1)
     Wend
     FEX2% = FEX1: ABC$ = ABC$ + MKI$(FEX2%) + MKI$(0)
     If Len(ABC$) < 216 Then GoTo 20
     PROGENT$(0) = ABC$
     
     INISEM(0).Caption = Left$(FECHATEX$(FEX1), 5)
125  FIN& = ULTREG&(APROGRA$)
     For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       P00$ = REGLEIDO$(APROGRA$, I&)
       P1$ = Left$(P00$, 2)
       CI% = CVI(P1$)
       '
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           If STATUIT%(CI%) >= 0 Then
             '
             FENTRE% = CVI(Mid$(P00$, 3, 2))
             FEXEN = FENTRE%
             CANPRO = CVS(Mid$(P00$, 7, 4))
             '
             For KKI% = 1 To CACODIP%
               If CODIPRO%(KKI%) = CI% Then GoTo 10
             Next KKI%
             CACODIP% = CACODIP% + 1
             CODIPRO%(CACODIP%) = CI%
             KKI% = CACODIP%
             '
             If FENTRE% < COMESACT% Then
               If TRAPROGANT% = 0 Then
                 TTXX$ = "Hay Entregas Programadas\Anteriores al Comienzo del Mes Actual\"
                 TTXX$ = TTXX$ + "  \Defina el Tratamiento Correspondiente a Estos Programas:\\Anular\Reprogramar al " + FECHATEX$(FECOMACT)
                 TRAPROGANT% = COMALTER%(TTXX$)
                 If TRAPROGANT% <> 2 Then TRAPROGANT% = 1
               End If
               If TRAPROGANT% <> 2 Then GoTo 129 ' anular
               FENTRE% = COMESACT%
8              If DIASEM%(FENTRE%) <> 1 Then     ' PARA LLEVARLO AL LUNES SIGUIENTE - EPB 25/02/11
                  FENTRE% = DIPOST(FENTRE%)
                  GoTo 8
               End If
               FEXEN = FENTRE%
             End If
             '
10           If FENTRE% >= COMESACT% Or SEMANU%(FEXEN, DELX, ALX) = SMX0% Then
               For KUI% = 0 To 52
                 If FENTRE% >= CVI(Mid$(PROGENT$(0), 1 + 4 * KUI%, 2)) Then
                   If FENTRE% < CVI(Mid$(PROGENT$(0), 1 + 4 * (KUI% + 1), 2)) Then
                     ISEMA% = KUI%
                     GoTo 30
                   End If
                 End If
               Next KUI%
               GoTo 129
               '
30             If ISEMA% >= 0 Then
                 If ISEMA% <= 53 Then
                   CANPRO = CANPRO + CVS(Mid$(PROGENT$(KKI%), 1 + 4 * ISEMA%, 4))
                   Mid$(PROGENT$(KKI%), 1 + 4 * ISEMA%, 4) = MKS$(CANPRO)
                 End If
               End If
             End If
           End If
         End If
       End If
129  Next I&
     '
     If TRAPROGANT% = 2 Then
        Call COMUNI("Recuerde GUARDAR el Plan de Entrega / Fabricacion")
     End If
     '
     Screen.MousePointer = 1
End Sub

Sub GRAPROGENT()
     '
     Screen.MousePointer = 11
     FEX = COMESACT%
     SMX0% = SEMANU%(FEX, DELX, ALX)
     IJ& = 0
     FIN& = List1.ListCount
     '
     APROGRA$ = "PROGENT"
     If MODOPLAN% = 1 Then APROGRA$ = "PROGFAB"
     '
     For KKL& = 1 To List1.ListCount
       Call TRACE(20, KKL&, FIN&)
       CI% = CODINT%(Left$(List1.List(KKL& - 1), 16))
       For UI& = 1 To CACODIP%
         If CODIPRO%(UI&) = CI% Then
           FEX = COMESACT%
10         If DIASEM%(FEX) <> 1 Then
             FEX = DIANTE(FEX)
             GoTo 10
           End If
           For JJ% = 0 To 52
             CAPROG = CVS(Mid$(PROGENT$(UI&), 1 + 4 * JJ%, 4))
             FENTRE% = CVI(Mid$(PROGENT$(0), 1 + 4 * JJ%, 2))
             If CAPROG > 0.05 Then
               RGXX$ = REGBLAN$(APROGRA$)
               Call REPLA(RGXX$, MKI$(CI%), 1, 2)
               If FENTRE% < COMESACT% Then FENTRE% = COMESACT%
               Call REPLA(RGXX$, MKI$(FENTRE%), 3, 2)
               Call REPLA(RGXX$, MKS$(CAPROG), 7, 4)
               IJ& = IJ& + 1
               Call GRAREG(APROGRA$, RGXX$, IJ&)
             End If
           Next JJ%
         End If
       Next UI&
     Next KKL&
     '
     Call SETULTREG(APROGRA$, IJ&)
     Call CIERRARCH(APROGRA$)
     '
     Screen.MousePointer = 1

End Sub

Sub MUEPROGENT()
   '
   FEX = COMESACT
   SMX0% = SEMANU%(FEX, DELX, ALX)
   TPXA = List1.TopIndex
   '
   Screen.MousePointer = 11
   List1.Clear
   For UI& = 1 To CACODIP%
     CI% = CODIPRO%(UI&)
     If CI% > 0 Then
       If CI% <= NUMAS% Then
         If PROGENT$(UI&) <> String$(212, 0) Then
           For KKI% = (-1) To 7: CAPRO(KKI%) = 0: Next KKI%
           SUMACOD = 0
           For JJ% = 0 To 52
             CAPROG = CVS(Mid$(PROGENT$(UI&), 1 + 4 * JJ%, 4))
             FEPRG% = CVI(Mid$(PROGENT$(0), 1 + 4 * JJ%, 2))
             '
             If FEPRG% < ISEMI%(0) Then
                 CAPRO(-1) = CAPRO(-1) + CAPROG
               Else
                 For KKI% = 0 To 5
                   If FEPRG% >= ISEMI%(KKI%) Then
                     If FEPRG% < ISEMI%(KKI% + 1) Then
                       CAPRO(KKI%) = CAPRO(KKI%) + CAPROG
                       GoTo 15
                     End If
                   End If
                 Next KKI%
                 CAPRO(6) = CAPRO(6) + CAPROG
             End If
15         Next JJ%
           TTXX$ = Space$(128)
           Call REPLA(TTXX$, CODEXT$(CI%), 1, 16)
           Call REPLA(TTXX$, DESCRIT0$(CI%), 18, 48)
           For KKI% = -1 To 6
             Call REPLA(TTXX$, CSTRING$(MKS$(CAPRO(KKI%)), 3, 6, 0, 2), 45 + 6 * KKI%, 6)
           Next KKI%
           Call REPLA(TTXX$, CSTRING$(MKS$(SUMACOD), 3, 6, 0, 2), 87, 6)
           ABUS$ = Left$(TTXX$, 16)
           For KKJ& = List1.ListCount To 1 Step -1
             If Left$(List1.List(KKJ& - 1), 16) < ABUS$ Then
               List1.AddItem TTXX$, KKJ&
               GoTo 19
             End If
           Next KKJ&
           List1.AddItem TTXX$, 0
         End If
       End If
     End If
19 Next UI&
   On Error Resume Next
   List1.TopIndex = TPXA
   On Error GoTo 0
   DoEvents
   Screen.MousePointer = 1
   '
End Sub

Sub MUEPROCAR()
   '
   FEX = COMESACT
   SMX0% = SEMANU%(FEX, DELX, ALX)
   '
   CI0% = CODINT%(Text1)
   For KKI% = (-1) To 7: CAPRO(KKI%) = 0: Next KKI%
   For UI& = 1 To CACODIP%
     CI% = CODIPRO%(UI&)
     If CI% = CI0% Then
       For JJ% = 0 To 52
         CAPROG = CVS(Mid$(PROGENT$(UI&), 1 + 4 * JJ%, 4))
         FEPRG% = CVI(Mid$(PROGENT$(0), 1 + 4 * JJ%, 2))
         If FEPRG% < ISEMI%(0) Then
             CAPRO(-1) = CAPRO(-1) + CAPROG
           Else
             For KKI% = 0 To 5
               If FEPRG% >= ISEMI%(KKI%) Then
                 If FEPRG% < ISEMI%(KKI% + 1) Then
                   CAPRO(KKI%) = CAPRO(KKI%) + CAPROG
                   GoTo 15
                 End If
               End If
             Next KKI%
             CAPRO(6) = CAPRO(6) + CAPROG
         End If
15     Next JJ%
       GoTo 25
       '
     End If
   Next UI&
   '
25 For KKI% = 0 To 5
     Text3(KKI%) = CSTRING$(MKS$(CAPRO(KKI%)), 3, 5, 0, 2)
   Next KKI%
   '
   ABUS$ = AJUSTI$(Text1, 16)
   For IKL& = 1 To List1.ListCount
     If Left$(List1.List(IKL& - 1), 16) = ABUS$ Then
       List1.ListIndex = IKL& - 1
       GoTo 99
     End If
   Next IKL&
   List1.ListIndex = (-1)
   '
99 DoEvents
   '
End Sub

Sub CARLINVE()
   '
   FLINVE$ = LOADFILE$(LUDAT$ + "INVERT.DAT")
   For UI& = 1 To NUINV%
     LINVE$(UI&) = Mid$(FLINVE$, 1 + 18 * UI&, 18)
   Next UI&
End Sub

Private Sub Label8_DblClick()
   If OPCIPLAN% = 1 Then
     If MODOPLAN% = 0 Then
         MODOPLAN% = 1
         Frame1.ForeColor = RGB(80, 0, 0)
         Frame1.Caption = "Plan Maestro de Fabricación"
         Label8.Visible = False: Label9.Visible = True
         Command6.ToolTipText = "Consulta Plan de Fabricación por Artículo (Pantalla)"
         Command25.ToolTipText = "Listado General del Plan Maestro de Fabricación"
       Else
         MODOPLAN% = 0
         Frame1.ForeColor = RGB(0, 0, 170)
         Frame1.Caption = "Programa de Entregas"
         Label9.Visible = False: Label8.Visible = True
         Command6.ToolTipText = "Consulta Programa de Entregas por Artículo (Pantalla)"
         Command25.ToolTipText = "Listado General del Programa de Entregas"
     End If
     Frame2.ForeColor = Frame1.ForeColor
     Frame2.Caption = "Actualización del " + Frame1.Caption
     Call CARPROGENT
     Call MUEPROGENT
   End If
End Sub

Private Sub Label9_DblClick()
   If OPCIPLAN% = 1 Then
     If MODOPLAN% = 0 Then
         MODOPLAN% = 1
         Frame1.ForeColor = RGB(80, 0, 0)
         Frame1.Caption = "Plan Maestro de Fabricación"
         Label8.Visible = False: Label9.Visible = True
         Command6.ToolTipText = "Consulta Plan de Fabricación por Artículo (Pantalla)"
         Command25.ToolTipText = "Listado General del Plan Maestro de Fabricación"
       Else
         MODOPLAN% = 0
         Frame1.ForeColor = RGB(0, 0, 170)
         Frame1.Caption = "Programa de Entregas"
         Label9.Visible = False: Label8.Visible = True
         Command6.ToolTipText = "Consulta Programa de Entregas por Artículo (Pantalla)"
         Command25.ToolTipText = "Listado General del Programa de Entregas"
     End If
     Frame2.ForeColor = Frame1.ForeColor
     Frame2.Caption = "Actualización del " + Frame1.Caption
     Call CARPROGENT
     Call MUEPROGENT
   End If
End Sub

Private Sub List1_DblClick()
   Text1.TEXT = Left$(List1.TEXT, 16)
End Sub

Private Sub Text1_Change()
   If CODINT%(Text1.TEXT) > 0 Then
       CI% = CODINT(Text1.TEXT)
       Label5 = DESCRIT0$(CI%)
       Call MUEPROCAR
       For KKI% = 0 To 5
         If Text3(KKI%).Enabled = True Then
           Text3(KKI%).SetFocus
           Exit Sub
         End If
       Next KKI%
     Else
       Label5 = ""
       For KKI% = 0 To 5
         Text3(KKI%) = ""
       Next KKI%
       List1.ListIndex = (-1)
   End If
End Sub

Sub GRAPROCAR()
   '
   FEX = COMESACT
   SMX0% = SEMANU%(FEX, DELX, ALX)
   '
   For KKI% = 0 To 5
     If SEMI%(KKI%) < SMX0% Then Text3(KKI%) = ""
   Next KKI%
   '
   CI% = CODINT%(Text1)
   If CI% < 1 Or CI% > NUMAS% Then Exit Sub
   '
   For KKI% = 1 To CACODIP%
     If CODIPRO%(KKI%) = CI% Then GoTo 10
   Next KKI%
   CACODIP% = CACODIP% + 1
   CODIPRO%(KKI%) = CI%
   PROGENT(KKI%) = String$(212, 0)
   KKI% = CACODIP%
   '
10 For KKJ% = 0 To 5
     FENTRE% = ISEMI%(KKJ%)
     For KUI% = 0 To 52
       If FENTRE% >= CVI(Mid$(PROGENT$(0), 1 + 4 * KUI%, 2)) Then
         If FENTRE% < CVI(Mid$(PROGENT$(0), 1 + 4 * (KUI% + 1), 2)) Then
           ISEMA% = KUI%
           GoTo 30
         End If
       End If
     Next KUI%
     GoTo 39
     '
30   If ISEMA% >= 0 Then
       If ISEMA% <= 53 Then
         CANPRO = Val(Text3(KKJ%))
         Mid$(PROGENT$(KKI%), 1 + 4 * ISEMA%, 4) = MKS$(CANPRO)
       End If
     End If
39 Next KKJ%
   '
   Call MUEPROGENT
   '
   ABUS$ = AJUSTI$(Text1, 16)
   For IKL& = 1 To List1.ListCount
     If Left$(List1.List(IKL& - 1), 16) = ABUS$ Then
       List1.ListIndex = IKL& - 1
       GoTo 99
     End If
   Next IKL&
   List1.ListIndex = (-1)
   '
99 DoEvents
   '
End Sub

Private Sub Text3_GotFocus(Index As Integer)
   Text3(Index).SelStart = 0
   Text3(Index).SelLength = Len(Text3(Index))
End Sub

Private Sub Text3_LostFocus(Index As Integer)
   Text3(Index) = CSTRING$(MKS$(Val(Text3(Index))), 3, 5, 0, 2)
End Sub

Sub TRANECOMAI()
   '
   DESTIMAI$ = LCase$(CONTROL$("NECOMPRA", "DESTINEC"))
10 DESTIMAI$ = TRIM$(LCase$(InputBox("Direccion Electronica Destino", "Transferir Necesidades de Compras", DESTIMAI$)))
   If DESTIMAI$ = "" Then Exit Sub
   '
   If InStr(DESTIMAI$, "@") < 2 Then
     Call MENSERR(24, "Destino no Válido")
     GoTo 10
   End If
   '
   Call GRACONTROL("NECOMPRA", "DESTINEC", DESTIMAI$)
   '
   Screen.MousePointer = 11
   TTYY$ = "": TTXX$ = ""
   Call FILESORT("PLACOMP", "", 1, 3)
   Screen.MousePointer = 11
   CIANT% = 0
   For U& = 1 To ULTREG&("PLACOMP")
     XX$ = REGLEIDO$("PLACOMP", U&)
     CIXI% = CVI(Left$(XX$, 2))
     If CIXI% <> CIANT% Then
       If TTXX$ <> "" Then
         TTYY$ = TTYY$ + TTXX$ + Chr$(13) + Chr$(10)
       End If
       CIANT% = CIXI%
       TTXX$ = AJUSTI$(CODEXT$(CIXI%), 16)
       TTXX$ = TTXX$ + AJUSTI$(DESCRIT0$(CIXI%), 64)
       TTXX$ = TTXX$ + AJUSTI$(UNIMED$(CIXI%), 4)
       TTXX$ = TTXX$ + Space$(10)
       TTXX$ = TTXX$ + FORMATNUM$(CIXI%, "F6.0")
     End If
     ' LONGITUD DE TEXTO QUEDO AJUSTADA A 100 CARACTERES
     FEX = CVI(Mid$(XX$, 3, 2))
     CANEC = CVS(Mid$(XX$, 21, 4))
     TTXX$ = TTXX$ + " " + FECHATEX$(FEX) + FORMATNUM$(CANEC, "F10.1") + " "
   Next U&
   TTYY$ = TTYY$ + TTXX$ + Chr$(13) + Chr$(10)
   '
   IDXS$ = RANDSTRING$(-6) ' NEGATIVO PARA FORZAR VALORES NUMERICOS
   IDENMESS$ = IDXS$ + "." + CONTRACLA$(IDXS$)
   '
   MAPISession1.DownLoadMail = False
   MAPISession1.LogonUI = False
   MAPISession1.Action = 1
   MAPIMessages1.SessionID = MAPISession1.SessionID
   '
   MAPIMessages1.MsgIndex = -1
   MAPIMessages1.MsgSubject = "Necesidades de Compras (" + IDENMESS$ + ")"
   MAPIMessages1.MsgNoteText = TTYY$
   MAPIMessages1.RecipIndex = 0
   MAPIMessages1.RecipType = 1
   'MAPIMessages1.RecipDisplayName = MUESTRA$
   MAPIMessages1.RecipAddress = DESTIMAI$
   '
   On Error Resume Next
   MAPIMessages1.Action = 3   ' vbMessageSend
   If Err Then
     MsgBox "Ocurrio Error " & Err
     Resume 90
   End If
   '
   Call COMUNI("Necesidades de Compras Transferidas.")
   '
90 On Error GoTo 0
   Screen.MousePointer = 1
   MAPISession1.Action = 2  ' !!!
   DoEvents
   '
End Sub

Sub BAJMAICOMP()
   '
   OPX% = COMALTER%("Desea Refrescar Información\Remota de Compras ?\\Refrescar Compras\Continuar")
   If OPX% <> 1 Then Exit Sub
   '
   MAPISession1.DownLoadMail = True
   MAPISession1.LogonUI = False
   MAPISession1.Action = 1
   MAPIMessages1.SessionID = MAPISession1.SessionID
   DoEvents
   '
   MAPIMessages1.Fetch
   CAMENSA& = MAPIMessages1.MsgCount
   '
   List2.Clear
   For U& = CAMENSA& To 1 Step -1
     Call TRACE(20, U&, CAMENSA&)
     MAPIMessages1.MsgIndex = U& - 1
     ASUNTO$ = MAPIMessages1.MsgSubject
     If InStr(ASUNTO$, "Informacion de Compras (") > 0 Then
       CLA1$ = Mid$(ASUNTO$, 25, 6)
       CLA2$ = Mid$(ASUNTO$, 32, 6)
       If CLA2$ = CONTRACLA$(CLA1$) Then
         TEXTO$ = MAPIMessages1.MsgNoteText
         List2.Clear
         While Len(TEXTO$) > 0
           PPXX% = InStr(TEXTO$, Chr$(13) + Chr$(10))
           If PPXX% < 1 Then PPXX% = 1 + Len(TEXTO$)
           TTXX$ = Left$(TEXTO$, PPXX% - 1)
           TEXTO$ = Mid$(TEXTO$, PPXX% + 2)
           If TRIM$(TTXX$) <> "" Then List2.AddItem TTXX$
         Wend
         MAPIMessages1.Delete
         GoTo 99
       End If
     End If
     '
89 Next U&
   '
   Call COMUNI("No se Encontró\Nueva Información de Compras")
   '
99 MAPISession1.Action = 2  ' !!!
   If List2.ListCount < 1 Then Exit Sub
   '
   HOII% = HOY(HOS$)
   For U& = 1 To List2.ListCount
     TTXX$ = List2.List(U& - 1)
     NROC& = Val(Mid$(TTXX$, 9, 8))
     FEMI% = FECHANUM(Mid$(TTXX$, 17, 8))
     NPROL& = Val(Mid$(TTXX$, 25, 8))
     If NPROL& > 0 Then
       If NPROL& < 32767 Then
         NPRO% = NPROL&
         DPRO$ = TRIM$(Mid$(TTXX$, 33, 48))
         CIXI% = 0
         On Error Resume Next
         CIXI% = Val(Mid$(TTXX$, 97, 8))
         On Error GoTo 0
         If CIXI% > 0 Then
           PRUNI# = Val(Mid$(TTXX$, 105, 16))
           FECAN$ = TRIM$(Mid$(TTXX$, 121))
           For KU% = 1 To Len(FECAN$) Step 20
             FESOL% = FECHANUM(Mid$(FECAN$, KU%, 8))
             CASOL = Val(Mid$(FECAN$, KU% + 8, 12))
             If FESOL% < HOII% Then FESOL% = HOII%
             If CASOL >= 0.05 Then
               ENCAO$ = FILTERGETRECORD$("OCOMENCA", 1, MKS$(NROC&))
               ' puede que sea nuevo registro o esté anulada
               If CVS(Left$(ENCAO$, 4)) < 1 Or Asc(Mid$(ENCAO$, 54, 1)) = 9 Then
                   Call REPLA(ENCAO$, MKS$(NROC&), 1, 4)
                     FEX = FEMI%
                     REFE$ = FECHATEX$(FEX) + " - " + DPRO$
                   Call REPLA(ENCAO$, REFE$, 5, 44)
                   Call REPLA(ENCAO$, MKI$(FEMI%), 49, 2)
                   Call REPLA(ENCAO$, MKI$(NPRO%), 51, 2)
                   Call REPLA(ENCAO$, Chr$(0), 53, 1)
                   Call REPLA(ENCAO$, Chr$(0), 54, 1)
                     MONECOI% = 1: MONEMI% = 1: TICAMBX = 1
                   Call REPLA(ENCAO$, Chr$(MONECOI%), 83, 1)
                   Call REPLA(ENCAO$, Chr$(MONEMI%), 84, 1)
                   Call REPLA(ENCAO$, MKS$(TICAMBX), 85, 4)
                     HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
                   Call REPLA(ENCAO$, MKI$(HOII%), 55, 2)
                   Call REPLA(ENCAO$, MKI$(Int(HP)), 57, 2)
                   Call REPLA(ENCAO$, MKI$(USNBR%), 59, 2)
                   '
                   PRIDETA& = 1 + ULTREG&("OCOMDETA")
                   ULTDETA& = PRIDETA&
                   '
                   ' grabar encabezado
                   Call REPLA(ENCAO$, MKS$(PRIDETA&), 155, 4)
                   Call REPLA(ENCAO$, MKS$(ULTDETA&), 159, 4)
                   Call REGAPP("OCOMENCA", ENCAO$)
                   Call CIERRARCH("OCOMENCA")
                   '
                 Else
                   If CVI(Mid$(ENCAO$, 51, 2)) <> NPRO% Then ' no coincide Proveedor
                     Call MENSERR(24, "Imposible Procesar O/C " + TRIM$(Str$(NROC&)) + ".\No Coincide Proveedor.")
                     GoTo 90
                   End If
               End If
               '
               RFF$ = RECFILT$("OCOMDETA", 1, MKS$(NROC&))
               For U1& = 1 To Len(RFF$) Step 4
                 REOCO& = CVS(Mid$(RFF$, U1&, 4))
                 OCDT$ = REGLEIDO$("OCOMDETA", REOCO&)
                 If CVI(Mid$(OCDT$, 9, 2)) = CIXI% Then
                   If CVI(Mid$(OCDT$, 95, 2)) = FESOL% Then
                     If Asc(Mid$(OCDT$, 124, 1)) < 9 Then
                       'COINCIDE O/C, CODIGO Y FECHA
                       GoTo 90
                     End If
                   End If
                 End If
               Next U1&
               OCDT$ = REGBLAN$("OCOMDETA")
               Call REPLA(OCDT$, MKS$(NROC&), 1, 4)
               Call REPLA(OCDT$, MKI$(FEMI%), 5, 2)
               Call REPLA(OCDT$, MKI$(NPRO%), 7, 2)
               Call REPLA(OCDT$, MKI$(CIXI%), 9, 2)
               Call REPLA(OCDT$, UNIMED$(CIXI%), 11, 4)
               Call REPLA(OCDT$, MKS$(CASOL), 51, 4)
               Call REPLA(OCDT$, MKS$(CASOL), 55, 4)
               Call REPLA(OCDT$, UNIMED$(CIXI%), 59, 4)
               Call REPLA(OCDT$, MKS$(1), 63, 4)
               Call REPLA(OCDT$, MKD$(PRUNI#), 67, 8)
               Call REPLA(OCDT$, MKD$(PRUNI#), 79, 8)
               Call REPLA(OCDT$, MKD$(PRUNI# * CASOL), 87, 8)
               Call REPLA(OCDT$, MKI$(FESOL%), 95, 8)
               Call REGAPP("OCOMDETA", OCDT$)
             End If
           Next KU%
         End If
       End If
     End If
90 Next U&
   Call CIERRARCH("OCOMDETA")
   Call CIERRARCH("*.*")
   '
   Call COMUNI("Importación de Compras Completa")
   '
End Sub



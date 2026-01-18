VERSION 5.00
Begin VB.Form PROMAT04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0DFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MRP-II - Programación de Materiales"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11880
   Icon            =   "PROMAT04.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   11880
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00006000&
      Height          =   2265
      Left            =   11000
      ScaleHeight     =   2205
      ScaleWidth      =   1005
      TabIndex        =   63
      Top             =   6825
      Width           =   1065
      Begin VB.PictureBox Picture7 
         Height          =   2265
         Left            =   0
         ScaleHeight     =   2205
         ScaleWidth      =   1005
         TabIndex        =   67
         Top             =   0
         Visible         =   0   'False
         Width           =   1065
         Begin VB.Image Image4 
            Height          =   1200
            Left            =   0
            Picture         =   "PROMAT04.frx":030A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1065
         End
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   64
         Top             =   210
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   65
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "PROMAT04.frx":26A2
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   610
         Width           =   1515
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00006000&
      Height          =   1800
      Left            =   11000
      ScaleHeight     =   1740
      ScaleWidth      =   900
      TabIndex        =   60
      Top             =   0
      Width           =   960
      Begin VB.PictureBox Picture6 
         Height          =   1800
         Left            =   0
         ScaleHeight     =   1740
         ScaleWidth      =   900
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   960
         Begin VB.Image Image3 
            Height          =   1800
            Left            =   0
            Picture         =   "PROMAT04.frx":45C4
            Stretch         =   -1  'True
            Top             =   0
            Width           =   960
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
         Height          =   690
         Left            =   105
         Picture         =   "PROMAT04.frx":695C
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "C"
         Top             =   260
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "PROMAT04.frx":6AA6
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Guarda las Modificaciones"
         Top             =   950
         Width           =   650
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0DFC0&
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
      Left            =   3600
      TabIndex        =   9
      Top             =   160
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
         Picture         =   "PROMAT04.frx":6DB0
         Style           =   1  'Graphical
         TabIndex        =   50
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
         TabIndex        =   10
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
            Picture         =   "PROMAT04.frx":71F2
            Style           =   1  'Graphical
            TabIndex        =   11
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
            TabIndex        =   12
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
         TabIndex        =   51
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
         Top             =   1275
         Width           =   960
      End
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   3570
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   330
      Left            =   315
      TabIndex        =   58
      Top             =   105
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0DFC0&
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
      TabIndex        =   46
      Top             =   160
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
         Picture         =   "PROMAT04.frx":74FC
         Style           =   1  'Graphical
         TabIndex        =   47
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
         Picture         =   "PROMAT04.frx":7806
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Consulta Programa de Entregas por Artículo (Pantalla)"
         Top             =   420
         Width           =   600
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   57
         ToolTipText     =   "Doble-Click para Cambiar Modalidad de Programación"
         Top             =   1050
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   56
         ToolTipText     =   "Doble-Click para Cambiar Modalidad de Programación"
         Top             =   350
         Visible         =   0   'False
         Width           =   380
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0DFC0&
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
      Left            =   7300
      TabIndex        =   3
      Top             =   160
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
         TabIndex        =   55
         Top             =   1155
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0DFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   330
         Left            =   630
         TabIndex        =   52
         Top             =   1260
         Width           =   2745
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0DFC0&
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
            TabIndex        =   54
            Top             =   0
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0DFC0&
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
            TabIndex        =   53
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
         TabIndex        =   4
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
            Picture         =   "PROMAT04.frx":7B10
            Style           =   1  'Graphical
            TabIndex        =   7
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
            Picture         =   "PROMAT04.frx":7F52
            Style           =   1  'Graphical
            TabIndex        =   6
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
            Picture         =   "PROMAT04.frx":825C
            Style           =   1  'Graphical
            TabIndex        =   5
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
            TabIndex        =   8
            Top             =   105
            Width           =   1170
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0DFC0&
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
      Top             =   6720
      Width           =   10750
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   9135
         TabIndex        =   32
         Top             =   630
         Width           =   730
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   8410
         TabIndex        =   31
         Top             =   630
         Width           =   730
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   7685
         TabIndex        =   30
         Top             =   630
         Width           =   730
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   6960
         TabIndex        =   29
         Top             =   630
         Width           =   730
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   6235
         TabIndex        =   28
         Top             =   630
         Width           =   730
      End
      Begin VB.CommandButton Command4 
         Height          =   670
         Left            =   9975
         Picture         =   "PROMAT04.frx":8566
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   315
         Width           =   600
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0DFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   225
         Left            =   5565
         TabIndex        =   34
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
            Index           =   10
            Left            =   2880
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
            Index           =   9
            Left            =   2160
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
            Index           =   8
            Left            =   1420
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
            Index           =   7
            Left            =   705
            TabIndex        =   36
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
            TabIndex        =   35
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
         Picture         =   "PROMAT04.frx":8870
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   630
         Width           =   175
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   5525
         TabIndex        =   27
         Top             =   630
         Width           =   730
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
         TabIndex        =   25
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
         TabIndex        =   41
         Top             =   400
         Width           =   3270
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0DFC0&
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
      Height          =   4740
      Left            =   105
      TabIndex        =   0
      ToolTipText     =   "Doble-Click para Cambiar Modalidad de Programación"
      Top             =   1880
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
         Height          =   3435
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
         TabIndex        =   13
         Top             =   315
         Width           =   11460
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   225
            Left            =   5490
            TabIndex        =   18
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
               TabIndex        =   24
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
               TabIndex        =   23
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
               TabIndex        =   22
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
               TabIndex        =   21
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
               TabIndex        =   20
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
               TabIndex        =   19
               Top             =   0
               Width           =   645
            End
         End
         Begin VB.CommandButton Command3 
            Height          =   330
            Left            =   9135
            Picture         =   "PROMAT04.frx":8B7A
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton Command2 
            Height          =   330
            Left            =   5670
            Picture         =   "PROMAT04.frx":8E84
            Style           =   1  'Graphical
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
            Top             =   315
            Width           =   11460
         End
      End
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11865
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   840
      Picture         =   "PROMAT04.frx":918E
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
      TabIndex        =   49
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
   Begin VB.Menu Archiv 
      Caption         =   "Archivos"
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "PROMAT04"
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

Dim MODIFPLAN%        ' SEÑAL DE CONTROL SI SE MODIFICó EL PLAN SIN GUARDARLO

Private Sub BOTSEL_Click(Index As Integer)
   Call SELECHO("MASTER")
   Text1 = VALACT1$("MASTER")
End Sub

Private Sub Command1_Click()
   '
   If MODIFPLAN% > 0 Then
     If COMALTER%("¿ Abandona sin Guardar Cambios ?\\No, Cancelar\Si, Descartar Cambios") = 2 Then
       Call CIERRARCH("*.*")
       Call FINAL("")
     End If
   Else
     Call CIERRARCH("*.*")
     Call FINAL("")
   End If
   '
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   OPPROMRP04.Show 1
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
    TTXX$ = "Cobertura Completa\Cobertura Inmediata"
    TTXX$ = TTXX$ + "\Resumen de Cobertura\Disponibilidad Multinivel"
    OP% = ALTERNA%(TTXX$)
    If OP% = 3 Then
         ANACOB04.Show 1
         GoTo 90
       ElseIf OP% = 4 Then
         Call DISPOMUL
         GoTo 90
    End If
    '
    Select Case OP%
      Case 1
        If Option1.Value = True Then
            Call CONECRUN("*COBERFA", "")
          Else
            Call FILESORT("COBERFA", "COBERFB", 2, 2, "ASC")
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
90  Command15.Enabled = True
End Sub

Sub DISPOMUL()
    '
    Dim TIENEPEDI%(32767)
    Dim EXIS(16384, 6), ACTI%(16384)
    '
    Call ABREPEDCLI
    SQLX$ = " SELECT * FROM PEDDETA "
    SQLX$ = SQLX$ + "WHERE (Status >= 0 AND Status < 8) "   ' para que no este anulada ni cumplida ni pendiente de aprobacion
    SQLX$ = SQLX$ + "AND CantEnt < (CanPed - 0.005) "
    SQLX$ = SQLX$ + "AND VenRep = 'V' "
    SQLX$ = SQLX$ + "ORDER BY FeSolEnt DESC "
    '
    Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    For JL& = 1 To 32767 'PONGO EN 0 EL VECTOR
       TIENEPEDI%(JL&) = 0
    Next JL&
    '
    ' CARGA UN VECTOR CON LOS CLIENTES CON PEDIDOS PENDIENTES
    With PEDTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
          NCLIE1% = !NROCLIE
          If NCLIE1% > 0 Then
             TIENEPEDI%(NCLIE1%) = 1 'MARCO LOS CLIENTES CON PEDIDOS
          End If
        .MoveNext
        Loop
      End If
    End With
    '
    For JL& = 1 To 32767
      If TIENEPEDI%(JL&) = 1 Then
         NCLIE2$ = FORMATNUM$(JL&, "F6.0")
         RASOCLIE2$ = RASOCLI$(JL&)
         Z$ = Space(128)
         Call REPLA(Z$, NCLIE2$, 1, 6)
         Call REPLA(Z$, RASOCLIE2$, 15, 64)
         With DISMULT04
          .LICOPEN.AddItem Z$
         End With
      End If
    Next JL&
    Screen.MousePointer = 1
    If DISMULT04.LICOPEN.ListCount > 0 Then
       DISMULT04.Show 1
       Call FRESHALL
       If DISMULT04.Label2.Caption <> "OK" Then
         Call COMUNI("Cancelado !!!")
         GoTo 9999
       End If
       '
       ' ACA VUELVE A PONER EN '0' EL VECTOR DE CLIENTES
       CORTEPED% = FECHANUM(DISMULT04.Label7)
       CORTECOM% = FECHANUM(DISMULT04.label3)
       For JL& = 1 To 32767
         TIENEPEDI%(JL&) = 0
       Next JL&
       '
       ' ACA RE-CARGA EL VECTOR PONIENDO EN '1' LOS CLIENTES SELECCIONADOS
       TCLISEL$ = ""
       For JL& = 1 To DISMULT04.LICOPEN.ListCount
         If DISMULT04.LICOPEN.Selected(JL& - 1) = True Then
           ICLIE1& = Val(Left$(DISMULT04.LICOPEN.List(JL& - 1), 6))
           If ICLIE1& > 0 Then
             If ICLIE1& <= 32767 Then
               TIENEPEDI%(ICLIE1&) = 1
               NCLISE% = ICLIE1&
               If TCLISEL$ <> "" Then TCLISEL$ = TCLISEL$ + " - "
               TCLISEL$ = TCLISEL$ + TRIM$(FANTCLI$(NCLISE%))
             End If
           End If
         End If
       Next JL&
     End If
     '
     With PEDTEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           NCLIE1% = !NROCLIE
           If NCLIE1% > 0 Then
             If TIENEPEDI%(NCLIE1%) > 0 Then 'MARCO LOS CLIENTES CON PEDIDOS
               CAPEN = !CanPed - !CantEnt
               '
               If CAPEN > 0 Then
                 If CORTEPED% < 1 Or !FeSolEnt <= CVDAT(CORTEPED%) Then ' OJO FECHA DE CORTE
                    CIIX% = !CodiInt
                    EXIS(CIIX%, 4) = EXIS(CIIX%, 4) + CAPEN
                    ACTI%(CIIX%) = 1
                 End If
               End If
               '
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     '
     For CIXI% = 1 To NUMAS%
       If ACTI%(CIXI%) = 1 Then
         EXIS(CIXI%, 1) = STOCKACT(CIXI%)
         EXIS(CIXI%, 5) = STOMIN(CIXI%)
         NECENE = EXIS(CIXI%, 4) + EXIS(CIXI%, 5) - EXIS(CIXI%, 1)
         If NECENE >= 0.05 Then
           Call EXPLOMUL(CIXI%)
           For JJ1% = 1 To CONTAMUL%
             CIX2% = CIK%(JJ1%)
             EXIS(CIX2%, 6) = EXIS(CIX2%, 6) + CANT(JJ1%) * NECENE
             ACTI%(CIX2%) = 2
           Next JJ1%
         End If
       End If
     Next CIXI%
     '
     FIN& = NUMAS%
     For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       CIIX% = I&
       If ACTI%(CIIX%) > 0 Then
         EXIS(I&, 1) = STOCKACT(CIIX%)
         EXIS(I&, 5) = STOMIN(CIIX%)
       End If
     Next I&
     '
     Call ABRECOMPRAS
     SQLX$ = " SELECT * FROM OCOMDETA "
      'SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + "WHERE Stat_Ocom < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND  Stat_Ocom <> -1 " 'Para que no  hayan O/C's sin aprobar
     SQLX$ = SQLX$ + "AND Cant_Rec < (Cant_Ocom - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY Nume_Ocom ASC "
     '
     Set OcoTemp = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     With OcoTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CIIX% = !COD_INTE
           If ACTI%(CIIX%) > 0 Then
              If CORTECOM% < 1 Or !Fentre_Sol <= CVDAT(CORTECOM%) Then ' OJO FECHA DE CORTE
                 NUORSE& = !N_OrServ
                 If NUORSE& < 1 Then
                   COEFCAN = !Coef_Unids
                   If COEFCAN < 0.005 Then COEFCAN = 1
                   CACOM = !Cant_Ocom * COEFCAN
                   CAREC = !Cant_Rec * COEFCAN
                   CAPEN = CACOM - CAREC
                   If CAPEN > 0 Then
                      EXIS(CIIX%, 3) = EXIS(CIIX%, 3) + CAPEN
                   End If
                 End If
              End If
           End If
         .MoveNext
         Loop
       End If
     End With
     '
     '\\\OT-06-0149-WAR-25/04/06///
     Call ABRECOMPRAS
     SQLX$ = " SELECT * FROM AUTOENT "
       'SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + "WHERE Stat_Aten < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND  Stat_Aten <> -1 " 'Para que no  hayan O/C's sin aprobar
     SQLX$ = SQLX$ + "AND Acuen_Aten < (Cant_Aten - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY Nume_Aten ASC "
     '
     Set AtenTemp = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     With AtenTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CIIX% = !COD_INTE
           If ACTI%(CIIX%) > 0 Then
              NUORSE& = 0 ' !N_OrServ
              If NUORSE& < 1 Then
                COEFCAN = 1 ' !Coef_Unids
                If COEFCAN < 0.005 Then COEFCAN = 1
                CACOM = !Cant_Aten * COEFCAN
                CAREC = !Acuen_Aten * COEFCAN
                CAPEN = CACOM - CAREC
                If CAPEN > 0 Then
                  EXIS(CIIX%, 3) = EXIS(CIIX%, 3) + CAPEN
                End If
              End If
           End If
         .MoveNext
         Loop
       End If
     End With
     '\\\OT-06-0149-WAR-FIN///
     Call ABRESTOCKS
     SQLX$ = " SELECT * FROM BOLRECEP "
     SQLX$ = SQLX$ + " WHERE Status = 0 "
     '
     Set COBRECEPTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With COBRECEPTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While (.EOF = False)
           CIIX% = !COD_INTE
           If ACTI%(CIIX%) > 0 Then
              CACOM = !CANTNOM
              CAREC = !CanToApro
              CAPEN = CACOM - CAREC
              If CAPEN > 0 Then
                EXIS(CIIX%, 2) = EXIS(CIIX%, 2) + CAPEN
              End If
           End If
         .MoveNext
         Loop
       End If
     End With
     '
50   J& = 0
     NVII& = NUMAS%
     For I& = 1 To NVII&
       Call TRACE(20, I&, NVII&)
       CIIX% = I&
       If STATUIT%(CIIX%) > 0 Then
         CODE$ = CODEXT$(CIIX%)
         If ACTI%(CIIX%) > 0 Then
         'If CODE$ >= COD1$ Then
          'If CODE$ <= COD2$ Then
            'If TMT% = 0 Or TIMATE%(CIIX%) = TMT% Then
              X$ = REGBLAN$("STODISP")
              Call REPLA(X$, MKI$(CIIX%), 1, 2)
              Call REPLA(X$, Chr$(TIMATE%(CIIX%)), 3, 1)
              SUMA = 0: HAYMO% = 0
              For K% = 1 To 6
                SIG% = 1: If K% >= 4 Then SIG% = (-1)
                SUMA = SUMA + SIG% * EXIS(CIIX%, K%)
                If Abs(EXIS(CIIX%, K%)) >= 0.05 Then HAYMO% = 1
                Call REPLA(X$, MKS$(EXIS(CIIX%, K%)), 4 * K% + 1, 4)
              Next K%
              '
              If EXIS(CIIX%, 4) > 0.05 Then   ' HAY PEDIDOS
                 SUMA = (-1) * EXIS(CIIX%, 6)
              End If
              '
55            If SUMA > 0.05 Then
                  Call REPLA(X$, MKS$(SUMA), 29, 4)
                ElseIf SUMA < (-0.05) Then
                  Call REPLA(X$, MKS$((-1) * SUMA), 33, 4)
                  'Call CAROEMI(CIIX%, Abs(SUMA), HOII%)
              End If
              '
              'If (NIVEFIL% = 1) Or (NIVEFIL% = 2 And HAYMO% = 1) Or (NIVEFIL% = 3 And SUMA <= (-0.05)) Then
                Call REPLA(X$, MKS$(LOTESTAN(CIIX%)), 37, 4)
                J& = J& + 1
                Call GRAREG("STODISP", X$, J&)
              'End If
            End If
         ' End If
        'End If
      End If
    Next I&
    '
60  Call SETULTREG("STODISP", J&)
    Call CIERRARCH("STODISP")
    Call FILESORT("STODISP", "", 1, 1, "ASC")
    '
    Call CIERRARCH("*.*")
    '
    TTXX$ = TCLISEL$
    If CORTEPED% > 0 Then
       FEX = CORTEPED%
       TTXX$ = TTXX$ + " / Corte Ped: " + FECHATEX$(FEX)
    End If
    If CORTECOM% > 0 Then
       FEX = CORTECOM%
       TTXX$ = TTXX$ + " / Corte Com: " + FECHATEX$(FEX)
    End If
    Call HEADERS("STODISP", "")
    Call HEADERS("STODISP", "Filtros: " + TTXX$)
    '
    Call CONECRUN("*DISPOMU", "")
    Exit Sub
    '
9999 End Sub
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
    APROGRA$ = "*LPROGEN"
    If MODOPLAN% = 1 Then APROGRA$ = "*LPROGFA"
    Call CONECRUN(APROGRA$, "")
    Command25.Enabled = True
End Sub

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
     MODIFPLAN% = 1
   End If
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   If DERACCE%("GUARDMRP", "Guardar Plan Maestro Ventas / Fabricación") >= 2 Then
     If OPCIPLAN% = 1 Then
       MODOP$ = "ENTREGA": If MODOPLAN% = 1 Then MODOP$ = "FABRICA"
       Call GRACONTROL("PROMRP", "MODOPLAN", MODOP$)
     End If
     Call GRAPROGENT
     MODIFPLAN% = 0
   End If
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
   If MODIFPLAN% > 0 Then
     Call MENSERR(24, "Imposible Re-Cálculo sin Guardar Cambios")
     GoTo 99
   End If
   '
   If DERACCE%("CALCUMRP", "Cálculo Programación de Materiales MRP") >= 2 Then
     FORCANEC04.Show 1
     Refresh
     DoEvents
     If App.EXEName <> "FLEXO00" Then
       Call CONECRUN(App.EXEName, "")
       Call FINAL("")
     End If
   End If
99 Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    If Option1.Value = True Then
          Call FINAL("*CPLAFAB")
        Else
          OPPY% = COMALTER%("Formato de Salida:\\Listado\Tabla")
          If OPPY% = 2 Then
             Call CUAPROCOF("PLAFABR")
             GoTo 99
           Else
             Call FINAL("*LPLAFAB")
          End If
    End If
99  Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    '
    Command9.Enabled = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    AXXX$ = "": ARPLA$ = "PLACOMP"
    If InStr(EPAC$, "FERVI") > 0 Then
       AXXX$ = "Q"
       ARPLA$ = "PLACOMQ"
    End If
    If Option1.Value = True Then
        Call FINAL("*CPLACOM" + AXXX$)
      Else
        OPPY% = COMALTER%("Formato de Salida:\\Listado\Tabla")
        If OPPY% = 2 Then
           Call CUAPROCOF(ARPLA$)
           GoTo 99
        End If
        '
        OPX% = COMALTER%("Elija Secuencia / Agrupamiento\\Por Proveedor\Por Artículo")
        If OPX% = 1 Then
            Call FILESORT(ARPLA$, "", 8, 8, "ASC")
          ElseIf OPX% = 2 Then
            Call FILESORT(ARPLA$, "", 1, 3, "ASC")
        End If
        '
        TNT% = 0
        If AXXX$ = "" Then
10        TMT% = Asc(Left$(REGSEL$("TATIMAT"), 1))
          '\\\OT-06-0130-WAR-05/04/06///
          Call SELECHO("FECHA/Fecha de Corte")
          VDEVU$ = VALACT1$("FECHA")
          FLIM1% = FECHANUM(VDEVU$)
          If TRIM$(VDEVU$) = "//" Then FLIM1% = 32767
          '\\\OT-06-0130-WAR-FIN///
          Screen.MousePointer = 11
          Call COPYSTRU("PLACOMP", "PLACOMF")
          JJ& = 0: FIN& = ULTREG&("PLACOMP")
          For U& = 1 To FIN&
            Call TRACE(20, U&, FIN&)
            X$ = REGLEIDO$("PLACOMP", U&)
            CIXI% = CVI(Left$(X$, 2))
            F1% = CVI(Mid$(X$, 3, 2))
            If (TIMATE%(CIXI%) = TMT% Or TMT% = 0) Then
              If F1% <= FLIM1% Then '\\\OT-06-0130-WAR-05/04/06///
                JJ& = JJ& + 1
                Call GRAREG("PLACOMF", X$, JJ&)
              End If
            End If
          Next U&
          Screen.MousePointer = 1
          If JJ& < 1 Then
            Call MENSERR(24, "Sin Necesidades\para la Selección.")
            GoTo 99
          End If
          Call SETULTREG("PLACOMF", JJ&)
          Call HEADERS("PLACOMF", "")
          Call HEADERS("PLACOMF", "Tipo de Material: " + ECOARCH$("TATIMAT", Chr$(TMT%)))
          Call FINAL("*LPLACOF")
          GoTo 99
        End If
        Call FINAL("*LPLACOM" + AXXX$)
    End If
99  Command9.Enabled = True
End Sub

Sub CUAPROCOF(ARPLA$)
    '
    Dim PEDIPRO#(32767, 12), TIENEPED%(32767)
    '
    TNT% = 0
    TMT% = Asc(Left$(REGSEL$("TATIMAT"), 1))
    '
    Call GENVECMES
    VECPEDI$ = ""
    FIN& = ULTREG&(ARPLA$)
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$(ARPLA$, U&)
       CI% = CVI(Left$(XX$, 2))
       FENTRESOL% = CVI(Mid$(XX$, 3, 2))
       CANPRO = CVS(Mid$(XX$, 21, 4))
       If CI% > 0 Then
         If (TMT% = 0 Or TIMATE%(CI%) = TMT%) Then
           If Abs(CANPRO) > 0 Then
             For KKI% = 0 To 13
                If FENTRESOL% >= INIMES%(KKI%) Then
                  If FENTRESOL% < INIMES%(KKI% + 1) Then
                    PEDIPRO#(CI%, KKI%) = PEDIPRO#(CI%, KKI%) + CANPRO
                    TIENEPED%(CI%) = 1
                    GoTo 19
                  End If
                End If
             Next KKI%
           End If
         End If
       End If
19  Next U&
    '
    JJ& = 0
    For NC1% = 1 To 32766
      If TIENEPED%(NC1%) > 0 Then
        TTXX$ = MKI$(NC1%) + AJUSTI$(UNIMED$(NC1%), 4) + String(10, 0)
        SUMATO# = 0
        For KKI% = 0 To 12
           TTXX$ = TTXX$ + MKD$(PEDIPRO#(NC1%, KKI%))
           SUMATO# = SUMATO# + PEDIPRO#(NC1%, KKI%)
        Next KKI%
        TTXX$ = TTXX$ + MKD$(SUMATO#)
        JJ& = JJ& + 1
        Call GRAREG("TAPROCOM", TTXX$, JJ&)
      End If
    Next NC1%
    Call SETULTREG("TAPROCOM", JJ&)
    Call FILESORT("TAPROCOM", "", 1, 1, "ASC")
    '
    ' ESTO ES PARA PONER EL ENCABEZADO DEL LISTADOR
    TTXX$ = AJUSTI$("Código", 16) + " Descripción                      U-M "
    For KKII% = 0 To 12
      TTXX$ = TTXX$ + AJUSTD$(DENOMES$(KKII%), 12)
    Next KKII%
    TTXX$ = TTXX$ + AJUSTD$("Total", 12)
    Call SAVEFILE(LUCOM$ + "TAPROCOM.HDR", TTXX$)
    '
    Call HEADERS("TAPROCOM", "")
    If ARPLA$ = "PLAFABR" Then
        Call HEADERS("TAPROCOM", "Fabricacion / " + ECOARCH$("TATIMAT", Chr$(TMT%)))
    Else
        Call HEADERS("TAPROCOM", "Compras / " + ECOARCH$("TATIMAT", Chr$(TMT%)))
    End If
    '
30  Call FINAL("*TAPROCOM")
    '
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

   Call CENTRAFORM(Me)
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.Picture6.Visible = True
      Me.Picture6.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Frame7.BackColor = &HFCD7B6
      Me.Frame8.BackColor = &HFCD7B6
      Me.Option1.BackColor = &HFCD7B6
      Me.Option2.BackColor = &HFCD7B6
      Me.Label1.BackColor = &HFCD7B6
      Me.Label2.BackColor = &HFCD7B6
      Me.Picture1.BackColor = &HFCD7B6
      
      Me.Picture7.Visible = True
      MARCOBARRA.Top = 800
      
      Set MARCOBARRA.Container = Me.Picture7
      MARCOBARRA.ZOrder 0
    End If

   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
     '
10   Caption = Caption + "  -  " + EPAC$
     Call GRATITFOR(Caption)
     '
   End If
   '
   OPENFORMS = DoEvents
   '
   FECHOR$ = Control$("CANECOM", "FECHOCAL")
   Label7.Caption = Left$(FECHOR$, 8): Label7.Refresh
   Label95.Caption = Mid$(FECHOR$, 11, 5): Label95.Refresh
   '
   PRISEMIN% = 53
   If Val(Control$("", "PRISEMIN")) = 1 Then PRISEMIN% = 1
   '
   Call CARPROGENT
   '
   FECHAC% = HOY(HO$)
   Label2.Caption = METEXTO$(FECHAC%)
   '
   Show
   Refresh

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

Private Sub grabar_Click()
    Call Command5_Click
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
   Call FILESORT("CAPANECE", "", 1, 3, "ASC")
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
     MODIFPLAN% = 0
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
               FEXEN = FENTRE%
               MODIFPLAN% = 1
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
     Text3(KKI%) = CSTRING$(MKS$(CAPRO(KKI%)), 3, 6, 0, 2)
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

Private Sub salir_Click()
   Call Command1_Click
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
   Text3(Index) = CSTRING$(MKS$(Val(Text3(Index))), 3, 6, 0, 2)
End Sub

Sub TRANECOMAI()
   '
   DESTIMAI$ = LCase$(Control$("NECOMPRA", "DESTINEC"))
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
   Call FILESORT("PLACOMP", "", 1, 3, "ASC")
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



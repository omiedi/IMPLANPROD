VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form PROMAT07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0DFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MRP-II - Programación de Materiales"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11880
   Icon            =   "PROMAT07.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   11880
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   360
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   58
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00006000&
      Height          =   2265
      Left            =   11000
      ScaleHeight     =   2205
      ScaleWidth      =   1005
      TabIndex        =   54
      Top             =   6825
      Width           =   1065
      Begin VB.CommandButton Command13 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   57
         ToolTipText     =   "Copia de MRP"
         Top             =   45
         Width           =   690
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   55
         Top             =   405
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   56
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "PROMAT07.frx":030A
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
      TabIndex        =   51
      Top             =   0
      Width           =   960
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "PROMAT07.frx":222C
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   260
         Width           =   650
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "PROMAT07.frx":2376
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Inicia Proceso de Cálculo"
         Top             =   950
         Width           =   650
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0DFC0&
      Caption         =   "Cálculo de Necesidades"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
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
      Width           =   3345
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
         Left            =   2400
         Picture         =   "PROMAT07.frx":2680
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Configuración - Valores por Omisión"
         Top             =   360
         Width           =   660
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   180
         ScaleHeight     =   585
         ScaleWidth      =   2055
         TabIndex        =   10
         Top             =   360
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
            Picture         =   "PROMAT07.frx":2AC2
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
               Name            =   "Arial"
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
         Left            =   840
         TabIndex        =   44
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
         Left            =   2070
         TabIndex        =   38
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label Label121 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "hs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2805
         TabIndex        =   37
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
         Left            =   1935
         TabIndex        =   36
         Top             =   1155
         Width           =   120
      End
      Begin VB.Label Label123 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ult.Cálc:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -240
         TabIndex        =   35
         Top             =   1275
         Width           =   960
      End
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   3570
      TabIndex        =   50
      Top             =   0
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Command12"
      Height          =   330
      Left            =   315
      TabIndex        =   49
      Top             =   105
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0DFC0&
      Caption         =   "Programa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1640
      Left            =   2100
      TabIndex        =   39
      Top             =   160
      Width           =   1380
      Begin VB.PictureBox Picture7 
         Height          =   495
         Left            =   360
         ScaleHeight     =   435
         ScaleWidth      =   915
         TabIndex        =   66
         Top             =   960
         Width           =   975
         Begin VB.Label Label8 
            BorderStyle     =   1  'Fixed Single
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
            Height          =   460
            Left            =   460
            TabIndex        =   68
            ToolTipText     =   "Doble-Click para Cambiar Modalidad de Programación"
            Top             =   -20
            Visible         =   0   'False
            Width           =   470
         End
         Begin VB.Label Label9 
            BorderStyle     =   1  'Fixed Single
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
            Height          =   460
            Left            =   -20
            TabIndex        =   67
            ToolTipText     =   "Doble-Click para Cambiar Modalidad de Programación"
            Top             =   -20
            Visible         =   0   'False
            Width           =   470
         End
      End
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
         Left            =   720
         Picture         =   "PROMAT07.frx":2DCC
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Listado General del Programa de Entregas"
         Top             =   240
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
         Left            =   120
         Picture         =   "PROMAT07.frx":30D6
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Consulta Programa de Entregas por Artículo (Pantalla)"
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0DFC0&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1640
      Left            =   7080
      TabIndex        =   3
      Top             =   160
      Width           =   3825
      Begin VB.CommandButton Command14 
         Caption         =   "Command14"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   1200
         Visible         =   0   'False
         Width           =   135
      End
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
         Left            =   3000
         TabIndex        =   48
         Top             =   1200
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0DFC0&
         BorderStyle     =   0  'None
         Height          =   450
         Left            =   240
         TabIndex        =   45
         Top             =   1100
         Width           =   2700
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0DFC0&
            Caption         =   "Impresión"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1320
            TabIndex        =   47
            Top             =   140
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0DFC0&
            Caption         =   "Pantalla"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   46
            Top             =   140
            Value           =   -1  'True
            Width           =   1170
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   150
         ScaleHeight     =   585
         ScaleWidth      =   3465
         TabIndex        =   4
         Top             =   420
         Width           =   3525
         Begin VB.CommandButton AGenRep 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2880
            Picture         =   "PROMAT07.frx":33E0
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
            Top             =   0
            Width           =   600
         End
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
            Left            =   1750
            Picture         =   "PROMAT07.frx":36EA
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
            Left            =   2330
            Picture         =   "PROMAT07.frx":3B2C
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
            Picture         =   "PROMAT07.frx":3E36
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   600
            TabIndex        =   8
            Top             =   120
            Width           =   1170
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0DFC0&
      Caption         =   "Actualización de Programa de Entregas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   120
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   630
         Width           =   730
      End
      Begin VB.CommandButton Command4 
         Height          =   670
         Left            =   9975
         Picture         =   "PROMAT07.frx":4140
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   315
         Width           =   600
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0DFC0&
         BorderStyle     =   0  'None
         Height          =   350
         Left            =   5565
         TabIndex        =   27
         Top             =   220
         Width           =   4290
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
            Height          =   180
            Index           =   11
            Left            =   3600
            TabIndex        =   33
            Top             =   120
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
            Height          =   180
            Index           =   10
            Left            =   2880
            TabIndex        =   32
            Top             =   120
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
            Height          =   180
            Index           =   9
            Left            =   2160
            TabIndex        =   31
            Top             =   120
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
            Height          =   180
            Index           =   8
            Left            =   1420
            TabIndex        =   30
            Top             =   120
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
            Height          =   180
            Index           =   7
            Left            =   705
            TabIndex        =   29
            Top             =   120
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
            Height          =   180
            Index           =   6
            Left            =   0
            TabIndex        =   28
            Top             =   120
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
         Picture         =   "PROMAT07.frx":444A
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   600
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
         TabIndex        =   20
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
         TabIndex        =   18
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
         TabIndex        =   34
         Top             =   400
         Width           =   3270
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0DFC0&
      Caption         =   "Programa de Entregas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
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
         Height          =   3210
         Left            =   105
         TabIndex        =   2
         Top             =   1320
         Width           =   11460
      End
      Begin VB.PictureBox Picture1 
         Height          =   975
         Left            =   120
         ScaleHeight     =   915
         ScaleWidth      =   11400
         TabIndex        =   13
         Top             =   315
         Width           =   11460
         Begin VB.CommandButton Command3 
            Height          =   330
            Left            =   9135
            Picture         =   "PROMAT07.frx":4754
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   0
            Width           =   540
         End
         Begin VB.CommandButton Command2 
            Height          =   330
            Left            =   5670
            Picture         =   "PROMAT07.frx":4A5E
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   0
            Width           =   540
         End
         Begin VB.Label INISEM 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   6225
            TabIndex        =   65
            Top             =   675
            Width           =   645
         End
         Begin VB.Label INISEM 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   5520
            TabIndex        =   64
            Top             =   675
            Width           =   645
         End
         Begin VB.Label INISEM 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   6945
            TabIndex        =   63
            Top             =   675
            Width           =   645
         End
         Begin VB.Label INISEM 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   7680
            TabIndex        =   62
            Top             =   675
            Width           =   645
         End
         Begin VB.Label INISEM 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   8400
            TabIndex        =   61
            Top             =   675
            Width           =   645
         End
         Begin VB.Label INISEM 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Left            =   9120
            TabIndex        =   60
            Top             =   675
            Width           =   645
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
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   0
            TabIndex        =   14
            Top             =   350
            Width           =   11460
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "PROMAT07.frx":4D68
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "PROMAT07.frx":4F9C
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
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
      Picture         =   "PROMAT07.frx":5016
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
      TabIndex        =   42
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
   Begin VB.Menu OpcionesVarias 
      Caption         =   "Opciones Varias"
      Begin VB.Menu ImportaPrograma 
         Caption         =   "Importador de Programa de Entregas"
      End
   End
End
Attribute VB_Name = "PROMAT07"
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

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

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
   If DERACCE%("STPMRP", "Configuracion de Funcionamiento de MRP") = 2 Then
      OPPROMRP07.Show 1
   End If
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
    For i& = 1 To ULTREG&("SUCURSAL")
      X$ = REGLEIDO$("SUCURSAL", i&)
      If i& <= 62 Then
        If Asc(Left$(X$, 1)) > 32 Then
          NUSU% = i&
          COSU%(i&) = Asc(Mid$(X$, 64, 1))
          SUCU$(i&) = Left$(X$, 20)
        End If
      End If
    Next i&
    Call CIERRARCH("SUCURSAL")
    '
    Call COPYSTRU("SALDEPOS", "SALDEPO")
    '
    FIN& = NUMAS% + 1
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      SSQ$ = String$(255, 0) + Chr$(1)
      Call GRAREG("SALDEPO", SSQ$, i&)
    Next i&
    Call SETULTREG("SALDEPO", FIN&)
    '
    FIN& = ULTREG&("MOVISTO")
    NMS% = NUMAS%
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      MVS$ = REGLEIDO$("MOVISTO", i&)
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
119 Next i&
    '
    Call CIERRARCH("MOVISTO")
    Call CIERRARCH("SALDEPO")
    '
    FIN& = NUMAS%
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      XX$ = Left$(REGLEIDO$("SALDEPOS", i& + 1), 248)
      YY$ = Left$(REGLEIDO$("SALDEPO", i& + 1), 248)
      If XX$ <> YY$ Then
        CAX = 0: CAY = 0
        For JI% = 1 To 248 Step 4
          CAX = CAX + CVS(Mid$(XX$, JI%, 4))
          CAY = CAY + CVS(Mid$(YY$, JI%, 4))
        Next JI%
          If (Abs(CAX - CAY)) > 0.05 Then
            CIXI% = i&
            MsgBox i& & "   " & CODEXT$(CIXI%)
            GoTo 129
          End If
        'Next JI%
      End If
129 Next i&
    Call CIERRARCH("*.*")
    MsgBox "Control Completo"
    '
End Sub

Private Sub Command13_Click()
   Command3.Enabled = False
   If DERACCE%("COPYMRP", "Copia de MRP") <> 2 Then GoTo 99
     
     Call CARSELMESA
     
     AA$ = OBJPLA("COPIMESA", VDEF$)
     If AA$ = "" Then GoTo 99
      '
      Screen.MousePointer = 11
      '
      Call MESANO(TRIM(Mid(AA$, 17, 16)), DES2%, HAS2%) ' RANGO HASTA
      Call MESANO(TRIM(Mid(AA$, 1, 16)), DES1%, HAS1%) 'RANGO DESDE
      '
      DIFDIAS% = DIENTRE(DES1%, DES2%) 'DES2% - DES1% + 1
      JR& = 0
      APROGRA$ = "PROGENT"
      If MODOPLAN% = 1 Then APROGRA$ = "PROGFAB"

      
      'BORRAR REGISTROS
      For U& = 1 To ULTREG&(APROGRA$)
        X$ = REGLEIDO$(APROGRA$, U&)
        FECHA = CVI(Mid(X$, 3, 2))
        If FECHA >= DES2% Then
           If FECHA <= HAS2% Then
              GoTo 19
           End If
        End If
        JR& = JR& + 1
        Call GRAREG(APROGRA$, X$, JR&)
19    Next U&
      Call SETULTREG(APROGRA$, JR&)
      '
      'AGREGAR REGISTROS
      For U& = 1 To ULTREG&(APROGRA$)
         X$ = REGLEIDO$(APROGRA$, U&)
         FECHA = CVI(Mid(X$, 3, 2))
         '
         'SI esta dentro del Rango entonces se agrega
         If FECHA >= DES1% And FECHA <= HAS1% Then
            FECHA = DIASPOST(FECHA, DIFDIAS%)
            Call REPLA(X$, MKI(FECHA), 3, 2)
            Call REGAPP(APROGRA$, X$)
         End If
         '
      Next U&
      '
      Call PROMAT07.CARPROGENT
      '
      Call COMUNI("Proceso de Copia Finalizado")
      '
99 Screen.MousePointer = 1
   Command3.Enabled = True
End Sub

Private Sub Command14_Click()

     Call BORRAREGISTROS("OCOEMI", "COD_INTE <> 0", REG&)
     Dim Ocoemi As ADODB.Recordset
     Set Ocoemi = New ADODB.Recordset
     SQLX1$ = "SELECT * FROM OCOEMI WHERE COD_INTE = 0"
25   On Error Resume Next
     Ocoemi.Open SQLX1$, FLEXCONN, adOpenDynamic, adLockPessimistic
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
         GoTo 25
     End If
     On Error GoTo 0

     For U& = 1 To ULTREG&("PLACOMP")
       XX$ = REGLEIDO$("PLACOMP", U&)
       CII% = CVI(Left$(XX$, 2))
       If CII% = 346 Then
        A = B
       End If
       If CII% <> CIAX% Then
         CIAX% = CII%
         'Stakk = STODEPOS(CII%, DEPOSTAKK%): If Stakk < 0 Then Stakk = 0
         'Call REPLA(XX$, MKS$(Stakk), 29, 4)
         'Call REPLA(XX$, MKS$(STOMIN(CII%)), 55, 4)
         'Call REPLA(XX$, MKS$(LOTESTAN(CII%)), 59, 4)
       End If
       '
       LT% = LEAD_TIME%(CII%)
       FeNec = CVI(Mid$(XX$, 3, 2))
       FANT% = DIANTE(FeNec, LT%)
       'Si la fecha de solicitud ya ha pasado, entonces el sistema pone la del dia del calculo
       If FANT% < HOY(HO$) Then
          FANT% = HOY(HO$)
       End If
       Call REPLA(XX$, MKI$(FANT%), 63, 2)
       '
       'Se graba en OCOEMI
       CI% = CVI(Mid$(XX$, 1, 2))
       FF% = CVI(Mid$(XX$, 3, 2))
       PERIO$ = Mid$(XX$, 5, 16)
       CANTX = CVS(Mid$(XX$, 21, 4))
       PHB% = CVI(Mid$(XX$, 33, 2))
       COSUN = CVS(Mid(XX$, 35, 4))
       'Se Graba OCOEMI
       With Ocoemi
           .AddNew
           !COD_INTE = CI%
           !CANTIDAD = CANTX
           !Fentre_Sol = FECHATEX$(FF%)
           !Npro_Ocom = PHB%
           !Raso_prov = Left(rasocli$(PHB% * -1), 48)
           !Uni_Mas = Unimed$(CI%)
           !FechRecep = FECHATEX$(FANT%)
           !COSUNI = COSUN
           !PERIODO = PERIO$
           !CodiEmpr = CodiEmp%
           !StokAct = Stakk
           !STOMIN = STOMIN(CI%)
           !LOREPOS = LOTESTAN(CI%)
           .Update
         End With
       '
'AQUI PONER FECHA LIMITE EMISION O/COMPRA
'FECHA DE NECESIDAD EN 3,2
'EN 63,2 PONER MKI$(DIANTE(FECHA DE NECESIDAD, LEAD TIME))
'       Call GRAREG("PLACOMP", XX$, U&)
     Next U&
     Set Ocoemi = Nothing
     Call CIERRARCH("PLACOMP")
'     Call CARPROSEL

End Sub

Private Sub Command15_Click()
    Dim CADISP(32000) As Single
    '
    Command15.Enabled = False
    TTXX$ = "Cobertura Completa\Cobertura Inmediata"
    TTXX$ = TTXX$ + "\Resumen de Cobertura\Disponibilidad Multinivel"
    OP% = ALTERNA%(TTXX$)
    If OP% = 3 Then
         ANACOB07.Show 1
         GoTo 90
       ElseIf OP% = 4 Then
         Call DISPOMUL
         GoTo 90
    End If
    '
    Select Case OP%
      Case 1
        If Option1.Value = True Then
           If EXISTESQR_ENSQL%("COBECOMP") > 0 Then
             SQLX$ = "select top 1 ID  from coberfa with(nolock) order by ID desc"
             Set RS = READSET(SQLX$)
             IDX = XVALO(RS!ID)
             If IDX > 0 Then 'si no hay ningun registro en la tabla muestra el dat
               CODI$ = "00"
               Call FINAL("?COBECOM2(" & IDX & ";" & CODI$ & ")")
               Screen.MousePointer = 1
             Else: GoTo 20
             End If
           Else
20           Call CONECRUN("*COBERFA", "")
           End If
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
    Screen.MousePointer = 11

    Call ABREPEDCLI
    SQLX$ = " SELECT * FROM PEDDETA "
    SQLX$ = SQLX$ + "WHERE (Status >= 0 AND Status < 8) "   ' para que no este anulada ni cumplida ni pendiente de aprobacion
    SQLX$ = SQLX$ + "AND CantEnt < (CanPed - 0.005) "
    SQLX$ = SQLX$ + "AND VenRep = 'V' "
    SQLX$ = SQLX$ + "ORDER BY FeSolEnt DESC "
    '
    'Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim PEDTEMP As ADODB.Recordset
    Set PEDTEMP = New ADODB.Recordset
    PEDTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    For JL& = 1 To 32767 'PONGO EN 0 EL VECTOR
       TIENEPEDI%(JL&) = 0
    Next JL&
    '
    ' CARGA UN VECTOR CON LOS CLIENTES CON PEDIDOS PENDIENTES
    With PEDTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          NCLIE1% = !NROCLIE
          If NCLIE1% > 0 Then
             TIENEPEDI%(NCLIE1%) = 1 'MARCO LOS CLIENTES CON PEDIDOS
          End If
        .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    '
    For JL& = 1 To 32767
      If TIENEPEDI%(JL&) = 1 Then
         NCLIE2$ = FORMATNUM$(JL&, "F6.0")
         RASOCLIE2$ = rasocli$(JL&)
         Z$ = Space(128)
         Call REPLA(Z$, NCLIE2$, 1, 6)
         Call REPLA(Z$, RASOCLIE2$, 15, 64)
         With DISMULT07
          .LICOPEN.AddItem Z$
         End With
      End If
    Next JL&
    Screen.MousePointer = 1
    If DISMULT07.LICOPEN.ListCount > 0 Then
       DISMULT07.Show 1
       Call FRESHALL
       If DISMULT07.Label2.Caption <> "OK" Then
         Call COMUNI("Cancelado !!!")
         GoTo 9999
       End If
       '
       ' ACA VUELVE A PONER EN '0' EL VECTOR DE CLIENTES
       CORTEPED% = FECHANUM(DISMULT07.Label7)
       CORTECOM% = FECHANUM(DISMULT07.Label3)
       For JL& = 1 To 32767
         TIENEPEDI%(JL&) = 0
       Next JL&
       '
       ' ACA RE-CARGA EL VECTOR PONIENDO EN '1' LOS CLIENTES SELECCIONADOS
       TCLISEL$ = ""
       For JL& = 1 To DISMULT07.LICOPEN.ListCount
         If DISMULT07.LICOPEN.Selected(JL& - 1) = True Then
           ICLIE1& = Val(Left$(DISMULT07.LICOPEN.List(JL& - 1), 6))
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
     With PEDTEMP  ' ACA LO ABRE DE VUELTA PERO SIN CONSULTA
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           NCLIE1% = !NROCLIE
           If NCLIE1% > 0 Then
             If TIENEPEDI%(NCLIE1%) > 0 Then 'MARCO LOS CLIENTES CON PEDIDOS
               CAPEN = !CanPed - !CANTENT
               '
               If CAPEN > 0 Then
                 If CORTEPED% < 1 Or !FeSolEnt <= CVDAT(CORTEPED%) Then ' OJO FECHA DE CORTE
                    CIIX% = !CODIINT
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
       On Error GoTo 0
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
     For i& = 1 To FIN&
       Call TRACE(20, i&, FIN&)
       CIIX% = i&
       If ACTI%(CIIX%) > 0 Then
         EXIS(i&, 1) = STOCKACT(CIIX%)
         EXIS(i&, 5) = STOMIN(CIIX%)
       End If
     Next i&
     '
     Call ABRECOMPRAS
     SQLX$ = " SELECT * FROM OCOMDETA "
      'SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + "WHERE Stat_Ocom < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND  Stat_Ocom <> -1 " 'Para que no  hayan O/C's sin aprobar
     SQLX$ = SQLX$ + "AND Cant_Rec < (Cant_Ocom - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY Nume_Ocom ASC "
     '
     'Set OcoTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     Dim OCOTEMP As ADODB.Recordset
     Set OCOTEMP = New ADODB.Recordset
     OCOTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     With OCOTEMP
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
     OCOTEMP.Close
     Set OCOTEMP = NOTTHING
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
     'Set AtenTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     Dim AtenTemp As ADODB.Recordset
     Set AtenTemp = New ADODB.Recordset
     AtenTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     With AtenTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
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
       On Error GoTo 0
     End With
     AtenTemp.Close
     Set AtenTemp = Nothing
     '\\\OT-06-0149-WAR-FIN///
     Call ABRESTOCKS
     SQLX$ = " SELECT * FROM BOLRECEP "
     SQLX$ = SQLX$ + " WHERE Status = 0 "
     '
     'Set COBRECEPTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     
     Dim COBRECEPTMP As ADODB.Recordset
     Set COBRECEPTMP = New ADODB.Recordset
     COBRECEPTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

     With COBRECEPTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
        On Error GoTo 0
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
       On Error GoTo 0
     End With
     COBRECEPTMP.Close
     Set COBRECEPTMP = Nothing
     '
50   J& = 0
     NVII& = NUMAS%
     For i& = 1 To NVII&
       Call TRACE(20, i&, NVII&)
       CIIX% = i&
       If STATUIT%(CIIX%) > 0 Then
         Code$ = CODEXT$(CIIX%)
         If ACTI%(CIIX%) > 0 Then
         'If CODE$ >= COD1$ Then
          'If CODE$ <= COD2$ Then
            'If TMT% = 0 Or TIMATE%(CIIX%) = TMT% Then
              X$ = REGBLAN$("STODISP")
              Call REPLA(X$, MKI$(CIIX%), 1, 2)
              Call REPLA(X$, Chr$(TIMATE%(CIIX%)), 3, 1)
              SUMA = 0: HAYMO% = 0
              For k% = 1 To 6
                SIG% = 1: If k% >= 4 Then SIG% = (-1)
                SUMA = SUMA + SIG% * EXIS(CIIX%, k%)
                If Abs(EXIS(CIIX%, k%)) >= 0.05 Then HAYMO% = 1
                Call REPLA(X$, MKS$(EXIS(CIIX%, k%)), 4 * k% + 1, 4)
              Next k%
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
    Next i&
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
    Screen.MousePointer = 1
    Exit Sub
    '
9999 End Sub
'
Private Sub Command2_Click()
   MESPAN$ = MESANO$(Label2, DES%, Has%)
   FEX = DES%
   FEPANUE% = DIANTE(FEX)
   If FEPANUE% < COMESACT% Then Exit Sub
   Label2 = METEXTO$(FEPANUE%)
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    OPX% = COMALTER%("Formato de Salida:\\Tabla\Listado")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    If OPX% = 1 Then
      Call LISTABPLAN
      GoTo 99
    End If
    APROGRA$ = "*LPROGEN"
    If MODOPLAN% = 1 Then APROGRA$ = "*LPROGFA"
    Call CONECRUN(APROGRA$, "")
99  Command25.Enabled = True
End Sub

Private Sub Command3_Click()
   MESPAN$ = MESANO$(Label2, DES%, Has%)
   FEX = Has%
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
     FORCANEC07.Show 1
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
          Call FINAL("*LPLAFAB")
    End If
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    AXXX$ = "": ARPLA$ = "PLACOMP"
'    If InStr(EPAC$, "FERVI") > 0 Then
'       AXXX$ = "Q"
'       ARPLA$ = "PLACOMQ"
'    End If
    MODOLIS% = 0
    If EXISTE%(LUDAT$ + "PLACOMC.RFS") > 0 Then MODOLIS% = 1
    '
5   OP% = ALTERNA%(" Cambio de Proveedor\Consultas/Listados")
    If OP% < 1 Or OP% > 2 Then GoTo 99
    If OP% = 1 Then
      PRORECO07.Show 1
      GoTo 5
    End If
    
    If Option1.Value = True Then
        Call FINAL("*CPLACOM" + AXXX$)
      Else
8       OPX% = COMALTER%("Elija Secuencia / Agrupamiento\\Por Proveedor\Por Artículo")
        If OPX% < 1 Or OPX% > 2 Then GoTo 5
        If OPX% = 1 Then
            Call FILESORT(ARPLA$, "", 8, 8, "ASC")
          ElseIf OPX% = 2 Then
            Call FILESORT(ARPLA$, "", 1, 3, "ASC")
        End If
        '
        TNT% = 0
        If AXXX$ = "" Then
10        'TMT% = Asc(Left$(REGSEL$("TATIMAT"), 1))
          '\\\OT-06-0130-WAR-05/04/06///
          'Call SELECHO("FECHA/Fecha de Corte")
          'VDEVU$ = VALACT1$("FECHA")
          
          'FLIM1% = FECHANUM(VDEVU$)
         ' OBX$ = OBJPLA$("FILNECOM", VDEF$)
         ' AA = Asc(Left$(OBX$, 1))
         ' TMT% = 0: If Asc(Left$(OBX$, 1)) > 0 Then TMT% = Asc(Left$(OBX$, 1))
         ' FNECE% = CVI(Mid$(OBX$, 2, 2))
         ' FLIM1% = CVI(Mid$(OBX$, 4, 2))

          'If TRIM$(VDEVU$) = "//" Then FLIM1% = 32767
          '\\\OT-06-0130-WAR-FIN///
          
          SELTIMA.Show 1
          If SELTIMA.LICOPEN.ListCount < 1 Then
             Unload SELTIMA
             GoTo 8
          End If
          '
          FNECE% = FECHANUM(SELTIMA.Label7)
          FLIM1% = FECHANUM(SELTIMA.Label3)
          TOFECHA% = 0
          If SELTIMA.Check2.Value = 1 Then TOFECHA% = 1

          ' VERIFICA SI EXISTE ALGUNO SELECCIONADO
          TOMATE% = 0: TMT% = 0
          If SELTIMA.Check1.Value = 1 Then TMT% = 1: TOMATE% = 1
          If TOMATE% < 1 Then
             For TIMA1% = 1 To SELTIMA.LICOPEN.ListCount
               If SELTIMA.LICOPEN.Selected(TIMA1% - 1) = True Then
                 TMT% = 1: Exit For
               End If
             Next TIMA1%
          End If
          '
          Screen.MousePointer = 11
          Call COPYSTRU("PLACOMP", "PLACOMF")
          JJ& = 0: FIN& = ULTREG&("PLACOMP")
          For U& = 1 To FIN&
            Call TRACE(20, U&, FIN&)
            X$ = REGLEIDO$("PLACOMP", U&)
            CIXI% = CVI(Left$(X$, 2))
            F1% = CVI(Mid$(X$, 3, 2))
            F2% = CVI(Mid$(X$, 63, 2))
            '
            If TMT% = 1 Then
               TIPOSELECT% = 0
               For TIMA1% = 1 To SELTIMA.LICOPEN.ListCount
                  If SELTIMA.LICOPEN.Selected(TIMA1% - 1) = True Then
                    NROTIMA% = XVALO(Mid$(SELTIMA.LICOPEN.List(TIMA1% - 1), 1, 3))
                    If NROTIMA% = TIMATE%(CIXI%) Then
                      TIPOSELECT% = 1: Exit For
                    End If
                  End If
               Next TIMA1%
            End If
            '
            'If (TIMATE%(CIXI%) = TMT% Or TMT% = 0) Then
            If (TIPOSELECT% = 1 Or TMT% = 0 Or TOMATE% = 1) Then
              If F1% <= FLIM1% Or FLIM1% = 0 Or TOFECHA% = 1 Then '\\\OT-06-0130-WAR-05/04/06///
                If F2% <= FNECE% Or FNECE% = 0 Or TOFECHA% = 1 Then
                  JJ& = JJ& + 1
                  Call GRAREG("PLACOMF", X$, JJ&)
                End If
              End If
            End If
          Next U&
          Call SETULTREG("PLACOMF", JJ&)
          Screen.MousePointer = 1
          '
          If JJ& < 1 Then
            Call MENSERR(24, "Sin Necesidades\para la Selección.")
            GoTo 99
          End If
          Call SETULTREG("PLACOMF", JJ&)
          ARENTRA$ = "PLACOMF"
          If MODOLIS% = 1 Then GoTo 60
          '
          CANTISEL& = SELTIMA.LICOPEN.SelCount  ' SI SE SELECCINO MAS DE UN TIPO DE MATERIAL
          If CANTISEL& > 1 Then
            TITUTIPO$ = "VARIOS"
          Else
            TITUTIPO$ = ECOARCH$("TATIMAT", Chr$(NROTIMA%))
          End If
          Call HEADERS("PLACOMF", "")
          If TMT% > 0 And TOMATE% < 1 Then Call HEADERS("PLACOMF", "Tipo de Material: " + TITUTIPO$)
          If FLIM1% > 0 And TOFECHA% < 1 Then Call HEADERS("PLACOMF", "Mat.Necesario Antes de: " + FECHATEX(FLIM1%))
          If FNECE% > 0 And TOFECHA% < 1 Then Call HEADERS("PLACOMF", "O/C's a Emitir Antes de: " + FECHATEX(FNECE%))
          Call FINAL("*LPLACOF")
          GoTo 99
        End If
        Call FINAL("*LPLACOM" + AXXX$)
        GoTo 99
    End If
    GoTo 99
    '
60  Dim CANCO(32767), CANRE(32767), CANPRO(32767)
    NMS% = NUMAS%
    Screen.MousePointer = 11

'    FIN& = ULTREG&("OCOMDETA")
'    JJ& = 0
'    For U& = 1 To FIN&
'      Call TRACE(20, U&, FIN&)
'      XXX$ = REGLEIDO$("OCOMDETA", U&)
'      If Asc(Mid$(XXX$, 124, 1)) < 8 Then ' ordenes abiertas Y NO CUMPLIDAS
'        CAPEN = CVS(Mid$(XXX$, 55, 4)) - CVS(Mid$(XXX$, 99, 4))
'        If CAPEN < 0.05 Then CAPEN = 0
'        CIXI% = CVI(Mid$(XXX$, 9, 2))
'        If CIXI% > 0 Then
'          If CIXI% <= NMS% Then
'            CANCO(CIXI%) = CANCO(CIXI%) + CAPEN
'          End If
'        End If
'      End If
'    Next U&
    '
    '\\\OT-08-00906-OD-29/09/08///
    '
    SQLX$ = "SELECT * FROM OCOMDETA"
    SQLX$ = SQLX$ + " WHERE Stat_Ocom < 8"
    SQLX$ = SQLX$ + " AND Cod_Inte > 0"
    SQLX$ = SQLX$ + " ORDER BY Cod_Inte"
    
    Dim OCOMTMP As ADODB.Recordset
    Set OCOMTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With OCOMTMP
      Do While Not .EOF
        CAPEN# = XVALO(!Cant_Ocom) - XVALO(!Cant_Rec)
        CIXI% = XVALO(!COD_INTE)
        If CAPEN# < 0.05 Then CAPEN# = 0
        If CIXI% > 0 Then
            If CIXI% <= NMS% Then
              CANCO(CIXI%) = CANCO(CIXI%) + CAPEN#
            End If
        End If
        .MoveNext
      Loop
    End With
    '
    OCOMTMP.Close
    Set OCOMTMP = Nothing
    '
    '\\\OT-08-00906-OD-FIN//
    
'    FIN& = ULTREG&("BOLREPEN")
'    JJ& = 0
'    For U& = 1 To FIN&
'      Call TRACE(20, U&, FIN&)
'      XXX$ = REGLEIDO$("BOLREPEN", U&)
'      CAPEN = CVD(Mid$(XXX$, 89, 8))
'      If CAPEN < 0.05 Then CAPEN = 0
'      CIXI% = CVI(Mid$(XXX$, 83, 2))
'      If CIXI% > 0 Then
'        If CIXI% <= NMS% Then
'          CANRE(CIXI%) = CANRE(CIXI%) + CAPEN
'        End If
'      End If
'    Next U&
    '\\\OT-08-00906-OD-29/09/08///
    SQLX$ = "SELECT * FROM BOLRECEP"
    SQLX$ = SQLX$ + " WHERE Status = 0"
    SQLX$ = SQLX$ + " AND Cod_Inte > 0"
    SQLX$ = SQLX$ + " ORDER BY Cod_Inte"
    '
    Dim BOLRETMP As ADODB.Recordset
    Set BOLRETMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With BOLRETMP
      Do While Not .EOF
        CAPEN# = XVALO(!CANTNOM)
        CIXI% = XVALO(!COD_INTE)
        If CAPEN# < 0.05 Then CAPEN# = 0
            If CIXI% <= NMS% Then
              CANRE(CIXI%) = CANRE(CIXI%) + CAPEN#
            End If
        .MoveNext
      Loop
    End With
    BOLRETMP.Close
    Set BOLRETMP = Nothing
    '\\\OT-08-00906-OD-FIN///
    '
    FIN& = ULTREG&(ARENTRA$)
    For U& = 1 To FIN&
      XX$ = REGLEIDO$(ARENTRA$, U&)
      CIXI% = CVI(Mid$(XX$, 1, 2))
      CANPROGRA# = CVS(Mid$(XX$, 21, 4))
      CANPRO(CIXI%) = CANPRO(CIXI%) + CANPROGRA#
    Next U&
    '
    REBLA$ = REGBLAN$("PLACOMC")
    CIA% = 0
    JJ& = 0
    PASO% = 0
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
          SALDOPEN# = CANPRO(CIXI%) + (CVS(Mid$(XX$, 29, 4)) + CANCO(CIXI%) + CANRE(CIXI%))
          Call REPLA(YY$, MKS$(SALDOPEN#), 41, 4)
          
          'GRABA EL RESTO DE LAS ORDENES DE COMPRA PENDIENTES
          If PASO% = 2 Then
            With OCOMTMP1
              Do While Not .EOF
                CAPEN# = XVALO(!Cant_Ocom) - XVALO(!Cant_Rec)
                If CAPEN# > 0.05 Then
                    FES% = CVINT%(!Fentre_Sol)
                    NOCOM& = XVALO(!NUME_OCOM)
                    Call REPLA(ZZ$, Space(69), 21, 68)
                    Call REPLA(ZZ$, MKI$(0), 45, 2)
                    Call REPLA(ZZ$, MKS$(NOCOM&), 89, 4)
                    Call REPLA(ZZ$, MKI$(FES%), 93, 2)
                    Call REPLA(ZZ$, MKS$(CAPEN#), 95, 4)
                    JJ& = JJ& + 1
                    Call GRAREG("PLACOMC", ZZ$, JJ&)
                End If
                .MoveNext
              Loop
            End With
          End If
          
          On Error Resume Next
          OCOMTMP1.Close
          Set OCOMTMP1 = Nothing
          On Error GoTo 0
          
          'CREAR RECORDSET CON LAS OC PENDIENTES PARA CIXI%
          PASO% = 0
          If PASO% = 0 Then
            SQLX$ = "SELECT * FROM OCOMDETA WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE Stat_Ocom < 8"
            SQLX$ = SQLX$ + " AND Cod_Inte =" & CIXI%
            SQLX$ = SQLX$ + " AND (Cant_Ocom - Cant_Rec)>0.05"
            SQLX$ = SQLX$ + " ORDER BY Fentre_Sol"
            Set OCOMTMP1 = READSET(SQLX$)
            If Not (OCOMTMP1.EOF And OCOMTMP1.BOF) Then PASO% = 1
          End If
          '
        End If
        
        Call REPLA(YY$, MKS$(CANCO(CIXI%)), 33, 4)
        Call REPLA(YY$, MKS$(CANRE(CIXI%)), 37, 4)
        If InStr(EPAC$, "FERVI") < 1 Then
            Call REPLA(YY$, MKS$(CVS(Mid$(XX$, 29, 4)) + CANCO(CIXI%) + CANRE(CIXI%)), 41, 4)
        End If
        '
        ' AGREGA CONSUMO ESTADISTICO MINIMO, MAXIMO Y PROMEDIO PARA DOSIVAC
        If InStr(EPAC$, "DOSIVAC") > 0 Then
          Codigo$ = CODEXT$(CIXI%)
          REG$ = RECFILT$("STATSTOK", 1, Codigo$)
          If CVS(REG$) > 0 Then
            ZZ$ = REGLEIDO$("STATSTOK", CVS(REG$))
            CONSUMINIMO# = CVS(Mid$(ZZ$, 65, 4))
            CONSUMAXIMO# = CVS(Mid$(ZZ$, 69, 4))
            CONSUPROMEDIO# = CVS(Mid$(ZZ$, 73, 4))
            Call REPLA(YY$, MKS$(CONSUMINIMO#), 77, 4)
            Call REPLA(YY$, MKS$(CONSUMAXIMO#), 81, 4)
            Call REPLA(YY$, MKS$(CONSUPROMEDIO#), 85, 4)
          End If
        End If
      End If
    
      'ESCRIBE EN CADA REGISTRO DE NECESIDAD LAS OC PENDIENTES
      If InStr(EPAC$, "FERVI") > 0 And PASO% > 0 Then
            With OCOMTMP1
              If Not .EOF Then
                CAPEN# = XVALO(!Cant_Ocom) - XVALO(!Cant_Rec)
                If CAPEN# > 0.05 Then
                      FES% = CVINT%(!Fentre_Sol)
                      NOCOM& = XVALO(!NUME_OCOM)
                      Call REPLA(YY$, MKS$(NOCOM&), 89, 4)
                      Call REPLA(YY$, MKI$(FES%), 93, 2)
                      Call REPLA(YY$, MKS$(CAPEN#), 95, 4)
                End If
                .MoveNext
              End If
            End With
            PASO% = 2
      End If
            
      JJ& = JJ& + 1
      Call GRAREG("PLACOMC", YY$, JJ&)
      
      ZZ$ = YY$
    Next U&
    '
    Call SETULTREG("PLACOMC", JJ&)
    Call CIERRARCH("PLACOMC")
    '
    If OPX1% = 1 Then
         Call FILESORT("PLACOMC", "", 11, 11)
       ElseIf OPX1% = 2 Then
         Call FILESORT("PLACOMC", "", 1, 3)
    End If
    Screen.MousePointer = 1
    
    ARCHCOM$ = "PLACOMC"
    If InStr(EPAC$, "FERVI") > 0 Then ARCHCOM$ = "PLACOMF"
    Call FINAL("*" & ARCHCOM$)
    GoTo 99
    '
    
99  Command9.Enabled = True
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Refresh
     Openforms = DoEvents
   Wend
End Sub

Private Sub Form_Load()
   '
   Call CENTRAFORM(Me)
'   Show
'   Refresh
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   Call CONAPLI(APLINVO$, TERMINA%) ' Busca la aplicacion a mostrar
   If TERMINA% > 0 Then Call FINAL("")
   '
'   EPAC$ = TRIM$(EMPREAC$)
'   If EPAC$ <> "" Then
'     '
'10   Caption = Caption + "  -  " + EPAC$
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
10     Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    '
    '
     Call GRATITFOR(Caption)
     '
   '
   Openforms = DoEvents
   '
   OPCIPLAN% = 1
   ' SE COMENTA LAS SIGUIENTES DOS LINEAS DE CODIGO PARA DEJAR LAS MODALIDADES
   ' DE CALCULO POR FABRICACION O VENTAS PARA TODOS SIN EXCEPCION
   'If InStr(UCase$(EPAC$), "DOSIVA") > 0 Then OPCIPLAN% = 1
   'If InStr(UCase$(EPAC$), "FLEXOFT") > 0 Then OPCIPLAN% = 1
   Call GRATITFOR(Caption)
   '
   MODOPLAN% = 0
   If OPCIPLAN% = 1 Then
     If Control("PROMRP", "MODOPLAN") = "FABRICA" Then
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
   UTILIZA_SKIN% = 1
   Call APLICACIONSKINS(Me, Picture4)
   Show

   FECHOR$ = Control$("CANECOM", "FECHOCAL")
   Label7.Caption = Left$(FECHOR$, 8): Label7.Refresh
   Label95.Caption = Mid$(FECHOR$, 11, 5): Label95.Refresh
   '
   PRISEMIN% = 53
   If Val(Control$("", "PRISEMIN")) = 1 Then PRISEMIN% = 1
   '
   Call CARPROGENT
   '
   Call CREATABLA_COBERFA

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

Private Sub Frame7_DblClick()
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "GUIDI") > 0 Then
      FORACPLAN07.Show 1
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
   For i& = 1 To FIN&
     Call TRACE(20, i&, FIN&)
     P00$ = REGLEIDO$(APROGRA$, i&)
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
   Next i&
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
          codex$ = CODEXT$(CIXI%)
          CACOM = CANT(UI&)
          '
          RFF$ = RECFILT$("SECOPER", 1, MKI$(CIXI%))
          IOP% = 0
          For UJ& = 1 To Len(RFF$) Step 4
            i& = CVS(Mid$(RFF$, UJ&, 4))
            If i& > 0 Then
              If i& <= URE& Then
                SCP$ = REGLEIDO$("SECOPER", i&)
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

Private Sub ImportaPrograma_Click()

    Call CONECRUN("ARCHIG07/IMPOMRP", "Importador de Programa de Entregas")

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
   MESPAN$ = MESANO$(Label2, DES%, Has%)
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
     For i& = 1 To FIN&
       Call TRACE(20, i&, FIN&)
       P00$ = REGLEIDO$(APROGRA$, i&)
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
               If CODIPRO%(KKI%) = CI% Then GoTo 7
             Next KKI%
             CACODIP% = CACODIP% + 1
             CODIPRO%(CACODIP%) = CI%
             KKI% = CACODIP%
             '
7            If FENTRE% < COMESACT% Then
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
129  Next i&
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
       TTXX$ = TTXX$ + AJUSTI$(Unimed$(CIXI%), 4)
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
         TEXTO1$ = MAPIMessages1.MsgNoteText
         List2.Clear
         While Len(TEXTO1$) > 0
           PPXX% = InStr(TEXTO1$, Chr$(13) + Chr$(10))
           If PPXX% < 1 Then PPXX% = 1 + Len(TEXTO1$)
           TTXX$ = Left$(TEXTO1$, PPXX% - 1)
           TEXTO1$ = Mid$(TEXTO1$, PPXX% + 2)
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
               Call REPLA(OCDT$, Unimed$(CIXI%), 11, 4)
               Call REPLA(OCDT$, MKS$(CASOL), 51, 4)
               Call REPLA(OCDT$, MKS$(CASOL), 55, 4)
               Call REPLA(OCDT$, Unimed$(CIXI%), 59, 4)
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
Sub CONAPLI(APLINVO$, TERMINA%)
   TERMINA% = 1
   Select Case APLINVO$
       '
       'TRANSACCIONES
       '
       Case "RECALNEC"
           Call Command7_Click                 ' RE-CALCULO DE NECESIDADES
       '
       'CONSULTAS Y LISTADOS
       '
       Case "ANACOBER"
           Option2.Value = True
           Call Command15_Click               ' ANALISIS DE COBERTURA
       Case "NECECOMPRA"
           Option2.Value = True
           Call Command9_Click                ' NECESIDADES DE COMPRAS
       Case "NECEFABRI"
           Option2.Value = True
           Call Command8_Click                ' NECESIDADES DE FABRICACION
       Case "PLAFABART"
           Call Command6_Click                ' PLAN DE FABRICACION POR ARTICULO
       Case "PLAMAEFAB"
           Call Command25_Click               ' PLAN MAESTRO DE FABRICACION
       '
       'CONFIGURACION MRP II
       '
       Case "CONGENMRP"
           Call Command10_Click               ' CONFIGURACION DE FUNCIONAMIENTO MRP II
       Case Else
           TERMINA% = 0
   End Select
End Sub


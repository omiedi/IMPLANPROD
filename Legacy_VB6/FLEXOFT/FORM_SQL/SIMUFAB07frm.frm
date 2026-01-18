VERSION 5.00
Begin VB.Form SIMUFAB07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simulación de Restricciones de Stock"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11880
   Icon            =   "SIMUFAB07frm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   11880
   Begin VB.Frame Frame17 
      BackColor       =   &H00CDDADA&
      Caption         =   "STOCKS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   975
      Left            =   5880
      TabIndex        =   51
      Top             =   570
      Width           =   2370
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "No Considera Stocks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   165
         TabIndex        =   53
         Top             =   650
         Width           =   2130
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Stock NO Reservado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   220
         Left            =   165
         TabIndex        =   52
         Top             =   300
         Width           =   2130
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00CDDADA&
      Caption         =   "PROYECCION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   8925
      TabIndex        =   47
      Top             =   1590
      Visible         =   0   'False
      Width           =   1845
      Begin VB.Frame Frame16 
         BackColor       =   &H00C0E0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   330
         Left            =   735
         TabIndex        =   50
         Top             =   1155
         Width           =   3465
      End
      Begin VB.CommandButton Command23 
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
         Left            =   945
         Picture         =   "SIMUFAB07frm.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Genera Tabla de Proyección de Compras"
         Top             =   375
         Width           =   600
      End
      Begin VB.CommandButton Command9 
         Height          =   600
         Left            =   315
         Picture         =   "SIMUFAB07frm.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Genera Tabla de Ordenes de Fabricación a Emitir"
         Top             =   375
         Width           =   600
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00CDDADA&
      Caption         =   "FABRICACION A EMITIR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   5880
      TabIndex        =   39
      Top             =   2760
      Width           =   4890
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         IntegralHeight  =   0   'False
         Left            =   105
         TabIndex        =   41
         Top             =   550
         Width           =   4680
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   330
         Left            =   360
         TabIndex        =   40
         Top             =   2340
         Width           =   3465
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Codigo                    Bruto     En Curso         Neto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   105
         TabIndex        =   42
         Top             =   315
         Width           =   4680
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00CDDADA&
      Caption         =   "RESTRICCIONES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   105
      TabIndex        =   33
      Top             =   6000
      Width           =   5650
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1470
         Left            =   105
         TabIndex        =   35
         Top             =   555
         Width           =   5400
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   330
         Left            =   360
         TabIndex        =   34
         Top             =   2340
         Width           =   3465
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  Codigo                 Descripcion                 Uso    Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   105
         TabIndex        =   36
         Top             =   315
         Width           =   5400
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00CDDADA&
      Caption         =   "INICIO / ULTIMO CALCULO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   8400
      TabIndex        =   28
      Top             =   130
      Width           =   2385
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
         Height          =   555
         Left            =   1440
         Picture         =   "SIMUFAB07frm.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Re-Cálculo de Capacidad"
         Top             =   300
         Width           =   720
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   330
         Left            =   360
         TabIndex        =   29
         Top             =   2340
         Width           =   3465
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   56
         Top             =   660
         Width           =   600
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   720
         TabIndex        =   54
         Top             =   490
         UseMnemonic     =   0   'False
         Width           =   495
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha  y Hora"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   105
         TabIndex        =   32
         Top             =   930
         Width           =   600
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   720
         TabIndex        =   31
         Top             =   1050
         Width           =   795
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1560
         TabIndex        =   30
         Top             =   1050
         Width           =   675
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   4950
      Left            =   10920
      ScaleHeight     =   4920
      ScaleWidth      =   1350
      TabIndex        =   23
      Top             =   -50
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   150
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   57
         Top             =   4440
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   58
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command3 
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
         Height          =   750
         Left            =   150
         Picture         =   "SIMUFAB07frm.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Maestro de Items de Stock"
         Top             =   3570
         Width           =   650
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
         Left            =   150
         Picture         =   "SIMUFAB07frm.frx":0D72
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   650
      End
      Begin VB.CommandButton Command28 
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
         Height          =   690
         Left            =   105
         Picture         =   "SIMUFAB07frm.frx":0EBC
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   8730
         Width           =   650
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00CDDADA&
      Caption         =   "NECESIDADES DE COMPONENTES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      Left            =   5880
      TabIndex        =   11
      Top             =   4920
      Width           =   5940
      Begin VB.ListBox List4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2190
         ItemData        =   "SIMUFAB07frm.frx":11C6
         Left            =   90
         List            =   "SIMUFAB07frm.frx":11C8
         TabIndex        =   14
         Top             =   660
         Width           =   5760
      End
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   90
         ScaleHeight     =   405
         ScaleWidth      =   5730
         TabIndex        =   12
         Top             =   360
         Width           =   5760
         Begin VB.CommandButton Command8 
            Caption         =   "  Compras"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4500
            TabIndex        =   37
            Top             =   0
            Width           =   1000
         End
         Begin VB.CommandButton Command10 
            Caption         =   "      Falta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3550
            TabIndex        =   27
            Top             =   0
            Width           =   950
         End
         Begin VB.CommandButton Command6 
            Caption         =   "    Stock"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2600
            TabIndex        =   18
            Top             =   0
            Width           =   950
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Q.Neces"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1650
            TabIndex        =   17
            Top             =   0
            Width           =   950
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Codigo               "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   1650
         End
         Begin VB.Frame Frame7 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   225
            Left            =   5490
            TabIndex        =   13
            Top             =   550
            Width           =   4425
         End
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   260
         Left            =   90
         TabIndex        =   38
         Top             =   2940
         Width           =   5760
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00CDDADA&
      Caption         =   "CONSULTAS Y LISTADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   5880
      TabIndex        =   3
      Top             =   1590
      Width           =   2930
      Begin VB.CommandButton Command16 
         Height          =   600
         Left            =   210
         Picture         =   "SIMUFAB07frm.frx":11CA
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Emite Listado del Programa de Fabricación Activo"
         Top             =   375
         Width           =   600
      End
      Begin VB.CommandButton Command12 
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
         Left            =   2100
         Picture         =   "SIMUFAB07frm.frx":14D4
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Análisis de Restricciones"
         Top             =   375
         Width           =   600
      End
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
         Left            =   1470
         Picture         =   "SIMUFAB07frm.frx":17DE
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Necesidad de Componentes"
         Top             =   375
         Width           =   600
      End
      Begin VB.CommandButton Command11 
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
         Left            =   840
         Picture         =   "SIMUFAB07frm.frx":1AE8
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Ordenes de Fabricación a Emitir"
         Top             =   375
         Width           =   600
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0E0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   330
         Left            =   735
         TabIndex        =   19
         Top             =   1155
         Width           =   3465
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "SELECCION DE PROGRAMA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1600
      Left            =   105
      TabIndex        =   1
      Top             =   1100
      Width           =   5650
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
         Height          =   570
         Left            =   4620
         Picture         =   "SIMUFAB07frm.frx":1DF2
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Baja Lógica Individual"
         Top             =   915
         Width           =   870
      End
      Begin VB.CommandButton Command21 
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
         Left            =   4620
         Picture         =   "SIMUFAB07frm.frx":20FC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Nuevo Programa de Fabricacion"
         Top             =   315
         Width           =   435
      End
      Begin VB.CommandButton Command22 
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
         Left            =   5040
         Picture         =   "SIMUFAB07frm.frx":2246
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Nuevo Programa por Copia"
         Top             =   315
         Width           =   450
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   105
         TabIndex        =   7
         Top             =   315
         Width           =   4425
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "PROGRAMA DE FABRICACION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3140
      Left            =   105
      TabIndex        =   0
      Top             =   2760
      Width           =   5650
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2370
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   2
         Top             =   600
         Width           =   5400
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   915
         Left            =   120
         ScaleHeight     =   885
         ScaleWidth      =   5370
         TabIndex        =   4
         Top             =   300
         Width           =   5400
         Begin VB.CommandButton Command15 
            Caption         =   "Codigo               "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   0
            TabIndex        =   22
            Top             =   0
            Width           =   1620
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Descripcion de Producto   "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   1575
            TabIndex        =   21
            Top             =   0
            Width           =   2640
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Cantidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   4200
            TabIndex        =   20
            Top             =   0
            Width           =   945
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   225
            Left            =   5490
            TabIndex        =   6
            Top             =   550
            Width           =   4425
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo           Descripción              Cant."
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
            TabIndex        =   5
            Top             =   360
            Width           =   11460
         End
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Simulación de restricciones de fabricación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   700
      TabIndex        =   15
      Top             =   120
      Width           =   4395
   End
End
Attribute VB_Name = "SIMUFAB07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IPROGRA&

Private Sub Check1_Click()
   If Check1.Value = 1 Then
     Check2.Value = 0
   End If
End Sub

Private Sub Check2_Click()
   If Check2.Value = 1 Then
     Check1.Value = 0
   End If

End Sub


Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Call FINAL("")
   '
End Sub

Private Sub Command10_Click()
   ' '\\\OT-07-0386-FL-16/10/07///
   Screen.MousePointer = 11
   '
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
   If List4.TopIndex > 0 Then
      List4.TopIndex = 0
   End If
   '
10 For KIL& = 1 To List4.ListCount - 1
     AAA$ = Mid$(List4.List(KIL& - 1), 34, 9)
     BBB$ = Mid$(List4.List(KIL&), 34, 9)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
   Screen.MousePointer = 1
   ' '\\\OT-07-0386-FIN///
End Sub

Private Sub Command11_Click()
  Command11.Enabled = False
  Call CIERRARCH("*.*")
  Call ORFEMIT
  Command11.Enabled = True
End Sub

Private Sub Command12_Click()
  Command12.Enabled = False
  Call CIERRARCH("*.*")
  Call ANARESTRI
  Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   '
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
10 For KIL& = 1 To List1.ListCount - 1
     AAA$ = Mid$(List1.List(KIL& - 1), 39, 10)
     BBB$ = Mid$(List1.List(KIL&), 39, 10)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List1.List(KIL& - 1)
          List1.RemoveItem (KIL& - 1)
          List1.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List1.List(KIL& - 1)
          List1.RemoveItem (KIL& - 1)
          List1.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
   Call List4_Click
   '
End Sub

Private Sub Command14_Click()
   '
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
10 For KIL& = 1 To List1.ListCount - 1
     AAA$ = Mid$(List1.List(KIL& - 1), 17, 22)
     BBB$ = Mid$(List1.List(KIL&), 17, 22)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List1.List(KIL& - 1)
          List1.RemoveItem (KIL& - 1)
          List1.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List1.List(KIL& - 1)
          List1.RemoveItem (KIL& - 1)
          List1.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
   Call List4_Click
   '
End Sub

Private Sub COMMAND15_CLICK()
   '
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
10 For KIL& = 1 To List1.ListCount - 1
     AAA$ = Mid$(List1.List(KIL& - 1), 1, 15)
     BBB$ = Mid$(List1.List(KIL&), 1, 15)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List1.List(KIL& - 1)
          List1.RemoveItem (KIL& - 1)
          List1.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List1.List(KIL& - 1)
          List1.RemoveItem (KIL& - 1)
          List1.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
   Call List4_Click
   '
End Sub

Private Sub Command16_Click()
   Command6.Enabled = False
   '
   If List1.ListCount > 0 Then
     Screen.MousePointer = 11
     Call COPYSTRU("DEPROFAB", "LPROFAB")
     RFF$ = RECFILT$("DEPROFAB", 1, MKS$(IPROGRA&))
     JJ& = 0
     For U& = 1 To Len(RFF$) Step 4
       REFI& = CVS(Mid$(RFF$, U&, 4))
       XX$ = REGLEIDO$("DEPROFAB", REFI&)
       JJ& = JJ& + 1
       Call GRAREG("LPROFAB", XX$, JJ&)
     Next U&
     Call SETULTREG("LPROFAB", JJ&)
     Call CIERRARCH("LPROFAB")
     '
     Call HEADERS("LPROFAB", "")
     Call HEADERS("LPROFAB", "Programa: " + List3.TEXT)
     Call FINAL("*LPROFAB")
     Screen.MousePointer = 1
   End If
   '
   Command6.Enabled = True
End Sub

Private Sub Command17_Click()
   Command17.Enabled = False
   Call COMUNI("No Disponible")
   Command17.Enabled = True
End Sub


Private Sub Command19_Click()
   Command19.Enabled = False
   Call CONECRUN("WINTABU", "Consultas y Listados")
   Command19.Enabled = True
End Sub

Private Sub Command2_Click()
  Command2.Enabled = False
  Call CIERRARCH("*.*")
  Call NECCOMP
  Command2.Enabled = True
End Sub

Private Sub CALCAPPED()
'
' \\\OT-5-537-FG-22/07/05///
Dim PEDIPEN(32766) As Single
Dim CLICONPEDID%(32767)
Dim ACTICOD%(32767)
'
' INICIALIZA EL VECTOR
For i% = 1 To 32766
  PEDIPEN(i%) = 0
  CLICONPEDID%(JJ1%) = 0
Next i%
For CIXI% = 1 To NUMAS%: ACTICOD%(CIXI%) = 0: Next CIXI%
'
' SELECCIONA EL RANGO DE FECHAS A CONSIDERAR
Call DESHASFECHA(DES%, HAS%, PERIO$)
If PERIO$ = "" Then Exit Sub
DESDEL = FETEXCOR1$(DES%)
HASTAL = FETEXCOR1$(HAS%)
'

     Call ABREPEDCLI
     SQLX$ = " SELECT * FROM PEDDETA WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE (Status >= 0 AND Status < 8) "   ' para que no este anulada ni cumplida ni pendiente de aprobacion
     SQLX$ = SQLX$ + " AND CantEnt < (CanPed - 0.005) "
     SQLX$ = SQLX$ + " AND VenRep = 'V' "
     SQLX$ = SQLX$ + " AND FESOLENT >= '" & DESDEL & "' "
     SQLX$ = SQLX$ + " AND FESOLENT <= '" & HASTAL & "' "
     SQLX$ = SQLX$ + " ORDER BY FeSolEnt DESC "
     '
     Set PEDTEMP = READSET(SQLX$)
     '
     ' CARGA UN VECTOR CON LOS CLIENTES CON PEDIDOS PENDIENTES
     FILPEDIPEN% = 1
     With PEDTEMP
         Do While Not .EOF
          NCLIE1 = XVALO(!NROCLIE)
          If NCLIE1 > 0 Then
            If NCLIE1 <= 32767 Then
              CLICONPEDID%(NCLIE1) = 1 'MARCO LOS CLIENTES CON PEDIDOS
            End If
          End If
          .MoveNext
         Loop
     End With
     PEDTEMP.Close
     Set PEDTEMP = Nothing
     '
     For JL& = 1 To 32767
      If CLICONPEDID%(JL&) = 1 Then
        NCLIE2$ = FORMATNUM$(JL&, "F6.0")
        RASOCLIE2$ = rasocli$(JL&)
        Z$ = Space(128)
        Call REPLA(Z$, NCLIE2$, 1, 6)
        Call REPLA(Z$, RASOCLIE2$, 15, 64)
        With CLICONPE07
          .LICOPEN.AddItem Z$
        End With
      End If
     Next JL&
     Screen.MousePointer = 1
     If CLICONPE07.LICOPEN.ListCount > 0 Then
       Descrip$ = "Pedidos Pendientes del " & FECHATEX$(DES%) & " al " & FECHATEX$(HAS%)
       CLICONPE07.Text1 = Descrip$
25     CLICONPE07.Show 1
       Call FRESHALL
       If CLICONPE07.Label2.Caption <> "OK" Then
         Call COMUNI("Cancelado !!!")
         GoTo 9999
       End If
       '
       ' VERIFICA REPETICIÓN DEL NOMBRE DEL PROGRAMA
       Descrip$ = TRIM$(CLICONPE07.Text1)
       If Descrip$ = "" Then
           Call MENSERR(24, "Falta Nombre del Programa")
           GoTo 25
       End If
       For U& = 1 To ULTREG&("PROFABRI")
          Y$ = REGLEIDO$("PROFABRI", U&)
          If TRIM$(UCase$(Mid$(Y$, 5, 48))) = UCase$(Descrip$) Then
             Call MENSERR(24, "Nombre Repetido - Elija Otra Denominación")
             GoTo 25
          End If
       Next U&
       '
       ' ACA VUELVE A PONER EN '0' EL VECTOR DE CLIENTES
       For JL& = 1 To 32767
         CLICONPEDID%(JL&) = 0
       Next JL&
       '
       For JL& = 1 To CLICONPE07.LICOPEN.ListCount
         If CLICONPE07.LICOPEN.Selected(JL& - 1) = True Then
           ICLIE1& = Val(Left$(CLICONPE07.LICOPEN.List(JL& - 1), 6))
           If ICLIE1& > 0 Then
             If ICLIE1& <= 32767 Then
               CLICONPEDID%(ICLIE1&) = 1
               If CLISELE$ <> "" Then
                  CLISELE$ = CLISELE$ + " + "
               End If
               NCXX% = ICLIE1&
               CLISELE$ = CLISELE$ + TRIM$(ABRECLI$(NCXX%)) + "(" + TRIM$(Str$(NCXX%)) + ")"
             End If
           End If
         End If
       Next JL&
     End If
     '
     
50   Screen.MousePointer = 11
     '\\\OT-05-0631-WAR-FIN///
     Set PEDTEMP = READSET(SQLX$)
     With PEDTEMP
         Do While Not .EOF
           '
           ' AQUI DESCARTA LOS PEDIDOS DE CLIENTES NO MARCADOS
           NCLIE1 = XVALO(!NROCLIE)
           If NCLIE1 > 0 Then
             If NCLIE1 <= 32767 Then
               If CLICONPEDID%(NCLIE1) <> 1 Then
                 GoTo 103
               End If
             End If
           End If
           '
           CIXI% = XVALO(!codiint)
           CAPEN = XVALO(!CanPed) - XVALO(!CANTENT)
           If CAPEN > 0 Then
              PEDIPEN(CIXI%) = PEDIPEN(CIXI%) + CAPEN
              If MAXCOD% < CIXI% Then MAXCOD% = CIXI%
           End If
           '
103      .MoveNext
         Loop
    End With
'
' GENERA EL PROGRAMA DE FABRICACION
If MAXCOD% > 0 Then
  ' GRABA EN EL INDICE DE PROGRAMAS
  FIN1& = ULTREG&("PROFABRI")
  XX$ = REGLEIDO$("PROFABRI", FIN1&)
  NUMPROG& = CVS(Mid$(XX$, 1, 4)) + 1
  '
  Y$ = REGBLAN$("PROFABRI")
  Call REPLA(Y$, MKS$(NUMPROG&), 1, 4)
  Call REPLA(Y$, Descrip$, 5, 48)
  Call REPLA(Y$, "1", 63, 1)
  Call REPLA(Y$, "1", 64, 1)
    FIN1& = FIN1& + 1
  Call GRAREG("PROFABRI", Y$, FIN1&)
  Call SETULTREG("PROFABRI", FIN1&)
  '
  ' GRABA EL DETALLE DEL PROGRAMA
  JJ& = ULTREG&("DEPROFAB")
  For i% = 1 To MAXCOD%
    If PEDIPEN(i%) > 0.005 Then
      Y$ = REGBLAN$("DEPROFAB")
      Call REPLA(Y$, MKS$(NUMPROG&), 1, 4)
        CI% = i%
      Call REPLA(Y$, MKI$(CI%), 5, 2)
        CAPEN = PEDIPEN(CI%)
      Call REPLA(Y$, MKS$(CAPEN), 7, 4)
      JJ& = JJ& + 1
      Call GRAREG("DEPROFAB", Y$, JJ&)
    End If
  Next i%
  Call SETULTREG("DEPROFAB", JJ&)
  '
  ' ACTUALIZA LA LISTA DE PROGRAMAS
  List3.Clear
  For U& = 1 To ULTREG&("PROFABRI")
    XX$ = REGLEIDO$("PROFABRI", U&)
    DENOPRO$ = TRIM$(Mid$(XX$, 5, 48))
    MODOCA$ = Mid$(XX$, 63, 2)
    If DENOPRO$ <> "" Then
      TTXX$ = Space$(130)
      Call REPLA(TTXX$, DENOPRO$, 1, 48)
      Call REPLA(TTXX$, TRIM$(Str$(U&)), 121, 8)
      Call REPLA(TTXX$, MODOCA$, 129, 2)
      List3.AddItem TTXX$
    End If
  Next U&
  If List3.ListCount > 0 Then
    List3.ListIndex = List3.ListCount - 1
    Call LISPROFAB
    DoEvents
  End If
  '
Else
  Call COMUNI("No Existen Pedidos Pendientes dentro del Rango Seleccionado")
End If
9999 Screen.MousePointer = 1
' \\\OT-5-537-FG-FIN///
'
End Sub

Private Sub CALCAPOF()
'
' \\\OT-5-537-FG-22/07/05///
Dim OFCERRA(32766) As Single
'
' INICIALIZA EL VECTOR
For i% = 1 To 32766
  OFCERRA(i%) = 0
Next i%
'
' SELECCIONA EL RANGO DE FECHAS A CONSIDERAR
Call DESHASFECHA(DES%, HAS%, PERIO$)
'
FIN& = ULTREG&("MOVISTO")
MAXCOD% = 0
'
' BUSCA LAS O/F CERRADAS
Screen.MousePointer = 11
For U& = 1 To FIN&
  Call TRACE(20, U&, FIN&)
  X$ = REGLEIDO$("MOVISTO", U&)
    FE% = CVI(Mid$(X$, 1, 2)) ' fecha del movimiento
    If FE% >= DES% Then
      If FE% <= HAS% Then
        CODMOV$ = Mid$(X$, 21, 2)
        If CODMOV$ = "OF" Then
          CANENT = CVD(Mid$(X$, 88, 8))
          If CANENT > 0.005 Then
            CI% = CVI(Mid$(X$, 3, 2))
            If CI% > 0 And CI% <= NUMAS% Then
              OFCERRA(CI%) = OFCERRA(CI%) + CANENT
              If MAXCOD% < CI% Then MAXCOD% = CI%
            End If
          End If
        End If
      End If
    End If
Next U&
'
' GENERA EL PROGRAMA DE FABRICACION
If MAXCOD% > 0 Then
  ' GRABA EN EL INDICE DE PROGRAMAS
  FIN1& = ULTREG&("PROFABRI")
  XX$ = REGLEIDO$("PROFABRI", FIN1&)
  NUMPROG& = CVS(Mid$(XX$, 1, 4)) + 1
  '
  Y$ = REGBLAN$("PROFABRI")
  Call REPLA(Y$, MKS$(NUMPROG&), 1, 4)
    Descrip$ = "Por O/F's Cerradas del " & FECHATEX$(DES%) & " al " & FECHATEX$(HAS%)
  Call REPLA(Y$, Descrip$, 5, 48)
  Call REPLA(Y$, "2", 63, 1)
  Call REPLA(Y$, "1", 64, 1)
    FIN1& = FIN1& + 1
  Call GRAREG("PROFABRI", Y$, FIN1&)
  Call SETULTREG("PROFABRI", FIN1&)
  '
  ' GRABA EL DETALLE DEL PROGRAMA
  JJ& = ULTREG&("DEPROFAB")
  For i% = 1 To MAXCOD%
    If OFCERRA(i%) > 0.005 Then
      Y$ = REGBLAN$("DEPROFAB")
      Call REPLA(Y$, MKS$(NUMPROG&), 1, 4)
        CI% = i%
      Call REPLA(Y$, MKI$(CI%), 5, 2)
        CANENT = OFCERRA(CI%)
      Call REPLA(Y$, MKS$(CANENT), 7, 4)
      JJ& = JJ& + 1
      Call GRAREG("DEPROFAB", Y$, JJ&)
    End If
  Next i%
  Call SETULTREG("DEPROFAB", JJ&)
  '
  ' ACTUALIZA LA LISTA DE PROGRAMAS
  List3.Clear
  For U& = 1 To ULTREG&("PROFABRI")
    XX$ = REGLEIDO$("PROFABRI", U&)
    DENOPRO$ = TRIM$(Mid$(XX$, 5, 48))
    MODOCA$ = Mid$(XX$, 63, 2)
    If DENOPRO$ <> "" Then
      TTXX$ = Space$(130)
      Call REPLA(TTXX$, DENOPRO$, 1, 48)
      Call REPLA(TTXX$, TRIM$(Str$(U&)), 121, 8)
      Call REPLA(TTXX$, MODOCA$, 129, 2)
      List3.AddItem TTXX$
    End If
  Next U&
  If List3.ListCount > 0 Then
    List3.ListIndex = List3.ListCount - 1
    Call LISPROFAB
    DoEvents
  End If
  '
Else
  Call COMUNI("No Existen Entradas por Orden de Fabricación dentro del Rango Seleccionado")
End If
Screen.MousePointer = 1
' \\\OT-5-537-FG-FIN///
'
End Sub

Private Sub Command21_CLICK()
   Command21.Enabled = False
   '
   X% = COMALTER%("Opciones de Generación de Programa\\Manual\Por Pedidos Pendientes")
   '
   If X% = 1 Then
     URE& = ULTREG&("PROFABRI")
     VDEF$ = MKS$(URE& + 1) + Space$(48)
25   OBX$ = OBJPLA$("NUEPROFAB", VDEF$)
     If OBX$ <> "" Then
        VDEF$ = OBX$
        DENOPRO$ = TRIM$(Mid$(OBX$, 5, 48))
        If DENOPRO$ <> "" Then
           '
           ' VERIFICA NOMBRE NO REPETIDO
           For U& = 1 To ULTREG&("PROFABRI")
              Y$ = REGLEIDO$("PROFABRI", U&)
              If TRIM$(UCase$(Mid$(Y$, 5, 48))) = UCase$(DENOPRO$) Then
                 Call MENSERR(24, "Nombre Repetido - Elija Otra Denominación")
                 GoTo 25
              End If
           Next U&
           '
           Call BLOQARCH("PROFABRI")
           REBLA$ = REGBLAN$("PROFABRI")
           URE& = ULTREG&("PROFABRI")
           Call REPLA(REBLA$, MKS$(URE& + 1), 1, 4)
           Call REPLA(REBLA$, DENOPRO$, 5, 48)
           Call REPLA(Y$, "  ", 63, 2)
           Call REGAPP("PROFABRI", REBLA$)
           Call CIERRARCH("PROFABRI")
           '
           ILISTA& = List3.ListCount
           TTXX$ = Space$(130)
           Call REPLA(TTXX$, DENOPRO$, 1, 48)
           Call REPLA(TTXX$, TRIM$(Str$(URE& + 1)), 121, 8)
           Call REPLA(TTXX$, MODOCA$, 129, 2)
           List3.AddItem TTXX$
           List3.ListIndex = ILISTA&
        End If
     End If
   ElseIf X% = 2 Then
     Call CALCAPPED
   End If
   Command21.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   If List3.ListIndex >= 0 Then
     REGA& = Val(Mid$(List3.TEXT, 121, 8))
     DENOA$ = TRIM$(Left$(List3.TEXT, 48))
     URE& = ULTREG&("PROFABRI")
10   VDEF$ = MKS$(URE& + 1) + "Copia de " + DENOA$
     OBX$ = OBJPLA$("NUEPROFAB", VDEF$)
     If OBX$ <> "" Then
       DENOPRO$ = TRIM$(Mid$(OBX$, 5, 48))
       If DENOPRO$ <> "" Then
         For KU& = 1 To List3.ListCount
           If UCase$(DENOPRO$) = UCase$(TRIM$(Left$(List3.List(KU& - 1), 48))) Then
             Call MENSERR(24, "Denominacion ya Utilizada.\Debe Cambiar Denominacion del Plan")
             GoTo 10
           End If
         Next KU&
         '
         SCMA = Screen.MousePointer
         Screen.MousePointer = 11
         Call BLOQARCH("PROFABRI")
         Call BLOQARCH("DEPROFAB")
         REBLA$ = REGBLAN$("PROFABRI")
         URE& = ULTREG&("PROFABRI")
         Call REPLA(REBLA$, MKS$(URE& + 1), 1, 4)
         Call REPLA(REBLA$, DENOPRO$, 5, 48)
         Call REPLA(REBLA$, "  ", 63, 2)
         Call REGAPP("PROFABRI", REBLA$)
         Call CIERRARCH("PROFABRI")
         '
         ILISTA& = List3.ListCount
         TTXX$ = Space$(128)
         Call REPLA(TTXX$, DENOPRO$, 1, 48)
         Call REPLA(TTXX$, TRIM$(Str$(URE& + 1)), 121, 8)
         MODOCA$ = "  "
         Call REPLA(TTXX$, MODOCA$, 129, 2)
         List3.AddItem TTXX$
         List3.ListIndex = ILISTA&
         '
         RFF$ = RECFILT$("DEPROFAB", 1, MKS$(REGA&))
         List1.Clear
         For U& = 1 To Len(RFF$) Step 4
           REFI& = CVS(Mid$(RFF$, U&, 4))
           XX$ = REGLEIDO$("DEPROFAB", REFI&)
           Call REPLA(XX$, MKS$(URE& + 1), 1, 4)
           Call REGAPP("DEPROFAB", XX$)
           '
           CIXI% = CVI(Mid$(XX$, 5, 2))
           CAXX = CVS(Mid$(XX$, 7, 4))
           TTXX$ = Space$(64)
           Call REPLA(TTXX$, CODEXT$(CIXI%), 1, 16)
           Call REPLA(TTXX$, DESCRIT0$(CIXI%), 17, 22)
           Call REPLA(TTXX$, FORMATNUM$(CAXX, "F10.1"), 39, 10)
           List1.AddItem TTXX$
         Next U&
         Call CIERRARCH("DEPROFAB")
         Screen.MousePointer = SCMA
       End If
     End If
   End If
   Command22.Enabled = True
End Sub

Private Sub Command23_Click()
   Command23.Enabled = False
   Call CIERRARCH("*.*")
   Call PLACOMP
   Command23.Enabled = True
End Sub

Private Sub Command24_Click()
   Command24.Enabled = False
   If List3.ListIndex >= 0 Then
      If COMALTER%("Está Seguro de Anular el Plan Activo ?\\No, Conservarlo\Si, Anularlo") = 2 Then
        REGA& = Val(Mid$(List3.TEXT, 121, 8))
        XX$ = REGLEIDO$("PROFABRI", REGA&)
        Call REPLA(XX$, Space$(48), 5, 48)
        Call GRAREG("PROFABRI", XX$, REGA&)
        '
        '\\\OT-5-0537-FG-12/08/05///
        ' ANULACION DEL DETALLE DEL PROGRAMA
        NUMPROG& = CVS(Mid$(XX$, 1, 4))
        FIN& = ULTREG&("DEPROFAB")
        JJ& = 0
        For U& = 1 To FIN&
          YY$ = REGLEIDO$("DEPROFAB", U&)
          NPRO& = CVS(Mid$(YY$, 1, 4))
          If NPRO& <> NUMPROG& Then
            JJ& = JJ& + 1
            Call GRAREG("DEPROFAB", YY$, JJ&)
          End If
        Next U&
        '
        Call SETULTREG("DEPROFAB", JJ&)
        '\\\OT-5-0537-FG-FIN///
        '
        List3.Clear
        Screen.MousePointer = 11
        For U& = 1 To ULTREG&("PROFABRI")
          XX$ = REGLEIDO$("PROFABRI", U&)
          DENOPRO$ = TRIM$(Mid$(XX$, 5, 48))
          MODOCA$ = Mid$(XX$, 63, 2)
          If DENOPRO$ <> "" Then
            TTXX$ = Space$(128)
            Call REPLA(TTXX$, DENOPRO$, 1, 48)
            Call REPLA(TTXX$, TRIM$(Str$(U&)), 121, 8)
            Call REPLA(TTXX$, MODOCA$, 129, 2)
            List3.AddItem TTXX$
          End If
        Next U&
        If List3.ListCount > 0 Then
          List3.ListIndex = List3.ListCount - 1
        End If
        Screen.MousePointer = 1
      End If
   End If
   Command24.Enabled = True
End Sub

Private Sub COMMAND3_Click()
   Command3.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG07/AMASTO"
   Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
   Command3.Enabled = True
End Sub


Private Sub Command4_Click()
   ' '\\\OT-07-0386-FL-16/10/07///
   Screen.MousePointer = 11
   '
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
   If List4.TopIndex > 0 Then
      List4.TopIndex = 0
   End If
   '
10 For KIL& = 1 To List4.ListCount - 1
     AAA$ = Mid$(List4.List(KIL& - 1), 1, 15)
     BBB$ = Mid$(List4.List(KIL&), 1, 15)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
   Screen.MousePointer = 1
    '\\\OT-07-0386-FL-FIN///
End Sub

Private Sub Command5_Click()
' '\\\OT-07-0386-FL-16/10/07///
Screen.MousePointer = 11
'
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
   If List4.TopIndex > 0 Then
      List4.TopIndex = 0
   End If
   '
10 For KIL& = 1 To List4.ListCount - 1
     AAA$ = Mid$(List4.List(KIL& - 1), 16, 9)
     BBB$ = Mid$(List4.List(KIL&), 16, 9)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
   Screen.MousePointer = 1
   ' '\\\OT-07-0386-FIN///
End Sub

Private Sub Command6_Click()
' '\\\OT-07-0386-FL-16/10/07///
   Screen.MousePointer = 11
   '
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
   If List4.TopIndex > 0 Then
      List4.TopIndex = 0
   End If
   
10 For KIL& = 1 To List4.ListCount - 1
     AAA$ = Mid$(List4.List(KIL& - 1), 25, 9)
     BBB$ = Mid$(List4.List(KIL&), 25, 9)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
   Screen.MousePointer = 1
   ' '\\\OT-07-0386-///
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   Call CALCUMAT
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   ' '\\\OT-07-0386-FL-16/10/07///
   Screen.MousePointer = 11
   '
   Static SENTI%
   If Abs(SENTI%) < 1 Then SENTI% = -1
   SENTI% = (-1) * SENTI%
   '
    If List4.TopIndex > 0 Then
      List4.TopIndex = 0
    End If
   
10 For KIL& = 1 To List4.ListCount - 1
     AAA$ = Mid$(List4.List(KIL& - 1), 43, 9)
     BBB$ = Mid$(List4.List(KIL&), 43, 9)
     If SENTI% > 0 Then
        If AAA$ > BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
      Else
        If AAA$ < BBB$ Then
          CCC$ = List4.List(KIL& - 1)
          List4.RemoveItem (KIL& - 1)
          List4.AddItem CCC$
          GoTo 10
        End If
     End If
   Next KIL&
   '
   Screen.MousePointer = 1
   ' '\\\OT-07-0386-FIN///
End Sub



Private Sub Command9_Click()
    Command9.Enabled = False
    Call GENORFEM
    Command9.Enabled = True
End Sub

Private Sub Form_Load()
   '
   'Call CENTRAFORM(Me)
   'Show
   'Refresh
   ' VERANTER$ = "PROMRP01"
   Call VERJOBID(OKEY%)

   If OKEY% < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
10   Caption = Caption + "  -  " + EPAC$
     Call GRATITFOR(Caption)
   End If
   Openforms = DoEvents
   '
   Screen.MousePointer = 11
   For U& = 1 To ULTREG&("PROFABRI")
     XX$ = REGLEIDO$("PROFABRI", U&)
     DENOPRO$ = TRIM$(Mid$(XX$, 5, 48))
     MODOCA$ = Mid$(XX$, 63, 2)
     FECA% = CVI(Mid$(XX$, 65, 2))
     HOCA% = CVI(Mid$(XX$, 67, 2))
     If DENOPRO$ <> "" Then
       TTXX$ = Space$(130)
       Call REPLA(TTXX$, DENOPRO$, 1, 48)
       Call REPLA(TTXX$, TRIM$(Str$(U&)), 121, 8)
       Call REPLA(TTXX$, MODOCA$, 129, 2)
       Call REPLA(TTXX$, FECHATEX$(FECA%), 131, 8)
       Call REPLA(TTXX$, HORATEX$(HOCA%), 139, 8)
       List3.AddItem TTXX$
     End If
   Next U&
   If List3.ListCount > 0 Then
     List3.ListIndex = List3.ListCount - 1
   End If
   
   If InStr(1, EPAC$, "ASEMBLI") > 0 Then
    Frame15.Visible = True
   End If
   
   Screen.MousePointer = 1
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

Sub CALCUMAT_MAL()
   '
   ' calculo de capacidad
   '
   Screen.MousePointer = 11
   Call RECUSALSTOK
   ILIS3& = List3.ListIndex
   '
   Command6.Caption = "    Stock"
   If Check1.Value = 1 Then Command6.Caption = "Stk. Neto"
   If Check2.Value = 1 Then Command6.Caption = "Stk.0"
   '
   Dim COEQUI$(1024), TIEQUI(1024)
   CANEQ% = 0
   Dim CANECE(32767), TICATO(1024)
   CANCA% = 0
   '
   'REBLA$ = REGBLAN$("CAPANECE")
   JJ& = 0
   FIP& = List1.ListCount
   '
   For i& = 1 To FIP&
     Screen.MousePointer = 11
     Call TRACE(20, i&, FIP&)
     VECOMPO$ = ""
     Label13 = TRIM$(Str$(i&))
     Label13.Refresh
     List1.ListIndex = i& - 1
     P00$ = Left$(List1.List(i& - 1), 64)
     P1$ = TRIM$(Left$(P00$, 16))
     CIKI% = CODINT(P1$)
     '
     If CIKI% > 0 Then
       If CIKI% <= NUMAS% Then
         CANPROG = Val(Mid$(P00$, 39, 10))
         Call EXPLOCOM(CIKI%, 1)
         If CONTAMUL% > 0 Then
           
           ACTIX% = 0
           For K% = 1 To CONTAMUL%
             CIX1% = CIK%(K%)
             CAN = CANT(K%)
             PD% = 0
             If CIX1% > 0 Then
               If CIX1% <= NUMAS% Then
                  If Abs(CAN) > 0.0005 Then
                     CANECE(CIX1%) = CANECE(CIX1%) + CAN * CANPROG
                     TBUS$ = TRIM$(Str$(CIX1%)) + ","
                     If InStr(VECOMPO$, TBUS$) < 1 Then
                        VECOMPO$ = VECOMPO$ + TBUS$
                     End If
                  End If
               End If
             End If
           Next K%
         End If
       End If
     End If
     TTXY$ = AJUSTI$(List1.List(i& - 1), 64) + VECOMPO$
     List1.RemoveItem i& - 1
     List1.AddItem TTXY$, i& - 1
   Next i&
   '
   
   
   List4.Clear: SUTIE = 0: J4& = 0
    Call LEEEXPLO(CODEXT$(CIKI%), 1, 1)
    For II% = 1 To CAIT%
        CIX1% = CIJ%(CAIT%)
        CAN = USOI(CAIT%)
        If Abs(CAN) > 0.0005 Then
                CANECE(CIX1%) = CANECE(CIX1%) + CAN * CANPROG
                TBUS$ = TRIM$(Str$(CIX1%)) + ","
                If InStr(VECOMPO$, TBUS$) < 1 Then
                   VECOMPO$ = VECOMPO$ + TBUS$
                End If
        End If
        STOACT = 0
        TUI% = CIX1%

        If Check1.Value = 1 Then
          STOACT = STOCKACT(TUI%) - STOCKRES(TUI%)
        End If
        If Check2.Value = 0 And Check1.Value = 0 Then  'CONSIDERA EL STOCKS EN EL CASO QUE NO ESTE NINGUNO MARCADO
           STOACT = STOCKACT(TUI%)
       End If
       '
       CANTNEC = CANECE(TUI%)
       If Abs(CANTNEC) >= 0.05 Then
           TTXX$ = Space$(128)
           Call REPLA(TTXX$, CODEXT$(TUI%), 1, 15)
           Call REPLA(TTXX$, FORMATNUM$(CANTNEC, "F9.1"), 16, 9)
           Call REPLA(TTXX$, FORMATNUM$(STOACT, "F9.1"), 25, 9)
           If STOACT < CANTNEC Then
              Call REPLA(TTXX$, FORMATNUM$(CANTNEC - STOACT, "F9.1"), 34, 9)
           End If
           Call REPLA(TTXX$, FORMATNUM$(CANOCPEN(TUI%) + CANRECEP(TUI%), "F9.1"), 43, 9)
           List4.AddItem (TTXX$)
         End If
    Next II%

'   For TUI% = 1 To NUMAS%
'       STOACT = 0
'       If Check1.Value = 1 Then
'          STOACT = STOCKACT(TUI%) - STOCKRES(TUI%)
'       End If
'       If Check2.Value = 0 And Check1.Value = 0 Then  'CONSIDERA EL STOCKS EN EL CASO QUE NO ESTE NINGUNO MARCADO
'           STOACT = STOCKACT(TUI%)
'           CODIGOEX$ = CODEXT$(TUI%)
'       End If
'       '
'       CANTNEC = CANECE(TUI%)
'         If Abs(CANTNEC) >= 0.05 Then
'           TTXX$ = Space$(128)
'           Call REPLA(TTXX$, CODEXT$(TUI%), 1, 15)
'           Call REPLA(TTXX$, FORMATNUM$(CANTNEC, "F9.1"), 16, 9)
'           Call REPLA(TTXX$, FORMATNUM$(STOACT, "F9.1"), 25, 9)
'           If STOACT < CANTNEC Then
'              Call REPLA(TTXX$, FORMATNUM$(CANTNEC - STOACT, "F9.1"), 34, 9)
'           End If
'           Call REPLA(TTXX$, FORMATNUM$(CANOCPEN(TUI%) + CANRECEP(TUI%), "F9.1"), 43, 9)
'           List4.AddItem (TTXX$)
'         End If
'   Next TUI%
   List4.Refresh
   '
   Label15 = FECHATEX$(HOY(HOS$))
   Label17 = Format$(Time, "HH:MM:SS")
   Screen.MousePointer = 1
   '
End Sub

Sub CALCUMAT()
   '
   ' calculo de capacidad
   '
   Screen.MousePointer = 11
   'Call RECUSALSTOK
   ILIS3& = List3.ListIndex
   '
   Command6.Caption = "    Stock"
   If Check1.Value = 1 Then Command6.Caption = "Stk. Neto"
   If Check2.Value = 1 Then Command6.Caption = "Stk.0"
   '
   Dim COEQUI$(1024), TIEQUI(1024)
   CANEQ% = 0
   Dim CANECE(32767), TICATO(1024)
   CANCA% = 0
   '
   'REBLA$ = REGBLAN$("CAPANECE")
   JJ& = 0
   FIP& = List1.ListCount
   '
   For i& = 1 To FIP&
     Screen.MousePointer = 11
     Call TRACE(20, i&, FIP&)
     VECOMPO$ = ""
     Label13 = TRIM$(Str$(i&))
     Label13.Refresh
     List1.ListIndex = i& - 1
     P00$ = Left$(List1.List(i& - 1), 64)
     P1$ = TRIM$(Left$(P00$, 16))
     CIKI% = CODINT(P1$)
     '
     If CIKI% > 0 Then
       If CIKI% <= NUMAS% Then
         CANPROG = Val(Mid$(P00$, 39, 10))
         Call EXPLOCOM(CIKI%, 1)
         If CONTAMUL% > 0 Then
           '
           ACTIX% = 0
           For K% = 1 To CONTAMUL%
             CIX1% = CIK%(K%)
             CAN = CANT(K%)
             PD% = 0
             If CIX1% > 0 Then
               If CIX1% <= NUMAS% Then
                  If Abs(CAN) > 0.0005 Then
                     CANECE(CIX1%) = CANECE(CIX1%) + CAN * CANPROG
                     TBUS$ = TRIM$(Str$(CIX1%)) + ","
                     If InStr(VECOMPO$, TBUS$) < 1 Then
                        VECOMPO$ = VECOMPO$ + TBUS$
                     End If
                  End If
               End If
             End If
           Next K%
         End If
       End If
     End If
     TTXY$ = AJUSTI$(List1.List(i& - 1), 64) + VECOMPO$
     List1.RemoveItem i& - 1
     List1.AddItem TTXY$, i& - 1
   Next i&
   '
   List4.Clear: SUTIE = 0: J4& = 0
   For TUI% = 1 To NUMAS%
       STOACT = 0
       If Check1.Value = 1 Then
          STOACT = STOCKACT(TUI%) - STOCKRES(TUI%)
       End If
       If Check2.Value = 0 And Check1.Value = 0 Then  'CONSIDERA EL STOCKS EN EL CASO QUE NO ESTE NINGUNO MARCADO
           STOACT = STOCKACT(TUI%)
       End If
       '
       CANTNEC = CANECE(TUI%)
         If Abs(CANTNEC) >= 0.05 Then
           TTXX$ = Space$(128)
           Call REPLA(TTXX$, CODEXT$(TUI%), 1, 15)
           Call REPLA(TTXX$, FORMATNUM$(CANTNEC, "F9.1"), 16, 9)
           Call REPLA(TTXX$, FORMATNUM$(STOACT, "F9.1"), 25, 9)
           If STOACT < CANTNEC Then
              Call REPLA(TTXX$, FORMATNUM$(CANTNEC - STOACT, "F9.1"), 34, 9)
           End If
           Call REPLA(TTXX$, FORMATNUM$(CANOCPEN(TUI%) + CANRECEP(TUI%), "F9.1"), 43, 9)
           List4.AddItem (TTXX$)
         End If
   Next TUI%
   List4.Refresh
   '
   Label15 = FECHATEX$(HOY(HOS$))
   Label17 = Format$(Time, "HH:MM")
   Screen.MousePointer = 1
   
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Label1_DblClick()
   Call CARPROFAB
End Sub

Sub CARPROFAB()
   DEPROGRA$ = List3.TEXT
   IPROGRA& = Val(Mid$(DEPROGRA$, 121, 8))
   TIPROGRA% = Val(Mid$(DEPROGRA, 129, 1))
   '
   If IPROGRA& < 1 Then
        Call MENSERR(24, "Imposible Abrir Programa '0'")
        Exit Sub
      Else
        Call TRALOFIL("DEPROFAB", 1, MKS$(IPROGRA&))
        VTB% = VENTABU%("CARPROFA")
        If VTB% >= 0 Then
          Call TRACENFIL("DEPROFAB", 1, MKS$(IPROGRA&))
          Call LISPROFAB
        End If
   End If
End Sub
'
Sub LISPROFAB()
   DEPROGRA$ = List3.TEXT
   'Label15 = TRIM$(Mid$(DEPROGRA$, 131, 8))
   'Label17 = TRIM$(Mid$(DEPROGRA$, 139, 8))
   IPROGRA& = Val(Mid$(DEPROGRA$, 121, 8))
   If IPROGRA& < 1 Then
        Call MENSERR(24, "Imposible Abrir Programa '0'")
        Exit Sub
      Else
        SCMA = Screen.MousePointer
        Screen.MousePointer = 11
        RFF$ = RECFILT$("DEPROFAB", 1, MKS$(IPROGRA&))
        List1.Clear
        List2.Clear
        List6.Clear
        Label8 = ""
       For U& = 1 To Len(RFF$) Step 4
          REFI& = CVS(Mid$(RFF$, U&, 4))
          XX$ = REGLEIDO$("DEPROFAB", REFI&)
          CIXI% = CVI(Mid$(XX$, 5, 2))
          CAXX = CVS(Mid$(XX$, 7, 4))
          TTXX$ = Space$(64)
          Call REPLA(TTXX$, CODEXT$(CIXI%), 1, 16)
          Call REPLA(TTXX$, DESCRIT0$(CIXI%), 17, 22)
          Call REPLA(TTXX$, FORMATNUM$(CAXX, "F10.1"), 39, 10)
          List1.AddItem TTXX$
          '\\\OT-07-0386-FL-17/10/07///
          TTXX2$ = Space$(50)
          CANFAP = (CANFAPEN(CIXI%))
          Call REPLA(TTXX2$, CODEXT$(CIXI%), 1, 15)
          Call REPLA(TTXX2$, FORMATNUM(CAXX, "F8.1"), 15, 8)
          Call REPLA(TTXX2$, FORMATNUM(CANFAP, "F8.1"), 25, 8)
          '
          If CAXX > CANFAP Then
            Call REPLA(TTXX2$, FORMATNUM(CAXX - CANFAP, "F8.1"), 34, 8)
          End If
          '
          List6.AddItem TTXX2$
          '\\\OT-07-0382-FL-12/10/07///
        Next U&
        '
        Call TRALOFIL("RECALMAQ", 1, MKS$(IPROGRA&))
        Call TRALOFIL("RECALMOD", 1, MKS$(IPROGRA&))
        '
        List4.Clear: TOTIX = 0
        For U& = 1 To ULTREG&("!RECALMAQ")
          TTXX$ = Mid$(REGLEIDO$("!RECALMAQ", U&), 5)
          List4.AddItem TTXX$
          TOTIX = TOTIX + Val(Mid$(TTXX$, 34, 2)) + 60 * Val(Mid$(TTXX$, 31, 2)) + 3600 * Val(Mid$(TTXX$, 26, 4))
        Next U&
        List4.Refresh
        '
        Label10 = "": TOTIX = 0
        For U& = 1 To ULTREG&("!RECALMOD")
          TTXX$ = Mid$(REGLEIDO$("!RECALMOD", U&), 5)
          List5.AddItem TTXX$
          TOTIX = TOTIX + Val(Mid$(TTXX$, 34, 2)) + 60 * Val(Mid$(TTXX$, 31, 2)) + 3600 * Val(Mid$(TTXX$, 26, 4))
        Next U&
        
        If TOTIX >= 1 Then
          Label10 = HORMINSE$(TOTIX)
        End If
        '
        Screen.MousePointer = SCMA
   End If
End Sub
'

Private Sub Label10_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub List1_DblClick()
   Call CARPROFAB
End Sub

Private Sub List3_Click()
   Label15 = ""
   Label17 = ""
   Label13 = ""
   Call LISPROFAB
End Sub

Private Sub LIST3_DBLCLICK()
   Call CARPROFAB
End Sub

Private Sub List4_Click()
    Screen.MousePointer = 11
    List2.Clear
    CICOMPO% = CODINT%(TRIM$(Left$(List4.TEXT, 16)))
    ABUS$ = TRIM$(Str$(CICOMPO%)) + ","
    For U& = 1 To List1.ListCount
       List1.Selected(U& - 1) = False
       TTXY$ = Mid$(List1.List(U& - 1), 65)
       If InStr(TTXY$, ABUS$) > 0 Then
          TTXX$ = List1.List(U& - 1)
          CODICON% = CODINT%(TRIM$(Left$(TTXX$, 16)))
          USOUN = 0
          If CODICON% > 0 Then
             Call EXPLOCOM(CODICON%, 1)
             If CONTAMUL% > 0 Then
                For K% = 1 To CONTAMUL%
                   If CIK%(K%) = CICOMPO% Then
                      USOUN = USOUN + CANT(K%)
                   End If
                Next K%
             End If
          End If
          USOX$ = FORMATNUM$(USOUN, "F8.4")
          Call REPLA(TTXX$, USOX$, 32, 8)
          List2.AddItem TTXX$
          List2.Refresh
          List1.Selected(U& - 1) = True
       End If
    Next U&
    Label8 = DESCRIT0$(CICOMPO%)
    Screen.MousePointer = 1
End Sub

Sub ANARESTRI()
    If TRIM$(Label15) = "" Or TRIM$(Label17) = "" Then
       Call MENSERR(24, "Falta Re-Calcular")
       Exit Sub
    End If
    '
    If List4.ListCount < 1 Then
       Call MENSERR(24, "No Hay Necesidad de Componentes")
       Exit Sub
    End If
    '
    ILISTA& = List4.ListIndex
    JJ& = 0
    REBLA$ = REGBLAN$("ANAREST")
    For U& = 1 To List4.ListCount
       TTXX$ = List4.List(U& - 1)
       FALTA# = XVALO(Mid$(TTXX$, 34, 9))
       '
       If FALTA > 0.05 Then
          List4.ListIndex = U& - 1
          Call List4_Click
          Code$ = TRIM$(Mid$(TTXX$, 1, 15))
          CANTNEC# = XVALO(Mid$(TTXX$, 16, 9))
          STOACT# = XVALO(Mid$(TTXX$, 25, 9))
          OCPEN# = XVALO(Mid$(TTXX$, 43, 9))
          XX0$ = REBLA$
          '
          Call REPLA(XX0$, MKI$(CODINT%(Code$)), 1, 2)
          Call REPLA(XX0$, MKD$(CANTNEC#), 3, 8)
          Call REPLA(XX0$, MKD$(STOACT#), 11, 8)
          Call REPLA(XX0$, MKD$(FALTA#), 19, 8)
          Call REPLA(XX0$, MKD$(OCPEN#), 27, 8)
          '
          For U2& = 1 To List2.ListCount
             TTYY$ = List2.List(U2& - 1)
             Code$ = TRIM$(Left$(TTYY$, 16))
             USOX# = XVALO(Mid$(TTYY$, 32, 8))
             CANTU# = XVALO(Mid$(TTYY$, 40, 9))
             XX$ = XX0$
             Call REPLA(XX$, MKI$(CODINT%(Code$)), 65, 2)
             Call REPLA(XX$, MKD$(USOX#), 67, 8)
             Call REPLA(XX$, MKD$(CANTU#), 75, 8)
             Call REPLA(XX$, MKD$(CANTU# * USOX#), 83, 8)
             '
             JJ& = JJ& + 1
             Call GRAREG("ANAREST", XX$, JJ&)
          Next U2&
       End If
    Next U&
    '
    Call SETULTREG("ANAREST", JJ&)
    List4.ListIndex = ILISTA&
    If JJ& < 1 Then
       Call COMUNI("No Hay Restricciones de Stocks\para el Programa Elegido.")
       Exit Sub
    End If
    '
    Call HEADERS("ANAREST", "")
    Call HEADERS("ANAREST", "Programa: " + List3.TEXT)
    Call FINAL("*ANAREST")
    '
End Sub

Sub ORFEMIT()
    '
    ILISTA& = List6.ListIndex
    JJ& = 0
    REBLA$ = REGBLAN$("ORFEMIT")
    For U& = 1 To List6.ListCount
       TTXX$ = List6.List(U& - 1)
       Code$ = TRIM$(Mid$(TTXX$, 1, 15))
       CANTNEC# = XVALO(Mid$(TTXX$, 15, 8))
       STOACT# = XVALO(Mid$(TTXX$, 25, 8))
       FALTA# = XVALO(Mid$(TTXX$, 34, 8))
       XX$ = REBLA$
       '
       Call REPLA(XX$, MKI$(CODINT%(Code$)), 1, 2)
       Call REPLA(XX$, MKD$(CANTNEC#), 3, 8)
       Call REPLA(XX$, MKD$(STOACT#), 11, 8)
       Call REPLA(XX$, MKD$(FALTA#), 19, 8)
       '
       JJ& = JJ& + 1
       Call GRAREG("ORFEMIT", XX$, JJ&)
    Next U&
    '
    Call SETULTREG("ORFEMIT", JJ&)
    List6.ListIndex = ILISTA&
    If JJ& < 1 Then
       Call COMUNI("No Hay Ordenes de Fabricación a Emitir\para el Programa Elegido.")
       Exit Sub
    End If
    '
    Call HEADERS("ORFEMIT", "")
    Call HEADERS("ORFEMIT", "Programa: " + List3.TEXT)
    Call FINAL("*ORFEMIT")
    '
End Sub

Sub NECCOMP()
    If TRIM$(Label15) = "" Or TRIM$(Label17) = "" Then
       Call MENSERR(24, "Falta Re-Calcular")
       Exit Sub
    End If
    '
    If List4.ListCount < 1 Then
       Call MENSERR(24, "No Hay Necesidad de Componentes")
       Exit Sub
    End If
    '
    ILISTA& = List4.ListIndex
    JJ& = 0
    REBLA$ = REGBLAN$("NECCOMP")
    For U& = 1 To List4.ListCount
       TTXX$ = List4.List(U& - 1)
       FALTA# = XVALO(Mid$(TTXX$, 34, 9))
       Code$ = TRIM$(Mid$(TTXX$, 1, 15))
       CANTNEC# = XVALO(Mid$(TTXX$, 16, 9))
       STOACT# = XVALO(Mid$(TTXX$, 25, 9))
       OCPEN# = XVALO(Mid$(TTXX$, 43, 9))
       DISPON# = STOACT# + OCPEN# - CANTNEC#
       '
       XX$ = REBLA$
       '
       Call REPLA(XX$, MKI$(CODINT%(Code$)), 1, 2)
       Call REPLA(XX$, MKD$(CANTNEC#), 3, 8)
       Call REPLA(XX$, MKD$(STOACT#), 11, 8)
       Call REPLA(XX$, MKD$(FALTA#), 19, 8)
       Call REPLA(XX$, MKD$(OCPEN#), 27, 8)
       '
       If DISPON# > 0 Then
            Call REPLA(XX$, MKD$(DISPON#), 35, 8)
          Else
            Call REPLA(XX$, MKD$(Abs(DISPON#)), 43, 8)
       End If
       '
       JJ& = JJ& + 1
       Call GRAREG("NECCOMP", XX$, JJ&)
    Next U&
    '
    Call SETULTREG("NECCOMP", JJ&)
    List4.ListIndex = ILISTA&
    If JJ& < 1 Then
       Call COMUNI("No Hay Necesidad de Componentes\para el Programa Elegido.")
       Exit Sub
    End If
    '
    Call HEADERS("NECCOMP", "")
    Call HEADERS("NECCOMP", "Programa: " + List3.TEXT)
    Call FINAL("*NECCOMP")
    '
End Sub

Sub GENORFEM()
    '
    Screen.MousePointer = 11
    ILISTA& = List6.ListIndex
    JJ& = 0
    REBLA$ = REGBLAN$("ORFEMIT")
    For U& = 1 To List6.ListCount
       TTXX$ = List6.List(U& - 1)
       Code$ = TRIM$(Mid$(TTXX$, 1, 15))
         CI% = CODINT%(Code$)
       CANTX = XVALO(Mid$(TTXX$, 15, 8))
         FF% = HOY(HOS$)
       STOACT# = XVALO(Mid$(TTXX$, 25, 8))
       FALTA# = XVALO(Mid$(TTXX$, 34, 8))
       '
       If FALTA# > 0 Then
          FF1$ = REGBLAN$("ORFAEMI")
          Call REPLA(FF1$, MKI$(CI%), 1, 2)
          Call REPLA(FF1$, DESCRIT$(CI%), 3, 26)
          Call REPLA(FF1$, CSTRING$(MKS$(CANTX), 4, 10, 1, 2), 29, 10)
            fex = FF%
          Call REPLA(FF1$, FECHATEX$(fex), 41, 8)
          Call REPLA(FF1$, MKD$(CDbl(CANTX)), 49, 8)
          Call REPLA(FF1$, MKI$(FF%), 57, 2)
          Call REPLA(FF1$, MKI$(CI%), 59, 2)
          JJ& = JJ& + 1
          Call GRAREG("ORFAEMI", FF1$, JJ&)
       End If
    Next U&
    '
    Call SETULTREG("ORFAEMI", JJ&)
    List6.ListIndex = ILISTA&
    Screen.MousePointer = 1
    If JJ& < 1 Then
       Call COMUNI("No Hay Ordenes de Fabricación a Emitir\para el Programa Elegido.")
       Exit Sub
    End If
    '
    Call COMUNI("Tabla Generada con " & JJ& & " Registros")
    '
End Sub

Sub PLACOMP()
    If TRIM$(Label15) = "" Or TRIM$(Label17) = "" Then
       Call MENSERR(24, "Falta Re-Calcular")
       Exit Sub
    End If
    '
    If List4.ListCount < 1 Then
       Call MENSERR(24, "No Hay Necesidad de Componentes")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    ILISTA& = List4.ListIndex
    JJ& = 0
    REBLA$ = REGBLAN$("PLACOMP")
    For U& = 1 To List4.ListCount
       TTXX$ = List4.List(U& - 1)
       FALTA# = XVALO(Mid$(TTXX$, 34, 9))
       Code$ = TRIM$(Mid$(TTXX$, 1, 15))
       CANTNEC# = XVALO(Mid$(TTXX$, 16, 9))
       STOACT# = XVALO(Mid$(TTXX$, 25, 9))
       OCPEN# = XVALO(Mid$(TTXX$, 43, 9))
       DISPON# = STOACT# + OCPEN# - CANTNEC#
       '
       If DISPON# < 0 Then
          XX$ = REBLA$
          CI% = CODINT%(Code$)
          FF% = HOY(HOS$)
          PERIO$ = "INMEDIATO"
          CANTX = Abs(DISPON#)
          PHB% = PROVHABI%(CI%) ' PROVEEDOR HABITUAL
          MONEI% = MONECOSTO%(CI%)
          COSUDOL = COSUNI(CI%)
          COSUN = COSUDOL * TICAMONE(MONEI%)
          PHB% = PROVHABI%(CI%) ' PROVEEDOR HABITUAL
          '
          Z$ = REGBLAN$("PLACOMP")
          Call REPLA(Z$, MKI$(CI%), 1, 2)
          Call REPLA(Z$, MKI$(FF%), 3, 2)
          Call REPLA(Z$, PERIO$, 5, 16)
          Call REPLA(Z$, MKS$(CANTX), 21, 4)
          Call REPLA(Z$, MKS$(0), 25, 4) ' tot.progra.por item
          Call REPLA(Z$, MKS$(0), 29, 4) ' filler
          Call REPLA(Z$, MKI$(PHB%), 33, 2)
          Call REPLA(Z$, MKS$(COSUN), 35, 4)
          Call REPLA(Z$, MKD$(COSUN * CANTX), 39, 8)
          '
          JJ& = JJ& + 1
          Call GRAREG("PLACOMP", Z$, JJ&)
          '
       End If
       '
    Next U&
    '
    Call SETULTREG("PLACOMP", JJ&)
    Screen.MousePointer = 1
    If JJ& < 1 Then
       Call COMUNI("No Hay Necesidad de Componentes\para el Programa Elegido.")
       Exit Sub
    End If
    '
    Call COMUNI("Tabla Generada con " & JJ& & " Registros")
    '
    '
End Sub


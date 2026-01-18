VERSION 5.00
Begin VB.Form PREPATI07 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preparacion de Tintas"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Tintas e Insumos Trazables"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2265
      Left            =   210
      TabIndex        =   21
      Top             =   1995
      Width           =   8730
      Begin VB.CommandButton Command18 
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
         Index           =   0
         Left            =   2520
         TabIndex        =   41
         Top             =   525
         Width           =   175
      End
      Begin VB.CommandButton Command19 
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
         Index           =   0
         Left            =   4410
         TabIndex        =   40
         Top             =   525
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   0
         Left            =   7455
         TabIndex        =   39
         Top             =   525
         Width           =   1065
      End
      Begin VB.CommandButton Command18 
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
         Index           =   1
         Left            =   2520
         TabIndex        =   38
         Top             =   840
         Width           =   175
      End
      Begin VB.CommandButton Command19 
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
         Index           =   1
         Left            =   4410
         TabIndex        =   37
         Top             =   840
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   7455
         TabIndex        =   36
         Top             =   840
         Width           =   1065
      End
      Begin VB.CommandButton Command18 
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
         Index           =   2
         Left            =   2520
         TabIndex        =   35
         Top             =   1155
         Width           =   175
      End
      Begin VB.CommandButton Command19 
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
         Index           =   2
         Left            =   4410
         TabIndex        =   34
         Top             =   1155
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   7455
         TabIndex        =   33
         Top             =   1155
         Width           =   1065
      End
      Begin VB.CommandButton Command18 
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
         Index           =   3
         Left            =   2520
         TabIndex        =   32
         Top             =   1470
         Width           =   175
      End
      Begin VB.CommandButton Command19 
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
         Index           =   3
         Left            =   4410
         TabIndex        =   31
         Top             =   1470
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   7455
         TabIndex        =   30
         Top             =   1470
         Width           =   1065
      End
      Begin VB.CommandButton Command2 
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
         Index           =   0
         Left            =   5145
         TabIndex        =   29
         Top             =   525
         Width           =   175
      End
      Begin VB.CommandButton Command2 
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
         Index           =   1
         Left            =   5145
         TabIndex        =   28
         Top             =   840
         Width           =   175
      End
      Begin VB.CommandButton Command2 
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
         Index           =   2
         Left            =   5145
         TabIndex        =   27
         Top             =   1155
         Width           =   175
      End
      Begin VB.CommandButton Command2 
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
         Index           =   3
         Left            =   5145
         TabIndex        =   26
         Top             =   1470
         Width           =   175
      End
      Begin VB.CommandButton Command2 
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
         Index           =   4
         Left            =   5145
         TabIndex        =   25
         Top             =   1785
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   7455
         TabIndex        =   24
         Top             =   1785
         Width           =   1065
      End
      Begin VB.CommandButton Command19 
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
         Index           =   4
         Left            =   4410
         TabIndex        =   23
         Top             =   1785
         Width           =   175
      End
      Begin VB.CommandButton Command18 
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
         Index           =   4
         Left            =   2520
         TabIndex        =   22
         Top             =   1785
         Width           =   175
      End
      Begin VB.Label Label39 
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
         Index           =   0
         Left            =   840
         TabIndex        =   77
         ToolTipText     =   "Doble-Click Para Lectura Código de Barras"
         Top             =   525
         Width           =   1755
      End
      Begin VB.Label Label40 
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
         Index           =   0
         Left            =   3045
         TabIndex        =   76
         ToolTipText     =   "Doble-Click Para Lectura Código de Barras"
         Top             =   525
         Width           =   1440
      End
      Begin VB.Label Label41 
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
         Index           =   0
         Left            =   5550
         TabIndex        =   75
         Top             =   525
         Width           =   1050
      End
      Begin VB.Label Label42 
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
         Index           =   0
         Left            =   6825
         TabIndex        =   74
         Top             =   525
         Width           =   420
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6150
         TabIndex        =   73
         Top             =   300
         Width           =   1965
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -1155
         TabIndex        =   72
         Top             =   555
         Width           =   1965
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1365
         TabIndex        =   71
         Top             =   300
         Width           =   1965
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5040
         TabIndex        =   70
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label47 
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
         Height          =   300
         Left            =   6050
         TabIndex        =   69
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label Label39 
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
         Index           =   1
         Left            =   840
         TabIndex        =   68
         ToolTipText     =   "Código"
         Top             =   840
         Width           =   1755
      End
      Begin VB.Label Label40 
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
         Index           =   1
         Left            =   3045
         TabIndex        =   67
         ToolTipText     =   "Código"
         Top             =   840
         Width           =   1440
      End
      Begin VB.Label Label41 
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
         Index           =   1
         Left            =   5550
         TabIndex        =   66
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label42 
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
         Index           =   1
         Left            =   6825
         TabIndex        =   65
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label39 
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
         Index           =   2
         Left            =   840
         TabIndex        =   64
         ToolTipText     =   "Código"
         Top             =   1155
         Width           =   1755
      End
      Begin VB.Label Label40 
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
         Index           =   2
         Left            =   3045
         TabIndex        =   63
         ToolTipText     =   "Código"
         Top             =   1155
         Width           =   1440
      End
      Begin VB.Label Label41 
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
         Index           =   2
         Left            =   5550
         TabIndex        =   62
         Top             =   1155
         Width           =   1050
      End
      Begin VB.Label Label42 
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
         Index           =   2
         Left            =   6825
         TabIndex        =   61
         Top             =   1155
         Width           =   420
      End
      Begin VB.Label Label39 
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
         Index           =   3
         Left            =   840
         TabIndex        =   60
         ToolTipText     =   "Código"
         Top             =   1470
         Width           =   1755
      End
      Begin VB.Label Label40 
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
         Index           =   3
         Left            =   3045
         TabIndex        =   59
         ToolTipText     =   "Código"
         Top             =   1470
         Width           =   1440
      End
      Begin VB.Label Label41 
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
         Index           =   3
         Left            =   5550
         TabIndex        =   58
         Top             =   1470
         Width           =   1050
      End
      Begin VB.Label Label42 
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
         Index           =   3
         Left            =   6825
         TabIndex        =   57
         Top             =   1470
         Width           =   420
      End
      Begin VB.Label Label21 
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
         Index           =   0
         Left            =   2835
         TabIndex        =   56
         Top             =   525
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label21 
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
         Index           =   1
         Left            =   2835
         TabIndex        =   55
         Top             =   840
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label21 
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
         Index           =   2
         Left            =   2835
         TabIndex        =   54
         Top             =   1155
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label21 
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
         Index           =   3
         Left            =   2835
         TabIndex        =   53
         Top             =   1470
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label52 
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
         Index           =   0
         Left            =   4830
         TabIndex        =   52
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   525
         Width           =   330
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3570
         TabIndex        =   51
         Top             =   300
         Width           =   1545
      End
      Begin VB.Label Label52 
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
         Index           =   1
         Left            =   4830
         TabIndex        =   50
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   840
         Width           =   330
      End
      Begin VB.Label Label52 
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
         Index           =   2
         Left            =   4830
         TabIndex        =   49
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   1155
         Width           =   330
      End
      Begin VB.Label Label52 
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
         Index           =   3
         Left            =   4830
         TabIndex        =   48
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   1470
         Width           =   330
      End
      Begin VB.Label Label52 
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
         Index           =   4
         Left            =   4830
         TabIndex        =   47
         ToolTipText     =   "Doble-Click para Blanquear"
         Top             =   1785
         Width           =   330
      End
      Begin VB.Label Label21 
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
         Index           =   4
         Left            =   2835
         TabIndex        =   46
         Top             =   1785
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label42 
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
         Index           =   4
         Left            =   6825
         TabIndex        =   45
         Top             =   1785
         Width           =   420
      End
      Begin VB.Label Label41 
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
         Index           =   4
         Left            =   5550
         TabIndex        =   44
         Top             =   1785
         Width           =   1050
      End
      Begin VB.Label Label40 
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
         Index           =   4
         Left            =   3045
         TabIndex        =   43
         ToolTipText     =   "Código"
         Top             =   1785
         Width           =   1440
      End
      Begin VB.Label Label39 
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
         Index           =   4
         Left            =   840
         TabIndex        =   42
         ToolTipText     =   "Código"
         Top             =   1785
         Width           =   1755
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   3360
      Left            =   9135
      ScaleHeight     =   3300
      ScaleWidth      =   1320
      TabIndex        =   14
      Top             =   0
      Width           =   1380
      Begin VB.CommandButton Command1 
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
         Picture         =   "PREPATI07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Salir"
         Top             =   250
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   18
         Top             =   7455
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command3 
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
         Picture         =   "PREPATI07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   945
         Width           =   650
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Kdx"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         Picture         =   "PREPATI07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   1830
         Width           =   650
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
         Height          =   650
         Left            =   120
         Picture         =   "PREPATI07.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Configuracion de Funcionamiento"
         Top             =   2490
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "PREPATI07.frx":0BA0
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   7770
         Width           =   1515
      End
   End
   Begin VB.CommandButton Command5 
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
      Left            =   2835
      TabIndex        =   13
      Top             =   315
      Width           =   175
   End
   Begin VB.CommandButton Command4 
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
      Left            =   5040
      TabIndex        =   12
      Top             =   315
      Width           =   175
   End
   Begin VB.CommandButton Command7 
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
      Left            =   1575
      TabIndex        =   11
      Top             =   1470
      Width           =   175
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1050
      TabIndex        =   10
      Top             =   1470
      Width           =   540
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFD0&
      Caption         =   "Fechas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   7245
      TabIndex        =   3
      Top             =   210
      Width           =   1695
      Begin VB.TextBox Text12 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   315
         TabIndex        =   7
         Top             =   1155
         Width           =   1065
      End
      Begin VB.CommandButton Command6 
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
         Left            =   1365
         TabIndex        =   6
         Top             =   1155
         Width           =   175
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   315
         TabIndex        =   5
         Top             =   525
         Width           =   1065
      End
      Begin VB.CommandButton Command8 
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
         Left            =   1365
         TabIndex        =   4
         Top             =   525
         Width           =   175
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimiento:"
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
         Left            =   105
         TabIndex        =   9
         Top             =   945
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricación:"
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
         Left            =   105
         TabIndex        =   8
         Top             =   315
         Width           =   1335
      End
   End
   Begin VB.TextBox Text17 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   5250
      TabIndex        =   2
      Top             =   1470
      Width           =   1800
   End
   Begin VB.CommandButton Command13 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   9135
      Picture         =   "PREPATI07.frx":2AC2
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Aprueba Preparación"
      Top             =   3465
      Width           =   750
   End
   Begin VB.Label Label24 
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
      Left            =   4095
      TabIndex        =   90
      Top             =   315
      Width           =   1020
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2730
      TabIndex        =   89
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -630
      TabIndex        =   88
      ToolTipText     =   "Doble-Click para Blanquear"
      Top             =   360
      Width           =   1545
   End
   Begin VB.Label Label30 
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
      Height          =   495
      Left            =   1050
      TabIndex        =   87
      Top             =   735
      Width           =   3480
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descrip:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -210
      TabIndex        =   86
      Top             =   735
      Width           =   1125
   End
   Begin VB.Label Label32 
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
      Left            =   1050
      TabIndex        =   85
      ToolTipText     =   "Doble-Click Para Lectura Código de Barras"
      Top             =   315
      Width           =   1800
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Lote:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4305
      TabIndex        =   84
      Top             =   975
      Width           =   1335
   End
   Begin VB.Label Label51 
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
      Left            =   5670
      TabIndex        =   83
      ToolTipText     =   "Doble-Click para Blanquear"
      Top             =   945
      Width           =   1380
   End
   Begin VB.Label Label34 
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
      Left            =   1890
      TabIndex        =   82
      Top             =   1470
      Width           =   2175
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Depósito:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -840
      TabIndex        =   81
      Top             =   1515
      Width           =   1755
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicacion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3465
      TabIndex        =   80
      Top             =   1515
      Width           =   1755
   End
   Begin VB.Label Label49 
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
      Left            =   6000
      TabIndex        =   79
      Top             =   315
      Width           =   1050
   End
   Begin VB.Label Label48 
      Alignment       =   1  'Right Justify
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
      Height          =   300
      Left            =   3675
      TabIndex        =   78
      Top             =   360
      Width           =   2280
   End
   Begin VB.Label Label7 
      Height          =   225
      Left            =   315
      TabIndex        =   0
      Top             =   5355
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "PREPATI07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BRTEMP As Recordset
Dim BAUTOTRA%

Private Sub Command1_Click()
   '
   If UCase$(TRIM$(Label7)) = "AUTO" Then
      If COMALTER%("¿ Realmente Cancela la Asignación\Automática de Lotes Propuesta ?\\No, Conservarla\Si, Cancelarla") = 2 Then GoTo 10
      Exit Sub
   End If
   '
10 Call CIERRARCH("*.*")
   Hide
   '
End Sub

Sub Command10_Click()

    Command10.Enabled = False
    FIKARDEX09.Show 1
    Command10.Enabled = True
    
End Sub

Sub GRABACONSUM()
     '
     FF% = FECHANUM(Label24)
     CMO$ = "CF"    ' CONSUMO DE FABRICACION
     DOPRI$ = Label51
     DOSEC$ = Label32
     REFE$ = "Sol-Mat-Dep. p/" + Label32
     NUORIT% = 0
     '
'     For YX& = 0 To 2
'        NLOTX$ = TRIM$(Label23(YX&))
'        If NLOTX$ <> "" Then
'           CIXI% = CODINT%(Label27)
'           NUORIT% = NUORIT% + 1
'           CANTI = (-1) * XVALO(Label35(YX&))
'           LOTE$ = Label23(YX&)
'           DEPORIG% = DEPOLOTE%(CIXI%, Label23(YX&))
'           Call MOVISTOK(FF%, CIXI%, LOTE$, CMO$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, 0, 0, DEPORIG%, CANTI)                         '\\\OT-06-0244-WAR-20/06/06///
'        End If
'     Next YX&
'     '
     For YX& = 0 To 3
        CODEX$ = TRIM$(Label39(YX&))
        If CODEX$ <> "" Then
           CIXI% = CODINT%(CODEX$)
           NUORIT% = NUORIT% + 1
           LOTE$ = Label40(YX&)
           CANTI = (-1) * XVALO(Text1(YX&))
'           DEPORIG% = DEPOLOTE%(CIXI%, Label23(YX&))
           Call MOVISTOK(FF%, CIXI%, LOTE$, CMO$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, 0, 0, DepOrig%, CANTI)                                   '\\\OT-06-0244-WAR-20/06/06///
        End If
     Next YX&
     '
     For YX& = 4 To 7
        CODEX$ = TRIM$(Label39(YX&))
        If CODEX$ <> "" Then
           CIXI% = CODINT%(CODEX$)
           NUORIT% = NUORIT% + 1
           LOTE$ = ""
           CANTI = (-1) * XVALO(Text1(YX&))
           DepOrig% = DEPOPRE%(CIXI%)
           Call MOVISTOK(FF%, CIXI%, LOTE$, CMO$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, 0, 0, DepOrig%, CANTI)                                   '\\\OT-06-0244-WAR-20/06/06///
        End If
     Next YX&
     '
End Sub
'
Sub PRIVALEMA()
     '
     FORIPRE$ = TRIM$(Control$("", "FORCONSU"))
     If FORIPRE$ = "" Then
        Call MENSERR(24, "Falta Definir Formulario Vale de Materiales")
        Exit Sub
     End If
     '
     If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
           RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
           Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Vale de Consumo de Materiales:\  \'" + RMCC$ + "'.")
           Exit Sub
        End If
     End If
     '
     Screen.MousePointer = 11
     FECHDOC$ = Label24
     NUMEDOC$ = Label51
     TIPODOC$ = "Vale de Consumo de Materiales"
     '
     'CODPARAM$(1) = "REF-MAT1": DOCPARAM$(1) = REFE$
     'CODPARAM$(2) = "DEPOSITO": DOCPARAM$(2) = DEPOSI$
     CAPARDOC% = 2
     '
     CODTABU1$(1) = "ORD-ITEM"
     CODTABU1$(2) = "CANTIDAD"
     CODTABU1$(3) = "COD-MAT"
     CODTABU1$(4) = "DES-MAT"
     CODTABU1$(5) = "NRO-LOTE"
     CODTABU1$(6) = "REF-MAT1"
     CACOLTAB1% = 6
     '
     DESTIDOC% = 0
     CAITAB1% = 0
     '
     For YX& = 0 To 2
'        NLOTX$ = TRIM$(Label23(YX&))
        If NLOTX$ <> "" Then
'           CIXI% = CODINT%(Label27)
           CAITAB1% = CAITAB1% + 1
           TETABU1$(1, CAITAB1%) = TRIM$(Str$(CAITAB1%))
'           TETABU1$(2, CAITAB1%) = Label35(YX&)
           TETABU1$(3, CAITAB1%) = Label27
           TETABU1$(4, CAITAB1%) = Label6
'           TETABU1$(5, CAITAB1%) = Label23(YX&)
'           TETABU1$(6, CAITAB1%) = UBILOTE$(CIXI%, Label23(YX&))
        End If
     Next YX&
     '
     For YX& = 0 To 3
        CODEX$ = TRIM$(Label39(YX&))
        If CODEX$ <> "" Then
           CIXI% = CODINT%(CODEX$)
           CAITAB1% = CAITAB1% + 1
           TETABU1$(1, CAITAB1%) = TRIM$(Str$(CAITAB1%))
           TETABU1$(2, CAITAB1%) = Text1(YX&)
           TETABU1$(3, CAITAB1%) = CODEX$
           TETABU1$(4, CAITAB1%) = DESCRIT0$(CIXI%)
           TETABU1$(5, CAITAB1%) = Label40(YX&)
           TETABU1$(6, CAITAB1%) = UBILOTE$(CIXI%, Label40(YX&))
        End If
     Next YX&
     '
     For YX& = 4 To 7
        CODEX$ = TRIM$(Label39(YX&))
        If CODEX$ <> "" Then
           CIXI% = CODINT%(CODEX$)
           CAITAB1% = CAITAB1% + 1
           TETABU1$(1, CAITAB1%) = TRIM$(Str$(CAITAB1%))
           TETABU1$(2, CAITAB1%) = Text1(YX&)
           TETABU1$(3, CAITAB1%) = CODEX$
           TETABU1$(4, CAITAB1%) = DESCRIT0$(CIXI%)
           TETABU1$(5, CAITAB1%) = ""
           TETABU1$(6, CAITAB1%) = UBISTO(CIXI%)
        End If
     Next YX&
     '
     Call PRINTDOC("FORCONSU")   ' Remito en Consignacion
     Screen.MousePointer = 1
     '
     Call CIERRARCH("*.*")
     '
End Sub

Private Sub Command13_Click()
    '\\\OT-08-0302-OD-27/05/08///
    If CODINT(Label32) < 1 Then
      Call MENSERR(24, " Falta Código ")
       GoTo 99
    End If
    '
    FEX% = FECHANUM(Label24)
    If FEX% < 1 Then
          Call MENSERR(24, "Fecha no Válida")
          GoTo 99
    End If
    '
    If XVALO(Label49) < 0.05 Then
       Call MENSERR(24, "Falta Completar Cantidad")
       GoTo 99
    End If
    '
    If Label51 = "" Then
      Call MENSERR(24, "Falta Completar el Lote")
      GoTo 99
    End If
    '
    If ECOARCH$("TADEPOSI", Chr$(XVALO(Text2))) = "" Then
      Call MENSERR(24, "Falta Completar el Deposito")
      GoTo 99
    End If
    '
    If Len(TRIM$(Text17)) < 1 Then
      Call MENSERR(24, "Falta Completar Ubicación")
      GoTo 99
    End If
    '
    FEX1% = FECHANUM(Text8)
    If FEX1% < 1 Then
          Call MENSERR(24, "Fecha de Fabricación no Válida")
          GoTo 99
    End If
    '
    If FECHANUM(Text12) > 0 Then
      FexVen% = FECHANUM(Text12)
      If FexVen% < FEX1% Then
        Call MENSERR(24, "Fecha de Vencimiento Mayor a Fecha Fabricación")
         GoTo 99
      End If
      HOII% = HOY(HOS$)
      If FexVen% < HOII% Then
        Call MENSERR(24, "Fecha de Vencimiento Menor a Fecha de Hoy")
      End If
    End If
    '/////////////////////////////////////////////////////////////
    For Index = 0 To 4
      
      CODIMP$ = Label39(Index)
      LOTEMP$ = TRIM$(Label40(Index))
      DEPOMP$ = Label52(Index)
      CANTMP$ = Text1(Index)
      SALDOMP$ = Label41(Index)
      
      If CODIMP$ = "" And LOTEMP$ = "" And DEPOMP$ = "" And CANTMP$ = "" And SALDOMP$ = "" Then Exit For
      
      If CODINT(CODIMP$) < 1 Then
        Call MENSERR(24, " Falta Código Insumo ")
         GoTo 99
      End If
    '
      If Len(LOTEMP$) < 1 Then
        Call MENSERR(24, " Falta Lote de Insumo ")
         GoTo 99
      End If
      
      If XVALO(DEPOMP$) < 1 Then
        Call MENSERR(24, " Falta Depósito de Insumo ")
         GoTo 99
      End If
      
      If XVALO(CANTMP$) < 1 Or XVALO(CANTMP$) > XVALO(SALDOMP$) Then
          Call MENSERR(24, " La Cantidad No Puede Superar el Saldo .")
          GoTo 99
      End If
    Next Index
    '
    'Crea Nuevo Lote
    '
    FECHA% = FEX%
    FechFab% = FEX1%
    FecVen% = FexVen%
    CIXI% = CODINT(Label32)
    LOTE$ = TRIM$(Label51)
    CANTID# = TRIM$(FORMATNUM$(XVALO(Label49), "F10.1"))
    DEPOX% = XVALO(Text2)
    Ubicx$ = TRIM$(Text17)
    '
    Call CREALOTIN(FECHA%, FechFab%, FecVen%, CIXI%, LOTE$, "", "", CANTID#, 0, 0, DEPOX%, Ubicx$)
    '
    'Graba Movimiento de Stocks
    '
    REFE$ = "Preparacion de Tintas"
    FECHA% = FEX1%
    CODT% = CODINT%(Label32)
    LOTE$ = TRIM$(Label51)
    DEPOSITO% = XVALO(Text2)
    Cantidad = XVALO(Label49)
    '
    Call MOVISTOK(FECHA%, CODT%, LOTE$, "PT", LOTE$, LOTE$, LOTE$, LOTE$, REFE$, 0, 0, 0, 0, DEPOSITO%, Cantidad)
    '
    'GRABA MOVIMINTO DE CONSUMO DE LAS TINTAS
    For Ind% = 0 To 4
       FECHA% = FEX%
       CODT1% = CODINT%(Label39(Ind%))
       LOTE1$ = TRIM$(Label40(Ind%))
       LOTE$ = TRIM$(Label51)
       REFE$ = "Mezcla de Tintas"
       DEPOSITO% = XVALO(Label52(Ind%))
       Cantidad = XVALO(Text1(Ind%)) * (-1) 'cantidad negativa para que descuente
       '
       If CODT1% > 0 Then
        If LOTE$ <> "" Then
         If DEPOSITO% > 0 Then
          If Abs(Cantidad) > 0 Then
            Call MOVISTOK(FECHA%, CODT1%, LOTE1$, "MT", LOTE$, LOTE$, LOTE$, LOTE$, REFE$, 0, 0, 0, 0, DEPOSITO%, Cantidad)
          End If
         End If
        End If
       End If
    Next Ind%
    
   '\\\OT-08-0302-OD-FIN///
   BLANFORMU
   For U& = 0 To 4
     Label39(U&) = ""
     Text1(U&) = ""
     Label42(U&) = ""
     LIMPINSUM (U&)
   Next U&
99 Screen.MousePointer = 1
End Sub

Sub Command14_Click()
    '\\\OT-06-0476-RG-07/12/06///
    '
    If UCase$(TRIM$(Label7)) = "AUTO" Then Exit Sub
    '
    OPOFTRA07.Show 1
    '\\\OT-06-0476-RG-FIN///
End Sub

Private Sub Command18_Click(Index As Integer)
   Call CommSel("MASTER", Command18(Index), Label39(Index))
'  vdevu1$ = Label39(Index)
'   If Label32 <> "" Then
'    Call SELECHO("MASTER")
'    vdevu$ = VALACT1$("MASTER")
'    'If vdevu1$ <> vdevu$ Then LIMPINSUM (Index)
'    If vdevu$ <> "" Then
'       CICOMPO% = CODINT%(vdevu$)
'       Label39(Index) = vdevu$
'       Label21(Index) = DESCRIT0$(CICOMPO%)
'       Label42(Index) = UNIMED(CICOMPO%)
'    End If
'   End If
End Sub

Private Sub Command19_Click(Index As Integer)
  If Label32 <> "" Then
     CIXI% = CODINT%(Label39(Index))
     If CIXI% > 0 Then
        If ESTRAZABLE(CIXI%) < 1 Then
           Call MENSERR(24, "Material No Trazable")
           Exit Sub
        End If
        Call ARMALILOT(CIXI%, 0, 0, 1)
        Call SELECHO("!LILOTSE")
        VDEVU$ = VALACT1$("!LILOTSE")
        VDEV2$ = VALACT2$("!LILOTSE")
        Label40(Index) = VDEVU$
    End If
  End If
End Sub

Private Sub Command2_Click(Index As Integer)
   Call SELECHO("TADEPOSI")
   NUMDEP% = XVALO(VALACT1$("TADEPOSI"))
   Label52(Index) = VALACT1$("TADEPOSI")
'   CIXI% = CODINT%(TRIM$(Label39(INDEX)))
'   IDLOT$ = TRIM$(Label40(INDEX))
'   Label41(INDEX) = SaldoLote(CIXI%, IDLOT$, NUMDEP%)
End Sub

Private Sub COMMAND3_Click()
   Command3.Enabled = False
   CUDI$ = UCase$(CurDir)
   PPXX% = InStr(CUDI$, "\WINDOWS")
   If PPXX% > 0 Then
      rutax$ = Left$(CUDI$, PPXX%) + "MANUALES\"
'      If EXISTE%(rutax$ + "TRAZAB04.PDF") > 0 Then
'         FORHLPDF.Pdf1.setShowToolbar (True)
'         On Error Resume Next
'         FORHLPDF.Pdf1.LOADFILE rutax$ + "TRAZAB04.pdf"
'         If Err Then GoTo 20
'         On Error GoTo 0
'         FORHLPDF.Show 1
'         GoTo 99
'      End If
   End If
   '
20 Call MENSERR(24, "Imposible Abrir Archivo de Ayuda")
   '
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Call CommSel("SELFECHA", Command4, Label24)
End Sub

Private Sub Command5_Click()
    '
'    If UCase$(TRIM$(Label7)) = "AUTO" Then Exit Sub
'    '
'    Call ACTECOR(2)     ' para seleccionar solo las activas lanzadas
'    Call CommSel("ECORDEP/Indice de Ordenes de Fabricación Pendientes", Command5, Label32)
'    If Label32 <> "" Then
'       On Error Resume Next
'       Text8.SetFocus
'       On Error GoTo 0
'    End If
    '
   Call SELECHO("MASTER")
   NA$ = VALACT1$("MASTER")
   If NA$ <> "" Then
      Label32 = NA$
      Label30 = VALACT2$("MASTER")
   End If

     End Sub
'
Function CARPRODOF$(NROOF$)
  '
  CARPRODOF$ = ""
  Call ABREORFAB
  '
  SQLX$ = " SELECT * FROM OPERFAB "
  SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NROOF$ & "' "
  Set OPFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
  '
  With OPFABTMP
    '
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      Exit Function
    End If
    On Error GoTo 0
    '
    CI% = !COD_INTE
    CARPRODOF$ = CODEXT$(CI%)
  End With
  '
End Function

Private Sub Command6_Click()
   Call CommSel("SELFECHA", Command4, Text12)
End Sub

Private Sub Command7_Click()
'     Call CommSel("TADEPOSI", Command7, Text2)
'     Label34 = ECOARCH$("TADEPOSI", Text2)
'   Call SELECHO("TADEPOSI")
'   Text2 = VALACT1$("TADEPOSI")
'   Label34 = VALACT2$("TADEPOSI")
    Call CommSel("TADEPOSI", Command7, Text2)
End Sub

Private Sub Command8_Click()
   Call CommSel("SELFECHA", Command4, Text8)
End Sub


Private Sub Form_Load()
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If INTERPRE% = 1 Then
       'Frame9.Visible = True
    End If
    '
    Call FRESHALL
    '
    If Left$(APLINVO$, 7) = "OFABRIC" Then
       Label32 = TRIM$(Mid$(APLINVO$, 8))
       If Label32 <> "" Then
          If ECOARCH$("ECORDEP", Label32) <> "" Then
            HOII% = HOY(HOS$)
            Label24 = HOS$
          End If
       End If
    End If
    '
End Sub


Private Sub Label22_DblClick()
    Label32 = ""
    Call Command5_Click
End Sub

Private Sub Label32_Change()
       NVM$ = TRIM$(ProNroComprobante("PT"))
       While Len(NVM$) < 6: NVM$ = "0" + NVM$: Wend
    Label51 = "PT-" + NVM$ + "-00"
'    Label17 = ""
'    Label20 = ""
'    'Label21 = ""
'    'Label33 = ""
'    HOII% = HOY(HOS$)
'    Label24 = HOS$
'    'Text8.Enabled = True
'    ' PONER EN BLANCO TODOS LOS RESTANTES CAMPOS Y LIST'
'    'List1.Clear
'    Label28 = ""
'    Label30 = ""
'    Text8 = NORFA$
'    'Text1 = ""
'    Text5 = ""
'    Text6 = ""
'    Text2 = ""
'    Text3 = ""
'    Label34 = ""
'    Label27 = ""
'    Label6 = ""
'    Label16 = ""
'    Text4 = ""
'    '
'    For KKU& = 0 To 2
''       Label23(KKU&) = ""
''       Text11(KKU&) = ""
''       Label12(KKU&) = ""
''       Label25(KKU&) = ""
'    Next KKU&
'    '
''    If UCase$(TRIM$(Label7)) = "AUTO" Then
''       Call CARLICONSU(1)
''       GoTo 99
''    End If
''    '
'    If NORFA$ <> "" Then
'         ' presenta datos de la orden de fabricación
'         Call ABREORFAB
'         SQLX$ = " SELECT * FROM ORDEFABR "
'         SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "'"
'         '
'         Dim ORDEFABTMP As ADODB.Recordset
'         Set ORDEFABTMP = New ADODB.Recordset
'         Set ORDEFABTMP = FLEXCONN.Execute(SQLX$)
'         '
'         With ORDEFABTMP
'           On Error Resume Next
'           .MoveFirst
'           If Err Then
'              On Error GoTo 0
'              Call MENSERR(24, "Orden de Fabricación no Válida. Consulte.")
'              Call FINAL("")
'           End If
'           On Error GoTo 0
'           CIXI% = !Cod_Inte
'           Label28 = CODEXT$(CIXI%)
'           Label30 = DESCRIT0$(CIXI%)
'              CAPENDE# = !CanProy - !CanApro: If CAPENDE < 0 Then CAPENDE = 0
'              If !Statuorf >= 3 Then CAPENDE# = 0
'           Label34 = FORMATNUM$(!CanProy, "F8.1")
'           Label37 = FORMATNUM$(!CanApro, "F8.1")
'           Label38 = FORMATNUM$(CAPENDE#, "F8.1")
'           '
'         End With
'         '
'    End If
'99  Screen.MousePointer = 1
    
End Sub
'
Private Sub Label32_DblClick()
    '
10  TTXX$ = InputBox$("Pinche sobre Barras", "Lectura Código de Barras")
    If Len(TTXX$) > 1 Then
       Label32.Caption = TRIM$(Left$(TTXX$, 16))
    End If
    '
End Sub

'
Sub LIMPLAB()

    Label27.Caption = ""
    Label6.Caption = ""
    Text4.TEXT = ""
  
End Sub

Sub LIMPTEXT()

    Text5 = ""
    Text3 = ""
    Text6 = ""
    Text8 = ""
    Text9 = ""
    Text14 = ""
    Text12 = ""
  
End Sub


Private Sub Label39_Change(Index As Integer)
    VDEVU$ = Label39(Index).Caption
    Call LIMPINSUM(Index)
    CICOMPO% = CODINT%(VDEVU$)
    Label21(Index) = DESCRIT0$(CICOMPO%)
    Label42(Index) = UNIMED(CICOMPO%)
End Sub

Private Sub Label39_DblClick(Index As Integer)
    '
10  TTXX$ = InputBox$("Pinche sobre Barras", "Lectura Código de Barras")
    If Len(TTXX$) > 1 Then
       Label39(Index).Caption = TRIM$(Left$(TTXX$, 16))
       LOTSE$ = TRIM$(Mid$(TTXX$, 17, 12))
       Label40(Index).Caption = LOTSE$
    End If
    '
End Sub

Private Sub Label40_Change(Index As Integer)
   Call MUESTRASEL(Index)
End Sub

Private Sub Label40_DblClick(Index As Integer)
   Call Label39_DblClick(Index)
End Sub

Private Sub Label41_Change(Index As Integer)
   Text1(Index) = TRIM$(Label41(Index))
End Sub

Private Sub Label48_DblClick()
   Call Command4_Click
End Sub

Private Sub Label52_Change(Index As Integer)
   Call MUESTRASEL(Index)
End Sub

Private Sub Label52_DblClick(Index As Integer)
   Call Command2_Click(Index)
End Sub

Private Sub Text1_Change(Index As Integer)
    SUMATRA = 0
    For KKI% = 0 To 3
       If Label39(KKI%) <> "" Then
           SUMATRA = SUMATRA + XVALO(Text1(KKI%))
         Else
           Text1(KKI%) = ""
       End If
    Next KKI%
    Label49 = TRIM$(FORMATNUM$(SUMATRA, "F10.1"))
End Sub

Private Sub text1_lostfocus(Index As Integer)
   If XVALO(Text1(Index)) <= XVALO(Label41(Index)) Then
     Text1(Index) = FORMATNUM(XVALO(Text1(Index)), "F10.1")
    Else
     Call MENSERR(24, "La Cantidad no Puede Superar el Saldo")
   End If
End Sub
Private Sub Text12_DblClick()
    Call Command6_Click
End Sub

Private Sub Text2_Change()
         Label34 = ""
         Label34 = ECOARCH$("TADEPOSI", Chr$(XVALO(Text2)))
End Sub

Private Sub Text2_DblClick()
   Call Command7_Click
End Sub

Private Sub Text8_DblClick()
   Call Command8_Click
End Sub
Sub MUESTRASEL(INDEX1)
   '
   CIXI% = CODINT%(TRIM$(Label39(INDEX1)))
   IDLOT$ = TRIM$(Label40(INDEX1))
   NUMDEP% = TRIM$(XVALO(Label52(INDEX1)))
   If ESTRAZABLE%(CIXI%) Then
       Label41(INDEX1) = FORMATNUM$(SaldoLote(CIXI%, IDLOT$, NUMDEP%), "F10.1")
     Else
       Label41(INDEX1) = ""
   End If
   '
End Sub
Sub LIMPINSUM(Index)
    Label40(Index) = ""
    Label41(Index) = ""
    Label52(Index) = ""
End Sub
Sub BLANFORMU()
  Label32 = ""
  Label24 = ""
  Label49 = ""
  Label30 = ""
  Label51 = ""
  Label34 = ""
  Text17 = ""
  Text2 = ""
  Text12 = ""
  Text8 = ""
End Sub

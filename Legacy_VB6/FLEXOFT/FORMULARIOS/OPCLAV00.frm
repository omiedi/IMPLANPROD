VERSION 5.00
Begin VB.Form OPCLAV00 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Resúmenes Mensuales de Ventas"
   ClientHeight    =   7920
   ClientLeft      =   1650
   ClientTop       =   300
   ClientWidth     =   8715
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command11 
      Caption         =   "Cuentas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7665
      Picture         =   "OPCLAV00.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   70
      ToolTipText     =   "Archivo Maestro de Cuentas Contables"
      Top             =   6195
      Width           =   855
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ajustes de Cuenta Corriente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   210
      TabIndex        =   60
      Top             =   6930
      Width           =   7260
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
         Height          =   285
         Index           =   10
         Left            =   1995
         TabIndex        =   64
         Text            =   " "
         Top             =   420
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   10
         Left            =   3150
         TabIndex        =   63
         Top             =   420
         Width           =   150
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   300
         Left            =   25
         TabIndex        =   61
         Top             =   315
         Width           =   1905
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Left            =   0
            TabIndex        =   62
            Top             =   105
            Width           =   1905
         End
      End
      Begin VB.Label Label4 
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
         Index           =   10
         Left            =   3465
         TabIndex        =   65
         Top             =   420
         Width           =   3585
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Clasificación de Débitos y Créditos Varios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   210
      TabIndex        =   50
      Top             =   5250
      Width           =   7260
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
         Height          =   285
         Index           =   11
         Left            =   1995
         TabIndex        =   68
         Text            =   " "
         Top             =   1050
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   3150
         TabIndex        =   67
         Top             =   1050
         Width           =   150
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   3150
         TabIndex        =   57
         Top             =   735
         Width           =   150
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
         Height          =   285
         Index           =   9
         Left            =   1995
         TabIndex        =   56
         Text            =   " "
         Top             =   735
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   3150
         TabIndex        =   55
         Top             =   420
         Width           =   150
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
         Height          =   285
         Index           =   8
         Left            =   1995
         TabIndex        =   54
         Text            =   " "
         Top             =   420
         Width           =   1170
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   950
         Left            =   25
         TabIndex        =   51
         Top             =   480
         Width           =   1905
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Left            =   0
            TabIndex        =   66
            Top             =   630
            Width           =   1905
         End
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Left            =   0
            TabIndex        =   53
            Top             =   315
            Width           =   1905
         End
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Left            =   0
            TabIndex        =   52
            Top             =   0
            Width           =   1905
         End
      End
      Begin VB.Label Label4 
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
         Index           =   11
         Left            =   3465
         TabIndex        =   69
         Top             =   1050
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Index           =   9
         Left            =   3465
         TabIndex        =   59
         Top             =   735
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Index           =   8
         Left            =   3465
         TabIndex        =   58
         Top             =   420
         Width           =   3585
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cuentas Contables I.V.A."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   210
      TabIndex        =   4
      Top             =   105
      Width           =   7260
      Begin VB.TextBox Text2 
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
         Left            =   1995
         TabIndex        =   10
         Text            =   " "
         Top             =   420
         Width           =   1170
      End
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3150
         TabIndex        =   9
         Top             =   420
         Width           =   150
      End
      Begin VB.TextBox Text2 
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
         Left            =   1995
         TabIndex        =   8
         Text            =   " "
         Top             =   840
         Width           =   1170
      End
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3150
         TabIndex        =   7
         Top             =   840
         Width           =   150
      End
      Begin VB.TextBox Text2 
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
         Left            =   1995
         TabIndex        =   6
         Text            =   " "
         Top             =   1260
         Width           =   1170
      End
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3150
         TabIndex        =   5
         Top             =   1260
         Width           =   150
      End
      Begin VB.Label Label3 
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
         Left            =   3465
         TabIndex        =   41
         Top             =   1260
         Width           =   3585
      End
      Begin VB.Label Label3 
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
         Left            =   3465
         TabIndex        =   40
         Top             =   840
         Width           =   3585
      End
      Begin VB.Label Label3 
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
         Left            =   3465
         TabIndex        =   39
         Top             =   420
         Width           =   3585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta IVA Ventas:"
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
         TabIndex        =   13
         Top             =   450
         Width           =   1905
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta IVA R.N.I:"
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
         Left            =   0
         TabIndex        =   12
         Top             =   870
         Width           =   1905
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Retención IVA:"
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
         Left            =   0
         TabIndex        =   11
         Top             =   1290
         Width           =   1905
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Clasificación de Ventas Según Posición de IVA Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   210
      TabIndex        =   3
      Top             =   1995
      Width           =   7260
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   2535
         Left            =   25
         TabIndex        =   30
         Top             =   480
         Width           =   1905
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Left            =   0
            TabIndex        =   38
            Top             =   2205
            Width           =   1905
         End
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            TabIndex        =   37
            Top             =   1890
            Width           =   1905
         End
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Left            =   0
            TabIndex        =   36
            Top             =   1575
            Width           =   1905
         End
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Left            =   0
            TabIndex        =   35
            Top             =   945
            Width           =   1905
         End
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Left            =   0
            TabIndex        =   34
            Top             =   630
            Width           =   1905
         End
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Left            =   0
            TabIndex        =   33
            Top             =   315
            Width           =   1905
         End
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            TabIndex        =   32
            Top             =   0
            Width           =   1905
         End
         Begin VB.Label LABEL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta IVA Ventas:"
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
            Left            =   0
            TabIndex        =   31
            Top             =   1260
            Width           =   1905
         End
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
         Height          =   285
         Index           =   7
         Left            =   1995
         TabIndex        =   29
         Text            =   " "
         Top             =   2625
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   3150
         TabIndex        =   28
         Top             =   2625
         Width           =   150
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
         Height          =   285
         Index           =   6
         Left            =   1995
         TabIndex        =   27
         Text            =   " "
         Top             =   2310
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   3150
         TabIndex        =   26
         Top             =   2310
         Width           =   150
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
         Height          =   285
         Index           =   5
         Left            =   1995
         TabIndex        =   25
         Text            =   " "
         Top             =   1995
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   3150
         TabIndex        =   24
         Top             =   1995
         Width           =   150
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
         Height          =   285
         Index           =   4
         Left            =   1995
         TabIndex        =   23
         Text            =   " "
         Top             =   1680
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   3150
         TabIndex        =   22
         Top             =   1680
         Width           =   150
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
         Height          =   285
         Index           =   2
         Left            =   1995
         TabIndex        =   21
         Text            =   " "
         Top             =   1050
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3150
         TabIndex        =   20
         Top             =   1050
         Width           =   150
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
         Height          =   285
         Index           =   1
         Left            =   1995
         TabIndex        =   19
         Text            =   " "
         Top             =   735
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3150
         TabIndex        =   18
         Top             =   735
         Width           =   150
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
         Height          =   285
         Index           =   0
         Left            =   1995
         TabIndex        =   17
         Text            =   " "
         Top             =   420
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3150
         TabIndex        =   16
         Top             =   420
         Width           =   150
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
         Height          =   285
         Index           =   3
         Left            =   1995
         TabIndex        =   15
         Text            =   " "
         Top             =   1365
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   3150
         TabIndex        =   14
         Top             =   1365
         Width           =   150
      End
      Begin VB.Label Label4 
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
         Index           =   7
         Left            =   3465
         TabIndex        =   49
         Top             =   2625
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Index           =   6
         Left            =   3465
         TabIndex        =   48
         Top             =   2310
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Index           =   5
         Left            =   3465
         TabIndex        =   47
         Top             =   1995
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Left            =   3465
         TabIndex        =   46
         Top             =   1680
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Left            =   3465
         TabIndex        =   45
         Top             =   1050
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Left            =   3465
         TabIndex        =   44
         Top             =   735
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Left            =   3465
         TabIndex        =   43
         Top             =   420
         Width           =   3585
      End
      Begin VB.Label Label4 
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
         Left            =   3465
         TabIndex        =   42
         Top             =   1365
         Width           =   3585
      End
   End
   Begin VB.CommandButton Command7 
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
      Height          =   855
      Left            =   7665
      Picture         =   "OPCLAV00.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2205
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7665
      Picture         =   "OPCLAV00.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1260
      Width           =   855
   End
   Begin VB.CommandButton command1 
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
      Height          =   855
      Left            =   7665
      Picture         =   "OPCLAV00.frx":075E
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   7140
      Picture         =   "OPCLAV00.frx":08A8
      Top             =   7270
      Width           =   2010
   End
End
Attribute VB_Name = "OPCLAV00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command8_Click(Index As Integer)
    Call SELECHO("ECUECON")
    If VALACT1$("ECUECON") <> "" Then
        Text4(Index).TEXT = VALACT1$("ECUECON")
        Label4(Index) = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Form_Load()
   '
   CUEIVA$ = CONTROL$("SVD451S", "CUEIVA")
   Text2(0) = CUEIVA$
   Label3(0) = ECOARCH$("ECUECON", CUEIVA$)
   '
   CUEADU$ = CONTROL$("SVD451S", "CUEADU")
   Text2(1) = CUEADU$
   Label3(1) = ECOARCH$("ECUECON", CUEADU$)
   '
   CUERET$ = CONTROL$("SVD451S", "CUERET")
   Text2(2) = CUERET$
   Label3(2) = ECOARCH$("ECUECON", CUERET$)
   '
   Label2(0) = "Consumidor Final:"
   Label2(1) = "IVA Resp.Inscripto"
   Label2(2) = "Resp. No Inscripto"
   Label2(3) = "No Responsable IVA"
   Label2(4) = "IVA Exento"
   Label2(5) = "Comercio Exterior"
   Label2(6) = "Monotributo"
   Label2(7) = "Zona Promovida"
   '
   Label2(8) = "Cheques Rechazados"
   Label2(9) = "Diferencia de Cambio"
   Label2(10) = "Ajustes de Saldo"
   Label2(11) = "Otros Deb/Creditos"
   '
   For KKI% = 0 To 11
     Text4(KKI%).TEXT = CONTROL$("SVD451S", "CUEVE-" + TRIM$(Str$(KKI%)))
   Next KKI%
   '
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command2_Click()
    '
    CUEIVA$ = AJUSTI$(Text2(0).TEXT, 12)
    CUEADU$ = AJUSTI$(Text2(1).TEXT, 12)
    CUERET$ = AJUSTI$(Text2(2).TEXT, 12)
    '
    CUEIVA1% = 0: CUEADU1% = 0: CUERET1% = 0
    For U& = 1 To ULTREG&("ECUECON")
       X$ = REGLEIDO$("ECUECON", U&)
       If Left$(X$, 12) = CUEIVA$ Then CUEIVA1% = CVI(Mid$(X$, 41, 2)): DECIVA$ = Mid$(X$, 13, 28)
       If Left$(X$, 12) = CUEADU$ Then CUEADU1% = CVI(Mid$(X$, 41, 2)): DECADU$ = Mid$(X$, 13, 28)
       If Left$(X$, 12) = CUERET$ Then CUERET1% = CVI(Mid$(X$, 41, 2)): DECRET$ = Mid$(X$, 13, 28)
    Next U&
    '
    If CUEIVA1% < 1 Then
       Call MENSERR(24, "Falta Cuenta Contable I.V.A.")
       Text2(0).SetFocus
       Exit Sub
    End If
    '
    If CUEADU1% < 1 Then
       Call MENSERR(24, "Falta Cuenta Contable I.V.A. R.N.I.")
       Text2(1).SetFocus
       Exit Sub
    End If
    '
    If CUERET1% < 1 Then
       Call MENSERR(24, "Falta Cuenta Contable Retenciones I.V.A.")
       Text2(2).SetFocus
       Exit Sub
    End If
    '
    YA1% = 0: YA2% = 0: YA3% = 0
    For U& = 1 To ULTREG&("CUECON")
       X$ = REGLEIDO$("CUECON", U&)
       If CVI(Left$(X$, 2)) = CUEIVA1% Then YA1% = 1
       If CVI(Left$(X$, 2)) = CUEADU1% Then YA2% = 1
       If CVI(Left$(X$, 2)) = CUERET1% Then YA3% = 1
    Next U&
    '
    If YA1% < 1 Then
       X$ = REGBLAN$("CUECON")
       Call REPLA(X$, MKI$(CUEIVA1%), 1, 2)
       Call REPLA(X$, DECIVA$, 3, 29)
       Call REPLA(X$, CUEIVA$, 33, 12)
       Call REPLA(X$, CUEIVA$, 45, 12)
       Call REGAPP("CUECON", X$)
    End If
    '
    If YA2% < 1 Then
       X$ = REGBLAN$("CUECON")
       Call REPLA(X$, MKI$(CUEADU1%), 1, 2)
       Call REPLA(X$, DECADU$, 3, 29)
       Call REPLA(X$, CUEADU$, 33, 12)
       Call REPLA(X$, CUEADU$, 45, 12)
       Call REGAPP("CUECON", X$)
    End If
    '
    If YA3% < 1 Then
       X$ = REGBLAN$("CUECON")
       Call REPLA(X$, MKI$(CUERET1%), 1, 2)
       Call REPLA(X$, DECRET$, 3, 29)
       Call REPLA(X$, CUERET$, 33, 12)
       Call REPLA(X$, CUERET$, 45, 12)
       Call REGAPP("CUECON", X$)
    End If
    '
    Call GRACONTROL("SVD451S", "CUEIVA", CUEIVA$)
    Call GRACONTROL("SVD451S", "CUEADU", CUEADU$)
    Call GRACONTROL("SVD451S", "CUERET", CUERET$)
    '
    MOCALVE$ = "PIVACLI"
    Call GRACONTROL("REMEVENT", "MODCALVE", MOCALVE$)
    '
    If MOCALVE$ = "PIVACLI" Then
      For KKI% = 0 To 11
        CUEVENTA$(KKI%) = AJUSTI$(Text4(KKI%).TEXT, 12)
      Next KKI%
      '
      For U& = 1 To ULTREG&("ECUECON")
        X$ = REGLEIDO$("ECUECON", U&)
        For KKI% = 0 To 11
          If Left$(X$, 12) = CUEVENTA$(KKI%) Then COCUEVEN%(KKI%) = CVI(Mid$(X$, 41, 2)): DECIVA$ = Mid$(X$, 13, 28)
        Next KKI%
      Next U&
      '
      For KKI% = 0 To 11
        If COCUEVEN%(KKI%) > 0 Then
            Call GRACONTROL("SVD451S", "CUEVE-" + TRIM$(Str$(KKI%)), CUEVENTA$(KKI%))
          Else
            Call MENSERR(24, "Faltan Cuentas Contables de Venta.")
            Text4(KKI%).SetFocus
            Exit Sub
        End If
      Next KKI%
    End If
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
    '
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Private Sub Text2_Change(Index As Integer)
    DECUE$ = ECOARCH("ECUECON", TRIM$(Text2(Index).TEXT))
    If DECUE$ <> "" Then
        Label3(Index) = DECUE$
      Else
        Label3(Index) = ""
    End If
End Sub

Private Sub Text2_DblClick(Index As Integer)
    Call SELECHO("ECUECON")
    If VALACT1$("ECUECON") <> "" Then
        Text2(Index).TEXT = VALACT1$("ECUECON")
        Label3(Index) = VALACT2$("ECUECON")
    End If
End Sub
'
Private Sub Command4_Click(Index As Integer)
    Call SELECHO("ECUECON")
    If VALACT1$("ECUECON") <> "" Then
        Text2(Index).TEXT = VALACT1$("ECUECON")
        Label3(Index) = VALACT2$("ECUECON")
    End If
End Sub
'
Private Sub Text4_Change(Index As Integer)
    DECUE$ = ECOARCH("ECUECON", TRIM$(Text4(Index).TEXT))
    If DECUE$ <> "" Then
        Label4(Index) = DECUE$
      Else
        Label4(Index) = ""
    End If
End Sub

Private Sub Text4_DblClick(Index As Integer)
    Call SELECHO("ECUECON")
    If VALACT1$("ECUECON") <> "" Then
        Text4(Index).TEXT = VALACT1$("ECUECON")
        Label4(Index) = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Command11_Click()
   '
   Command11.Enabled = False
   RETEJ% = EJER%
   EJER% = 1
   Call CARCUEINGAS
   EJER% = RETEJ%
   EJU$ = EJACT$
   Command11.Enabled = True
   '
End Sub



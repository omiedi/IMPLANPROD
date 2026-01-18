VERSION 5.00
Begin VB.Form OPFACTU01 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Facturación"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
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
   ScaleHeight     =   6780
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Puertos de Impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1300
      Left            =   5500
      TabIndex        =   40
      Top             =   210
      Width           =   5160
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   420
         Width           =   3795
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1155
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   840
         Width           =   3795
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remitos:"
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
         Left            =   -105
         TabIndex        =   42
         Top             =   920
         Width           =   1170
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Facturas:"
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
         Left            =   -210
         TabIndex        =   41
         Top             =   500
         Width           =   1275
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Formularios de Impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6470
      Left            =   210
      TabIndex        =   15
      Top             =   210
      Width           =   5160
      Begin VB.Frame Frame14 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Descuento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   3465
         TabIndex        =   60
         Top             =   5355
         Width           =   1485
         Begin VB.OptionButton Option14 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Porciento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   62
            Top             =   630
            Width           =   1170
         End
         Begin VB.OptionButton Option13 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pr.Neto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   61
            Top             =   315
            Value           =   -1  'True
            Width           =   1170
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Formulario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   1995
         TabIndex        =   57
         Top             =   5340
         Width           =   1380
         Begin VB.OptionButton Option12 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Continuo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   59
            Top             =   315
            Value           =   -1  'True
            Width           =   1100
         End
         Begin VB.OptionButton Option11 
            BackColor       =   &H00FFFFC0&
            Caption         =   "H.Suelta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   58
            Top             =   630
            Width           =   1100
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   210
         TabIndex        =   54
         Top             =   5340
         Width           =   1695
         Begin VB.OptionButton Option10 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Auto-Numera"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   56
            Top             =   630
            Width           =   1420
         End
         Begin VB.OptionButton Option9 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pre-Impresa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   210
            TabIndex        =   55
            Top             =   315
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Formulario de Remitos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1150
         Index           =   3
         Left            =   210
         TabIndex        =   34
         Top             =   4095
         Width           =   4740
         Begin VB.CommandButton Command3 
            Caption         =   "Obs."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   68
            ToolTipText     =   "Permite Definir Observaciones Normalizadas Habituales"
            Top             =   280
            Width           =   735
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Abrir"
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
            Index           =   3
            Left            =   3780
            Picture         =   "OPFACTU01.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   260
            Width           =   750
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
            Height          =   300
            Index           =   3
            Left            =   3465
            TabIndex        =   36
            Top             =   260
            Width           =   175
         End
         Begin VB.TextBox Text1 
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
            Index           =   3
            Left            =   2415
            TabIndex        =   35
            Top             =   260
            Width           =   1065
         End
         Begin VB.Label DEFOR1 
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
            Height          =   280
            Index           =   3
            Left            =   150
            TabIndex        =   39
            Top             =   685
            Width           =   3480
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Formulario:"
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
            TabIndex        =   38
            Top             =   335
            Width           =   2325
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Facturación de Exportación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1150
         Index           =   2
         Left            =   210
         TabIndex        =   28
         Top             =   2835
         Width           =   4740
         Begin VB.CommandButton Command3 
            Caption         =   "Obs."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   67
            ToolTipText     =   "Permite Definir Observaciones Normalizadas Habituales"
            Top             =   280
            Width           =   735
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Abrir"
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
            Index           =   2
            Left            =   3780
            Picture         =   "OPFACTU01.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   260
            Width           =   750
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
            Height          =   300
            Index           =   2
            Left            =   3465
            TabIndex        =   30
            Top             =   260
            Width           =   175
         End
         Begin VB.TextBox Text1 
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
            Index           =   2
            Left            =   2415
            TabIndex        =   29
            Top             =   260
            Width           =   1065
         End
         Begin VB.Label DEFOR1 
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
            Height          =   280
            Index           =   2
            Left            =   150
            TabIndex        =   33
            Top             =   685
            Width           =   3480
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Formulario:"
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
            TabIndex        =   32
            Top             =   335
            Width           =   2325
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Facturación Tipo ""A"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1150
         Index           =   0
         Left            =   210
         TabIndex        =   22
         Top             =   315
         Width           =   4740
         Begin VB.CommandButton Command3 
            Caption         =   "Obs."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   65
            ToolTipText     =   "Permite Definir Observaciones Normalizadas Habituales"
            Top             =   280
            Width           =   735
         End
         Begin VB.TextBox Text1 
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
            Index           =   0
            Left            =   2415
            TabIndex        =   25
            Top             =   260
            Width           =   1065
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
            Height          =   300
            Index           =   0
            Left            =   3465
            TabIndex        =   24
            Top             =   260
            Width           =   175
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Abrir"
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
            Index           =   0
            Left            =   3780
            Picture         =   "OPFACTU01.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   260
            Width           =   750
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Formulario:"
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
            TabIndex        =   27
            Top             =   335
            Width           =   2325
         End
         Begin VB.Label DEFOR1 
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
            Left            =   150
            TabIndex        =   26
            Top             =   685
            Width           =   3480
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Facturación Tipo ""B"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1150
         Index           =   1
         Left            =   210
         TabIndex        =   16
         Top             =   1575
         Width           =   4740
         Begin VB.CommandButton Command3 
            Caption         =   "Obs."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   66
            ToolTipText     =   "Permite Definir Observaciones Normalizadas Habituales"
            Top             =   280
            Width           =   735
         End
         Begin VB.TextBox Text1 
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
            Index           =   1
            Left            =   2415
            TabIndex        =   19
            Top             =   260
            Width           =   1065
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
            Height          =   300
            Index           =   1
            Left            =   3465
            TabIndex        =   18
            Top             =   260
            Width           =   175
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Abrir"
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
            Index           =   1
            Left            =   3780
            Picture         =   "OPFACTU01.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   260
            Width           =   750
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Formulario:"
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
            TabIndex        =   21
            Top             =   335
            Width           =   2325
         End
         Begin VB.Label DEFOR1 
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
            Height          =   280
            Index           =   1
            Left            =   150
            TabIndex        =   20
            Top             =   685
            Width           =   3480
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Controles Operativos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5000
      Left            =   5500
      TabIndex        =   3
      Top             =   1680
      Width           =   5160
      Begin VB.Frame Frame15 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   3490
         TabIndex        =   63
         Top             =   2700
         Width           =   1430
         Begin VB.CheckBox Check3 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Reempl. Original"
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
            Left            =   210
            TabIndex        =   64
            Top             =   420
            Width           =   1095
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Pre-Facturación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   2415
         TabIndex        =   51
         Top             =   1610
         Width           =   2520
         Begin VB.OptionButton Option8 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Nota de Venta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   53
            Top             =   315
            Value           =   -1  'True
            Width           =   2010
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Orden de Despacho"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   315
            TabIndex        =   52
            Top             =   630
            Width           =   2115
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Expedición - Depósito de Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Left            =   210
         TabIndex        =   46
         Top             =   4020
         Width           =   4695
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   1155
            TabIndex        =   48
            Text            =   " "
            Top             =   350
            Width           =   540
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
            Left            =   1680
            TabIndex        =   47
            Top             =   350
            Width           =   175
         End
         Begin VB.Label Label1 
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
            Height          =   280
            Left            =   1995
            TabIndex        =   50
            Top             =   350
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Depósito:"
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
            TabIndex        =   49
            Top             =   420
            Width           =   1065
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Verificaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   210
         TabIndex        =   43
         Top             =   2700
         Width           =   3180
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Límite de Crédito del Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   210
            TabIndex        =   45
            Top             =   315
            Width           =   1275
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Existencia de Piezas a Despachar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   1680
            TabIndex        =   44
            Top             =   315
            Width           =   1380
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   210
         TabIndex        =   9
         Top             =   420
         Width           =   3750
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            Caption         =   "Frame7"
            Height          =   750
            Left            =   1890
            TabIndex        =   12
            Top             =   210
            Width           =   1695
            Begin VB.OptionButton Option6 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Independiente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   105
               TabIndex        =   14
               Top             =   420
               Width           =   1590
            End
            Begin VB.OptionButton Option5 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Correlativa"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   105
               TabIndex        =   13
               Top             =   105
               Value           =   -1  'True
               Width           =   1380
            End
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Manual"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   420
            TabIndex        =   11
            Top             =   315
            Width           =   1380
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Automatica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   420
            TabIndex        =   10
            Top             =   630
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   4095
         TabIndex        =   7
         Top             =   420
         Width           =   855
         Begin VB.TextBox Text9 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   190
            TabIndex        =   8
            Text            =   "FC"
            Top             =   525
            Width           =   450
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Formato de Carga"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   210
         TabIndex        =   4
         Top             =   1610
         Width           =   2100
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Itemizado "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   420
            TabIndex        =   6
            Top             =   630
            Width           =   1590
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Normal"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   420
            TabIndex        =   5
            Top             =   315
            Value           =   -1  'True
            Width           =   1380
         End
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
      Left            =   10845
      Picture         =   "OPFACTU01.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1260
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
      Left            =   10845
      Picture         =   "OPFACTU01.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2205
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
      Left            =   10845
      Picture         =   "OPFACTU01.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10290
      Picture         =   "OPFACTU01.frx":1386
      Top             =   6150
      Width           =   2010
   End
End
Attribute VB_Name = "OPFACTU01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "FORFAC-A", Text1(0).TEXT)
    Call GRACONTROL("", "FORFAC-B", Text1(1).TEXT)
    Call GRACONTROL("", "FORFAC-E", Text1(2).TEXT)
    Call GRACONTROL("", "FOREMIFA", Text1(3).TEXT)
    Call GRACONTROL("", "DEPOEXPE", Text7.TEXT)
    Call GRACONTROL("", "CODFACTU", Text9.TEXT)
    '
    Call GRACONTROL(IDENTER$, "PORTFAC", Combo1.TEXT)
    Call GRACONTROL(IDENTER$, "PORTREM", Combo2.TEXT)
    '
    MONUMERA$ = "AUTO"
    If Option3.Value = True Then
        MONUMERA$ = "MANUAL"
    End If
    Call GRACONTROL("", "NUMFACTU", MONUMERA$)
    '
    MONUME$ = "NORMAL": FORCAFAC% = 1
    If Option2.Value = True Then
        MONUME$ = "ITEMIZADO": FORCAFAC% = 2
    End If
    Call GRACONTROL("FACTUMAI", "FORMACAR", MONUME$)
    '
    MONUME$ = "CORRELA": NUMECORR% = 1
    If Option6.Value = True Then
        MONUME$ = "INDEPEN": NUMECORR% = (-1)
    End If
    Call GRACONTROL("NUMEFACT", "NUMECORR", MONUME$)
    '
    PREFACTU$ = "PEDIDO"
    If Option7.Value = True Then
        PREFACTU$ = "ORDESPA"
    End If
    Call GRACONTROL("FACTURA", "PREFACTU", PREFACTU$)
    '
    NUMERA$ = "PREIMPRE"
    If Option10.Value = True Then NUMERA$ = "AUTONUME"
    Call GRACONTROL("FACTUREM", "IMPRENUM", NUMERA$)
    '
    TIFORMU$ = "CONTINUO"
    If Option11.Value = True Then TIFORMU$ = "HOJASUEL"
    Call GRACONTROL("FACTUREM", "FORMULAR", TIFORMU$)
    '
    MODESCUE$ = "PRENETO"
    If Option14.Value = True Then MODESCUE$ = "PORCIENTO"
    Call GRACONTROL("FACTUREM", "MODESCUE", MODESCUE$)
    '
    MONUME$ = "NO"
    If Check2.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("FACTURA", "VERLICRE", MONUME$)
    '
    MONUME$ = "NO"
    If Check3.Value = 1 Then MONUME$ = "SI"
    Call GRACONTROL("REMSTAN", "REPLAORI", MONUME$)
    '
    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completada !\Re-Inicie Ahora el Módulo de Facturación.")
    Call FINAL("")
    '
End Sub

Private Sub Command3_Click(Index As Integer)
    Command3(Index).Enabled = False
    Call TRALOFIL("LEYFAC", 1, MKI$(Index))
    VTB% = VENTABU%("LEYFAC/TITUPAN=" + DEFOR1(Index))
    If VTB% >= 0 Then
      JJ& = 0
      For U& = 1 To ULTREG&("LEYFAC")
        XX$ = REGLEIDO$("LEYFAC", U&)
        If CVI(Left$(XX$, 2)) <> Index Then
          JJ& = JJ& + 1
          Call GRAREG("LEYFAC", XX$, JJ&)
        End If
      Next U&
      For U& = 1 To ULTREG&("!LEYFAC")
        XX$ = REGLEIDO$("!LEYFAC", U&)
        Call REPLA(XX$, MKI$(Index), 1, 2)
        JJ& = JJ& + 1
        Call GRAREG("LEYFAC", XX$, JJ&)
      Next U&
      Call SETULTREG("LEYFAC", JJ&)
    End If
    Command3(Index).Enabled = True
End Sub

'
Private Sub Command4_Click(Index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1(Index).TEXT = VDEVU$
        DEFOR1(Index).Caption = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command6_Click(Index As Integer)
    MsgBox "Para Datos del Transporte Usar Parametros '-COMP'"
    Call CONECRUN("FLFORMS/" + TRIM$(Text1(Index).TEXT), "Formularios de Impresion")
End Sub

Private Sub Command8_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text3.TEXT), "Formularios de Impresion")
End Sub

Private Sub Form_Load()
    '
    J& = 0
    For U& = 1 To ULTREG&("SUCURSAL")
      X$ = REGLEIDO$("SUCURSAL", U&)
      CODEPO% = Asc(Mid$(X$, 64, 1))
      If CODEPO% > 0 Then
        If CODEPO% <> 99 Then
          If J& < 62 Then
            J& = J& + 1
            Y$ = REGBLAN$("TADEPOSI")
            Call REPLA(Y$, Chr$(CODEPO%), 1, 1)
            Call REPLA(Y$, Left$(X$, 63), 2, 63)
            Call GRAREG("TADEPOSI", Y$, J&)
          End If
        End If
      End If
    Next U&
    Call SETULTREG("TADEPOSI", J&)
    Call CIERRARCH("SUCURSAL")
    Call CIERRARCH("TADEPOSI")
    Screen.MousePointer = 1
    '
    Text1(0).TEXT = CONTROL$("", "FORFAC-A")
    DEFOR1(0).Caption = ECOARCH$("INDIFRM", Text1(0).TEXT)
    Text1(1).TEXT = CONTROL$("", "FORFAC-B")
    DEFOR1(1).Caption = ECOARCH$("INDIFRM", Text1(1).TEXT)
    Text1(2).TEXT = CONTROL$("", "FORFAC-E")
    DEFOR1(2).Caption = ECOARCH$("INDIFRM", Text1(2).TEXT)
    Text1(3).TEXT = CONTROL$("", "FOREMIFA")
    DEFOR1(3).Caption = ECOARCH$("INDIFRM", Text1(3).TEXT)
    '
    PFAC$ = CONTROL$(IDENTER$, "PORTFAC")
    Combo1.Clear
    Dim PRX As Printer
    On Error Resume Next
    Combo1.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo1.AddItem PRX.DeviceName
      If PRX.DeviceName = PFAC$ Then
        Combo1.TEXT = PFAC$
      End If
    Next
    '
    PREM$ = CONTROL$(IDENTER$, "PORTREM")
    Combo2.Clear
    On Error Resume Next
    Combo2.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo2.AddItem PRX.DeviceName
      If PRX.DeviceName = PREM$ Then
        Combo2.TEXT = PREM$
      End If
    Next
    
    Text7.TEXT = CONTROL$("", "DEPOEXPE")
    Label1.Caption = ECOARCH$("TADEPOSI", Chr$(Val(Text7.TEXT)))
    '
    If TRIM$(CONTROL$("", "NUMFACTU")) = "MANUAL" Then Option3.Value = True
    If CONTROL$("NUMEFACT", "NUMECORR") = "INDEPEN" Then Option6.Value = True
    If CONTROL$("FACTUMAI", "FORMACAR") = "ITEMIZADO" Then Option2.Value = True
    If CONTROL$("FACTUREM", "IMPRENUM") = "AUTONUME" Then Option10.Value = True
    If CONTROL$("FACTUREM", "FORMULAR") = "HOJASUEL" Then Option11.Value = True
    If CONTROL$("FACTUREM", "MODESCUE") = "PORCIENTO" Then Option14.Value = True
    '
    PREFACTU$ = CONTROL("FACTURA", "PREFACTU")
    If PREFACTU$ = "ORDESPA" Then Option7 = True
    '
    CMO$ = TRIM$(CONTROL$("", "CODFACTU")): If CMO$ <> "" Then Text9.TEXT = CMO$
    '
    If CONTROL$("FACTURA", "VERLICRE") = "SI" Then Check2.Value = 1
    '
    If CONTROL$("REMSTAN", "REPLAORI") = "SI" Then
       Check3.Value = 1
    End If
    '
End Sub

Private Sub Text1_DblClick(Index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1(Index).TEXT = VDEVU$
        DEFOR1(Index).Caption = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text4_Click()
    Text3.SetFocus
End Sub

Private Sub Text6_Click()
    Text5.SetFocus
End Sub

Private Sub Text8_Click()
    Text7.SetFocus
End Sub


Private Sub Text3_DBLCLICK()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text5_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text5.TEXT = VDEVU$
        Text6.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text7_DblClick()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Label1.Caption = VALACT2$("TADEPOSI")
    End If
End Sub
'
Private Sub Command9_Click()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
        Text7.TEXT = VDEVU$
        Label1.Caption = VALACT2$("TADEPOSI")
    End If
End Sub
'


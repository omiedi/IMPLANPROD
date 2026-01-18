VERSION 5.00
Begin VB.Form FORMFAC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command12 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   275
      Left            =   11590
      Picture         =   "FORMFAC.frx":0000
      TabIndex        =   60
      Top             =   25
      Width           =   275
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1905
      Left            =   0
      TabIndex        =   23
      Top             =   6300
      Width           =   10800
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   210
         ScaleHeight     =   75
         ScaleWidth      =   5205
         TabIndex        =   61
         Top             =   210
         Width           =   5265
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   62
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.TextBox PORIVA 
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
         Left            =   7560
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   " "
         Top             =   990
         Width           =   1065
      End
      Begin VB.TextBox PORDESCU 
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
         Left            =   7560
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   600
         Width           =   1065
      End
      Begin VB.TextBox Total 
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
         Left            =   9030
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   " "
         Top             =   990
         Width           =   1650
      End
      Begin VB.TextBox Total 
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
         Left            =   9030
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   " "
         Top             =   600
         Width           =   1650
      End
      Begin VB.TextBox Total 
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
         Left            =   9030
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   " "
         Top             =   210
         Width           =   1650
      End
      Begin VB.TextBox Total 
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
         Left            =   9030
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1380
         Width           =   1650
      End
      Begin VB.TextBox Eco21 
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
         Left            =   1890
         TabIndex        =   35
         Text            =   " "
         Top             =   1380
         Width           =   3585
      End
      Begin VB.TextBox TEXT21 
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
         Left            =   1155
         TabIndex        =   33
         Text            =   " "
         Top             =   1380
         Width           =   435
      End
      Begin VB.CommandButton Command21 
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
         Index           =   2
         Left            =   1575
         TabIndex        =   32
         Top             =   1380
         Width           =   180
      End
      Begin VB.TextBox Eco21 
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
         Left            =   1890
         TabIndex        =   31
         Text            =   " "
         Top             =   990
         Width           =   3585
      End
      Begin VB.TextBox TEXT21 
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
         Left            =   1155
         TabIndex        =   29
         Text            =   " "
         Top             =   990
         Width           =   435
      End
      Begin VB.CommandButton Command21 
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
         Index           =   1
         Left            =   1575
         TabIndex        =   28
         Top             =   990
         Width           =   180
      End
      Begin VB.TextBox Eco21 
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
         Left            =   1890
         TabIndex        =   27
         Text            =   " "
         Top             =   600
         Width           =   3585
      End
      Begin VB.TextBox TEXT21 
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
         Left            =   1155
         TabIndex        =   25
         Text            =   " "
         Top             =   600
         Width           =   435
      End
      Begin VB.CommandButton Command21 
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
         Index           =   0
         Left            =   1575
         TabIndex        =   24
         Top             =   600
         Width           =   180
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "%:"
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
         Left            =   8715
         TabIndex        =   47
         Top             =   990
         Width           =   225
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "%:"
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
         Left            =   8715
         TabIndex        =   46
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Documento:"
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
         Index           =   3
         Left            =   7035
         TabIndex        =   43
         Top             =   1380
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A.:"
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
         Index           =   2
         Left            =   5565
         TabIndex        =   42
         Top             =   990
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento:"
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
         Index           =   1
         Left            =   5565
         TabIndex        =   41
         Top             =   600
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Neto:"
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
         Index           =   0
         Left            =   7035
         TabIndex        =   40
         Top             =   210
         Width           =   1905
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vended:"
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
         Index           =   2
         Left            =   105
         TabIndex        =   34
         Top             =   1380
         Width           =   960
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C.Pago:"
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
         Index           =   1
         Left            =   105
         TabIndex        =   30
         Top             =   990
         Width           =   960
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "L.Precio:"
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
         Index           =   0
         Left            =   105
         TabIndex        =   26
         Top             =   600
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8200
      Left            =   0
      TabIndex        =   0
      Top             =   315
      Width           =   12000
      Begin VB.CommandButton Command11 
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
         Height          =   855
         Left            =   10920
         Picture         =   "FORMFAC.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Apertura de Nueva Cuenta de Cliente"
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Pendien"
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
         Left            =   10920
         Picture         =   "FORMFAC.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Remitos Pendientes de Facturación"
         Top             =   5145
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "C.Pago"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   730
         Left            =   10920
         Picture         =   "FORMFAC.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Condiciones de Pago"
         Top             =   6960
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Precios"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   10920
         Picture         =   "FORMFAC.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Listas de Precios de Venta"
         Top             =   6170
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "C.Cte."
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
         Left            =   10920
         Picture         =   "FORMFAC.frx":0EAA
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Resumen de Cuenta Corriente"
         Top             =   2625
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Notas"
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
         Left            =   10920
         Picture         =   "FORMFAC.frx":11B4
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Notas de Clientes"
         Top             =   3570
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
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
         Left            =   10920
         Picture         =   "FORMFAC.frx":14BE
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   315
         Width           =   855
      End
      Begin VB.CommandButton BOTNEXT 
         Caption         =   "Next"
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
         Left            =   10920
         Picture         =   "FORMFAC.frx":17C8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   735
         Width           =   855
      End
      Begin VB.TextBox Text11 
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
         Left            =   9030
         TabIndex        =   21
         Text            =   " "
         Top             =   1155
         Width           =   1650
      End
      Begin VB.TextBox Text10 
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
         Left            =   9030
         TabIndex        =   19
         Text            =   " "
         Top             =   735
         Width           =   1485
      End
      Begin VB.CommandButton Command10 
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
         Left            =   10500
         TabIndex        =   18
         Top             =   735
         Width           =   180
      End
      Begin VB.TextBox Text9 
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
         Left            =   9345
         TabIndex        =   16
         Text            =   " "
         Top             =   315
         Width           =   1170
      End
      Begin VB.CommandButton Command9 
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
         Left            =   10500
         TabIndex        =   15
         Top             =   315
         Width           =   180
      End
      Begin VB.TextBox Text7 
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
         Left            =   5565
         TabIndex        =   13
         Text            =   " "
         Top             =   735
         Width           =   1695
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
         Height          =   300
         Left            =   5565
         TabIndex        =   12
         Text            =   " "
         Top             =   1155
         Width           =   1695
      End
      Begin VB.CommandButton Command6 
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
         Left            =   7035
         TabIndex        =   11
         Top             =   315
         Width           =   180
      End
      Begin VB.TextBox Text6 
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
         Left            =   6615
         TabIndex        =   10
         Text            =   " "
         Top             =   315
         Width           =   435
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
         Height          =   300
         Left            =   2310
         TabIndex        =   8
         Text            =   " "
         Top             =   1155
         Width           =   2850
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
         Height          =   300
         Left            =   1470
         TabIndex        =   7
         Text            =   " "
         Top             =   1155
         Width           =   750
      End
      Begin VB.TextBox Text3 
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
         Left            =   1470
         TabIndex        =   6
         Text            =   " "
         Top             =   735
         Width           =   3690
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
         Height          =   300
         Left            =   1470
         TabIndex        =   5
         Text            =   " "
         Top             =   315
         Width           =   3690
      End
      Begin VB.CommandButton Command1 
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
         Left            =   1050
         TabIndex        =   4
         Top             =   1155
         Width           =   180
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
         Left            =   210
         TabIndex        =   3
         Text            =   " "
         Top             =   1155
         Width           =   855
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00E0E0E0&
         Height          =   4530
         Left            =   210
         ScaleHeight     =   4470
         ScaleWidth      =   10455
         TabIndex        =   49
         Top             =   1680
         Width           =   10515
         Begin VB.Frame Frame4 
            BackColor       =   &H0080C0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   540
            Left            =   0
            TabIndex        =   51
            Top             =   0
            Width           =   10515
            Begin VB.Line Line6 
               X1              =   8820
               X2              =   8820
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line5 
               X1              =   7560
               X2              =   7560
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line4 
               X1              =   2310
               X2              =   2310
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line1 
               X1              =   840
               X2              =   840
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   $"FORMFAC.frx":1AD2
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
               Left            =   0
               TabIndex        =   52
               Top             =   140
               Width           =   10410
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0F0FF&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5580
            Left            =   0
            TabIndex        =   50
            Top             =   -1050
            Width           =   10515
            Begin VB.Line Line2 
               Index           =   3
               X1              =   8820
               X2              =   8820
               Y1              =   525
               Y2              =   5565
            End
            Begin VB.Line Line2 
               Index           =   2
               X1              =   7560
               X2              =   7560
               Y1              =   525
               Y2              =   5565
            End
            Begin VB.Line Line2 
               Index           =   1
               X1              =   2310
               X2              =   2310
               Y1              =   525
               Y2              =   5565
            End
            Begin VB.Line Line2 
               Index           =   0
               X1              =   840
               X2              =   840
               Y1              =   525
               Y2              =   5565
            End
         End
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   420
         Picture         =   "FORMFAC.frx":1B63
         ToolTipText     =   "Click Aquí para cambiar Modalidad de Facturación"
         Top             =   210
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Doc.Numero:"
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
         Left            =   7455
         TabIndex        =   22
         Top             =   1155
         Width           =   1590
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rto./ Resum.:"
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
         Left            =   7245
         TabIndex        =   20
         Top             =   735
         Width           =   1800
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
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
         Left            =   8610
         TabIndex        =   17
         Top             =   315
         Width           =   750
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pos.IVA:"
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
         Left            =   5670
         TabIndex        =   9
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
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
         Left            =   315
         TabIndex        =   1
         Top             =   840
         Width           =   960
      End
   End
   Begin VB.Label Label10 
      BackColor       =   &H00800000&
      Caption         =   "  Facturacion - Debitos y Creditos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   330
      Left            =   0
      TabIndex        =   59
      ToolTipText     =   "Doble Click Aquí para Cambiar Modalidad de Facturación "
      Top             =   0
      Width           =   11895
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pos.IVA:"
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
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "FORMFAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTNEXT_Click()
     '
     If InStr(FORMFAC.Label10.Caption, "Anulacion de Comprobantes") > 0 Then
         Call ANUFAC
         Exit Sub
     End If
     '
     NCII% = Val(Text1.TEXT)
     '
     If NC% < 1 Or NCII% <> NC% Then
        Text1.SetFocus
        Exit Sub
     End If
     '
     If TRIM$(Text2.TEXT) <> RASOCLI$(NC%) Then
        Text1.SetFocus
        Exit Sub
     End If
     '
     FFI1 = FECHANUM(TRIM$(FORMFAC.Text9.TEXT))
     FECHALIM% = Val(CONTROL$("", "CIEMEVEN"))
     If FFI1 < 1 Or FFI1 <= FECHALIM% Then
        Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
        FORMFAC.Text9.SetFocus
        Exit Sub
     End If
     '
     Call SIGUEFAC1
     '
End Sub

Private Sub Command1_Click()
    '
    Call SELECHO("DEUDOR1")
    NC% = Val(TRIM$(VALACT1$("DEUDOR1")))
    Text1.TEXT = TRIM$(Str$(NC%))
    Call CARDACLI
    '
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("AMACLI/NEW", "Alta de Nueva Cuenta de Cliente")
    Call FRESHECHO("DEUDOR1")
    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
    If COMALTER%("Fin de Facturación\\Reanudar\Salir") <> 2 Then
       Exit Sub
    End If
    Call BLANPEDP
    Call FINAL("")
End Sub

Private Sub Command2_Click()
    If COMALTER%("Fin de Facturación\\Reanudar\Salir") <> 2 Then
       Exit Sub
    End If
    Call BLANPEDP
    Call FINAL("")
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NCLI$ = TRIM$(Text1.TEXT)
    Call CONECRUN("*NOTACLI/" + NCLI$, "Notas de Clientes")
    ' Shell LUBAS$ + "WINTABU NOTACLI/" + NCLI$, vbNormalFocus
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RESCUDEU.Text1.TEXT = Text1.TEXT
    Hide
    RESCUDEU.Show
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call CONECRUN("ADMIPRE", "Listas de Precios de Venta")
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Call SELECHO("CAPOSIV")
    If TRIM$(VALACT1$("CAPOSIV")) <> "" Then
        Text6.TEXT = TRIM$(VALACT1$("CAPOSIV"))
      Else
        Text6.TEXT = "1"
    End If
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    Call CONECRUN("#SVD151", "Condiciones de Pago Clientes")
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Screen.MousePointer = 11
    V& = 0
    For U& = 1 To ULTREG&("REMPEN")
        X$ = REGLEIDO$("REMPEN", U&)
        If CVI(Left$(X$, 2)) > 0 Then
            If CVS(Mid$(X$, 27, 4)) >= 0.05 Then
                V& = V& + 1
                Call GRAREG("REMPEN", X$, V&)
            End If
        End If
    Next U&
    Call SETULTREG("REMPEN", V&)
    Call CIERRARCH("REMPEN")
    Screen.MousePointer = 1
    '
    MDX% = ALTERNA%("Archivo Completo\Por Cliente\Por Número de Remito")
    If MDX% = 1 Then
        Call CONECRUN("*COREPEN", "Remitos Pendientes de Facturación")
      ElseIf MDX% = 2 Then
        Screen.MousePointer = 11
        Call COPYSTRU("DEUDOR1", "DEUDOREM")
        Dim NCAC%(16384)
        For U& = 1 To ULTREG&("REMPEN")
            X$ = REGLEIDO$("REMPEN", U&)
            NCX% = CVI(Left$(X$, 2))
            If NCX% > 0 Then
                If NCX% <= 16384 Then
                    NCAC%(NCX%) = 1
                End If
            End If
        Next U&
        '
        V& = 0
        For U& = 1 To ULTREG&("DEUDOR1")
            X$ = REGLEIDO$("DEUDOR1", U&)
            NCX% = CVI(Left$(X$, 2))
            If NCX% > 0 Then
                If NCX% <= 16384 Then
                    If NCAC%(NCX%) > 0 Then
                        V& = V& + 1
                        Call GRAREG("DEUDOREM", X$, V&)
                    End If
                End If
            End If
        Next U&
        '
        Call SETULTREG&("DEUDOREM", V&)
        Call CIERRARCH("DEUDOREM")
        Call FRESHECHO("DEUDOREM")
        Screen.MousePointer = 1
        '
        Call CONECRUN("*REPECLI", "Remitos Pendientes por Cliente")
      ElseIf MDX% = 3 Then
        Call CONECRUN("*REMPEN", "Remitos Pendientes por Numero")
    End If
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Call SELECHO("SELFECHA")
    Text9.TEXT = VALACT1$("SELFECHA")
    FF% = FECHANUM(Text9.TEXT)
End Sub

Private Sub Form_Load()
    '
    'EPAC$ = TRIM$(EMPREAC$)
    'If EPAC$ <> "" Then
    '    Caption = Caption + "  -  " + EPAC$
    'End If
    '
    Call VERJOBID
    Show
    ''Call FRESHALL
    '
    J& = 0
    TDF% = 1
10  FIN& = ULTREG&("TIDOCFAC")
    For U& = 1 To FIN&
        X$ = REGLEIDO$("TIDOCFAC", U&)
        If Asc(Left$(X$, 1)) = TDF% Then
            J& = J& + 1
            Y$ = Space$(128)
            X1$ = REPLACAR$(TRIM$(Mid$(X$, 2, 23)) + " - " + TRIM$(Mid$(X$, 33, 32)), Chr$(160), "a")
            X2$ = REPLACAR$(X1$, Chr$(130), "e")
            X3$ = REPLACAR$(X2$, Chr$(162), "o")
            Call REPLA(Y$, X3$, 1, 56)
            Call REPLA(Y$, Mid$(X$, 65, 64), 63, 64)
            Call REPLA(Y$, TRIM$(Str$(TDF%)), 127, 1)
            Call REPLA(Y$, TRIM$(Str$(Asc(Mid$(X$, 129, 1)))), 128, 1)
            Call GRAREG("SELTIFAC", Y$, J&)
            If J& = 1 Then
                TDFAC$ = Y$
            End If
        End If
    Next U&
    If TDF% < 3 Then
        TDF% = TDF% + 1
        GoTo 10
    End If
    J& = J& + 1
    Y$ = REGBLAN$("SELTIFAC")
    Call REPLA(Y$, "Anulacion de Comprobantes", 1, 56)
    Call GRAREG("SELTIFAC", Y$, J&)
    '
    Call SETULTREG("SELTIFAC", J&)
    Call CIERRARCH("SELTIFAC")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
          ENCABE$ = "  " + EPAC$ + "  -  "
        Else
          ENCABE$ = "  "
    End If
    ENCABE$ = ENCABE$ + TRIM$(Left$(TDFAC$, 56))
    FORMFAC.Label10.Caption = ENCABE$
    '
    RGTIFA& = 1
    CONTROFA$ = REGLEIDO$("TIDOCFAC", RGTIFA&)
    Call CIERRARCH("TIDOCFAC")
    '
    Call VAPUBLI
    '
    If ULTREG&("REMPEN") < 1 Then
        Command8.Enabled = False
    End If
    '
End Sub
'

Private Sub Image1_Click()
    '
10  Call SELTIFAC
    If InStr(FORMFAC.Label10.Caption, "Anulacion de Comprobantes") = 0 Then
       Call VAPUBLI
     Else
       Call ANUFAC
       Call SELTIFAC
       GoTo 10
    End If
    '
End Sub

Private Sub Label10_DblClick()
10  Call SELTIFAC
    If InStr(FORMFAC.Label10.Caption, "Anulacion de Comprobantes") = 0 Then
       Call VAPUBLI
     Else
       Call ANUFAC
       Call SELTIFAC
       GoTo 10
    End If
End Sub

Private Sub PORDESCU_Click()
   If DERACCE%("CAMBDESC", "Modificar Descuento de Cliente") < 2 Then
      Text1.SetFocus
   End If
End Sub

Private Sub PORDESCU_LostFocus()
    If TRIM$(PORDESCU.TEXT) <> TRIM$(DESCUCLI$(NC%)) Then
        If ALTERNA%("Registra este Valor como Standard") = 1 Then
            DESCUNUE0$ = TRIM$(PORDESCU.TEXT)
            DESCUNUE$ = REPLACAR$(DESCUNUE0$, "%", " ")
            For U& = 1 To ULTREG("DESCUCLI")
                X1$ = REGLEIDO$("DESCUCLI", U&)
                If CVI(Left$(X1$, 2)) = NC% Then
                    Call REPLA(X1$, DESCUNUE$, 3, 14)
                    Call GRAREG("DESCUCLI", X1$, U&)
                    Call CIERRARCH("DESCUCLI")
                    Exit Sub
                End If
            Next U&
            X1$ = MKI$(NC%) + DESCUNUE$
            Call REGAPP("DESCUCLI", X1$)
            Call CIERRARCH("DESCUCLI")
        End If
    End If
End Sub
'

Private Sub PORIVA_LOSTFOCUS()
    If Val(PORIVA.TEXT) < 0.5 Then
        PORIVA.TEXT = ""
        ALICU = 0
        ALI1 = 0
        ALI2 = 0
      Else
        PORIVA.TEXT = ALICU
        PORIVA.Refresh
    End If
End Sub

Private Sub Text1_DblClick()
    '
    Call SELECHO("DEUDOR1")
    NC% = Val(TRIM$(VALACT1$("DEUDOR1")))
    Text1.TEXT = TRIM$(Str$(NC%))
    Call CARDACLI
    '
End Sub

Private Sub Text1_GotFocus()
    If ALICU < 0.5 Then
        Call FORMPRIN(LU$, FLIM%, ALI1, ALI2, TC1, TLET$, INTER%, TFOR%, ANFO%, LARFO%, DIFER, CAFOR%, MARIZ%, MARINT%, REDON%, MPRI$, MCEN$, FACTO, CANCON%, COMA%, PARDEC%)
        TICA = TC1
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Val(Text1.TEXT) > 0 Then
          If Val(Text1.TEXT) <= 32767 Then
             NC% = Val(Text1.TEXT)
             If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
                Call CARDACLI
                BOTNEXT.SetFocus
                Exit Sub
             End If
          End If
       End If
       '
       Beep
       Text1.TEXT = ""
       '
    End If
    '
End Sub

Private Sub Text11_LostFocus()
    '
    LU$ = LUDAT$
    CLI$ = RASOCLI$(NC%)
    FLIM% = FECHANUM("01/01/99")
    If ULTREG&("FORMFAC") > 0 Then
        FFS$ = REGLEIDO$("FORMFAC", 2&)
        FLIM% = CVI(Mid$(FFS$, 19, 2))
    End If
    '
    If Val(Text11.TEXT) < 1 Then
        Call MENSERR(24, "Numero de Documento no Permitido")
        Call GRAFAC98(LU$, ARCHIFAC$, 1, RD%, FLIM%, TIPO%, VA%, CAFOR%, NRO, NC%, CLI$, IVAS$, FFI%, FVI%, LPRE%, VENDE%, TOTAFAC#())
        '
        NROPRO = NRO
        FORMFAC.Text11.TEXT = NROPRO
        Exit Sub
    End If
    '
    NRO = Val(Text11.TEXT)
    FLIM% = FECHANUM("01/01/99")
    If ULTREG&("FORMFAC") > 0 Then
        FFS$ = REGLEIDO$("FORMFAC", 2&)
        FLIM% = CVI(Mid$(FFS$, 19, 2))
    End If
    '
    Call GRAFAC98(LU$, ARCHIFAC$, 2, RD%, FLIM%, TIPO%, VA%, CAFOR%, NRO, NC%, CLI$, IVAS$, FFI%, FVI%, LPRE%, VENDE%, TOTAFAC#())
    '
    If NRO <> Val(Text11.TEXT) Then
        Call MENSERR(24, "Numero de Documento Pre-existente")
        Text11.TEXT = NRO
        Exit Sub
    End If
    '
End Sub

Private Sub Text6_Change()
    PIVA% = Val(Text6.TEXT)
    If ECOARCH$("CAPOSIV", Chr$(PIVA%)) <> "" Then
        XX$ = REGACT$("CAPOSIV")
        Text7.TEXT = UCase$(Mid$(XX$, 33, 8))
        FACIVA = 1
        CATIVA$ = "A"
        VA% = 1
        If PIVA% = 0 Or PIVA% = 3 Or PIVA% = 4 Or PIVA% = 6 Then
           FACIVA = 1 + ALICU / 100
           CATIVAS$ = "B"
           VA% = 2
           '
          ElseIf PIVA% = 2 Then
           ALICU = ALI1 + ALI2
           RNI% = 1
           '
          ElseIf PIVA% = 5 Then
           ALICU = 0
           CATIVAS$ = " "
           VA% = 0
           If LETREXPO$ = "A" Then
              VA% = 1
              CATIVAS$ = "A"
           End If
        End If
        Call VAPUBLI
      Else
        Text6.TEXT = "1"
    End If
End Sub

Private Sub Text6_DblClick()
    Call SELECHO("CAPOSIV")
    If TRIM$(VALACT1$("CAPOSIV")) <> "" Then
        Text6.TEXT = TRIM$(VALACT1$("CAPOSIV"))
      Else
        Text6.TEXT = "1"
    End If
End Sub

Private Sub TEXT21_Change(Index As Integer)
    '
    If Index = 0 Then
        Eco21(Index).TEXT = ECOARCH$("LISTAS", MKI$(Val(TEXT21(Index).TEXT)))
      ElseIf Index = 1 Then
        Eco21(Index).TEXT = ECOARCH$("CONPAG", MKI$(Val(TEXT21(Index).TEXT)))
      ElseIf Index = 2 Then
        Eco21(Index).TEXT = ECOARCH$("VENDEDOR", Chr$(Val(TEXT21(Index).TEXT)))
    End If
    '
End Sub

Private Sub Command21_Click(Index As Integer)
    '
    If Index = 0 Then
        ARCHEC$ = "LISTAS": LOCA% = 2
      ElseIf Index = 1 Then
        ARCHEC$ = "CONPAG": LOCA% = 2
      ElseIf Index = 2 Then
        ARCHEC$ = "VENDEDOR": LOCA% = 1
    End If
    ARGU$ = CBIN$(TRIM$(TEXT21(Index).TEXT), 1, LOCA%)
    Call SELECHO(ARCHEC$)
    TEXT21(Index).TEXT = VALACT1$(ARCHEC$)
    '
End Sub



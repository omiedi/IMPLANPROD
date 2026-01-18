VERSION 5.00
Begin VB.Form FORMUFAC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturacion"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   11910
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2115
      Left            =   0
      TabIndex        =   23
      Top             =   6510
      Width           =   10800
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
      BackColor       =   &H00C0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8400
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12000
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   10920
         ScaleHeight     =   75
         ScaleWidth      =   810
         TabIndex        =   60
         Top             =   6720
         Width           =   875
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   61
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command12 
         Caption         =   "   O.K.   Print "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   10920
         Picture         =   "FORMUFAC.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Condiciones de Pago"
         Top             =   7035
         Width           =   855
      End
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
         Height          =   800
         Left            =   10920
         Picture         =   "FORMUFAC.frx":066A
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
         Height          =   800
         Left            =   10920
         Picture         =   "FORMUFAC.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "Ordenes de Despacho Pendientes por Cliente"
         Top             =   4200
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
         Height          =   700
         Left            =   10920
         Picture         =   "FORMUFAC.frx":0EEE
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Condiciones de Pago"
         Top             =   5800
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
         Height          =   700
         Left            =   10920
         Picture         =   "FORMUFAC.frx":11F8
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Listas de Precios de Venta"
         Top             =   5040
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
         Height          =   800
         Left            =   10920
         Picture         =   "FORMUFAC.frx":1342
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Resumen de Cuenta Corriente"
         Top             =   2520
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
         Height          =   800
         Left            =   10920
         Picture         =   "FORMUFAC.frx":164C
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Notas de Clientes"
         Top             =   3360
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
         Picture         =   "FORMUFAC.frx":1956
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
         Picture         =   "FORMUFAC.frx":1C60
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
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         Height          =   4820
         Left            =   210
         ScaleHeight     =   4755
         ScaleWidth      =   10455
         TabIndex        =   49
         Top             =   1680
         Width           =   10515
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   540
            Left            =   0
            TabIndex        =   51
            Top             =   0
            Width           =   10515
            Begin VB.Line Line7 
               X1              =   0
               X2              =   10500
               Y1              =   525
               Y2              =   525
            End
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
               Caption         =   $"FORMUFAC.frx":1F6A
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
               Top             =   135
               Width           =   10410
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   5900
            Left            =   0
            TabIndex        =   50
            Top             =   -1050
            Width           =   10515
            Begin VB.Line Line2 
               Index           =   3
               X1              =   8820
               X2              =   8820
               Y1              =   525
               Y2              =   6490
            End
            Begin VB.Line Line2 
               Index           =   2
               X1              =   7560
               X2              =   7560
               Y1              =   525
               Y2              =   6000
            End
            Begin VB.Line Line2 
               Index           =   1
               X1              =   2310
               X2              =   2310
               Y1              =   525
               Y2              =   6200
            End
            Begin VB.Line Line2 
               Index           =   0
               X1              =   840
               X2              =   840
               Y1              =   525
               Y2              =   6200
            End
         End
      End
      Begin VB.Line Line9 
         X1              =   210
         X2              =   10720
         Y1              =   6490
         Y2              =   6490
      End
      Begin VB.Line Line8 
         X1              =   10720
         X2              =   10720
         Y1              =   1680
         Y2              =   6490
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   420
         Picture         =   "FORMUFAC.frx":1FFB
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
Attribute VB_Name = "FORMUFAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NUPEDSEL&(1024)
Dim RF$(4)
Dim DATCLI$(32)
'
Dim POI%(10)
Dim T$(2)
Dim RG&(16)
'
Dim TPX%(8), TIPCOMS$(8), VTIPS$(8)
Dim TIFORM%
Dim VARIFA$
Dim FORVENTA$
Dim FORIMPRE$
Dim VENCADET$
Dim VEPIEDET$
Dim MAXITEMS%
Dim UREN%, RG0%
Dim RTO$, TRXP$
Dim CAPED%
'
Dim TIPCOM$
Dim TICA
Dim FFI%, FVI%
Dim PORDES, MODESCU%, VERISTO%
Dim FACIVA
Dim CATIVAS$
'
Dim SELORDESPA%
'
Dim EXPOR%, FINFACEX$
Dim SOLOFAC%
Dim YACERRA%

Private Sub BOTNEXT_Click()
     '
     If InStr(Caption, "Anulacion de Comprobantes") > 0 Then
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
     FFI1 = FECHANUM(TRIM$(Text9.TEXT))
     FECHALIM% = Val(CONTROL$("", "CIEMEVEN"))
     If FFI1 < 1 Or FFI1 <= FECHALIM% Then
        Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
        Text9.SetFocus
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
    If EXISTE%("AMACLI02.EXE") > 0 Then
        Call CONECRUN("AMACLI02", "Alta de Nueva Cuenta de Cliente")
      Else
        Call CONECRUN("AMACLI/NEW", "Alta de Nueva Cuenta de Cliente")
    End If
    Call FRESHECHO("DEUDOR1")
    Command11.Enabled = True
End Sub


Private Sub Command2_Click()
    If COMALTER%("Fin de Facturación\\Reanudar\Salir") <> 2 Then
       Exit Sub
    End If
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    YACERRA% = 1
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
    PROLLAMA$ = Name
    RESCUDEU.Show
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    If EXISTE%("LISPRE02.EXE") > 0 Then
         Call CONECRUN("LISPRE02", "Listas de Precios de Venta")
       Else
         Call CONECRUN("ADMIPRE", "Listas de Precios de Venta")
    End If
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
    Call CARCONPAGOS
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CONODESCLI
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Call SELECHO("SELFECHA")
    Text9.TEXT = VALACT1$("SELFECHA")
    FF% = FECHANUM(Text9.TEXT)
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + " - " + EPAC$
    End If
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
          ENCABE$ = "Facturación - " + EPAC$ + " - "
        Else
          ENCABE$ = "Facturación - "
    End If
    ENCABE$ = ENCABE$ + TRIM$(Left$(TDFAC$, 56))
    Caption = ENCABE$
    '
    RGTIFA& = 1
    CONTROFA$ = REGLEIDO$("TIDOCFAC", RGTIFA&)
    Call CIERRARCH("TIDOCFAC")
    '
    Call VERIMOVI
    '
    Call INIFACTU
    Call VAPUBLI
    '
    Printer.TrackDefault = False
    PORTUFAC$ = CONTROL$("*", "PORTFAC")
    PORTUREM$ = CONTROL$("*", "PORTREM")
    DIREPRIN.Show 1
    If DIREPRIN.Label3 <> "OK" Then
      Call FINAL("")
    End If
    '
    PORTUFAC$ = DIREPRIN.Combo1.TEXT
    PORTUREM$ = DIREPRIN.Combo2.TEXT
    '
    If PORTUFAC$ <> "" Then
      Call SETPRINPORT(PORTUFAC$)
      Call SETPRINPORT("RESTORE")
    End If
    
    If PORTUREM$ <> "" Then
      Call SETPRINPORT(PORTUREM$)
      Call SETPRINPORT("RESTORE")
    End If
    
    Call GRACONTROL("*", "PORTFAC", PORTUFAC$)
    Call GRACONTROL("*", "PORTREM", PORTUREM$)
    PRINTERPORT$ = PORTUFAC$
    '
End Sub
'

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call GRATITFOR(Caption)
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   If YACERRA% < 1 Then
     If COMALTER%("Fin de Facturación\\Reanudar\Salir") <> 2 Then
       Cancel = 1
       Exit Sub
     End If
   End If
   '
   Cancel = 0
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Image1_Click()
    '
10  Call SELTIFAC
    If InStr(Caption, "Anulacion de Comprobantes") = 0 Then
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
          ElseIf PIVA% = 5 Or PIVA% = 7 Then
           ALICU = 0
           CATIVAS$ = "X"
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

Sub VAPUBLI()
    '
    If CONTROFA$ = "" Then
       Call MENSERR(24, "No se ha Establecido\el Controlador de Facturación.\  \Consulte a su Soporte de Sistemas.")
       Call FINAL("")
    End If
    '
    TIDOC% = Asc(Left$(CONTROFA$, 1))
    VTIP2$ = Space$(9)
    '
    On TIDOC% GoTo 210, 220, 230
    Call FINAL("")
    '
210 TIPO% = 2: VTIP2$ = "Remito(s)"
    TIPCOM$ = "Factura"
    MONEDA% = 1: GoTo 300
    '
220 TIPO% = 3
    TIPCOM$ = "N.Debito"
    MONEDA% = 1: GoTo 300
    '
230 TIPO% = 6
    TIPCOM$ = "Credito"
    MONEDA% = 1: GoTo 300
    '
300 Text11.TEXT = ""
    If NC% > 0 Then
       Call GRAFAC98(LU$, ARCHIFAC$, 1, RD%, FLIM%, TIPO%, VA%, CAFOR%, NRO, NC%, CLI$, IVAS$, FFI%, FVI%, LPRE%, VENDE%, TOTAFAC#())
       NROPRO = NRO
       Text11.TEXT = NROPRO
    End If
    '
End Sub
'
Sub INIFACTU()
    '
    Static CTX%
    '
    FORMAIMPRE$ = "WIN"
    EPAC$ = EMPREAC$
    If InStr(UCase$(EPAC$), "SABO") > 0 Then FORMAIMPRE$ = "DOS"
    If InStr(UCase$(EPAC$), "NOVOESPU") > 0 Then FORMAIMPRE$ = "DOS"
    '
    UTI$ = LUBAS$
    If Mid$(UTI$, 2, 1) <> ":" Then
        UTI$ = Left$(CurDir, 2) + UTI$
    End If
'    If EXISTE%(UTI$ + "PRINTFRM.EXE") < 1 Then
'        Call COMUNI("Imposible Procesar Facturación.\  \Falta Utilitario 'PRINTFRM.EXE'\en Directorio '" + REPLACAR$(UTI$, "\", "/") + "'.")
'        Call FINAL("")
'    End If
'    '
    ORCOP$(1) = "ORIGINAL"
    ORCOP$(2) = "DUPLICADO"
    ORCOP$(3) = "TRIPLICADO"
    ORCOP$(4) = "CUADRUPLICADO"
    ORCOP$(5) = "QUINTUPLICADO"
    ORCOP$(6) = "SEXTUPLICADO"
    
100 Call CIERRARCH("*.*")
    If CTX% = 0 Then
        IQ& = 0
        IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(0) + AJUSTI$("Consumidor Final            ", 31) + AJUSTI$("Cons.Fin", 9) + String$(23, 0), IQ&)
        IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(1) + AJUSTI$("IVA Responsable Inscripto   ", 31) + AJUSTI$("Resp.Insc", 9) + String$(23, 0), IQ&)
        IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(2) + AJUSTI$("IVA Responsable no Inscripto", 31) + AJUSTI$("R.No Insc", 9) + String$(23, 0), IQ&)
        IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(3) + AJUSTI$("IVA No Responsable          ", 31) + AJUSTI$("No Resp.", 9) + String$(23, 0), IQ&)
        IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(4) + AJUSTI$("IVA Exento                  ", 31) + AJUSTI$("Exento  ", 9) + String$(23, 0), IQ&)
        IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(5) + AJUSTI$("IVA Exportador Exento       ", 31) + AJUSTI$("Export. ", 9) + String$(23, 0), IQ&)
        IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(6) + AJUSTI$("Contribuyente Monotributo   ", 31) + AJUSTI$("Monotrib", 9) + String$(23, 0), IQ&)
        IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(7) + AJUSTI$("IVA Exento Zona Promovida   ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
        Call SETULTREG&("CAPOSIV", IQ&)
        Call CIERRARCH("CAPOSIV")
        CTX% = 1
    End If
    '
102 Call CAREMPRE(LU$, A$())
    Call FORMPRIN(LU$, FLIM%, ALI1, ALI2, TC1, TLET$, INTER%, TFOR%, ANFO%, LARFO%, DIFER, CAFOR%, MARIZ%, MARINT%, REDON%, MPRI$, MCEN$, FACTO, CANCON%, COMA%, PARDEC%)
    TICA = TC1
    '
109 VTIP2$ = "Remito(s)"
    ARCHIFAC$ = "FACTUR"
    ARCHIDEU$ = "SVD202"         ' replantear estas variantes segun facstan
    ARCHICLI$ = "DEUDOR"
    ARCHIMAT$ = "MATVEN"
    ARCHISTO$ = "SALSTOCK"
    '
' ***************************************************************************

200 Screen.MousePointer = 1
    '
    Call VAPUBLI
    '
' ***************************************************************************

400 FFI% = HOY(HO$)
    Text9.TEXT = HO$
    VARIFA$ = TRIM$(TIPCOM$) + " - " + TRIM$(Mid$(CONTROFA$, 33, 32))
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
    '
End Sub
'
Sub CARDACLI()
    '
    Screen.MousePointer = 11
    '
    NC% = Val(Text1.TEXT)
    If NC% > 0 Then
        Text2.TEXT = RASOCLI$(NC%)
        Text3.TEXT = DOMICLI$(NC%)
        Text4.TEXT = CPOSCLI$(NC%)
        Text5.TEXT = LOCACLI$(NC%)
        Text6.TEXT = PIVACLI%(NC%)
        Text7.TEXT = POSIVAC$(NC%)
        Text8.TEXT = CUITCLI$(NC%)
        '
        TEXT21(0).TEXT = Str$(LIPRCLI%(NC%))
        TEXT21(1).TEXT = Str$(CPAGCLI%(NC%))
        TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
        '
        If TIDOC% > 1 Then
          TEXT21(1).TEXT = " 1"  ' fuerza C.P.=Contado
        End If
        '
        If Val(TEXT21(2).TEXT) < 1 Then
           Eco21(2).TEXT = "GERENCIA"
        End If
        '
        PIVA% = Val(Text6.TEXT)
        If PIVA% < 0 Or PIVA% > 7 Then
            Text6.SetFocus
            Screen.MousePointer = 1
            Exit Sub
        End If
        '
        PORDESCU.TEXT = DESCUCLI$(NC%)
        '
        GoSub 81                        ' determina posicion IVA y CATIVAS$
        PORIVA.TEXT = ALICU
        '
1050    LU$ = LUDAT$
        CLI$ = RASOCLI$(NC%)
        FLIM% = FECHANUM("01/01/99")
        If ULTREG&("FORMFAC") > 0 Then
            FFS$ = REGLEIDO$("FORMFAC", 2&)
            FLIM% = CVI(Mid$(FFS$, 19, 2))
        End If
        '
        Call GRAFAC98(LU$, ARCHIFAC$, 1, RD%, FLIM%, TIPO%, VA%, CAFOR%, NRO, NC%, CLI$, IVAS$, FFI%, FVI%, LPRE%, VENDE%, TOTAFAC#())
        '
        NROPRO = NRO
        Text11.TEXT = NROPRO
        '
    End If
90  Screen.MousePointer = 1
    Exit Sub
    '
    '
81  FACIVA = 1
    CATIVAS$ = "A"
    VA% = 1
    ALICU = ALI1
    RNI% = 0
    '
    If PIVA% = 0 Or PIVA% = 3 Or PIVA% = 4 Or PIVA% = 6 Then
          FACIVA = 1 + ALICU / 100
          CATIVAS$ = "B"
          VA% = 2
          '
        ElseIf PIVA% = 2 Then
          ALICU = ALI1 + ALI2
          RNI% = 1
          '
        ElseIf PIVA% = 5 Or PIVA% = 7 Then
          ALICU = 0
          CATIVAS$ = " "
          VA% = 0
          If LETREXPO$ = "A" Then
              VA% = 1
              CATIVAS$ = "A"
          End If
    End If
    '
Return
'
End Sub
'
Sub SIGUEFAC1()
    '
    LU$ = LUDAT$
    NCC$ = TRIM$(Text1.TEXT)
    If Val(NCC$) < 1 Then
        Text1.SetFocus
        Exit Sub
    End If
    '
    HOI1 = HOY(HO$)
    FFI1 = FECHANUM(TRIM$(Text9.TEXT))
    If FFI1 < HOI1 - 30 Or DIENTRE%(FFI1, HOI1) > 5 Then
        Call MENSERR(24, "Fecha no Válida")
        Text9.TEXT = FECHATEX$(HOI1)
        Exit Sub
    End If
    If DIENTRE%(HOI1, FFI1) > 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text9.TEXT = FECHATEX$(HOI1)
        Exit Sub
    End If
    FFI% = FFI1
    '
    NC% = Val(NCC$)
502 PIVA% = Val(Text6.TEXT)
    If ECOARCH$("CAPOSIV", Chr$(PIVA%)) = "" Then
        Text1.TEXT = "1"
        Text6.SetFocus
        Exit Sub
    End If
    '
    If PIVA% <> 0 Then
       If PIVA% <> 5 Then
          CUITT$ = TRIM$(Text8.TEXT)
          If CUITT$ = "" Or VALICUIT%(CUITT$) <> 1 Then
             Call MENSERR(24, "Imposible Facturar a esta Cuenta.\  \Número de CUIT no es Válido.")
             Text1.SetFocus
             Exit Sub
          End If
       End If
    End If
    '
    If TRIM$(DOMICLI$(NC%)) = "" Then
       Call MENSERR(24, "Imposible Facturar a esta Cuenta.\  \Falta Domicilio Legal.")
       Text1.SetFocus
       Exit Sub
    End If
    '
    LPRE% = Val(TEXT21(0).TEXT)
    If ECOARCH("LISTAS", MKI$(LPRE%)) = "" Then
        TEXT21(0).SetFocus
        Exit Sub
    End If
    '
    CPAG% = Val(TEXT21(1).TEXT)
    If ECOARCH("CONPAG", MKI$(CPAG%)) = "" Then
        TEXT21(1).SetFocus
        Exit Sub
    End If
    '
    VENDE% = Val(TEXT21(2).TEXT)
    If VENDE% > 0 Then
        If ECOARCH("VENDEDOR", Chr$(VENDE%)) = "" Then
            TEXT21(2).SetFocus
            Exit Sub
        End If
    End If
    '
    PROV$ = NPROCLI$(NC%)          ' numero de proveedor
    '
    Call CLOSEFORMPRIN("")
    FORVENTA$ = TRIM$(Mid$(CONTROFA$, 81, 8))
    FORIMPRE$ = TRIM$(Mid$(CONTROFA$, 89, 8))
    If PIVA% = 1 Or PIVA% = 2 Then
        FORVENTA$ = TRIM$(Mid$(CONTROFA$, 65, 8))
        FORIMPRE$ = TRIM$(Mid$(CONTROFA$, 73, 8))
      ElseIf PIVA% = 5 Or PIVA% = 7 Then
        FORVENTA$ = TRIM$(Mid$(CONTROFA$, 97, 8))
        FORIMPRE$ = TRIM$(Mid$(CONTROFA$, 105, 8))
    End If
    VENCADET$ = TRIM$(Mid$(CONTROFA$, 113, 8))
    VEPIEDET$ = TRIM$(Mid$(CONTROFA$, 121, 8))
    VERISTO% = Asc(Mid$(CONTROFA$, 129, 1))
    MODESCU% = Asc(Mid$(CONTROFA$, 130, 1))
    '
    If FORVENTA$ = "" Then
         Call COMUNI("Falta Definir Controlador\de Ventana de Carga de Factura")
         Call FINAL("")
      ElseIf EXISTE%(FORVENTA$ + ".PRF") <> 1 Then
        Call COMUNI("Falta Definir Controlador\de Ventana de Carga de Factura")
         Call FINAL("")
    End If
    '
    If FORIMPRE$ = "" Then
         Call COMUNI("Falta Definir Controlador\de Impresion de Documento")
         Call FINAL("")
       ElseIf EXISTE%(LU$ + FORIMPRE$ + ".FRM") <> 1 Then
         Call COMUNI("Falta Definir Controlador\de Impresion de Documento")
         Call FINAL("")
    End If
    '
    MAXITEMS% = Val(ATRIFORM$(FORIMPRE$, "MAX-ITEM"))
    CONTIHOJA% = 0: If TRIM$(ATRIFORM$(FORIMPRE$, "MAX-ITEM")) = "HOJACONT" Then CONTIHOJA% = 1
    MAXRENGLO% = 0
    If MAXITEMS% > 0 Then
       PRIREDET% = Val(ATRIFORMDOS$(FORIMPRE$, "DET-DES"))
       ULTREDET% = Val(ATRIFORMDOS$(FORIMPRE$, "DET-HAS"))
       MAXRENGLO% = ULTREDET% - PRIREDET%
    End If
    '
    Call CLOSEFORMPRIN("")
    '
    BOTNEXT.Enabled = False
    EXPOR% = 0: If PIVA% = 5 And TIPO% = 2 Then EXPOR% = 1

'*********************************************************** fecha de factura

1200 CAPED% = 0
     Call SETULTREG("!CARDEFAC", 0)
     Call CIERRARCH("!CARDEFAC")
     For KKUU1% = 0 To 3
       For KKUU2% = 0 To 10
         TOTAFAC#(KKUU1%, KKUU2%) = 0
       Next KKUU2%
     Next KKUU1%
     If TIDOC% = 1 Then
       If InStr(UCase$(FORVENTA$), "LIB") < 1 Then
         Call CARODESPEN
       End If
     End If

1205  TINI = Timer
1211  TOTAFAC#(1, 9) = 0
      '
      ARCHIPRE$ = TRIM$(Str$(LPRE%))
      While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
      ARCHIPRE$ = "LIPRE" + ARCHIPRE$
      '
1215  ATRI$ = "/COLORPAN=B-N/PS/TITUPAN=" + VARIFA$
      ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture1.Top + 50))
      ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture1.Left))
      '
      If InStr(UCase$(FORVENTA$), "LIB") < 1 Then ATRI$ = ATRI$ + "/ARCHECO(2)=" + ARCHIPRE$
      ' lo anterior para permitir facturacion libre
      'Picture1.Visible = False
      ECOPRECIO% = 0
      If TRIM$(UCase$(FORVENTA$)) = "FACRENO" Then ECOPRECIO% = (-1)
      If TRIM$(UCase$(FORVENTA$)) = "FACSEVE" Then ECOPRECIO% = (-1)
      If TRIM$(UCase$(FORVENTA$)) = "FACFORD" Then ECOPRECIO% = (-1)
      XI% = VENTABU%(FORVENTA$ + ATRI$)
        On Error Resume Next
        DoEvents
        On Error GoTo 0
      'Picture1.Visible = True
      If XI% < 0 Then
          GoTo 3500
          BOTNEXT.Enabled = True
          Text1.SetFocus
          Exit Sub
      End If
      '
      Call CIERRARCH("*.*")
      '
      If ULTREG&("!CARDEFAC") < 1 Then GoTo 3500
      '
      Screen.MousePointer = 11
      ATX% = 0: CIANT% = 0: FIN& = ULTREG&("!CARDEFAC")
      If InStr(UCase$(FORVENTA$), "LIB") > 0 Then GoTo 1310
      If TIDOC% <> 1 Then
        If TIDOC% <> 3 Or InStr(UCase$(VARIFA$), "DEVOL") < 1 Then
          GoTo 1310
        End If
      End If
      '
      CARENEC% = 0
      For U& = 1 To FIN&
         '
         Call TRACE(20, U&, FIN&)
         '
         X$ = REGLEIDO$("!CARDEFAC", U&)
         CI% = CVI(Mid$(X$, 5, 2))
         CAN = CVS(Left$(X$, 4))
         CARENEC% = CARENEC% + 1
         '
         If VERISTO% = 1 Then
           If CVS(Mid$(X$, 509, 4)) >= 0 Then ' suprime remitos
             If CI% > 0 Then
               STDIS = STOCKACT(CI%) - CAN
               If STDIS < 0 Then
                  If TIDOC% <> 3 Then
                     Call COMUNI("Factura no Procesable.\Stock Negativo en C¢digo " + TRIM$(CODEXT$(CI%)))
                     GoTo 1215
                  End If
               End If
             End If
           End If
         End If
         '
         PRECIOD# = CDbl(CVS(Mid$(X$, 421, 4)))
         '
         If PRECIOD# < 0.00005 Then
           If CAN > 0.05 Then
             If CI% > 0 Then
               If CI% <= NUMAS% Then
                 PRECIOD# = PRELISTA#(ARCHIPRE$, CI%)
                 PRECIOD# = (Int(10000 * PRECIOD# + 0.5)) / 10000
                 Call REPLA(X$, MKS$(PRECIOD#), 421, 4)
                 DE$ = DESCRIT$(CI%)
                 Call REPLA(X$, DE$, 7, 64)
                 Call GRAREG("!CARDEFAC", X$, U&)
               End If
             End If
           End If
         End If
         '
         NORI$ = NUORIG$(ARCHIPRE$, CI%)
         Call REPLA(X$, NORI$, 339, 16)
         '
         OCOM$ = NUOCOM$(NC%, CI%)
         Call REPLA(X$, OCOM$, 355, 16)
         '
         PORDESCUNI = CVS(Mid$(X$, 425, 4))
         If PORDESCUNI < 0 Or PORDESCUNI > 99 Then
           Call MENSERR(24, "Descuento no Procesable")
           GoTo 1215
         End If
         PREUNI# = (Int(PRECIOD# * (10000 - 100 * PORDESCUNI) + 0.5)) / 10000
         Call REPLA(X$, MKS$(PREUNI#), 429, 4)
         '
         TOT# = (Int(100 * CDbl(CAN * PREUNI#) + 0.5)) / 100
         If TOT# > 0.005 Then Call REPLA(X$, MKD$(TOT#), 433, 8)
         '
         TXT$ = ""
         CX% = 0: RF$(1) = "": RF$(2) = "": RF$(3) = ""
         If NORI$ <> "" Then TXT$ = "Nro.Orig.: " + NORI$
         If OCOM$ <> "" Then TXT$ = TXT$ + "    O.Compra: " + OCOM$
         NORI$ = "": OCOM$ = ""
         TXT$ = TRIM$(TXT$)
         If TXT$ <> "" Then
             CX% = CX% + 1
             CARENEC% = CARENEC% + 1
             RF$(CX%) = TXT$
         End If
         '
         If CI% > 0 Then
             TXT$ = TRIM$(DESPACHO$(CI%))
             If TXT$ <> "" Then
                 CX% = CX% + 1
                 CARENEC% = CARENEC% + 1
                 RF$(CX%) = "Desp: " + TXT$
             End If
         End If
         '
         For CY% = 1 To CX%
            Call REPLA(X$, RF$(CY%), 11 + 64 * CY%, 64)
         Next CY%
         '
         Call GRAREG("!CARDEFAC", X$, U&)
         '
         If CVS(Mid$(X$, 429, 4)) < 0.00005 Then
            ATX% = 1
         End If
         '
      Next U&
      Screen.MousePointer = 1
      '
      X$ = "": DE$ = "": TXT$ = "": RF$(1) = "": RF$(2) = "": RF$(3) = ""
      If MAXRENGLO% > 0 Then
        If CARENEC% > MAXRENGLO% Then
          Call CIERRARCH("*.*")
          Call COMUNI("Factura no Procesable\Demasiados Items.")
          GoTo 1215
        End If
      End If
      '
      If ATX% > 0 Then
         Call CIERRARCH("*.*")
         Call COMUNI("Factura no Procesable\hay Items sin Precio.")
         GoTo 1215
      End If
      '
1310  Call CIERRARCH("*.*")
      GoSub 8200                          ' calcular IVA de cada item
      '
      GoSub 8700                                        ' presentar totales
      TITU$ = "FORMATO NORMAL"
1330  '
      FACIVA = 1
      CATIVAS$ = "A"
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
          ElseIf PIVA% = 5 Or PIVA% = 7 Then
           ALICU = 0
           CATIVAS$ = " "
           VA% = 0
           If LETREXPO$ = "A" Then
              VA% = 1
              CATIVAS$ = "A"
           End If
      End If
      FOPRES$ = "PRESEFAC": If FACIVA <> 1 Then FOPRES$ = "PRESFACB"
      ATRI$ = ""
      ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture1.Top + 50))
      ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture1.Left))
      'Picture1.Visible = False
      XI% = VENTABU%(FOPRES$ + ATRI$ + "/COLORPAN=B-N/NC/TITUPAN=" + TITU$)
      On Error Resume Next
      DoEvents
      On Error GoTo 0
      'Picture1.Visible = True
      If XI% < 0 Then
          GoTo 3500
          BOTNEXT.Enabled = True
          Text1.SetFocus
          Exit Sub
      End If
      '
      ATX% = 0
      FIN& = ULTREG&("!CARDEFAC")
      '
      For U& = 1 To FIN&
         X$ = REGLEIDO$("!CARDEFAC", U&)
         CAN = CVS(Left$(X$, 4))
         PREUNI# = CDbl(CVS(Mid$(X$, 429, 4)))
         If CAN > 0 Then
           If PREUNI# > 0 Then
             TOT# = (Int(100 * CDbl(CAN * PREUNI#) + 0.5)) / 100
             If Abs(TOT# - CVD(Mid$(X$, 433, 8))) >= 0.005 Then
               ATX% = 1
               Call REPLA(X$, MKD$(TOT#), 433, 8)
               Call GRAREG("!CARDEFAC", X$, U&)
             End If
           End If
         End If
      Next U&
      '
      If ATX% <> 0 Then GoTo 1310                ' volver a presentar factura
'
'****************************************************************************
'
        If TIDOC% = 1 Or (TIDOC% = 3 And InStr(UCase$(TDFAC$), "DEVOLUC") > 0) Then                   ' solo facturas
            If DESCUPRE% = 0 Then
                DESC$ = TRIM$(PORDESCU.TEXT)
                Call DESCUEN(PORDES, DESC$)
                '
2011            If PORDES > 0 Then
                    COEF = 1 - PORDES / 100
                    If MODESCU% = 1 Then COEF = 1
                    For U& = 1 To ULTREG&("!CARDEFAC")
                       X$ = REGLEIDO$("!CARDEFAC", U&)
                       CAN = CVS(Left$(X$, 4))
                       PREUNI# = (Int(10000 * (CDbl(CVS(Mid$(X$, 429, 4)))) * COEF + 0.5)) / 10000
                       Call REPLA(X$, MKS$(PREUNI#), 429, 4)
                       TOT# = (Int(100 * CDbl(CAN * PREUNI#) + 0.5)) / 100
                       If CVI(Mid$(X$, 5, 2)) < 1 Then
                           TOT# = (Int(100 * CVD(Mid$(X$, 433, 8)) * COEF + 0.5)) / 100
                       End If
                       '
                       IVAITEM# = TOT# * (ALICU / 100)
                       BRUITEM# = TOT# + IVAITEM#
                       '
                       Call REPLA(X$, MKD$(TOT#), 433, 8)
                       Call REPLA(X$, MKS$(ALICU), 441, 4)
                       Call REPLA(X$, MKD$(IVAITEM#), 445, 8)
                       Call REPLA(X$, MKD$(BRUITEM#), 453, 8)
                       Call GRAREG("!CARDEFAC", X$, U&)
                    Next U&
                End If
                '
            End If
        End If
        '
        GoSub 8700
        If TOTAFAC#(1, 5) < 1 Then
          If TIDOC% = 1 Or InStr(UCase$(TDFAC$), "LIB") < 1 Then
            Call MENSERR(24, "Imposible Registrar - Falta Importe.")
            GoTo 1200
          End If
        End If
        '
        EPAC$ = TRIM$(UCase$(EMPREAC$))
        If InStr(EPAC$, "SABO") > 0 Then
          If TIDOC% <= 1 Then
            If TOTAFAC#(1, 5) < 100 Then
              OPXX1% = COMALTER%("Importe Factura debe Superar $ 100.-\\Cancelar\Emitir Igual")
              If OPXX1% <> 2 Then GoTo 1200
            End If
          End If
        End If
        '
'****************************************************************************
'
2020 '  calculo percepcion ingresos brutos
        '
        TOTAFAC#(1, 10) = 0
        If TIDOC% > 1 Then
          TOTAFAC#(1, 10) = Val(InputBox$("Importe Percepción", "Percepciones Ingresos Brutos", "0.00"))
          GoSub 8700
          GoTo 2025
        End If
        '
        TASAA = Val(CONTROL$("PERIBRU", "TASA-A"))
        TASAB = Val(CONTROL$("PERIBRU", "TASA-B"))
        TASAEXP = Val(CONTROL$("PERIBRU", "TASA-EXP"))
        MINIPER = Val(CONTROL$("PERIBRU", "MINIPER"))
        TASAX = TASAA
        If VA% = 2 Then TASAX = TASAB
        If VA% = 0 Then TASAX = TASAEXP
        '
        REPERCI$ = FILTERGETRECORD$("DAPIBCLI", 1, MKI$(NC%))
        CATIB1% = Asc(Mid$(REPERCI$, 3, 1))
        PORCONVE = CVS(Mid$(REPERCI$, 5, 4))
          If PORCONVE > 100 Then PORCONVE = 100
          '
        If CATIB1% > 0 Then         ' NO ESTA EXENTO
          If PORCONVE >= 10.1 Then  ' PARTICIPACION CONVENIO MULTILATERAL
            BIMPON# = TOTAFAC#(1, 5): If VA% = 2 Then BIMPON# = BIMPON# + TOTAFAC#(1, 6)
            SUFANT# = BIMPON#: SUPERANT# = 0
            'For KUL& = ULTREG&("FACTUR") To 1 Step -1
            '  RFC$ = REGLEIDO$("FACTUR", KUL&)
            '  If CVI(Mid$(RFC$, 1, 2)) <> NC% Then GoTo 2021
            '  SIGNX% = 1: If Asc(Mid$(RFC$, 13, 1)) > 3 Then SIGNX% = (-1) 'excluye creditos
            '  SUFANT# = SUFANT# + SIGNX% * Abs(CVD(Mid$(RFC$, 15, 8)))
            '  If VA% = 2 Then SUFANT# = SUFANT# + SIGNX% * Abs(CVD(Mid$(RFC$, 32, 8)))
            '  SUPERANT# = SUPERANT# + SIGNX% * Abs(CVD(Mid$(RFC$, 57, 8)))
            'Next KUL&
            '
2021        If SUFANT# >= MINIPER Then
              PERTOTAX# = (SUFANT# * PORCONVE / 100) * (TASAX / 100)
            'If PERTOTAX# >= MINIPER Then
              PERSALTA# = PERTOTAX# - SUPERANT#
              TOTAFAC#(1, 10) = (Int(100 * PERSALTA# + 0.5)) / 100
              GoSub 8700
            End If
            '
          End If
        End If
        '
2025    If TIDOC% > 1 Then ' debitos y creditos
          If Abs(TOTAFAC#(1, 5)) < 0.005 Then
            If FACIVA = 1 Then                      ' solo facturas 'A'
              IVA# = TOTAFAC#(1, 6)
              If PIVA% = 5 Or PIVA% = 7 Then IVA# = 0
              If PIVA% <> 5 Then
                If PIVA% <> 7 Then
20251             VDEF$ = MKD$(IVA#)
                  OBX$ = OBJPLA$("ENTRAIVA", VDEF$)
                  If OBX$ = "" Then GoTo 3500
                  IVA# = CVD(OBX$)
                  If IVA# < 0.005 Then GoTo 20251
                End If
              End If
              '
              TOTAFAC#(1, 6) = IVA#
              TOTAFAC#(1, 7) = 0
              TOTAFAC#(1, 8) = TOTAFAC#(1, 5) + TOTAFAC#(1, 6) + TOTAFAC#(1, 7) + TOTAFAC#(1, 9) + TOTAFAC#(1, 10)
            End If
          End If
        End If
        '
2028    For I% = 1 To 10
            TOTAFAC#(2, I%) = TOTAFAC#(1, I%) / TICA
        Next I%
'
'****************************************************************************
'
2200    If TIDOC% > 1 Then GoTo 2900
        '
        TRXP$ = OBJPLA$("TRANSPOR", TRXP$)
        '
        ATRI$ = "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture1.Top + 50))
        ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture1.Left))
2300    If VENCADET$ <> "" Then
           VTBX% = VENTABU%(VENCADET$ + "/PS" + ATRI$)
           On Error Resume Next
           DoEvents
           On Error GoTo 0
        End If
        '
        If VEPIEDET$ <> "" Then
           VTBX% = VENTABU%(VEPIEDET$ + "/PS/NC" + ATRI$)
           On Error Resume Next
           DoEvents
           On Error GoTo 0
       End If
        '
2900  ' aprobacion de factura
        '
        XI% = COMALTER%("Fin de Carga de " + TIPCOM$ + " " + TRIM$(Str$(NRO)) + "\\Aprobar\Corregir\Anular")
        ''Call FRESHALL
        'Show
        
        '
        On XI% GoTo 2910, 1200, 3500
        Beep
        GoTo 2900
        '
2910  ' reserva numero de documento
        '
        GRAICON% = 0
        If TIDOC% > 1 Then
          If TIDOC% <> 3 Or InStr(UCase$(TDFAC$), "DEVOLUC") < 1 Then
           VDEF$ = "ND": If TIDOC% = 3 Then VDEF$ = "NC"
           VDEF$ = VDEF$ + String$(94, 0)
           Call REPLA(VDEF$, MKS$(NRO), 3, 4)
           Call REPLA(VDEF$, MKI$(NC%), 7, 2)
           Call REPLA(VDEF$, MKI$(FFI%), 9, 2)
           Call REPLA(VDEF$, MKD$(TOTAFAC#(1, 1)), 19, 8)
2915       XD$ = OBJPLA$("ICOVEN", VDEF$)
           If XD$ = "" Then GoTo 2918
           '
           TOIMPU# = 0
           For QQ% = 17 To 96 Step 10
              CUE% = CVI(Mid$(XD$, QQ%, 2))
              IMPO# = (Int(100 * CVD(Mid$(XD$, QQ% + 2, 8)) + 0.5)) / 100
              If CUE% > 0 Then
                 If ECOARCH$("CUECON", MKI$(CUE%)) <> "" Then
                    TOIMPU# = TOIMPU# + IMPO#
                 End If
              End If
           Next QQ%
           '
           If Abs(TOIMPU# - TOTAFAC#(1, 1)) >= 0.005 Then
               VDEF$ = XD$
               GoTo 2915
           End If
           '
           GRAICON% = 1
           '
          End If
        End If
        '
2918    FINFACEX$ = "": TICAMBX# = 1
        If EXPOR% = 1 Then
          FINFACEX$ = String$(68, 0)
          Call REPLA(FINFACEX$, MKD$(TOTAFAC#(1, 8)), 1, 8)
          Call REPLA(FINFACEX$, "AEREO       ", 9, 12)
          Call REPLA(FINFACEX$, MKD$(0), 21, 8)
          Call REPLA(FINFACEX$, "CFR Bs.As.  ", 29, 12)
          Call REPLA(FINFACEX$, MKD$(TOTAFAC#(1, 8)), 41, 8)
          Call REPLA(FINFACEX$, "DOLARES     ", 49, 12)
          Call REPLA(FINFACEX$, MKD$(0), 61, 8)
          '
          VDEF$ = FINFACEX$
2919      OBX$ = OBJPLA$("FINFACEXPO", VDEF$)
          If OBX$ = "" Then GoTo 2918
          '
          VDEF$ = OBX$
          TICAMBX# = CVD(Mid$(OBX$, 61, 8))
          If TICAMBX# < 0.1 Then
            Call MENSERR(24, "Ingrese Tipo de Cambio")
            GoTo 2919
          End If
          '
          VAFLE# = CVD(Mid$(OBX$, 21, 8))
          VAFOB# = CVD(Mid$(OBX$, 1, 8)) + VAFLE#
          If Abs(VAFOB# - CVD(Mid$(OBX$, 41, 8))) > 0.005 Then
            Call REPLA(VDEF$, MKD$(VAFOB#), 41, 8)
            GoTo 2919
          End If
          '
          FINFACEX$ = OBX$
          TOTAFAC#(1, 9) = VAFLE#
          TOTAFAC#(1, 8) = VAFOB#
          '
        End If
        '
        If EXPOR% <> 1 Then
          If PIVA% = 5 Then ' OPERACIONES DE EXPORTACION
2818        FINFACEX$ = String$(68, 0)
            Call REPLA(FINFACEX$, MKD$(TOTAFAC#(1, 8)), 1, 8)
            Call REPLA(FINFACEX$, "DOLARES     ", 49, 12)
            Call REPLA(FINFACEX$, MKD$(0), 61, 8)
            '
            VDEF$ = FINFACEX$
2819        OBX$ = OBJPLA$("TICAEXPOR", VDEF$)
            If OBX$ = "" Then GoTo 2818
            '
            VDEF$ = OBX$
            TICAMBX# = CVD(Mid$(OBX$, 61, 8))
            If TICAMBX# < 0.1 Then
              Call MENSERR(24, "Ingrese Tipo de Cambio")
              GoTo 2819
            End If
          End If
        End If
        '
        If EXPOR% = 1 Or PIVA% = 5 Then
          For KKI1% = 0 To 10
            TOTAFAC#(1, KKI1%) = TOTAFAC#(1, KKI1%) * TICAMBX#
          Next KKI1%
        End If
        '
        Call BLOQARCH("SVD202")
        Call BLOQARCH("DEUDOR2")
        Call BLOQARCH("FACTUR")
        Call BLOQARCH("MOVISTO")
        Call BLOQARCH("SALSTOCK")
        Call BLOQARCH("MAVENDID")
        '
        URE2021& = ULTREG&("SVD202")
        UREFAC1& = ULTREG&("FACTUR")
        UREMOV1& = ULTREG&("MOVISTO")
        '
        If (NC% > 0 And NC% < 10000) Then
             '
             'Call BLIMESS("Calculando Vencimientos ...")
             Call GRADEU(LU$, ARCHIDEU$, RD%, TIPO%, VA%, NRO, NC%, FFI%, CPAG%, TOTAFAC#(1, 8))
             URE2022& = ULTREG&("SVD202")
             Call CIERRARCH("SVD202")
             '
             SIGX% = 1: If TIPO% = 6 Then SIGX% = (-1)
             TTF# = TOTAFAC#(1, 8) * SIGX%
             '
             'Call BLIMESS("Actualizando Saldo de Cuenta ...")
             Call GRACLI(LU$, ARCHICLI$, NC%, TTF#, DATCLI$(), 2)
             Call CIERRARCH("DEUDOR2")
             '
        End If
        '
        Call GRAFAC98(LU$, ARCHIFAC$, 3, RD%, FLIM%, TIPO%, VA%, CAFOR%, NRO, NC%, CLI$, IVAS$, FFI%, FVI%, LPRE%, VENDEX%, TOTAFAC#())
        Call GRAFAC98(LU$, ARCHIFAC$, 4, RD%, FLIM%, TIPO%, VA%, CAFOR%, NRO, NC%, CLI$, IVAS$, FFI%, FVI%, LPRE%, VENDEX%, TOTAFAC#())
        UREFAC2& = ULTREG&("FACTUR")
        '
        'Call BLIMESS("Grabando Movimientos de Material ...")
2917    NRO1 = NRO: If TIPO% = 6 Then NRO1 = (-1) * NRO
        '
        If InStr(EMPREAC$, "SABO") > 0 Then
           REMI = NRO                         ' para SABO el numero es el mismo
           If EXPOR% = 1 Then
              REMI = NRO - 300
              GoTo 2924
           End If
           '
           If NRO <= 28000 Then
                REMI = NRO - 12500
              Else
                'REMI = NRO - 12200 '12350
                'REMI = NRO - 12050  '12350
                REMI = NRO - 11950 '\\\OT-06-0108-FG-28/03/06///
           End If
           If VA% = 2 Then
              'REMI = NRO + 15500 - 750
              'REMI = NRO + 35900 '\\\OT-06-0050-WAR-15/02/06///
              REMI = NRO + 52900
           End If
        End If
        '
2924    REMI1 = REMI: If TIPO% = 3 Then REMI1 = (-1) * Abs(NRO)
                      If TIPO% = 6 Then REMI1 = NRO
        Call GRAMOVEN(NRO1, REMI1)
        UREMOV2& = ULTREG&("MOVISTO")
        Call CIERRARCH("MOVISTO")
        Call CIERRARCH("MAVENDID")
        Call CIERRARCH("SALSTOCK")
        Call CIERRARCH("SALDEPOS")
        Call BAJALDISCO
        '
        If EXPOR% = 1 Or PIVA% = 5 Then
          For KKI1% = 0 To 10
            TOTAFAC#(1, KKI1%) = TOTAFAC#(1, KKI1%) / TICAMBX#
          Next KKI1%
        End If
        '
2920    Call CARIMPRE(CCI$())
        '
        Call CARSETUP("NEW")
        Call CARSETUP(CCI$(1))            ' POR SI IMPRESORA APAGADA
        If InStr(EMPRELI$, "NOVOESPU") > 0 Then
            Call CARSETUP(CCI$(4))
            Call CARSETUP(CCI$(10) + CCI$(11))
          Else
            Call CARSETUP(CCI$(3))
            Call CARSETUP(CCI$(10) + CCI$(12))
        End If
        '
        VENDEX% = VENDE% + VVX%
        FLIM% = FECHANUM("01/01/99")
        If ULTREG&("FORMFAC") > 0 Then
            FFS$ = REGLEIDO$("FORMFAC", 2&)
            FLIM% = CVI(Mid$(FFS$, 19, 2))
        End If
        '
        CPAGX% = CPAG% + VVX%
        TPC$ = TIPCOM$
        NPR$ = TRIM$(NPROCLI$(NC%))        ' Numero de proveedor
        '
        RTO$ = "": RTO2$ = ""
        If InStr(EMPREAC$, "SABO") > 0 Then GoTo 2930
        '
        For UV& = 1 To ULTREG&("!CARDEFAC")
           Z$ = REGLEIDO$("!CARDEFAC", UV&)
           RTO1$ = TRIM$(Mid$(Z$, 131, 16))
           If RTO1$ = RTO2$ Then RTO1$ = ""
           If RTO1$ <> "" Then
               RTO2$ = RTO1$
               If RTO$ <> "" Then RTO$ = RTO$ + "-"
               RTO$ = RTO$ + RTO1$
           End If
        Next UV&
        If Len(RTO$) > 16 Then RTO$ = "VARIOS"
        '
2930    If RTO$ = "" Then
           REMI = NRO                         ' para SABO el numero es el mismo
           '
           If InStr(EMPREAC$, "SABO") > 0 Then
              If EXPOR% = 1 Or PIVA% = 7 Then
                REMI = NRO - 300
                GoTo 2931
              End If
              '
              If NRO <= 28000 Then
                  REMI = NRO - 12500
                Else
                  'REMI = NRO - 12200 ' 12350
                  'REMI = NRO - 12050 '12350
                  REMI = NRO - 11950 '\\\OT-06-0108-FG-28/03/06///
              End If
              If VA% = 2 Then
                 'REMI = NRO + 15500 - 750
                 'REMI = NRO + 35900 '\\\OT-06-0050-WAR-15/02/06///
                 REMI = NRO + 52900
              End If
           End If
           '
2931       RTO$ = TRIM$(Str$(REMI))
           While Len(RTO$) < 8: RTO$ = "0" + RTO$: Wend
           PREFI$ = ATRIFORM$(FORIMPRE$, "PREFIFAC")
           RTO$ = PREFI$ + "-" + RTO$
        End If
        If TRIM$(Text10.TEXT) <> "" Then
            RTO$ = TRIM$(Text10.TEXT)
        End If
        '
        NROCON& = NRO
        VALOR# = TOTAFAC#(1, 8)
        '
        If FORMAIMPRE$ = "DOS" Then
              Call PRIFADOS
           Else
              GoSub 6000                         ' imprimir
        End If
        '
        NRXX& = NRO + 1
        While NRXX& < NROCON& + 0.1
           NRT = NRXX&
           For I% = 0 To 3: For J% = 0 To 10: TOTAFAC#(I%, J%) = 0: Next J%: Next I%
           Call GRAFAC98(LU$, ARCHIFAC$, 3, RD%, FLIM%, TIPO%, VA%, CAFOR%, NRT, NC%, CLI$, IVAS$, FFI%, FVI%, LPRE%, VENDEX%, TOTAFAC#())
           NRXX& = NRXX& + 1
        Wend
        Call CIERRARCH("FACTUR")
        '
2990    TPXS$ = "Factura"
        If TIPO% = 3 Then TPXS$ = "Nota de Débito"
        If TIPO% = 6 Then TPXS$ = "Nota de Crédito"
        '
        On Error Resume Next
        TITUPA$ = FORMUFAC.Caption
        AppActivate (TITUPA$)
        FORMUFAC.Visible = True
        Call FRESHALL
        FORMUFAC.WindowState = 0
        DoEvents
        On Error GoTo 0
        '
        AI% = COMALTER%("Revise la " + TPXS$ + " Impresa\\Esta O.K.\Anular Documento")
        ''Call FRESHALL
        ''Show
        If AI% = 2 Then
           '
           For U& = URE2021& + 1 To URE2022&
              X$ = REGBLAN$("SVD202")
              Call GRAREG("SVD202", X$, U&)
           Next U&
           '
           For U& = UREFAC1& + 1 To UREFAC2&
              X$ = REGLEIDO$("FACTUR", U&)
              Call REPLA(X$, String$(50, 0), 15, 50)
              Call GRAREG("FACTUR", X$, U&)
           Next U&
           '
           HOII% = HOY(HOS$)
           FIN& = ULTREG&("!CARDEFAC")
           For U& = 1 To FIN&
             Call TRACE(20, U&, FIN&)
             X$ = REGLEIDO$("!CARDEFAC", U&)
             REODES& = CVS(Mid$(X$, 505, 4))
             If REODES& > 0 Then
               Y$ = REGLEIDO$("MAVENDID", REODES&)
               Call REPLA(Y$, MKS$(0), 225, 4)  ' CANTIDAD = 0
               Call REPLA(Y$, Chr$(9), 246, 1) ' MARCA ANULADO
               Call REPLA(Y$, MKI$(HOII%), 251, 2)
               Call REPLA(Y$, MKI$(HORANUM%(Time$)), 253, 2)
               Call REPLA(Y$, Chr$(USNBR% Mod 100), 255, 1)
               Call GRAREG("MAVENDID", Y$, REODES&)
               REMOVI& = CVS(Mid$(Y$, 247, 4))
               If REMOVI& > 0 Then
                 Z$ = REGLEIDO$("MOVISTO", REMOVI&)
                 Call REPLA(Z$, String$(16, 0), 88, 16)
                 Call GRAREG("MOVISTO", Z$, REMOVI&)
               End If
             End If
           Next U&
           '
           Call CIERRARCH("MOVISTO")
           Call CIERRARCH("FACTUR")
           Call CIERRARCH("SVD202")
           Call CIERRARCH("MAVENDID")
           Call CIERRARCH("*.*")
           Call BAJALDISCO
           '
           Call COMUNI("Los Pedidos y Ordenes de Despacho\Correspondientes a este Documento Anulado\Deben Volver a Cargarse !!!")
           '
           GoTo 3500
           '
        End If
        '
3000    Screen.MousePointer = 11
        '
        VENDEX% = VENDE% + VVX%
        Call CIERRARCH("*.*")
        '
        FLIM% = FECHANUM("01/01/99")
        If ULTREG&("FORMFAC") > 0 Then
            FFS$ = REGLEIDO$("FORMFAC", 2&)
            FLIM% = CVI(Mid$(FFS$, 19, 2))
        End If
        '
        Call CIERRARCH("MOVISTO")
        Call CIERRARCH("PEDDETA")
        Call CIERRARCH("PEDDETP")
        Call CIERRARCH("MAVENDID")
        Call BAJALDISCO
        '
      ' imputacion contable de ventas
        '
        If GRAICON% = 1 Then
           '
           XD1$ = REGLEIDO$("ICOVEN", 1)
           MMXX$ = "ENEFEBMARABRMAYJUNJULAGOSETOCTNOVDIC"
           FEX = FFI%: FEFA$ = FECHATEX$(FEX)
           MESI% = Val(Mid$(FEFA$, 4, 2))
           ARCHICO$ = "ICV" + Mid$(MMXX$, 3 * MESI% - 2, 3) + Mid$(FEFA$, 7)
           Call COPYSTRU("ICOVEN", ARCHICO$)
           Call REGAPP(ARCHICO$, XD$)
           '
        End If
        '
        ATRI$ = ""
        '
        Z$ = ""

3500  ' aqui limpia pantalla y vectores ****************
        '
        Call CIERRARCH("*.*")
        Call BAJALDISCO
        '
        BOTNEXT.Enabled = True
        NC% = 0
        Call BLANARCH("!CARDEFAC")
        '
        Text1.TEXT = "": Text1.Refresh
        Text2.TEXT = "": Text2.Refresh
        Text3.TEXT = "": Text3.Refresh
        Text4.TEXT = "": Text4.Refresh
        Text5.TEXT = "": Text5.Refresh
        Text6.TEXT = "": Text6.Refresh
        Text7.TEXT = "": Text7.Refresh
        Text8.TEXT = "": Text8.Refresh
        Text9.TEXT = "": Text9.Refresh
        Text10.TEXT = "": Text10.Refresh
        Text11.TEXT = "": Text11.Refresh
        '
        For KK% = 0 To 2
           TEXT21(KK%).TEXT = ""
           Eco21(KK%) = ""
        Next KK%
        '
        For KK% = 0 To 3
           Total(KK%).TEXT = ""
        Next KK%
        '
        PORDESCU.TEXT = ""
        PORIVA.TEXT = ""
        FFI% = HOY(HO$)
        Text9.TEXT = HO$
        '
        Text1.SetFocus
        Screen.MousePointer = 1
        '
        Exit Sub
        '
'****************************************************************************
'
4000    If X$ = "" Then GoTo 3500
        '
        Call BLANARCH("!FACSEL")
        NX% = FILENBR%("FACTUR")
        Dim AX As String * 64
        FIN& = LOF(NX%) / 64
        For IL& = 1 To FIN&
           'Call TRACE(20, IL&, FIN&)
           Get #NX%, IL&, AX$
           If CVI(Mid$(AX$, 7, 2)) >= DES% Then
              If CVI(Mid$(AX$, 7, 2)) <= HAS% Then
                 NC% = CVI(Left$(AX$, 2))
                 CLI$ = AJUSTI$(RASOCLI$(NC%), 30)
                 X$ = Mid$(AX$, 3, 4) + CLI$ + MKS$(IL&)
                 Call REGAPP("!FACSEL", X$)
              End If
           End If
        Next IL&
        Call CIERRARCH("*.*")
        '
        'X$ = REGSEL$("FACSEL")
        If Len(X$) <= 4 Then GoTo 3500
        REFA& = CVS(Mid$(X$, 35, 4))
        '
        If ALTERNA%("-Confirma Anulación") <> 1 Then GoTo 3500
        If REFA& > 0 Then
            If REFA& <= ULTREG&("FACTUR") Then
               X$ = REGLEIDO$("FACTUR", REFA&)
               Call REPLA(X$, String$(50, 0), 15, 50)
               Call GRAREG("FACTUR", X$, REFA&)
               Call CIERRARCH("*.*")
               Call FINAL("DIAGNOS")
            End If
        End If
        '
        GoTo 3500
'
'****************************************************************************
'
6000    PORIAN = Printer.Orientation  ' retiene orientacion
        CODFOR$ = FORIMPRE$
        GoSub 6800                            ' LEER ATRIBUTOS DEL FORMULARIO
        '
        Printer.Orientation = 1
        If ORIENTA% = 2 Then
            Printer.Orientation = 2
        End If
        '
        Call COMUNI("Verifique Encuadre y Pulse\<ENTER> para Continuar ...")
        PRIDET0& = 1: HOJACONT% = 0: TOTITEMS = 0
6005    TRANSPOR% = 0
        ORDECOP% = 0
           For COP% = 1 To CACOPI%
              PRIDET& = PRIDET0&
              Call CARFORWIN
              GoSub 7100                      ' DATOS GENERALES DE FACTURA
              YAC = DETDESMM
              GoSub 7300: UREDET% = UREN%     ' CARGAR DETALLE DE FACTURACION
              GoSub 7500                      ' CARGAR LEYENDAS
              '
              'If COP% >= PRICOFAC% And COP% <= ULTCOFAC% Then
                  If HOJACONT% = 0 Then        ' SOLO EN PRIMERA HOJA
                     If CONTIHOJA% = 0 Then
                        GoSub 7400                '                CARGAR TOTALES
                     End If
                  End If
              'End If
              '
              If CONTIHOJA% > 0 Then
                  If TRANSPOR% = 0 Then
                      GoSub 7400
                  End If
              End If
              '
              Printer.EndDoc
              '
        Next COP%
        '
        PRIDET0& = PRIDET&
        '
        If TRANSPOR% > 0 Then                 ' TRANSPORTA A HOJA SIGUIENTE
            HOJACONT% = HOJACONT% + 1
            If CONTIHOJA% = 0 Then
                NROCON& = NROCON& + 1
            End If
            '
            Printer.EndDoc
            '
            GoTo 6005
        End If
        '
        If VEPIEDET$ <> "" Then               ' LEYENDAS COMPLEMENTARIAS
           For COP% = 1 To CANCOP%
              Call CARFORWIN
              GoSub 7100                      ' DATOS GENERALES DE FACTURA
              '
              YAC = DETDESMM
              For U& = 1 To ULTREG&("!FACEXPO2")
                  XXX$ = REGLEIDO$("!FACEXPO2", U&)
                  TXT$ = XXX$
                  Call APPFORM(TXT$, "ESPE-DET")
                  YAC = YAC + INTERLI
              Next U&
              '
              Printer.EndDoc    ' imprime hoja
              '
           Next COP%
           If CONTIHOJA% = 0 Then
               NROCON& = NROCON& + 1
           End If
        End If
        '
6007    AI5% = COMALTER%("Puede Repetir Impresión\\Continuar\Repite Impresion")
        ''Call FRESHALL
        'Show
        If AI5% = 2 Then
            NROCON& = NRO                   ' EMPIEZA DESDE EL PRIMERO
            GoTo 6000
        End If
        Printer.Orientation = PORIAN
        '
        Call CIERRARCH("FORMUFAC")
        '
        Return
        '
'****************************************************************************
        '
6800    PRICOFAC% = Val(ATRIFORM$(CODFOR$, "IMP-DES"))
        ULTCOFAC% = Val(ATRIFORM$(CODFOR$, "IMP-HAS"))
        MONPRI$ = ATRIFORM$(CODFOR$, "MON-PRI")
        MONSEC$ = ATRIFORM$(CODFOR$, "MON-SEC")
        PREFI$ = ATRIFORM$(CODFOR$, "PREFIFAC")
        LOIMPLET% = Val(ATRIFORM$(CODFOR$, "LOIMPLET"))
        If LOIMPLET% < 16 Then LOIMPLET% = 32
        '
        MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
        MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
        DETDESMM = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
        DETHASMM = Val(ATRIFORM$(CODFOR$, "DETHASMM"))
        INTERLI = Val(ATRIFORM$(CODFOR$, "INTERLIN"))
        If INTERLI < 1 Then
            Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
            INTERLI = 25.4 / 6
        End If
        '
        CACOPI% = Val(ATRIFORM$(CODFOR$, "CAN-COP"))
        If CACOPI% < 1 Or CACOPI% > 6 Then
            Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
            CACOPI% = 1
        End If
        '
        ORIENTA% = 1
        PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "ORIENTA")))
        If PORI$ = "PAISAJE" Or PORI$ = "2" Then
            ORIENTA% = 2
        End If
        '
        Return
        '
'****************************************************************************
        '
7100    FEX = FFI%
        TCMP$ = UCase$(TIPCOM$): If COP% > ULTCOFAC% Then TCMP$ = "REMITO"
        Call PRINTFORWIN(TCMP$, "TIPO-COM", 0, 0)
        
        Call PRINTFORWIN(FECHATEX$(FEX), "FECH-EMI", 0, 0)
        Call PRINTFORWIN(Text2.TEXT, "RASO-CLI", 0, 0)
        Call PRINTFORWIN(Text8.TEXT, "CUIT-CLI", 0, 0)
        Call PRINTFORWIN(Text3.TEXT, "DOMI-CLI", 0, 0)
        Call PRINTFORWIN(NIBRCLI$(NC%), "IBRU-CLI", 0, 0)
        Call PRINTFORWIN(Text4.TEXT, "CPOS-CLI", 0, 0)
        Call PRINTFORWIN(Text5.TEXT, "LOCA-CLI", 0, 0)
        Call PRINTFORWIN(NPROCLI$(NC%), "PROV-CLI", 0, 0)
        Call PRINTFORWIN(Str$(NC%), "CUEN-CLI", 0, 0)
        POSIVA$ = ECOARCH$("CAPOSIV", Chr$(PIVA%))
        Call PRINTFORWIN(POSIVA$, "PIVA-CLI", 0, 0)
        TCMP$ = UCase$(CATIVAS$): If COP% > ULTCOFAC% Then TCMP$ = "X"
        Call PRINTFORWIN(TCMP$, "CIVA-CLI", 0, 0)
        Call PRINTFORWIN(DECONPA$(CPAG%), "COND-PAG", 0, 0)
        '
        Call PRINTFORWIN(RTO$, "NRO-REMI", 0, 0)
        '
        Call PRINTFORWIN(Mid$(TRXP$, 1, 32), "TRANSPOR", 0, 0)
        Call PRINTFORWIN(Mid$(TRXP$, 33, 32), "DOMI-TRA", 0, 0)
        Call PRINTFORWIN(Mid$(TRXP$, 65, 16), "CUIT-TRA", 0, 0)
        Call PRINTFORWIN(Mid$(TRXP$, 81, 16), "PATE-TRA", 0, 0)
        Call PRINTFORWIN(Mid$(TRXP$, 97, 32), "COND-TRA", 0, 0)
        '
        NCN$ = TRIM$(Str$(NROCON&))
        While Len(NCN$) < 8: NCN$ = "0" + NCN$: Wend
        PREFI$ = ATRIFORM$(CODFOR$, "PREFIFAC")
        NCN$ = PREFI$ + "-" + NCN$
        Call PRINTFORWIN(NCN$, "NUME-COM", 0, 0)
        COPXX% = COP%: If COPXX% > ULTCOFAC% Then COPXX% = COP% - ULTCOFAC%
        NCN$ = ORCOP$(COPXX%)
        Call PRINTFORWIN(NCN$, "ORD-COPI", 0, 0)
         '
7190    Return
        '
'****************************************************************************
        '
7300    If PRIDET& = 1 Then
           If VENCADET$ <> "" Then
              For U& = 1 To ULTREG&("!FACEXPO1")
                  XXX$ = REGLEIDO$("!FACEXPO1", U&)
                  TXT$ = XXX$
                  Call APPFORM(TXT$, "ENCA-DET")
                  RG0% = RGII% + 1
              Next U&
              YAC = IAC + INTERLI
           End If
        End If
        '
        For U& = PRIDET& To ULTREG&("!CARDEFAC")
            '
            If YAC > DETHASMM - 4 * INTERLI Then                      ' SALTO DE HOJA
               PRIDET& = U&
               TRANSPOR% = 1
               GoTo 7390
            End If
            '
            XXX$ = REGLEIDO$("!CARDEFAC", U&)
            '
            TOTITEMS = TOTITEMS + CVS(Left$(XXX$, 4))
            TXT$ = Str$(CVS(Left$(XXX$, 4)))       ' cantidad
                Call APPFORM(TXT$, "CANTIDAD")
                '
            TXT$ = CODEXT$(CVI(Mid$(XXX$, 5, 2)))     ' codigo
                Call APPFORM(TXT$, "COD-MAT")
                '
            TXT$ = Mid$(XXX$, 7, 64)                  ' descripcion
                Call APPFORM(TXT$, "DES-MAT")
                '
            TXT$ = Mid$(XXX$, 71, 4)                  ' unidad de medida
                Call APPFORM(TXT$, "UNIMED")
                '
            TXT$ = Mid$(XXX$, 75, 64)                 ' referencia 1
                Call APPFORM(TXT$, "REF-MAT1")
                '
            TXT$ = Mid$(XXX$, 139, 64)                ' referencia 2
                Call APPFORM(TXT$, "REF-MAT2")
                '
            TXT$ = Mid$(XXX$, 203, 64)                ' referencia 3
                Call APPFORM(TXT$, "REF-MAT3")
                '
            TXT$ = Mid$(XXX$, 315, 24)                ' numero de plano
                Call APPFORM(TXT$, "N-PLANO")
                '
            TXT$ = Mid$(XXX$, 339, 16)                ' numero original
                Call APPFORM(TXT$, "N-ORIGIN")
                '
            TXT$ = Mid$(XXX$, 355, 16)                ' numero de o.compra
                Call APPFORM(TXT$, "N-OCOMPR")
                '
            TXT$ = Mid$(XXX$, 387, 16)                ' numero de remito
                Call APPFORM(TXT$, "N-REMITO")
                '
            TXT$ = Mid$(XXX$, 403, 16)                ' numero de doc-recepcion
                Call APPFORM(TXT$, "N-BOLREC")
                '
            FEX = CVI(Mid$(XXX$, 419, 2))             ' fecha de recepcion
                TXT$ = FECHATEX$(FEX)
                Call APPFORM(TXT$, "FECH-REC")
                '
          If COP% >= PRICOFAC% And COP% <= ULTCOFAC% Then
              TXT$ = Str$(FACIVA * CVS(Mid$(XXX$, 421, 4))) ' precio de lista
                Call APPFORM(TXT$, "PRE-LIST")
                '
              TXT$ = Str$(FACIVA * CVS(Mid$(XXX$, 429, 4))) ' precio unitario
                Call APPFORM(TXT$, "PRE-NETO")
                '
              TXT$ = Str$(CVD(Mid$(XXX$, 433, 8)))   ' total neto
                Call APPFORM(TXT$, "NET-ITEM")
                '
              TXT$ = Str$(CVS(Mid$(XXX$, 441, 8)))   ' alicuota iva
                Call APPFORM(TXT$, "ALICUIVA")
                '
              TXT$ = Str$(CVD(Mid$(XXX$, 445, 8)))   ' importe iva del item
                Call APPFORM(TXT$, "IVA-ITEM")
                '
              TXT$ = Str$(CVD(Mid$(XXX$, 453, 8)))   ' total bruto
                Call APPFORM(TXT$, "BRU-ITEM")
                '
            Else
               '
               TXT$ = "********************"
               Call APPFORM(TXT$, "PRE-LIST")
               Call APPFORM(TXT$, "PRE-NETO")
               Call APPFORM(TXT$, "NET-ITEM")
               Call APPFORM(TXT$, "ALICUIVA")
               Call APPFORM(TXT$, "IVA-ITEM")
               Call APPFORM(TXT$, "BRU-ITEM")
               '
          End If
          '
          YAC = YAC + INTERLI
          '
        Next U&
        '
        YAC = YAC + INTERLI
        Call APPFORM(String$(32, "="), "LIN-CIE")
        YAC = YAC + INTERLI
        Call APPFORM(Str$(TOTITEMS), "TOTA-ITS")
        Call APPFORM("Total de Items", "LETOT-IT")
        '
7390    Return
        '
'****************************************************************************
        '
7400    If COP% > ULTCOFAC% Then GoTo 7410
        '
        Call PRINTFORWIN(Str$(FACIVA * TOTAFAC#(1, 1)), "IMP-NETO", 0, 0)
        DESCUE$ = TRIM$(PORDESCU.TEXT)
        Call PRINTFORWIN(DESCUE$, "POR-DESC", 0, 0)
        Call PRINTFORWIN(Str$(FACIVA * TOTAFAC#(1, 4)), "IMP-DESC", 0, 0)
        Call PRINTFORWIN(Str$(FACIVA * TOTAFAC#(1, 5)), "IMPNEDES", 0, 0)
        Call PRINTFORWIN(Str$(TOTAFAC#(1, 6)), "IMP-IVA", 0, 0)
        Call PRINTFORWIN(Str$(TOTAFAC#(1, 7)), "IMP-RNI", 0, 0)
        Call PRINTFORWIN(Str$(TOTAFAC#(1, 8)), "IMP-TOTA", 0, 0)
        Call PRINTFORWIN(Str$(TOTAFAC#(1, 9)), "IMP-VARS", 0, 0)
        Call PRINTFORWIN(Str$(TOTAFAC#(1, 10)), "PERC-IVA", 0, 0)
        '
        TTTT# = TOTAFAC#(1, 8)
        TOTFAC$ = NUMTEX$(TTTT#)
        Call FRATEX(TOTFAC$, NRX1$, NRX2$, NRX3$, LOIMPLET%)
        Call PRINTFORWIN(NRX1$, "IMP-LET1", 0, 0)
        Call PRINTFORWIN(NRX2$, "IMP-LET2", 0, 0)
        Call PRINTFORWIN(NRX3$, "IMP-LET3", 0, 0)
        '
        Return
        '
7410    Call PRINTFORWIN("******************", "IMP-NETO", 0, 0)
        Call PRINTFORWIN("****************", "POR-DESC", 0, 0)
        Call PRINTFORWIN("****************", "IMP-DESC", 0, 0)
        Call PRINTFORWIN("****************", "IMPNEDES", 0, 0)
        Call PRINTFORWIN("****************", "IMP-IVA", 0, 0)
        Call PRINTFORWIN("****************", "IMP-RNI", 0, 0)
        Call PRINTFORWIN("****************", "IMP-TOTA", 0, 0)
        Call PRINTFORWIN("****************", "IMP-VARS", 0, 0)
        Call PRINTFORWIN("****************", "PERC-IVA", 0, 0)
        '
        Call PRINTFORWIN(String$(80, "*"), "IMP-LET1", 0, 0)
        Call PRINTFORWIN(String$(80, "*"), "IMP-LET2", 0, 0)
        Call PRINTFORWIN(String$(80, "*"), "IMP-LET3", 0, 0)
        '
        Return
        '
'****************************************************************************
        '
7500    Call CARLEYEN("")
        '
        Return
        '
'****************************************************************************
        '
'****************************************************************************
'
8200    ' calcular IVA de cada item
        '
        For U& = 1 To ULTREG&("!CARDEFAC")
            X$ = REGLEIDO$("!CARDEFAC", U&)
            TOT# = CVD(Mid$(X$, 433, 8))
            IVAITEM# = TOT# * (ALICU / 100)
            BRUITEM# = TOT# + IVAITEM#
            '
            Call REPLA(X$, MKS$(ALICU), 441, 4)
            Call REPLA(X$, MKD$(IVAITEM#), 445, 8)
            Call REPLA(X$, MKD$(BRUITEM#), 453, 8)
            '
            Call GRAREG("!CARDEFAC", X$, U&)
        Next U&
        '
        Return
'
'****************************************************************************
'
8700 TOT1# = 0: TOT2# = 0: TOT3# = 0: TOT4# = 0
     RGXI% = RGII%: CLII% = 66: LCII% = 13: TV% = 3: LD% = 2: JTII% = 2
8702 For JJ& = 1 To ULTREG&("!CARDEFAC")
        TOT# = CVD(Mid$(REGLEIDO$("!CARDEFAC", JJ&), 433, 8) + String$(8, 0))
8715 TOT1# = (Int(100 * TOT1# + 0.5)) / 100 + (Int(100 * TOT# + 0.5)) / 100
8720 Next JJ&
8722 '
8727 'por ahora los descuentos se aplican directamente sobre el pr.unit
     TOT2# = 0
     If PORDES > 0 Then
        If MODESCU% = 1 Then
8730        TOT2# = (Int(TOT1# * PORDES + 0.5)) / 100
8732        TOTAFAC#(1, 9) = (Int(TOTAFAC#(1, 9) * (100 - PORDES) + 0.5)) / 100
        End If
     End If
8735 '
8743 TOT3# = (Int((TOT1# - TOT2#) * ALICU + 0.5)) / 100:
     TOT4# = TOT1# - TOT2# + TOT3#
8748 '
8750 If FACIVA <> 1 Then GoTo 8760
8751 RGII% = 19: VD# = TOT1#
     Total(0).TEXT = CSTRING$(MKD$(VD#), 4, 18, 2, 2)
        TOTAFAC#(1, 1) = TOT1#
        TOTAFAC#(1, 3) = TOT1#
8752 RGII% = 20: VD# = TOT2#
     Total(1).TEXT = CSTRING$(MKD$(VD#), 4, 18, 2, 2)
     'Call PRESENTA(20, 0, CL%, LC%, LK%, TV%, LD%, JT%, VDI%, TOT2#, VDS$, CT%)
        TOTAFAC#(1, 4) = TOT2#
        TOTAFAC#(3, 4) = 0
                If DESCUPRE% = 0 Then
                        TOTAFAC#(3, 4) = CDbl(PORDES)
                End If
        TOTAFAC#(1, 5) = ((Int(100 * TOTAFAC#(1, 3) + 0.5)) - (Int(100 * TOTAFAC#(1, 4) + 0.5))) / 100
8753 RGII% = 21: VD# = TOT3#
     Total(2).TEXT = CSTRING$(MKD$(VD#), 4, 18, 2, 2)
     'Call PRESENTA(21, 0, CL%, LC%, LK%, TV%, LD%, JT%, VDI%, TOT3# + TOTAFAC#(1, 10), VDS$, CT%)
        TOTAFAC#(1, 6) = (Int(TOTAFAC#(1, 5) * ALI1 + 0.5)) / 100
        TOTAFAC#(3, 6) = CDbl(ALI1)
        TOTAFAC#(1, 7) = (Int(TOTAFAC#(1, 5) * ALI2 + 0.5)) / 100
        TOTAFAC#(3, 7) = CDbl(ALI2)
        If ALICU = ALI1 Then TOTAFAC#(1, 7) = 0: TOTAFAC#(3, 7) = 0
        If ALICU < 0.01 Then TOTAFAC#(1, 6) = 0: TOTAFAC#(3, 6) = 0: TOTAFAC#(1, 7) = 0: TOTAFAC#(3, 7) = 0
        TOTAFAC#(1, 8) = TOTAFAC#(1, 5) + TOTAFAC#(1, 6) + TOTAFAC#(1, 7) + TOTAFAC#(1, 9) + TOTAFAC#(1, 10)
8754 RGII% = 22: VD# = TOT4#
     Total(3).TEXT = CSTRING$(MKD$(VD#), 4, 18, 2, 2)
     'Call PRESENTA(22, 0, CL%, LC%, LK%, TV%, LD%, JT%, VDI%, TOT4# + TOTAFAC#(1, 9) + TOTAFAC#(1, 10), VDS$, CT%)
8756 GoTo 8768
8757 '
8760 RGII% = 19: VD# = TOT4#
     'Call PRESENTA(19, 0, CL%, LC%, LK%, TV%, LD%, JT%, VDI%, TOT4#, VDS$, CT%)
        TOTAFAC#(1, 1) = TOT1#
        TOTAFAC#(1, 3) = TOT1#
8762 RGII% = 20: VD# = TOT2#
     'Call PRESENTA(20, 0, CL%, LC%, LK%, TV%, LD%, JT%, VDI%, TOT2#, VDS$, CT%)
        TOTAFAC#(1, 4) = TOT2#
        TOTAFAC#(3, 4) = 0
                If DESCUPRE% = 0 Then
                        TOTAFAC#(3, 4) = CDbl(PORDES)
                End If
        TOTAFAC#(1, 5) = ((Int(100 * TOTAFAC#(1, 3) + 0.5)) - (Int(100 * TOTAFAC#(1, 4) + 0.5))) / 100
8763 RGII% = 21: VD# = TOT4#
     'Call PRESENTA(21, 0, CL%, LC%, LK%, TV%, LD%, JT%, VDI%, TOT4#, VDS$, CT%)
        TOTAFAC#(1, 8) = TOT4# + TOTAFAC#(1, 10) ' AGREGA PERCEPCION
8764 RGII% = 22: VD# = TOT3#
     'Call PRESENTA(22, 0, CL%, LC%, LK%, TV%, LD%, JT%, VDI%, TOT3#, VDS$, CT%)
        TOTAFAC#(1, 6) = TOT3#
        TOTAFAC#(1, 7) = 0

8768 For I% = 0 To 3
        Total(I%).Refresh
     Next I%
     RGII% = RGXI%

8800 For I% = 1 To 10
        TOTAFAC#(2, I%) = TOTAFAC#(1, I%) / TICA
     Next I%

     Return

End Sub
'****************************************************************************

Sub FORMPRIN(LU$, FLIM%, ALI1, ALI2, TC1, TLET$, INTER%, TFOR%, ANFO%, LARFO%, DIFER, CAFOR%, MARIZ%, MARINT%, REDON%, MPRI$, MCEN$, FACTO, CANCON%, COMA%, PARDEC%)
      '
      If ULTREG&("FORMFAC") > 0 Then
            FFS$ = REGLEIDO$("FORMFAC", 2&)
            ALI1 = CVS(Mid$(FFS$, 2, 4))
            ALI2 = CVS(Mid$(FFS$, 7, 4))
            TC1 = CVS(Mid$(FFS$, 11, 4))
            '
            TFOR% = Asc(Mid$(FFS$, 6, 1))
            CAFOR% = 2: If TFOR% > 10 Then CAFOR% = (TFOR% Mod 10): TFOR% = Int(TFOR% / 10)
            MPRI$ = TRIM$(Mid$(FFS$, 24, 30))
            MCEN$ = TRIM$(Mid$(FFS$, 54, 30))
            '
            If TFOR% < 1 Or TFOR% > 2 Then TFOR% = 1
         Else
            Call MENSERR(24, "Falta Archivo de Formato de Impresion de Factura")
      End If
      '
End Sub
'
Sub SELTIFAC()
    '
    TDFAC$ = REGSEL$("SELTIFAC")
    If Len(TDFAC$) <= 4 Then
        Call CIERRARCH("*.*")
        Call FINAL("")
    End If
    '
    RGTIFA& = CVS(Mid$(TDFAC$, 129, 4))
    CONTROFA$ = REGLEIDO$("TIDOCFAC", RGTIFA&)
    Call CIERRARCH("TIDOCFAC")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
          ENCABE$ = "Facturación - " + EPAC$ + " - "
        Else
          ENCABE$ = "Facturación - "
    End If
    ENCABE$ = ENCABE$ + TRIM$(Left$(TDFAC$, 56))
    Caption = ENCABE$
    '
    TIDOC% = Asc(Left$(CONTROFA$, 1))
    VTIP2$ = Space$(9)
    '
    If TIDOC% > 1 Then
      TEXT21(1).TEXT = " 1"  ' fuerza C.P.=Contado
    End If
    '
    On TIDOC% GoTo 210, 220, 230
    If InStr(Caption, "Anulacion de Comprobantes") > 0 Then
      Exit Sub
    End If
    Call FINAL("")
    '
210 TIPO% = 2: VTIP2$ = "Remito(s)"
    TIPCOM$ = "Factura"
    MONEDA% = 1
    NCC$ = TRIM$(Text1.TEXT)
    If Val(NCC$) > 0 Then
      If Val(NCC$) < 32767 Then
        NCXX% = Val(NCC$)
        If ECOARCH$("DEUDOR1", MKI$(NCXX%)) <> "" Then
          TEXT21(1).TEXT = CPAGCLI%(NCXX%)
        End If
      End If
    End If
    GoTo 300
    '
220 TIPO% = 3
    TIPCOM$ = "N.Debito"
    MONEDA% = 1: GoTo 300
    '
230 TIPO% = 6
    TIPCOM$ = "Credito"
    MONEDA% = 1: GoTo 300
    '
300 VARIFA$ = TRIM$(TIPCOM$) + " - " + TRIM$(Mid$(CONTROFA$, 33, 32))
    '
End Sub
'
Sub CARODESPEN()
   '
   Screen.MousePointer = 11
   '
   Call SETULTREG("!CARDEFAC", 0)
   Call CIERRARCH("!CARDEFAC")
   FINF = 32767: FSUP = 0: CAPED% = 0
   REBLA$ = REGBLAN$("CARDEFAC")
   NCXI% = Val(Text1)
   FORSELPED.LIPEDPE.Clear
   RPEDX$ = RECFILT$("MAVENDID", 19, MKI$(NCXI%))
   NRODESA& = 0
   FIN& = Len(RPEDX$)
   '
   For UU& = 1 To FIN& Step 4
     Call TRACE(20, UU&, FIN&)
     U& = CVS(Mid$(RPEDX$, UU&, 4))
     X$ = REGLEIDO$("MAVENDID", U&)
     If CVI(Mid$(X$, 1, 2)) = NCXI% Then   ' coincide cliente
       If Asc(Mid$(X$, 63, 1)) = 1 Then    ' STATUS = 1
         CAN = CVS(Mid$(X$, 177, 4))
         CADESPACH = CVS(Mid$(X$, 225, 4))
         '
         If CADESPACH >= 0.05 Then         ' MARCA COMO CUMPLIDA
           Call REPLA(X$, Chr$(2), 63, 1)
           Call REPLA(X$, MKI$(0), 65, 2)
           Call REPLA(X$, MKI$(0), 67, 2)
           Call GRAREG("MAVENDID", X$, U&)
           Call CIERRARCH("MAVENDID")
           GoTo 19
         End If
         '
         CAPENDE = CAN - CADESPACH
         If CAN >= 0.05 Then
           NRODES& = CVS(Mid$(X$, 87, 4))
           If NRODES& <> NRODESA& Then
             TTXX$ = AJUSTI$("OD-" + TRIM$(CSTRING$(MKS$(NRODES&), 3, 8, 0, 2)), 11)
             FEX = CVI(Mid$(X$, 125, 2))
             TTXX$ = TTXX$ + "   Sol." + FECHATEX$(FEX)
             FORSELPED.LIPEDPE.AddItem TTXX$
           End If
           NRODESA& = NRODES&
           CAPED% = CAPED% + 1
           FENTRE% = CVI(Mid$(X$, 125, 2))
           If FENTRE% < FINF Then FINF = FENTRE%
           If FENTRE% > FSUP Then FSUP = FENTRE%
         End If
       End If
     End If
19 Next UU&
   '
   If FORSELPED.LIPEDPE.ListCount < 1 Then
     Screen.MousePointer = 1
     Call COMUNI("No hay Ordenes de Despacho en Curso\para " + RASOCLI$(NCXI%) + ".")
     GoTo 99
   End If
   '
   Screen.MousePointer = 1
   FORSELPED.Show 1
   '
   CAPESEL% = 0
   For U& = 1 To FORSELPED.LIPEDPE.ListCount
     If FORSELPED.LIPEDPE.Selected(U& - 1) = True Then
       CAPESEL% = CAPESEL% + 1
       TTXX$ = FORSELPED.LIPEDPE.List(U& - 1)
       NUPEDSEL&(CAPESEL%) = Val(Left$(TTXX$, 10))
       If Left$(TTXX$, 3) = "OD-" Then
           NUPEDSEL&(CAPESEL%) = Val(Mid$(TTXX$, 4, 10))
         ElseIf Left$(TTXX$, 3) = "RT-" Then
           NUPEDSEL&(CAPESEL%) = (-1) * Val(Mid$(TTXX$, 4, 10))
       End If
     End If
   Next U&
   '
   If CAPESEL% < 1 Then GoTo 99
   '
   LPRE% = Val(TEXT21(0))
   'If LPRE% < 1 Then LPRE% = 1
   ARCHIPRE$ = TRIM$(Str$(LPRE%))
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   '
   CAITCAR% = 0
   FIN& = Len(RPEDX$)
   SUMATO# = 0
   For UU& = 1 To Len(RPEDX$) Step 4
     If MAXITEMS% > 0 Then
       If CAITCAR% >= MAXITEMS% Then
         GoTo 99
       End If
     End If
     U& = CVS(Mid$(RPEDX$, UU&, 4))
     X$ = REGLEIDO$("MAVENDID", U&)
     REGODES& = U&
     REGPEDA& = CVS(Mid$(X1$, 99, 4))
     NODES& = CVS(Mid$(X$, 87, 4))
     '
     For UI% = 1 To CAPESEL%
       If NUPEDSEL&(UI%) = NODES& Then GoTo 20
     Next UI%
     GoTo 39
     '
20   CAN = CVS(Mid$(X$, 177, 4)) - CVS(Mid$(X$, 225, 4))
     If CAN >= 0.05 Then
       CI% = CVI(Mid$(X$, 5, 2))
       '
       If VERISTO% = 1 Then
         If CI% > 0 Then
           Call VERISAL(CI%)
           STACT = STOCKACT(CI%)
           STDIS = STACT - CAN
           If STDIS < 0 Then
             If TIDOC% <> 3 Then
               CAN = STACT
               If CAN < 0.05 Then GoTo 39
             End If
           End If
         End If
       End If
       '
       Z$ = REBLA$
       Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
           COD$ = CODEXT$(CI%)
       Call REPLA(Z$, MKI$(CI%), 5, 2)
       Call REPLA(Z$, DESCRIT$(CI%), 7, 64)
           TACOMED$ = Mid$(X$, 25, 8) + Mid$(X$, 9, 12) + Mid$(X$, 33, 12)
       Call REPLA(Z$, TACOMED$, 283, 32)
           PRELIS = CVD(Mid$(X$, 127, 8))
           If PRELIS < 0.0001 Then PRELIS = CVD(Mid$(X$, 139, 8))
           If PRELIS < 0.0001 Then PRELIS = PRELISTA#(ARCHIPRE$, CI%)
           PRELIS = (Int(10000 * PRELIS + 0.5)) / 10000
       Call REPLA(Z$, MKS$(PRELIS), 421, 4)
           PRECIO = CVD(Mid$(X$, 139, 8))
           If PRECIO < 0.0001 Then PRECIO = PRELIS
           PRECIO = (Int(10000 * PRECIO + 0.5)) / 10000
       Call REPLA(Z$, MKS$(PRECIO), 429, 4)
           IMPO# = (Int(100 * CDbl(PRECIO) * CAN + 0.5)) / 100
       Call REPLA(Z$, MKD$(IMPO#), 433, 8)
       Call REPLA(Z$, MKS$(REGODES&), 505, 4)
       Call REPLA(Z$, MKS$(REGPEDA&), 509, 4)
       '
       Call REGAPP("!CARDEFAC", Z$)
       CAITCAR% = CAITCAR% + 1
       '
     End If
     '
39 Next UU&
   '
99 Screen.MousePointer = 1
   '
End Sub

Sub PRIFADOS()

Call BAJALDISCO
6000 Call CLEARFORMDOS(FORIMPRE$)
     GoSub 6800                            ' LEER ATRIBUTOS DEL FORMULARIO
     '
     PRIDET& = 1: HOJACONT% = 0: TOTITEMS = 0
     PVCL% = PIVACLI%(NC%)
     If PVCL% = 5 Then  ' FACTURAS DE EXPORTACION
       NXFC% = FILEOPEN%(LUCOM$ + "FC" + TRIM$(Str$(NROCON&)) + ".TXT", 2, 128)
     End If
     '
6005 TRANSPOR% = 0
     For COP% = 1 To CANCOP%
         Call CLEARFORMDOS(FORIMPRE$)     ' GENERAR FORMULARIO EN BLANCO
         GoSub 7100                       ' DATOS GENERALES DE FACTURA
         GoSub 7300: UREDET% = UREN%      ' CARGAR DETALLE DE FACTURACION
         GoSub 7500                       ' CARGAR LEYENDAS
         '
         'If COP% >= PRICOFAC% And COP% <= ULTCOFAC% Then
             If HOJACONT% = 0 Then        ' SOLO EN PRIMERA HOJA
               GoSub 7400                 ' CARGAR TOTALES
             End If
         'End If
         '
         Call PRINTFORMDOS("")            ' IMPRIMIR
         If PVCL% = 5 Then
           If COP% = 1 Then
             FORMACT$ = FORIMPRE$
             LUX$ = LUCOM$
             NYYY% = FILEOPEN%(LUX$ + "\" + FORMACT$, 0, 256)
             Dim FYYYY As String * 256
             For NNN444% = 1 To LOF(NYYY%) / 256
               Get #NYYY%, NNN444%, FYYYY$
               TTT555$ = FYYYY$
               Print #NXFC%, TTT555$
             Next NNN444%
             Close #NYYY%
             For KKI% = 1 To 10: Print #NXFC%, "   ": Next KKI%
           End If
         End If
         '
    Next COP%
    '
    If TRANSPOR% > 0 Then                 ' TRANSPORTA A HOJA SIGUIENTE
        HOJACONT% = HOJACONT% + 1
        NROCON& = NROCON& + 1
        GoTo 6005
    End If
    '
    If VEPIEDET$ <> "" Then               ' LEYENDAS COMPLEMENTARIAS
        NROCON& = NROCON& + 1
        For COP% = 1 To CANCOP%
           Call CLEARFORMDOS(FORIMPRE$)       ' GENERAR FORMULARIO EN BLANCO
           GoSub 7100                      ' DATOS GENERALES DE FACTURA
           '
           RG0% = PRIREDET%: UREN% = RG0% - 1
           For U& = 1 To ULTREG&("!FACEXPO2")
              XXX$ = REGLEIDO$("!FACEXPO2", U&)
              RGII% = RG0%
              TXT$ = XXX$
              Call APPFORMDOS("", RGII%, TXT$, "ESPE-DET")
              RG0% = RGII% + 1
           Next U&
           '
           Call PRINTFORMDOS("")              ' IMPRIMIR
           If PVCL% = 5 Then
             If COP% = 1 Then
               FORMACT$ = FORIMPRE$
               LUX$ = LUCOM$
               NYYY% = FILEOPEN%(LUX$ + "\" + FORMACT$, 0, 256)
               Dim FYYYZ As String * 256
               For NNN444% = 1 To LOF(NYYY%) / 256
                 Get #NYYY%, NNN444%, FYYYZ$
                 TTT555$ = FYYYZ$
                 Print #NXFC%, TTT555$
               Next NNN444%
               Close #NYYY%
               For KKI% = 1 To 10: Print #NXFC%, "   ": Next KKI%
             End If
           End If
           '
        Next COP%
    End If
    '
    If PVCL% = 5 Then
      If NXFC% > 0 Then Close #NXFC%: NXFC% = 0
      Call COMUNI("Archivo de Texto '" + "FC" + TRIM$(Str$(NROCON&)) + ".TXT" + "'\Generado en " + REPLACAR$(LUCOM$, "\", "/") + ".")
    End If
    '
    ''Call FRESHALL
    'AI5% = COMALTER%("Puede Repetir Impresión\\Continuar\Repetir Impresión")
    'If AI5% = 2 Then
    '    NROCON& = NRO                   ' EMPIEZA DESDE EL PRIMERO
    '    GoTo 6000
    'End If
    '
    Call CIERRARCH("FORMUFAC")
    '
    Exit Sub
    '
'****************************************************************************
    '
    '
6800 TILETRA$ = ATRIFORMDOS$("", "TI-LETRA")
     LINPULG% = Val(ATRIFORMDOS$("", "LIN-PULG"))
     CACOL% = Val(ATRIFORMDOS$("", "CAN-COL"))
     CAREN% = Val(ATRIFORMDOS$("", "CAN-REN"))
     MARSUP% = Val(ATRIFORMDOS$("", "MARG-SUP"))
     MARIZQ% = Val(ATRIFORMDOS$("", "MARG-IZQ"))
     CANCOP% = Val(ATRIFORMDOS$("", "CAN-COP"))
     PRICOFAC% = Val(ATRIFORMDOS$("", "IMP-DES"))
     ULTCOFAC% = Val(ATRIFORMDOS$("", "IMP-HAS"))
     PRIREDET% = Val(ATRIFORMDOS$("", "DET-DES"))
     ULTREDET% = Val(ATRIFORMDOS$("", "DET-HAS"))
     MONPRI$ = ATRIFORMDOS$("", "MON-PRI")
     MONSEC$ = ATRIFORMDOS$("", "MON-SEC")
     PREFI$ = ATRIFORMDOS$("", "PREFIFAC")
     '
     Return
     '
'****************************************************************************
     '
7100 If UCase$(App.EXEName) = "RFAC2001" Then
          FEX = FECHANUM(TRIM$(FORMRAC.Text9.TEXT))
          CPAG% = Val(FORMRAC.TEXT21(1).TEXT)
        Else
          FEX = FECHANUM(TRIM$(Text9.TEXT))
          CPAG% = Val(TEXT21(1).TEXT)
     End If
     
     Call SETFORMDOS("", TIPCOM$, "TIPO-COM")
     Call SETFORMDOS("", FECHATEX$(FEX), "FECH-EMI")
     Call SETFORMDOS("", RASOCLI$(NC%), "RASO-CLI")
     Call SETFORMDOS("", CUITCLI$(NC%), "CUIT-CLI")
     Call SETFORMDOS("", DOMICLI$(NC%), "DOMI-CLI")
     Call SETFORMDOS("", NIBRCLI$(NC%), "IBRU-CLI")
     Call SETFORMDOS("", CPOSCLI$(NC%), "CPOS-CLI")
     Call SETFORMDOS("", LOCACLI$(NC%), "LOCA-CLI")
     Call SETFORMDOS("", NPROCLI$(NC%), "PROV-CLI")
     Call SETFORMDOS("", Str$(NC%), "CUEN-CLI")
     Call SETFORMDOS("", POSIVAL$(NC%), "PIVA-CLI")
     Call SETFORMDOS("", DECONPA$(CPAG%), "COND-PAG")
     '
     Call SETFORMDOS("", RTO$, "NRO-REMI")
     '
     Call SETFORMDOS("", Mid$(TRXP$, 1, 32), "TRANSPOR")
     Call SETFORMDOS("", Mid$(TRXP$, 33, 32), "DOMI-TRA")
     Call SETFORMDOS("", Mid$(TRXP$, 65, 16), "CUIT-TRA")
     Call SETFORMDOS("", Mid$(TRXP$, 81, 16), "PATE-TRA")
     Call SETFORMDOS("", Mid$(TRXP$, 97, 32), "COND-TRA")
     '
     NCN$ = TRIM$(Str$(NROCON&))
     While Len(NCN$) < 8: NCN$ = "0" + NCN$: Wend
     PREFI$ = ATRIFORMDOS$("", "PREFIFAC")
     NCN$ = PREFI$ + "-" + NCN$
     Call SETFORMDOS("", NCN$, "NUME-COM")
     '
     TTTT# = TOTAFAC#(1, 8)
     TOTFAC$ = NUMTEX$(TTTT#)
     Call SETFORMDOS("", TOTFAC$, "IMP-LET1")
     '
7190 Return
     '
'****************************************************************************
     '
7300 RG0% = PRIREDET%: UREN% = RG0% - 1
     '
     If PRIDET& = 1 Then
       If VENCADET$ <> "" Then
          For U& = 1 To ULTREG&("!FACEXPO1")
          XXX$ = REGLEIDO$("!FACEXPO1", U&)
          RGII% = RG0%
          TXT$ = XXX$
          Call APPFORMDOS("", RGII%, TXT$, "ENCA-DET")
          RG0% = RGII% + 1
          Next U&
          RG0% = RG0% + 1
       End If
     End If
     '
     For U& = PRIDET& To ULTREG&("!CARDEFAC")
        '
        ULTREDET1% = ULTREDET%
        If EXPOR% = 1 Then ULTREDET1% = ULTREDET% - 7
        If RG0% > ULTREDET1% Then                          ' SALTO DE HOJA
           PRIDET& = U&
           TRANSPOR% = 1
           GoTo 7390
        End If
        '
        XXX$ = REGLEIDO$("!CARDEFAC", U&)
        RGII% = RG0%
        '
        TOTITEMS = TOTITEMS + CVS(Left$(XXX$, 4))
        TXT$ = Str$(CVS(Left$(XXX$, 4)))       ' cantidad
        Call APPFORMDOS("", RGII%, TXT$, "CANTIDAD")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = CODEXT$(CVI(Mid$(XXX$, 5, 2)))     ' codigo
        Call APPFORMDOS("", RGII%, TXT$, "COD-MAT")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Mid$(XXX$, 7, 64)                  ' descripcion
        If UCase$(App.EXEName) = "RFAC2001" Then
          If CVI(Mid$(XXX$, 5, 2)) > 0 Then
            TXT$ = DESCRIT$(CVI(Mid$(XXX$, 5, 2)))
          End If
        End If
        Call APPFORMDOS("", RGII%, TXT$, "DES-MAT")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Mid$(XXX$, 71, 4)                  ' unidad de medida
        Call APPFORMDOS("", RGII%, TXT$, "UNIMED")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Mid$(XXX$, 75, 64)                 ' referencia 1
        Call APPFORMDOS("", RGII%, TXT$, "REF-MAT1")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Mid$(XXX$, 139, 64)                ' referencia 2
        Call APPFORMDOS("", RGII%, TXT$, "REF-MAT2")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Mid$(XXX$, 203, 64)                ' referencia 3
        Call APPFORMDOS("", RGII%, TXT$, "REF-MAT3")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Mid$(XXX$, 315, 24)                ' numero de plano
        Call APPFORMDOS("", RGII%, TXT$, "N-PLANO")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Mid$(XXX$, 339, 16)                ' numero original
        Call APPFORMDOS("", RGII%, TXT$, "N-ORIGIN")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Mid$(XXX$, 355, 16)                ' numero de o.compra
        Call APPFORMDOS("", RGII%, TXT$, "N-OCOMPR")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Mid$(XXX$, 387, 16)                ' numero de remito
        Call APPFORMDOS("", RGII%, TXT$, "N-REMITO")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Mid$(XXX$, 403, 16)                ' numero de doc-recepcion
        Call APPFORMDOS("", RGII%, TXT$, "N-BOLREC")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        FEX = CVI(Mid$(XXX$, 419, 2))             ' fecha de recepcion
        TXT$ = FECHATEX$(FEX)
        Call APPFORMDOS("", RGII%, TXT$, "FECH-REC")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
      If COP% >= PRICOFAC% And COP% <= ULTCOFAC% Then
        TXT$ = Str$(FACIVA * CVS(Mid$(XXX$, 421, 4))) ' precio de lista
        Call APPFORMDOS("", RGII%, TXT$, "PRE-LIST")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Str$(FACIVA * CVS(Mid$(XXX$, 429, 4))) ' precio unitario
        Call APPFORMDOS("", RGII%, TXT$, "PRE-NETO")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Str$(CVD(Mid$(XXX$, 433, 8)))   ' total neto
        Call APPFORMDOS("", RGII%, TXT$, "NET-ITEM")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Str$(CVS(Mid$(XXX$, 441, 8)))   ' alicuota iva
        Call APPFORMDOS("", RGII%, TXT$, "ALICUIVA")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Str$(CVD(Mid$(XXX$, 445, 8)))   ' importe iva del item
        Call APPFORMDOS("", RGII%, TXT$, "IVA-ITEM")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Str$(CVD(Mid$(XXX$, 453, 8)))   ' total bruto
        Call APPFORMDOS("", RGII%, TXT$, "BRU-ITEM")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
      End If
      '
      RG0% = UREN% + 1
     Next U&
     '
     RGII% = RG0%
     Call APPFORMDOS("", RGII%, String$(32, "="), "LIN-CIE")
       If RGII% > UREN% Then UREN% = RGII%
     RGII% = RG0%
     Call APPFORMDOS("", RGII%, Str$(TOTITEMS), "TOTA-ITS")
       If RGII% > UREN% Then UREN% = RGII%
     RGII% = RG0%
     If UCase$(App.EXEName) = "RFAC2001" Then
          Call APPFORMDOS("", RGII%, "Total de Items", "LETOT-IT")
        Else
          Call APPFORMDOS("", RGII%, "Total Items", "LETOT-IT")
     End If
     If RGII% > UREN% Then UREN% = RGII%
     '
     If EXPOR% = 1 Then
       If UCase$(App.EXEName) = "RFAC2001" Then GoTo 7390
       '
       UREDTX% = UREN%
       SOLOFAC% = 1
       RGII% = UREDTX%
       Call APPFORMDOS("", RGII%, "Valor FOB " + " .......................................................", "DES-MAT")
       '
       RGII% = UREDTX%
       Call APPFORMDOS("", RGII%, CSTRING$(Mid$(FINFACEX$, 1, 8), 4, 12, 2, 2), "NET-ITEM")
       '
       RGII% = UREDTX% + 1
       Call APPFORMDOS("", RGII%, "FLETE " + TRIM$(Mid$(FINFACEX$, 9, 12)) + " ...................................", "DES-MAT")
       '
       RGII% = UREDTX% + 1
       Call APPFORMDOS("", RGII%, CSTRING$(Mid$(FINFACEX$, 21, 8), 4, 12, 2, 2), "NET-ITEM")
       '
       RGII% = UREDTX% + 3
       Call APPFORMDOS("", RGII%, "Valor " + TRIM$(Mid$(FINFACEX$, 29, 12)) + " ...................................", "DES-MAT")
       '
       RGII% = UREDTX% + 3
       Call APPFORMDOS("", RGII%, CSTRING$(Mid$(FINFACEX$, 41, 8), 4, 12, 2, 2), "NET-ITEM")
       '
       TTX111$ = "Para el presente documento expresado"
       TTX222$ = "en " + TRIM$(Mid$(FINFACEX$, 49, 12)) + " el tipo de cambio considerado"
       TTX333$ = "es 1 " + TRIM$(Mid$(FINFACEX$, 49, 12)) + " = " + TRIM$(CSTRING$(Mid$(FINFACEX$, 61, 8), 4, 12, 4, 2)) + " $."
       '
       RGII% = UREDTX% + 5
       Call APPFORMDOS("", RGII%, TTX111$, "DES-MAT")
       '
       RGII% = UREDTX% + 6
       Call APPFORMDOS("", RGII%, TTX222$, "DES-MAT")
       '
       RGII% = UREDTX% + 7
       Call APPFORMDOS("", RGII%, TTX333$, "DES-MAT")
       '
     End If
     SOLOFAC% = 0
     '
7390 Return
     '
'****************************************************************************
    '
7400 If COP% >= PRICOFAC% And COP% <= ULTCOFAC% Then
         Call SETFORMDOS("", Str$(FACIVA * TOTAFAC#(1, 1)), "IMP-NETO")
         Call SETFORMDOS("", Str$(PORDES), "POR-DESC")
         Call SETFORMDOS("", Str$(FACIVA * TOTAFAC#(1, 4)), "IMP-DESC")
         Call SETFORMDOS("", Str$(FACIVA * TOTAFAC#(1, 5)), "IMPNEDES")
         Call SETFORMDOS("", Str$(TOTAFAC#(1, 6)), "IMP-IVA")
         Call SETFORMDOS("", Str$(TOTAFAC#(1, 7)), "IMP-RNI")
         Call SETFORMDOS("", Str$(TOTAFAC#(1, 8)), "IMP-TOTA")
         Call SETFORMDOS("", Str$(TOTAFAC#(1, 9)), "IMP-VARS")
         Call SETFORMDOS("", Str$(TOTAFAC#(1, 10)), "PERC-IVA")
     End If
     '
     Return
     '
'****************************************************************************
    '
7500 Call CARLEYENDOS("")
     '
     Return
     '
End Sub


Sub FORMMGRDOS(Form$, RG%, CL%, VLX$, CMD$)
   '
   ' FORM$  =  Nombre de formulario
   ' RG%    =  Renglon
   ' CL%    =  Columna
   ' VLX$   =  Valor o texto a colocar
   ' CMD$   =  Comando a ejecutar
   '
   Static FORMACT$, NX%, NY%, ATOT%, CARE%
   Static FXX As String * 128
   Static FYY As String * 256
   Static RESPALTEX%
   '
   If REFRESHFORMDOS% > 0 Then    ' se ejecuto BAJALDISCO
      NX% = 0
      NY% = 0
      If Form$ = "" Then Form$ = FORMACT$
      FORMACT$ = ""
      REFRESHFORMDOS% = 0
   End If
   '
   If Form$ <> "" Then
      If UCase$(TRIM$(Form$)) <> FORMACT$ Then
         FORMACT$ = UCase$(TRIM$(Form$))
         FORMUDRV$ = LUDAT$ + FORMACT$ + ".FRM"
         If NX% > 0 Then Close #NX%
         NX% = FILEOPEN%(FORMUDRV$, 0, 128)
         '
         ATOT% = 256: CARE% = 128       ' ancho total y cantidad de renglones
         For K% = 1 To LOF(NX%) / 128
            Get #NX%, K%, FXX$
            If Left$(FXX$, 8) = "CAN-COL " Then
                  ATOT% = Val(TRIM$(Mid$(FXX$, 89, 8)))
                ElseIf Left$(FXX$, 8) = "CAN-REN " Then
                  CARE% = Val(TRIM$(Mid$(FXX$, 89, 8)))
            End If
         Next K%
      End If
   End If
   '
   If NY% < 1 Then
      LUX$ = LUCOM$
      NY% = FILEOPEN%(LUX$ + "\" + FORMACT$, 0, 256)
   End If
   COMA$ = UCase$(TRIM$(CMD$))
   '
   If Left$(COMA$, 5) = "CLEAR" Then           ' genera formulario en blanco
      LUX$ = LUCOM$
      If NY% > 0 Then Close #NY%
      LUX$ = LUCOM$
      NY% = FILEOPEN%(LUX$ + "\" + FORMACT$, 0, 256)
      '
      For I% = 1 To CARE%
         LSet FYY$ = Space$(256)
         Put #NY%, I%, FYY$
      Next I%
      '
      Exit Sub
      '
   ElseIf Left$(COMA$, 5) = "PRINT" Then
      Screen.MousePointer = 11
      '
 ' para imprimir directamente sin pasar por DOS - EPB 23/MAYO/2005
 '
 If PRINTERPORT$ <> "" Then
   Call SETPRINPORT(PRINTERPORT$)
 End If
 Printer.Orientation = 1
 Printer.Font = "Draft 17cpi"
 Printer.FontBold = False
 Printer.FontItalic = False
 Printer.FontUnderline = False
 Printer.FontName = "Draft 17cpi"
 '
 CAREMAX% = 0
 For NNN444% = 1 To CARE%
   Get #NY%, NNN444%, FYY$
   TTT555$ = FYY$
   If TRIM$(TTT555$) <> "" Then CAREMAX% = NNN444%
 Next NNN444%
 '
 INTERLI = 56.7 * 25.4 * 12 / 72
 For NNN444% = 1 To CAREMAX%
   Get #NY%, NNN444%, FYY$
   TTT555$ = FYY$
   Printer.CurrentY = (NNN444% - 1) * INTERLI
   Printer.Font = "Draft 17cpi"
   Printer.FontBold = False
   Printer.FontItalic = False
   Printer.FontUnderline = False
   Printer.FontName = "Draft 17cpi"
   Printer.Print TTT555$
 Next NNN444%
 Printer.EndDoc
 '
 If PRINTERPORT$ <> "" Then
   Call SETPRINPORT("RESTORE")
 End If
 '
 Exit Sub
' FIN EPB 23 MAYO 2005
'
      If RESPALTEX% = 0 Then
        RESPALTEX% = (-1)
        SVLC$ = SERIVOL$("C:")
        If SVLC$ = "081D-0BD9" Or SVLC$ = "2560-10F0" Then ' NOTEBOOK
          RESPALTEX% = 1
        End If
      End If
      '
      If RESPALTEX% = 1 Then
         TTT111& = 10 * Timer
         TTT222$ = TRIM$(Str$(TTT111&))
         NNN333% = FreeFile
         Open "C:\FLEXOFT\" + TTT222$ + ".TXT" For Output As #NNN333%
         For NNN444% = 1 To CARE%
           Get #NY%, NNN444%, FYY$
           TTT555$ = FYY$
           Print #NNN333%, TTT555$
         Next NNN444%
         Close #NNN333%
      End If
      '
      UTI$ = LUBAS$
      If Mid$(UTI$, 2, 1) <> ":" Then
          UTI$ = Left$(CurDir, 2) + UTI$
      End If
      LUX$ = LUCOM$
      ATRIBU$ = "/" + LUX$ + "\" + FORMACT$
      ATRIBU$ = ATRIBU$ + "/CAREN=" + TRIM$(Str$(CARE%))
      ATRIBU$ = ATRIBU$ + "/SETUP=" + SETUPSTRI$
      ATRIBU$ = ATRIBU$ + "/MARSU=" + TRIM$(Str$(MARSUP%))
      '
      NZ% = FILEOPEN%("C:\FLEXOFT\COMBAT.$$$", 2, 0)
      Print #NZ%, "FINBAT"
      Close #NZ%
      '
      On Error Resume Next
      Kill "C:\FLEXOFT\FINBAT.$$$"
      On Error GoTo 0
      '
      NZ% = FILEOPEN%("C:\FLEXOFT\DOSCOMM.BAT", 2, 0)
      Print #NZ%, UTI$ + "PRINTFRM.EXE " + ATRIBU$
      Close #NZ%
      '
      Shell "C:\FLEXOFT\DOSCOMM.PIF"
      TINI = Timer
      MENSA$ = "Se ha producido un error no recuperable" + Chr$(10) + Chr$(13)
      MENSA$ = MENSA$ + "en la impresion de este documento." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
      MENSA$ = MENSA$ + "Consulte a su Soporte de Sistemas FLEXOFT"
      '
149   'OPENFORMS = DoEvents
      If Timer > TINI + 30 Then
          Screen.MousePointer = 1
          MsgBox MENSA$, 4096 + 48, "Control de Impresion"
          On Error Resume Next
          Kill "C:\FLEXOFT\COMBAT.$$$"
          Kill "C:\FLEXOFT\FINBAT.$$$"
          On Error GoTo 0
          ABOPRIN% = 1
          Exit Sub
      End If
      '
      If EXISTE%("C:\FLEXOFT\FINBAT.$$$") < 1 Then GoTo 149
      Screen.MousePointer = 1
      ABOPRIN% = 0
      Exit Sub
      '
   ElseIf Left$(COMA$, 4) = "FSET" Then
      If RG% > 0 Then
        If RG% <= CARE% Then
          If CL% > 0 Then
            If CL% <= ATOT% Then
              Get #NY%, RG%, FYY$
              TXTT$ = FYY$
              Mid$(TXTT$, CL%, Len(VLX$)) = VLX$
              LSet FYY$ = TXTT$
              Put #NY%, RG%, FYY$
            End If
          End If
        End If
      End If
      Exit Sub
      '
   ElseIf Left$(COMA$, 5) = "LEYEN" Then
      For K% = 1 To LOF(NX%) / 128
        Get #NX%, K%, FXX$
        If Left$(FXX$, 4) = "LEYE" Then
          VLX$ = TRIM$(Mid$(FXX$, 9, 80))
          If VLX$ <> "" Then
            For TTI% = 97 To 128 Step 8
              RG% = Asc(Mid$(FXX$, TTI%, 1))
              If RG% > 0 Then
                If RG% <= CARE% Then
                  CL% = Asc(Mid$(FXX$, TTI% + 1, 1))
                  If CL% > 0 Then
                    If CL% <= ATOT% Then
                      FORMATO$ = TRIM$(Mid$(FXX$, TTI% + 2, 6))
                      If FORMATO$ <> "" Then
                        TX$ = TEXTFORM$(VLX$, FORMATO$)
                        Get #NY%, RG%, FYY$
                        TXTT$ = FYY$
                        Mid$(TXTT$, CL%, Len(TX$)) = TX$
                        LSet FYY$ = TXTT$
                        Put #NY%, RG%, FYY$
                      End If
                    End If
                  End If
                End If
              End If
            Next TTI%
          End If
        End If
      Next K%
      '
      Exit Sub
      '
   ElseIf Left$(COMA$, 1) = "@" Then
      CD$ = AJUSTI$(Mid$(COMA$, 2), 8)
      For K% = 1 To LOF(NX%) / 128
        Get #NX%, K%, FXX$
        If Left$(FXX$, 8) = CD$ Then
          RG0% = RG%
          UREGUT% = 0
          For TTI% = 97 To 128 Step 8
            RG% = RG0%
            RGX% = Asc(Mid$(FXX$, TTI%, 1))
            If RG% + RGX% > 0 Then
              If RG% + RGX% <= CARE% Then
                CL% = Asc(Mid$(FXX$, TTI% + 1, 1))
                If CL% > 0 Then
                  If CL% <= ATOT% Then
                    FORMATO$ = TRIM$(Mid$(FXX$, TTI% + 2, 6))
                    If FORMATO$ <> "" Then
                      ILETRA% = 0
                      '
                      'suprime linea de cierre bajo importe en refacturacion
                      If UCase$(App.EXEName) = "RFAC2001" Then
                        If TTI% = 105 Then
                          If TRIM$(CD$) = "LIN-CIE" Then
                            GoTo 1009
                          End If
                        End If
                      End If
                      '
                      If CD$ = "IMP-LET1" Then
                        ILETRA% = 1
                        LOTEX% = Val(Mid$(FORMATO$, 2, 4))
                        Call FRATEX(VLX$, NRX1$, NRX2$, NRX3$, LOTEX%)
                        VLX$ = NRX1$ + String$(LOTEX%, "*")
                        RL% = 1
                      End If
                      '
1000                  TX$ = TEXTFORM$(VLX$, FORMATO$)
                      RG% = RG% + RGX%
                      Get #NY%, RG%, FYY$
                      TXTT$ = FYY$
                      Mid$(TXTT$, CL%, Len(TX$)) = TX$
                      LSet FYY$ = TXTT$
                      Put #NY%, RG%, FYY$
                      If RG% > UREGUT% Then UREGUT% = RG%
                      '
                      If ILETRA% = 1 Then
                        If RL% = 1 Then
                          RL% = 2
                          VLX$ = NRX2$ + String$(LOTEX%, "*")
                          RGX% = 1
                          GoTo 1000
                        ElseIf RL% = 2 Then
                          RL% = 3
                          VLX$ = NRX3$ + String$(LOTEX%, "*")
                          RGX% = 1
                          GoTo 1000
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
            If SOLOFAC% = 1 Then GoTo 1010
1009      Next TTI%
1010      If UREGUT% > RG% Then RG% = UREGUT%
          Exit Sub
        End If
      Next K%
      Exit Sub
      '
   ElseIf Left$(COMA$, 1) = "?" Then
      CD$ = AJUSTI$(Mid$(COMA$, 2), 8)
      If NX% > 0 Then
        For K% = 1 To LOF(NX%) / 128
          Get #NX%, K%, FXX$
          If Left$(FXX$, 8) = CD$ Then
            VLX$ = TRIM$(Mid$(FXX$, 89, 8))
            Exit Sub
          End If
        Next K%
      End If
      VLX$ = ""
      Exit Sub
      '
   Else
      '
      Call COMUNI("Comando '" + COMA$ + "' no Procesable\en Rutina 'FORMMGR'\para Formulario '" + FORMACT$ + "'.")
      '
   End If
   '
End Sub
   '
'*****************************************************************************
   '
Sub SETFORMDOS(Form$, VLX$, CTRX$)
   '
   CMD$ = "@" + TRIM$(CTRX$)
   RGII% = 0: CLII% = 0
   Call FORMMGRDOS(Form$, RGII%, CLII%, VLX$, CMD$)
   ' sobre RG% y CL% devuelve renglon y columna utilizado
   '
End Sub
'
'*****************************************************************************
'
Sub CARLEYENDOS(Form$)
   '
   CMD$ = "LEYEN"
   RGII% = 0: CLII% = 0
   Call FORMMGRDOS(Form$, RGII%, CLII%, VLX$, CMD$)
   '
End Sub
'
'*****************************************************************************
'
Sub APPFORMDOS(Form$, RGII%, VLX$, CTRX$)
   '
   CMD$ = "@" + TRIM$(CTRX$)
   If TRIM$(VLX$) <> "" Then
       Call FORMMGRDOS(Form$, RGII%, CLII%, VLX$, CMD$) ' sobre RG% y CL% devuelve
   End If                                    '    renglon y columna utilizado
   '
End Sub
'
'*****************************************************************************
'
Sub CLEARFORMDOS(Form$)
   '
   Call FORMMGRDOS(Form$, RGII%, CLII%, VLX$, "CLEAR")
   '
End Sub
'
'*****************************************************************************
'
Sub PRINTFORMDOS(Form$)
   '
   Call FORMMGRDOS(Form$, RGII%, CLII%, VLX$, "PRINT")
   '
End Sub
'
'*****************************************************************************
'
Function ATRIFORMDOS$(Form$, CTRX$)
   '
   CMD$ = "?" + TRIM$(CTRX$)
   Call FORMMGRDOS(Form$, RGII%, CLII%, VLX$, CMD$)
   ATRIFORMDOS$ = VLX$
   '
End Function

Sub GRAMOVEN(NROFAC, REMI)
   '
   UREFA& = ULTREG&("FACTUR")
   '
   NC1% = Val(Text1)
   If NC1% < 1 Then NC1% = 10001
   REFE$ = ECOARCH$("DEUDOR1", MKI$(NC1%))
   FF% = FECHANUM(Text9)
   FEXPE$ = TRIM$(Text9)
   ANIOX% = Val(Mid$(FEXPE$, 7, 2))
   If ANIOX% < 80 Then ANIOX% = 100 + ANIOX%
   MESOX% = Val(Mid$(FEXPE$, 4, 2))
   CODIMEXPE% = 12 * ANIOX% + MESOX%
   '
   RTOX$ = TRIM$(Str$(Abs(REMI)))
   While Len(RTOX$) < 8: RTOX$ = "0" + RTOX$: Wend
   PREFI$ = ATRIFORM$(FORIMPRE$, "PREFIFAC")
   RTOX$ = PREFI$ + "-" + RTOX$
   '
   FACX$ = TRIM$(Str$(Abs(NROFAC)))
   While Len(FACX$) < 8: FACX$ = "0" + FACX$: Wend
   PREFI$ = ATRIFORM$(CODFOR$, "PREFIFAC")
   FACX$ = PREFI$ + "-" + FACX$
   '
   HOII% = HOY(HOS$)
   '
   SIG% = (-1)
   If NROFAC < 0 Then SIG% = 1
   '
   COEF = 1
   If MODESCU% = 1 Then
     DESC$ = TRIM$(PORDESCU)
     Call DESCUEN(PORDES, DESC$)
     If PORDES > 0 Then COEF = 1 - PORDES / 100
   End If
   '
   FIN& = ULTREG&("!CARDEFAC")
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     X$ = REGLEIDO$("!CARDEFAC", I&)
     CANTAA = CVS(Left$(X$, 4))
     If CANTAA < 0.05 Then
       If CVD(Mid$(X$, 433, 8)) > 0 Then
         CANTAA = 1
       End If
     End If
     '
     If CANTAA > 0 Then
       '
       CI% = CVI(Mid$(X$, 5, 2))
       COD$ = CODEXT$(CI%)
       PREUNIBRU# = CDbl(CVS(Mid$(X$, 421, 4)))
       If PREUNIBRU# < 0.00005 Then
         PREUNIBRU# = CVD(Mid$(X$, 433, 8)) / CANTAA
       End If
       PRE# = CDbl(CVS(Mid$(X$, 429, 4)))
       If PRE# < 0.00005 Then
         PRE# = CVD(Mid$(X$, 433, 8)) / CANTAA
       End If
       TASAIVA = CVS(Mid$(X$, 441, 4))
       LOTE$ = TRIM$(Mid$(X$, 267, 16))
       CMOV$ = "RT"
       DOPRI$ = "RT." + TRIM$(Str$(REMI))
       DOSEC$ = "FC." + TRIM$(Str$(NROFAC))
       '
       REODES& = CVS(Mid$(X$, 505, 4))
       REPEDA& = CVS(Mid$(X$, 509, 4))
       If REODES& < 1 Then
         '
         If REBLAMAVE$ = "" Then
           REBLAMAVE$ = REGBLAN$("MAVENDID")
         End If
         '
         Y$ = REBLAMAVE$
         Call REPLA(Y$, MKI$(NC1%), 1, 2)
         Call REPLA(Y$, Chr$(0), 3, 1) ' codigo de empresa
         Call REPLA(Y$, Chr$(0), 4, 1) ' CODIGO DE SUCURSAL
         Call REPLA(Y$, MKI$(CI%), 5, 2)
         Call REPLA(Y$, AJUSTI$(KCOLOR$, 16), 7, 16)
         Call REPLA(Y$, AJUSTI$(KTALLE$, 8), 25, 8)
         Call REPLA(Y$, MKS$(KLARGO), 33, 4)
         Call REPLA(Y$, MKS$(KANCHO), 37, 4)
         Call REPLA(Y$, MKS$(KESPES), 41, 4)
         Call REPLA(Y$, String$(4, 0), 45, 4)  ' FILLER
         UNI$ = UNIMED$(CI%)
         Call REPLA(Y$, UNI$, 49, 4)
         Call REPLA(Y$, UNI$, 53, 4)
         Call REPLA(Y$, MKS$(1), 57, 4)  ' COEFICIENTE CONVERSION UNIDADES
         Call REPLA(Y$, "FC", 61, 2)     ' MOVIMIENTO DE ORIGEN
         Call REPLA(Y$, Chr$(3), 63, 1)  ' STATUS FACTURADO
         Call REPLA(Y$, Chr$(0), 64, 1)  ' FILLER
         Call REPLA(Y$, MKI$(0), 65, 2)  ' CLIENTE
         Call REPLA(Y$, MKI$(0), 67, 2)  ' CODIGO
         Call REPLA(Y$, String$(12, 0), 69, 12) ' FILLER
         '
         Call REPLA(Y$, Space$(16), 105, 16)
         PRECOVE$ = Chr$(TEXT21(0)) + Chr$(TEXT21(1)) + Chr$(TEXT21(2))
         Call REPLA(Y$, PRECOVE$ + Chr$(0), 121, 4)
         Call REPLA(Y$, MKI$(FENTRESO%), 125, 2) ' FECHA O/DESPACHO
         Call REPLA(Y$, MKD$(PREUNIBRU#), 127, 8) ' PRECIO UNITARIO BRUTO
         Call REPLA(Y$, MKS$(0), 135, 4) ' PORCIENTO DESCUENTO item
         Call REPLA(Y$, MKD$(PRE#), 139, 8) ' PRECIO UNITARIO NETO
         Call REPLA(Y$, MKS$(0), 147, 4) ' PORCIENTO DESCUENTO general pedido
         Call REPLA(Y$, MKD$(PRE#), 151, 8) ' PRECIO UNITARIO NETO
         Call REPLA(Y$, Chr$(0), 159, 1)       ' MONEDA = $
         Call REPLA(Y$, Chr$(0), 160, 1)       ' FILLER
         Call REPLA(Y$, MKD$(1), 161, 8)       ' TIPO DE CAMBIO
         Call REPLA(Y$, MKD$(PRE#), 169, 8) ' valor unitario neto-neto EN PESOS
         Call REPLA(Y$, MKS$(-1 * CANTAA * SIG%), 177, 4) ' CANTIDAD ASIGNADA
         Call REPLA(Y$, MKS$(0), 181, 4)  ' POCEN.COMIS.VENDEDOR
         Call REPLA(Y$, String$(2, 0), 185, 2) ' FILLER
         Call REPLA(Y$, String$(6, 0), 187, 6) ' ANULACION (F/H/U/T)
         '
         Call REGAPP("MAVENDID", Y$)
         '
         REODES& = ULTREG&("MAVENDID")
         Call REPLA(X$, MKS$(REODES&), 505, 4)
         Call GRAREG("!CARDEFAC", X$, I&)
         '
       End If
       '
       If NROFAC < 0 Then
          CMOV$ = "RD"
          DOPRI$ = "DV." + TRIM$(Str$(Abs(REMI)))
          DOSEC$ = "NC." + TRIM$(Str$(Abs(NROFAC)))
       End If
       '
       If REMI < 0 Then
          CMOV$ = "ND"
          DOPRI$ = "ND." + TRIM$(Str$(Abs(REMI)))
          DOSEC$ = "ND." + TRIM$(Str$(Abs(NROFAC)))
       End If
       '
       Call MOVISTO(FF%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, PRE# * COEF, "$", NC1%, 1, SIG% * CANTAA)
       REMOVI& = ULTREG&("MOVISTO")
       '
       Y$ = REGLEIDO$("MAVENDID", REODES&)
       Call REPLA(Y$, Chr$(3), 63, 1)     ' MARCA FACTURADO
       Call REPLA(Y$, MKI$(0), 65, 2)
       Call REPLA(Y$, MKI$(-1), 67, 2)    ' MARCA PENDIENTE VERIFICACION
       '
       Call REPLA(Y$, MKI$(HOII%), 193, 2)
       Call REPLA(Y$, MKI$(HORANUM%(Time$)), 195, 2)
       Call REPLA(Y$, Chr$(USNBR% Mod 100), 197, 1)
       Call REPLA(Y$, MKI$(FF%), 199, 2)
       Call REPLA(Y$, MKI$(CODIMEXPE%), 201, 2)
       Call REPLA(Y$, MKS$(Abs(REMI)), 203, 2)
       Call REPLA(Y$, Left$(DOPRI$, 2) + "-" + RTOX$, 209, 16)
       Call REPLA(Y$, MKS$((-1) * CANTAA * SIG%), 225, 4)
       Call REPLA(Y$, LOTE$, 229, 16)
       Call REPLA(Y$, Chr$(1), 245, 1)
       Call REPLA(Y$, Chr$(3), 246, 1)
       Call REPLA(Y$, MKS$(REMOVI&), 247, 4)
       Call REPLA(Y$, String$(6, 0), 251, 6) ' ANULACION (F/H/U/T)
       '
       Call REPLA(Y$, MKI$(HOII%), 257, 2)
       Call REPLA(Y$, MKI$(HORANUM%(Time$)), 259, 2)
       Call REPLA(Y$, Chr$(USNBR% Mod 100), 261, 1)
       Call REPLA(Y$, MKI$(FF%), 263, 2)
       Call REPLA(Y$, MKI$(CODIMEXPE%), 265, 2)
       Call REPLA(Y$, MKS$(Abs(NROFAC)), 267, 4)
       Call REPLA(Y$, Left$(DOSEC$, 2) + "-" + CATIVA$ + "-" + FACX$, 273, 24)
       Call REPLA(Y$, MKS$((-1) * CANTAA * SIG%), 297, 4)
       Call REPLA(Y$, MKD$(PRE# * COEF), 301, 8) ' valor unitario neto-neto EN PESOS
       Call REPLA(Y$, Chr$(0), 309, 1)       ' MONEDA = $
       Call REPLA(Y$, Chr$(0), 310, 1)       ' FILLER
       Call REPLA(Y$, MKD$(1), 311, 8)       ' TIPO DE CAMBIO
       Call REPLA(Y$, MKD$(PRE# * COEF), 319, 8) ' valor unitario neto-neto EN PESOS
       Call REPLA(Y$, MKS$(TASAIVA), 343, 4) ' TASA PORCIENTO IVA
       Call REPLA(Y$, MKS$(UREFA&), 347, 4) ' PUNTERO A FACTURACION
       '
       Call GRAREG("MAVENDID", Y$, REODES&)
       If REPEDA& < 1 Then
         REPEDA& = CVS(Mid$(Y$, 99, 4))
       End If
       '
       If SIG% < 0 Then
         If REPEDA& > 0 Then
           Y$ = REGLEIDO$("PEDDETA", REPEDA&)
           Call REPLA(Y$, MKS$(0), 75, 4)
           Call REPLA(Y$, MKS$(CVS(Mid$(Y$, 95, 4)) + CANTAA), 95, 4)
           SALPENOD = CVS(Mid$(Y$, 79, 4)) - CVS(Mid$(Y$, 95, 4))
           If SALPENOD < 0.05 Then SALPENOD = 0
           Call REPLA(Y$, MKS$(SALPENOD), 99, 4)
           STATU% = 1: If SALPENOD < 0.05 Then STATU% = 3
           Call REPLA(Y$, Chr$(STATU%), 27, 1)
           Call REPLA(Y$, MKI$(FF%), 103, 2)
           Call REPLA(Y$, MKS$(Abs(REMI)), 105, 4)
           Call GRAREG("PEDDETA", Y$, REPEDA&)
           '
           REPEDP& = REGBUS&("PEDDETP", 28, MKS$(REPEDA&))
           If REPEDP& > 0 Then
             Call REPLA(Y$, MKS$(REPEDA&), 109, 4)
             Call GRAREG("PEDDETP", Y$, REPEDP&)
           End If
           '
         End If
       End If
     End If
   Next I&
   '
   Call CIERRARCH("MAVENDID")
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("PEDDETA")
   Call CIERRARCH("PEDDETP")
   '
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     X$ = REGLEIDO$("!CARDEFAC", I&)
     CANTAA = CVS(Left$(X$, 4))
     If CANTAA > 0 Then
       '
       CI% = CVI(Mid$(X$, 5, 2))
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           RXX$ = RECFILT$("DESPAIMP", 1, MKI$(CI%))
           CANTU = CANTAA * SIG% * (-1)
           '
           For VV% = 1 To Len(RXX$) Step 4
             V& = CVS(Mid$(RXX$, VV%, 4))
             If V& > 0 Then
               Y$ = REGLEIDO$("DESPAIMP", V&)
               If CVI(Left$(Y$, 2)) = CI% Then
                 CAPEN# = CVD(Mid$(Y$, 41, 8))
                 If CAPEN# >= CANTU Then
                     CAPEN# = CAPEN# - CDbl(CANTU)
                     Call REPLA(Y$, MKD$(CAPEN#), 41, 8)
                     Call GRAREG("DESPAIMP", Y$, V&)
                     CANTU = 0
                     GoTo 98
                   Else
                     CANTU = CANTU - CAPEN#
                     CAPEN# = 0
                     Call REPLA(Y$, MKD$(0), 41, 8)
                     Call GRAREG("DESPAIMP", Y$, V&)
                 End If
               End If
             End If
           Next VV%
           '
         End If
       End If
     End If
     '
98 Next I&
   '
End Sub
'
Sub ANUFAC()
    '
    If DERACCE%("ANU-FAC", "Anulacion Facturas-Debitos-Creditos") < 1 Then
        GoTo 999
    End If
    '
    Call CIERRARCH("*.*")
    Call FRESHALL
    '
    OPX% = ALTERNA%("Comprobantes Emitidos\Facturas 2848/49/50 del 17/01/06")
    If OPX% < 1 Or OPX% > 2 Then
      Exit Sub
    End If
    If OPX% = 2 Then
      For KU& = 2848 To 2850
        TTXX$ = String$(64, 0)
        FEXI% = FECHANUM("17/01/06")
        Call REPLA(TTXX$, MKI$(5090), 1, 2)  ' CLIENTE
        Call REPLA(TTXX$, MKS$(KU&), 3, 4) ' NUMERO
        Call REPLA(TTXX$, MKI$(FEXI%), 7, 2)
        Call REPLA(TTXX$, Chr$(2), 13, 1)  ' FACTURA
        Call REPLA(TTXX$, Chr$(0), 14, 1)
        '
        Screen.MousePointer = 11
        TTYY$ = Mid$(TTXX$, 3, 12)
        FIN& = ULTREG&("FACTUR")
        For U& = 1 To FIN&
          Call TRACE(20, U&, FIN&)
          XX$ = REGLEIDO$("FACTUR", U&)
          If Mid$(XX$, 3, 12) = TTYY$ Then
            Call MENSERR(24, "Documento " + TRIM$(Str$(KU&)) + " ya Registrado")
            GoTo 3
          End If
        Next U&
        '
        Call REGAPP("FACTUR", TTXX$)
        Call CIERRARCH("FACTUR")
3       Screen.MousePointer = 1
      Next KU&
      Exit Sub
    End If
    '
    '
    Dim TIS$(10)
    TDAT$ = "AS,VC,FC,ND,RB,DC,NC,AS"
    For II% = 0 To 7
        TIS$(II%) = Mid$(TDAT$, 3 * II% + 1, 2)
    Next II%
    '
    Call CARSELMESA
5   Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
        GoTo 999
    End If
    '
    MMXX0$ = TRIM$(MESASEL$)
    MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
    '
    FECHALIM% = Val(CONTROL$("", "CIEMEVEN"))
    If HAS% <= FECHALIM% Then
       Call MENSERR(24, "Período no Válido.\  \Ese Período Mensual Está Cerrado.")
       GoTo 5
    End If
    '
10  Screen.MousePointer = 11
    Call FRESHECHO("!FACSEL1")
    Call BLANARCH("!FACSEL1")
    FIN& = ULTREG&("FACTUR")
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        AX$ = REGLEIDO$("FACTUR", I&)
        If CVI(Mid$(AX$, 7, 2)) >= DES% Then
            If CVI(Mid$(AX$, 7, 2)) <= HAS% Then
                NC% = CVI(Left$(AX$, 2))
                CLI$ = AJUSTI$(RASOCLI$(NC%), 30)
                NROL& = CVS(Mid$(AX$, 3, 4))
                FEX = CVI(Mid$(AX$, 7, 2))
                TI% = Asc(Mid$(AX$, 13, 1)): If TI% < 0 Or TI% > 7 Then TI% = 2
                TIX$ = TIS$(TI%)
                VA% = Asc(Mid$(AX$, 14, 1))
                VAX$ = "A": If VA% < 1 Or VA% > 2 Then VAX$ = "X"
                            If VA% = 2 Then VAX$ = "B"
                IMPO# = CVD(Mid$(AX$, 15, 8))
                IMPO# = IMPO# + CVD(Mid$(AX$, 23, 8))
                IMPO# = IMPO# + CVD(Mid$(AX$, 31, 8))
                IMPO# = IMPO# + CVD(Mid$(AX$, 49, 8))
                IMPO# = IMPO# + CVD(Mid$(AX$, 57, 8))
                SIG% = 1: If TI% > 3 Then SIG% = (-1)
                '
                If Abs(IMPO#) >= 0.005 Then
                    X$ = MKS$(I&) + TIX$ + "." + TRIM$(Str$(NROL&)) + "-" + VAX$
                    X$ = X$ + " - " + FECHATEX$(FEX)
                    X$ = X$ + " - " + TRIM$(CSTRING$(MKD$(Abs(IMPO#)), 3, 12, 2, 2))
                    X$ = X$ + " - " + CLI$
                    Call REGAPP("!FACSEL1", X$)
                End If
                '
            End If
        End If
    Next I&
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
20  Call SELECHO("!FACSEL1")
    X$ = VALACT1$("!FACSEL1")
    If X$ = "" Then
        GoTo 5
    End If
    REFA& = Val(X$)
    If REFA& < 1 Or REFA& > ULTREG&("FACTUR") Then
        GoTo 999
    End If
    '
    AX$ = REGLEIDO$("FACTUR", REFA&)
    NC% = CVI(Left$(AX$, 2))
    NROL& = CVS(Mid$(AX$, 3, 4))
    FF% = CVI(Mid$(AX$, 7, 2))
    TI% = Asc(Mid$(AX$, 13, 1)): If TI% < 0 Or TI% > 7 Then TI% = 2
    TIX$ = TIS$(TI%)
    VA% = Asc(Mid$(AX$, 14, 1))
    VAX$ = "A": If VA% < 1 Or VA% > 2 Then VAX$ = "X"
                If VA% = 2 Then VAX$ = "B"
    NETO# = CVD(Mid$(AX$, 15, 8))
    IMPO# = CVD(Mid$(AX$, 15, 8))
    IMPO# = IMPO# + CVD(Mid$(AX$, 23, 8))
    IMPO# = IMPO# + CVD(Mid$(AX$, 31, 8))
    IMPO# = IMPO# + CVD(Mid$(AX$, 49, 8))
    IMPO# = IMPO# + CVD(Mid$(AX$, 57, 8))
    SIG% = 1: If TI% > 3 Then SIG% = (-1)
    NETO# = NETO# * SIG%: IMPO# = IMPO# * SIG%
    '
    VDEF$ = String$(40, 0)
    Call REPLA(VDEF$, TIX$ + "." + TRIM$(Str$(NROL&)) + "-" + VAX$, 1, 18)
    Call REPLA(VDEF$, MKI$(NC%), 19, 2)
    Call REPLA(VDEF$, MKI$(FF%), 21, 2)
    Call REPLA(VDEF$, MKD$(NETO#), 23, 8)
    Call REPLA(VDEF$, MKD$(IMPO#), 31, 8)
    '
    OBX$ = OBJPLA$("MUESTRAFAC", VDEF$)
    If OBX$ = "" Then GoTo 20
    '
    Screen.MousePointer = 11
    DOCANU$ = TIX$ + "." + TRIM$(Str$(NROL&))
    FIN& = ULTREG&("SVD202")
    '
    SUMADEU# = 0
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        X$ = REGLEIDO$("SVD202", I&)
        If CVI(Left$(X$, 2)) = NC% Then
            If CVS(Mid$(X$, 5, 4)) = NROL& Then
                If Asc(Mid$(X$, 3, 1)) = TI% Then
                    SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
                    SALDO# = CVD(Mid$(X$, 11, 8))
                    IORIG# = CVD(Mid$(X$, 21, 8))
                    If TI% <> 0 Then SALDO# = Abs(SALDO#)
                    SALDO# = SG% * SALDO#
                    SUMADEU# = SUMADEU# + SALDO#
                End If
            End If
        End If
    Next I&
    '
    If Abs(IMPO# - SUMADEU#) > 0.005 Then
        Call COMUNI("La Anulacion de este Documento\no es Posible por Registrar\Cancelaciones Totales o Parciales.\  \Emita Notas de Debito, Credito o\Ajustes de Saldo según Corresponda.")
        GoTo 20
    End If
    '
    AX$ = REGLEIDO$("FACTUR", REFA&)
    Call REPLA(AX$, String$(50, 0), 15, 50)
    Call GRAREG("FACTUR", AX$, REFA&)
    Call CIERRARCH("FACTUR")
    '
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        X$ = REGLEIDO$("SVD202", I&)
        If CVI(Left$(X$, 2)) = NC% Then
            If CVS(Mid$(X$, 5, 4)) = NROL& Then
                If Asc(Mid$(X$, 3, 1)) = TI% Then
                    Call REPLA(X$, MKD$(0), 11, 8)
                    Call REPLA(X$, MKD$(0), 21, 8)
                    Call GRAREG("SVD202", X$, I&)
                End If
            End If
        End If
    Next I&
    '
    RGXX& = REGICLI&(NC%)
    If RGXX& > 0 Then
        If RGXX& <= ULTREG&("DEUDOR2") Then
            TTXS$ = REGLEIDO$("DEUDOR2", RGXX&)
            SALNUE# = (Int(100 * CVD(Mid$(TTXS$, 109, 8)) + 0.5)) / 100 - (Int(100 * IMPO# + 0.5)) / 100
            Call REPLA(TTXS$, MKD$(SALNUE#), 109, 8)
            Call GRAREG("DEUDOR2", TTXS$, RGXX&)
        End If
    End If
    Call CIERRARCH("DEUDOR2")
    '
    FIN& = ULTREG&("MOVISTO")
    '
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        X$ = REGLEIDO$("MOVISTO", I&)
        If TRIM$(Mid$(X$, 33, 10)) = DOCANU$ Then
            If CVI(Left$(X$, 2)) = FF% Then
                If CVI(Mid$(X$, 84, 2)) = NC% Then
                    CI% = CVI(Mid$(X$, 3, 2))
                    LOTE$ = Mid$(X$, 5, 16)
                    CMOV$ = Mid$(X$, 21, 2)
                    DOPRI$ = Mid$(X$, 23, 10)
                    DOSEC$ = DOCANU$
                    REFE$ = "Anulacion " + DOCANU$
                    VUNI# = CVD(Mid$(X$, 73, 8))
                    MONE$ = Mid$(X$, 81, 2)
                    DEPO% = Asc(Mid$(X$, 83, 1))
                    CANTI = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                    Call MOVISTO(FF%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, CANTI)
                End If
            End If
        End If
    Next I&
    Call CIERRARCH("MOVISTO")
    '
    FIN& = ULTREG&("MAVENDID")
    HOII% = HOY(HOS$)
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        X$ = REGLEIDO$("MAVENDID", I&)
        If CVI(Left(X$, 2)) = NC% Then
          If CVI(Mid$(X$, 263, 2)) = FF% Then
            If TRIM$(Mid$(X$, 273, 2)) = TIX$ Then
              If CVS(Mid$(X$, 267, 4)) = NROL& Then
                Call REPLA(X$, Chr$(9), 246, 1) ' MARCA ANULADO
                Call REPLA(X$, MKI$(HOII%), 251, 2)
                Call REPLA(X$, MKI$(HORANUM%(Time$)), 253, 2)
                Call REPLA(X$, Chr$(USNBR% Mod 100), 255, 1)
                Call GRAREG("MAVENDID", X$, I&)
              End If
            End If
          End If
        End If
    Next I&
    Call CIERRARCH("MAVENDID")
    Call CIERRARCH("*.*")
    '
    Call COMUNI("Anulación Completa")
    GoTo 10
    '
999 Call CIERRARCH("*.*")
    If UCase$(App.EXEName) = "FACTUMAI" Then
       OPFACMAI.Option1.Value = True
       OPFACMAI.Option1.Refresh
    End If
    '
End Sub
'
Sub VERIMOVI()
   '
   Screen.MousePointer = 11
   HOII% = HOY(HOS$)
   UREMOVI& = ULTREG&("MOVISTO")
   REBLAMO$ = REGBLAN$("MOVISTO")
   FIN& = ULTREG&("MAVENDID")
   '
   If EXISTE%(LUDAT$ + "MAVENDID.X20") < 1 Then
     For U& = 1 To FIN&
       XX$ = REGLEIDO$("MAVENDID", U&)
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("MAVENDID", U&)
       If CVS(Mid$(XX$, 203, 4)) > 0 Or TRIM$(Mid$(XX$, 209, 16)) <> "" Then
         If CVI(Mid$(XX$, 67, 2)) = 0 Then
           Call REPLA(XX$, MKI$(-1), 67, 2)
           Call GRAREG("MAVENDID", XX$, U&)
         End If
       End If
     Next U&
     Call CIERRARCH("MAVENDID")
   End If
   '
   RFM$ = RECFILT$("MAVENDID", 20, MKI$(-1))
   '
   For UK& = 1 To Len(RFM$) Step 4
     U& = CVS(Mid$(RFM$, UK&, 4))
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("MAVENDID", U&)
     REMOVI& = CVS(Mid$(XX$, 247, 4))
     '
     If REMOVI& <= 0 Then
       If CVI(Mid$(XX$, 67, 2)) = (-1) Then
         Call REPLA(XX$, MKI$(0), 67, 2)
         Call GRAREG("MAVENDID", XX$, U&)
       End If
     End If
     '
     If REMOVI& > 0 Then
       FEMOX% = CVI(Mid$(XX$, 199, 2))
       CIPIX% = CVI(Mid$(XX$, 5, 2))
       LOTEX$ = Mid$(XX$, 229, 16)
       CMOVX$ = Mid$(XX$, 209, 2)
       DOPRI$ = CMOVX$ + "." + TRIM$(Str$(Val(Mid$(XX$, 217, 8))))
       If CMOVX$ = "DV" Then CMOVX$ = "RD"
       DOSEC$ = Mid$(XX$, 273, 2) + "." + TRIM$(Str$(Val(Mid$(XX$, 283, 8))))
       NCX1% = CVI(Left$(XX$, 2))
       REFE$ = AJUSTI$(RASOCLI$(NCX1%), 30)
       VALOX# = CVD(Mid$(XX$, 301, 8))
       MONEX$ = "$ "
       DEPOX% = Asc(Mid$(XX$, 245, 1))
       If DEPOX% < 1 Then DEPOX% = 1
       CASALX# = CVS(Mid$(XX$, 225, 4))
       GENER$ = Mid$(XX$, 193, 6)
       MARCASTA% = Asc(Mid$(XX$, 246, 1))
       '
       If REMOVI& <= UREMOVI& Then
         YY$ = REGLEIDO$("MOVISTO", REMOVI&)
         If CVI(Left$(YY$, 2)) = FEMOX% Then
           If CVI(Mid$(YY$, 3, 2)) = CIPIX% Then
             If TRIM$(Mid$(YY$, 21, 2)) = CMOVX$ Then
               If TRIM$(Mid$(YY$, 23, 10)) = DOPRI$ Then
                 If CVI(Mid$(YY$, 84, 2)) = NCX1% Then
                   If Abs(CASALX# - (CVD(Mid$(YY$, 96, 8)) - CVD(Mid$(YY$, 88, 8)))) < 0.05 Then
                     ' MOVIMIENTO O.K.
                     If FEMOX% < HOII% Then
                       Call REPLA(XX$, MKI$(0), 67, 2)
                       Call GRAREG("MAVENDID", XX$, U&)
                     End If
                     GoTo 49
                   End If
                 End If
               End If
             End If
           End If
         End If
       End If
       '
       REMOREC$ = REBLAMO$
       Call REPLA(REMOREC$, MKI$(FEMOX%), 1, 2)
       Call REPLA(REMOREC$, MKI$(CIPIX%), 3, 2)
       Call REPLA(REMOREC$, LOTEX$, 5, 16)
       Call REPLA(REMOREC$, CMOVX$, 21, 2)
       Call REPLA(REMOREC$, AJUSTI$(DOPRI$, 10), 23, 10)
       Call REPLA(REMOREC$, AJUSTI$(DOSEC$, 10), 33, 10)
       Call REPLA(REMOREC$, REFE$, 43, 30)
       Call REPLA(REMOREC$, MKD$(VALOX#), 73, 8)
       Call REPLA(REMOREC$, MONEX$, 81, 2)
       Call REPLA(REMOREC$, Chr$(DEPOX%), 83, 1)
       Call REPLA(REMOREC$, MKI$(NCX1%), 84, 2)
       If CASALX# > 0 Then
           Call REPLA(REMOREC$, MKD$(CASALX#), 96, 8)
         Else
           Call REPLA(REMOREC$, MKD$(Abs(CASALX#)), 88, 8)
       End If
       Call REPLA(REMOREC$, GENER$, 122, 5)
       Call REGAPP("MOVISTO", REMOREC$)
       '
       REMOVI& = ULTREG&("MOVISTO")
       XX$ = REGLEIDO$("MAVENDID", U&)
       Call REPLA(XX$, MKS$(REMOVI&), 247, 4)
       Call GRAREG("MAVENDID", XX$, U&)
       '
       REMOREK$ = String$(12, 0) + REMOREC$
       Call REPLA(REMOREK$, "MOVISTO ", 1, 8)
       Call REPLA(REMOREK$, MKS$(REMOVI), 9, 4)
       Call REGAPP("MOVIRECU", REMOREK$)
       '
     End If
49 Next UK&
   '
   Call CIERRARCH("MAVENDID")
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("MOVIRECU")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   Screen.MousePointer = 1
   '
End Sub


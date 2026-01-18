VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form LISTAP14 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación de Orden de Pago por Impuestos"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   11145
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text21 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Text            =   "Generación de Lista para Pagos de Impuestos "
      Top             =   120
      Width           =   10920
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   34
      Top             =   7680
      Width           =   10935
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
         Height          =   645
         Left            =   120
         Picture         =   "LISTAP14.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Invierte Selección"
         Top             =   240
         Width           =   4245
      End
      Begin VB.TextBox Text23 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   600
         Width           =   1365
      End
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Index           =   0
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "Cant. Item Selecc.:"
         ToolTipText     =   "Cantidad de Item Seleccionados"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Index           =   0
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "Cantidad ItemsTotal:"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text19 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   240
         Width           =   1365
      End
      Begin VB.TextBox Text18 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         Width           =   1245
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "Imorte Total:"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   600
         Width           =   1245
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   4560
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "Total Seleccionado:"
         Top             =   600
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   10935
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   4035
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   10680
         _ExtentX        =   18838
         _ExtentY        =   7117
         _Version        =   393216
         BackColorFixed  =   16501405
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Anulación Total"
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
         Left            =   4320
         TabIndex        =   45
         Top             =   6300
         Width           =   2175
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Salir - Modo Edición"
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
         Left            =   120
         TabIndex        =   44
         Top             =   6300
         Width           =   2175
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Modificar - Imprimir"
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
         Left            =   8600
         TabIndex        =   43
         Top             =   6300
         Width           =   2175
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "Referencia"
         Top             =   6800
         Width           =   975
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1080
         MaxLength       =   64
         TabIndex        =   25
         Top             =   6800
         Width           =   7485
      End
      Begin VB.Frame Frame5 
         Height          =   1695
         Left            =   7920
         TabIndex        =   12
         Top             =   120
         Width           =   2895
         Begin VB.CommandButton Command4 
            Caption         =   "Listar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   31
            Top             =   1320
            Width           =   2680
         End
         Begin VB.TextBox Text13 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   1800
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   960
            Width           =   1005
         End
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Text            =   "Fecha - Generación"
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   960
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   600
            Width           =   1850
         End
         Begin VB.TextBox Text10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Text            =   "Usuario:"
            Top             =   600
            Width           =   860
         End
         Begin VB.CommandButton Command2 
            Caption         =   "!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1680
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   240
            Width           =   195
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0FF&
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
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   23
            TabStop         =   0   'False
            Text            =   "Listas Generadas"
            Top             =   240
            Width           =   1550
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1880
            TabIndex        =   22
            Top             =   240
            Width           =   930
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1695
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   7815
         Begin VB.TextBox TXTNROOPAGO 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   32
            TabStop         =   0   'False
            Text            =   "O.Pago Nº"
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   5280
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Text            =   "Total:"
            Top             =   1320
            Width           =   855
         End
         Begin VB.TextBox TXTTOTAL 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1575
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Listar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   120
            TabIndex        =   19
            Top             =   1320
            Width           =   2680
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Rango de Fechas"
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
            Left            =   2640
            TabIndex        =   18
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Index           =   0
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "Desde:"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox TXTDESDE 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   5115
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   240
            Width           =   915
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Index           =   1
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "Hasta:"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox TXTHASTA 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Left            =   6795
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   240
            Width           =   915
         End
         Begin VB.TextBox Text6 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Index           =   5
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   960
            Width           =   6495
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   1395
            TabIndex        =   1
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FBCA9D&
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
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   10
            TabStop         =   0   'False
            Text            =   "Cod.Imp.:"
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
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
            Index           =   4
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   9
            TabStop         =   0   'False
            Text            =   "Descripción"
            Top             =   960
            Width           =   1095
         End
         Begin VB.CommandButton Command5 
            Caption         =   "!"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   240
            Width           =   195
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Grabar - Imprimir"
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
         Left            =   8600
         TabIndex        =   2
         Top             =   6800
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
         Height          =   225
         Index           =   0
         Left            =   0
         OleObjectBlob   =   "LISTAP14.frx":014A
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   1485
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   375
      Index           =   0
      Left            =   2640
      ScaleHeight     =   315
      ScaleWidth      =   1515
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "LISTAP14.frx":01A8
      Top             =   120
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "LISTAP14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function DESCRIIMPU$(CODIMPU%)
    CONDI$ = "CODINT = " & CODIMPU%
    DESCRIIMPU$ = VALOBADA("ITEMS", "DESCRIPCION", CONDI$, "NOMESS")
End Function

Function IMPUESTO_YA_INGRESADO%()
  IMPUESTO_YA_INGRESADO% = 0
  For I& = 1 To MSF1.Rows - 1
     CODIMPU% = XVALO(MSF1.TextMatrix(I&, 5))
     If CODIMPU% = XVALO(Text7) Then
        IMPUESTO_YA_INGRESADO% = 1
        Exit For
     End If
  Next I&
End Function

Private Sub Command1_Click()
   Command1.Enabled = False
   '
   NORPAG& = XVALO(TXTNROOPAGO)
   
   If MSF1.Rows < 2 Then
      Call COMUNI("Falta Detalle de Lista de Impuestos a Pagar")
      GoTo 99
   End If
   '
   If XVALO(Text23) < 1 Then
      Call COMUNI("No Hay Items Seleccionados")
      GoTo 99
   End If
   '
   If TRIM$(Text5) = "" Then
     Call COMUNI("Falta Referencia de Lista")
     Call PINTATEXT(Text5)
     Text5.SetFocus
     GoTo 99
   End If
   '
   CONDI$ = "NUMLISTA >= 0 "
   NUMLI& = 1 + XVALO(VALOBADA("CAJABANC", "MAX(NUMLISTA)", CONDI$, "NOMESS"))
   NUSER1% = USNBR% Mod 100
   '
   If PRONUORPA& <> NORPAG& Then
       Call COMUNI("Se va a Actualizar el Número de Orden de Pago")
       TXTNROOPAGO = PRONUORPA&
       GoTo 99
   End If
   '
   
   COCOR$ = "00000000" & TRIM$(TXTNROOPAGO): COCOR$ = Right$(COCOR$, 8)
   COCOR$ = "OP-" & COCOR$
   COLAR$ = COCOR$
   '
   'PRESENTO VENTANA PARA QUE EL USUARIO INDIQUE CUANTO DINERO SE LLEVA PARA EL PAGO
    RETIPAGO.Text1 = FORMATNUM$(XVALO(Text15), "F12.2")
    RETIPAGO.IMPORTE_ORIGINAL = XVALO(Text15)
    RETIPAGO.Show 1
    If XVALO(RETIPAGO.Text1) < 1 Then
       GoTo 99
    End If
    IMPO_OPAGO# = XVALO(RETIPAGO.Text1)
    REFERENCIA_LISTA$ = Left$(Text5, 64)

'   Dim RS2 As ADODB.Recordset
'   Set RS2 = New ADODB.Recordset
'25 On Error Resume Next
'   SQLX$ = "SELECT * FROM CAJABANC  WHERE nuclien = 1  "
'   RS2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
'   FLEXCONN.BeginTrans
'   With RS2

    Screen.MousePointer = 11
    For I& = 1 To MSF1.Rows - 1
      If MSF1.TextMatrix(I&, 0) = "*" Then
        
        NC% = XVALO(MSF1.TextMatrix(I&, 1))
        CONCEPTO$ = MSF1.TextMatrix(I&, 3)
        IMPOX# = XVALO(MSF1.TextMatrix(I&, 4))
        CODIMPU% = MSF1.TextMatrix(I&, 5)
        IDENTIF$ = MSF1.TextMatrix(I&, 6)
        
        CONDI$ = "ID = '" & IDENTIF$ & "'"
        CONDI$ = CONDI$ + " And NUCLIEN = " & NC%
        CONDI$ = CONDI$ + " And NUMEIMPU = " & CODIMPU%
        Call ACTUREGISTRO("CAJABANC", "STATUS", CONDI$, 2, REGAF&, "NOMESS")
        Call ACTUREGISTRO("CAJABANC", "NUMLISTA", CONDI$, NUMLI&, REGAF&, "NOMESS")
        Call ACTUREGISTRO("CAJABANC", "VALORREAL_IMP", CONDI$, NUMLI&, REGAF&, "NOMESS")
        '
      End If
    Next I&
          '
   'INICIALIZO
   VUELTA% = 0
  
65 ABUS& = RANDSTRING$(-8)
   If CantRegistros("CAJABANC", "NUMECOMP = " & ABUS&, "NOMESS") > 0 Then GoTo 65
   Call CreaRegistro("CAJABANC", "NUMECOMP", ABUS&, "NOMESS")
   '
   CONDI$ = "NUMECOMP =   " & ABUS&
   CUENTACONTABLE% = CUEINT(Control$("AMACLI", "ECUECON"))
   Call ACTUREGISTRO("CAJABANC", "NUCLIEN", CONDI$, 0, REGAF&, "NOMESS")
   '
   If VUELTA% = 0 Then
    Call ACTUREGISTRO("CAJABANC", "!IDEBECOMP ", CONDI$, 0, REGAF&, "NOMESS")
    Call ACTUREGISTRO("CAJABANC", "IHABECOMP", CONDI$, FORMATNUM$(IMPO_OPAGO#, "F16.5"), REGAF&, "NOMESS")
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", CONDI$, CUENTACONTABLE%, REGAF&, "NOMESS")
    Call ACTUREGISTRO("CAJABANC", "TIPOMOVIM", CONDI$, "OPAG", REGAF&, "NOMESS")
  Else
    Call ACTUREGISTRO("CAJABANC", "!IDEBECOMP ", CONDI$, FORMATNUM$(IMPO_OPAGO#, "F16.5"), REGAF&, "NOMESS")
    Call ACTUREGISTRO("CAJABANC", "IHABECOMP", CONDI$, 0, REGAF&, "NOMESS")
    Call ACTUREGISTRO("CAJABANC", "CUECONTAI", CONDI$, CUENTACONTABLE%, REGAF&, "NOMESS")
    Call ACTUREGISTRO("CAJABANC", "TIPOMOVIM", CONDI$, "CHENTR", REGAF&, "NOMESS")
  End If
  '
  Call ACTUREGISTRO("CAJABANC", "CODICOM_COR", CONDI$, COCOR$, REGAF&, "NOMESS")
  Call ACTUREGISTRO("CAJABANC", "CODICOM_LAR", CONDI$, COCOR$, REGAF&, "NOMESS")
  Call ACTUREGISTRO("CAJABANC", "NUMEIMPU", CONDI$, CODIMPU%, REGAF&, "NOMESS")
  Call ACTUREGISTRO("CAJABANC", "Status", CONDI$, 0, REGAF&, "NOMESS")
  Call ACTUREGISTRO("CAJABANC", "ID", CONDI$, "", REGAF&, "NOMESS")
  Call ACTUREGISTRO("CAJABANC", "NUMLISTA", CONDI$, NUMLI&, REGAF&, "NOMESS")
  Call ACTUREGISTRO("CAJABANC", "NUMEUSER", CONDI$, NUSER1%, REGAF&, "NOMESS")
  Call ACTUREGISTRO("CAJABANC", "FECHEMISI", CONDI$, CVDAT(HOY(HO$)), REGAF&, "NOMESS")
  Call ACTUREGISTRO("CAJABANC", "REFERENCIA_LISTA", CONDI$, REFERENCIA_LISTA$, REGAF&, "NOMESS")
  Call ACTUREGISTRO("CAJABANC", "NUMECOMP", CONDI$, NORPAG&, REGAF&, "NOMESS")
  If VUELTA% = 0 Then
    VUELTA% = 1
    CUENTACONTABLE% = CUEINT(Control$("", "CUENCAJA"))
    GoTo 65
  End If

'   SQLX$ = "SELECT * FROM CAJABANC  WHERE NUMECOMP =   " & ABUS&
'   RS2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
   
'
'   With RS2
'    '
'      !NUCLIEN = 0
'      If VUELTA% = 0 Then
'        !IDEBECOMP = 0
'        !IHABECOMP = IMPO_OPAGO#
'        !CUECONTAI = CUENTACONTABLE%
'        !TIPOMOVIM = "OPAG"
'      Else
'        !IDEBECOMP = IMPO_OPAGO#
'        !IHABECOMP = 0
'        !CUECONTAI = CUENTACONTABLE%
'        !TIPOMOVIM = "CHENTR"
'      End If
'      !CODICOM_COR = COCOR$
'      !CODICOM_LAR = COCOR$
'      !NUMEIMPU = CODIMPU%
'      !Status = 0
'      !ID = ""
'      !NUMLISTA = NUMLI&
'      !NUMEUSER = NUSER1%
'      !FECHEMISI = CVDAT(HOY(HO$))
'      !NUMECOMP = NORPAG&
'      !REFERENCIA_LISTA = REFERENCIA_LISTA$
'      .Update
'      If VUELTA% = 0 Then
'         VUELTA% = 1
'         CUENTACONTABLE% = CUEINT(Control$("", "CUENCAJA"))
'
'         GoTo 65
'      End If
'   End With
'   RS2.Close
'   Set RS2 = Nothing
'      Call COMUNI("Grabación Realizada Exitosamente")
   
    FREP = HOY(HO$)
    FILTX$ = TRIM$(Str$(NUMLI&)) & ";" & TRIM$(FREP) & ";" & TRIM$(Str$(IMPO_OPAGO#))
    Call STDFORM("LIST-GTS", FILTX$)
    '
    Text2 = ""
    Text5 = ""
    Text7 = ""
    TXTNROOPAGO = ""
    MSF1.Rows = 1
    TXTNROOPAGO = PRONUORPA&
99 Command1.Enabled = True
   Screen.MousePointer = 1
End Sub



Private Sub Command10_Click()
   Command10.Enabled = False
   If DERACCE%("ANULORPA", "Anulación de Orden de Pago de Impuestos") < 2 Then Exit Sub
   '
   NULI& = XVALO(Text2)
   CONDI$ = "NUMLISTA = " & NULI& & " AND TIPOMOVIM = 'OPAG'"
   NORPA$ = VALOBADA("CAJABANC", "CODICOM_COR", CONDI$, "NOMESS")

   If COMALTER%("Confirmación:\\¿Confirma Eliminar Orden de Pago Nª: ? " & TRIM$(NORPA$) & "\Cancelar") <> 1 Then
       GoTo 99
   End If
   '
   'BORRO LOS 2 MOVIMIENTOS QUE SE GENERAN CON EL IMPORTE TOTAL DE LA ORDEN DE PAGO DE IMPUESTOS
   NULI& = XVALO(Text2)
   CONDI$ = "NUMLISTA = " & NULI& & " AND TIPOMOVIM = 'OPAG'"
   NORPA$ = VALOBADA("CAJABANC", "CODICOM_COR", CONDI$, "NOMESS")
   '
   CONDI$ = "CODICOM_COR = '" & NORPA$ & "'"
   Call BORRAREGISTROS("CAJABANC", CONDI$, REGAF&, "NOMESS")
   '
   'VUELVO PARA ATRAS EL STATUS Y EL NUMERO  DE LISTA PONIENDOLOS EN 0
   For I& = 1 To MSF1.Rows - 1
        NC% = XVALO(MSF1.TextMatrix(I&, 1))
        CONCEPTO$ = MSF1.TextMatrix(I&, 3)
        IMPOX# = XVALO(MSF1.TextMatrix(I&, 4))
        CODIMPU% = MSF1.TextMatrix(I&, 5)
        IDENTIF$ = MSF1.TextMatrix(I&, 6)
        CONDI$ = "ID = '" & IDENTIF$ & "'"
        CONDI$ = CONDI$ + " And NUCLIEN = " & NC%
        CONDI$ = CONDI$ + " And NUMEIMPU = " & CODIMPU%
        Call ACTUREGISTRO("CAJABANC", "STATUS", CONDI$, 0, REGAF&, "NOMESS")
        Call ACTUREGISTRO("CAJABANC", "NUMLISTA", CONDI$, 0, REGAF&, "NOMESS")
    Next I&
   '
   TXTNROOPAGO = PRONUORPA&
   Call Command9_Click
99 Command10.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call ABRECABAN
   CONDI$ = " NUMLISTA > 0 AND TIPOMOVIM = 'OPAG'"
   Call CARGALISEL("LISTEMITI", "CAJABANC", "NUMLISTA/F6.0;REFERENCIA_LISTA/A32", CONDI$, "DISTINCT/NOMESS")
   Call SELECHO("LISTEMITI/Listas Emitidas (Todas)")
   LISTX& = XVALO(VALACT1$("LISTEMITI"))
   If LISTX& < 1 Then GoTo 99
   '
   Text2 = LISTX&
   
99 Command2.Enabled = True
End Sub


Private Sub Command3_Click()
    Command3.Enabled = False
    OBX$ = ""
    If TRIM$(TXTDESDE) <> "" Then
       Call REPLA(OBX$, MKI$(CVINT(TXTDESDE)), 1, 2)
    End If
    If TRIM$(TXTHASTA) <> "" Then
       Call REPLA(OBX$, MKI$(CVINT(TXTHASTA)), 3, 2)
    End If
    VDEF$ = OBX$
10  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        VDEF$ = OBX$
        GoTo 10
    End If
    DESDE = FECHATEX$(DES%)
    HASTA = FECHATEX$(HAS%)
    '
    TXTDESDE = "" ' PARA QUE PROVOQUE EL EVENTO CHANGE DE ACTUALIZAR
    TXTHASTA.TEXT = HASTA
    TXTDESDE = DESDE
    '
99  Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Text5 = ""
   Text11 = ""
   Text13 = ""
   TXTDESDE = ""
   TXTHASTA = ""
   Text6(5) = ""
   Text7 = ""
   Command8.Enabled = True
   Command9.Enabled = True
   Command10.Enabled = True
   '
   'TRAE EL NUMERO DE OREDEN DE PAGO
   NULI& = XVALO(Text2)
   CONDI$ = "NUMLISTA = " & NULI& & " AND TIPOMOVIM = 'OPAG'"
   NORPA$ = VALOBADA("CAJABANC", "CODICOM_COR", CONDI$, "NOMESS")
   TXTNROOPAGO = XVALO(Mid$(NORPA$, 4))
   '
   Command1.Visible = False
   MSF1.Height = 4275
   
   MSF1.Rows = 1
   MSF1.BackColorFixed = &HC0C0FF
   LISTX& = XVALO(Text2)
   SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE NUMLISTA = " & LISTX&
   SQLX$ = SQLX$ + " AND TIPOMOVIM = 'CHRCIM'"
   VUELTA% = 0
   Set RS = READSET(SQLX$)
   With RS
    Do While Not .EOF
          I& = MSF1.Rows
          MSF1.Rows = MSF1.Rows + 1
          '
          NC% = XVALO(!NUCLIEN)
          IMPUESTO$ = SAFETEXT$(!DESCRI_IMPU)
          IMPOX# = XVALO(!IDEBECOMP)
          CONCEPTO$ = SAFETEXT$(!DESCRIMOV)
          CODIMPU% = SAFETEXT$(!NUMEIMPU)
          IDENTIF$ = SAFETEXT$(!ID)
          MSF1.TextMatrix(I&, 1) = SAFETEXT$(NC%)
          MSF1.TextMatrix(I&, 2) = RASOCLI$(NC%)
          MSF1.TextMatrix(I&, 3) = CONCEPTO$
          MSF1.TextMatrix(I&, 4) = FORMATNUM$(IMPOX#, "F12.2")
          MSF1.TextMatrix(I&, 5) = SAFETEXT$(CODIMPU%)
          MSF1.TextMatrix(I&, 6) = IDENTIF$
          If VUELTA% = 0 Then
            VUELTA% = 1
              CONDI = "TIPOMOVIM= 'OPAG' AND NUMLISTA = " & LISTX&
              Text5 = VALOBADA("CAJABANC", "REFERENCIA_LISTA", CONDI$, "NOMESS")
              Text11 = USERNAM$(XVALO(!NUMEUSER))
              Text13 = SAFETEXT$(!FECHEMISI)
              Text6(5) = IMPUESTO$
          End If
          '
          .MoveNext
     Loop
   End With
   '
   Call ACTUTOTALLISTA

End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call SELECHO("ITEMDAT/IMPUESTOS")
    CI1% = XVALO(VALACT1$("ITEMDAT"))
    Text7 = CI1%
    Command5.Enabled = True

End Sub

Private Sub Command6_Click()
  '
  Command6.Enabled = False
  Text5 = ""
  Text11 = ""
  Text13 = ""
  Text2 = ""
  '
  'ACTUALIZA EL NUMERO DE ORDEN DE PAGO SIGUIENTE
  TXTNROOPAGO = PRONUORPA&
  '
  Command1.Visible = True
  MSF1.Height = 4755
  MSF1.ZOrder 0
  MSF1.Rows = 1
  MSF1.BackColorFixed = &HFBCA9D
  '
  Screen.MousePointer = 11
  If IMPUESTO_YA_INGRESADO% > 0 Then
     Call COMUNI("El Impuesto Seleccionado Ya Fue Ingresado en la Lista, Verifique !!!")
  End If
  
  If MSF1.Rows > 1 Then
   If COMALTER%("Confirmación:\\Agregar al Final \Sobre-Escribir") = 2 Then
       MSF1.Rows = 1
   End If
  End If

  CONDI$ = " STATUS = 0  AND TipoMovim = 'CHRCIM' AND NUMEIMPU = " & XVALO(Text7)
  SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)WHERE " & CONDI$
  'SI SE ELIGIO FECHA FILTRO POR LA FECHA
  If FECHANUM(TXTDESDE) > 0 And FECHANUM(TXTHASTA) > 0 Then
     DESDE = FETEXCOR1$(FECHANUM(TXTDESDE))
     HASTA = FETEXCOR1$(FECHANUM(TXTHASTA))
     SQLX$ = SQLX$ + " AND FECHEMISI >= '" & DESDE & "' AND FECHEMISI <= '" & HASTA & "'"
  End If
  '
  Set RS = READSET(SQLX$)
  With RS
    Do While Not .EOF
          I& = MSF1.Rows
          MSF1.Rows = MSF1.Rows + 1
          NC% = XVALO(!NUCLIEN)
          IMPUESTO$ = SAFETEXT$(!DESCRI_IMPU)
          IMPOX# = XVALO(!IDEBECOMP)
          CONCEPTO$ = SAFETEXT$(!DESCRIMOV)
          CODIMPU% = XVALO(!NUMEIMPU)
          IDENT$ = SAFETEXT$(!ID)
          '
          MSF1.TextMatrix(I&, 1) = SAFETEXT$(NC%)
          MSF1.TextMatrix(I&, 2) = RASOCLI$(NC%)
          MSF1.TextMatrix(I&, 3) = CONCEPTO$
          MSF1.TextMatrix(I&, 4) = FORMATNUM$(IMPOX#, "F12.2")
          MSF1.TextMatrix(I&, 5) = SAFETEXT$(CODIMPU%)
          MSF1.TextMatrix(I&, 6) = (IDENT$)
          .MoveNext
    Loop
  End With
  'Call ACTUTOTALLISTA
   Call ACTUTOTALES

  Command6.Enabled = True
  Screen.MousePointer = 1
End Sub
Sub ACTUTOTALLISTA()
   '
   ACUMU# = 0
   For I& = 1 To MSF1.Rows - 1
      IMPOX# = ROUND(XVALO(MSF1.TextMatrix(I&, 4)), 2)
      ACUMU# = ACUMU# + IMPOX#
   Next I&
   TXTTOTAL = FORMATNUM$(ACUMU#, "F12.2")
   '
End Sub

Private Sub MnuSalimp_Click()

End Sub


Private Sub Command7_Click()
   Command7.Enabled = False
   For I& = 1 To MSF1.Rows - 1
     Call COLOREAR_GRILLA_SOLAFILA(LISTAP14.MSF1, &HFFFF80, I&)
   Next I&
   Call ACTUTOTALES
   'Call ACTUCANTISELECCION
   
   Command7.Enabled = True

End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   '
   If DERACCE%("MODIORPA", "MODIFICAR ORDEN DE PAGO DE IMPUESTOS") < 2 Then
      GoTo 99
   End If
   
   NORPAG& = XVALO(TXTNROOPAGO)
   
   If MSF1.Rows < 2 Then
      Call COMUNI("Falta Detalle de Lista de Impuestos a Pagar")
      GoTo 99
   End If
   '
   If XVALO(Text23) < 1 Then
      Call COMUNI("No Hay Items Seleccionados")
      GoTo 99
   End If
   '
   If TRIM$(Text5) = "" Then
     Call COMUNI("Falta Referencia de Lista")
     Call PINTATEXT(Text5)
     Text5.SetFocus
     GoTo 99
   End If
   '
   If XVALO(Text2) < 1 Then
     Call COMUNI("Falta Número de Listado de Pago")
     Call PINTATEXT(Text2)
     Text2.SetFocus
   End If
   
   NUMLI& = XVALO(Text2)
   '
   COCOR$ = "00000000" & TRIM$(TXTNROOPAGO): COCOR$ = Right$(COCOR$, 8)
   COCOR$ = "OP-" & COCOR$
   COLAR$ = COCOR$
   '
   'PRESENTO VENTANA PARA QUE EL USUARIO INDIQUE CUANTO DINERO SE LLEVA PARA EL PAGO
    RETIPAGO.Text1 = FORMATNUM$(XVALO(Text15), "F12.2")
    RETIPAGO.IMPORTE_ORIGINAL = XVALO(Text15)
    RETIPAGO.Show 1
    If XVALO(RETIPAGO.Text1) < 1 Then
       GoTo 99
    End If
    IMPO_OPAGO# = XVALO(RETIPAGO.Text1)
    REFERENCIA_LISTA$ = Left$(Text5, 64)

    For I& = 1 To MSF1.Rows - 1
        NC% = XVALO(MSF1.TextMatrix(I&, 1))
        CONCEPTO$ = MSF1.TextMatrix(I&, 3)
        IMPOX# = XVALO(MSF1.TextMatrix(I&, 4))
        CODIMPU% = MSF1.TextMatrix(I&, 5)
        IDENTIF$ = MSF1.TextMatrix(I&, 6)
        CONDI$ = "ID = '" & IDENTIF$ & "'"
        CONDI$ = CONDI$ + " And NUCLIEN = " & NC%
        CONDI$ = CONDI$ + " And NUMEIMPU = " & CODIMPU%
      If MSF1.TextMatrix(I&, 0) = "*" Then
        'SI ESTA SELECCIONADO LO REGISTRA NUEVAMENTE CON STATUS = 2 Y LE GRABA EL NUMERO DE LISTA
        Call ACTUREGISTRO("CAJABANC", "STATUS", CONDI$, 2, REGAF&, "NOMESS")
        Call ACTUREGISTRO("CAJABANC", "NUMLISTA", CONDI$, NUMLI&, REGAF&, "NOMESS")
      Else
        'CUANDO NO ESTA SELECCIONADO LE CAMBIA EL STATUS = 0 Y NUMERO DE LISTA A 0
        Call ACTUREGISTRO("CAJABANC", "STATUS", CONDI$, 0, REGAF&, "NOMESS")
        Call ACTUREGISTRO("CAJABANC", "NUMLISTA", CONDI$, 0, REGAF&, "NOMESS")
      End If
    Next I&
          '
    'BORRA LOS MOVIMIENTOS GENEERADOS DE LA ORDEN DE PAGO POR LOS TOTALES
    CONDI$ = "NUMLISTA = " & NUMLI& & " AND TIPOMOVIM = 'OPAG'"
    NORPA$ = VALOBADA("CAJABANC", "CODICOM_COR", CONDI$, "NOMESS")
    '
    CONDI$ = "CODICOM_COR = '" & NORPA$ & "'  AND TIPOMOVIM = 'OPAG'"
    Call ACTUREGISTRO("CAJABANC", "IHABECOMP", CONDI$, FORMATNUM$(IMPO_OPAGO#, "F16.5"), REGAF&, "NOMESS")
    Call ACTUREGISTRO("CAJABANC", "REFERENCIA_LISTA", CONDI$, Text5, REGAF&, "NOMESS")
    '
    CONDI$ = "CODICOM_COR = '" & NORPA$ & "'  AND TIPOMOVIM = 'CHENTR'"
    Call ACTUREGISTRO("CAJABANC", "IDEBECOMP", CONDI$, FORMATNUM$(IMPO_OPAGO#, "F16.5"), REGAF&, "NOMESS")
    Call ACTUREGISTRO("CAJABANC", "REFERENCIA_LISTA", CONDI$, Text5, REGAF&, "NOMESS")
    '
    
'   VUELTA% = 0
'   Dim RS1 As ADODB.Recordset
'   Set RS1 = New ADODB.Recordset
'   ' ACA PRIMERO BORRAR LOS REGISTROS O VER SI SOLO NO HAY QUE CAMBIAR LOS DATOS DEL IMPORTE
'   SQLX$ = "SELECT * FROM CAJABANC  WHERE nuclien < 1  "
'   RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic
'   With RS1
'65    .AddNew
'      !NUCLIEN = NC%
'      If VUELTA% = 0 Then
'        !IDEBECOMP = 0
'        !IHABECOMP = IMPO_OPAGO#
'        !CUECONTAI = CUENTACONTABLE%
'        !TIPOMOVIM = "OPAG"
'      Else
'        !IDEBECOMP = IMPO_OPAGO#
'        !IHABECOMP = 0
'        CUENTACONTABLE% = CUEINT(Control$("", "CUENCAJA"))
'        !CUECONTAI = CUENTACONTABLE%
'        !TIPOMOVIM = "CHENTR"
'      End If
'      !CODICOM_COR = COCOR$
'      !CODICOM_LAR = COCOR$
'      !NUMEIMPU = CODIMPU%
'      !Status = 0
'      !ID = ""
'      !NUMLISTA = NUMLI&
'      !NUMEUSER = NUSER1%
'      !FECHEMISI = CVDAT(HOY(HO$))
'      !NUMECOMP = NORPAG&
'      !REFERENCIA_LISTA = REFERENCIA_LISTA$
'      .Update
'      If VUELTA% = 0 Then
'         VUELTA% = 1
'         GoTo 65
'      End If
'   End With
   
   
    FREP = HOY(HO$)
    FILTX$ = TRIM$(Str$(NUMLI&)) & ";" & TRIM$(FREP) & ";" & TRIM$(Str$(IMPO_OPAGO#))
    Call STDFORM("LIST-GTS", FILTX$)
    
    Text2 = ""
    Text5 = ""
    Text7 = ""
    MSF1.Rows = 1
    Call Command9_Click
    TXTNROOPAGO = PRONUORPA&
99 Command8.Enabled = True

End Sub

Private Sub Command9_Click()
  Command9.Enabled = False
  Command6.Enabled = True
  Text5 = ""
  Text11 = ""
  Text13 = ""
  Text2 = ""
  TXTNROOPAGO = ""
  '
  Command8.Enabled = False
  Command9.Enabled = False
  Command10.Enabled = False
  MSF1.Rows = 1
  MSF1.Height = 4755
  Command9.Enabled = True
End Sub

Private Sub Form_Load()

    Call CENTRAFORM(Me)
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    Call CARGA_ENCABEZADO(LISTAP14.MSF1, "Cta/F6;Nombre/A32;Concepto/A32;Importe/F12.2;Cod_Impuesto/A0;IDENTIFICADOR/A0;ValorRealPagado/F12.2", LISTAP14)
    '
    TXTNROOPAGO = ""
    TXTNROOPAGO = PRONUORPA&
    MSF1.Height = 4755
    '
End Sub

Sub ACTUCANTISELECCION()
   ACUMUCANTI = 0
   For I& = 1 To MSF1.Rows - 1
     If MSF1.TextMatrix(I&, 0) = "*" Then
       ACUMUCANTI = ACUMUCANTI + 1
     End If
   Next I&
   Text23 = ACUMUCANTI
End Sub

Private Sub Mnumod_Click()

End Sub

Private Sub MSF1_DblClick()
    I& = MSF1.MouseRow
    j& = MSF1.MouseCol
    If I& = 0 Or I& > MSF1.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF1.Row = I&
    MSF1.COL = j&
    On Error Resume Next

    REG& = I&
    If REG& < 1 Then
        Exit Sub
    End If
    Call COLOREAR_GRILLA_SOLAFILA(LISTAP14.MSF1, &HFFFF80, REG&)
    Call ACTUTOTALES
End Sub
Sub ACTUTOTALES()
   Text18 = SUMA_COLUMNA_SELECCION(LISTAP14.MSF1, 4) 'TOTAL IMPORTES TODOS
   Text15 = SUMA_COLUMNA_SELECCION(LISTAP14.MSF1, 4, "*") 'TOTAL SELECCIONADO
   'Call ACTUTOTALLISTA
   Text19 = TOTALFILAS&(LISTAP14.MSF1)
   Call ACTUCANTISELECCION

End Sub

Private Sub Text2_Change()
  '
   MSF1.Rows = 1
   If TRIM$(Text2) = "" Then
        Command1.Caption = "Grabar - Imprimir"
        Text5 = ""
        Text11 = ""
        Text13 = ""
        TXTDESDE = ""
        TXTHASTA = ""
        Text6(5) = ""
        TXTTOTAL = ""
        Text5 = ""
        TXTNROOPAGO = PRONUORPA&
   Else
     Command1.Caption = "Imprimir"
   End If
  '
'
'   MSF1.Rows = 1
'   LISTX& = XVALO(Text2)
'   SQLX$ = "SELECT * FROM LISTAIMPU WITH(NOLOCK) "
'   SQLX$ = SQLX$ + " WHERE NUMLISTA = " & LISTX&
'
'   Set RS = READSET(SQLX$)
'   With RS
'    Do While Not .EOF
'          I& = MSF1.Rows
'          MSF1.Rows = MSF1.Rows + 1
'          '
'          NC% = XVALO(!NUCLIEN)
'          IMPUESTO$ = SAFETEXT$(!DESCRI_IMPU)
'          IMPOX# = XVALO(!IMPORTE)
'          CONCEPTO$ = SAFETEXT$(!CONCEPTO)
'          CODIMPU% = SAFETEXT$(!NUMEIMPU)
'          'VER SI NO HABRIA QUE MOSTRAR USUARIO
'          MSF1.TextMatrix(I&, 1) = SAFETEXT$(NC%)
'          MSF1.TextMatrix(I&, 2) = RASOCLI$(NC%)
'          MSF1.TextMatrix(I&, 3) = CONCEPTO$
'          MSF1.TextMatrix(I&, 4) = FORMATNUM$(IMPOX#, "F12.2")
'          MSF1.TextMatrix(I&, 5) = SAFETEXT$(CODIMPU%)
'          Text5 = SAFETEXT$(!REFERENCIA_LISTA)
'          Text11 = USERNAM$(XVALO(!NUMEUSER))
'          Text13 = SAFETEXT$(!FECHEMISI)
'          '
'          .MoveNext
'     Loop
'   End With
'   '
'   Call ACTUTOTALLISTA
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call Command4_Click
   End If
End Sub


Private Sub Text7_Change()
 
   MSF1.Rows = 1
   CI% = XVALO(Text7)
'   If CI% < 1 Then Exit Sub
   '
   CONDI$ = "CODINT = " & CI% & " AND STATUS >= 0 "
   If CantRegistros("ITEMS", CONDI$, "NOMESS") > 0 Then
      SQLX$ = "SELECT * FROM ITEMS WITH(NOLOCK) WHERE CODINT = " & CI%
      Set RS = READSET(SQLX$)
      With RS
        Text6(5) = SAFETEXT$(!Descripcion)
      End With
   Else
      Text6(5) = ""
   End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call Command6_Click
   End If

End Sub

Function PRONUORPA&()
    
    Call ABRECONEXION
    SQLX$ = "SELECT NumeComp FROM CAJABANC WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'OPAG'   ORDER BY NUMECOMP DESC"
    Set RS = READSET(SQLX$)
    With RS
      AA = XVALO(!NUMECOMP)
      SIGUIENTE& = 1 + XVALO(!NUMECOMP)
    End With
    RS.Close
    Set RS = Nothing
    
    PRONUORPA& = SIGUIENTE&
    '
End Function

Private Sub TXTDESDE_Change()
   Call Command6_Click
End Sub

Private Sub TXTDESDE_DblClick()
   TXTDESDE = ""
End Sub

Private Sub TXTHASTA_Change()
   Call Command6_Click
End Sub

Private Sub TXTHASTA_DblClick()
     TXTHASTA = ""
End Sub

VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FIMPUC07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operaciones Comerciales - Imputación de Créditos y Anticipos"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11580
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   45
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00624E28&
      ForeColor       =   &H80000008&
      Height          =   7365
      Left            =   10680
      ScaleHeight     =   7335
      ScaleWidth      =   1260
      TabIndex        =   33
      Top             =   0
      Width           =   1290
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   49
         Top             =   5400
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   50
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Auto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   140
         Picture         =   "FIMPUC07.frx":0000
         TabIndex        =   44
         ToolTipText     =   "Aprueba - Graba la Imputación"
         Top             =   4515
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         Caption         =   "R-Cue"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   790
         Left            =   140
         Picture         =   "FIMPUC07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Resumen de Cuenta Corriente"
         Top             =   1785
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   140
         Picture         =   "FIMPUC07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Notas de Clientes"
         Top             =   2580
         Width           =   650
      End
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   790
         Left            =   140
         Picture         =   "FIMPUC07.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Aprueba - Graba la Imputación"
         Top             =   3675
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Help"
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
         Left            =   140
         Picture         =   "FIMPUC07.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   890
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Left            =   140
         Picture         =   "FIMPUC07.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   200
         Width           =   650
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8200
      Left            =   0
      TabIndex        =   0
      Top             =   -360
      Width           =   12000
      Begin VB.TextBox TICAMBIS 
         Enabled         =   0   'False
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
         Left            =   7065
         MaxLength       =   14
         TabIndex        =   31
         Top             =   7455
         Width           =   1065
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2D3C0&
         Caption         =   "DOCUMENTO A IMPUTAR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   4920
         TabIndex        =   22
         Top             =   480
         Width           =   5655
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Documento"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   240
            TabIndex        =   51
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label Label32 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   1800
            TabIndex        =   43
            Top             =   1035
            Visible         =   0   'False
            Width           =   330
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cotización"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   42
            Top             =   1500
            Width           =   750
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   41
            Top             =   1130
            Width           =   750
         End
         Begin VB.Label Label28 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   960
            TabIndex        =   40
            Top             =   1035
            Visible         =   0   'False
            Width           =   645
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   960
            TabIndex        =   39
            Top             =   1400
            Visible         =   0   'False
            Width           =   1170
         End
         Begin VB.Label Label15 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   4410
            TabIndex        =   30
            Top             =   105
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Saldo No Imputado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2040
            TabIndex        =   29
            Top             =   1500
            Width           =   1605
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total Bruto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   2150
            TabIndex        =   28
            Top             =   1130
            Width           =   1485
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Emisión"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   3000
            TabIndex        =   27
            Top             =   630
            Width           =   1275
         End
         Begin VB.Label Label25 
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
            Height          =   300
            Left            =   240
            TabIndex        =   26
            Top             =   525
            Width           =   2220
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   4350
            TabIndex        =   25
            Top             =   525
            Width           =   1170
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   3720
            TabIndex        =   24
            Top             =   1035
            Width           =   1800
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   3720
            TabIndex        =   23
            Top             =   1400
            Width           =   1800
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2D3C0&
         Caption         =   "CUENTA / CLIENTE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   4635
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   1155
            TabIndex        =   21
            Top             =   1450
            Width           =   3270
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   210
            TabIndex        =   20
            Top             =   1450
            Width           =   750
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   210
            TabIndex        =   19
            Top             =   1100
            Width           =   4215
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   210
            TabIndex        =   18
            Top             =   735
            Width           =   4215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   3570
            TabIndex        =   17
            Top             =   315
            Width           =   855
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2800
         Left            =   120
         ScaleHeight     =   2775
         ScaleWidth      =   10425
         TabIndex        =   2
         Top             =   2640
         Width           =   10450
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   380
            Left            =   -10
            ScaleHeight     =   345
            ScaleWidth      =   10425
            TabIndex        =   47
            Top             =   -10
            Width           =   10450
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   $"FIMPUC07.frx":107C
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   0
               TabIndex        =   48
               Top             =   60
               Width           =   10410
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Height          =   5630
            Left            =   0
            TabIndex        =   3
            Top             =   315
            Width           =   10515
            Begin VB.ListBox List2 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2055
               ItemData        =   "FIMPUC07.frx":1129
               Left            =   -15
               List            =   "FIMPUC07.frx":112B
               TabIndex        =   4
               Top             =   0
               Width           =   10455
            End
            Begin VB.ListBox List3 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2085
               Left            =   0
               TabIndex        =   15
               Top             =   0
               Width           =   10455
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   330
               Left            =   8355
               TabIndex        =   8
               Top             =   2100
               Width           =   1560
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   330
               Left            =   6600
               TabIndex        =   7
               Top             =   2100
               Width           =   1650
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
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
               Height          =   330
               Left            =   3500
               TabIndex        =   6
               Top             =   2100
               Width           =   1695
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "TOTAL GENERAL"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   30
               TabIndex        =   5
               Top             =   2205
               Width           =   1815
            End
         End
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "T-C:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6825
         TabIndex        =   32
         ToolTipText     =   "Doble-Click para Ajustar T-C"
         Top             =   7215
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   8490
         TabIndex        =   14
         Top             =   5745
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pendiente"
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
         Left            =   6915
         TabIndex        =   13
         Top             =   5835
         Width           =   1485
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5025
         TabIndex        =   12
         Top             =   5745
         Width           =   1695
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Imputado"
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
         Left            =   3450
         TabIndex        =   11
         Top             =   5835
         Width           =   1485
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1665
         TabIndex        =   10
         Top             =   5745
         Width           =   1695
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Importe Total"
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
         Left            =   -645
         TabIndex        =   9
         Top             =   5835
         Width           =   2220
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FIMPUC07.frx":112D
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "FIMPUC07.frx":1361
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
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
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "FIMPUC07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$
Dim MODOMONEI%

Sub Command1_Click()
    Command1.Enabled = False
    For U& = 1 To List2.ListCount
       If TRIM$(Mid$(List2.List(U& - 1), 70, 14)) <> "" Then
          Call MENSERR(24, "Imposible Aplicación Automática.")
          GoTo 99
       End If
    Next U&
    For U& = 1 To List2.ListCount
       List2.ListIndex = U& - 1
    Next U&
    APLIAUTO% = 1
    Command1.Enabled = True
    Call Command5_Click
    Exit Sub
    '
99  Command1.Enabled = True

End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NCLI$ = TRIM$(Label1.Caption)
    'Call CONECRUN("*NOTACLI/" + NCLI$, "Notas de Clientes")
    NC% = XVALO(Label1.Caption)
    Call NotaCli(NC%, 2)
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RECUEC07.Text1 = Label1
    RECUEC07.Show 1
    Command4.Enabled = True
End Sub

Private Sub Command5_Click()
 'jandy sql
 
    Command5.Enabled = False
    TOTAVAL# = CDbl(XVALO(Label14.Caption))
    TOTACAN# = CDbl(XVALO(Label18.Caption))
    If TOTACAN# > TOTAVAL# Then
        Call COMUNI("Operación Incorrecta.\  \Las Cancelaciones no Pueden\Exceder al Documento a Aplicar.")
        Command5.Enabled = True
        Exit Sub
    End If
    '
    TICAPLI# = 1
    If MODOMONEI% = 2 Then
       If XVALO(Label16) > 0.5 Then
          TICAPLI# = XVALO(Label16)
       End If
    End If
    '
    ' VERIFICAR COINCIDENCIA DE MONEDA DE EMISIÓN DE DOCUMENTOS
    MEMI% = XVALO(Label32)
    If MODOMONEI% = 2 Then
       YACAN% = 0
       For U& = 1 To List2.ListCount
         Z$ = List2.List(U& - 1)
         DOCU$ = Left$(Z$, 20)
         If DOCU$ <> "Anticipo no Aplicado" Then
            CANCE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
            If CANCE# > 0 Then
               YACAN% = YACAN% + 1
               DOCUCA$ = TRIM$(Left$(Z$, 18))
               FEMIX = FECHANUM(Mid$(Z$, 23, 8))
               NCX1% = XVALO(Label1)
               MORI% = MONEMISOR%(DOCUCA$, FEMIX, NCX1%)
               If MEMI% > 1 Then
                  If MORI% <> MEMI% Then
                     Call MENSERR(24, "Imposible Aplicar Sobre Comprobante\Número '" + DOCUCA$ + "'\por Diferir Moneda de Emisión.")
                     Command5.Enabled = True
                     Exit Sub
                  End If
               End If
            End If
         End If
       Next U&
       '
       If TICAPLI# <> 1 Then
         If YACAN% > 1 Then
           Call MENSERR(24, "Transaccion Cancelada.\Debe Aplicar a Un Solo\Documento por Vez.")
           Command5.Enabled = True
           Exit Sub
         End If
       End If
       '
    End If
    '
    Call APERBASDAT("CABAN")
    RECILAR$ = Label25
    CoComLar$ = RECILAR$
'    CodiEmp% = 0
'    CodiSuc% = SUC%
    NumeClie = XVALO(Label1)
    NumeProv% = 0
    ATX% = 1
    '
    ' DESCUENTA SALDO DEL RECIBO APLICADO
    SUDEBE# = 0
    
    'jandy sql
    strsql = "Select * from SaDeuPen where NuClien=" & NumeClie
    strsql = strsql & " And CODICOM_LAR = '" & RECILAR$ & "'"
    strsql = strsql & " And SALDAcre > 0"
    
    Dim SaDeuPenAcre As ADODB.Recordset
    Set SaDeuPenAcre = New ADODB.Recordset
25  On Error Resume Next
    SaDeuPenAcre.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    With SaDeuPenAcre
     ' .FindFirst "NuClien = " & NumeClie & " AND CODICOM_LAR = '" & RECILAR$ & "' AND SALDAcre > 0"
     ' If .NoMatch = False Then
       If Not (.EOF And .BOF) Then
        CoComCor$ = !CodiCom_Cor
        '.Edit
        !SALDAcre = !SALDAcre - ROUND#(TOTACAN# * TICAPLI#)
        If MODOMONEI% = 2 Then
           If MEMI% > 1 Then
              !SaldAcre_DOLA = !SaldAcre_DOLA - TOTACAN#
              If !SaldAcre_DOLA < 0 Then
                 !SaldAcre_DOLA = !SALDAcre / TICAPLI#
              End If
              If Left$(CoComCor$, 2) = "RB" Then
                !SaldDebe_DOLA = 0
                !SaldAcre_DOLA = 0
              End If
           End If
        End If
        .Update
        SUDEBE# = SUDEBE# + ROUND#(TOTACAN# * TICAPLI#)
        ATX% = 0
      End If
    End With
    SaDeuPenAcre.Close
    Set SaDeuPenAcre = Nothing
    
    If ATX% > 0 Then
      Call MENSERR(24, "Imposible Completar Transacción\de Imputación de Creditos")
      GoTo 99
    End If
    '
    NRECIB$ = CoComCor$
    RefApli$ = "Imputación Manual " + NRECIB$
    '
    XX% = 0: VACON% = 1
    SUCANCE# = 0
    For U& = 1 To List2.ListCount
      Z$ = List2.List(U& - 1)
      DOCU$ = Left$(Z$, 20)
      If DOCU$ <> "Anticipo no Aplicado" Then
        CANCE# = ROUND#(TICAPLI# * CDbl(XVALO(Mid$(Z$, 70, 14))))
        If CANCE# > 0 Then
          DOCUCA$ = TRIM$(Left$(Z$, 18))
          ICAN# = CANCE#
          TOTAPLI# = 0
          '
          DocOrCor$ = ""
          TotOrig# = 0
          IvaOrig# = 0
          '
          Call APERBASDAT("CABAN")
          strsql = "Select * from SaDeuPen where NuClien=" & NumeClie
          strsql = strsql & " And CODICOM_LAR = '" & DOCUCA$ & "'"
          strsql = strsql & " And SALDDebe > 0"
          Dim SaDeuPenDebe As ADODB.Recordset
          Set SaDeuPenDebe = New ADODB.Recordset
26        On Error Resume Next
          SaDeuPenDebe.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 26
          End If
          On Error GoTo 0
          With SaDeuPenDebe
            '.FindFirst "NuClien = " & NumeClie & " AND CODICOM_LAR = '" & DOCUCA$ & "' AND SALDdebe > 0"
850         'If .NoMatch = False Then
          If Not (.EOF And .BOF) Then
             .MoveFirst
            Do While Not (.EOF)
                 PARCAN# = ICAN#: If !SaldDebe < ICAN# Then PARCAN# = !SaldDebe
                 '.Edit
                 !SaldDebe = !SaldDebe - PARCAN#
                 .Update
                 TOTAPLI# = TOTAPLI# + PARCAN#   ' SUMA APLICACIONES -  agregado por tema bimonetario 22/03
                 '
                 DocOrCor$ = !CodiCom_Cor
                 TotOrig# = !IBruComp
                 IvaOrig# = 0
                 On Error Resume Next
                 IvaOrig# = !IIvaComp
                 On Error GoTo 0
                 '
                 ICAN# = ICAN# - PARCAN#
                 'If ICAN# >= 0.005 Then
                 '  .FindNext "NuClien = " & NumeClie & "AND CODICOM_LAR = '" & DOCUCA$ & "' AND SALDdebe > 0"
                 '  GoTo 850
                 If ICAN# >= 0.005 Then 'me indico flavio rutina de grarecista
                    .MoveNext
                 Else
                    Exit Do
                 End If
              Loop
           End If
          End With
          SaDeuPenDebe.Close
          Set SaDeuPenDebe = Nothing

          '
          DocOrLar$ = TRIM$(Left$(Z$, 18))
          ImpApli# = TOTAPLI#
          SUCANCE# = SUCANCE# + TOTAPLI#
          Call GRAPLICOB
'          If EsFacturaAnticipo%(NumeClie, DocOrLar$) > 0 Then
'             Call GRAANTICIPO(NumeClie, RECILAR$, DocOrLar$, CANCE#)
'             If DifireCuenta%(NumeClie, RECILAR$, DocOrLar$) > 0 Then
'                Call COMUNI("Atención !! Se Modificaran Las Cuentas Contables del " & RECILAR$ & " Según Configuración Modalidad Facturas Anticipo")
'                Call MODICUENTA(NumeClie, RECILAR$, DocOrLar$) 'esta funcion deberia modificar las cuentas segun la configuracion.
'             End If
'          End If
          '
        End If
      End If
    Next U&
    '
    DIFECAN# = SUDEBE# - SUCANCE#
    If DIFECAN# >= 0.005 Then
       If MODOMONEI% = 2 Then
           IAJU% = CUEINT%(CONTROL$("SVD451S", "CUEVE-9"))
          Else
           IAJU% = CUEINT%(CONTROL$("SVD451S", "CUEVE-11"))
       End If
       If IAJU% < 1 Then
          Call MENSERR(24, "Errores al Grabar este Recibo.\Debe Anularlo y Cargarlo de Nuevo.")
          GoTo 99
       End If
       '
       MONEMANT% = MONEMI%
       MONEMI% = 1
       NCX1% = XVALO(Label1)
       ' graba ajuste diferencia de cambio
       RECILA$ = Label25
       CoComCor$ = "AS-" + Right$(RECILA$, 7)
       CoComLar$ = "AS" + Mid$(RECILA$, 3)
       FeContab = CVDAT(Int((FECHANUM(Label24))))
         CUECO% = CUECOCLI%(NCX1%)
       CuentaCo% = CUECO%
       NumeClie = NCX1%
       NumeProv% = 0
       TiMovi = "AJUD"            ' AJUSTE DE SALDO DEUDORES
       OpMovi = "AS"              ' AJUSTE DE SALDO
       VaMovi = "CCORR"           ' CUENTA CORRIENTE
       Descrip = "Aj.Saldo - " + rasocli$(NCX1%)
       FechEmi = FeContab
       FeAcredi = FeContab
       NumeTal = XVALO(Mid$(RECILA$, 6, 4))
'       CodiEmp% = 0
'       CodiSuc% = SUCUI%
       NumeCom = XVALO(Mid$(RECILA$, 11))
       StatInit = 0
       ImpoDeb# = DIFECAN#
       Call GRACABAN
       '
       ' graba contrapartida del ajuste
       CuentaCo% = IAJU%
       TiMovi = "AJUSALD"         ' AJUSTE DE SALDO DEUDORES
       OpMovi = "AS"              ' AJUSTE DE SALDO
       VaMovi = "NO-IMP"          ' NO IMPUTADO
       Descrip = "Aj.Saldo - " + rasocli$(NCX1%)
       ImpoDeb# = (-1) * DIFECAN#
       Call GRACABAN
       '
       MONEMI% = MONEMANT%
       '
    End If
    '
    If MODOMONEI% > 0 Then
      NCI% = NumeClie
      RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCI%))
      If Len(RFF$) > 0 Then
        For U& = 1 To List2.ListCount
          X$ = List2.List(U& - 1)
          CANCE# = CDbl(XVALO(Mid$(X$, 70, 14)))
          If CANCE# >= 0.005 Then
            ABUSQUE$ = TRIM$(Left$(X$, 18))
            For JKLX& = 1 To Len(RFF$) Step 4
              RECABI& = CVS(Mid$(RFF$, JKLX&, 4))
              YCB$ = REGLEIDO$("CARBIMON", RECABI&)
              If TRIM$(Mid$(YCB$, 13, 18)) = ABUSQUE$ Or TRIM$(Mid$(YCB$, 45, 18)) = ABUSQUE$ Then
                DOCUCABE$ = TRIM$(Mid$(YCB$, 13, 18))
                For JKLY& = 1 To Len(RFF$) Step 4
                  RECABI& = CVS(Mid$(RFF$, JKLY&, 4))
                  YCB$ = REGLEIDO$("CARBIMON", RECABI&)
                  If TRIM$(Mid$(YCB$, 13, 18)) = DOCUCABE$ Then
                      MONEMORI% = CVI(Mid$(YCB$, 31, 2))
                      TICAMBX# = TICAMONE(MONEMORI%)
                      If MODOMONEI% = 2 Then
                         TICAMBX# = XVALO(Label16)
                      End If
                      FEDOAPLI% = FECHANUM(Label24)
                      NROO& = XVALO(Mid$(RECILA$, 11))
                      Call REPLA(YCB$, MKI$(FEDOAPLI%), 33, 2)
                      Call REPLA(YCB$, MKI$(NCI%), 35, 2)
                      Call REPLA(YCB$, Chr$(TIDOAP%), 37, 1)
                      Call REPLA(YCB$, Chr$(VADOAP%), 38, 1)
                      Call REPLA(YCB$, MKS$(NROO&), 39, 4)
                      Call REPLA(YCB$, MKI$(SUC%), 43, 2)
                      Call REPLA(YCB$, RECILAR$, 45, 18)
                      Call REPLA(YCB$, MKI$(1), 63, 2) ' RECIBO SIEMPRE EN PESOS
                        If MODOMONEI% = 2 Then
                          Call REPLA(YCB$, MKI$(XVALO(Label32)), 63, 2) ' RECIBO SIEMPRE EN PESOS
                        End If
                      Call REPLA(YCB$, MKD$(0), 65, 8)
                      Call REPLA(YCB$, MKD$(CANCE# * TICAMBX#), 73, 8)
                      Call REPLA(YCB$, MKD$(0), 81, 8)
                      Call REPLA(YCB$, MKD$(TICAMBX#), 89, 8)
                      Call REPLA(YCB$, MKD$(0), 97, 8)
                      Call REPLA(YCB$, MKD$(ROUND#(CANCE#)), 105, 8)
                      Call REPLA(YCB$, MKD$(0), 113, 8)
                        ' ESTO ES PARA QUE SI SE APLICA DOCUMENTO EN PESOS,
                        ' SOLO CANCELE EL SALDO EN PESOS
                        If MODOMONEI% = 2 And MEMI% <= 1 Then
                           Call REPLA(YCB$, String$(32, 0), 89, 32)
                        End If
                      Call REGAPP("CARBIMON", YCB$)
                      '
                      If DIFECAN# >= 0.005 Then
                        Call REPLA(YCB$, "AS" + Mid$(RECILA$, 3), 45, 18)
                        Call REPLA(YCB$, MKI$(1), 63, 2) ' RECIBO SIEMPRE EN PESOS
                        Call REPLA(YCB$, MKD$(DIFECAN#), 65, 8)
                        Call REPLA(YCB$, MKD$(0), 73, 8)
                        Call REPLA(YCB$, MKD$(0), 81, 8)
                        Call REPLA(YCB$, MKD$(TICAMBX#), 89, 8)
                        Call REPLA(YCB$, MKD$(0), 97, 8)
                        Call REPLA(YCB$, MKD$(0), 105, 8)
                        Call REPLA(YCB$, MKD$(0), 113, 8)
                        Call REGAPP("CARBIMON", YCB$)
                      End If
                      GoTo 835
                  End If
                Next JKLY&
              End If
            Next JKLX&
          End If
835     Next U&
      End If
      Call CIERRARCH("CARBIMON")
      '
      ' RECUPERAR EL SALDO EN DOLARES
      For U& = 1 To List2.ListCount
        X$ = List2.List(U& - 1)
        CANCE# = CDbl(XVALO(Mid$(X$, 70, 14)))
        If CANCE# >= 0.005 Then
          ABUSQUE$ = TRIM$(Left$(X$, 18))
          Call RECUSALDOL(ABUSQUE$, NCI%)
        End If
      Next U&
      '
    End If
    '
    
    Call COMUNI("Imputación Completada.")
    Hide
    '
99  VACON% = 0
    Command5.Enabled = True
End Sub

Private Sub Form_Load()
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture1)
    '
'    EPAC$ = TRIM$(EMPREAC$)
'    If EPAC$ <> "" Then
'        Caption = Caption + "  -  " + EPAC$
'    End If
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    '
    TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
    If XVALO(TICAMBIS) < 0.1 Then TICAMBIS = "   1.0000"
    '
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESVENTA", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
    End If
    '
    Call AplicaFuenteAnchoFijo(List2, , 10)
    Call AplicaFuenteAnchoFijo(List3, , 10)
    '
    Call AplicaFuenteAnchoFijo(Label23, , 10)
    Call AplicaFuenteAnchoFijo(Label22, , 10)
    '
    Call AplicaFuenteAnchoFijo(Label4, , 10)
    Call AplicaFuenteAnchoFijo(Label8, , 10)
    Call AplicaFuenteAnchoFijo(Label10, , 10)
    Call AplicaFuenteAnchoFijo(Label14, , 10)
    Call AplicaFuenteAnchoFijo(Label18, , 10)
    Call AplicaFuenteAnchoFijo(Label20, , 10)

End Sub

Private Sub List2_Click()
    '
    If VACON% <> 0 Then
        Exit Sub
    End If
    '
    Z$ = List2.TEXT
    LIST2ANT$ = Z$
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    ACREMAX# = CDbl(XVALO(Label14.Caption)) - CDbl(XVALO(Label18.Caption))
    If ACREMAX# < 0 Then ACREMAX# = 0
    IPEN$ = Mid$(Z$, 56, 14)
    If XVALO(IPEN$) > ACREMAX# Then
        IPEN$ = CSTRING$(MKD$(ACREMAX#), 3, 14, 2, 2)
    End If
    SUMACRE# = XVALO(Label18.Caption)
    If IACRE$ = "" Then
         Call REPLA(Z$, IPEN$, 70, 14)
         SUMACRE# = SUMACRE# + CDbl(XVALO(IPEN$))
       Else
         Call REPLA(Z$, Space$(14), 70, 14)
         SUMACRE# = SUMACRE# - CDbl(XVALO(IACRE$))
    End If
    VACON% = 1
    List2.AddItem Z$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    List2.Refresh
    List2.ListIndex = ILISTA&
    VACON% = 0
    '
    Label10.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    Label10.Refresh
    Label18.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Label18.Refresh
    '
    ANTI# = CDbl(XVALO(Label14.Caption)) - SUMACRE#
    If ANTI# < 0 Then
          Label19.Caption = "Descuentos:"
        Else
          Label19.Caption = "Pendiente:"
    End If
    Label20.Caption = TRIM$(CSTRING$(MKD$(ANTI#), 4, 14, 2, 2))
    Label20.Refresh
    '
End Sub
'

Private Sub List2_DblClick()
    '
    VACON% = 1
    ILIST2& = List2.ListIndex
    TOLIST2& = List2.TopIndex
    '
    Z$ = List2.TEXT
    ILISTA& = List2.ListIndex
    IACRE$ = TRIM$(Mid$(Z$, 70, 14))
    IPEN$ = Mid$(Z$, 56, 14)
    SUMACRE# = XVALO(Label18.Caption)
    SUMACRE# = SUMACRE# - CDbl(XVALO(IACRE$))
    SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(LIST2ANT$, 70, 14)))
    List2.Visible = False
    List2.AddItem LIST2ANT$, ILISTA&
    List2.RemoveItem ILISTA& + 1
    '
    Call IMPUCREVEN
    '
    VACON% = 1
    List2.TopIndex = TOLIST2&
    List2.Visible = True
    If List2.ListCount > ILIST2& Then
          List2.ListIndex = ILIST2&
        Else
          Call CIERRARCH("!CARMECO")
          Call SETULTREG("!CARMECO", 0)
          Call FRESHGRID
          Call BLANARCH("!CARMECO")
          Call CIERRARCH("!APLIRECI")
          Call BLANARCH("!APLIRECI")
          List2.Clear
          Label4.Caption = ""
          Label8.Caption = ""
          Label10.Caption = ""
          Label14.Caption = ""
          Label18.Caption = ""
          Label20.Caption = ""
          Command5.Enabled = True
          Command5.Enabled = False
    End If
    VACON% = 0
    '
End Sub

Private Sub Label29_DblClick()
   TICAMBIS.Enabled = True
   TICAMBIS.SetFocus
End Sub

Private Sub TICAMBIS_Change()
   If TICAMBIS.Enabled = True Then
     TICACH(MONEMI%) = XVALO(TICAMBIS)
   End If
End Sub

Private Sub TICAMBIS_LostFocus()
   If XVALO(TICAMBIS) < 0.1 Then TICAMBIS = "1"
   TICAMBIS = FORMATNUM(XVALO(TICAMBIS), "F9.4")
   TICAMBIS.Enabled = False
End Sub

Sub IMPUCREVEN()
    '
    If List2.ListCount < 1 Then
        Exit Sub
    End If
    '
    Call BLANARCH("!APLIRECI")
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        Z$ = List2.TEXT
        TINUMS$ = Left$(Z$, 20)
        'NROL& = XVALO(Mid$(Z$, 13, 8))
        FEMII% = FECHANUM(Mid$(Z$, 23, 8))
        IPED# = CDbl(XVALO(Mid$(Z$, 31, 14)))
        FEVEI% = FECHANUM(Mid$(Z$, 48, 8))
        SALD# = CDbl(XVALO(Mid$(Z$, 56, 14)))
        ACRE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
        DOCU$ = Mid$(Z$, 97, 18)
        REG2& = XVALO(Mid$(Z$, 121, 6))
        '
        Y$ = REGBLAN$("!APLIRECI")
        Call REPLA(Y$, TINUMS$, 1, 20)
        'Call REPLA(Y$, MKS$(NROL&), 17, 4)
        Call REPLA(Y$, MKI$(FEMII%), 21, 2)
        Call REPLA(Y$, MKD$(IPED#), 23, 8)
        Call REPLA(Y$, MKI$(FEVEI%), 31, 2)
        Call REPLA(Y$, MKD$(SALD#), 33, 8)
        Call REPLA(Y$, MKD$(ACRE#), 41, 8)
        Call REPLA(Y$, DOCU$, 49, 18)
        Call REPLA(Y$, MKS$(REG2&), 67, 4)
        '
        Call REGAPP("!APLIRECI", Y$)
    Next U&
    Call CIERRARCH("!APLIRECI")
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Aplicación de Cobranza"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Picture2.Top + Frame1.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Picture2.Left))
    '
    VTB% = VENTABU%("APLIRECJ" + ATRI$)
    '
    VACON% = 1
    If VTB% >= 0 Then
        List2.Clear
        For U& = 1 To ULTREG&("!APLIRECI")
            Y$ = REGLEIDO$("!APLIRECI", U&)
            TINUMS$ = Left$(Y$, 20)
            'NROL& = CVS(Mid$(Y$, 17, 4))
            FEMI = CVI(Mid$(Y$, 21, 2))
            IPED# = CVD(Mid$(Y$, 23, 8))
            IMPTEX$ = CSTRING$(MKD$(IPED#), 4, 14, 2, 2)
            FEVE = CVI(Mid$(Y$, 31, 8))
            SALD# = CVD(Mid$(Y$, 33, 8))
            SALTEX$ = CSTRING$(MKD$(SALD#), 4, 14, 2, 2)
            ACRE# = CVD(Mid$(Y$, 41, 8))
            '
            If ACRE# < 0 Then
                Call MENSERR(24, "No se Permite Aplicación Negativa")
                GoTo 200
            End If
            If ACRE# > SALD# Then
                Call MENSERR(24, "El Importe Acreditado Debe Ser\Inferior al Saldo del Documento")
                GoTo 200
            End If
            ACRTEX$ = CSTRING$(MKD$(ACRE#), 4, 14, 2, 2)
            DOCU$ = Mid$(Y$, 49, 18)
            REG2& = CVS(Mid$(Y$, 67, 4))
            '
            Z$ = Space$(128)
            Call REPLA(Z$, TINUMS$, 1, 20)
            'Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 13, 8)
            Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
            Call REPLA(Z$, IMPTEX$, 31, 14)
            Call REPLA(Z$, FECHATEX$(FEVE), 48, 8)
            Call REPLA(Z$, SALTEX$, 56, 14)
            Call REPLA(Z$, ACRTEX$, 70, 14)
            Call REPLA(Z$, DOCU$, 97, 18)
            Call REPLA(Z$, TRIM$(Str$(REG2&)), 121, 6)
            '
            List2.AddItem Z$
        Next U&
    End If
    '
    SUMACRE# = 0
    For U& = 1 To List2.ListCount
        List2.ListIndex = U& - 1
        Z$ = List2.TEXT
        SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(Z$, 70, 14)))
    Next U&
    '
    Label10.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    Label10.Refresh
    Label18.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    Label18.Refresh
    ANTI# = CDbl(XVALO(Label14.Caption)) - SUMACRE#
    If ANTI# < 0 Then
          Label19.Caption = "Descuentos:"
        Else
          Label19.Caption = "Pendiente:"
    End If
    Label20.Caption = TRIM$(CSTRING$(MKD$(ANTI#), 4, 14, 2, 2))
    Label20.Refresh
    '
    List2.Refresh
    VACON% = 1
    '
End Sub
'

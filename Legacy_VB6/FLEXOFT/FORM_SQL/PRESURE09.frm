VERSION 5.00
Begin VB.Form PRESURE09 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuestos por Reparaciones"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   13620
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "PRESURE09.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   13620
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame6 
      BackColor       =   &H00CDDADA&
      Caption         =   "CODIGO - MODELO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2445
      Left            =   0
      TabIndex        =   16
      Top             =   75
      Width           =   12705
      Begin VB.Frame Frame9 
         BackColor       =   &H00CDDADA&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   1
         Left            =   165
         TabIndex        =   50
         Top             =   240
         Width           =   5950
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   210
            Width           =   375
         End
         Begin VB.Label Label13 
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
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   1155
            TabIndex        =   54
            Top             =   210
            Width           =   2295
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Técnico:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   -40
            TabIndex        =   53
            Top             =   225
            Width           =   720
         End
         Begin VB.Label Label11 
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
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   4515
            TabIndex        =   52
            Top             =   210
            Width           =   1335
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Talón Nro"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   3720
            TabIndex        =   51
            Top             =   225
            Width           =   720
         End
      End
      Begin VB.Label Label2 
         BackColor       =   &H00CDDADA&
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   6330
         TabIndex        =   27
         Top             =   240
         Width           =   3315
      End
      Begin VB.Label Label20 
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
         Height          =   1005
         Left            =   1005
         TabIndex        =   26
         Top             =   1320
         Width           =   5085
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Equipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   -90
         TabIndex        =   25
         Top             =   1005
         Width           =   1005
      End
      Begin VB.Label Label22 
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
         Left            =   1005
         TabIndex        =   24
         Top             =   945
         Width           =   5085
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Defectos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   -90
         TabIndex        =   23
         Top             =   1350
         Width           =   1005
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Talón Nro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   9525
         TabIndex        =   22
         Top             =   495
         Width           =   1800
      End
      Begin VB.Label Label25 
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
         Left            =   11415
         TabIndex        =   21
         Top             =   435
         Width           =   1140
      End
      Begin VB.Label Label26 
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
         Left            =   8970
         TabIndex        =   20
         Top             =   945
         Width           =   3585
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Serie"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   8040
         TabIndex        =   19
         Top             =   705
         Width           =   1245
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
         Height          =   1005
         Left            =   7710
         TabIndex        =   18
         Top             =   1320
         Width           =   4845
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Accesorios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7095
         TabIndex        =   17
         Top             =   1065
         Width           =   1245
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   12750
      Picture         =   "PRESURE09.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6840
      Width           =   765
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "TRABAJOS A EFECTUAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2800
      Left            =   0
      TabIndex        =   8
      Top             =   5250
      Width           =   12705
      Begin VB.TextBox TXTMOBRAENGTIA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   6795
         TabIndex        =   56
         Top             =   2400
         Width           =   1360
      End
      Begin VB.TextBox Text10 
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
         Height          =   300
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1995
         Width           =   1920
      End
      Begin VB.TextBox PORDESCUMC 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   1995
         Width           =   735
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00CDDADA&
         Caption         =   "TOTAL PRESUPUESTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   240
         TabIndex        =   41
         Top             =   1800
         Width           =   4935
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Presupuesto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1800
            TabIndex        =   45
            Top             =   465
            Width           =   1560
         End
         Begin VB.Label Label6 
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
            Left            =   3435
            TabIndex        =   44
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "IVA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   120
            TabIndex        =   43
            Top             =   460
            Width           =   360
         End
         Begin VB.Label Label5 
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
            Left            =   555
            TabIndex        =   42
            Top             =   375
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "."
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
         Left            =   6000
         TabIndex        =   40
         Top             =   1995
         Width           =   175
      End
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   2415
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5445
         TabIndex        =   37
         Top             =   1995
         Width           =   495
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   1995
         Width           =   1335
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox PORDESCU 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10200
         TabIndex        =   32
         Top             =   1155
         Width           =   735
      End
      Begin VB.TextBox Text22 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1155
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   11160
         TabIndex        =   29
         Top             =   735
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Height          =   300
         Left            =   11160
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   315
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   420
         Width           =   8000
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "M.O. tec. externo (Incluye IVA)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   1
         Left            =   5280
         TabIndex        =   57
         Top             =   2320
         Width           =   1440
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   9540
         TabIndex        =   39
         Top             =   2505
         Width           =   1560
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Medio de Cobranza"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   5445
         TabIndex        =   36
         Top             =   1740
         Width           =   1920
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Subtotal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   9540
         TabIndex        =   34
         Top             =   1650
         Width           =   1560
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dto."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   9675
         TabIndex        =   31
         Top             =   1245
         Width           =   480
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Repuestos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   9540
         TabIndex        =   15
         Top             =   405
         Width           =   1560
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mano de Obra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   9540
         TabIndex        =   9
         Top             =   855
         Width           =   1560
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      Caption         =   "REPUESTOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2655
      Left            =   0
      TabIndex        =   5
      Top             =   2535
      Width           =   12705
      Begin VB.Frame Frame2 
         Caption         =   "                                                                                                                  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Index           =   2
         Left            =   11040
         TabIndex        =   59
         Top             =   600
         Width           =   15
      End
      Begin VB.Frame Frame2 
         Caption         =   "                                                                                                                  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Index           =   1
         Left            =   10120
         TabIndex        =   58
         Top             =   600
         Width           =   15
      End
      Begin VB.Frame Frame8 
         Caption         =   "                                                                                                                  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   6480
         TabIndex        =   14
         Top             =   620
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "                                                                                                                  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   7520
         TabIndex        =   13
         Top             =   620
         Width           =   15
      End
      Begin VB.Frame Frame5 
         Caption         =   "                                                                                                                  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame2 
         Caption         =   "                                                                                                                  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Index           =   0
         Left            =   8820
         TabIndex        =   11
         Top             =   620
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "                                                                                                                  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   2010
         TabIndex        =   10
         Top             =   620
         Width           =   15
      End
      Begin VB.ListBox List1 
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
         Height          =   1830
         ItemData        =   "PRESURE09.frx":0614
         Left            =   75
         List            =   "PRESURE09.frx":0616
         TabIndex        =   6
         Top             =   600
         Width           =   12555
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codigo          Descripción                              Cant.    P.Unit.   Importe  %Desc.    Importe"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   75
         TabIndex        =   7
         Top             =   315
         Width           =   12555
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5715
      Left            =   12750
      ScaleHeight     =   5685
      ScaleWidth      =   1995
      TabIndex        =   3
      Top             =   0
      Width           =   2025
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   48
         Top             =   5000
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   49
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
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
         Height          =   650
         Left            =   105
         Picture         =   "PRESURE09.frx":0618
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   80
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
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
         Height          =   650
         Left            =   105
         Picture         =   "PRESURE09.frx":0762
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   800
         Width           =   650
      End
   End
End
Attribute VB_Name = "PRESURE09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EstadoDeReparacion%
Dim ItemsAgregados%

Public ModoAutomaticoPorImporteMinimo%
Sub ABREVENTANANEW(NUREP&)
    CAMPOS$ = "Codigo/A16;Descripción/A32;Cant./F10.1;P.Unit./F10.2;Importe/F11.2;Desc./f7.2;Neto/f11.2"
    VENTANA.Inicializar = 0
    VENTANA.formula(5) = "3;*;4"
    
    VENTANA.formula(7) = "=ImporteConDescuento#;5;6"
    VENTANA.CampEditables = "4;6"
    VENTANA.ColEcoCampo(1) = "CODIGO;DESCRIPCION;MASTER"
    VENTANA.TOTALES(5) = "S"
    VENTANA.TOTALES(7) = "S"
    VENTANA.DejarLibre(6) = "S"
    VENTANA.AgregaRegistro = 0
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 0
    VENTANA.HabilitaInsertar = 0
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "ITEMS PARA REPARACIÓN"
    VENTANA.Inicializar = 1
    VENTANA.Campox = CAMPOS$



    For i& = 1 To List1.ListCount
       FILA$ = List1.List(i& - 1)
       CI% = CODINT%(Mid$(FILA$, 1, 16))
       CANTIX$ = Mid$(FILA$, 54, 8)
       PRECI1X$ = Mid$(FILA$, 63, 10)
       
       DESC$ = Mid$(FILA$, 85, 6)
       IMPOCONDESC$ = Mid$(FILA$, 92, 12)
       '
       VENTABNEW.MSF.Rows = i& + 1
       VENTABNEW.MSF.TextMatrix(i&, 1) = CODEXT(CI%)
       VENTABNEW.MSF.TextMatrix(i&, 2) = DESCRIT0$(CI%)
       VENTABNEW.MSF.TextMatrix(i&, 3) = CANTIX$
       VENTABNEW.MSF.TextMatrix(i&, 4) = PRECI1X$
       VENTABNEW.MSF.TextMatrix(i&, 5) = FORMATNUM$(XVALO(CANTIX$) * XVALO(PRECI1X$), "F10.2")
       VENTABNEW.MSF.TextMatrix(i&, 6) = DESC$
       VENTABNEW.MSF.TextMatrix(i&, 7) = IMPOCONDESC$
    Next i&
    '
    VENTABNEW.MSF.Rows = i&
    VENTABNEW.Show 1

    If VENTABNEW.APROBADO% > 0 Then
        List1.Clear
        For U& = 1 To VENTABNEW.MSF.Rows - 1
           CODI$ = VENTABNEW.MSF.TextMatrix(U&, 1)
           DESCRIX$ = VENTABNEW.MSF.TextMatrix(U&, 2)
           CANTIX$ = VENTABNEW.MSF.TextMatrix(U&, 3)
           PRECI1X$ = VENTABNEW.MSF.TextMatrix(U&, 4)
             PRECI1# = XVALO(PRECI1X$): CANTI = XVALO(CANTIX$)
           IMPOX$ = FORMATNUM$(XVALO(VENTABNEW.MSF.TextMatrix(U&, 5)), "F10.2")
           DESCU$ = VENTABNEW.MSF.TextMatrix(U&, 6)
           IMPOCONDESC$ = VENTABNEW.MSF.TextMatrix(U&, 7)
           '
           Call REPLA(TTXX$, CODI$, 1, 16)
           Call REPLA(TTXX$, DESCRIX$, 17, 36)
           Call REPLA(TTXX$, CANTIX$, 54, 8)
           Call REPLA(TTXX$, PRECI1X$, 63, 10)
           IMPOR# = PRECI1# * CANTI
           Call REPLA(TTXX$, FORMATNUM(XVALO(IMPOX$), "F11.2"), 73, 11)
           Call REPLA(TTXX$, DESCU$, 85, 6)
           Call REPLA(TTXX$, IMPOCONDESC$, 92, 12)
           RPTOS# = RPTOS# + XVALO(IMPOCONDESC$)
           List1.AddItem TTXX$
           '
        Next U&
        '
        Text1 = FORMATNUM$(RPTOS#, "F12.2")
        Call CALCULATOTAL
   End If
'  Dim obj As New RECORXET
'  Set RS1 = obj.RS_ITEMS_PARA_REPARACIÓN(NUREP&)
'  Call Carga_MSF_Recordset(RS1, CAMPOS$, VENTABNEW.MSF)
'  Set obj = Nothing
'  RS1.Close
'  Set RS1 = Nothing





Set VENTANA = Nothing
End Sub

Sub CALCULATOTAL()
   'PRIMER DESCUENTO
   
   DESC$ = TRIM$(PORDESCU.TEXT)
   '
   TTT1 = (XVALO(Text1) + XVALO(Text4)) * COEFDESCU(DESC$) / 100
   Text22.TEXT = FORMATNUM$(TTT1, "F12.2")
   '
   'ACTUALIZO EL SUBTOTAL PRIMERO
   Text6 = FORMATNUM$(XVALO(Text1) + XVALO(Text4) - TTT1, "F12.2")
   '
   'SEGUNDO DESCUENTO
   DESC$ = TRIM$(PORDESCUMC.TEXT)
     '
   TTT2 = XVALO(Text6) * COEFDESCU(DESC$) / 100
   Text2.TEXT = FORMATNUM$(TTT2, "F12.2")
   Text8 = FORMATNUM$(XVALO(Text6) + TTT2, "F12.2")
   '********************************************************************************
   'NOTA: TTT1 ES UN IMPORTE DE DESCUENTO Y TTT2 ES UN IMPORTE TRATADO COMO RECARGO.
   '********************************************************************************
   
   'ASIGNO IMPORTE A LABEL DE IVA
   PORIVA = ALIVA / 100
   TOTNETO# = XVALO(Text1) + XVALO(Text4) - TTT1 + TTT2
   Label5 = FORMATNUM$((PORIVA * TOTNETO#), "F12.2")
   Label6 = FORMATNUM$(TOTNETO# + XVALO(Label5), "F12.2")
   
End Sub

Sub GENERA_ORREPA(NUREP&, NUMESER$, NUMETEC%, NOMBRETEC$, TALON$, RACLI$)
    FEORCOM% = HOY(HO$)
    ORCOMPRA$ = TALON$
    '
    'TRASLADO DE REPUREPA A RESERVA
    NUMEREPA$ = "OR-" & TRIM$(Str$(NUREP&))
    Call BORRAREGISTROS("RESERVA", "IDSUBOR = '" & NUMEREPA$ & "'", REGAF&, "NOMESS")

    SQLX$ = "SELECT * FROM REPUREPA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NUMREPA = " & NUREP&
    SQLX$ = SQLX$ + " ORDER BY NUORD_ITEM"
    '
    Dim REPUTMP As ADODB.Recordset
    Set REPUTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With REPUTMP
    Do While Not .EOF
         SQLX1$ = "SELECT * FROM RESERVA"
         SQLX1$ = SQLX1$ + " WHERE COD_INTE < 1 "
         Dim RESETMP As ADODB.Recordset
         Set RESETMP = New ADODB.Recordset
         RESETMP.Open FILTSQL$(SQLX1$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
         With RESETMP
           .AddNew   ' YA CONVERTIDO
           On Error Resume Next
           !CodiEmpr = CodiEmp%
           On Error GoTo 0
           CI% = XVALO(REPUTMP!COD_INTE)
           CANTI = XVALO(REPUTMP!Cant_Sol)
           !COD_INTE = CI%
           !COD_EXTE = CODEXT$(CI%)
           !Descrip = Left$(DESCRIT$(CI%), 64) 'modificado por que devolvia un error de datos
           !IDSUBOR = NUMEREPA$
           !FechRes = Now
           !CANTRES = CANTI
           !CANTCONS = 0
           !Saldo_Entre = CANTI
           !IDENLOTE = ""
           !Deposito = 0
           !USOUNIT = CANTI
           !NuOrItem = XVALO(REPUTMP!NuOrd_Item)
           .Update
         End With
         RESETMP.Close
         Set RESETMP = Nothing
    .MoveNext
    Loop
    End With
    
    REPUTMP.Close
    Set REPUTMP = Nothing
    '
    Call ACTUREGISTRO("REPARA", "STATUS", "NUMREPA = " & NUREP&, 3, REGAF&, "NOMESS")
    Call ACTUREGISTRO("REPARA", "FOR_COMPRA", "NUMREPA = " & NUREP&, CVDAT(FEORCOM%), REGAF&, "NOMESS")
    Call ACTUREGISTRO("REPARA", "OR_COMPRA", "NUMREPA = " & NUREP&, ORCOMPRA$, REGAF&, "NOMESS")
    '
20  FILTX$ = TRIM$(Str$(NUREP&))
    Call STDFORM("OREP-MDT", FILTX$, "PRINT")
    
'    FEORCOM% = XVALO(VALOFORM("Text-1"))
'    ORCOMPRA$ = VALOFORM("Text-3")
    If TRIM$(ORCOMPRA$) = "" Or FEORCOM% < 1 Then
        Call COMUNI("No se han Ingresado Fecha y Nº de Orden de Compra")
        GoTo 20
    End If

    CUERPO$ = "Generación de Orden de Reparación" & vbCrLf
    CUERPO$ = "Automática (Talón Técnico: " & TALON$ & ")" & vbCrLf
    CUERPO$ = "Técnico: " & NUMETEC% & " - " & NOMBRETEC$ & vbCrLf
    CUERPO$ = CUERPO$ + "Número de Reparación: " & NUREP& & vbCrLf
    CUERPO$ = CUERPO$ + "Cliente Original: " & RACLI$
    Call ANOTANUSE(NUMESER$, CUERPO$, , "NOMESS")
    
    
    '
End Sub

Sub Command1_Click()
        Command1.Enabled = False
        '
        NUREP& = XVALO(Label25)
        Screen.MousePointer = 11
        'ACTUALIZO EL DETALLE
        '----------------------
        For i& = 1 To List1.ListCount
            FILA$ = List1.List(i& - 1)
            CI% = CODINT%(Mid$(FILA$, 1, 16))
            CANTI = XVALO(Mid$(FILA$, 54, 8))
            IMPUNI# = XVALO(Mid$(FILA$, 63, 10))
            PORDEX# = XVALO(Mid$(FILA$, 85, 6))
            
            SQLX$ = "SELECT PRUN_NETO,IMPO_NETO,PORDESC FROM REPUREPA "
            SQLX$ = SQLX$ + " WHERE NUMREPA = " & NUREP& & " AND NUORD_ITEM = " & i&
            Dim RS1 As ADODB.Recordset
            Set RS1 = New ADODB.Recordset
26          On Error Resume Next
            RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            '
            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 26
            End If
            On Error GoTo 0
            '
            With RS1
                Do While Not .EOF
                   !prun_neto = IMPUNI#
                       IMPNET# = XVALO(CANTI) * IMPUNI#
                   !IMPO_NETO = IMPNET#
                   !PORDESC = ROUND(PORDEX#, 2)
                   .Update
                   .MoveNext
                Loop
            End With
            RS1.Close
            Set RS1 = Nothing

        Next i&
        '----------------------------------
        
        SQLX$ = "SELECT Trab_Efect,REPARACION,STATUS,TOT_REPUESTOS, TOT_MANO_OBRA,DESCTO1,DESCTO2,MED_COBRANZA,MOBRA_EN_GTIA FROM REPARA "
        SQLX$ = SQLX$ + " WHERE NUMREPA = " & NUREP&
        Dim RS2 As ADODB.Recordset
        Set RS2 = New ADODB.Recordset
27      On Error Resume Next
        RS2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        '
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 27
        End If
        On Error GoTo 0
        '
        With RS2
          Do While Not .EOF
             !TRAB_EFECT = Text3
             If TRIM$(SAFETEXT$(!REPARACION)) = "" Then !REPARACION = Text3
              
             If EstadoDeReparacion% < 2 Then RS2!Status = 2
             ' SI PROVIENE DE UNA REPARACION CERRADA (STATUS = 4) Y SE AGREGARON NUEVOS REPUESTOS, HABILITA LA POSIBILIDAD DE TRANSFERIRLOS.
             If EstadoDeReparacion% = 4 And ItemsAgregados% = 1 Then RS2!Status = 2
             !TOT_REPUESTOS = XVALO(Text1)
             !TOT_MANO_OBRA = XVALO(Text4)
             !DESCTO1 = ROUND(XVALO(PORDESCU.TEXT), 2)
             !DESCTO2 = ROUND(XVALO(PORDESCUMC.TEXT), 2)
             !MED_COBRANZA = ROUND(XVALO(Text7.TEXT), 2)
             !MOBRA_EN_GTIA = ROUND(XVALO(TXTMOBRAENGTIA), 2)
             .Update
             .MoveNext
          Loop
         End With
         RS2.Close
         Set RS2 = Nothing
        '-----------------------------------
        NUMESER$ = TRIM$(Label26)
        RACLI$ = VALOBADA("REPARA", "RASO_CLI", "NUMREPA = " & NUREP& & "", "NOMESS")
        CUERPO$ = "Presupuestación de Reparación" & vbCrLf
        CUERPO$ = CUERPO$ + "Número de Reparación: " & NUREP& & vbCrLf
        CUERPO$ = CUERPO$ + "Cliente: " & RACLI$
        Call ANOTANUSE(NUMESER$, CUERPO$, , "NOMESS")
        
        If XVALO(Me.Text5) > 0 Then
         'si tiene tecnico externo esto esta dentro del punto 2
         Call GENERA_ORREPA(XVALO(Label25), NUMESER$, XVALO(Text5), TRIM$(Label13), TRIM$(Label11), RACLI$)
        End If
        
        Screen.MousePointer = 1

        CANCELPRINT% = 0
        TEX& = XVALO(Text5)

'        If TEX& <= 0 Then
            If ModoAutomaticoPorImporteMinimo% = 1 Then GoTo 21
            CONPRIRE.Show 1
21          If CONPRIRE.Check7.Value < 1 Or CONPRIRE.Check2.Value = 1 Or CONPRIRE.Check1.Value = 1 Or ModoAutomaticoPorImporteMinimo% = 1 Then CANCELPRINT% = 1
            Call PRIPRESU(1)    ' IMPRIME PRESUPUESTO INTERNO
               '
            CANCELPRINT% = 0
            If CONPRIRE.Check8.Value < 1 Or CONPRIRE.Check2.Value = 1 Or CONPRIRE.Check1.Value = 1 Or ModoAutomaticoPorImporteMinimo% = 1 Then CANCELPRINT% = 1
            '
            'Envio por Mail
            If CONPRIRE.Check1.Value = 1 Then 'Envio por mail
               CANCELPRINT% = 0
               NUREP& = XVALO(Label25)
               CONDI$ = "NUMREPA= " & NUREP&
               EMAI$ = VALOBADA("REPARA", "MAIL_CLI", CONDI$, "NOMESS")
               If EMAI$ = "" Or InStr(EMAI$, "@") < 1 Then 'TIENE DIRECCION VALIDA
                 Call MENSERR(24, "Imposible Enviar por Mail.\   \Falta Direccion de Email o la Existente es Inválida.")
                 GoTo 25 ' VA IGUAL A IMPRIMIR POR QUE SIEMPRE DEBE IMPRIMIR EN PAPEL O FANTASMA
               End If
               '
               SENDBYMAIL$ = EMAI$
               For Each PRX1 In Printers ' VERIFICA SI EXISTE LA IMPRESORA PDFWRITER
                  If InStr(UCase$(PRX1.DeviceName), "PDFWRITER") > 0 Then GoTo 25
               Next
               Call MENSERR(24, "Imposible Enviar por Mail.\   \Falta Impresora Virtual PDFWRITER.")
               SENDBYMAIL = ""
               GoTo 25
            End If
            '
            'SI ES TECNICO EXTERNO NO IMPRIME

25          Call PRIPRESU(2)
            '
            SENDBYMAIL = ""
            '
            CONPRIRE.Label1 = ""
            CANCELPRINT% = 0
'        End If
        '
        ModoAutomaticoPorImporteMinimo% = 0
    
        Screen.MousePointer = 1
        Call Command2_Click ' SE CIERRA
      
        Command1.Enabled = True

End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Unload Me
 End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Text7 = TablaCobranzaSel(1)
99  Command5.Enabled = True
End Sub

Private Sub Label25_Change()
    NUREP& = XVALO(Label25)
    Text1 = ""
    Text4 = ""
    EstadoDeReparacion% = 0
    TXTMOBRAENGTIA = ""
    
    CONDI$ = "NUMREPA = " & NUREP& & "" ' CARGO DATOS DE ENCABEZADO
'    Label22 = VALOBADA("REPARA", "DESCRIPCION", CONDI$, "NOMESS")
'    Label26 = VALOBADA("REPARA", "NUMEROSERIE", CONDI$, "NOMESS")
'    Label20 = VALOBADA("REPARA", "DEFEC_INFOR", CONDI$, "NOMESS")
'    Label28 = VALOBADA("REPARA", "ACCESORIOS", CONDI$, "NOMESS")
'    Text3 = VALOBADA("REPARA", "TRAB_EFECT", CONDI$, "NOMESS")
    SQLX$ = "SELECT * FROM REPARA WITH(NOLOCK) WHERE " & CONDI$
    Set RS = READSET(SQLX$)
    With RS
        CodigoProductoInterno% = XVALO(!COD_INTE)
        Label22 = SAFETEXT$(!Descripcion)
        Label26 = SAFETEXT$(!NUMEROSERIE)
        Label20 = SAFETEXT$(!DEFEC_INFOR)
        Label28 = SAFETEXT$(!ACCESORIOS)
        Text3 = SAFETEXT$(!TRAB_EFECT)
        PORDESCU.TEXT = ROUND(XVALO(!DESCTO1), 2)
        'MANTENER ESTE ORDEN DE LAS SIGUIENTES 2 FILAS PARA QUE PRIMERO IMPRIMA EL NUMERO DE MEDIO DE COBRANZA Y
        ' ACTUALICE EL DESCUENTO EN PORDESCUMC.TEXT Y LUEGO PONGA EL QUE TRAE DE LO GRABADO
        'DE ESTE MODO CUANDO SE ABRE UN PRESUPUESTO YA GRABADO MANTIENE EL DESCUENTO AUNQUE SE HAYA CAMBIADO EL VALOR DE LA TABLA.
        Text7.TEXT = XVALO(!MED_COBRANZA) ' ESTE PROVOCA EVENTO CHANGE
        PORDESCUMC.TEXT = ROUND(XVALO(!DESCTO2), 2) 'PISA EL DESCUENTO DE LA TABLA POR EL GRABADO LA ULT.VEZ
        Text5 = XVALO(!TEC_EXT)
        Label11.Caption = SAFETEXT$(!TALON_TEC)
        TXTMOBRAENGTIA = TRIM$(FORMATNUM$(XVALO(!MOBRA_EN_GTIA), "F12.2"))
        '*************************************************
        ESTGAR_REP% = XVALO(!Status_Gar)
        EstadoDeReparacion% = XVALO(!Status)
    End With
    RS.Close
    Set RS = Nothing
    '
    List1.Clear
    RPTOS# = 0
    ItemsAgregados% = 0
    '
    If XVALO(Text7) < 1 And EstadoDeReparacion% < 2 Then Text7 = Control$("REPARA", "FOREPAR") ' SI AUN NUNCA SE  HA GRABADO EL PRESUPUESTO PONE EL CONFIGURADO
    SQLX$ = "SELECT * FROM REPUREPA  WITH(NOLOCK)" 'CARGA DE LIST CON LOS REPUESTOS
    SQLX$ = SQLX$ + " WHERE NUMREPA = " & NUREP&
    SQLX$ = SQLX$ + " ORDER BY NUORD_ITEM "
    Dim REPURTMP As ADODB.Recordset
    Set REPURTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With REPURTMP
      Do While Not .EOF
       CI% = XVALO(!COD_INTE)
       CANTI = XVALO(!Cant_Sol)
       PRECI1# = XVALO(!prun_neto)
       IMPOR# = XVALO(!IMPO_NETO)
       DESCU$ = FORMATNUM$(XVALO(!PORDESC), "F6.2")
       MarcaItemAgregado% = !MarcaItemAgregado
       '
       Call REPLA(TTXX$, CODEXT$(CI%), 1, 16)
       Call REPLA(TTXX$, DESCRIT$(CI%), 17, 36)
       Call REPLA(TTXX$, FORMATNUM(CANTI, "F8.1"), 54, 8)
       Call REPLA(TTXX$, FORMATNUM(PRECI1#, "F10.2"), 63, 10)
       Call REPLA(TTXX$, FORMATNUM(IMPOR#, "F11.2"), 73, 11)
       Call REPLA(TTXX$, DESCU$, 85, 6)
       IMPOCONDESC = ImporteConDescuento#(IMPOR#, DESCU$)
       Call REPLA(TTXX$, FORMATNUM(IMPOCONDESC, "F11.2"), 92, 12)
       If EstadoDeReparacion% = 4 And MarcaItemAgregado% = 1 Then ItemsAgregados% = 1
       
       RPTOS# = RPTOS# + XVALO(IMPOCONDESC)
       List1.AddItem TTXX$
       .MoveNext
      Loop
    End With
    REPURTMP.Close
    Set REPURTMP = Nothing
    '
    'CALCULO DE TOTALES
    '
    Text1 = FORMATNUM$(RPTOS#, "F12.2")
    '
    TOTMAOBARA = XVALO(VALOBADA("REPARA", "TOT_MANO_OBRA", CONDI$, "NOMESS"))
    If TOTMAOBARA < 0.005 Then
        ' ** CALCULA EL COSTO DE MANO DE OBRA A PARTIR DE LO CONFIGURADO EN LA LISTA DE COSTOS DE REPARACION DEL SETUP
        If ESTGAR_REP% = 0 Then
            If Control$("REPARA", "MOCOSREP") = "SI" Then
                 CodigoProducto$ = CODEXT$(CodigoProductoInterno%)
                 IDLIST& = ObtenerIdTablaAuxiliar("LISCOSREPA")
                 CostoReparacionManoObra# = XVALO(VALOBADA("DATASQL", "CAMPO3", "ID_LISTA=" & CStr(IDLIST&) & " AND CAMPO1='" & CodigoProducto$ & "'"))
                 If CostoReparacionManoObra# > 0.005 Then
                     Text4 = FORMATNUM(CostoReparacionManoObra#, "F12.2")
                     GoTo 99
                 End If
            End If
            
            ' ** SI A ESTA ALTURA NO PUDO CALCULAR EL COSTO DE M-O, ENTONCES RECURRE EL AL VALOR HORA Y A LAS HORAS ESTIMADAS
            ' ** EN LA PLANILLA DE REPARACION
            CosHora# = XVALO(Control$("REPARA", "VALHORA"))
            If CosHora# < 1 Then CosHora# = 1
            HORAX# = XVALO(VALOBADA("REPARA", "Estima_Horas", "NUMREPA = " & NUREP& & "", "NOMESS"))
            PRSPTO# = ROUND#(HORAX# * CosHora#, 2) ' CAMBIE A ROUND POR QUE EL FORMATNUM NO FUNCIONABA
            Text4 = FORMATNUM(PRSPTO#, "F12.2")
        End If
    Else
       Text4 = FORMATNUM(TOTMAOBARA, "F12.2")
    End If
    '
99

End Sub

Private Sub Label9_DblClick()

    If DERACCE%("MODETREP", "Modificación detalle de Reparación") < 2 Then Exit Sub
    '
    
    'VALIDACION DE ESTADO SOLO SE PUEDE MODFIICAR EL DETALLE SI ES NO ESTA CERRADA.
    NUREP& = XVALO(Label25)
    STAT% = XVALO(VALOBADA("REPARA", "STATUS", "NUMREPA = " & NUREP&, "NOMESS"))
    If STAT% > 3 Then Exit Sub 'SI LA ORDEN YA GENERO CIERRE DE REPARACION, NO SE PUEDE EDITAR
    '******************************************************************************************
    Call ABREVENTANANEW(NUREP&)
    Exit Sub

    For i& = 1 To List1.ListCount
       FILA$ = List1.List(i& - 1)
       CI% = CODINT%(Mid$(FILA$, 1, 16))
       CANTI = XVALO(Mid$(FILA$, 54, 8))
       
       PRECI1# = XVALO(Mid$(FILA$, 63, 10))
       '
       Call REPLA(XX$, MKI$(CI%), 1, 2)
       Call REPLA(XX$, MKS$(CANTI), 5, 4)
       Call REPLA(XX$, MKD$(PRECI1#), 9, 8)
       Call GRAREG("!MATEROF", XX$, i&)
       '
    Next i&
    '
    Call SETULTREG("!MATEROF", i& - 1)
    Call CIERRARCH("!MATEROF")
    VTB% = VENTABU%("EDITREPA/TITUPAN=DETALLE DE REPARACIÓN")
    If VTB% < 0 Then Exit Sub
    '
    RPTOS# = 0
    List1.Clear
    For U& = 1 To ULTREG&("!MATEROF")
       ZZ$ = REGLEIDO$("!MATEROF", U&)
       TTXX$ = Space$(128)
       CI% = CVI(Left$(ZZ$, 2))
       CANTI = CVS(Mid$(ZZ$, 5, 4))
       PRECI1# = CVD(Mid$(ZZ$, 9, 8))
       '
       Call REPLA(TTXX$, CODEXT$(CI%), 1, 16)
       Call REPLA(TTXX$, DESCRIT$(CI%), 17, 36)
       Call REPLA(TTXX$, FORMATNUM(CANTI, "F8.1"), 54, 8)
       Call REPLA(TTXX$, FORMATNUM(PRECI1#, "F10.2"), 63, 10)
       IMPOR# = PRECI1# * CANTI
       Call REPLA(TTXX$, FORMATNUM(IMPOR#, "F11.2"), 73, 11)
       RPTOS# = RPTOS# + IMPOR#
       List1.AddItem TTXX$
       '
    Next U&
    '
    Text1 = FORMATNUM$(RPTOS#, "F12.2") '    Call CALCULATOTAL
    '
End Sub

Private Sub List1_DblClick()
   If List1.ListCount > 0 Then Call Label9_DblClick
End Sub


Private Sub PORDESCU_Change()

   Call CALCULATOTAL
End Sub

Private Sub PORDESCUMC_Change()
   Call CALCULATOTAL
End Sub

Private Sub Text1_Change()
   Call CALCULATOTAL
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text4.SetFocus
   End If
End Sub

Private Sub Text1_LostFocus()
   Text1 = FORMATNUM$(Text1, "F12.2")

End Sub


Sub PRIPRESU(OPCI%)
    '
    PORIVA# = ALIVA
    SUBTOT# = XVALO(Text1) + XVALO(Text4)
    If OPCI% = 2 Then GoTo 12
    
    FORIPRE$ = TRIM$(Control$("FOREPAR", "DOCUREP1"))
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Imposible Imprimir Presupuesto de Reparación.\Falta Definir Formulario Electrónico.")
       Exit Sub
    End If
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Presupuesto de Reparación.\No existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
10  NUREP& = XVALO(Label25)
    NCLI = XVALO(VALOBADA("REPARA", "Nume_cli", "NUMREPA = " & NUREP& & "", "NOMESS"))
    TIPODOC$ = "Presupuesto de Reparación"
    DESTIDOCNUE = NCLI
    NUMEDOC$ = TRIM$(Label25)
    FE% = HOY(HO$)
    FECHDOC$ = FECHATEX$(FE%) 'TRIM$(Text9.TEXT)
    CODI$ = ""
    CI% = XVALO(VALOBADA("REPARA", "COD_INTE", "NUMEROSERIE = '" & Label26 & "'", "MOMESS"))
    If CI% > 0 Then
       CODI$ = CODEXT$(CI%)
    End If
    '
    CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = CODI$
    CODPARAM$(2) = "DES-CONJ": DOCPARAM$(2) = TRIM$(Label22)
    CODPARAM$(3) = "REF-MAT1": DOCPARAM$(3) = TRIM$(Label26)
    CODPARAM$(4) = "MEDIDA-1": DOCPARAM$(4) = "" 'TRIM$(Text10)
      HORAX# = XVALO(VALOBADA("REPARA", "Estima_Horas", "NUMREPA = " & NUREP& & "", "NOMESS"))
      ESTIHS$ = FORMATNUM$(HORAX#, "F5.2")
    CODPARAM$(5) = "REF-MAT2": DOCPARAM$(5) = ESTIHS$
      CosHor# = XVALO(Control$("REPARA", "VALHORA"))
      VALHOR$ = FORMATNUM$(CosHor#, "F10.2")
    CODPARAM$(6) = "REF-MAT3": DOCPARAM$(6) = VALHOR$
    CODPARAM$(7) = "PRE-NETO": DOCPARAM$(7) = TRIM$(Text1)
    CODPARAM$(8) = "BRU-ITEM": DOCPARAM$(8) = TRIM$(Text4)
    '
'    ' ***  CALCULO DE IVA
'    APLIVA% = 1
'    If NCLIE% > 0 Then
'       PVCLI% = PIVACLI%(NCLIE%)
'       If PVCLI% = 2 Then RNI% = 1
'       If PVCLI% = 5 Or PVCLI% = 7 Then APLIVA% = 0
'     End If
'     '
'     IMPIVA21 = TRIM$(FORMATNUM$(APLIVA% * XVALO(TEXT4) * (PORCEIVA(0) / 100), "F12.2"))
'     IMPIVA105 = ""
'     For U& = 1 To List1.ListCount - 1
'       ZZ$ = List1.List(U&)
'       CIXXI% = CODINT(Left$(ZZ$, 16))
'       CANTU = XVALO(Mid$(ZZ$, 54, 8))
'       PRUNI = XVALO(Mid$(ZZ$, 63, 10))
'       '
'       TOT# = ROUND(CANTU * PRUNI)
'       SUIVA# = SUIVA# + (Int(TOT# * PORCEIVA(CIXXI%) + 0.5)) / 100
'       '
'       If GRUPIVA%(CIXXI%) < 1 Then
'             IMPOIVA# = APLIVA% * (XVALO(IMPIVA21) + ROUND#(TOT# * (PORCEIVA(CIXXI%) / 100) * (1 - COEFDESCU(DESC$) / 100), 2))
'             IMPIVA21 = TRIM$(FORMATNUM$(IMPOIVA#, "F12.2"))
'           Else
'             IMPOIVA# = APLIVA% * (XVALO(IMPIVA105) + ROUND#(TOT# * (PORCEIVA(CIXXI%) / 100) * (1 - COEFDESCU(DESC$) / 100), 2))
'             IMPIVA105 = TRIM$(FORMATNUM$(IMPOIVA#, "F12.2"))
'       End If
'     Next U&
'     '
'     SUIVA# = XVALO(IMPIVA21) + XVALO(IMPIVA105)
'     SUIVARNI# = SUIVA# * RNI% / 2
'     TOTIVA = TRIM$(CSTRING$(MKD$(SUIVA# + SUIVARNI#), 4, 12, 2, 2))
'     '
'     ' *****

     '
    CODPARAM$(9) = "IMP-NETO": DOCPARAM$(9) = TRIM$(FORMATNUM$(SUBTOT#, "F12.2"))
    'TOTIVA = FORMATNUM$(XVALO(Label6) * PORIVA# / 100, "F12.2")
    CODPARAM$(10) = "IMP-IVA": DOCPARAM$(10) = TRIM$(Label5)
     IMPOTOTAL# = XVALO(Label6) 'SUBTOT# + XVALO(Label5)
    CODPARAM$(11) = "IMP-TOTA": DOCPARAM$(11) = TRIM$(FORMATNUM$(IMPOTOTAL#, "F12.2"))
    
    PORI$ = ""
    On Error Resume Next
    'PORI$ = TRIM$(Str$(100 * TOTAL(1) / TOTAL(0)))
    On Error GoTo 0
    CODPARAM$(12) = "POR-IVA": DOCPARAM$(12) = "" 'TRIM$(PORI$)
    CODPARAM$(13) = "MEDIDA-3": DOCPARAM$(13) = TRIM$(Label25)
    CODPARAM$(14) = "PERC-IVA": DOCPARAM$(14) = "" 'TRIM$(TOTPER)
    '
    
    'DIFER# = XVALO(Label6) - SUBTOT#
    '
    RACLIENTE$ = VALOBADA("REPARA", "RASO_CLI", "NUMREPA = " & NUREP& & "")
    'CODPARAM$(15) = "IMPNEDES": DOCPARAM$(15) = FORMATNUM$(SUBTOT#, "F12.2") 'TRIM$(Label6)
    CODPARAM$(15) = "IMPNEDES": DOCPARAM$(15) = FORMATNUM$(XVALO(Text6.TEXT), "F12.2") 'TRIM$(Label6)
    CODPARAM$(16) = "RASO-CLI": DOCPARAM$(16) = RACLIENTE$
    
    CODPARAM$(17) = "ITO-GRA2": DOCPARAM$(17) = FORMATNUM$(PORDESCU.TEXT, "F6.2")
    CODPARAM$(18) = "PORDESCU": DOCPARAM$(18) = FORMATNUM$(XVALO(Text22), "F12.2")
    CODPARAM$(19) = "ITO-GRA1": DOCPARAM$(19) = FORMATNUM$(XVALO(PORDESCUMC.TEXT), "F12.2")
    CODPARAM$(20) = "IMP-DESC": DOCPARAM$(20) = FORMATNUM$(XVALO(Text2.TEXT), "F12.2")
    CODPARAM$(21) = "SUB-TOTA": DOCPARAM$(21) = FORMATNUM$(XVALO(Text8.TEXT), "F12.2")
    CAPARDOC% = 21
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CODTABU1$(4) = "PRE-LIST"
    CODTABU1$(5) = "NET-ITEM"
    
    CACOLTAB1% = 5
    '
    CAITAB1% = 0
    For U& = 1 To List1.ListCount
      ZZ$ = List1.List(U& - 1)
      CAITAB1% = CAITAB1% + 1
      TETABU1$(1, CAITAB1%) = Left$(ZZ$, 16)
      CII% = CODINT%(Left$(ZZ$, 16))
      TETABU1$(2, CAITAB1%) = DESCRIT$(CII%)
      TETABU1$(3, CAITAB1%) = TRIM$(Mid$(ZZ$, 54, 8))
      IMPOX# = XVALO(Mid$(ZZ$, 63, 10))
      DESCU$ = Mid$(ZZ$, 85, 6)
      If XVALO(DESCU$) > 0 Then IMPOX# = ImporteConDescuento#(IMPOX#, DESCU$)
      TETABU1$(4, CAITAB1%) = TRIM$(FORMATNUM$(IMPOX#, "F12.2"))
      'TETABU1$(5, CAITAB1%) = TRIM$(Mid$(ZZ$, 73, 11))
      TETABU1$(5, CAITAB1%) = TRIM$(Mid$(ZZ$, 92, 12))
    Next U&
    '
    Call SETULTREG("!OBSERVOF", 0)
    Call FRATEXTO(Text3.TEXT, 50)
    For YX& = 1 To CARETEX%
       Call REGAPP("!OBSERVOF", RETEX$(YX&))
    Next YX&
    '
    Call PRINTDOC("@" + FORIPRE$)
    '
    Exit Sub
    
12  FORIPRE$ = TRIM$(Control$("FOREPAR", "DOCUREP3"))
    If FORIPRE$ = "" Then
       Exit Sub
    End If
    '
    If Control("REPARA", "PUERTOIM") <> "" Then PUERREPA$ = Control("REPARA", "PUERTOIM")

    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       Call MENSERR(24, "Imposible Imprimir Presupuesto de Reparación.\No existe Formulario Electrónico '" + FORIPRE$ + ".FRM'.")
       Exit Sub
    End If
    '
    NUREP& = XVALO(Label25)
    VDEF$ = Space$(96) + Space$(48)
    Call REPLA(VDEF$, "5 Días hábiles de recibida su órden de compra.", 1, 48)
    Call REPLA(VDEF$, "Contado contra entrega", 49, 48)
        ATENTX$ = VALOBADA("REPARA", "Ctac_Cli", "NUMREPA = " & NUREP& & "", "NOMESS")
    Call REPLA(VDEF$, ATENTX$, 97, 48)
15  If ModoAutomaticoPorImporteMinimo% = 1 Then
        OBX$ = VDEF$
        GoTo 17
    End If
    
    OBX$ = OBJPLA$("PRESUREPA", VDEF$)
17  If OBX$ = "" Then GoTo 15
    PLAENTRE$ = Left$(OBX$, 48)
    CONDIPAG$ = Mid$(OBX$, 49, 48)
    ATENTX$ = Mid$(OBX$, 97, 48)
    If TRIM$(PLAENTRE$) = "" Or TRIM$(CONDIPAG$) = "" Then
      Call MENSERR(24, "Datos Incompletos")
      VDEF$ = OBX$
      GoTo 15
    End If
    '
20  NUREP& = XVALO(Label25)
    NCLI = XVALO(VALOBADA("REPARA", "Nume_cli", "NUMREPA = " & NUREP& & "", "NOMESS"))
    TIPODOC$ = "Presup.Reparación Externo"
    DESTIDOCNUE = NCLI
    NUMEDOC$ = TRIM$(Label25)
    FE% = HOY(HO$)
    FECHDOC$ = FECHATEX$(FE%) 'TRIM$(Text9.TEXT)
    '
    CODI$ = ""
    CI% = XVALO(VALOBADA("REPARA", "COD_INTE", "NUMEROSERIE = '" & Label26 & "'", "MOMESS"))
    If CI% > 0 Then
       CODI$ = CODEXT$(CI%)
    End If
    '
    CODPARAM$(1) = "COD-CONJ": DOCPARAM$(1) = CODI$ 'TRIM$(Label22)
    CODPARAM$(2) = "DES-CONJ": DOCPARAM$(2) = TRIM$(Label22) 'DESCRIT$(CI%)
    CODPARAM$(3) = "REF-MAT1": DOCPARAM$(3) = TRIM$(Label26)
    CODPARAM$(4) = "MEDIDA-1": DOCPARAM$(4) = "" 'TRIM$(Text10)
    CODPARAM$(5) = "REF-MAT2": DOCPARAM$(5) = ""
    CODPARAM$(6) = "REF-MAT3": DOCPARAM$(6) = ATENTX$ '""
    CODPARAM$(7) = "PRE-NETO": DOCPARAM$(7) = "" 'TRIM$(TEXT1)
    CODPARAM$(8) = "BRU-ITEM": DOCPARAM$(8) = "" 'TRIM$(TEXT4)
    'CODPARAM$(9) = "IMP-NETO": DOCPARAM$(9) = FORMATNUM$(XVALO(Text1) + XVALO(Text4), "F12.2") ' TRIM$(Label6)
    CODPARAM$(9) = "IMP-NETO": DOCPARAM$(9) = FORMATNUM$(XVALO(Text8.TEXT), "F12.2") ' TRIM$(Label6)
    CODPARAM$(10) = "IMP-IVA": DOCPARAM$(10) = TRIM$(Label5)
'    TOTALCONIVA# = XVALO(Label5) + SUBTOT#
'    CODPARAM$(11) = "IMP-TOTA": DOCPARAM$(11) = TRIM$(FORMATNUM$(TOTALCONIVA#, "F10.2")) 'TRIM$(TOTAL(2))
    CODPARAM$(11) = "IMP-TOTA": DOCPARAM$(11) = TRIM$(FORMATNUM$(XVALO(Label6), "F10.2")) 'TRIM$(TOTAL(2))
    PORI$ = ""
    On Error Resume Next
    'PORI$ = TRIM$(Str$(100 * TOTAL(1) / TOTAL(0)))
    On Error GoTo 0
    CODPARAM$(12) = "POR-IVA": DOCPARAM$(12) = "" 'TRIM$(PORI$)
    CODPARAM$(13) = "MEDIDA-3": DOCPARAM$(13) = TRIM$(Label25)
    CODPARAM$(14) = "REF-MAT2": DOCPARAM$(14) = TRIM$(PLAENTRE$)
    CODPARAM$(15) = "COND-PAG": DOCPARAM$(15) = TRIM$(CONDIPAG$)
    CODPARAM$(16) = "PERC-IVA": DOCPARAM$(16) = "" 'TRIM$(TOTPER)
    CODPARAM$(17) = "RASO-CLI": DOCPARAM$(17) = RACLIENTE$
    CODPARAM$(18) = "REF-MAT4": DOCPARAM$(18) = Text3
    If XVALO(PORDESCU) > 0 Then ' SI TIENE DESCUENTO SE ARMA UN TEXTO INFORMANDO EL MISMO.
      TXX1$ = " Descuento Aplicado: $ " & TRIM$(FORMATNUM$(XVALO(Text22.TEXT), "F12.2"))
      TXX1$ = TXX1$ + " Sobre Valor Original Neto del Presupuesto "
        CODPARAM$(19) = "OBS-OPER": DOCPARAM$(19) = TXX1$
    End If
    CAPARDOC% = 19
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CACOLTAB1% = 3
    '
    CAITAB1% = 0
    For U& = 1 To List1.ListCount
      ZZ$ = List1.List(U& - 1)
      CAITAB1% = CAITAB1% + 1
      TETABU1$(1, CAITAB1%) = Left$(ZZ$, 16)
      CII% = CODINT%(Left$(ZZ$, 16))
      TETABU1$(2, CAITAB1%) = DESCRIT$(CII%)
      TETABU1$(3, CAITAB1%) = TRIM$(Mid$(ZZ$, 54, 8))
    Next U&
    '
    If PUERREPA$ <> "" Then Call SETPRINPORT(PUERREPA$) ' TOMA LA IMPRESORA CONFIGURADA
    Call PRINTDOC("@" + FORIPRE$) ' REPARACION
    If PUERREPA$ <> "" Then Call SETPRINPORT("RESTORE")
    
    '
End Sub

Private Sub Text4_Change()
   Call CALCULATOTAL
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Text3.SetFocus
    End If
End Sub

Private Sub Text4_LostFocus()
   Text4 = FORMATNUM$(Text4, "F12.2")
End Sub

Private Sub Text5_Change()
   If XVALO(Text5) > 0 Then
     Frame9(1).Visible = True
     Label13 = NOMBRE_TEC_EXT$(XVALO(Text5))
     Label19(1).Visible = True
     TXTMOBRAENGTIA.Visible = True
     
   Else
     Frame9(1).Visible = False
     Label19(1).Visible = False
     TXTMOBRAENGTIA.Visible = False
   End If
End Sub

Private Sub Text7_Change()
   Text10 = VALOBADA("DATASQL", "CAMPO2", "LISTA = 'MEDIOS COBRANZA-MDT' AND CAMPO1 = '" & TRIM$(FORMATOCON0$(XVALO(Text7), 7)) & "'")
   PORDESCUMC = FORMATNUM$(XVALO(VALOBADA("DATASQL", "CAMPO3", "LISTA = 'MEDIOS COBRANZA-MDT' AND CAMPO1 = '" & TRIM$(FORMATOCON0$(XVALO(Text7), 7)) & "'")), "F6.2")
End Sub
Function TablaCobranzaSel(COLUMNA%)
    'COLUMNA% ES LA QUE VA A DEVOLVER EN LA FUNCION TablaCobranzaSel
    If IsFormLoaded(FRMZELECHO) Then
        ESTAENMEMORIA = 1
    Else
        Call HookForm(FRMZELECHO.msf2) 'estaba tirando un error cuando se cierra el formulario con doble clic
        Call HookForm(FRMZELECHO.MSF1) 'probablemente por que hace un hide, SI NO ESTA EN MEMORIA SE APLICA LA FUNCION.
    End If

    CAMPOS$ = "Código/f6;Referencia/A32;Porcentaje/F10.2"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
    FRMZELECHO.Caption = "TABLA DE MEDIOS DE COBRANZA"
    
    Dim obj As New RECORXET
    Set RS1 = obj.RS_MEDIOSCOBRANZA_MDT()
    Call Carga_MSF_Recordset(RS1, CAMPOS$, FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 10000
    FRMZELECHO.Show 1
    TablaCobranzaSel = TRIM$(RESP_ZELE_VECT(COLUMNA%))

End Function

Private Sub TXTMOBRAENGTIA_LostFocus()
    TXTMOBRAENGTIA.TEXT = FORMATNUM$(XVALO(TXTMOBRAENGTIA.TEXT), "F12.2")
End Sub

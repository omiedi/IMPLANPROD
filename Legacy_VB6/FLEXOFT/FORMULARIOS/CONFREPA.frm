VERSION 5.00
Begin VB.Form CONFREPA 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reparaciones - Configuración de Funcionamiento"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10950
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
   ScaleHeight     =   7935
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Precios de Repuestos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   210
      TabIndex        =   58
      Top             =   4920
      Width           =   10560
      Begin VB.CommandButton Command5 
         Height          =   520
         Left            =   9680
         Picture         =   "CONFREPA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Acceso a Lista de Precios"
         Top             =   240
         Width           =   540
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
         Left            =   4275
         TabIndex        =   61
         Top             =   480
         Width           =   175
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Tabla de Costos Calculados"
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
         Left            =   360
         TabIndex        =   59
         Top             =   320
         Width           =   1815
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
         Left            =   4680
         TabIndex        =   63
         Top             =   480
         Width           =   4850
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   3840
         TabIndex        =   62
         Top             =   480
         Width           =   450
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Lista de Precios Nro:"
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
         Left            =   2640
         TabIndex        =   60
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ingreso de Equipos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   49
      Top             =   210
      Width           =   5160
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Index           =   5
         Left            =   105
         TabIndex        =   50
         Top             =   315
         Width           =   4950
         Begin VB.TextBox FORMUNAME 
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
            Left            =   2625
            TabIndex        =   53
            Text            =   " "
            Top             =   105
            Width           =   1275
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Index           =   5
            Left            =   3885
            TabIndex        =   52
            Top             =   105
            Width           =   175
         End
         Begin VB.CommandButton BOTOPENFORM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   620
            Index           =   5
            Left            =   4200
            Picture         =   "CONFREPA.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   105
            Width           =   540
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nota de Recepción:"
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
            Left            =   210
            TabIndex        =   55
            Top             =   160
            Width           =   2325
         End
         Begin VB.Label FORMUECO 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   210
            TabIndex        =   54
            Top             =   430
            Width           =   3855
         End
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Código Genérico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5565
      TabIndex        =   48
      Top             =   1575
      Width           =   4215
      Begin VB.CommandButton Command3 
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
         Left            =   3150
         TabIndex        =   57
         Top             =   420
         Width           =   175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   1155
         TabIndex        =   56
         Top             =   420
         Width           =   2010
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Reparaciones Típicas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   5565
      TabIndex        =   44
      Top             =   210
      Width           =   4215
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   3000
         TabIndex        =   45
         Top             =   420
         Width           =   3060
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
            Left            =   0
            Picture         =   "CONFREPA.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Reparaciones Normalizadas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Actualizar B.Datos de Reparaciones Normales"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   47
            Top             =   105
            Width           =   2220
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Condiciones Comerciales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3045
      TabIndex        =   39
      Top             =   6030
      Width           =   7725
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
         Left            =   630
         TabIndex        =   41
         Top             =   580
         Width           =   6840
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
         Left            =   630
         TabIndex        =   40
         Top             =   1210
         Width           =   6840
      End
      Begin VB.Label FORMUDES 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Plazo de Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   8
         Left            =   -630
         TabIndex        =   43
         Top             =   315
         Width           =   2325
      End
      Begin VB.Label FORMUDES 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pago:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   7
         Left            =   0
         TabIndex        =   42
         Top             =   945
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Orden de Reparación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   5565
      TabIndex        =   26
      Top             =   2520
      Width           =   5205
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Index           =   0
         Left            =   105
         TabIndex        =   33
         Top             =   1260
         Width           =   4950
         Begin VB.TextBox FORMUNAME 
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
            Left            =   2625
            TabIndex        =   36
            Text            =   " "
            Top             =   105
            Width           =   1275
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Left            =   3885
            TabIndex        =   35
            Top             =   105
            Width           =   175
         End
         Begin VB.CommandButton BOTOPENFORM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   620
            Index           =   0
            Left            =   4200
            Picture         =   "CONFREPA.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   105
            Width           =   540
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hoja de Procedimientos:"
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
            Left            =   210
            TabIndex        =   38
            Top             =   160
            Width           =   2325
         End
         Begin VB.Label FORMUECO 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   210
            TabIndex        =   37
            Top             =   430
            Width           =   3855
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Index           =   2
         Left            =   105
         TabIndex        =   27
         Top             =   315
         Width           =   4950
         Begin VB.CommandButton BOTOPENFORM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   620
            Index           =   2
            Left            =   4200
            Picture         =   "CONFREPA.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   105
            Width           =   540
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Left            =   3885
            TabIndex        =   29
            Top             =   105
            Width           =   175
         End
         Begin VB.TextBox FORMUNAME 
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
            Left            =   2625
            TabIndex        =   28
            Text            =   " "
            Top             =   105
            Width           =   1275
         End
         Begin VB.Label FORMUECO 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   210
            TabIndex        =   32
            Top             =   430
            Width           =   3855
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Formulario de O/Rep:"
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
            Left            =   105
            TabIndex        =   31
            Top             =   165
            Width           =   2430
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Parámetros de Cálculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   210
      TabIndex        =   15
      Top             =   6030
      Width           =   2685
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
         Left            =   1365
         TabIndex        =   19
         Top             =   1210
         Width           =   1065
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
         Left            =   1365
         TabIndex        =   17
         Top             =   580
         Width           =   1065
      End
      Begin VB.Label FORMUDES 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Coefic. Presupuesto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   5
         Left            =   105
         TabIndex        =   18
         Top             =   1080
         Width           =   1170
      End
      Begin VB.Label FORMUDES 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C-H Mano de Obra:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Index           =   4
         Left            =   210
         TabIndex        =   16
         Top             =   470
         Width           =   1065
      End
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
      Height          =   750
      Left            =   9975
      Picture         =   "CONFREPA.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   315
      Width           =   800
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
      Height          =   750
      Left            =   9975
      Picture         =   "CONFREPA.frx":107C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1090
      Width           =   800
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Formularios Cotización"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   210
      TabIndex        =   2
      Top             =   1575
      Width           =   5205
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Index           =   4
         Left            =   105
         TabIndex        =   20
         Top             =   2205
         Width           =   4950
         Begin VB.TextBox FORMUNAME 
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
            Left            =   2625
            TabIndex        =   23
            Text            =   " "
            Top             =   105
            Width           =   1275
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Left            =   3885
            TabIndex        =   22
            Top             =   105
            Width           =   175
         End
         Begin VB.CommandButton BOTOPENFORM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   620
            Index           =   4
            Left            =   4200
            Picture         =   "CONFREPA.frx":1386
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   105
            Width           =   540
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Secuencia Operaciones:"
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
            Left            =   210
            TabIndex        =   25
            Top             =   160
            Width           =   2325
         End
         Begin VB.Label FORMUECO 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   210
            TabIndex        =   24
            Top             =   430
            Width           =   3855
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Index           =   3
         Left            =   105
         TabIndex        =   9
         Top             =   1260
         Width           =   4950
         Begin VB.CommandButton BOTOPENFORM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   620
            Index           =   3
            Left            =   4200
            Picture         =   "CONFREPA.frx":1690
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   105
            Width           =   540
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Left            =   3885
            TabIndex        =   11
            Top             =   105
            Width           =   175
         End
         Begin VB.TextBox FORMUNAME 
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
            Left            =   2625
            TabIndex        =   10
            Text            =   " "
            Top             =   105
            Width           =   1275
         End
         Begin VB.Label FORMUECO 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   210
            TabIndex        =   14
            Top             =   430
            Width           =   3855
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Presupuesto Cliente:"
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
            Left            =   105
            TabIndex        =   13
            Top             =   165
            Width           =   2430
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Index           =   1
         Left            =   105
         TabIndex        =   3
         Top             =   315
         Width           =   4950
         Begin VB.TextBox FORMUNAME 
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
            Left            =   2625
            TabIndex        =   6
            Text            =   " "
            Top             =   105
            Width           =   1275
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Left            =   3885
            TabIndex        =   5
            Top             =   105
            Width           =   175
         End
         Begin VB.CommandButton BOTOPENFORM 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   620
            Index           =   1
            Left            =   4200
            Picture         =   "CONFREPA.frx":199A
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   105
            Width           =   540
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Presupuesto Interno:"
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
            Left            =   210
            TabIndex        =   8
            Top             =   160
            Width           =   2325
         End
         Begin VB.Label FORMUECO 
            BackColor       =   &H80000009&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   210
            TabIndex        =   7
            Top             =   430
            Width           =   3855
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   9370
      Picture         =   "CONFREPA.frx":1CA4
      Top             =   1910
      Width           =   2010
   End
End
Attribute VB_Name = "CONFREPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTOPENFORM_Click(Index As Integer)
    Call CONECRUN("FLFORMS/" + TRIM$(FORMUNAME(Index).TEXT), "Formularios de Impresion")
    If EXISTE%("FLFORMS.EXE") < 1 Then
        Call FINAL("")
    End If
End Sub

Private Sub BOTSELFORM_Click(Index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        FORMUNAME(Index).TEXT = VDEVU$
        FORMUECO(Index).Caption = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Check1_Click()
   If Check1.Value = 1 Then
      Label3 = "0"
   End If
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    If CODINT%(Label1) < 1 Then
       Call MENSERR(24, "Falta Código Genérico de Reparaciones")
       Exit Sub
    End If
    '
    For KK% = 0 To 5
       On Error Resume Next
       Call GRACONTROL("FOREPAR", "DOCUREP" + TRIM$(Str$(KK%)), FORMUNAME(KK%).TEXT)
    Next KK%
    '
    Call GRACONTROL("REPARA", "VALHORA", Text1.TEXT)
    Call GRACONTROL("REPARA", "COEFPREA", Text2.TEXT)
    Call GRACONTROL("REPARA", "CODGENER", Label1)
    Call GRACONTROL("REPARA", "LIPREREP", Label3)
    '
    Call GRACONTROL$("", "PLAENREP", Text4)
    Call GRACONTROL$("", "CONPAREP", Text3)
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call SELECHO("MASTER")
   VDEVU$ = TRIM$(VALACT1$("MASTER"))
   If VDEVU$ <> "" Then
     Label1 = VDEVU$
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Call SELECHO("LISTAS")
    VDEVU$ = VALACT1$("LISTAS")
    If VDEVU$ <> "" Then
        Label3 = VDEVU$
        Check1.Value = 0
    End If
End Sub

Sub Command5_Click()
   Command5.Enabled = False
   If Val(Label3) > 0 Then
      Call CONECRUN("LISPRE02/" + TRIM$(Label3), "Listas de Precios de Venta")
   End If
   Command5.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   On Error Resume Next
   REPABOMB.Show 1
   If Err Then
      Call MENSERR(24, "No Disponible en la Presente Versión")
   End If
   Command8.Enabled = True
End Sub

Private Sub Form_Load()
    '
    For KK% = 0 To 5
       On Error Resume Next
       FORMUNAME(KK%).TEXT = CONTROL$("FOREPAR", "DOCUREP" + TRIM$(Str$(KK%)))
       FORMUECO(KK%).Caption = ECOARCH$("INDIFRM", FORMUNAME(KK%).TEXT)
    Next KK%
    CUHOR = Val(CONTROL$("REPARA", "VALHORA"))
    Text1 = CSTRING$(MKS$(CUHOR), 4, 8, 2, 2)
    COEFI = Val(CONTROL$("REPARA", "COEFPREA"))
    Text2 = CSTRING$(MKS$(COEFI), 4, 8, 2, 2)
    Label1 = CONTROL$("REPARA", "CODGENER")
    Label3 = TRIM$(CONTROL$("REPARA", "LIPREREP"))
      If Val(Label3) < 1 Then Check1.Value = 1
      '
    PLAENREP$ = TRIM$(CONTROL$("", "PLAENREP"))
    If PLAENREP$ = "" Then PLAENREP$ = "5 Días hábiles de recibida su órden de compra."
    Text4 = PLAENREP$
    '
    CONPAREP$ = TRIM$(CONTROL$("", "CONPAREP"))
    If CONPAREP$ = "" Then CONPAREP$ = "Contado contra entrega"
    Text3 = CONPAREP$
    '
End Sub

Private Sub FORMUNAME_DblClick(Index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        FORMUNAME(Index).TEXT = VDEVU$
        FORMUECO(Index).Caption = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Label1_DblClick()
   Call SELECHO("MASTER")
   VDEVU$ = TRIM$(VALACT1$("MASTER"))
   If VDEVU$ <> "" Then
     Label1 = VDEVU$
   End If
End Sub

Private Sub Label3_Change()
   If Val(Label3) < 1 Then
       Label4 = "Costo Calculado / Informado"
     Else
       Label4 = ECOARCH("LISTAS", Chr$(Val(Label3)))
   End If
End Sub

Private Sub Label3_DblClick()
   Call Command4_Click
End Sub

VERSION 5.00
Begin VB.Form ConfigDepo 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Depósitos  (Producción)"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10095
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
   ScaleHeight     =   5475
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame10 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Vale de Consumo de Materiales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5325
      Left            =   3840
      TabIndex        =   16
      Top             =   120
      Width           =   5160
      Begin VB.TextBox TXTDepBobinasAImprimir 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   3960
         TabIndex        =   59
         Top             =   3720
         Width           =   1000
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   15
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Text            =   "Dep.Bobinas Para Impresión (Quilmes)"
         ToolTipText     =   "Dep.Bobinas Para Impresión"
         Top             =   3720
         Width           =   3525
      End
      Begin VB.CommandButton Command15 
         Caption         =   "."
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
         Left            =   3765
         TabIndex        =   57
         Top             =   3720
         Width           =   180
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   14
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   4080
         Width           =   4725
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4095
         Picture         =   "ConfigDepo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   49
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1605
         Width           =   3750
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2415
         TabIndex        =   48
         Text            =   " "
         Top             =   1200
         Width           =   1380
      End
      Begin VB.CommandButton Command12 
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
         Height          =   360
         Left            =   3780
         TabIndex        =   47
         Top             =   1200
         Width           =   175
      End
      Begin VB.TextBox TXTFAMILIA 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   210
         TabIndex        =   37
         Top             =   4800
         Width           =   1000
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   7
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "Depósito Reserva Liner Asociadas a OF"
         Top             =   4800
         Width           =   3525
      End
      Begin VB.CommandButton Command7 
         Caption         =   "."
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
         Left            =   1235
         TabIndex        =   35
         Top             =   4800
         Width           =   180
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   6
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   3240
         Width           =   4750
      End
      Begin VB.CommandButton Command6 
         Caption         =   "."
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
         Left            =   3765
         TabIndex        =   33
         Top             =   2880
         Width           =   180
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   5
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "Depósito Reserva Liner Asociadas a OF"
         Top             =   2880
         Width           =   3525
      End
      Begin VB.TextBox TXTDEPOLINER 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   3960
         TabIndex        =   31
         Top             =   2880
         Width           =   1000
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Cód."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4095
         TabIndex        =   24
         Top             =   2160
         Width           =   855
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   280
            Left            =   190
            TabIndex        =   25
            Text            =   "VM"
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Numeración"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   210
         TabIndex        =   21
         Top             =   2235
         Width           =   3750
         Begin VB.OptionButton Option7 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Automatica"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2205
            TabIndex        =   23
            Top             =   220
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option8 
            BackColor       =   &H00E2D3C0&
            Caption         =   "Manual"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   525
            TabIndex        =   22
            Top             =   220
            Width           =   1380
         End
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
         Height          =   360
         Left            =   3780
         TabIndex        =   20
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2415
         TabIndex        =   19
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   " "
         Top             =   720
         Width           =   3750
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   4095
         Picture         =   "ConfigDepo.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Etiqueta Scrap A4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   51
         Top             =   1305
         Width           =   2325
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Configurar Familia Liner"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   38
         Top             =   4560
         Width           =   2325
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   26
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   6500
      Left            =   9060
      ScaleHeight     =   6435
      ScaleWidth      =   945
      TabIndex        =   4
      Top             =   0
      Width           =   1000
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
         Height          =   500
         Left            =   120
         Picture         =   "ConfigDepo.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4200
         Width           =   645
      End
      Begin VB.CommandButton command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   105
         Picture         =   "ConfigDepo.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   105
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -300
         Picture         =   "ConfigDepo.frx":0A68
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4980
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Presentación de Costo y Precios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   120
      TabIndex        =   0
      Top             =   9000
      Width           =   5160
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Calculado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Precio Base (C.Rep. + 15 %)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3360
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Informado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1905
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2D3C0&
      Height          =   5325
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3615
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   13
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   4800
         Width           =   3400
      End
      Begin VB.CommandButton Command14 
         Caption         =   "."
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
         Left            =   2325
         TabIndex        =   54
         Top             =   4440
         Width           =   180
      End
      Begin VB.TextBox TXTDepMatPrimas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2520
         TabIndex        =   53
         Top             =   4440
         Width           =   1000
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   12
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Text            =   "Dep-Cons. (Mat. Primas)"
         ToolTipText     =   "Dep-Consumo (Materias Primas)"
         Top             =   4440
         Width           =   2200
      End
      Begin VB.TextBox TXTDEPBAJA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2520
         TabIndex        =   46
         Top             =   3600
         Width           =   1000
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   11
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "Dep-Baja (Scrap Susp.)"
         ToolTipText     =   "Dep-Baja (Scrap en Suspenso)"
         Top             =   3600
         Width           =   2200
      End
      Begin VB.CommandButton Command11 
         Caption         =   "."
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
         Left            =   2325
         TabIndex        =   44
         Top             =   3600
         Width           =   180
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   10
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   3960
         Width           =   3400
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   9
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   3120
         Width           =   3400
      End
      Begin VB.CommandButton Command10 
         Caption         =   "."
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
         Left            =   2325
         TabIndex        =   41
         Top             =   2760
         Width           =   180
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   8
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "Dep- Materiales P/Bolsas"
         Top             =   2760
         Width           =   2200
      End
      Begin VB.TextBox TXTBOLSA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2520
         TabIndex        =   39
         Top             =   2760
         Width           =   1000
      End
      Begin VB.TextBox TXTDEPOSCRAP 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2520
         TabIndex        =   30
         Top             =   1800
         Width           =   1000
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   4
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "Depósito Scrap"
         Top             =   1800
         Width           =   2200
      End
      Begin VB.CommandButton Command5 
         Caption         =   "."
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
         Left            =   2325
         TabIndex        =   28
         Top             =   1800
         Width           =   180
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   3
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2160
         Width           =   3400
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1360
         Width           =   3400
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   600
         Width           =   3400
      End
      Begin VB.CommandButton Command3 
         Caption         =   "."
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
         Left            =   2320
         TabIndex        =   13
         Top             =   1000
         Width           =   180
      End
      Begin VB.CommandButton Command9 
         Caption         =   "."
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
         Left            =   2320
         TabIndex        =   12
         Top             =   240
         Width           =   180
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Depósito Corte Eslinga"
         Top             =   1000
         Width           =   2200
      End
      Begin VB.TextBox TXTDEPOESLINGA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2520
         TabIndex        =   10
         Top             =   1000
         Width           =   1000
      End
      Begin VB.TextBox TXTDEPOCINTILLA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   2520
         TabIndex        =   9
         Top             =   240
         Width           =   1000
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         Index           =   30
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "Depósito Corte Cintilla"
         Top             =   240
         Width           =   2200
      End
   End
End
Attribute VB_Name = "ConfigDepo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
    Call SELECHO("TADEPOSI")
    DEPO$ = VALACT1$("TADEPOSI")
    If DEPO$ = "" Then Exit Sub
    DepOrig% = XVALO(DEPO$)
    Me.TXTBOLSA = DepOrig%
End Sub

Private Sub Command11_Click()
    Call SELECHO("TADEPOSI")
    DEPO$ = VALACT1$("TADEPOSI")
    If DEPO$ = "" Then Exit Sub
    DepBaja% = XVALO(DEPO$)
    Me.TXTDEPBAJA = DepBaja%

End Sub

Private Sub Command12_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command13_Click()

   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1.TEXT), "Formularios de Impresion")

End Sub

Private Sub Command14_Click()
    Call SELECHO("TADEPOSI")
    DEPO$ = VALACT1$("TADEPOSI")
    If DEPO$ = "" Then Exit Sub
    DepConsMatPrima% = XVALO(DEPO$)
    Me.TXTDepMatPrimas = DepConsMatPrima%
End Sub

Private Sub command15_Click()
    Call SELECHO("TADEPOSI")
    DEPO$ = VALACT1$("TADEPOSI")
    If DEPO$ = "" Then Exit Sub
    DepOrig% = XVALO(DEPO$)
    Me.TXTDepBobinasAImprimir = DepOrig%
End Sub

Private Sub Command2_Click()


    Call GRACONTROL("CONTEPLA", "DEPOESLI", TXTDEPOESLINGA.TEXT)
    Call GRACONTROL("CONTEPLA", "DEPOCINT", TXTDEPOCINTILLA.TEXT)
    Call GRACONTROL("CONTEPLA", "DEPOSCRP", Me.TXTDEPOSCRAP.TEXT)
    Call GRACONTROL("CONTEPLA", "DEPLINER", Me.TXTDEPOLINER.TEXT)
    Call GRACONTROL("CONTEPLA", "FAMILINE", Me.TXTFAMILIA.TEXT)
    Call GRACONTROL("CONTEPLA", "DEMATBOL", Me.TXTBOLSA.TEXT)
    Call GRACONTROL("CONTEPLA", "DEPOBAJA", Me.TXTDEPBAJA.TEXT) 'DEPOSITO DE BAJA EN SUSPENSO PARA SCRAP USO EN DEVOLUCION DE BOBINA FRM INGREGOB
    Call GRACONTROL("CONTEPLA", "DECOMAPR", Me.TXTDepMatPrimas.TEXT)
    Call GRACONTROL("CONTEPLA", "DEBOAIMP", Me.TXTDepBobinasAImprimir)
    
    Call GRACONTROL("", "FORCONSU", Text8.TEXT)
    Call GRACONTROL("CONTEPLA", "ETISCRAP", Text1.TEXT)
    
    MONUME$ = "AUTO": If Option8.Value = True Then MONUME$ = "MANUAL"
    Call GRACONTROL("", "NUMVALMA", MONUME$)
    Call GRACONTROL("", "CODCONSU", Text7.TEXT)


    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completada.\Re-Inicie Ahora la Aplicación '" + UCase$(App.EXEName) + "'.")
    Call FINAL("")
    '
End Sub


Private Sub Command3_Click()
    Call SELECHO("TADEPOSI")
    DEPO$ = VALACT1$("TADEPOSI")
    If DEPO$ = "" Then Exit Sub
    DepOrig% = XVALO(DEPO$)
    TXTDEPOESLINGA = DepOrig%

End Sub

Private Sub Command5_Click()
    Call SELECHO("TADEPOSI")
    DEPO$ = VALACT1$("TADEPOSI")
    If DEPO$ = "" Then Exit Sub
    DepOrig% = XVALO(DEPO$)
    TXTDEPOSCRAP = DepOrig%

End Sub

Private Sub Command6_Click()
    Call SELECHO("TADEPOSI")
    DEPO$ = VALACT1$("TADEPOSI")
    If DEPO$ = "" Then Exit Sub
    DepOrig% = XVALO(DEPO$)
    TXTDEPOLINER = DepOrig%
End Sub

Private Sub Command7_Click()
   Call SELECHO("TAFAMIL")
   VALONUE$ = VALACT1$("TAFAMIL")
   If VALONUE$ <> "" Then
      TXTFAMILIA.TEXT = VALONUE$
   End If

End Sub

Private Sub Command8_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text8.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
Private Sub Command4_Click()
   ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
   Call CONECRUN(ACONEC$ + "/" + TRIM$(Text8.TEXT), "Formularios de Impresion")
End Sub


Private Sub Command9_Click()
    Call SELECHO("TADEPOSI")
    DEPO$ = VALACT1$("TADEPOSI")
    If DEPO$ = "" Then Exit Sub
    DepOrig% = XVALO(DEPO$)
    TXTDEPOCINTILLA = DepOrig%

End Sub

Private Sub Form_Load()
    '
    TXTDEPOESLINGA = Control("CONTEPLA", "DEPOESLI")
    TXTDEPOCINTILLA = Control("CONTEPLA", "DEPOCINT")
    TXTDEPOSCRAP = Control("CONTEPLA", "DEPOSCRP")
    Me.TXTDEPOLINER = Control("CONTEPLA", "DEPLINER")
    Me.TXTFAMILIA = Control("CONTEPLA", "FAMILINE")
    Me.TXTBOLSA = Control("CONTEPLA", "DEMATBOL")
    Me.TXTDEPBAJA = Control("CONTEPLA", "DEPOBAJA") 'DEPOSITO DE BAJA EN SUSPENSO PARA SCRAP USO EN DEVOLUCION DE BOBINA FRM INGREGOB
    Me.TXTDepMatPrimas = Control("CONTEPLA", "DECOMAPR")
    TXTDepBobinasAImprimir = Control("CONTEPLA", "DEBOAIMP")
    Text8.TEXT = Control$("", "FORCONSU")
    Text1.TEXT = Control("CONTEPLA", "ETISCRAP")
    CMO$ = TRIM$(Control$("", "CODCONSU")): If CMO$ <> "" Then Text7.TEXT = CMO$


   '
End Sub

Private Sub Text8_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text8.TEXT = VDEVU$
        Text9.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
Private Sub TXTGRAMAJEPAPEL_Change()
End Sub

Private Sub TXTBOLSA_Change()
    TXTLABEL(9) = TRIM$(ECOARCH$("TADEPOSI", Chr$(XVALO(Me.TXTBOLSA))))
End Sub

Private Sub TXTDEPBAJA_Change()
    TXTLABEL(10) = TRIM$(ECOARCH$("TADEPOSI", Chr$(XVALO(Me.TXTDEPBAJA))))
End Sub

Private Sub TXTDepBobinasAImprimir_Change()
    TXTLABEL(14) = TRIM$(ECOARCH$("TADEPOSI", Chr$(XVALO(Me.TXTDepBobinasAImprimir))))
End Sub

Private Sub TXTDepMatPrimas_Change()
 TXTLABEL(13) = TRIM$(ECOARCH$("TADEPOSI", Chr$(XVALO(Me.TXTDepMatPrimas))))
End Sub

Private Sub TXTDEPOCINTILLA_Change()
   TXTLABEL(1) = TRIM$(ECOARCH$("TADEPOSI", Chr$(XVALO(TXTDEPOCINTILLA))))
End Sub

Private Sub TXTDEPOESLINGA_Change()
    TXTLABEL(2) = TRIM$(ECOARCH$("TADEPOSI", Chr$(XVALO(TXTDEPOESLINGA))))
End Sub

Private Sub TXTDEPOLINER_Change()
    TXTLABEL(6) = TRIM$(ECOARCH$("TADEPOSI", Chr$(XVALO(Me.TXTDEPOLINER))))

End Sub

Private Sub TXTDEPOSCRAP_Change()
    TXTLABEL(3) = TRIM$(ECOARCH$("TADEPOSI", Chr$(XVALO(TXTDEPOSCRAP))))

End Sub

Private Sub TXTFAMILIA_Change()
   TAFAM$ = TRIM$(Me.TXTFAMILIA)
   If TAFAM$ = "" Then
      TXTLABEL(7) = ""
      Exit Sub
   End If
   TXTLABEL(7) = ECOARCH$("TAFAMIL", TAFAM$)


End Sub

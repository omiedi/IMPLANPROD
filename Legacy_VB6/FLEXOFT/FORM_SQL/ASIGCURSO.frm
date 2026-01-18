VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ASIGCURSO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Cursos a Alumnos"
   ClientHeight    =   8490
   ClientLeft      =   -15
   ClientTop       =   675
   ClientWidth     =   15600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   15600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Ver Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10200
      TabIndex        =   25
      ToolTipText     =   "Conmuta Entre Ver todos y Solo los que la fecha de inicio no supera los 28 días "
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8640
      Left            =   0
      TabIndex        =   1
      Top             =   -100
      Width           =   15555
      Begin VB.TextBox TXTRESPONSABLETEXT 
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
         Left            =   13200
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   7560
         Width           =   2160
      End
      Begin VB.TextBox Text17 
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
         Index           =   7
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   88
         TabStop         =   0   'False
         Text            =   "Responsable:"
         Top             =   7560
         Width           =   1125
      End
      Begin VB.TextBox TXTRESPONSABLENUM 
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
         Left            =   12795
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   7560
         Width           =   360
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4100
         Left            =   120
         TabIndex        =   57
         Top             =   4440
         Width           =   9675
         Begin VB.CommandButton Command3 
            Caption         =   "Asignar Alumno a Curso Seleccionado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3960
            TabIndex        =   86
            Top             =   240
            Width           =   5415
         End
         Begin VB.TextBox TEXT13 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1755
            Left            =   5160
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   85
            Top             =   1320
            Width           =   4245
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
            Left            =   6340
            Locked          =   -1  'True
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   930
            Width           =   1000
         End
         Begin VB.TextBox Text5 
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
            Left            =   8160
            Locked          =   -1  'True
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   570
            Width           =   1215
         End
         Begin VB.TextBox Text16 
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
            Index           =   3
            Left            =   315
            Locked          =   -1  'True
            TabIndex        =   82
            TabStop         =   0   'False
            Text            =   "Cel. Contacto:"
            Top             =   3450
            Width           =   1300
         End
         Begin VB.TextBox Text18 
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   81
            Top             =   3450
            Width           =   3250
         End
         Begin VB.TextBox Text17 
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
            Index           =   5
            Left            =   315
            Locked          =   -1  'True
            TabIndex        =   80
            TabStop         =   0   'False
            Text            =   "E-Mail:"
            Top             =   2730
            Width           =   1300
         End
         Begin VB.TextBox Text16 
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
            Index           =   2
            Left            =   315
            Locked          =   -1  'True
            TabIndex        =   79
            TabStop         =   0   'False
            Text            =   "Contacto:"
            Top             =   3090
            Width           =   1300
         End
         Begin VB.TextBox Text4 
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
            Index           =   2
            Left            =   315
            Locked          =   -1  'True
            TabIndex        =   78
            TabStop         =   0   'False
            Text            =   "Celular:"
            Top             =   2370
            Width           =   1300
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
            Index           =   19
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   77
            TabStop         =   0   'False
            Text            =   "Doc.:"
            Top             =   570
            Width           =   555
         End
         Begin VB.TextBox Text4 
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
            Left            =   315
            Locked          =   -1  'True
            TabIndex        =   76
            TabStop         =   0   'False
            Text            =   "Télefono:"
            Top             =   2010
            Width           =   1300
         End
         Begin VB.TextBox Text8 
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   75
            Top             =   2730
            Width           =   3250
         End
         Begin VB.TextBox Text9 
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   74
            Top             =   3090
            Width           =   3250
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   73
            Top             =   2370
            Width           =   3250
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   72
            Top             =   2010
            Width           =   3250
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
            Left            =   5760
            Locked          =   -1  'True
            TabIndex        =   71
            Top             =   570
            Width           =   1605
         End
         Begin VB.TextBox Text44 
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   70
            Top             =   1290
            Width           =   3255
         End
         Begin VB.TextBox Text3 
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   69
            Top             =   930
            Width           =   3255
         End
         Begin VB.TextBox Text14 
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   570
            Width           =   3255
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
            Index           =   18
            Left            =   315
            Locked          =   -1  'True
            TabIndex        =   67
            TabStop         =   0   'False
            Text            =   "Localidad:"
            Top             =   1290
            Width           =   1300
         End
         Begin VB.TextBox Text16 
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
            Left            =   315
            Locked          =   -1  'True
            TabIndex        =   66
            TabStop         =   0   'False
            Text            =   "Domicilio:"
            Top             =   930
            Width           =   1300
         End
         Begin VB.TextBox Text17 
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
            Left            =   315
            Locked          =   -1  'True
            TabIndex        =   65
            TabStop         =   0   'False
            Text            =   "Nomb. y Ape.:"
            Top             =   570
            Width           =   1300
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
            Index           =   17
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   64
            TabStop         =   0   'False
            Text            =   "F. de Nacim.:"
            Top             =   930
            Width           =   1100
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
            Index           =   12
            Left            =   7440
            Locked          =   -1  'True
            TabIndex        =   63
            TabStop         =   0   'False
            Text            =   "Sexo:"
            Top             =   570
            Width           =   650
         End
         Begin VB.TextBox Text17 
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
            Index           =   3
            Left            =   315
            Locked          =   -1  'True
            TabIndex        =   62
            TabStop         =   0   'False
            Text            =   "Legajo:"
            Top             =   210
            Width           =   1300
         End
         Begin VB.CommandButton Command7 
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
            Height          =   315
            Left            =   3585
            TabIndex        =   61
            Top             =   210
            Width           =   255
         End
         Begin VB.TextBox Text2 
            Appearance      =   0  'Flat
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
            Left            =   1680
            TabIndex        =   60
            Top             =   210
            Width           =   1875
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
            Index           =   10
            Left            =   315
            Locked          =   -1  'True
            TabIndex        =   59
            TabStop         =   0   'False
            Text            =   "Provinicia"
            Top             =   1650
            Width           =   1300
         End
         Begin VB.TextBox TXTDESCRIPROOVINCIA 
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
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   1650
            Width           =   3255
         End
      End
      Begin VB.TextBox TXTCANTIHORAS1 
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
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Text            =   "Observaciones:"
         Top             =   6120
         Width           =   1515
      End
      Begin VB.TextBox TXTAULA 
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
         Left            =   12000
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   3960
         Width           =   3315
      End
      Begin VB.TextBox TXTPROFESOR 
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
         Left            =   12000
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   3600
         Width           =   3315
      End
      Begin VB.TextBox TXTHHASTA 
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
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   3240
         Width           =   3675
      End
      Begin VB.TextBox TXTHDESDE 
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
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   52
         Top             =   2880
         Width           =   3675
      End
      Begin VB.TextBox TXTSEDICTA 
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
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   2520
         Width           =   3675
      End
      Begin VB.TextBox TXTFINALIZA 
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
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   2160
         Width           =   3675
      End
      Begin VB.TextBox TXTINICIO 
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
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   49
         Top             =   1800
         Width           =   3675
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
         Index           =   9
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "Aula:"
         Top             =   3960
         Width           =   1515
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
         Index           =   8
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "Profesor:"
         Top             =   3600
         Width           =   1515
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
         Index           =   7
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "H.Hasta:"
         Top             =   3240
         Width           =   1515
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
         Index           =   4
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "H.Desde:"
         Top             =   2880
         Width           =   1515
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
         Index           =   2
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "Se dicta:"
         Top             =   2520
         Width           =   1515
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
         Index           =   1
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "Finaliza:"
         Top             =   2160
         Width           =   1515
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
         Index           =   6
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "Inicio:"
         Top             =   1800
         Width           =   1515
      End
      Begin VB.TextBox TXTNUMPROFESOR 
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
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   3600
         Width           =   315
      End
      Begin VB.TextBox TXTNUMAULA 
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
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   3960
         Width           =   315
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
         Index           =   3
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "Código Interno:"
         Top             =   1440
         Width           =   1515
      End
      Begin VB.TextBox TXTCODINT 
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
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   1440
         Width           =   3675
      End
      Begin VB.TextBox TXTDEREXAMEN 
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
         Left            =   13800
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   960
         Width           =   1515
      End
      Begin VB.TextBox Text17 
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
         Index           =   2
         Left            =   11760
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "Derecho de Examen"
         Top             =   960
         Width           =   1965
      End
      Begin VB.TextBox TXTMATRICULA2 
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
         Left            =   13800
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   600
         Width           =   1515
      End
      Begin VB.TextBox Text17 
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
         Left            =   11760
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Matrícula:"
         Top             =   600
         Width           =   1965
      End
      Begin VB.TextBox TXTCURSOTOT 
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
         Left            =   13800
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   240
         Width           =   1515
      End
      Begin VB.TextBox Text17 
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
         Left            =   11760
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   " Total del Curso:"
         Top             =   240
         Width           =   1965
      End
      Begin VB.Frame Frame2 
         Height          =   600
         Left            =   120
         TabIndex        =   26
         Top             =   3840
         Width           =   7695
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "ASIGCURSO.frx":0000
            TabIndex        =   29
            Top             =   240
            Width           =   495
         End
         Begin VB.ComboBox CMBFILTRO 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4440
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   160
            Width           =   3135
         End
         Begin VB.TextBox TXTBUSCAR 
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
            Left            =   720
            TabIndex        =   27
            Top             =   160
            Width           =   2535
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   3480
            OleObjectBlob   =   "ASIGCURSO.frx":006E
            TabIndex        =   30
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox TXTMASID 
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
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   6840
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.TextBox TXTREFERENCIA 
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
         Left            =   10080
         TabIndex        =   23
         Top             =   8200
         Width           =   5280
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
         Index           =   16
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Referencia Curso:"
         Top             =   7800
         Width           =   1515
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
         Index           =   5
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Area:"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text17 
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
         Index           =   6
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Código de Curso:"
         Top             =   240
         Width           =   1965
      End
      Begin VB.CommandButton Command1 
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
         Height          =   315
         Left            =   4800
         TabIndex        =   18
         Top             =   240
         Width           =   255
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   2595
      End
      Begin VB.TextBox TXTDESCRIAREA 
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   600
         Width           =   6795
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
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Descirpción de Curso:"
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox TXTDESCRICURSO 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   960
         Width           =   7515
      End
      Begin VB.TextBox TXTCODAREA 
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
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   675
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
         Index           =   11
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Cant. de Cuotas:"
         Top             =   4320
         Width           =   1515
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
         Index           =   13
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Capacidad:"
         Top             =   4680
         Width           =   1515
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
         Index           =   14
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "Anotados:"
         Top             =   5040
         Width           =   1515
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
         Index           =   15
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "Vacantes:"
         Top             =   5400
         Width           =   1515
      End
      Begin VB.TextBox TXTCANTIHORAS1 
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
         Index           =   16
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "Cant.Horas x Día:"
         Top             =   5760
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox TXTCANTICUOTAS 
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
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   4320
         Width           =   3675
      End
      Begin VB.TextBox TXTCAPACIDAD 
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
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   4680
         Width           =   3675
      End
      Begin VB.TextBox TXTANOTADOS 
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
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   5040
         Width           =   3675
      End
      Begin VB.TextBox TXTVACANTES 
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
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   5400
         Width           =   3675
      End
      Begin VB.TextBox TXTCANTIHORAS 
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
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   5760
         Visible         =   0   'False
         Width           =   3675
      End
      Begin VB.TextBox TXTOBSERVACIONES 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   11640
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   6120
         Width           =   3645
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   2565
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   4524
         _Version        =   393216
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         GridLines       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
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
      Begin MSFlexGridLib.MSFlexGrid MSF_2 
         Height          =   2565
         Left            =   120
         TabIndex        =   31
         ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
         Top             =   1320
         Width           =   9550
         _ExtentX        =   16854
         _ExtentY        =   4524
         _Version        =   393216
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   12632319
         BackColorBkg    =   16777215
         FillStyle       =   1
         GridLines       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
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
   End
   Begin MSComDlg.CommonDialog CMD1 
      Left            =   6600
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   10815
      OleObjectBlob   =   "ASIGCURSO.frx":00DA
      Top             =   105
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu ABM 
      Caption         =   ""
   End
   Begin VB.Menu MNUARCHIVO 
      Caption         =   "Archivo"
      Visible         =   0   'False
      Begin VB.Menu mnucopiar 
         Caption         =   "Copiar Celda Seleccionada"
         Visible         =   0   'False
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "ASIGCURSO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GLOBALFRM_strContenidoCelda$
Dim GLOBALFRM_FILA&
Dim GLOBALFRM_INDIX%


Sub CARGA_DATOS_CURSOS()

   Curso& = XVALO(Text1)
   If Curso& <= 0 Then Exit Sub
   CONDI$ = "CURSOS_ID = " & Curso&
   If CantRegistros&("CURSOS", "CURSOS_ID = " & Curso&, "NOMESS") < 1 Then Exit Sub
   '
   SQLX$ = "SELECT * FROM CURSOS WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " + CONDI$
   Set rs = FLEXCONN.Execute(SQLX$)
   With rs
        Me.TXTCODAREA = SAFETEXT$(!COD_AREA)
        Me.TXTDESCRIAREA = SAFETEXT$(!AREA)
        Me.TXTDESCRICURSO = SAFETEXT$(!Descripcion)
        Me.TXTCURSOTOT = FORMATNUM$(XVALO(!ValCurso), "F12.2")
        Me.TXTMATRICULA2 = FORMATNUM$(XVALO(!ValMatricula), "F12.2")
        Me.TXTDEREXAMEN = FORMATNUM$(XVALO(!ValDerExamen), "F12.2")
    End With
    rs.Close
    Set rs = Nothing
End Sub







Sub CopiarMsfAOtroMsf_Solo_Encabezado(MSF As MSFlexGrid, msf2 As MSFlexGrid, CAMPOS$, Form1 As Form)
     msf2.Rows = 1
     Call CARGA_ENCABEZADO(msf2, CAMPOS$, Form1)
     msf2.Rows = 2
End Sub

Function CURSOSEL()
    Campox$ = "C.Curso/F6;C.Area/F5;Area/A30;Desc.Curso/A58;Curso Cmp./F10.2;Matrícula/F10.2;Der. Examen/F10.2"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campox$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campox$, FRMZELECHO)
    Dim obj As New RECORXET
    Set RS1 = obj.RS_CURSOS()
    Call Carga_MSF_Recordset(RS1, Campox, FRMZELECHO.msf2, 1)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    FRMZELECHO.Width = 15000
    FRMZELECHO.Show 1

    NumeroCliente$ = TRIM(RESP_ZELE_VECT(1))
    CURSOSEL = XVALO(NumeroCliente$)

End Function

Sub CARGAR_GRILLA_CURSOS()
    
    'Campox$ = "F.Inicio/D8;F.Fin/D8;C/F1;Días/A12;C/F1;Días/A12;H.Desde/H12;H.Hasta/H12;Id Prof/F8;Profesor/A32;C.A./f4;Aula/A18;Val.Curso/f10.2;Val.Matrícula/F11.2;Cant.Meses/F10;Capacidad/F10;Ord/F0;CODINT/F7"
     Campox$ = "F.Inicio/D8;F.Fin/D8;C/A0;Días/A12;C/A0;Días/A12;C/A0;Días/A12;H.Desde/H12;H.Hasta/H12;Id Prof/A0;Profesor/A32;C.A./A0;Aula/A18;Val.Curso/f10.2;Val.Matrícula/F11.2;Cant.Meses/F10;Capacidad/F10;Ord/A0;CODINT/F7"

    Call CARGA_ENCABEZADO(Me.MSF, Campox$, Me)
    Dim obj As New RECORXET
    Set RS1 = obj.RS_CURSOS_ACT(XVALO(Text1), 0, Check1.Value)
    REG& = 0
    With RS1
    Do While Not .EOF
       REG& = REG& + 1
       MSF.Rows = REG& + 1
       MSF.TextMatrix(REG&, 1) = FECHATEX$(CVINT(!feinicurs))
       MSF.TextMatrix(REG&, 2) = FECHATEX$(CVINT(!FeFinCurs))
       MSF.TextMatrix(REG&, 3) = XVALO(!COD_DIA1)
       MSF.TextMatrix(REG&, 4) = SAFETEXT$(!desc_dia1)
       MSF.TextMatrix(REG&, 5) = XVALO(!COD_DIA2)
       MSF.TextMatrix(REG&, 6) = SAFETEXT$(!desc_dia2)
       MSF.TextMatrix(REG&, 7) = XVALO(!COD_DIA3)
       MSF.TextMatrix(REG&, 8) = SAFETEXT$(!desc_dia3)
        
       MSF.TextMatrix(REG&, 9) = SAFETEXT$(!H_DESDE)
       MSF.TextMatrix(REG&, 10) = SAFETEXT$(!H_HASTA)
       MSF.TextMatrix(REG&, 11) = XVALO(!ID_PROFESOR)
       MSF.TextMatrix(REG&, 12) = VALOBADA("PROFESORES", "RASO_CLI", "NUME_CLI = " & XVALO(!ID_PROFESOR))
       MSF.TextMatrix(REG&, 13) = XVALO(!ID_AULA)
       CONDI$ = "LISTA = 'AULAS' And CAMPO1 = '" & FORMATOCON0$(XVALO(!ID_AULA), 4) & "'"
       MSF.TextMatrix(REG&, 14) = VALOBADA("DATASQL", "CAMPO2", CONDI$, "MOMES")
       MSF.TextMatrix(REG&, 15) = FORMATNUM$(XVALO(!COSUNI), "F12.2")
       MSF.TextMatrix(REG&, 16) = FORMATNUM$(XVALO(!ValMatricula), "F12.2")
       MSF.TextMatrix(REG&, 17) = XVALO(!CantMeses)
       MSF.TextMatrix(REG&, 18) = XVALO(!CAPACIDAD)
       MSF.TextMatrix(REG&, 19) = XVALO(!MASTER_ID)
       MSF.TextMatrix(REG&, 20) = XVALO(!CODINT)
      .MoveNext
    Loop
    End With
    RS1.Close
    Set RS1 = Nothing

End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If
End Sub


Sub DOBLE_CLICK_EN_GRILLA(MSF As MSFlexGrid)
    Unload VENTABNEW
    REG& = MSF.MouseRow
    If REG& < 1 Then Exit Sub
    If FilaVacia%(Me.MSF, 0, 0) > 0 Then Exit Sub
    'CONFIGURACION DE LA VENTANA
    VENTANA.ColEcoCampo(3) = "CAMPO1;CAMPO2;DATASQL;LISTA  = 'DIAS_SEMANA'"
    VENTANA.ColEcoCampo(5) = "CAMPO1;CAMPO2;DATASQL;LISTA  = 'DIAS_SEMANA'"
    VENTANA.ColEcoCampo(11) = "CAMPO1;CAMPO2;DATASQL;LISTA  = 'AULAS'"
    VENTANA.ColEcoCampo(9) = "NUME_CLI;RASO_CLI;PROFESORES"
    VENTANA.CampEditables = "1;2;3;5;7;8;9;11;13;14;15;16"
    VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 0
    VENTANA.HabilitaInsertar = 1
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "EDICIÓN DE LA FILA SELECCIONADA"
    VENTANA.Inicializar = 1
    
    Campox1$ = "F.Inicio/D8;F.Fin/D8;C/F1;Días/A12;C/F1;Días/A12;H.Desde/H12;H.Hasta/H12;Id Prof/F8;Profesor/A32;C.A./f4;Aula/A18;Val.Curso/f10.2;Val.Matrícula/F11.2;Cant.Meses/F10;Capacidad/F10;Ord/F0;CODINT/F7"
    VENTANA.Campox = Campox1$
    
    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
    A = FilaVacia%(VENTABNEW.MSF, 1)
25  Call EDITARFILA(MSF, VENTABNEW.MSF, REG&)
    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
    A = FilaVacia%(VENTABNEW.MSF, 1)
30  VENTABNEW.Show 1

    If VENTABNEW.APROBADO% > 0 Then
       If VALIDAR_CURSO_OK%(VENTABNEW.MSF) > 0 Then
        Call PEGARFILAEDITADA(VENTABNEW.MSF, MSF, REG&)
        Call UPDATE_CURSO_MATER(REG&, XVALO(MSF.TextMatrix(REG&, 18)))
       Else
         GoTo 30
       End If
    End If
'    Call ResaltarFila(MSF, REG&)
End Sub

Sub EVENTOSELCHANGE(MSF As MSFlexGrid, REG1&)
   'REG1& = MSF.MouseRow
   Call MOSTRAR_DATOS_FILA(MSF, REG1&)
   TXTMASID = XVALO(MSF.TextMatrix(REG1&, 18))

End Sub

''Sub GRABACUOTA(NC, DESCRICURSO$, Tipo%, LETRA$, NUMPVTA%, DOCUMENTO$, IMPORTE#, MODO$, RET_NUMFACTURALAR$, RET_NUMFACTUCOR$, FEINICIO%, DESCRICURSO$)
''    'GRABA REGISTRO POR EL TOTAL DE LA FACTURA
''    '
''    TICOMX$ = "FC"
''    If Tipo% = 3 Then TICOMX$ = "ND"
''    If Tipo% = 6 Then TICOMX$ = "NC"
''    FFI% = HOY(HO$)
''    '
''    NPVTA& = NUMPVTA%
''    PVTA$ = FORMATOCON0$(NPVTA&, 4)
''    NFCZ$ = FORMATOCON0$(TRIM$(Str(SIG_NUMERO_FACTURA&(2, "B"))), 8)
''    NFCX$ = TICOMX$ + "-" + Mid$(NFCZ$, 3) + "-" + LETRA$
''    NFCY$ = TICOMX$ + "-" + PVTA$ + "-" + NFCZ$ + "-" + LETRA$
''    RET_NUMFACTUCOR$ = NFCX$
''    RET_NUMFACTURALAR$ = NFCY$
''    OPMOVIX$ = UCase(Left$(NFCY$, 2))
''    '
''    SIGX% = 1
''    If OPMOVIX$ = "NC" Then SIGX% = (-1)  ' SI ES NOTA DE CREDITO
''    '
''    CoComCor$ = NFCX$
''    CoComLar$ = NFCY$
''    NumeClie = Int(GRUDIVCLI(NC))
''    NumeProv% = 0
''    FechEmi = CVDAT(FEINICIO%)
''    Referen$ = rasocli$(NC)
''    FeContab = CVDAT(FFI%)
''    FeAcredi = CVDAT(FFI%)
''    CuentaCo% = CUECOCLI%(NC)
''    TiMovi$ = "FVEN"
''    OpMovi$ = OPMOVIX$
''    VaMovi$ = LETRA$
''    DOCUMEN$ = "Factura:": If OPMOVIX$ = "NC" Then DOCUMEN$ = "N.Crédito:": If OPMOVIX$ = "ND" Then DOCUMEN$ = "N.Débito:"
''    Descrip$ = DOCUMEN$ + " " + LETRA$ + " (Mat.Codif.)" + " " + PVTA$ + "-" + TRIM$(Str(XVALO(Mid$(NFCY$, 9, 8))))
''    NumeTal = NUMPVTA%
''    NumeCom = XVALO(Mid$(NFCY$, 9, 8))
''    StatInit = 0
''    ImpoDeb = IMPORTE# * SIGX%
''    ValoDola = IMPORTE# 'TOTADOLAR# * SIGX% ' VER COMO SE VA A GRABAR SI ES A VALOR DOLAR ACTUALIZADO
''    Observa$ = DESCRICURSO$
''    NumeroCae$ = ""
''    IdentiCae$ = ""
''    ObservaFactu$ = DESCRICURSO$ & " Correspondiente a " & METEXTO$(FEINICIO%)  ' MATRICULA ,CUOTA O DERECHO DE EXAMEN"
''    Call GRACABAN
''    '
''    'GRABO EL MOVIMIENTO TIPO VENTA POR EL IMPORTE TOTAL POR QUE AQUI NO HAY IVA
''    TOTANETX# = IMPORTE#
''    If TOTANETX# >= 0.005 Then
''      PPVI% = 0
''      If PPVI% < 0 Or PPVI% > 7 Then PPVI% = 1
''      CUENTAX$ = Control$("SVD451S", "CUEVE-" + TRIM$(Str$(PPVI%)))
''      QQ1% = CUECOINT%(CUENTAX$)
''      CuentaCo% = QQ1%
''      TiMovi$ = "VENTA"
''      OpMovi$ = OPMOVIX$
''      VaMovi$ = LETRA$
''      Descrip$ = "Mercaderia Vendida " + NFCX$
''      StatInit = 0
''      ImpoDeb = TOTANETX# * SIGX% * (-1)
''        NENETODOL# = TOTANETX# '(NENETO# / TICAMONE(2))
''        ValoDola = NENETODOL# * SIGX% * (-1)
''      Call GRACABAN
''    End If
''
''    VARIDOC% = 1
''    '
''    If LETRA$ = "B" Then VARIDOC% = 2
''    If LETRA$ = "E" Then VARIDOC% = 3
''    Tipo% = 2
''    If OPMOVIX$ = "ND" Then Tipo% = 3
''    If OPMOVIX$ = "NC" Then Tipo% = 6
''
''    COPAGI% = 1
''   '
''    Call GRADEUFA(Tipo%, VARIDOC%, NumeCom, NC, FEINICIO%, COPAGI%, IMPORTE#)
''
'''SELECT * FROM CURSOS  WHERE CURSOS_ID  = 5
'''SELECT * FROM MASTER WHERE Codint = 38
''End Sub
''
Sub GRADEUFA(Tipo%, VA%, NRO, NCL, FFAC%, COPAGI%, TOTFAC#)
   '
   Dim FVI%(6), PLAII%(6), IM#(6), FEX(3)
   '
   LU$ = LUDAT$
   IORIG# = TOTFAC#
   '
   If Tipo% = 0 Then
     If IORIG# < 0 Then
       Tipo% = 7
     End If
   End If
   '
   If Tipo% > 3 Then
     NUPAG% = 1
     FVI%(1) = FFAC%
     IM#(1) = TOTFAC#     ' IORIG#
     GoTo 41
   End If
   '
   ' busca condicion de pago segun COPAGI%
20 Call BUCONPA(LU$, COPAGI%, CoPag$, FREF%, PLA$)
   '
   ' determina vencimientos segun total factura y condicion de pago
30 FF = FFAC%: If FF < 1 Then FF = FECHANUM("01/01/80")
   FFS$ = FECHATEX$(FF)
   PLA$ = Left$(PLA$ + String$(6, 0), 6)

   If (FREF% < 1 Or FREF% > 4) Then
           MENSERR 24, "Mal Fecha Referencia en Cond.Pago Nro." + Str$(COPAGI%)
           FREF% = 1
   End If

   If FREF% = 2 Then FREF% = 1   ' AUN NO FUNCIONA SEMANA DE FACTURA
   If FREF% <> 3 Then GoTo 35
   If XVALO(Left$(FFS$, 2)) <= 15 Then FEBAS$ = "15" + Mid$(FFS$, 3)
   If XVALO(Left$(FFS$, 2)) > 15 Then FEBAS$ = "31" + Mid$(FFS$, 3)
35 If FREF% = 1 Then FEBAS$ = FFS$
   If FREF% = 4 Then FEBAS$ = "31" + Mid$(FFS$, 3)
   '
   For i% = 1 To 6
     FVI%(i%) = 0
   Next i%
   '
   FEFA = FECHANUM(FEBAS$)
   For i% = 1 To 6
     PLAI% = Asc(Mid$(PLA$, i%, 1))
     PLAII%(i%) = PLAI%
     If i% > 1 Then
       If (PLAII%(i%) < PLAII%(i% - 1) Or PLAII%(i%) <= 0) Then
         GoTo 31
       End If
     End If
     '
     FEXU = FEFA + PLAI%
     While DIENTRE(FEFA, FEXU) < PLAI%
       FEXU = DIPOST(FEXU)
     Wend
     '
     FVI%(i%) = FEXU
     NUPAG% = i%
   Next i%

31 FVIJ% = FVI%(NUPAG%)
   If NUPAG% < 2 Then FVIJ% = FFAC%
   '
   ' determina importes segun numero de pagos
40 If NUPAG% < 1 Then NUPAG% = 1
   IM#(1) = (Int(100 * IORIG# / NUPAG% + 0.5)) / 100

   For H% = 2 To NUPAG% - 1
     IM#(H%) = IM#(1)
   Next H%
   IM#(NUPAG%) = IORIG#

   For H% = 1 To NUPAG% - 1
     IM#(NUPAG%) = IM#(NUPAG%) - IM#(H%)
   Next H%
   '
   ICANCE# = IORIG# - TOTFAC#
   For H% = 1 To NUPAG%
     If ICANCE# > 0 Then
       If ICANCE# >= IM#(H%) Then
           ICANCE# = ICANCE# - IM#(H%)
           IM#(H%) = 0
         Else
           IM#(H%) = IM#(H%) - ICANCE#
           ICANCE# = 0
       End If
     End If
   Next H%
   '
41 NuorVen% = 0
   SIGX% = 1: If Tipo% > 3 Then SIGX% = (-1)
   For i% = 1 To NUPAG%
     If Abs(IM#(i%)) >= 0.005 Then
       FechEmi = CVDAT(FFAC%)
       IBruTot# = IORIG#
       NuorVen% = NuorVen% + 1
       FeVenc = CVDAT(FVI%(i%))
       ImVenc# = (Abs(IM#(i%))) * SIGX%
       Call GRADEUPEN
     End If
   Next i%
   FVENCIM = CVDAT(FVI%(1))
   If IsDate(FVENCIM) = False Then
     FVENCIM = ""
   End If
   '
End Sub
'


Sub MENUXELDA(MSF As MSFlexGrid)
     If DERACCE("MODCURSO", "MODIFICAR DATOS DE CURSO") < 2 Then Exit Sub
     GLOBALFRM_INDIX% = Index
     REG& = MSF.MouseRow
     COL% = MSF.MouseCol
     Call MENU_CELDA(MSF, REG&, COL%)
End Sub

Function SIG_NUMERO_FACTURA&(TIFADOC%, LETRA$)

   DATEDOC% = HOY(HO$)
   '
   Screen.MousePointer = 11
   '
   OPMOVIX$ = "FC"
   If TIFADOC% = 3 Then OPMOVIX$ = "ND"
   If TIFADOC% = 6 Then OPMOVIX$ = "NC"
   VAMOVIX$ = LETRA$
   
   ASELE$ = "SELECT NumeComp FROM CajaBanc WITH(NOLOCK)"
   If OPMOVIX$ = "ND" Or OPMOVIX$ = "NC" Then
     If NUMECORR% = (-2) Then    ' Unificada debitos y creditos
       ASELE$ = ASELE$ + "WHERE ((TipoMovim = 'FVEN' AND (OpciMovim = 'ND' OR OpciMovim = 'NC')) "
       ASELE$ = ASELE$ + "OR (TipoMovim = 'FCOM' AND (OpciMovim = 'DP' OR OpciMovim = 'CP'))) "
       GoTo 25
     End If
   End If
   '
   ASELE$ = ASELE$ + "WHERE TipoMovim = 'FVEN' "
   If NUMECORR% <> 1 Then
       ASELE$ = ASELE$ + "AND OpciMovim = '" & OPMOVIX$ & "' "
     Else
       ASELE$ = ASELE$ + "AND (OpciMovim = 'FC' OR OpciMovim = 'ND' OR OpciMovim = 'NC' OR OpciMovim = 'ANUL') "
   End If
   '
25 ASELE$ = ASELE$ + "AND VariMovim = '" & VAMOVIX$ & "' "
   If XVALO(Label12) <= 1 Then
       ASELE$ = ASELE$ + "AND (NumeTalo = 0 OR NumeTalo = 1) "
     Else
       ASELE$ = ASELE$ + "AND NumeTalo = " & XVALO(Label12) & " "
   End If
   ASELE$ = ASELE$ + "AND FechEmisi >= '" & FLIMI & "' "
   ASELE$ = ASELE$ + " ORDER BY NUMECOMP DESC "
   '
   NUMAXI& = 0
   Dim MIRS As ADODB.Recordset
   Set MIRS = New ADODB.Recordset
   MIRS.CursorLocation = adUseServer
   Set MIRS = READSET(ASELE$)
    NUMAXI& = XVALO(MIRS!NUMECOMP)
   MIRS.Close
   Set MIRS = Nothing
   '
   NPP& = NUMAXI& + 1
   '
   
   SIG_NUMERO_FACTURA& = NPP&
   '
   
999 Screen.MousePointer = 1
   '
End Function

Private Sub Command1_Click()
    Command1.Enabled = False
    '
'    If SKALTA.Visible = True Then
'      TX$ = "PARA SELECCIONAR CURSO EXISTENTE" + vbCrLf
'      TX$ = TX$ + "DEBE VOLVER A MODO EDICIÓN (Presionando el Botón 'Edición')"
'      MsgBox TX$
'      GoTo 99
'    End If
    '
    Call ABRECONEXION
    '
    If CURSOSEL > 0 Then
      Text1 = XVALO(RESP_ZELE_VECT(1))
    End If
    '
99  Command1.Enabled = True
    '
    
End Sub







Private Sub Command9_Click()
   RUTAX$ = SELECCION_Archivo$("*", CMD1)
   If RUTAX$ <> "" Then
        On Error Resume Next
        Image1.Picture = LoadPicture(RUTAX$)
        If Err > 0 Then MsgBox Err.Description
   Else
      GoTo 99
   End If
   TXTRUTAFIRMA = RUTAX$
99 Screen.MousePointer = 1

End Sub

Private Sub DTP1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If
End Sub





Private Sub Command3_Click()
    Command3.Enabled = False
    If XVALO(Me.TXTCODINT) < 1 Then
       Call COMUNI("Debe Seleccionar un Curso para Asociar")
       GoTo 99
    End If
    If XVALO(Me.Text2) < 1 Then
       Call COMUNI("Debe Seleccionar un Alumno para Asociar")
       GoTo 99
    End If
    '
    T1$ = "Confirma la Asignación del Curso:\"
    T1$ = T1$ + "" & TXTDESCRICURSO.Text
    T1$ = T1$ + "\Al Alumno: " & Text14.Text
    T1$ = T1$ + "\\Cancelar\Confirmar"
    OPX% = COMALTER%(T1$)
    If OPX% < 2 Then GoTo 99
    '
    Call UPDATEALUMYCURSO(XVALO(Text2), XVALO(Me.TXTCODINT), OK%)
    '
'    If OK% >= 0 Then
'        ACONEC$ = "RECIBO10/RECICUTA" + TRIM$(Text2.Text) ' NUME
'        Call CONECRUN(ACONEC$, "Generación de Recibo")
'    End If
    Call MOSTRAR_DATOS_FILA(MSF, MSF.MouseRow)
    TXTMASID = XVALO(MSF.TextMatrix(MSF.MouseRow, 18))


99  Command3.Enabled = True
End Sub

Sub UPDATEALUMYCURSO(LEGAJO&, CIX1&, OK%)
    'GRABACION DE ASOCIASION DE ALUMNO A CURSO
    SQLX$ = "SELECT * FROM ALUMNO_CURSO"
    SQLX$ = SQLX$ + " WHERE NUME_CLI = " & LEGAJO&
    SQLX$ = SQLX$ + " AND  CURSOS_ID = " & CIX1&
    '
    UN% = USNBR% Mod 100
    On Error GoTo ERROR_GUARDAR
    Set rs = New ADODB.Recordset
    rs.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    FLEXCONN.BeginTrans
    With rs
        If (Not .EOF) Then
           FLEXCONN.RollbackTrans
           rs.Close
           Set rs = Nothing
           Call COMUNI("Yá Está Inscripto él Alumno a Este Curso")
           OK% = -1
           Exit Sub
        Else
          .AddNew
          !NUME_CLI = LEGAJO&
          !CURSOS_ID = CIX1& ' codint
          
          !FECHA = CVDAT(HOY(HO$))
          !Descripcion = TRIM$(TXTDESCRICURSO.Text)
          !Status = 0
          !USUARIO = UN%
          .Update
        End If
   End With
   '
   TXT1$ = " ALTA DE ALUMNO A CURSO: " & TXTDESCRICURSO
   Call ANOTAX_ALUMNO(LEGAJO&, TXT1$)
   
   'GENERACION DE FACTURA SEGUN CONDICION DE PAGO
   CONDI$ = "CODINT = " & CIX1
   CANT_CUOTAS% = XVALO(VALOBADA("MASTER", "CantMeses", CONDI$, "NOMESS"))
   ID_CURSO = XVALO(VALOBADA("MASTER", "ID_CURSO", CONDI$, "NOMESS"))
   FEINICIO% = CVINT(VALOBADA("MASTER", "FeIniCurs", CONDI$, "NOMESS"))
   
   CONDI$ = "CURSOS_ID = " & ID_CURSO
   VAL_CUOTA# = XVALO(VALOBADA("CURSOS", "VALCURSO", CONDI$, "NOMESS"))
   VAL_MATRICULA# = XVALO(VALOBADA("CURSOS", "VALMATRICULA", CONDI$, "NOMESS"))
   CUOTAMESANO$ = METEXTO$(FEINICIO%)
   MESYANOX$ = METEXTO$(FEINICIO%)
   
   'Call GRABACUOTA(LEGAJO&, TXTDESCRICURSO, 2, "B", 1, CUOTAMESANO$, VAL_CUOTA#, "ALTA", RET_NUMFACTURALAR$, RET_NUMFACTUCOR$, FEINICIO%)
   Call GRABACUOTA(LEGAJO&, TXTDESCRICURSO, 2, "B", 1, "CUOTA INICIAL ", VAL_CUOTA#, "ALTA", RET_NUMFACTURALAR$, RET_NUMFACTUCOR$, FEINICIO%)
   NPVX$ = "0001"
   CIX% = CIX1&
   Call NUREMIFA(NROREMIT&, NPVX$)
   Call MOVISTOK(HOY(HO$), CIX%, METEXTO$(FEINICIO%), "RT", "RT-0001-" + FORMATOCON0(NROREMIT&, 8), "RT-" & NROREMIT&, RET_NUMFACTURALAR$, RET_NUMFACTUCOR$, "CUOTA INICIAL", 1, VAL_CUOTA#, 1, LEGAJO&, 1, 1, 0, 1, TXTDESCRICURSO, "", "CUOTA 1/" & CANT_CUOTAS%, VAL_CUOTA#, , , 1)
   
   Call GRABACUOTA(LEGAJO&, TXTDESCRICURSO, 2, "B", 1, "MATRICULA ", VAL_MATRICULA#, "ALTA", RET_NUMFACTURALAR$, RET_NUMFACTUCOR$, FEINICIO%)
   Call MOVISTOK(HOY(HO$), CIX%, METEXTO$(FEINICIO%), "RT", "RT-0001-" + FORMATOCON0(NROREMIT&, 8), "RT-" & NROREMIT&, RET_NUMFACTURALAR$, RET_NUMFACTUCOR$, "MATRICULA", 1, VAL_MATRICULA#, 1, LEGAJO&, 1, 1, 0, 1, TXTDESCRICURSO, "", "MATRICULA", VAL_MATRICULA#, , , 1)
   
   FLEXCONN.CommitTrans
   rs.Close
   Set rs = Nothing
   Screen.MousePointer = 1
   
   Exit Sub
   
ERROR_GUARDAR:

    If Err <> 0 Then
      FLEXCONN.RollbackTrans
      MsgBox "ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description
      Screen.MousePointer = 1
      OK% = -1
      GoTo 99
      
    End If
    '

99
End Sub
Private Sub Command7_Click()
    Command7.Enabled = False
    '
    '
    Call ABRECONEXION
    '
    If ALUMONOSEL > 0 Then
      Text2 = XVALO(RESP_ZELE_VECT(1))
    End If
    '
99  Command7.Enabled = True

End Sub

Private Sub Form_Load()
    ''
    Call ABRECONEXION
    Call ABRECLIEN
'    Call CREACAMPO("", "CURSOSACTI", "ID_CURSO", 4, 4, 1)
'    Call CREACAMPO("", "CURSOSACTI", "F_INICIO", 8, 0)
'    Call CREACAMPO("", "CURSOSACTI", "F_FIN", 8, 0)
'    Call CREACAMPO("", "CURSOSACTI", "DIA1", 10, 24)
'    Call CREACAMPO("", "CURSOSACTI", "DIA2", 10, 24)
'    Call CREACAMPO("", "CURSOSACTI", "H_DES", 10, 24)
'    Call CREACAMPO("", "CURSOSACTI", "H_HAS", 10, 24)
'    Call CREACAMPO("", "CURSOSACTI", "CAN_H_TRAB", 7, 8)
'    Call CREACAMPO("", "CURSOSACTI", "ANO", 3, 0, 1)
'    Call CREACAMPO("", "CURSOSACTI", "ID_PROFESOR", 4, 4, 1)
'    If ValidaExisteCampo%("CURSOSACTI_ID", "CURSOSACTI") < 1 Then
'      SQLX$ = "ALTER TABLE CURSOSACTI ADD CURSOSACTI_ID INT IDENTITY"
'      FLEXCONN.Execute (SQLX$)
'    End If
'    Call APLICAR_SKIN1(Me, App.Path & "\SKINS\dogmax.skn")
'    Campox$ = "F.Inicio/D8;F.Fin/D8;C/F1;Días/A10;C/F1;Días/A10;H.Desde/H10;H.Hasta/H10;Id Prof/F8;Profesor/A32;C.A./f4;Aula/A18;Val.Curso/f10.2;Val.Matrícula/F11.2;Cant.Meses/F10;Capacidad/F10;Ord/F0;CODINT/F7"
    Campox$ = "F.Inicio/D8;F.Fin/D8;C/A0;Días/A12;C/A0;Días/A12;C/A0;Días/A12;H.Desde/H12;H.Hasta/H12;Id Prof/A0;Profesor/A32;C.A./A0;Aula/A18;Val.Curso/f10.2;Val.Matrícula/F11.2;Cant.Meses/F10;Capacidad/F10;Ord/A0;CODINT/F7"
    
    Call CARGA_ENCABEZADO(Me.MSF, Campox$, Me, 1)
    Call CARGA_ENCABEZADO(Me.MSF_2, Campox$, Me, 0)
    MSF.Rows = 2
    'PROPIEDADES DE SELECCION DE FILA
    MSF.SelectionMode = flexSelectionByRow
    MSF.FocusRect = flexFocusNone
    '

End Sub





Private Sub MSF_2_DblClick()
   REG& = Me.MSF_2.MouseRow
   If REG& < 1 Then Exit Sub
   Call T_Espera(1 / 3)
   MASID& = XVALO(Me.MSF_2.TextMatrix(REG&, 18)) 'EL ID DE LA TABLA
   
   TXTBUSCAR = ""
   Me.MSF_2.Visible = False
   MSF.Visible = True
   For i& = 1 To MSF.Rows - 1
        IDGRILLA& = XVALO((Me.MSF.TextMatrix(i&, 18)))
        If IDGRILLA& = MASID& Then
             Call ResaltarFila(Me.MSF, i&)
             Call T_Espera(1 / 3)
             Call EVENTOSELCHANGE(Me.MSF, i&) ' PARA QUE MUESTRE EN LA PARTE INFERIOR  LOS DATOS
             Exit For
        End If
   Next i&
   
End Sub


Private Sub MSF_2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
  End If

End Sub


Private Sub MSF_2_RowColChange()
   Call EVENTOSELCHANGE(Me.MSF_2, Me.MSF_2.Row)
End Sub

Private Sub MSF_RowColChange()
      Call EVENTOSELCHANGE(Me.MSF, MSF.Row)
End Sub

Private Sub MSF_SelChange()
   'Call EVENTOSELCHANGE(ME.MSF, MSF.MouseRow)
End Sub

Private Sub Text2_Change()
'     Call LIMPIAR
     Call CARGA_DATOS_ALUMNMOS

End Sub

Private Sub TXTBUSCAR_Change()
    '
    If TRIM$(TXTBUSCAR) = "" Then
      MSF_2.Visible = False
      MSF.Visible = True
      Exit Sub
    End If
    '
    MSF_2.Rows = 2
    For i& = 1 To MSF_2.Cols - 1
       MSF_2.TextMatrix(1, i&) = ""
    Next i&
    REG& = 0
    MSF_2.Visible = True
    MSF_2.ZOrder 0
    MSF_2.Cols = MSF.Cols
    MSF_2.Redraw = False
    MSF.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For i& = 1 To MSF.Rows - 1
      For j& = 1 To MSF.Cols - 1
        If (InStr(UCase$(MSF.TextMatrix(i&, j&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.Text)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(MSF.TextMatrix(0, j&))) Then
            REG& = REG& + 1
            MSF_2.Rows = REG& + 1
            For H& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(REG&, H&) = MSF.TextMatrix(i&, H&)
            Next H&
          End If
          If (TRIM$(CMBFILTRO.Text)) = "" Then Exit For
        End If
      Next j&
    Next i&
    '
    MSF_2.Redraw = True
    
    '
End Sub

Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      TXTBUSCAR.Text = ""
   End If
End Sub

Private Sub CMBFILTRO_Change()
   TXTBUSCAR_Change
End Sub
Private Sub MSF_DblClick()
  'Call DOBLE_CLICK_EN_GRILLA(Me.MSF)
End Sub

Function VALIDAR_CURSO_OK%(MSF1 As MSFlexGrid)

       VALHOR& = VALIDARHORAS&(MSF1, 7, 8)
       If VALHOR& > 0 Then
          Call COMUNI("Rango Horario Inválido en Fila: " & VALHOR&)
          GoTo 30
       End If
       '
       If FECHANUM(MSF1.TextMatrix(1, 1)) < 1 Then
          Call COMUNI("Fecha Desde Inválida")
          GoTo 30
       End If
       If FECHANUM(MSF1.TextMatrix(1, 2)) < 1 Then
          Call COMUNI("Fecha Hasta Inválida")
          GoTo 30
       End If
       If FECHANUM(MSF1.TextMatrix(1, 7)) > FECHANUM(MSF1.TextMatrix(1, 8)) Then
           Call COMUNI("Fecha Ingresada es Menor el Inicio que la Fecha de Finalización")
           GoTo 30
       End If
       '
       If XVALO(MSF1.TextMatrix(1, 3)) > 0 Then
        If VALIDAR_ECO(MSF1, 4, 0) > 0 Then
           Call COMUNI("Día 1 Inválido")
           GoTo 30
        End If
       End If
       If XVALO(MSF1.TextMatrix(1, 5)) > 0 Then
        If VALIDAR_ECO(MSF1, 6, 0) > 0 Then
           Call COMUNI("Día 2 Inválido")
           GoTo 30
        End If
       End If
       '
       If XVALO(MSF1.TextMatrix(1, 10)) < 1 Then
        If VALIDAR_ECO(MSF1, 11, 0) > 0 Then
           Call COMUNI("Profesor Inválido")
           GoTo 30
        End If
       Else
           Call COMUNI("Falta Ingresar Profesor")
           GoTo 30
       End If
       If XVALO(MSF1.TextMatrix(1, 11)) > 0 Then
        If VALIDAR_ECO(MSF1, 12, 0) > 0 Then
           Call COMUNI("Aula Inválida")
           GoTo 30
        End If
       Else
           Call COMUNI("Falta Ingresar el Aula")
           GoTo 30
       End If
       If XVALO(MSF1.TextMatrix(1, 13)) < 1 Then
           Call COMUNI("Falta Ingresar el Costo del Curso")
           GoTo 30
       End If
       If XVALO(MSF1.TextMatrix(1, 14)) < 1 Then
           Call COMUNI("Falta Ingresar el Costo de la Matrícula")
           GoTo 30
       End If
       If XVALO(MSF1.TextMatrix(1, 15)) < 1 Then
           Call COMUNI("Falta Ingresar La Cantidad de Meses")
           GoTo 30
       End If
       If XVALO(MSF1.TextMatrix(1, 16)) < 1 Then
           Call COMUNI("Falta Ingresar La Capacidad")
           GoTo 30
       End If
       
       VALIDAR_CURSO_OK% = 1
       Exit Function
       
30     VALIDAR_CURSO_OK% = 0: Exit Function
End Function
Private Sub MSF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
     Call MENUXELDA(MSF)
  End If
End Sub
Sub MENU_CELDA(MSF As MSFlexGrid, REG&, COL%)

        If MSF.Rows <= 1 Then Exit Sub
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        GLOBALFRM_FILA& = REG&
        PopupMenu MnuArchivo
        Call INVICTRL
        On Error Resume Next
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA
        On Error GoTo 0
End Sub

 Sub INVICTRL()
   mnucopiar.Visible = Not (mnucopiar.Visible)
   mnuQuitar.Visible = Not (mnuQuitar.Visible)
End Sub

Private Sub mnuquitar_Click()
End Sub
Private Sub mnucopiar_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub
Sub LIMPIAR()
            Me.TXTCODAREA = ""
            Me.TXTDESCRIAREA = ""
            Me.TXTDESCRICURSO = ""
            Me.MSF.Rows = 2
            
            Me.TXTSEDICTA = ""
            Me.TXTHDESDE = ""
            Me.TXTHHASTA = ""
            Me.TXTNUMPROFESOR = ""
            Me.TXTPROFESOR = ""
            Me.TXTNUMAULA = ""
            Me.TXTAULA = ""
            
'            Me.TXTVALORTOTAL = ""
'            Me.TXTMATRICULA = ""
            Me.TXTCANTICUOTAS = ""
            Me.TXTCAPACIDAD = ""
            Me.TXTCODINT = ""
            Me.TXTOBSERVACIONES = ""
            Me.TXTANOTADOS = ""
            Me.TXTCANTICUOTAS = ""
            Me.TXTREFERENCIA = ""
            
            
            Call LIMPIAR_MSF_SIN_BORRAR_REGISTROS(MSF, 0)
'            Check1.Value = 0
End Sub

Sub UPDATE_CURSO_MATER(Optional REG&, Optional MAS_ID&) 'ES ALTA O EDICION
    
    If MAS_ID& > 0 Then
       SQLX$ = "SELECT * FROM MASTER WHERE MASTER_ID =  " & MAS_ID&
    Else
       SQLX$ = "SELECT * FROM MASTER WHERE CODINT < 1 "
        NUEVOCODINT& = 1 + XVALO(VALOBADA("MASTER", "MAX(CODINT)", "CODINT > 0", "NOMESS"))
    End If
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
25  On Error Resume Next
    rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    '
    With rs
        If .EOF Then
            .AddNew   ' YA CONVERTIDO
            !CODINT = NUEVOCODINT&
            !FeALTA = CVDAT(HOY(HO$))
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
        End If
       
       !feinicurs = CVDAT(CVINT(MSF.TextMatrix(REG&, 1)))
       !FeFinCurs = CVDAT(CVINT(MSF.TextMatrix(REG&, 2)))
       !COD_DIA1 = XVALO(MSF.TextMatrix(REG&, 3))
       !COD_DIA2 = XVALO(MSF.TextMatrix(REG&, 5))
       !COD_DIA3 = XVALO(MSF.TextMatrix(REG&, 7))
       
       !desc_dia1 = MSF.TextMatrix(REG&, 4)
       !desc_dia2 = MSF.TextMatrix(REG&, 6)
       !desc_dia3 = MSF.TextMatrix(REG&, 8)
       
       !H_DESDE = MSF.TextMatrix(REG&, 9)
       !H_HASTA = MSF.TextMatrix(REG&, 10)
       H_TRABAJADAS = XVALO(DIFERENCIAHORASCENTESIMAL(HORANUM(MSF.TextMatrix(REG&, 9)), HORANUM(MSF.TextMatrix(REG&, 10))))
       !CantHorCso = H_TRABAJADAS
       !ID_PROFESOR = MSF.TextMatrix(REG&, 11)
       !ID_AULA = MSF.TextMatrix(REG&, 13)
       !COSUNI = ROUND(MSF.TextMatrix(REG&, 15), 2)
       !ValMatricula = ROUND(MSF.TextMatrix(REG&, 16), 2)
       !CantMeses = MSF.TextMatrix(REG&, 17)
       !CAPACIDAD = MSF.TextMatrix(REG&, 18)
       
       !Ferevis = CVDAT(HOY(HO$))
       !COD_AREA = XVALO(Me.TXTCODAREA)
       !DESCR_AREA = Left$(Me.TXTDESCRIAREA, 128)
       !DESCRICURSO = Left$(Me.TXTDESCRICURSO, 256)
       !referencia = Left$(Me.TXTREFERENCIA, 256)
       !Observa = Me.TXTOBSERVACIONES
       !ID_CURSO = XVALO(Me.Text1)
'       LOSDIAS$ = DIA1$: If TRIM$(DIA2$) <> "" Then LOSDIAS$ = LOSDIAS$ & " Y " & DIA2$
       .Update
       End With
       IDCURS& = XVALO(Text1)
       Text1 = ""
       Text1 = IDCURS&
       

End Sub

Sub MOSTRAR_DATOS_FILA(MSF As MSFlexGrid, REG&)
       Me.TXTINICIO = TRIM$(MSF.TextMatrix(REG&, 1))
       Me.TXTFINALIZA = TRIM$(MSF.TextMatrix(REG&, 2))
       DIA1$ = TRIM$(MSF.TextMatrix(REG&, 4))
       DIA2$ = TRIM$(MSF.TextMatrix(REG&, 6))
       DIA3$ = TRIM$(MSF.TextMatrix(REG&, 8))
       LOSDIAS$ = DIA1$: If TRIM$(DIA2$) <> "" Then LOSDIAS$ = LOSDIAS$ & " Y " & DIA2$
       If TRIM$(DIA3$) <> "" Then LOSDIAS$ = DIA1$: If TRIM$(DIA2$) <> "" Then LOSDIAS$ = LOSDIAS$ & " , " & DIA2$ & " Y " & DIA3$
       Me.TXTSEDICTA = TRIM$(LOSDIAS$)
       Me.TXTHDESDE = TRIM$(MSF.TextMatrix(REG&, 9))
       Me.TXTHHASTA = TRIM$(MSF.TextMatrix(REG&, 10))
       Me.TXTNUMPROFESOR = TRIM$(MSF.TextMatrix(REG&, 11))
       Me.TXTPROFESOR = TRIM$(MSF.TextMatrix(REG&, 12))
       Me.TXTNUMAULA = TRIM$(MSF.TextMatrix(REG&, 13))
       Me.TXTAULA = TRIM$(MSF.TextMatrix(REG&, 14))
       
'       Me.TXTVALORTOTAL = TRIM$(FORMATNUM$(MSF.TextMatrix(REG&, 13), "f12.2"))
'       Me.TXTMATRICULA = TRIM$(FORMATNUM$(MSF.TextMatrix(REG&, 14), "f12.2"))
       Me.TXTCANTICUOTAS = TRIM$(MSF.TextMatrix(REG&, 17))
       Me.TXTCAPACIDAD = TRIM$(MSF.TextMatrix(REG&, 18))
       Me.TXTCODINT = MSF.TextMatrix(REG&, 20)
       Me.TXTCANTIHORAS = H_TRABAJADAS = XVALO(DIFERENCIAHORASCENTESIMAL(HORANUM(TXTHDESDE), HORANUM(TXTHHASTA)))
       
       CONDI$ = "CODINT = " & XVALO(TXTCODINT)
       SQLX$ = "SELECT OBSERVA,REFERENCIA,COD_RESPONSABLE,DESC_RESPONSABLE FROM MASTER WITH(NOLOCK) WHERE " & CONDI$
       Set rs = READSET(SQLX$)
       Me.TXTOBSERVACIONES = SAFETEXT$(rs!Observa) 'VALOBADA("MASTER", "OBSERVA", "CODINT = " & XVALO(TXTCODINT), "NOMESS")
       Me.TXTREFERENCIA = SAFETEXT$(rs!referencia) 'VALOBADA("MASTER", "REFERENCIA", "CODINT = " & XVALO(TXTCODINT), "NOMESS")
       Me.TXTRESPONSABLENUM = XVALO(rs!COD_RESPONSABLE)
       Me.TXTRESPONSABLETEXT = SAFETEXT$(rs!DESC_RESPONSABLE)
       rs.Close
       Set rs = Nothing

'       Me.TXTOBSERVACIONES = VALOBADA("MASTER", "OBSERVA", "CODINT = " & XVALO(TXTCODINT), "NOMESS")
'       Me.TXTREFERENCIA = VALOBADA("MASTER", "REFERENCIA", "CODINT = " & XVALO(TXTCODINT), "NOMESS")
           TXTANOTADOS = CantRegistros("ALUMNO_CURSO", "CURSOS_ID = " & XVALO(MSF.TextMatrix(REG&, 20)) & " AND STATUS >= 0", "NOMESS")
       Me.TXTVACANTES = XVALO(TXTCAPACIDAD) - XVALO(TXTANOTADOS)

'       Me. = MSF.TextMatrix(REG&, i&)
'       Me.TXTINICIO = MSF.TextMatrix(REG&, i&)
'       Me.TXTINICIO = MSF.TextMatrix(REG&, i&)
'       Me.TXTINICIO = MSF.TextMatrix(REG&, i&)
'       Me.TXTINICIO = MSF.TextMatrix(REG&, i&)
'       Me.TXTINICIO = MSF.TextMatrix(REG&, i&)
'       Me.TXTINICIO = MSF.TextMatrix(REG&, i&)
       
       
    
End Sub


Private Sub Text1_Change()
        Call LIMPIAR
        Call CARGA_DATOS_CURSOS
        Call CARGAR_GRILLA_CURSOS

End Sub
Private Sub Text1_DblClick()
   Call Command1_Click
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If
End Sub

Private Sub TEXT14_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text17_DblClick(Index As Integer)
    If Index = 6 Then Command1_Click
    
End Sub


Private Sub Text18_Change()
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Sub CARGA_DATOS_ALUMNMOS()

   LEGAJO& = XVALO(Text2)
   If LEGAJO& <= 0 Then Exit Sub
   CONDI$ = "NUME_CLI = " & LEGAJO&
   If CantRegistros&("DEUDOR12", "NUME_CLI = " & LEGAJO&, "NOMESS") < 1 Then Exit Sub
   '
   SQLX$ = "SELECT * FROM DEUDOR12 WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " + CONDI$
   Set ALUMTMP = FLEXCONN.Execute(SQLX$)
   With ALUMTMP
        Text14 = SAFETEXT$(!RASO_CLI)
        Text3 = SAFETEXT$(!DOMI_CLI)
        Text44 = SAFETEXT$(!LOCA_CLI)
        TXTDESCRIPROOVINCIA = VALOBADA("TAPROVINCIA", "DESCRI_PROV", "CODITEX= '" & SAFETEXT$(!PROV_CLI) & "'", "NOMESS")
        'Combo1(1).Text = SAFETEXT$(!TIPODOCU_ALUM)
        Text5 = SAFETEXT$(!SEXO_CLI)
        Text12.Text = SAFETEXT$(!DOCU_CLI)
        Text11.Text = SAFETEXT$(!Tele_Cli)
        Text10.Text = SAFETEXT$(!CELU_CLI)
        Text8.Text = SAFETEXT$(!MAIL_CLI)
        Text9.Text = SAFETEXT$(!CTAC_CLI)
        Text18.Text = SAFETEXT$(!CELCONTAC_CLI)
        TEXT13 = SAFETEXT$(!nota_Cli)
        Text6 = FECHATEX$(CVINT(!FECHNAC))
      
    End With
    ALUMTMP.Close
    Set ALUMTMP = Nothing
    
    


End Sub
Private Sub Text20_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text21_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text25_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text26_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text27_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text28_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text29_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub


Private Sub Text32_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

End Sub

'Private Sub TXCOMBO1_Change(Index As Integer)
'   TXCOMBO1(Index) = UCase(TXCOMBO1(Index))
'End Sub


'Private Sub TXTCODPROVINCIA_Change()
'     If TRIM$(Me.TXTCODPROVINCIA) <> "" Then
'         Me.TXTDESCRIPROOVINCIA = VALOBADA("TAPROVINCIA", "DESCRI_PROV", "CODITEX= '" & Me.TXTCODPROVINCIA & "'", "NOMESS")
'     Else
'         Me.TXTDESCRIPROOVINCIA = ""
'     End If
'End Sub


'Private Sub TXTRUTAFIRMA_Change()
'   Image1.Picture = LoadPicture("")
'   If EXISTE%(Me.TXTRUTAFIRMA) > 0 Then
'     On Error Resume Next
'     Image1.Picture = LoadPicture(TXTRUTAFIRMA)
'     On Error GoTo 0
'   Else
'
'   End If
'
'End Sub


Private Sub TXTRESPONSABLENUM_Change()

End Sub

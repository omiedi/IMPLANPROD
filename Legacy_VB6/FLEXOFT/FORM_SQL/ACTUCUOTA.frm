VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ACTUCUOTA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MODULO DE ACTUALIZACIÓN DE CUOTAS/MATRÍCULAS/DERECHOS DE EXAMEN"
   ClientHeight    =   7635
   ClientLeft      =   -15
   ClientTop       =   675
   ClientWidth     =   15600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
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
      TabIndex        =   12
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
      Top             =   0
      Width           =   15555
      Begin VB.CommandButton Command2 
         Caption         =   "Generación Masiva de Cuotas"
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
         Left            =   11760
         TabIndex        =   57
         Top             =   3360
         Width           =   3495
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
         Height          =   1335
         Left            =   120
         TabIndex        =   28
         Top             =   6240
         Width           =   15195
         Begin VB.CommandButton Command4 
            Caption         =   "Generación de Cuotas por Alumno"
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
            Left            =   10200
            TabIndex        =   58
            Top             =   360
            Width           =   4815
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
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   1650
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
            Index           =   10
            Left            =   315
            Locked          =   -1  'True
            TabIndex        =   55
            TabStop         =   0   'False
            Text            =   "Provinicia"
            Top             =   1650
            Width           =   1300
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
            TabIndex        =   54
            Top             =   210
            Width           =   1875
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
            TabIndex        =   53
            Top             =   210
            Width           =   255
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
            TabIndex        =   52
            TabStop         =   0   'False
            Text            =   "Legajo:"
            Top             =   210
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
            Index           =   12
            Left            =   7440
            Locked          =   -1  'True
            TabIndex        =   51
            TabStop         =   0   'False
            Text            =   "Sexo:"
            Top             =   570
            Width           =   650
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
            TabIndex        =   50
            TabStop         =   0   'False
            Text            =   "F. de Nacim.:"
            Top             =   3330
            Width           =   1100
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
            TabIndex        =   49
            TabStop         =   0   'False
            Text            =   "Nomb. y Ape.:"
            Top             =   570
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
            TabIndex        =   48
            TabStop         =   0   'False
            Text            =   "Domicilio:"
            Top             =   930
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
            Index           =   18
            Left            =   315
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            Text            =   "Localidad:"
            Top             =   1410
            Width           =   1300
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
            TabIndex        =   46
            Top             =   570
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
            TabIndex        =   45
            Top             =   930
            Width           =   3255
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
            TabIndex        =   44
            Top             =   1410
            Width           =   3255
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
            TabIndex        =   43
            Top             =   570
            Width           =   1605
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
            TabIndex        =   42
            Top             =   2010
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
            TabIndex        =   41
            Top             =   2370
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
            TabIndex        =   40
            Top             =   3090
            Width           =   3250
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
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   39
            Top             =   930
            Width           =   3250
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
            TabIndex        =   38
            TabStop         =   0   'False
            Text            =   "Télefono:"
            Top             =   2010
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
            TabIndex        =   37
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
            Index           =   2
            Left            =   315
            Locked          =   -1  'True
            TabIndex        =   36
            TabStop         =   0   'False
            Text            =   "Celular:"
            Top             =   2370
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
            TabIndex        =   35
            TabStop         =   0   'False
            Text            =   "Contacto:"
            Top             =   3090
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
            Index           =   5
            Left            =   5160
            Locked          =   -1  'True
            TabIndex        =   34
            TabStop         =   0   'False
            Text            =   "E-Mail:"
            Top             =   930
            Width           =   850
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
            TabIndex        =   33
            Top             =   3450
            Width           =   3250
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
            TabIndex        =   32
            TabStop         =   0   'False
            Text            =   "Cel. Contacto:"
            Top             =   3450
            Width           =   1300
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
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   570
            Width           =   1215
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
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   3330
            Width           =   1000
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
            TabIndex        =   29
            Top             =   1440
            Width           =   4245
         End
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
         Left            =   5640
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "Código Interno:"
         Top             =   240
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
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   1395
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Valor de la Cuota:"
         Top             =   240
         Width           =   1965
      End
      Begin VB.Frame Frame2 
         Height          =   600
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   7695
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "ACTUCUOTA.frx":0000
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
            Top             =   160
            Width           =   2535
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   3480
            OleObjectBlob   =   "ACTUCUOTA.frx":006E
            TabIndex        =   17
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
         Left            =   10440
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   645
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   600
         Width           =   675
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   1875
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Doble Click Para Seleccionar hacia la Grilla Inferior"
         Top             =   1320
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   3307
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
         Height          =   1755
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
         Top             =   1320
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   3096
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
      Begin MSFlexGridLib.MSFlexGrid MSF3 
         Height          =   2475
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "Doble Click Para Quitar de la Grilla"
         Top             =   3840
         Width           =   15195
         _ExtentX        =   26802
         _ExtentY        =   4366
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
      OleObjectBlob   =   "ACTUCUOTA.frx":00DA
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
   Begin VB.Menu mnuEdicion 
      Caption         =   "Edición"
      Begin VB.Menu mnuAnularFactura 
         Caption         =   "Anular Factura"
      End
   End
End
Attribute VB_Name = "ACTUCUOTA"
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
       MSF.TextMatrix(REG&, 15) = FORMATNUM$(XVALO(TXTCURSOTOT), "F12.2")
       MSF.TextMatrix(REG&, 16) = FORMATNUM$(XVALO(TXTMATRICULA2), "F12.2")
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



Sub EVENTOSELCHANGE(MSF As MSFlexGrid, REG1&)
'   REG1& = MSF.MouseRow
'   Call MOSTRAR_DATOS_FILA(MSF, REG1&)
   TXTMASID = XVALO(MSF.TextMatrix(REG1&, 18))

End Sub

Sub GeneracionAutomaticaCuota(CI1, NC, MESASEL$, DES%, PROCESOOK%)
   
   'CARGO UN VECTOR CON LOS VALORES DE LAS CUOTAS CON EL INDICE DEL CURSO
   Dim VALORCUOTA(), VALORDERECHOEXAMEN()
   MESNUM1$ = TRIM$(Mid$(FECHATEX$(DES%), 4, 2))
   PROCESOOK% = 1
   SQLX$ = "SELECT ValCurso,CURSOS_ID,ValDerExamen FROM CURSOS WITH(NOLOCK)"
   SQLX$ = SQLX$ + " ORDER BY CURSOS_ID"
   Set rs = FLEXCONN.Execute(SQLX$)
   With rs
     Do While Not .EOF
        IDCUR& = XVALO(!CURSOS_ID)
        ReDim Preserve VALORCUOTA(IDCUR&)
        ReDim Preserve VALORDERECHOEXAMEN(IDCUR&)
        VALORCUOTA(IDCUR&) = XVALO(!ValCurso)
        VALORDERECHOEXAMEN(IDCUR&) = XVALO(!ValDerExamen)
        .MoveNext
     Loop
    End With
    rs.Close
    Set rs = Nothing
    'FIN CARGA DE VECOR DEL VALOR DE LAS CUOTAS
   '
   SQLX$ = "select M.Codint , M.FeIniCurs,M.FeFinCurs, M.CantMeses,A.NUME_CLI, C.VALCURSO  "
   SQLX$ = SQLX$ + " ,A.CURSOS_ID ,M.ID_CURSO, M.FeIniCurs,M.DESCRICURSO,M.CantMeses "
   SQLX$ = SQLX$ + " from alumno_curso a  WITH(NOLOCK) inner join master  m WITH(NOLOCK)  on m.codint = a.cursos_id inner  join cursos  c WITH(NOLOCK)"
   SQLX$ = SQLX$ + " on m.id_curso = c.cursos_id"
   SQLX$ = SQLX$ + " Where datediff(D, m.FeFinCurs, GETDATE()) <= 28 AND (M.STATUS >= 0 OR M.STATUS IS NULL) AND A.STATUS >= 0" ' SI SE DESAPLICO EL CURSO
   'SQLX$ = SQLX$ + " ON M.ID_CURSO = A.CURSOS_ID "
   If XVALO(CI1) > 0 Then SQLX$ = SQLX$ + " AND M.CODINT = " & CI1
   If XVALO(NC) > 0 Then SQLX$ = SQLX$ + " AND A.NUME_CLI = " & NC
   SQLX$ = SQLX$ + " ORDER BY A.NUME_CLI "
   Set rs = READSET(SQLX$)
   With rs
   Do While Not .EOF
      'VER SI HAY QUE VALIDAR QUE EL CLIENTE NO SE REPITA(ALUMNO)
      LEGAJO& = XVALO(!NUME_CLI)
      IDCUR& = XVALO(!ID_CURSO)
      CANTICUOTAS = XVALO(!CantMeses) 'CANTIDAD DE CUOTAS DEL CURSO
      CI1 = XVALO(!CODINT)
      FEINICIO% = CVINT(!feinicurs)
      DESCRICURSO$ = SAFETEXT$(!DESCRICURSO)
      '*** OJO QUE ESTA CONDI$ ** SE CONCATENA ABAJO AGREGANDO UNA CONDICION PARA OTRA NUEVA CONSULTA *****************************
      CONDI$ = "(left(IDENMOVI,5) = 'CUOTA' OR left(IDENMOVI,13) = 'CUOTA INICIAL') AND COD_INTE = " & CI1 & " AND NUME_CLIE = " & LEGAJO&
      CANT_CTAS_GENER_MOVISTO% = CantRegistros&("MOVISTO", CONDI$, "NOMESS")
      CONDI1$ = "Nume_Clie = " & LEGAJO& & " AND Cod_Inte = " & CI1 & " AND LEFT(IDENMOVI,5 )= 'CUOTA'"
      On Error GoTo ERROR_GUARDAR
      '
      'LE CAMBIO EL MES A LA FECAHA DE INICIO PARA QUE  FECHEMISI COINCIDA CON EL MISMO NNUMERO DE FECHA PERO CON EL  MES QUE SE ESTA CORRIENDO.
      FECHEMISIX% = FEINICIO%
      FECHA_CTA_TEX$ = FECHATEX$(FEINICIO%)
      CANTIDIAS_MES% = DiasDelMes(FETEXCOR1$(DES%))
35    If XVALO(Left(FECHA_CTA_TEX$, 2)) > CANTIDIAS_MES% Then
         FECHEMISIX% = DIANTE(FECHEMISIX%, 1): FECHA_CTA_TEX$ = FECHATEX$(FECHEMISIX%): GoTo 35
      End If
      FECHA_CTA_TEX$ = Left(FECHA_CTA_TEX$, 3) & MESNUM1$ & Right(FECHA_CTA_TEX$, 3)
      FECHEMISIX% = FECHANUM(FECHA_CTA_TEX$)
      '
      '******************************************************
      'ESTE BLOQUE CALCULA A PARTIR DE CUANDO EMPEZO EL CURSO QUE NUMERO DE CUOTA CORRESPONDE SEGUN
      'EL MES SELECCIONADO POR SI REALIZA GENERACION DE CUOTAS SIN CORRELATIVIDAD
      FEX% = FEINICIO%
      FEXX$ = FECHATEX(FEX%)
      For i& = 1 To CANTICUOTAS
         If Right(FEXX$, 5) = Right(FECHA_CTA_TEX$, 5) Then
            INDICE_CTA& = i&
            Exit For
         End If
         FEXX$ = Format(DateAdd("M", 1, FETEXCOR1$(FEX%)), "DD-MM-YY")
         FEX% = FECHANUM(FEXX$)
      Next i&
      If INDICE_CTA& < 1 Then GoTo 75 'NEXT 'SI NO ENCUENTRA EL INDICE ES POR QUE YA NO LE CORRESPONDE
      'AQUI TERMINA CALCULO DE NUMERO DE CUOTA.
      '******************************************************
      '
'      If CANT_CTAS_GENER_MOVISTO% < CANTICUOTAS Then
         CUOTAX$ = "CUOTA " & MESASEL$
         CONDI$ = "(left(IDENMOVI,5) = 'CUOTA' OR left(IDENMOVI,13) = 'CUOTA INICIAL') AND COD_INTE = " & CI1 & " AND NUME_CLIE = " & LEGAJO&
         CONDI$ = CONDI$ + " AND IDENLOTE =  '" & MESASEL$ & "'"
         If CantRegistros&("MOVISTO", CONDI$, "NOMESS") < 1 Then
            'SI NO EXISTE LA AGREGO
            CONDI$ = "ID_CURSO = " & IDCUR&
            VAL_CUOTA# = VALORCUOTA(IDCUR&)
            FLEXCONN.BeginTrans
            Call GRABACUOTA(LEGAJO&, DESCRICURSO$, 2, "B", 1, "CUOTA " & MESASEL$, VAL_CUOTA#, "AUTO", RET_NUMFACTURALAR$, RET_NUMFACTUCOR$, FECHEMISIX%)
            NPVX$ = "0001"
            CIX% = CI1
            Call NUREMIFA(NROREMIT&, NPVX$)
            CUOTA_DE$ = "CUOTA " & INDICE_CTA& & "/" & CANTICUOTAS
            Call MOVISTOK(HOY(HO$), CIX%, MESASEL$, "RT", "RT-0001-" + FORMATOCON0(NROREMIT&, 8), "RT-" & NROREMIT&, RET_NUMFACTURALAR$, RET_NUMFACTUCOR$, CUOTAX$, 1, VAL_CUOTA#, 1, LEGAJO&, 1, 1, 0, 1, DESCRICURSO$, "", CUOTA_DE$, VAL_CUOTA#, , , INDICE_CTA&)
            Call ANOTAX_ALUMNO(LEGAJO&, "SE GENERÓ CUOTA " & MESASEL$ & " " & CUOTA_DE$ & " POR IMPORTE: " & FORMATNUM$(VAL_CUOTA#, "F12.2"))
            FLEXCONN.CommitTrans
         Else
            'SI ENTRA ACA POR QUE ESA CUOTA YA EXISTE EN EL MOVISTO
            FACTURAMOVISTO$ = VALOBADA("MOVISTO", "DOSECLAR", CONDI$, "NOMESS")
            IDMOVI = XVALO(VALOBADA("MOVISTO", "MOVISTO_ID", CONDI$, "NOMESS"))
            CONDI$ = "CODOCOR_LAR = '" & FACTURAMOVISTO$ & "'"
            If CantRegistros&("APLICOBRA", CONDI$, "NOMESS") < 1 Then
               'NO TIENE APLICACIONES
               FLEXCONN.BeginTrans
               'ACTUALIZO EL CAJABANC
               SQLX$ = "UPDATE CAJABANC SET IDEBECOMP = " & VALORCUOTA(IDCUR&)
               SQLX$ = SQLX$ + " , IMPO_DOLA = " & VALORCUOTA(IDCUR&)
               SQLX$ = SQLX$ + " , VALABSCOMP = " & VALORCUOTA(IDCUR&)
               SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'FVEN' AND CODICOM_LAR = '" & FACTURAMOVISTO$ & "'"
               FLEXCONN.Execute (SQLX$)
               
               SQLX$ = "UPDATE CAJABANC SET IHABECOMP = " & VALORCUOTA(IDCUR&)
               SQLX$ = SQLX$ + " , IMPO_DOLA = " & VALORCUOTA(IDCUR&)
               SQLX$ = SQLX$ + " , VALABSCOMP = " & VALORCUOTA(IDCUR&)
               SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'VENTA' AND CODICOM_LAR = '" & FACTURAMOVISTO$ & "'"
               FLEXCONN.Execute (SQLX$)
               '
               'ACTUALIZO EL MOVISTO
               SQLX$ = "UPDATE MOVISTO SET ValoUnit = " & VALORCUOTA(IDCUR&)
               SQLX$ = SQLX$ + " , VALOUNITIVA = " & VALORCUOTA(IDCUR&)
               SQLX$ = SQLX$ + " , VALTOTIT = " & VALORCUOTA(IDCUR&)
               SQLX$ = SQLX$ + " , ValTotItIva = " & VALORCUOTA(IDCUR&)
               SQLX$ = SQLX$ + " WHERE MOVISTO_ID = " & IDMOVI
               FLEXCONN.Execute (SQLX$)
               Call ANOTAX_ALUMNO(LEGAJO&, "SE ACTUALIZO CUOTA " & MESASEL$ & " " & CUOTA_DE$ & " POR IMPORTE: " & FORMATNUM$(VAL_CUOTA#, "F12.2"))
               '
               FLEXCONN.CommitTrans
            End If
         End If
         'ACA GENERA EL DERECHO DE EXAMEN CUANDO EN EL CASO QUE LA CUOTRA GENERADA SEA LA ULTIMA.
         If INDICE_CTA& = CANTICUOTAS Then
            CUOTAX$ = "CUOTA " & MESASEL$
            CONDI$ = "left(IDENMOVI,8) = 'D.EXAMEN'  AND COD_INTE = " & CI1 & " AND NUME_CLIE = " & LEGAJO&
             If CantRegistros&("MOVISTO", CONDI$, "NOMESS") < 1 Then
                'SI NO EXISTE LA AGREGO
                CONDI$ = "ID_CURSO = " & IDCUR&
                VAL_CUOTA# = VALORDERECHOEXAMEN(IDCUR&)
                FLEXCONN.BeginTrans
                Call GRABACUOTA(LEGAJO&, DESCRICURSO$, 2, "B", 1, "D.EXAMEN " & MESASEL$, VAL_CUOTA#, "AUTO", RET_NUMFACTURALAR$, RET_NUMFACTUCOR$, FECHEMISIX%)
                NPVX$ = "0001"
                CIX% = CI1
                Call NUREMIFA(NROREMIT&, NPVX$)
                CUOTA_DE$ = "CUOTA " & INDICE_CTA& & "/" & CANTICUOTAS
                Call MOVISTOK(HOY(HO$), CIX%, MESASEL$, "RT", "RT-0001-" + FORMATOCON0(NROREMIT&, 8), "RT-" & NROREMIT&, RET_NUMFACTURALAR$, RET_NUMFACTUCOR$, "D.EXAMEN " & DESCRICURSO, 1, VAL_CUOTA#, 1, LEGAJO&, 1, 1, 0, 1, DESCRICURSO$, "", "D.EXAMEN " & DESCRICURSO$, VAL_CUOTA#, , , INDICE_CTA&)
                Call ANOTAX_ALUMNO(LEGAJO&, "SE GENERÓ DERECHO DE EXAMEN  POR IMPORTE: " & FORMATNUM$(VAL_CUOTA#, "F12.2"))

                FLEXCONN.CommitTrans
             Else
                'SI ENTRA ACA POR QUE ESA CUOTA YA EXISTE EN EL MOVISTO
                FACTURAMOVISTO$ = VALOBADA("MOVISTO", "DOSECLAR", CONDI$, "NOMESS")
                IDMOVI = XVALO(VALOBADA("MOVISTO", "MOVISTO_ID", CONDI$, "NOMESS"))
                CONDI$ = "CODOCOR_LAR = '" & FACTURAMOVISTO$ & "'"
                If CantRegistros&("APLICOBRA", CONDI$, "NOMESS") < 1 Then
                   'NO TIENE APLICACIONES
                   FLEXCONN.BeginTrans
                   'ACTUALIZO EL CAJABANC
                   SQLX$ = "UPDATE CAJABANC SET IDEBECOMP = " & VALORCUOTA(IDCUR&)
                   SQLX$ = SQLX$ + " , IMPO_DOLA = " & VALORCUOTA(IDCUR&)
                   SQLX$ = SQLX$ + " , VALABSCOMP = " & VALORCUOTA(IDCUR&)
                   SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'FVEN' AND CODICOM_LAR = '" & FACTURAMOVISTO$ & "'"
                   FLEXCONN.Execute (SQLX$)
                   '
                   SQLX$ = "UPDATE CAJABANC SET IHABECOMP = " & VALORCUOTA(IDCUR&)
                   SQLX$ = SQLX$ + " , IMPO_DOLA = " & VALORCUOTA(IDCUR&)
                   SQLX$ = SQLX$ + " , VALABSCOMP = " & VALORCUOTA(IDCUR&)
                   SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'VENTA' AND CODICOM_LAR = '" & FACTURAMOVISTO$ & "'"
                   FLEXCONN.Execute (SQLX$)
                   '
                   'ACTUALIZO EL MOVISTO
                   SQLX$ = "UPDATE MOVISTO SET ValoUnit = " & VALORCUOTA(IDCUR&)
                   SQLX$ = SQLX$ + " , VALOUNITIVA = " & VALORCUOTA(IDCUR&)
                   SQLX$ = SQLX$ + " , VALTOTIT = " & VALORCUOTA(IDCUR&)
                   SQLX$ = SQLX$ + " , ValTotItIva = " & VALORCUOTA(IDCUR&)
                   SQLX$ = SQLX$ + " WHERE MOVISTO_ID = " & IDMOVI
                   FLEXCONN.Execute (SQLX$)
                   '
                   FLEXCONN.CommitTrans
                End If
             End If
         End If
'      End If
75   .MoveNext
   Loop
   End With
   rs.Close
   Set rs = Nothing
   GoTo 99
   
   
ERROR_GUARDAR:
    If Err <> 0 Then
      FLEXCONN.RollbackTrans
      MsgBox "ERROR AL GUARDAR NO SE REALIZO LA TRANSACCIÓN " & vbCrLf & Err.Number & " " & Err.Description
      PROCESOOK% = -1
      GoTo 99
    End If
   
'UNA FALLA PROBABLE DE ESTE PROCESO ES QUE EN UN CASO  SE HAYA SALTEADO CORRERLO PARA UN MES Y COMO CUANDO SE VUELVA A CORRER PARA OTRO MES
'SI NO ALCANZA A LA CANTIDAD DE CUOTAS INFORMADAS POR EJEMPLO UN CURSO QUE DEBERIA GENERAR CUOTAS ENE-FEB-MAR-ABRIL GENERE ENE-MAR-ABRIL-MAYO.
'ES DIFICIL POR QUE SE VA A CORRER TODOS LOS MESES.
   
99
 
End Sub

Sub GeneracionAutomaticaCuota_original(CI1, NC, MESASEL$, DES%, PROCESOOK%)
   
   'CARGO UN VECTOR CON LOS VALORES DE LAS CUOTAS CON EL INDICE DEL CURSO
   Dim VALORCUOTA(), VALORDERECHOEXAMEN()
   MESNUM1$ = TRIM$(Mid$(FECHATEX$(DES%), 4, 2))
   PROCESOOK% = 1
   SQLX$ = "SELECT ValCurso,CURSOS_ID,ValDerExamen FROM CURSOS WITH(NOLOCK)"
   SQLX$ = SQLX$ + " ORDER BY CURSOS_ID"
   Set rs = FLEXCONN.Execute(SQLX$)
   With rs
     Do While Not .EOF
        IDCUR& = XVALO(!CURSOS_ID)
        ReDim Preserve VALORCUOTA(IDCUR&)
        ReDim Preserve VALORDERECHOEXAMEN(IDCUR&)
        VALORCUOTA(IDCUR&) = XVALO(!ValCurso)
        VALORDERECHOEXAMEN(IDCUR&) = XVALO(!ValDerExamen)
        .MoveNext
     Loop
    End With
    rs.Close
    Set rs = Nothing
    'FIN CARGA DE VECOR DEL VALOR DE LAS CUOTAS
   '
   SQLX$ = "SELECT M.Codint , M.FeIniCurs,M.FeFinCurs, M.CantMeses,A.NUME_CLI, C.VALCURSO, A.CURSOS_ID"
   SQLX$ = SQLX$ + " ,M.ID_CURSO, M.FeIniCurs,M.DESCRICURSO,M.CantMeses FROM  MASTER M WITH(NOLOCK)"
   SQLX$ = SQLX$ + " INNER JOIN (ALUMNO_CURSO A WITH(NOLOCK)"
   SQLX$ = SQLX$ + " INNER JOIN  CURSOS C WITH(NOLOCK) ON A.CURSOS_ID = C.CURSOS_ID)"
   SQLX$ = SQLX$ + " ON M.CODINT = A.CURSOS_ID "
   'SQLX$ = SQLX$ + " ON M.ID_CURSO = A.CURSOS_ID "
   SQLX$ = SQLX$ + " Where datediff(D, m.FeFinCurs, GETDATE()) <= 28 AND (M.STATUS >= 0 OR M.STATUS IS NULL)"
   SQLX$ = SQLX$ + " AND A.STATUS >= 0 " ' SI SE DESAPLICO EL CURSO
'VER COMO VALIDAR QUE SI SE ELIGE UN MES ANTERIOR PARA EL COMIENZO DE CURSO NO GENERE UNA CUOTA ANTERIOR A LA GENERADA
'INICIALMENTE.
   If XVALO(CI1) > 0 Then SQLX$ = SQLX$ + " AND M.CODINT = " & CI1
   If XVALO(NC) > 0 Then SQLX$ = SQLX$ + " AND A.NUME_CLI = " & NC
   SQLX$ = SQLX$ + " ORDER BY A.NUME_CLI "
   Set rs = READSET(SQLX$)
   With rs
   Do While Not .EOF
      'VER SI HAY QUE VALIDAR QUE EL CLIENTE NO SE REPITA(ALUMNO)
      LEGAJO& = XVALO(!NUME_CLI)
      IDCUR& = XVALO(!ID_CURSO)
      CANTICUOTAS = XVALO(!CantMeses) 'CANTIDAD DE CUOTAS DEL CURSO
      CI1 = XVALO(!CODINT)
      FEINICIO% = CVINT(!feinicurs)
      DESCRICURSO$ = SAFETEXT$(!DESCRICURSO)
      '*** OJO QUE ESTA CONDI$ ** SE CONCATENA ABAJO AGREGANDO UNA CONDICION PARA OTRA NUEVA CONSULTA *****************************
      CONDI$ = "(left(IDENMOVI,5) = 'CUOTA' OR left(IDENMOVI,13) = 'CUOTA INICIAL') AND COD_INTE = " & CI1 & " AND NUME_CLIE = " & LEGAJO&
      CANT_CTAS_GENER_MOVISTO% = CantRegistros&("MOVISTO", CONDI$, "NOMESS")
      CONDI1$ = "Nume_Clie = " & LEGAJO& & " AND Cod_Inte = " & CI1 & " AND LEFT(IDENMOVI,5 )= 'CUOTA'"
      On Error GoTo ERROR_GUARDAR
      '
      'LE CAMBIO EL MES A LA FECAHA DE INICIO PARA QUE  FECHEMISI COINCIDA CON EL MISMO NNUMERO DE FECHA PERO CON EL  MES QUE SE ESTA CORRIENDO.
      FECHEMISIX% = FEINICIO%
      FECHA_CTA_TEX$ = FECHATEX$(FEINICIO%)
      CANTIDIAS_MES% = DiasDelMes(FETEXCOR1$(DES%))
35    If XVALO(Left(FECHA_CTA_TEX$, 2)) > CANTIDIAS_MES% Then
         FECHEMISIX% = DIANTE(FECHEMISIX%, 1): FECHA_CTA_TEX$ = FECHATEX$(FECHEMISIX%): GoTo 35
      End If
      FECHA_CTA_TEX$ = Left(FECHA_CTA_TEX$, 3) & MESNUM1$ & Right(FECHA_CTA_TEX$, 3)
      FECHEMISIX% = FECHANUM(FECHA_CTA_TEX$)
      '
      '******************************************************
      'ESTE BLOQUE CALCULA A PARTIR DE CUANDO EMPEZO EL CURSO QUE NUMERO DE CUOTA CORRESPONDE SEGUN
      'EL MES SELECCIONADO POR SI REALIZA GENERACION DE CUOTAS SIN CORRELATIVIDAD
      FEX% = FEINICIO%
      FEXX$ = FECHATEX(FEX%)
      For i& = 1 To CANTICUOTAS
         If Right(FEXX$, 5) = Right(FECHA_CTA_TEX$, 5) Then
            INDICE_CTA& = i&
            Exit For
         End If
         FEXX$ = Format(DateAdd("M", 1, FETEXCOR1$(FEX%)), "DD-MM-YY")
         FEX% = FECHANUM(FEXX$)
      Next i&
      If INDICE_CTA& < 1 Then GoTo 75 'NEXT 'SI NO ENCUENTRA EL INDICE ES POR QUE YA NO LE CORRESPONDE
      'AQUI TERMINA CALCULO DE NUMERO DE CUOTA.
      '******************************************************
      '
'      If CANT_CTAS_GENER_MOVISTO% < CANTICUOTAS Then
         CUOTAX$ = "CUOTA " & MESASEL$
         CONDI$ = "(left(IDENMOVI,5) = 'CUOTA' OR left(IDENMOVI,13) = 'CUOTA INICIAL') AND COD_INTE = " & CI1 & " AND NUME_CLIE = " & LEGAJO&
         CONDI$ = CONDI$ + " AND IDENLOTE =  '" & MESASEL$ & "'"
         If CantRegistros&("MOVISTO", CONDI$, "NOMESS") < 1 Then
            'SI NO EXISTE LA AGREGO
            CONDI$ = "ID_CURSO = " & IDCUR&
            VAL_CUOTA# = VALORCUOTA(IDCUR&)
            FLEXCONN.BeginTrans
            Call GRABACUOTA(LEGAJO&, DESCRICURSO$, 2, "B", 1, "CUOTA " & MESASEL$, VAL_CUOTA#, "AUTO", RET_NUMFACTURALAR$, RET_NUMFACTUCOR$, FECHEMISIX%)
            NPVX$ = "0001"
            CIX% = CI1
            Call NUREMIFA(NROREMIT&, NPVX$)
            CUOTA_DE$ = "CUOTA " & INDICE_CTA& & "/" & CANTICUOTAS
            Call MOVISTOK(HOY(HO$), CIX%, MESASEL$, "RT", "RT-0001-" + FORMATOCON0(NROREMIT&, 8), "RT-" & NROREMIT&, RET_NUMFACTURALAR$, RET_NUMFACTUCOR$, CUOTAX$, 1, VAL_CUOTA#, 1, LEGAJO&, 1, 1, 0, 1, DESCRICURSO$, "", CUOTA_DE$, VAL_CUOTA#, , , INDICE_CTA&)
            Call ANOTAX_ALUMNO(LEGAJO&, "SE GENERÓ CUOTA " & MESASEL$ & " " & CUOTA_DE$ & " POR IMPORTE: " & FORMATNUM$(VAL_CUOTA#, "F12.2"))
            FLEXCONN.CommitTrans
         Else
            'SI ENTRA ACA POR QUE ESA CUOTA YA EXISTE EN EL MOVISTO
            FACTURAMOVISTO$ = VALOBADA("MOVISTO", "DOSECLAR", CONDI$, "NOMESS")
            IDMOVI = XVALO(VALOBADA("MOVISTO", "MOVISTO_ID", CONDI$, "NOMESS"))
            CONDI$ = "CODOCOR_LAR = '" & FACTURAMOVISTO$ & "'"
            If CantRegistros&("APLICOBRA", CONDI$, "NOMESS") < 1 Then
               'NO TIENE APLICACIONES
               FLEXCONN.BeginTrans
               'ACTUALIZO EL CAJABANC
               SQLX$ = "UPDATE CAJABANC SET IDEBECOMP = " & VALORCUOTA(IDCUR&)
               SQLX$ = SQLX$ + " , IMPO_DOLA = " & VALORCUOTA(IDCUR&)
               SQLX$ = SQLX$ + " , VALABSCOMP = " & VALORCUOTA(IDCUR&)
               SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'FVEN' AND CODICOM_LAR = '" & FACTURAMOVISTO$ & "'"
               FLEXCONN.Execute (SQLX$)
               
               SQLX$ = "UPDATE CAJABANC SET IHABECOMP = " & VALORCUOTA(IDCUR&)
               SQLX$ = SQLX$ + " , IMPO_DOLA = " & VALORCUOTA(IDCUR&)
               SQLX$ = SQLX$ + " , VALABSCOMP = " & VALORCUOTA(IDCUR&)
               SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'VENTA' AND CODICOM_LAR = '" & FACTURAMOVISTO$ & "'"
               FLEXCONN.Execute (SQLX$)
               '
               'ACTUALIZO EL MOVISTO
               SQLX$ = "UPDATE MOVISTO SET ValoUnit = " & VALORCUOTA(IDCUR&)
               SQLX$ = SQLX$ + " , VALOUNITIVA = " & VALORCUOTA(IDCUR&)
               SQLX$ = SQLX$ + " , VALTOTIT = " & VALORCUOTA(IDCUR&)
               SQLX$ = SQLX$ + " , ValTotItIva = " & VALORCUOTA(IDCUR&)
               SQLX$ = SQLX$ + " WHERE MOVISTO_ID = " & IDMOVI
               FLEXCONN.Execute (SQLX$)
               Call ANOTAX_ALUMNO(LEGAJO&, "SE ACTUALIZO CUOTA " & MESASEL$ & " " & CUOTA_DE$ & " POR IMPORTE: " & FORMATNUM$(VAL_CUOTA#, "F12.2"))
               '
               FLEXCONN.CommitTrans
            End If
         End If
         'ACA GENERA EL DERECHO DE EXAMEN CUANDO EN EL CASO QUE LA CUOTRA GENERADA SEA LA ULTIMA.
         If INDICE_CTA& = CANTICUOTAS Then
            CUOTAX$ = "CUOTA " & MESASEL$
            CONDI$ = "left(IDENMOVI,8) = 'D.EXAMEN'  AND COD_INTE = " & CI1 & " AND NUME_CLIE = " & LEGAJO&
             If CantRegistros&("MOVISTO", CONDI$, "NOMESS") < 1 Then
                'SI NO EXISTE LA AGREGO
                CONDI$ = "ID_CURSO = " & IDCUR&
                VAL_CUOTA# = VALORDERECHOEXAMEN(IDCUR&)
                FLEXCONN.BeginTrans
                Call GRABACUOTA(LEGAJO&, DESCRICURSO$, 2, "B", 1, "D.EXAMEN " & MESASEL$, VAL_CUOTA#, "AUTO", RET_NUMFACTURALAR$, RET_NUMFACTUCOR$, FECHEMISIX%)
                NPVX$ = "0001"
                CIX% = CI1
                Call NUREMIFA(NROREMIT&, NPVX$)
                CUOTA_DE$ = "CUOTA " & INDICE_CTA& & "/" & CANTICUOTAS
                Call MOVISTOK(HOY(HO$), CIX%, MESASEL$, "RT", "RT-0001-" + FORMATOCON0(NROREMIT&, 8), "RT-" & NROREMIT&, RET_NUMFACTURALAR$, RET_NUMFACTUCOR$, "D.EXAMEN " & DESCRICURSO, 1, VAL_CUOTA#, 1, LEGAJO&, 1, 1, 0, 1, DESCRICURSO$, "", "D.EXAMEN " & DESCRICURSO$, VAL_CUOTA#, , , INDICE_CTA&)
                Call ANOTAX_ALUMNO(LEGAJO&, "SE GENERÓ DERECHO DE EXAMEN  POR IMPORTE: " & FORMATNUM$(VAL_CUOTA#, "F12.2"))

                FLEXCONN.CommitTrans
             Else
                'SI ENTRA ACA POR QUE ESA CUOTA YA EXISTE EN EL MOVISTO
                FACTURAMOVISTO$ = VALOBADA("MOVISTO", "DOSECLAR", CONDI$, "NOMESS")
                IDMOVI = XVALO(VALOBADA("MOVISTO", "MOVISTO_ID", CONDI$, "NOMESS"))
                CONDI$ = "CODOCOR_LAR = '" & FACTURAMOVISTO$ & "'"
                If CantRegistros&("APLICOBRA", CONDI$, "NOMESS") < 1 Then
                   'NO TIENE APLICACIONES
                   FLEXCONN.BeginTrans
                   'ACTUALIZO EL CAJABANC
                   SQLX$ = "UPDATE CAJABANC SET IDEBECOMP = " & VALORCUOTA(IDCUR&)
                   SQLX$ = SQLX$ + " , IMPO_DOLA = " & VALORCUOTA(IDCUR&)
                   SQLX$ = SQLX$ + " , VALABSCOMP = " & VALORCUOTA(IDCUR&)
                   SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'FVEN' AND CODICOM_LAR = '" & FACTURAMOVISTO$ & "'"
                   FLEXCONN.Execute (SQLX$)
                   '
                   SQLX$ = "UPDATE CAJABANC SET IHABECOMP = " & VALORCUOTA(IDCUR&)
                   SQLX$ = SQLX$ + " , IMPO_DOLA = " & VALORCUOTA(IDCUR&)
                   SQLX$ = SQLX$ + " , VALABSCOMP = " & VALORCUOTA(IDCUR&)
                   SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'VENTA' AND CODICOM_LAR = '" & FACTURAMOVISTO$ & "'"
                   FLEXCONN.Execute (SQLX$)
                   '
                   'ACTUALIZO EL MOVISTO
                   SQLX$ = "UPDATE MOVISTO SET ValoUnit = " & VALORCUOTA(IDCUR&)
                   SQLX$ = SQLX$ + " , VALOUNITIVA = " & VALORCUOTA(IDCUR&)
                   SQLX$ = SQLX$ + " , VALTOTIT = " & VALORCUOTA(IDCUR&)
                   SQLX$ = SQLX$ + " , ValTotItIva = " & VALORCUOTA(IDCUR&)
                   SQLX$ = SQLX$ + " WHERE MOVISTO_ID = " & IDMOVI
                   FLEXCONN.Execute (SQLX$)
                   '
                   FLEXCONN.CommitTrans
                End If
             End If
         End If
'      End If
75   .MoveNext
   Loop
   End With
   rs.Close
   Set rs = Nothing
   GoTo 99
   
   
ERROR_GUARDAR:
    If Err <> 0 Then
      FLEXCONN.RollbackTrans
      MsgBox "ERROR AL GUARDAR NO SE REALIZO LA TRANSACCIÓN " & vbCrLf & Err.Number & " " & Err.Description
      PROCESOOK% = -1
      GoTo 99
    End If
   
'UNA FALLA PROBABLE DE ESTE PROCESO ES QUE EN UN CASO  SE HAYA SALTEADO CORRERLO PARA UN MES Y COMO CUANDO SE VUELVA A CORRER PARA OTRO MES
'SI NO ALCANZA A LA CANTIDAD DE CUOTAS INFORMADAS POR EJEMPLO UN CURSO QUE DEBERIA GENERAR CUOTAS ENE-FEB-MAR-ABRIL GENERE ENE-MAR-ABRIL-MAYO.
'ES DIFICIL POR QUE SE VA A CORRER TODOS LOS MESES.
   
99
 
End Sub


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

End Sub

Private Sub Command2_Click()
   If DERACCE%("GENECUTA", "GENERACIÓN MASIVAS DE CUOTAS") < 2 Then Exit Sub
   Call CARSELMESA(7)
   
   Call SELECHO("SELMESA")
   MESASEL$ = TRIM$(VALACT1$("SELMESA"))
   If MESASEL$ = "" Then
        Call COMUNI("Debe Seleccionar un Mes")
        GoTo 99
   End If
   '
   MMXX0$ = TRIM$(MESASEL$)
   MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
   MESASEL$ = METEXTO$(DES%) ' TIENE QUE TENER EL FORMATO: FEBRERO 2018 Y NO FEBRERO 18
   
   '
   Screen.MousePointer = 11
'   Dim CURSOSSELC()
   VUELTA% = 0

   Command2.Enabled = False
   HAYFILTROCURSO% = 0
   'VERIFICO SI HAY SELECCIONADO CURSOS
   If MSF3.Rows > 1 Then
      HAYFILTROCURSO% = 1
      For i& = 1 To MSF3.Rows - 1
        CI1 = MSF3.TextMatrix(i&, 18) 'ES EL CODINT DEL CURSO
        If CI1 > 0 Then
         Call GeneracionAutomaticaCuota(CI1, 0, MESASEL$, DES%, PROCESOOK%)
         CANTICURSOSSEL& = i&
        End If
      Next i&
   End If
   
   If HAYFILTROCURSO% < 1 Then Call GeneracionAutomaticaCuota(0, 0, MESASEL$, DES%, PROCESOOK%)
   Screen.MousePointer = 1
   If PROCESOOK% >= 0 Then
     Call COMUNI("Proceso Terminado")
   End If
   
99 Command2.Enabled = True
End Sub

Private Sub Command4_Click()
   If DERACCE%("GENECUT1", "GENERACIÓN DE CUOTAS X ALUMNO") < 2 Then Exit Sub
   Command4.Enabled = False
   
   NC = XVALO(Text2)
   If NC < 1 Then
     Call COMUNI("Falta Seleccionar Alumno al Cual Generar la Cuota")
     GoTo 99
   End If
   
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = TRIM$(VALACT1$("SELMESA"))
   If MESASEL$ = "" Then
        Call COMUNI("Debe Seleccionar un Mes")
        GoTo 99
   End If
   '
   MMXX0$ = TRIM$(MESASEL$)
   MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
   MESASEL$ = METEXTO$(DES%) ' TIENE QUE TENER EL FORMATO: FEBRERO 2018 Y NO FEBRERO 18
   '
   Screen.MousePointer = 11
'   Dim CURSOSSELC()
   VUELTA% = 0

   Command2.Enabled = False
   HAYFILTROCURSO% = 0
   'VERIFICO SI HAY SELECCIONADO CURSOS
   If MSF3.Rows > 1 Then
      HAYFILTROCURSO% = 1
      For i& = 1 To MSF3.Rows - 1
        CI1 = MSF3.TextMatrix(i&, 18) 'ES EL CODINT DEL CURSO
        If CI1 > 0 Then
         Call GeneracionAutomaticaCuota(CI1, NC, MESASEL$, DES%, PROCESOOK%)
         CANTICURSOSSEL& = i&
        End If
      Next i&
   End If
   
   If HAYFILTROCURSO% < 1 Then Call GeneracionAutomaticaCuota(0, NC, MESASEL$, DES%, PROCESOOK%)
   Screen.MousePointer = 1
   If PROCESOOK% >= 0 Then
     Call COMUNI("Proceso Terminado")
   End If
   
99 Command4.Enabled = True
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
    Campox$ = "F.Inicio/D8;F.Fin/D8;C/F1;Días/A10;C/F1;Días/A10;H.Desde/H10;H.Hasta/H10;Id Prof/F8;Profesor/A32;C.A./f4;Aula/A18;Val.Curso/f10.2;Val.Matrícula/F11.2;Cant.Meses/F10;Capacidad/F10;Ord/F0;CODINT/F7"
    Call CARGA_ENCABEZADO(Me.MSF, Campox$, Me, 1)
    Call CARGA_ENCABEZADO(Me.MSF_2, Campox$, Me, 0)
    Call CARGA_ENCABEZADO(Me.MSF3, Campox$, Me, 0)
    
    MSF.Rows = 2
    'PROPIEDADES DE SELECCION DE FILA
    MSF.SelectionMode = flexSelectionByRow
    MSF.FocusRect = flexFocusNone
    '

End Sub





Private Sub mnuAnularFactura_Click()
  Call ANUFAC
End Sub

Sub ANUFAC()
 'ANULACION DE FACTURAS YA EMITIDAS
    '
    If DERACCE%("ANUFAEM", "Anulación de Facturas Emitidas") < 2 Then
      Exit Sub
    End If
    '
    FECHALIM% = XVALO(CONTROL$("", "CIEMEVEN")) '\\\OT-06-0212-WAR-06/06/06///
    FECHALU = FETEXCOR1$(FECHALIM%)
    
    '
10  If ALUMONOSEL > 0 Then
      NCLI = XVALO(RESP_ZELE_VECT(1))
    Else
      Exit Sub
    End If

'    NCLI = XVALO(VALACT1$("DEUDOR1"))
    If NCLI < 1 Then Exit Sub
    '
    Call APERBASDAT("CABAN")
15  SQLX$ = " select  c.NuCLIEN, c.codicom_lar,c.obsergen, m.idenmovi,c.ValAbsComp from cajabanc c with(nolock)"
    SQLX$ = SQLX$ + " inner join movisto m with(nolock) on c.codicom_lar = m.doseclar"
    SQLX$ = SQLX$ + " WHERE (c.TipoMovim = 'FVEN' OR c.TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + " AND c.NuCLIEN = " & NCLI
    SQLX$ = SQLX$ + " AND c.FechEmisi > '" & FECHALU & " '"  '\\\OT-06-0212-WAR-06/06/06///
    SQLX$ = SQLX$ + "Order by c.FechEmisi ASC "
'=============================================================================
    CAMPOS$ = "FACTURA/A0;Descripción/A16;Referencia/A32;Importe/F12.2"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO, , , ANTOT)
    FRMZELECHO.Width = 10000
    FRMZELECHO.Caption = "Anulación de Cuotas/Matrícula"
    '
    REG& = 0: j& = 0
    FRMZELECHO.MSF1.Rows = 1
    REG& = 0
    Set RS1 = READSET(SQLX$)
    With RS1
    Do While Not .EOF
    j& = j& + 1
       REG& = REG& + 1
       FRMZELECHO.msf2.Rows = REG& + 1
       NCLI1 = !NuCLIEN
       IMPO# = !ValAbsComp
       MPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
       If TRIM$(IMPOS$) = "" Then IMPOS$ = AJUSTD$("0.00", 14)
        NPX& = XVALO(Mid$(!codicom_lar, 11, 8))
        If Left$(!codicom_lar, 2) = "CC" Then
          NPX& = XVALO(Mid$(!codicom_lar, 4))
        End If
       FRMZELECHO.msf2.TextMatrix(REG&, 1) = SAFETEXT$(!codicom_lar)
       FRMZELECHO.msf2.TextMatrix(REG&, 2) = SAFETEXT$(!obsergen)
       FRMZELECHO.msf2.TextMatrix(REG&, 3) = SAFETEXT$(!idenmovi)
       FRMZELECHO.msf2.TextMatrix(REG&, 4) = MPOS$
      .MoveNext
    Loop
    End With
    RS1.Close
    Set RS1 = Nothing
    '
    If j& < 1 Then
       Call MENSERR(24, "La Cuenta Elegida no tiene Cuotas \ que Puedan Anularse.")
       GoTo 10
    End If
    '
    DOCULA$ = ""
    FRMZELECHO.Show 1
    DOCULA$ = TRIM$(RESP_ZELE_VECT(1))
    '
    If DOCULA$ = "" Then GoTo 10
    DOCULA$ = RESP_ZELE_VECT(1)
    NUOPANU& = XVALO(Mid$(DOCULA$, 13, 6))
    IMPDOCU$ = XVALO(RESP_ZELE_VECT(4))
    If IMPDOCU$ = "" Then IMPDOCU$ = " 0.00"

    If COMALTER%("¿Realmente Quiere ANULAR el Documento: " & DOCULA$ & " ,\con Importe: $" & IMPDOCU$ & "  de la Cuenta " & rasocli(NCLI) & " ?\\No, Cancelar\Si, Continuar") <> 2 Then GoTo 10
    '
    'EN EL CASO QUE SE QUIERA ELIMINAR EL REGISTRO VALIDA SI TIENE CAE
    CONDI$ = " (TipoMovim = 'FVEN' OR TipoMovim = 'AJUD') "
    CONDI$ = CONDI$ + " AND CodiCom_Lar = '" & DOCULA$ & "' "
    CONDI$ = CONDI$ + " AND NuCLIEN = " & NCLI & " "
    CONDI$ = CONDI$ + " AND FechEmisi > '" & FECHALU & "' "
    SUCAE$ = TRIM$(VALOBADA("CAJABANC", "CAE", CONDI$, "NOMESS"))
    If SUCAE$ <> "" Then
       TXX1$ = "Imposible Anular el Comprobante: " & DOCULA$ & " con Cae: " & SUCAE$
       TXX1$ = TXX1$ + ".\Este Documento Electrónico Ya Se Encuentra Aprobado y Registrado en la AFIP."
       Call COMUNI(TXX1$)
       Screen.MousePointer = 1
       GoTo 15
    End If
    '
    OKX% = (-1)
    '
    Call APERBASDAT("CABAN")
    Screen.MousePointer = 11
    '
    SQLX$ = "Select * from CajaBanc where "
    SQLX$ = SQLX$ + CONDI$
    Dim MIRS As ADODB.Recordset
    Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With MIRS
    If (.EOF And .BOF) Then
        Call MENSERR(24, "Anulación Imposible. Consulte !")
        GoTo 10
    End If
    '
    REUTI% = 0
    If !OpciMovim = "CI" Then REUTI% = 1
    If REUTI% = 0 Then
       If DERACCE%("NULDEFAC/NOMESS", "Anulacion Reutilizable Facturas") = 2 Then
          RESERNUM% = COMALTER%("Desea Dejar Reservado\este Número de Comprobante ?\\Reservar Número\Re-Utilizarlo")
          If RESERNUM% < 1 Or RESERNUM% > 2 Then
             Call COMUNI("Transaccion Cancelada")
             GoTo 10
          End If
          REUTI% = RESERNUM% - 1
       End If
    End If
    '
    If REUTI% = 0 Then
         Call COMUNI("La Transacción RESERVARÁ el Número de Comprobante.")
       Else
         Call COMUNI("La Transacción PERMITIRÁ REUTILIZAR\el Número de Comprobante Anulado.")
    End If
    End With
    MIRS.Close
    Set MIRS = Nothing
    '
    'TOMO EL CLIREAL PARA USAR LUEGO EN EL MOVISTO
    CONDI$ = "CodiCom_Lar= '" & DOCULA$ & "' And NuClien= " & NCLI & " And FechEmisi > '" & FECHALU & "'"
    CLIENRE = XVALO(VALOBADA("CAJABANC", "CLI_REAL", CONDI$, "NOMESS"))
    NC1 = NCLI
    If CLIENRE > 0 Then NC1 = CLIENRE
    '
    If REUTI% = 1 Then
         strsql = "Delete CajaBanc where "
         strsql = strsql & " CodiCom_Lar= '" & DOCULA$ & "' "
         strsql = strsql & " And NuClien= " & NCLI & " And FechEmisi > '" & FECHALU & "' "
         Call EJECUTACOMANDO(strsql)
       Else
         strsql = "Update CajaBanc set  ValAbsComp = 0,IDEBECOMP = 0,IHABECOMP = 0 "
         strsql = strsql & " where CodiCom_Lar= '" & DOCULA$ & "' "
         strsql = strsql & " And NuClien= " & NCLI & " And FechEmisi > '" & FECHALU & "' "
         Call EJECUTACOMANDO(strsql)
    End If
    '
    'ACTUALIZADO CON BORRAREGISTRO
    CONDI$ = "CodiCom_Lar = '" & DOCULA$ & "' AND NuCLIEN = " & NCLI
    Call BORRAREGISTROS("SADEUPEN", CONDI$, REG&, "NOMESS")
    '
    CONDI$ = "(DoSecLar ='" & DOCULA$ & "' OR  DoPriLar = '" & DOCULA$ & "') AND  Nume_Clie = " & NC1
    Call BORRAREGISTROS("MOVISTO", CONDI$, REG&, "NOMESS")
     Call ACTUREGISTRO("MOVIREPA", "DOSECCOR", CONDI$, " ", REG&, "NOMESS")
     Call ACTUREGISTRO("MOVIREPA", "DOSECLAR", CONDI$, " ", REG&, "NOMESS")
    'ACTUALIZADO CON BORRAREGISTRO
    CONDI$ = "(CoDocOr_Lar = '" & DOCULA$ & "' OR CoDocAp_Lar = '" & DOCULA$ & "') AND NuCLIEN = " & NCLI
    Call BORRAREGISTROS("APLICOBRA", CONDI$, REG&, "NOMESS")
    '
    CAMPOS$ = "FACTURA/A0;Descripción/A16;Referencia/A32;Importe/F12.2"
    IMPDOCU$ = XVALO(RESP_ZELE_VECT(4))
    If IMPDOCU$ = "" Then IMPDOCU$ = " 0.00"
    LEGAJO& = NCLI
    Call ANOTAX_ALUMNO(LEGAJO&, "SE Anuló la Factura: " & DOCULA$ & " (" & RESP_ZELE_VECT(2) & " - " & RESP_ZELE_VECT(3) & ") \  POR IMPORTE: " & IMPDOCU$)

    Screen.MousePointer = 1
    
    Call COMUNI("Anulación Comprobante " + TRIM$(DOCULA$) + " Completa")
    '
999
    Screen.MousePointer = 1
    '
    GoTo 10
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

Private Sub MSF3_Click()
  If (MSF3.MouseRow = 0) Or MSF3.Rows <= 1 Then Exit Sub
  REG& = MSF3.MouseRow
  If REG& < 1 Or MSF3.Rows <= 1 Then Exit Sub
  Call cmdEliminaItem(MSF3, REG&)
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
  'PASA LA FILA SELECCIONADA A LA GRILLA DE ABAJO
  If (MSF.MouseRow = 0) Or MSF.Rows <= 1 Then Exit Sub
  REG& = MSF.MouseRow
  If REG& < 1 Or MSF.Rows <= 1 Then Exit Sub
  
  'PRIMERO SELECCIONO SI EL ID YA NO ESTA PARA IDENTIFICAR QUE NO SE VULEVA A BAJAR
  ID$ = TRIM$(MSF.TextMatrix(REG&, 18))
  NUMREPE& = VALIDA_COD_GRID%(ID$, MSF3, 18, "Este Curso yá Se Ha Seleccionado")
  If NUMREPE& > 0 Then
     Call ResaltarFila(MSF3, NUMREPE&)
     Exit Sub
  End If
  '
  'PASO LOS DATOS A LA GRILLA INFERIOR
  MSF3.Rows = MSF3.Rows + 1: MSF3.Cols = MSF3.Cols
  For i& = 1 To MSF.Cols - 1
     MSF3.TextMatrix(MSF3.Rows - 1, i&) = MSF.TextMatrix(REG&, i&)
  Next i&
  '
  'PINTO Y POSICIONO A LA FILA AGREGADA
  Call ResaltarFila(MSF3, MSF3.Rows - 1)

  MSF3.TopRow = MSF3.Rows - 1
  MSF3.Row = MSF3.Rows - 1
  MSF3.ColSel = 1
'  MSF3.SetFocus
  '
  'ELIMINO DE LA GRILLA SUPERIOR EL ITEM QUE SE BAJO
  Call cmdEliminaItem(MSF, REG&)
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
            
            
'            Me.TXTVALORTOTAL = ""
'            Me.TXTMATRICULA = ""
            
            
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
'       !REFERENCIA = Left$(Me.TXTREFERENCIA, 256)
'       !Observa = Me.TXTOBSERVACIONES
       !ID_CURSO = XVALO(Me.Text1)
'       LOSDIAS$ = DIA1$: If TRIM$(DIA2$) <> "" Then LOSDIAS$ = LOSDIAS$ & " Y " & DIA2$
       .Update
       End With
       rs.Close
       Set rs = Nothing
       IDCURS& = XVALO(Text1)
       Text1 = ""
       Text1 = IDCURS&
       

End Sub

Sub MOSTRAR_DATOS_FILA(MSF As MSFlexGrid, REG&)
'       Me.TXTINICIO = TRIM$(MSF.TextMatrix(REG&, 1))
'       Me.TXTFINALIZA = TRIM$(MSF.TextMatrix(REG&, 2))
'       DIA1$ = TRIM$(MSF.TextMatrix(REG&, 4))
'       DIA2$ = TRIM$(MSF.TextMatrix(REG&, 6))
'       LOSDIAS$ = DIA1$: If TRIM$(DIA2$) <> "" Then LOSDIAS$ = LOSDIAS$ & " Y " & DIA2$
'       Me.TXTSEDICTA = TRIM$(LOSDIAS$)
'       Me.TXTHDESDE = TRIM$(MSF.TextMatrix(REG&, 7))
'       Me.TXTHHASTA = TRIM$(MSF.TextMatrix(REG&, 8))
'       Me.TXTNUMPROFESOR = TRIM$(MSF.TextMatrix(REG&, 9))
'       Me.TXTPROFESOR = TRIM$(MSF.TextMatrix(REG&, 10))
'       Me.TXTNUMAULA = TRIM$(MSF.TextMatrix(REG&, 11))
'       Me.TXTAULA = TRIM$(MSF.TextMatrix(REG&, 12))
'
''       Me.TXTVALORTOTAL = TRIM$(FORMATNUM$(MSF.TextMatrix(REG&, 13), "f12.2"))
''       Me.TXTMATRICULA = TRIM$(FORMATNUM$(MSF.TextMatrix(REG&, 14), "f12.2"))
'       Me.TXTCANTICUOTAS = TRIM$(MSF.TextMatrix(REG&, 15))
'       Me.TXTCAPACIDAD = TRIM$(MSF.TextMatrix(REG&, 16))
'       Me.TXTCODINT = MSF.TextMatrix(REG&, 18)
'       Me.TXTCANTIHORAS = H_TRABAJADAS = XVALO(DIFERENCIAHORASCENTESIMAL(HORANUM(TXTHDESDE), HORANUM(TXTHHASTA)))
'       Me.TXTOBSERVACIONES = VALOBADA("MASTER", "OBSERVA", "CODINT = " & XVALO(TXTCODINT), "NOMESS")
'       Me.TXTREFERENCIA = VALOBADA("MASTER", "REFERENCIA", "CODINT = " & XVALO(TXTCODINT), "NOMESS")
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



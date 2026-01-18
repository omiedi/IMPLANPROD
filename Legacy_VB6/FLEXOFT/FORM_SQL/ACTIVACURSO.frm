VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ACTIVACURSO 
   Caption         =   "Alta de Nuevos Horarios"
   ClientHeight    =   7530
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   15705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   15705
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
      Left            =   10080
      TabIndex        =   49
      ToolTipText     =   "Conmuta Entre Ver todos y Solo los que la fecha de inicio no supera los 28 días "
      Top             =   840
      Width           =   1575
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
      Height          =   8040
      Left            =   120
      TabIndex        =   1
      Top             =   -120
      Width           =   15555
      Begin VB.TextBox TXTRESPONSABLENUM 
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
         Left            =   12680
         TabIndex        =   68
         Top             =   6840
         Width           =   360
      End
      Begin VB.CommandButton Command2 
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
         Left            =   13050
         TabIndex        =   67
         Top             =   6840
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
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   66
         TabStop         =   0   'False
         Text            =   "Responsable:"
         Top             =   6840
         Width           =   1125
      End
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
         Left            =   13320
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   6840
         Width           =   2040
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
         TabIndex        =   61
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
         TabIndex        =   60
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
         TabIndex        =   59
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
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   56
         TabStop         =   0   'False
         Text            =   "Valor Cuota:"
         Top             =   240
         Width           =   1965
      End
      Begin VB.Frame Frame2 
         Height          =   600
         Left            =   360
         TabIndex        =   50
         Top             =   3960
         Width           =   7455
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "ACTIVACURSO.frx":0000
            TabIndex        =   53
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
            Left            =   4200
            Style           =   2  'Dropdown List
            TabIndex        =   52
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
            TabIndex        =   51
            Top             =   160
            Width           =   2535
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   3480
            OleObjectBlob   =   "ACTIVACURSO.frx":006E
            TabIndex        =   54
            Top             =   240
            Width           =   615
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
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   7560
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
         Left            =   7560
         TabIndex        =   47
         Top             =   4680
         Width           =   7800
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
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "Referencia Curso:"
         Top             =   4680
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   4680
         Width           =   3915
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
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Text            =   "Código Interno:"
         Top             =   4680
         Width           =   1515
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Grabar Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11520
         TabIndex        =   43
         Top             =   7200
         Width           =   3830
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   7200
         Width           =   315
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   6840
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
         Index           =   5
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "Area:"
         Top             =   600
         Width           =   1935
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
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   "Inicio:"
         Top             =   5040
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
         Index           =   6
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   36
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
         Left            =   5025
         TabIndex        =   35
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   34
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
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   33
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
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   32
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   31
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   30
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
         Index           =   1
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "Finaliza:"
         Top             =   5400
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
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "Se dicta:"
         Top             =   5760
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
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "H.Desde:"
         Top             =   6120
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
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "H.Hasta:"
         Top             =   6480
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
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "Profesor:"
         Top             =   6840
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
         Index           =   9
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "Aula:"
         Top             =   7200
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
         Index           =   10
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "Valor Cuota:"
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
         Index           =   11
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Cant. de Cuotas:"
         Top             =   5400
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
         Index           =   12
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Matrícula:"
         Top             =   5760
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
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Capacidad:"
         Top             =   6120
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
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Anotados:"
         Top             =   6480
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
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "Vacantes:"
         Top             =   6840
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
         Left            =   6000
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "Cant.Horas x Día:"
         Top             =   7200
         Visible         =   0   'False
         Width           =   1515
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   5040
         Width           =   3915
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   5400
         Width           =   3915
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   5760
         Width           =   3915
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   6120
         Width           =   3915
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   6480
         Width           =   3915
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   6840
         Width           =   3555
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   7200
         Width           =   3555
      End
      Begin VB.TextBox TXTVALORTOTAL 
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
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   5040
         Width           =   3915
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
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   5400
         Width           =   3915
      End
      Begin VB.TextBox TXTMATRICULA 
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
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   5760
         Width           =   3915
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
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   6120
         Width           =   3915
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
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   6480
         Width           =   3915
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
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   6840
         Width           =   3915
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
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   7200
         Visible         =   0   'False
         Width           =   3915
      End
      Begin VB.TextBox TXTOBSERVACIONES 
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
         Height          =   1725
         Left            =   11520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   5040
         Width           =   3885
      End
      Begin ACTIVESKINLibCtl.SkinLabel SKALTA 
         Height          =   255
         Left            =   5880
         OleObjectBlob   =   "ACTIVACURSO.frx":00DA
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   1995
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   2325
         Left            =   360
         TabIndex        =   40
         ToolTipText     =   "Doble Click Para Editar"
         Top             =   1560
         Width           =   15045
         _ExtentX        =   26538
         _ExtentY        =   4101
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
         Height          =   2325
         Left            =   360
         TabIndex        =   55
         ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
         Top             =   1560
         Width           =   15045
         _ExtentX        =   26538
         _ExtentY        =   4101
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
      Begin VB.Label Label1 
         Caption         =   "Cursos Existentes"
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
         Left            =   360
         TabIndex        =   64
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "'Ver Todos': sin filtro por fecha."
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
         Left            =   7920
         TabIndex        =   63
         Top             =   4320
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "'Vista Filtrada': no superan el mes de inicio"
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
         Left            =   7920
         TabIndex        =   62
         Top             =   4080
         Width           =   3735
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
      OleObjectBlob   =   "ACTIVACURSO.frx":0154
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
      Caption         =   "ABM"
      Begin VB.Menu MNUALTA 
         Caption         =   "ALTA"
         Shortcut        =   ^A
      End
      Begin VB.Menu MNUBAJA 
         Caption         =   "BAJA"
      End
      Begin VB.Menu MNUREST 
         Caption         =   "RESTAURAR BAJA"
      End
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
      Begin VB.Menu mnuQuitar 
         Caption         =   "Dar de Baja Curso Seleccionado"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "ACTIVACURSO"
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




Private Sub Combo2_Change()
   TXCOMBO2 = Combo2.Text
End Sub

Private Sub Combo2_Click()
   TXCOMBO2 = Combo2.Text
End Sub

Private Sub Combo3_Change()
   TXCOMBO3 = Combo3.Text
End Sub

Private Sub Combo3_Click()
   TXCOMBO3 = Combo3.Text
End Sub

Private Sub Combo5_Change()
  TXCOMBO5 = Combo5.Text
End Sub

Private Sub Combo5_Click()
  TXCOMBO5 = Combo5.Text
End Sub


Private Sub Combo6_Change()
   TXCOMBO6 = Combo6.Text
End Sub

Private Sub Combo6_Click()
   TXCOMBO6 = Combo6.Text
End Sub

Private Sub Combo7_Change()
   TXCOMBO7 = Combo7.Text
End Sub

Private Sub Combo7_Click()
   TXCOMBO7 = Combo7.Text
End Sub

Private Sub Combo8_Change()
   TXCOMBO8 = Combo8.Text
End Sub

Private Sub Combo8_Click()
   TXCOMBO8 = Combo8.Text
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
    
'    Campox$ = "F.Inicio/D8;F.Fin/D8;Días/A12;Año/F5;H.Desde/H12;H.Hasta/H12;Id Prof/F8;Profesor/A32"
'    Call CARGA_ENCABEZADO(ACTIVACURSO.MSF, Campox$, ACTIVACURSO)
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


Sub EVENTOSELCHANGE(MSF As MSFlexGrid, REG1&)
   'REG1& = MSF.MouseRow
   Call MOSTRAR_DATOS_FILA(MSF, REG1&)
   TXTMASID = XVALO(MSF.TextMatrix(REG1&, 20))
End Sub

Sub MENUXELDA(MSF As MSFlexGrid)
     If DERACCE("MODCURSO", "MODIFICAR DATOS DE CURSO") < 2 Then Exit Sub
     GLOBALFRM_INDIX% = Index
     REG& = MSF.MouseRow
     COL% = MSF.MouseCol
     Call MENU_CELDA(MSF, REG&, COL%)

End Sub

Private Sub Check1_Click()
    If XVALO(Text1) < 1 Then
       Call COMUNI("Debe Seleccionar Primero un Curso Padre.")
       Check1.Value = 0
       Exit Sub
    End If
    '
    If Check1.Caption = "Ver Todos" Then
       Check1.Caption = "Vista Filtrada"
    Else
       Check1.Caption = "Ver Todos"
    End If
'    MAS28% = 1
'    If Check1.Value = 1 Then
'      MAS28% = 0
'    End If
    Dim obj As New RECORXET
    Set RS1 = obj.RS_CURSOS_ACT(XVALO(Text1), 0, Check1.Value)
    '
    MSF.Rows = 1
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

Private Sub Command1_Click()
    Command1.Enabled = False
    '
    If SKALTA.Visible = True Then
      TX$ = "PARA SELECCIONAR CURSO EXISTENTE" + vbCrLf
      TX$ = TX$ + "DEBE VOLVER A MODO EDICIÓN (Presionando el Botón 'Edición')"
      MsgBox TX$
      GoTo 99
    End If
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


Private Sub Command2_Click()
    Command2.Enabled = False
    CAMPOS$ = "Código/f6;Responsable/A48"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
    FRMZELECHO.Caption = "TABLA DE RESPONSABLES"
    
    Dim obj As New RECORXET
    Set RS1 = obj.RS_RESPONSABLES_CME()
    Call Carga_MSF_Recordset(RS1, CAMPOS$, FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 9000
50  FRMZELECHO.Show 1

    Me.TXTRESPONSABLENUM = TRIM(RESP_ZELE_VECT(1))


   
   
   
    Command2.Enabled = True

End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    If XVALO(Me.TXTMASID) > 0 Then
       SQLX$ = "UPDATE MASTER SET REFERENCIA = '" & TXTREFERENCIA & "' , OBSERVA = ' " & Me.TXTOBSERVACIONES & "'"
       SQLX$ = SQLX$ + " , COD_RESPONSABLE = " & XVALO(Me.TXTRESPONSABLENUM) & ", DESC_RESPONSABLE = '" & TRIM$(Me.TXTRESPONSABLETEXT) & "'"
       SQLX$ = SQLX$ + " WHERE CODINT = " & XVALO(Me.TXTMASID)
       FLEXCONN.Execute (SQLX$)
    End If
    
99  Command3.Enabled = True
    

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
    Campox$ = "F.Inicio/D8;F.Fin/D8;C/A0;Días/A12;C/A0;Días/A12;C/A0;Días/A12;H.Desde/H12;H.Hasta/H12;Id Prof/A0;Profesor/A32;C.A./A0;Aula/A18;Val.Curso/f10.2;Val.Matrícula/F11.2;Cant.Meses/F10;Capacidad/F10;Ord/A0;CODINT/F7"
    
    

    Call CARGA_ENCABEZADO(ACTIVACURSO.MSF, Campox$, ACTIVACURSO, 1)
    Call CARGA_ENCABEZADO(ACTIVACURSO.MSF_2, Campox$, ACTIVACURSO, 0)
    MSF.Rows = 2
    'PROPIEDADES DE SELECCION DE FILA
    MSF.SelectionMode = flexSelectionByRow
    MSF.FocusRect = flexFocusNone
    '

End Sub



Private Sub MSF_2_Click()
   REG& = MSF_2.MouseRow
   If REG& < 1 Then Exit Sub
   Call EVENTOSELCHANGE(ACTIVACURSO.MSF_2, REG&)

End Sub

Private Sub MSF_2_DblClick()
   REG& = MSF_2.MouseRow
   If REG& < 1 Then Exit Sub
   Call T_Espera(1 / 3)
   MASID& = XVALO(MSF_2.TextMatrix(REG&, 18)) 'EL ID DE LA TABLA
   
   TXTBUSCAR = ""
   MSF_2.Visible = False
   MSF.Visible = True
   For i& = 1 To MSF.Rows - 1
        IDGRILLA& = XVALO((MSF.TextMatrix(i&, 18)))
        If IDGRILLA& = MASID& Then
             Call ResaltarFila(MSF, i&)
             Call T_Espera(1 / 3)
             Call EVENTOSELCHANGE(MSF, i&) ' PARA QUE MUESTRE EN LA PARTE INFERIOR  LOS DATOS
             Exit For
        End If
   Next i&
   
End Sub


Private Sub MSF_2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
  End If

End Sub


Private Sub MSF_2_RowColChange()
   Call EVENTOSELCHANGE(ACTIVACURSO.MSF_2, MSF_2.Row)
End Sub

Private Sub MSF_Click()
   REG& = MSF.MouseRow
   If REG& < 1 Then Exit Sub
   Call EVENTOSELCHANGE(ACTIVACURSO.MSF, REG&)
End Sub

Private Sub MSF_RowColChange()
      'Call EVENTOSELCHANGE(ACTIVACURSO.MSF, MSF.Row)
End Sub

Private Sub MSF_SelChange()
   Call EVENTOSELCHANGE(ACTIVACURSO.MSF, MSF.Row)
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode 'SI TIENE UN TEXTO DE BUSQUEDA BORRA
'        Case 38 '(flecha arriba,abajo,39derecha,40izquierda)
'          If MSF.Visible = True Then
'            On Error Resume Next
'            MSF.SetFocus
'            If MSF.Row > 1 Then MSF.Row = MSF.Row - 1 'CUANDO TOMA EL FOCO DESDE EL TEXTBUSCAR, PARA QUE SUBA UNA POSICION
'            Call EVENTOSELCHANGE(MSF, MSF.Row)
'          End If
'        Case 40 '(flecha arriba,abajo,39derecha,40izquierda)
'          If MSF.Visible = True Then
'            On Error Resume Next
'            MSF.SetFocus
'            If MSF.Row < MSF.Rows Then MSF.Row = MSF.Row + 1 'CUANDO TOMA EL FOCO DESDE EL TEXTBUSCAR, PARA QUE BAJE UNA POSICION
'            Call EVENTOSELCHANGE(MSF, MSF.Row)
'
'          End If
'    End Select
End Sub

Private Sub Text3_Change()
   
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
  Call DOBLE_CLICK_EN_GRILLA(ACTIVACURSO.MSF)
End Sub

Sub DOBLE_CLICK_EN_GRILLA(MSF As MSFlexGrid)
    Unload VENTABNEW
    REG& = MSF.MouseRow
    If REG& < 1 Then Exit Sub
    If FilaVacia%(Me.MSF, 0, 0) > 0 Then Exit Sub
    'CONFIGURACION DE LA VENTANA
    VENTANA.ColEcoCampo(3) = "CAMPO1;CAMPO2;DATASQL;LISTA  = 'DIAS_SEMANA'"
    VENTANA.ColEcoCampo(5) = "CAMPO1;CAMPO2;DATASQL;LISTA  = 'DIAS_SEMANA'"
    VENTANA.ColEcoCampo(7) = "CAMPO1;CAMPO2;DATASQL;LISTA  = 'DIAS_SEMANA'"
    
    VENTANA.ColEcoCampo(13) = "CAMPO1;CAMPO2;DATASQL;LISTA  = 'AULAS'"
    VENTANA.ColEcoCampo(11) = "NUME_CLI;RASO_CLI;PROFESORES"
    VENTANA.CampEditables = "1;2;3;5;7;9;10;11;13;17;18"
    VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 0
    VENTANA.HabilitaInsertar = 1
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "EDICIÓN DE LA FILA SELECCIONADA"
    VENTANA.Inicializar = 1
    
    'Campox1$ = "F.Inicio/D8;F.Fin/D8;C/F1;Días/A12;C/F1;Días/A12;H.Desde/H12;H.Hasta/H12;Id Prof/F8;Profesor/A32;C.A./f4;Aula/A18;Val.Curso/f10;Val.Matrícula/F10;Cant.Meses/F10;Capacidad/F10;Ord/F0;CODINT/F7"
    Campox1$ = "F.Inicio/D8;F.Fin/D8;C/F1;Día/A12;C/F1;Día/A12;C/F1;Día/A12;H.Desde/H12;H.Hasta/H12;Id Prof/F8;Profesor/A32;C.A./f4;Aula/A18;Val.Curso/f0;Val.Matrícula/F0;Cant.Meses/F10;Capacidad/F10;Ord/F0;CODINT/F7"

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
        Call UPDATE_CURSO_MATER(REG&, XVALO(MSF.TextMatrix(REG&, 19)))
       Else
         GoTo 30
       End If
    End If
'    Call ResaltarFila(MSF, REG&)
End Sub


Function VALIDAR_CURSO_OK%(MSF1 As MSFlexGrid)

       VALHOR& = VALIDARHORAS&(MSF1, 9, 10)
       If VALHOR& > 0 Then
          Call COMUNI("Rango Horario Inválido en Fila: " & VALHOR&)
          GoTo 30
       End If
       '
       If FECHANUM(TRIM$(MSF1.TextMatrix(1, 1))) < 1 Then
          Call COMUNI("Fecha Desde Inválida")
          GoTo 30
       End If
       If FECHANUM(TRIM$(MSF1.TextMatrix(1, 2))) < 1 Then
          Call COMUNI("Fecha Hasta Inválida")
          GoTo 30
       End If
       If FECHANUM(MSF1.TextMatrix(1, 1)) > FECHANUM(MSF1.TextMatrix(1, 2)) Then
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
       If XVALO(MSF1.TextMatrix(1, 7)) > 0 Then
        If VALIDAR_ECO(MSF1, 8, 0) > 0 Then
           Call COMUNI("Día 3 Inválido")
           GoTo 30
        End If
       End If
       If XVALO(MSF1.TextMatrix(1, 11)) > 0 Then
        If VALIDAR_ECO(MSF1, 12, 0) > 0 Then
           Call COMUNI("Profesor Inválido")
           GoTo 30
        End If
       Else
           Call COMUNI("Falta Ingresar Profesor")
           GoTo 30
       End If
       If XVALO(MSF1.TextMatrix(1, 13)) > 0 Then
        If VALIDAR_ECO(MSF1, 14, 0) > 0 Then
           Call COMUNI("Aula Inválida")
           GoTo 30
        End If
       Else
           Call COMUNI("Falta Ingresar el Aula")
           GoTo 30
       End If
       If XVALO(Me.TXTCURSOTOT) < 1 Then
           Call COMUNI("Falta Ingresar el Costo del Curso")
           GoTo 30
       End If
       If XVALO(Me.TXTMATRICULA2) < 1 Then
           Call COMUNI("Falta Ingresar el Costo de la Matrícula")
           GoTo 30
       End If
       If XVALO(MSF1.TextMatrix(1, 17)) < 1 Then
           Call COMUNI("Falta Ingresar La Cantidad de Meses")
           GoTo 30
       End If
       If XVALO(MSF1.TextMatrix(1, 18)) < 1 Then
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
       If DERACCE%("BORRCURSO", "BORRAR CURSO SELECCIONADO") < 2 Then Exit Sub
       FILX& = GLOBALFRM_FILA&
       OPX% = COMALTER%("Confirma dar de Baja el Curso Seleccionado\\Cancelar\Dar de Baja")
       MAS_ID& = XVALO(MSF.TextMatrix(GLOBALFRM_FILA&, 18))
       SQLX$ = "UPDATE MASTER SET STATUS = -1 WHERE MASTER_ID = " & MAS_ID&
       FLEXCONN.Execute (SQLX$)
       Call cmdEliminaItem(MSF, GLOBALFRM_FILA&)
       If MSF.Rows < 2 Then MSF.Rows = 2
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
            
            Me.TXTVALORTOTAL = ""
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

Private Sub MNUALTA_Click()
    If XVALO(Text1) < 1 Then
       Call COMUNI("Debe Seleccionar Primero un Curso Padre Para Dar de Alta uno Dependiente.")
       Exit Sub
    End If
    If XVALO(Me.TXTCURSOTOT) < 1 Then
       Call COMUNI("Falta Ingresar Valor de la Cuota Para el Curso Activo")
       Exit Sub
    End If
    If XVALO(Me.TXTMATRICULA2) < 1 Then
       Call COMUNI("Falta Ingresar Valor de la Matrícula Para el Curso Activo")
       Exit Sub
    End If
    '
    VENTANA.Inicializar = 0
    VENTANA.ColEcoCampo(3) = "CAMPO1;CAMPO2;DATASQL;LISTA  = 'DIAS_SEMANA'"
    VENTANA.ColEcoCampo(5) = "CAMPO1;CAMPO2;DATASQL;LISTA  = 'DIAS_SEMANA'"
    VENTANA.ColEcoCampo(7) = "CAMPO1;CAMPO2;DATASQL;LISTA  = 'DIAS_SEMANA'"
    
    VENTANA.ColEcoCampo(13) = "CAMPO1;CAMPO2;DATASQL;LISTA  = 'AULAS'"
    VENTANA.ColEcoCampo(11) = "NUME_CLI;RASO_CLI;PROFESORES;STAT_CLI >=0"
    VENTANA.CampEditables = "1;2;3;5;7;9;10;11;13;17;18"
    VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 0
    VENTANA.HabilitaInsertar = 1
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = ""
    VENTANA.Inicializar = 1
    
    Campox1$ = "F.Inicio/D8;F.Fin/D8;C/F1;Día/A12;C/F1;Día/A12;C/F1;Día/A12;H.Desde/H12;H.Hasta/H12;Id Prof/F8;Profesor/A32;C.A./f4;Aula/A18;Val.Curso/f0;Val.Matrícula/F0;Cant.Meses/F10;Capacidad/F10;Ord/F0;CODINT/F7"
    VENTANA.Campox = Campox1$
    
    Call CopiarMsfAOtroMsf_Solo_Encabezado(MSF, VENTABNEW.MSF, Campox1$, VENTABNEW)
    '****************************************************************
    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
    
50  VENTABNEW.Show 1
    A = FilaVacia%(MSF, 1)
    If VENTABNEW.APROBADO% > 0 Then
      If VALIDAR_CURSO_OK%(VENTABNEW.MSF) > 0 Then
        For i& = 1 To VENTABNEW.MSF.Rows - 1
          U& = MSF.Rows
          MSF.Rows = U& + 1
          For j& = 1 To VENTABNEW.MSF.Cols - 1
              DATO$ = TRIM$(VENTABNEW.MSF.TextMatrix(i&, j&))
              MSF.TextMatrix(U&, j&) = DATO$
          Next j&
          Call UPDATE_CURSO_MATER(U&)
        Next i&
        Else
          GoTo 50
        End If
    End If
    
    Set VENTANA = Nothing

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
       !H_DESDE_FN = HORANUM(MSF.TextMatrix(REG&, 7), "NOMESS") 'guardo en formato numero para ordenamiento
       !H_HASTA = MSF.TextMatrix(REG&, 10)
       H_TRABAJADAS = XVALO(DIFERENCIAHORASCENTESIMAL(HORANUM(MSF.TextMatrix(REG&, 9)), HORANUM(MSF.TextMatrix(REG&, 10))))
       !CantHorCso = H_TRABAJADAS
       !ID_PROFESOR = MSF.TextMatrix(REG&, 11)
       !ID_AULA = MSF.TextMatrix(REG&, 13)
       !COSUNI = ROUND(XVALO(Me.TXTCURSOTOT), 2)
       !ValMatricula = ROUND(XVALO(Me.TXTMATRICULA2), 2)
       !CantMeses = MSF.TextMatrix(REG&, 17)
       !CAPACIDAD = MSF.TextMatrix(REG&, 18)
       
       !Ferevis = CVDAT(HOY(HO$))
       !COD_AREA = XVALO(Me.TXTCODAREA)
       !DESCR_AREA = Left$(Me.TXTDESCRIAREA, 128)
       !DESCRICURSO = Left$(Me.TXTDESCRICURSO, 256)
       !referencia = Left$(Me.TXTREFERENCIA, 256)
       !Observa = Me.TXTOBSERVACIONES
       !ID_CURSO = XVALO(Me.Text1)
       !Status = 0
       !COD_RESPONSABLE = XVALO(Me.TXTRESPONSABLENUM)
       !DESC_RESPONSABLE = Me.TXTRESPONSABLETEXT

'       LOSDIAS$ = DIA1$: If TRIM$(DIA2$) <> "" Then LOSDIAS$ = LOSDIAS$ & " Y " & DIA2$
       .Update
       End With
       rs.Close
       Set rs = Nothing
       IDCURS& = XVALO(Text1)
       Text1 = ""
       Text1 = IDCURS&
       

End Sub

Private Sub MNUREST_Click()

    CAMPOS$ = "Area/A20;Título Curos/A32;F.Inicio/D8;F.Fin/D8;C/A0;Días/A12;C/A0;Días/A12;C/A0;Días/A12;H.Desde/H12;H.Hasta/H12;Id Prof/F8;Profesor/A32;C.A./f4;Aula/A18;Val.Curso/f10.2;Val.Matrícula/F11.2;Cant.Meses/F10;Capacidad/F10;Ord/F0;CODINT/F7"

    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO, , , ANTOT)
    FRMZELECHO.Width = ANTOT + 13000
    FRMZELECHO.Caption = "Curos Dados de Baja"
    '
    REG& = 0
    FRMZELECHO.MSF1.Rows = 1
    Dim obj As New RECORXET
    Set RS1 = obj.RS_CURSOS_ACT(XVALO(Text1), 1)
    REG& = 0
    With RS1
    Do While Not .EOF
    j& = j& + 1
       REG& = REG& + 1
       FRMZELECHO.msf2.Rows = REG& + 1
       
       FRMZELECHO.msf2.TextMatrix(REG&, 1) = SAFETEXT$(!DESCR_AREA)
       FRMZELECHO.msf2.TextMatrix(REG&, 2) = SAFETEXT$(!DESCRICURSO)
       
       FRMZELECHO.msf2.TextMatrix(REG&, 3) = FECHATEX$(CVINT(!feinicurs))
       FRMZELECHO.msf2.TextMatrix(REG&, 4) = FECHATEX$(CVINT(!FeFinCurs))
       FRMZELECHO.msf2.TextMatrix(REG&, 5) = XVALO(!COD_DIA1)
       FRMZELECHO.msf2.TextMatrix(REG&, 6) = SAFETEXT$(!desc_dia1)
       FRMZELECHO.msf2.TextMatrix(REG&, 7) = XVALO(!COD_DIA2)
       FRMZELECHO.msf2.TextMatrix(REG&, 8) = SAFETEXT$(!desc_dia2)
       FRMZELECHO.msf2.TextMatrix(REG&, 9) = XVALO(!COD_DIA3)
       FRMZELECHO.msf2.TextMatrix(REG&, 10) = SAFETEXT$(!desc_dia3)
        
       FRMZELECHO.msf2.TextMatrix(REG&, 11) = SAFETEXT$(!H_DESDE)
       FRMZELECHO.msf2.TextMatrix(REG&, 12) = SAFETEXT$(!H_HASTA)
       FRMZELECHO.msf2.TextMatrix(REG&, 13) = XVALO(!ID_PROFESOR)
       FRMZELECHO.msf2.TextMatrix(REG&, 14) = VALOBADA("PROFESORES", "RASO_CLI", "NUME_CLI = " & XVALO(!ID_PROFESOR))
       FRMZELECHO.msf2.TextMatrix(REG&, 15) = XVALO(!ID_AULA)
       CONDI$ = "LISTA = 'AULAS' And CAMPO1 = '" & FORMATOCON0$(XVALO(!ID_AULA), 4) & "'"
       FRMZELECHO.msf2.TextMatrix(REG&, 16) = VALOBADA("DATASQL", "CAMPO2", CONDI$, "MOMES")
       FRMZELECHO.msf2.TextMatrix(REG&, 17) = FORMATNUM$(XVALO(!COSUNI), "F12.2")
       FRMZELECHO.msf2.TextMatrix(REG&, 18) = FORMATNUM$(XVALO(!ValMatricula), "F12.2")
       FRMZELECHO.msf2.TextMatrix(REG&, 19) = XVALO(!CantMeses)
       FRMZELECHO.msf2.TextMatrix(REG&, 20) = XVALO(!CAPACIDAD)
       FRMZELECHO.msf2.TextMatrix(REG&, 21) = XVALO(!MASTER_ID)
       FRMZELECHO.msf2.TextMatrix(REG&, 22) = XVALO(!CODINT)
       IDCUR& = XVALO(!ID_CURSO)
      .MoveNext
    Loop
    End With
    RS1.Close
    Set RS1 = Nothing
    '
    FRMZELECHO.Show 1

    MASID& = XVALO(RESP_ZELE_VECT(22))
    If MASID& > 0 Then
        SQLX$ = "UPDATE MASTER SET STATUS = 1 WHERE MASTER_ID = " & MASID&
        FLEXCONN.Execute (SQLX$)
        Call LIMPIAR
        Text1 = ""
        Text1 = IDCUR&
    End If
    
    
    
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
       
       Me.TXTVALORTOTAL = TRIM$(FORMATNUM$(MSF.TextMatrix(REG&, 15), "f12.2"))
       Me.TXTMATRICULA = TRIM$(FORMATNUM$(MSF.TextMatrix(REG&, 16), "f12.2"))
       Me.TXTCANTICUOTAS = TRIM$(MSF.TextMatrix(REG&, 17))
       Me.TXTCAPACIDAD = TRIM$(MSF.TextMatrix(REG&, 18))
       Me.TXTCODINT = MSF.TextMatrix(REG&, 20)
       Me.TXTCANTIHORAS = H_TRABAJADAS = XVALO(DIFERENCIAHORASCENTESIMAL(HORANUM(TXTHDESDE), HORANUM(TXTHHASTA)))
       '
       CONDI$ = "CODINT = " & XVALO(TXTCODINT)
       SQLX$ = "SELECT OBSERVA,REFERENCIA,COD_RESPONSABLE FROM MASTER WITH(NOLOCK) WHERE " & CONDI$
       Set rs = READSET(SQLX$)
       Me.TXTOBSERVACIONES = SAFETEXT$(rs!Observa) 'VALOBADA("MASTER", "OBSERVA", "CODINT = " & XVALO(TXTCODINT), "NOMESS")
       Me.TXTREFERENCIA = SAFETEXT$(rs!referencia) 'VALOBADA("MASTER", "REFERENCIA", "CODINT = " & XVALO(TXTCODINT), "NOMESS")
       Me.TXTRESPONSABLENUM = XVALO(rs!COD_RESPONSABLE)
       rs.Close
       Set rs = Nothing
       '
'       VALIDAR EL STATUS
       TXTANOTADOS = CantRegistros("ALUMNO_CURSO", "CURSOS_ID = " & XVALO(MSF.TextMatrix(REG&, 20)) & " AND STATUS >= 0", "NOMESS")
       Me.TXTVACANTES = XVALO(TXTCAPACIDAD) - XVALO(TXTANOTADOS)
       
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

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
         SendKeys "{TAB}"
   End If

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


Private Sub TXTCODINT_Change()

End Sub


Private Sub TXTRESPONSABLENUM_Change()
       CONDI$ = "LISTA = 'RESPONSABLES' And CAMPO1 = '" & FORMATOCON0$(XVALO(Me.TXTRESPONSABLENUM), 5) & "'"
       Me.TXTRESPONSABLETEXT = VALOBADA("DATASQL", "CAMPO2", CONDI$, "MOMES")

End Sub



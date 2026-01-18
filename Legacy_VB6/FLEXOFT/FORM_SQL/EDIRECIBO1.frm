VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form EDIRECIBO1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Edición"
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Generar Cuota / Matrícula"
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
      Left            =   6720
      TabIndex        =   72
      Top             =   7240
      Width           =   3135
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   4850
      TabIndex        =   70
      Top             =   600
      Width           =   870
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
         Height          =   795
         Left            =   40
         Picture         =   "EDIRECIBO1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Actualizar"
         Top             =   120
         Width           =   765
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Guardar "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6240
      TabIndex        =   66
      Top             =   8040
      Visible         =   0   'False
      Width           =   5970
   End
   Begin VB.Frame Frame5 
      Height          =   1485
      Left            =   5520
      TabIndex        =   55
      Top             =   8160
      Width           =   6105
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
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   65
         TabStop         =   0   'False
         Text            =   "COMPUTACIÓN"
         Top             =   555
         Width           =   1400
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
         Index           =   1
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         Text            =   "TOTAL"
         Top             =   915
         Width           =   1400
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
         Index           =   3
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   63
         TabStop         =   0   'False
         Text            =   "INGLÉS"
         Top             =   195
         Width           =   1400
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
         Index           =   0
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   62
         TabStop         =   0   'False
         Text            =   "EMERGENCIAS"
         Top             =   570
         Width           =   1400
      End
      Begin VB.TextBox Text20 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   61
         Top             =   555
         Width           =   1200
      End
      Begin VB.TextBox Text19 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   60
         Top             =   915
         Width           =   1200
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   195
         Width           =   1200
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   58
         Top             =   210
         Width           =   1200
      End
      Begin VB.TextBox TXCOMBO1 
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
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Text            =   "Ens.Programada"
         Top             =   210
         Width           =   1400
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   1635
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   570
         Width           =   1200
      End
   End
   Begin VB.Frame Frame6 
      Height          =   1485
      Left            =   120
      TabIndex        =   42
      Top             =   7920
      Width           =   5160
      Begin VB.TextBox Text29 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCF170&
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
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   835
         Width           =   1850
      End
      Begin VB.TextBox Text28 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCF170&
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
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   485
         Width           =   1850
      End
      Begin VB.TextBox TXTDCTO 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCF170&
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
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   135
         Width           =   1850
      End
      Begin VB.TextBox TXTANO 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCF170&
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
         Left            =   570
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   835
         Width           =   1850
      End
      Begin VB.TextBox TXTMES 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCF170&
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
         Left            =   570
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   485
         Width           =   1850
      End
      Begin VB.TextBox TXTTIPO 
         Appearance      =   0  'Flat
         BackColor       =   &H00FCF170&
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
         Left            =   570
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   135
         Width           =   1850
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   225
         Left            =   105
         OleObjectBlob   =   "EDIRECIBO1.frx":0578
         TabIndex        =   49
         Top             =   240
         Width           =   540
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   225
         Left            =   105
         OleObjectBlob   =   "EDIRECIBO1.frx":05E0
         TabIndex        =   50
         Top             =   590
         Width           =   540
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   225
         Left            =   105
         OleObjectBlob   =   "EDIRECIBO1.frx":0646
         TabIndex        =   51
         Top             =   890
         Width           =   540
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   225
         Left            =   2640
         OleObjectBlob   =   "EDIRECIBO1.frx":06AC
         TabIndex        =   52
         Top             =   240
         Width           =   540
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   225
         Left            =   2640
         OleObjectBlob   =   "EDIRECIBO1.frx":0716
         TabIndex        =   53
         Top             =   590
         Width           =   540
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   225
         Left            =   2640
         OleObjectBlob   =   "EDIRECIBO1.frx":0780
         TabIndex        =   54
         Top             =   940
         Width           =   540
      End
   End
   Begin VB.TextBox TEXT13 
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
      Height          =   870
      Left            =   5760
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   700
      Width           =   4095
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   21
      Top             =   600
      Width           =   4695
      Begin VB.OptionButton Option2 
         Caption         =   "Todos"
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
         Left            =   2040
         TabIndex        =   68
         Top             =   650
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Solo Pendientes"
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
         Left            =   120
         TabIndex        =   67
         Top             =   650
         Value           =   -1  'True
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
         Left            =   50
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "Legajo:"
         Top             =   240
         Width           =   1605
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
         Left            =   4320
         TabIndex        =   23
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox Text1 
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
         TabIndex        =   22
         Top             =   240
         Width           =   2595
      End
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FBCA9D&
      Height          =   600
      Index           =   3
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "Acciones Varias Sobre Cuotas / Matrículas"
      Top             =   0
      Width           =   9975
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
      Height          =   1350
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   9795
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
         Left            =   6825
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "Celular de Contacto:"
         Top             =   2085
         Width           =   1815
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   8715
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   2085
         Width           =   4950
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
         Left            =   6825
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "E-Mail:"
         Top             =   1365
         Width           =   1815
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
         Left            =   6825
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "Contacto:"
         Top             =   1725
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   8715
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   1365
         Width           =   4950
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   8715
         Locked          =   -1  'True
         TabIndex        =   36
         Top             =   1725
         Width           =   4950
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
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "Celular:"
         Top             =   2085
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
         Index           =   1
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Doc.:"
         Top             =   1365
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
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "Télefono:"
         Top             =   1725
         Width           =   1815
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2085
         Width           =   4950
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1725
         Width           =   4950
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   1995
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1365
         Width           =   4950
      End
      Begin VB.TextBox TXCOMBO1 
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
         Left            =   690
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1365
         Width           =   1230
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   930
         Width           =   3000
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   570
         Width           =   3000
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   210
         Width           =   3000
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
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Localidad:"
         Top             =   930
         Width           =   1815
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
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Domicilio:"
         Top             =   570
         Width           =   1815
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
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "Nombres y Apellido:"
         Top             =   210
         Width           =   1815
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
         Left            =   5085
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Sala de:"
         Top             =   570
         Width           =   1000
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
         Left            =   5085
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Grado:"
         Top             =   210
         Width           =   1000
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
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Año:"
         Top             =   210
         Width           =   405
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Alumno Becado"
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
         Left            =   5085
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   930
         Width           =   1680
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
         Index           =   9
         Left            =   6825
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "Turno:"
         Top             =   930
         Width           =   765
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
         Index           =   10
         Left            =   6825
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "Color:"
         Top             =   570
         Width           =   765
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
         Left            =   6810
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "Letra:"
         Top             =   210
         Width           =   525
      End
      Begin VB.TextBox TXCOMBO5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   6150
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   210
         Width           =   615
      End
      Begin VB.TextBox TXCOMBO7 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   6150
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   570
         Width           =   615
      End
      Begin VB.TextBox TXCOMBO3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   7410
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   210
         Width           =   510
      End
      Begin VB.TextBox TXCOMBO8 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   7695
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   570
         Width           =   1950
      End
      Begin VB.TextBox TXCOMBO2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   7695
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   930
         Width           =   1950
      End
      Begin VB.TextBox TXCOMBO6 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1F1F1&
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
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   210
         Width           =   1130
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7080
      OleObjectBlob   =   "EDIRECIBO1.frx":07EA
      Top             =   5760
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   4140
      Left            =   105
      TabIndex        =   26
      Top             =   3045
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   7303
      _Version        =   393216
      Cols            =   4
      BackColor       =   14737632
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      AllowUserResizing=   1
   End
   Begin VB.Frame Frame4 
      Height          =   1350
      Left            =   0
      TabIndex        =   28
      Top             =   7920
      Width           =   6060
   End
   Begin VB.Label LBLIMPORTE_ORIG 
      Caption         =   "Label1"
      Height          =   135
      Left            =   4440
      TabIndex        =   69
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   195
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "EDIRECIBO1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''VARIMOVI= CUOTA
'TIPOMOVIM = PROGRAMACION


Private Sub Check4_GotFocus()
   On Error Resume Next
   Text17(1).SetFocus
   On Error GoTo 0
End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
    '
    '
    Call ABRECONEXION
    '
    Call CARGA_INDICE_ALUMNOS
    '
    If RESPUESTA_BUSQUEDA$ = "" Then GoTo 99
    '
    Text1 = RESPUESTA_BUSQUEDA$
99  Command1.Enabled = True

End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   LGJO& = XVALO(Text1)
   Text1 = ""
   Text1 = SAFETEXT$(LGJO&)
   Command2.Enabled = True
   
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   LGJO& = XVALO(Text1)
   
   If LGJO& < 1 Then MsgBox "NÚMERO DE LEGAJO NO VÁLIDO": GoTo 99
   
   CONDI$ = "LEGAJO = " & LGJO&
   If CantRegistros&("ALUMNOS", CONDI$, "NOMESS") < 1 Then
      MsgBox "NÚMERO DE LEGAJO NO INEXISTENTE O NO VÁLIDO"
      GoTo 99
   End If
   '
   'PASO DATOS A FRM DE GENERACION DE CUOTAS
   GENERACUOTA.Text1(5) = SAFETEXT(LGJO&)
   GENERACUOTA.Text1(5).Visible = True
   GENERACUOTA.Text17(16).Visible = True
   GENERACUOTA.Show 1
   '
   Call Command2_Click ' ACTUALIZA
   
99 Command4.Enabled = True

End Sub

'Private Sub Command3_Click()
'
'   Command3.Enabled = False
'   '
'   LGJO& = XVALO(Text1)
'   SBTOT# = XVALO(Text19)
'   BONIF# = XVALO(TXTBONIFICACION)
'   RCRGO# = XVALO(TXTRECARGO)
'   TTAL# = XVALO(Text23)
'   MEX% = XVALO(TXTMES)
'   ANOX% = XVALO(TXTANO)
'   MESXX$ = "00" + TRIM$(Str(MEX%)): MESXX$ = Right(MESXX$, 2)
'   VARIMOVIX$ = "RCOB"
'   TIMOVI$ = "RECIBO COBRANZA"
'   REFE$ = "RECIBO PARA CUOTA : " & MES$(MEX%) & " AÑO " & ANOX%
'      If TRIM$(TXTTIPO) = "MATRICULA" Then CUOTA_MATRICULA$ = TRIM$(TXTTIPO)
'   If CUOTA_MATRICULA$ = "MATRICULA" Then
'      REFE$ = "RECIBO PARA MATRICULA : " & " AÑO " & ANOX%
'      VARIMOVIX$ = "COBMATRI"
'   End If
'
'   CUECONTA% = 1 ' DEUDORES POR CUOTA
'   VALOR_CUOTA_MATRICULA# = XVALO(Text19)
'   TTALDEBE# = 0 ' IMPORTE PARA SEGUNDO REGISTRO QUE GENERA QUE SE GUARDA EN EL DEBE
'   VUELTA% = 0
'   '
'   If LGJO& < 1 Then
'      MsgBox "FALTA SELECCIONAR ALUMNO"
'      GoTo 99
'   End If
'   '
'   If TRIM$(Text14) = "" Then
'      MsgBox "ALUMNO INGRESADO NO VÁLIDO (FALTA NOMBRE)"
'      GoTo 99
'   End If
'   '
'   If VALOR_CUOTA_MATRICULA# = 0 Then
'      MsgBox "DEBE SELECCIONAR UNA CUOTA O MATRÍCULA A COBRAR"
'      GoTo 99
'   End If
'   '
'   If BONIF# < 0 Then
'      MsgBox "NO VÁLIDO INGRESAR BONIFICACION CON VALORES NEGATIVOS"
'      GoTo 99
'   End If
'   '
'   If TTAL# < 0 Then
'      MsgBox "NO VÁLIDO,  TOTAL RECIBO CON VALORES NEGATIVOS"
'      GoTo 99
'   End If
'   '
'   If TTAL# = 0 Then
'      MsgBox "NO VÁLIDO,  TOTAL RECIBO CON VALORES NEGATIVOS"
'      GoTo 99
'   End If
'
'   If RCRGO# < 0 Then
'      MsgBox "NO VÁLIDO INGRESAR RECARGO CON VALORES NEGATIVOS"
'      GoTo 99
'   End If
'   '
'   'TRAE EL NUMERO DE RECIBO SEGUN SI ES MATRICULA O SI ES CUOTA
'   If CUOTA_MATRICULA$ = "MATRICULA" Then
'    NURECI& = NURECIBO&(FORMATEADO$, "MATRICULA")
'    COMLAR$ = FORMATEADO$
'   Else
'    NURECI& = NURECIBO&(FORMATEADO$)
'    COMLAR$ = FORMATEADO$
'   End If
'   '
'   '
'   SQLX$ = "SELECT * FROM CAJABANC_COL"
'   SQLX$ = SQLX$ + " WHERE LEGAJO < 1"
'
'   Set AUX = New ADODB.Recordset
'   AUX.Open (SQLX$), BASSSEX, adOpenKeyset, adLockPessimistic, adCmdText
'   With AUX
'30        .AddNew
'          !LEGAJO = LGJO&
'          !MES_CUOTA = MEX% ' ESTE ES EL MES DE LA CUOTA
'          !ANO_CUOTA = ANOX% 'ESTE ES EL ANO DE LA CUOTA
'          !TIPOMOVIM = TIMOVI$
''               REFE$ = REFE$
'          !REFERENCIA = Left$(REFE$, 64)
'          !CUENTA_CONT = CUECONTA%
'          !VARIMOVI = VARIMOVIX$
'          !CODICOM_LAR = COMLAR$
'
'          !NUMEDOC = NURECI&
'          !HABER = TTAL# 'LA SEGUNDA VEZ ESTA EN 0
'          !DEBE = TTALDEBE# 'LA PRIMERA VEZ ESTA EN 0 PASA DEVUELTA Y GRABA NEW REGISTRO CON OTRA CUENTA CONTABLE
'          !NUME_USUARIO = NUMEUSER%
'          !FECHEMISI = CVDAT(HOY(HO$))
'          !BONIFICACION = (-1) * BONIF# ' LO GUARDO EN NEGATIVO
'          !RECARGO = RCRGO#
'          !CUOTA_ORIGINAL = XVALO(Text28)
'          !ENSE_PROGRAMADA = XVALO(Text5)
'          !INGLES = XVALO(Text7)
'          !EMERGENCIA = XVALO(Text6)
'          !COMPUTACION = XVALO(Text20)
'          !DOCU_APLI = TRIM$(TXTDCTO)
'          !FeReMovi = Now
'          .Update
'          If VUELTA% < 1 Then
'             VUELTA% = 1
'             CUECONTA% = 2
'             VARIMOVIX$ = "CHRECI"
'             TIMOVI$ = "INGRESO A CAJA"
'             TTALDEBE# = TTAL#
'             TTAL# = 0
'             !FeReMovi = Now
'             GoTo 30
'          End If
'      End With
'      AUX.Close
'      Set AUX = Nothing
'      '
'      'GRABO NOTA DE ALUMNO
'      TXT$ = "Recibo de Cobro: " & COMLAR$ & " vbcrlf"
'      TXT$ = TXT$ + REFE$
'      Call ANOTAX_ALUMNO(LEGAJO&, TXT$)
'      '
'      If TRIM$(TXTTIPO) = "MATRICULA" Then
'         Call PRIMATRICULA(LGJO&, NURECI&, FORMATEADO$, "COBMATRI")
'      Else
'         Call PRIRECIBO(LGJO&, NURECI&, FORMATEADO$, "RCOB")
'      End If
'      '
'      MsgBox "HECHO"
'      '
'      Text1 = ""
'      Text1 = LGJO&
'
'99    Command3.Enabled = True
'End Sub
Private Sub Form_Load()
    Call APLICAR_SKIN(Me, App.Path & "\SKINS\dogmax.skn")
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
    LBLIMPORTE_ORIG = ""
    Timer1.Enabled = True
End Sub
Private Sub Timer1_Timer()
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
    Timer1.Enabled = False
End Sub


Private Sub MSF_Click()
'   Exit Sub
'   LGJO& = XVALO(Text1)
'   'ASIGNO POSICION SELECCIONADA A VARIABLES
'   i& = MSF.MouseRow
'   j& = MSF.MouseCol
'   '
'   If i& < 1 Then Exit Sub
'
'   'VALIDO
'   If i& = 0 Or i& > MSF.Rows Then Exit Sub
'   If j& > MSF.Cols Then Exit Sub
'   '
'   'RESTAURO AL COLOR ORIGINAL LA COLUMNA 0
'   MSF.Col = 0
'   For H& = 0 To MSF.Rows - 1
'     MSF.Row = H&
'     MSF.CellBackColor = &H8000000F
'   Next H&
'   '
'   'PINTO DE AMARILLO LA SELECCIONADA
'   MSF.Row = i&
'   MSF.CellBackColor = &H80FFFF   ' LA PINTA DE AMARILLA
'
'   If LGJO& < 1 Then Exit Sub
'   '
'   FECH% = FECHANUM(TRIM$(MSF.TextMatrix(MSF.Row, 1)))
'   DOCU$ = TRIM$(MSF.TextMatrix(MSF.Row, 2))
'   MEX% = XVALO(MSF.TextMatrix(MSF.Row, 3))
'   ANOX% = XVALO(TRIM$(MSF.TextMatrix(MSF.Row, 4)))
'   VARMOV$ = TRIM$(MSF.TextMatrix(MSF.Row, 5))
'   VALOR_CUOTA_ORIGINAL# = XVALO(MSF.TextMatrix(MSF.Row, 6))
'   HABE# = XVALO(MSF.TextMatrix(MSF.Row, 7))
'   '
'
'   SQLX$ = "SELECT * FROM CAJABANC_COL"
'   SQLX$ = SQLX$ + " WHERE LEGAJO = " & LGJO&
'   SQLX$ = SQLX$ + " AND (CODICOM_LAR = '" & DOCU$ & "'"
'   If TRIM$(DOCU$) = "" Then SQLX$ = SQLX$ + " OR CODICOM_LAR IS NULL "
'   SQLX$ = SQLX$ + ") AND MES_CUOTA = " & MEX%
'   SQLX$ = SQLX$ + " AND ANO_CUOTA = " & ANOX%
'   SQLX$ = SQLX$ + " AND VARIMOVI = '" & VARMOV$ & "'"
'
'   Set CTA_CTETMP = BaDaABRE_ADMIN_ALQ.OpenRecordset(SQLX$, dbOpenSnapshot)
'   With CTA_CTETMP
'      Do While Not .EOF
'        IMPO$ = FORMATNUM$(XVALO(!DEBE), "F12.2")
'        ENSE_PROGRAMADA# = XVALO(!ENSE_PROGRAMADA)
'        EMERGENCIA# = XVALO(!EMERGENCIA)
'        INGLES# = XVALO(!INGLES)
'        COMPUTACION# = XVALO(!COMPUTACION)
'        TIMOV$ = SAFETEXT$(!TIPOMOVIM)
'
'
'        Text5 = ENSE_PROGRAMADA#
'        Text6 = EMERGENCIA#
'        Text7 = INGLES#
'        Text20 = COMPUTACION#
'        .MoveNext
'      Loop
'   End With
'   '
'
'
'   If VARMOV$ = "MATRICULA" Then
'       Text19 = IMPO$
'   Else
'       Text19 = XVALO(Text5) + XVALO(Text6) + XVALO(Text7) + XVALO(Text20)
'   End If
'
'   TXTTIPO = VARMOV$
'   TXTMES = SAFETEXT$(MEX%)
'   TXTANO = SAFETEXT$(ANOX%)
'   TXTDCTO = DOCU$
'   Text28 = TRIM$(FORMATNUM$(VALOR_CUOTA_ORIGINAL#, "F12.2"))
''   Call SUMARSUBTOTAL
''   Call SUMARTOT
End Sub

Private Sub MSF_DblClick()
   FRMEDIT.LBLRESULTADO = ""
   LBLIMPORTE_ORIG = ""

   LGJO& = XVALO(Text1)
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   i& = MSF.MouseRow
   j& = MSF.MouseCol
   '
   If i& < 1 Then Exit Sub
   
   'VALIDO
   If i& = 0 Or i& > MSF.Rows Then Exit Sub
   If j& > MSF.Cols Then Exit Sub
   '
   'RESTAURO AL COLOR ORIGINAL LA COLUMNA 0
   MSF.COL = 0
   For H& = 0 To MSF.Rows - 1
     MSF.Row = H&
     MSF.CellBackColor = &H8000000F
   Next H&
   '
   'PINTO DE AMARILLO LA SELECCIONADA
   MSF.Row = i&
   MSF.CellBackColor = &H80FFFF   ' LA PINTA DE AMARILLA

   If LGJO& < 1 Then Exit Sub
   '

   FRMEDIT.Text1.Text = EDIRECIBO.MSF.TextMatrix(EDIRECIBO.MSF.Row, 1)
   FRMEDIT.TEXT2.Text = EDIRECIBO.MSF.TextMatrix(EDIRECIBO.MSF.Row, 2)
   FRMEDIT.Text3.Text = EDIRECIBO.MSF.TextMatrix(EDIRECIBO.MSF.Row, 3)
   FRMEDIT.Text4.Text = EDIRECIBO.MSF.TextMatrix(EDIRECIBO.MSF.Row, 4)
   IMPODEBE# = XVALO(MSF.TextMatrix(EDIRECIBO.MSF.Row, 6))
   FRMEDIT.Text5.Text = EDIRECIBO.MSF.TextMatrix(EDIRECIBO.MSF.Row, 5)
   FRMEDIT.Text6.Text = TRIM$(FORMATNUM$(IMPODEBE#, "F16.2"))
   FRMEDIT.Text7.Text = EDIRECIBO.MSF.TextMatrix(EDIRECIBO.MSF.Row, 7)
   FRMEDIT.Text8.Text = EDIRECIBO.MSF.TextMatrix(EDIRECIBO.MSF.Row, 8)
   FRMEDIT.Text9.Text = SAFETEXT$(LGJO&)
   LBLIMPORTE_ORIG = TRIM$(FORMATNUM$(IMPODEBE#, "F16.2"))
   FRMEDIT.Show 1
   Call GRABAMODIFICACION(FRMEDIT.LBLRESULTADO)
     
End Sub


Private Sub MSF_SelChange()
'   Text7 = TRIM$(MSF.TextMatrix(MSF.Row, 0))
   '
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   '
   i& = MSF.Row
   j& = MSF.COL
   '
   Call SELECCION_COLUMNA_0(MSF, i&, j&)
   
End Sub


Private Sub Option1_Click()
   Call Text1_Change
End Sub

Private Sub Option2_Click()
   Call Text1_Change
End Sub

Private Sub Text1_Change()
     Call LIMPIAR
     MSF.Rows = 1
     LGJO& = XVALO(Text1)
     If LGJO& < 1 Then Exit Sub
     
     Z& = NURECIBO&(FORMATEADO$)
     '
     Call CARGA_DATOS_ALUMNMOS
     Call CARGA_GRILLA_CTA_CTE(LGJO&)

End Sub
Sub CARGA_GRILLA_CTA_CTE(LEGAJO&)
     
     '
     LGJO& = LEGAJO&
     MSF.Rows = 1
     MSF.Cols = 9
     MSF.Redraw = False
     SQLX$ = "SELECT DEBE , HABER ,CODICOM_LAR,"
     SQLX$ = SQLX$ + " FECHEMISI,MES_CUOTA, ANO_CUOTA, VARIMOVI, NUMEDOC"
     SQLX$ = SQLX$ + " FROM CAJABANC_COL WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE LEGAJO = " & LGJO&
     SQLX$ = SQLX$ + " AND (TIPOMOVIM = 'CUOTA' OR  TIPOMOVIM = 'MATRICULA')"
     SQLX$ = SQLX$ + " AND (VARIMOVI = 'CUOTA' OR VARIMOVI = 'MATRICULA')"
     
     SQLX$ = SQLX$ + " order by ANO_CUOTA, MES_CUOTA"
     
     Set CTA_CTETMP = BASSSEX.Execute(SQLX$)
     With CTA_CTETMP
        Do While Not .EOF
           FEMI% = CVINT(!FECHEMISI)
           MEX% = XVALO(!MES_CUOTA)
           DOCUM$ = SAFETEXT$(!CODICOM_LAR)
           ANOX% = XVALO(!ANO_CUOTA)
           CONCEP = SAFETEXT$(!VARIMOVI)
           DEBEX = XVALO(!DEBE)
           PAGOALUMNO# = SUPAGOX#(DOCUM$) ' DEVUELVE LOS PAGOS REFERENCIADOS AL DOCUMENTO
           HABEX = PAGOALUMNO#
           SALDO = DEBEX - HABEX
           If Option1.Value = True Then
               If SALDO < 0.05 Then GoTo 50
           End If
           '
           FILA& = FILA& + 1
           MSF.Rows = FILA& + 1
           MSF.TextMatrix(FILA&, 1) = FECHATEX$(FEMI%)
           MSF.TextMatrix(FILA&, 2) = DOCUM$
           MESXX$ = "00" & TRIM$(Str(MEX%))
           MESXX$ = Right(MESXX$, 2)
           MSF.TextMatrix(FILA&, 3) = MESXX$
           MSF.TextMatrix(FILA&, 4) = SAFETEXT$(ANOX%)
           MSF.TextMatrix(FILA&, 5) = CONCEP
           MSF.TextMatrix(FILA&, 6) = FORMATNUM$(DEBEX, "F10.2")
           MSF.TextMatrix(FILA&, 7) = FORMATNUM$(HABEX, "F10.2")
           MSF.TextMatrix(FILA&, 8) = FORMATNUM$(SALDO, "F10.2")
           '
50      .MoveNext
        Loop
     End With
     CTA_CTETMP.Close
     Set CTA_CTETMP = Nothing
     NUCOLU% = 1: TEPRUEBA = "10": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = ""
     NUCOLU% = 2: TEPRUEBA = "10/10/20120": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Fecha Emis."
     NUCOLU% = 3: TEPRUEBA = "123456789000000000000": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Documento"
     NUCOLU% = 4: TEPRUEBA = "ABCDEFGHIH": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Mes"
     NUCOLU% = 5: TEPRUEBA = "CHAPA123456": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Año"
     NUCOLU% = 6: TEPRUEBA = "123456789012": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Concepto"
     NUCOLU% = 7: TEPRUEBA = "123456789012": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Debe"
     NUCOLU% = 8: TEPRUEBA = "123456789012": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Haber"
     NUCOLU% = 9: TEPRUEBA = "123456789012": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Saldo"
     MSF.Redraw = True
End Sub
   '
Sub CARGA_DATOS_ALUMNMOS()

   LEGAJO& = XVALO(Text1)
   If LEGAJO& <= 0 Then Exit Sub
   CONDI$ = "LEGAJO = " & LEGAJO& & "AND STATUS >= 0"
   
   If CantRegistros&("ALUMNOS", CONDI$, "NOMESS") < 1 Then Exit Sub

   SQLX$ = "SELECT * FROM ALUMNOS WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " + CONDI$
   Set ALUMTMP = BASSSEX.Execute(SQLX$)
   With ALUMTMP
        Text14 = SAFETEXT$(!NOMBRE_APELLIDO)
        TEXT2 = SAFETEXT$(!DOMCILIO)
        Text3 = SAFETEXT$(!LOCALIDAD)
        Text12.Text = SAFETEXT$(!documento)
        Text11.Text = SAFETEXT$(!TELEFONO)
        Text10.Text = SAFETEXT$(!CELULAR)
        Text8.Text = SAFETEXT$(!MAIL)
        '
        Text9.Text = SAFETEXT$(!CONTACTO)
        Text18.Text = SAFETEXT$(!CELU_CONTACTO)

        '
        TXCOMBO5.Text = CEROIGUALVACIO(XVALO(!GRADO))
        If XVALO(TXCOMBO5.Text) > 0 Then TXCOMBO5.Text = TXCOMBO5.Text & "º"
        TXCOMBO7.Text = CEROIGUALVACIO(XVALO(!SALA_DE))
          
        COMPLETANUM$ = "to"
          If XVALO(!AÑO) = 1 Or XVALO(!AÑO) = 3 Then COMPLETANUM$ = "ro"
          If XVALO(!AÑO) = 2 Then COMPLETANUM$ = "do"
        TXCOMBO6.Text = CEROIGUALVACIO(XVALO(!AÑO))
          If XVALO(!AÑO) > 0 Then TXCOMBO6.Text = TXCOMBO6.Text & COMPLETANUM$
        TXCOMBO7.Text = CEROIGUALVACIO(XVALO(!SALA_DE))
        
        TXCOMBO2.Text = SAFETEXT$(!TURNO)
        TXCOMBO3.Text = SAFETEXT$(!LETRA)
        TXCOMBO8.Text = SAFETEXT$(!SALA_COLOR)
        '
        TXCOMBO1(0).Text = SAFETEXT$(!TIPODOCU_ALUM)
        Check4.Value = XVALO(!BECADO)
        '
        TEXT13 = SAFETEXT$(!NOTA)
    End With
    ALUMTMP.Close
    Set ALUMTMP = Nothing

End Sub
Sub LIMPIAR()

For Each CONTROLX In RECIBO_COLEGIO.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
        If TypeOf CONTROLX Is TextBox Then   ' SI ES CAJA DE TEXTO
           If UCase$(CONTROLX.Name) <> "TEXT1" And (CONTROLX.BackColor = &HF1F1F1 Or CONTROLX.BackColor = &H80000005) Then
              CONTROLX.Text = ""
           End If
        End If
Next
End Sub

Sub GRABAMODIFICACION(MODO$)
   If TRIM$(MODO$) = "" Then Exit Sub
   
   fech% = FECHANUM(TRIM$(FRMEDIT.Text1)): FECHAX = ENROQUE_DE_FECHAS(FECHATEXLAR$(fech%))
   DOCU$ = TRIM$(FRMEDIT.TEXT2)
   MEX% = XVALO(FRMEDIT.Text3)
   ANOX% = XVALO(TRIM$(FRMEDIT.Text4))
   VARMOV$ = TRIM$(FRMEDIT.Text5)
   VALOR_CUOTA_MODIFICADO# = XVALO(FRMEDIT.Text6)
   HABE# = XVALO(FRMEDIT.Text7)
   LGJO& = XVALO(FRMEDIT.Text9)
   
   CONDI$ = "  LEGAJO = " & LGJO&
   CONDI$ = CONDI$ + " AND VARIMOVI = '" & VARMOV$ & "'"
   CONDI$ = CONDI$ + " AND MES_CUOTA = " & MEX%
   CONDI$ = CONDI$ + " AND ANO_CUOTA = " & ANOX%
   CONDI$ = CONDI$ + " AND CODICOM_LAR = '" & DOCU$ & "'"
   CONDI$ = CONDI$ + " AND FECHEMISI =  #" & FECHAX & "#"
   '
   'VALIDO SI SOLO ES UN REGISTRO EL QUE DEVUELVE LA CONDICION
   'CONSIDERANDO QUE SOLO DEBE SER 1 SI NO HAY ALGO QUE ESTA MAL
   If CantRegistros&("CAJABANC_COL", CONDI$, "MOMESS") <> 1 Then
      TX$ = "ATENCIÓN!!! TRANSACCIÓN IMPOSIBLE DE REALIZAR " + vbCrLf
      TX$ = TX$ + "POSIBLE INCONSISTENCIA DE DATOS " + vbCrLf
      TX$ = TX$ + "INFORME ESTA SITUACIÓN AL PROVEEDOR DEL SISTEMA" + vbCrLf
      MsgBox TX$
      
      Exit Sub
   End If
   '
   If MODO$ = "ELIMINA" Then
       Call BORRAREGISTROS("CAJABANC_COL", CONDI$, REGAF&, "NOMESS")
       If REGAF& > 0 Then
        'GRABO NOTA DE ALUMNO
        TX$ = "ANULACION DE " & VARMOV$ & ": " & DOCU$ & vbCrLf
          Call ANOTAX_ALUMNO(LGJO&, TX$)
       End If
   End If
   '
   If MODO$ = "GRABA" Then
      SQLX$ = " UPDATE CAJABANC_COL SET DEBE = " & FORMATNUM$(VALOR_CUOTA_MODIFICADO#, "F16.2")
      SQLX$ = SQLX$ + " WHERE " & CONDI$
      BASSSEX.Execute (SQLX$)
      'GRABO NOTA DE ALUMNO
      TX$ = "SE MODIFICO " & VARMOV$ & ": " & DOCU$ & vbCrLf
      TX$ = TX$ + "CON EL IMPORTE ORIGINAL DE : $ " & LBLIMPORTE_ORIG
      TX$ = TX$ + "LLEVANDO EL IMPORTE A : $ " & TRIM$(FORMATNUM$(VALOR_CUOTA_MODIFICADO#, "F16.2"))
      Call ANOTAX_ALUMNO(LGJO&, TX$)
   End If
   '
   '

   Text1_Change ' REFRESCO LA PANTALLA
   
   MsgBox "HECHO"

End Sub


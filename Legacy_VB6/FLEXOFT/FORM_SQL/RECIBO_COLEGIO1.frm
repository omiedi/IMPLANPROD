VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form RECIBO_COLEGIO1 
   Caption         =   "RECIBO"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   13905
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTP1 
      Height          =   330
      Left            =   12285
      TabIndex        =   60
      Top             =   1260
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   130088961
      CurrentDate     =   41123
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   225
      Left            =   12285
      OleObjectBlob   =   "RECIBO_COLEGIO1.frx":0000
      TabIndex        =   59
      Top             =   630
      Width           =   1065
   End
   Begin VB.TextBox Text24 
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
      Left            =   12285
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   860
      Width           =   1335
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
      Left            =   8520
      TabIndex        =   57
      Top             =   7920
      Width           =   5130
   End
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   8350
      TabIndex        =   54
      Top             =   7080
      Width           =   5535
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
         Index           =   6
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Text            =   "TOTAL"
         Top             =   180
         Width           =   1400
      End
      Begin VB.TextBox Text23 
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
         Left            =   4170
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   180
         Width           =   1200
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2895
      Left            =   8350
      TabIndex        =   41
      Top             =   4080
      Width           =   5535
      Begin VB.TextBox TXTRECARGO 
         Alignment       =   1  'Right Justify
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
         Left            =   4170
         TabIndex        =   74
         Top             =   2415
         Width           =   1200
      End
      Begin VB.TextBox TXTBONIFICACION 
         Alignment       =   1  'Right Justify
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
         Left            =   4170
         TabIndex        =   73
         Top             =   2040
         Width           =   1200
      End
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
         Left            =   630
         Locked          =   -1  'True
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   2310
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
         Left            =   630
         Locked          =   -1  'True
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   1890
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
         Left            =   630
         Locked          =   -1  'True
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   1470
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
         Left            =   630
         Locked          =   -1  'True
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1050
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
         Left            =   630
         Locked          =   -1  'True
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   630
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
         Left            =   630
         Locked          =   -1  'True
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   240
         Width           =   1850
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
         Index           =   5
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Text            =   "RECARGO"
         Top             =   2400
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
         Index           =   4
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Text            =   "BONIFICACIÓN"
         Top             =   2040
         Width           =   1400
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "COMPUTACIÓN"
         Top             =   1320
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Text            =   "VALOR CUOTA"
         Top             =   1680
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Text            =   "INGLÉS"
         Top             =   960
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "EMERGENCIAS"
         Top             =   600
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
         Left            =   4170
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   1320
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
         Left            =   4170
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   1680
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
         Left            =   4170
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   960
         Width           =   1200
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
         Left            =   4170
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   600
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
         Left            =   4170
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   240
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "Ens.Programada"
         Top             =   240
         Width           =   1400
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   225
         Left            =   180
         OleObjectBlob   =   "RECIBO_COLEGIO1.frx":0070
         TabIndex        =   61
         Top             =   350
         Width           =   540
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   225
         Left            =   200
         OleObjectBlob   =   "RECIBO_COLEGIO1.frx":00D8
         TabIndex        =   64
         Top             =   735
         Width           =   540
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   225
         Left            =   220
         OleObjectBlob   =   "RECIBO_COLEGIO1.frx":013E
         TabIndex        =   66
         Top             =   1155
         Width           =   540
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   225
         Left            =   105
         OleObjectBlob   =   "RECIBO_COLEGIO1.frx":01A4
         TabIndex        =   68
         Top             =   1575
         Width           =   540
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   225
         Left            =   105
         OleObjectBlob   =   "RECIBO_COLEGIO1.frx":020E
         TabIndex        =   70
         Top             =   1995
         Width           =   540
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   225
         Left            =   105
         OleObjectBlob   =   "RECIBO_COLEGIO1.frx":0278
         TabIndex        =   72
         Top             =   2415
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
      Left            =   5160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   38
      Top             =   700
      Width           =   6975
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   33
      Top             =   600
      Width           =   4935
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
         TabIndex        =   36
         TabStop         =   0   'False
         Text            =   "Legajo:"
         Top             =   360
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
         Left            =   4440
         TabIndex        =   35
         Top             =   360
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
         Left            =   1800
         TabIndex        =   34
         Top             =   360
         Width           =   2595
      End
   End
   Begin VB.TextBox Text17 
      Alignment       =   2  'Center
      BackColor       =   &H00400000&
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
      Height          =   555
      Index           =   3
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Text            =   "RECIBO DE COBRANZA"
      Top             =   0
      Width           =   13695
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
      Height          =   2430
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   13755
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
         Left            =   7370
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   210
         Width           =   1230
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
         Left            =   8670
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   210
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
         Left            =   8670
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   570
         Width           =   4950
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
         Left            =   8670
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   930
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
         Left            =   8670
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1650
         Width           =   4950
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
         Left            =   8670
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1290
         Width           =   4950
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
         TabIndex        =   26
         Top             =   930
         Width           =   4695
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
         TabIndex        =   25
         Top             =   570
         Width           =   4695
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
         TabIndex        =   24
         Top             =   210
         Width           =   4695
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Nombres y Apellido:"
         Top             =   210
         Width           =   1815
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
         Left            =   6780
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "Télefono:"
         Top             =   570
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
         Left            =   6780
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Doc.:"
         Top             =   210
         Width           =   555
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
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "Sala de:"
         Top             =   1650
         Width           =   1815
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
         Left            =   6780
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "Celular:"
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
         Index           =   2
         Left            =   6780
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Contacto:"
         Top             =   1650
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
         Index           =   2
         Left            =   6780
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "E-Mail:"
         Top             =   1290
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
         Left            =   8670
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2010
         Width           =   4950
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
         Left            =   6780
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Celular de Contacto:"
         Top             =   2010
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
         Index           =   6
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Grado:"
         Top             =   1290
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
         Index           =   8
         Left            =   4220
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "Año:"
         Top             =   1290
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
         Left            =   75
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1995
         Width           =   1800
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
         Left            =   3855
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "Turno:"
         Top             =   2015
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
         Left            =   3855
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "Color:"
         Top             =   1650
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "Letra:"
         Top             =   1290
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
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1300
         Width           =   850
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
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1690
         Width           =   850
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
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1300
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
         Left            =   4725
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1650
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
         Left            =   4725
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2015
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
         Left            =   4725
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1310
         Width           =   1950
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7080
      OleObjectBlob   =   "RECIBO_COLEGIO1.frx":02E2
      Top             =   5760
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   4365
      Left            =   120
      TabIndex        =   39
      Top             =   4155
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7699
      _Version        =   393216
      Cols            =   4
      BackColor       =   14737632
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   195
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "RECIBO_COLEGIO1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'VARIMOVI= CUOTA
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

Private Sub Command3_Click()

   Command3.Enabled = False
   LGJO& = XVALO(Text1)
   '
   GRAAÑO$ = GRADO_AÑO$(LGJO&)
   SAGRADO = SALA_GRADO_AÑO$(LGJO&) 'SALA DE , GRADO O AÑO
   LETRX$ = LETRAX$(LGJO&)
   TUXNO$ = TURNO$(LGJO&)
   '
   'PRESENTA FRM DE MODIFICACION DE ESTOS DATOS PARA LA IMPRESION DE LA MATRICULA
   'POR QUE EL ALUMNO CAPAS ESTA EN 4 Y LA MATRICULA ES PARA 5
   If TRIM$(TXTTIPO) = "MATRICULA" Then
        MODIGRADO.Text3 = GRAAÑO$
        MODIGRADO.TXTPARA = SAGRADO
        MODIGRADO.TXTLETRA = LETRX$
        MODIGRADO.TXTTURNO = TUXNO$
        MODIGRADO.Show 1
        'EN CASO QUE  NO CANCELA OSEA QUE PRESIONO ACEPTAR
        If MODIGRADO.CANCELA = False Then
            GRAAÑO$ = TRIM$(MODIGRADO.Text3)
            SAGRADO = TRIM$(MODIGRADO.TXTPARA) 'SALA DE , GRADO O AÑO
            LETRX$ = TRIM$(MODIGRADO.TXTLETRA)
            TUXNO$ = TRIM$(MODIGRADO.TXTTURNO)
        Else
            Unload MODIGRADO
            GoTo 99
        End If
   End If
   '
   Call CARGAVECTOR_MES
   SBTOT# = XVALO(Text19)
   BONIF# = XVALO(TXTBONIFICACION)
   RCRGO# = XVALO(TXTRECARGO)
   TTAL# = XVALO(Text23)
   MEX% = XVALO(TXTMES)
   ANOX% = XVALO(TXTANO)
   MESXX$ = "00" + TRIM$(Str(MEX%)): MESXX$ = Right(MESXX$, 2)
   VARIMOVIX$ = "RCOB"
   TiMovi$ = "RECIBO COBRANZA"
   REFE$ = "RECIBO PARA CUOTA : " & MES$(MEX%) & " AÑO " & ANOX%
      If TRIM$(TXTTIPO) = "MATRICULA" Then CUOTA_MATRICULA$ = TRIM$(TXTTIPO)
   If CUOTA_MATRICULA$ = "MATRICULA" Then
      REFE$ = "RECIBO PARA MATRICULA : " & " AÑO " & ANOX%
      VARIMOVIX$ = "COBMATRI"
   End If
   
   CUECONTA% = 1 ' DEUDORES POR CUOTA
   VALOR_CUOTA_MATRICULA# = XVALO(Text19)
   TTALDEBE# = 0 ' IMPORTE PARA SEGUNDO REGISTRO QUE GENERA QUE SE GUARDA EN EL DEBE
   VUELTA% = 0
   '
   If LGJO& < 1 Then
      MsgBox "FALTA SELECCIONAR ALUMNO"
      GoTo 99
   End If
   '
   If TRIM$(Text14) = "" Then
      MsgBox "ALUMNO INGRESADO NO VÁLIDO (FALTA NOMBRE)"
      GoTo 99
   End If
   '
   If VALOR_CUOTA_MATRICULA# = 0 Then
      MsgBox "DEBE SELECCIONAR UNA CUOTA O MATRÍCULA A COBRAR"
      GoTo 99
   End If
   '
   If BONIF# < 0 Then
      MsgBox "NO VÁLIDO INGRESAR BONIFICACION CON VALORES NEGATIVOS"
      GoTo 99
   End If
   '
   If TTAL# < 0 Then
      MsgBox "NO VÁLIDO,  TOTAL RECIBO CON VALORES NEGATIVOS"
      GoTo 99
   End If
   '
   If TTAL# = 0 Then
      MsgBox "NO VÁLIDO,  TOTAL RECIBO CON VALORES NEGATIVOS"
      GoTo 99
   End If

   If RCRGO# < 0 Then
      MsgBox "NO VÁLIDO INGRESAR RECARGO CON VALORES NEGATIVOS"
      GoTo 99
   End If
   '
   'TRAE EL NUMERO DE RECIBO SEGUN SI ES MATRICULA O SI ES CUOTA
   If CUOTA_MATRICULA$ = "MATRICULA" Then
    NURECI& = NURECIBO&(FORMATEADO$, "MATRICULA")
    COMLAR$ = FORMATEADO$
   Else
    NURECI& = NURECIBO&(FORMATEADO$)
    COMLAR$ = FORMATEADO$
   End If
   '
   '
   SQLX$ = "SELECT * FROM CAJABANC_COL"
   SQLX$ = SQLX$ + " WHERE LEGAJO < 1"
      
   Set AUX = New ADODB.Recordset
   AUX.Open (SQLX$), BASSSEX, adOpenKeyset, adLockPessimistic, adCmdText
   With AUX
30        .AddNew
          !LEGAJO = LGJO&
          !MES_CUOTA = MEX% ' ESTE ES EL MES DE LA CUOTA
          !ANO_CUOTA = ANOX% 'ESTE ES EL ANO DE LA CUOTA
          !TipoMovim = TiMovi$
'               REFE$ = REFE$
          !REFERENCIA = Left$(REFE$, 64)
          !CUENTA_CONT = CUECONTA%
          !VARIMOVI = VARIMOVIX$
          !CODICOM_LAR = COMLAR$
          
          !NUMEDOC = NURECI&
          !HABER = TTAL# 'LA SEGUNDA VEZ ESTA EN 0
          !DEBE = TTALDEBE# 'LA PRIMERA VEZ ESTA EN 0 PASA DEVUELTA Y GRABA NEW REGISTRO CON OTRA CUENTA CONTABLE
          !NUME_USUARIO = NUMEUSER%
          !FECHEMISI = CVDAT(HOY(HO$))
          !BONIFICACION = (-1) * BONIF# ' LO GUARDO EN NEGATIVO
          !RECARGO = RCRGO#
          !CUOTA_ORIGINAL = XVALO(Text28)
          !ENSE_PROGRAMADA = XVALO(Text5)
          !INGLES = XVALO(Text7)
          !EMERGENCIA = XVALO(Text6)
          !COMPUTACION = XVALO(Text20)
          !DOCU_APLI = TRIM$(TXTDCTO)
          !FeReMovi = Now
          'INICIALIZO
          !SALA_DE = 0
          !GRADO = 0
          !AÑO = 0
          Select Case MODIGRADO.TXTPARA
            Case "SALA DE"
               !SALA_DE = MODIGRADO.Text3
            Case "GRADO"
               !GRADO = MODIGRADO.Text3
            Case "AÑO"
               !AÑO = MODIGRADO.Text3
          End Select
          !LETRA = TRIM$(MODIGRADO.TXTLETRA)
          !TURNO = TRIM$(MODIGRADO.TXTTURNO)
          '
          If VUELTA% < 1 Then
             VUELTA% = 1
             CUECONTA% = 2
             VARIMOVIX$ = "CHRECI"
             TiMovi$ = "INGRESO A CAJA"
             TTALDEBE# = TTAL#
             TTAL# = 0
             !FeReMovi = Now
             GoTo 30
          End If
          .Update
      End With
      AUX.Close
      Set AUX = Nothing
      '
      'GRABO NOTA DE ALUMNO
      TXT$ = "Recibo de Cobro: " & COMLAR$ & vbCrLf
      TXT$ = TXT$ + REFE$
      Call ANOTAX_ALUMNO(LGJO&, TXT$)
      '
      If TRIM$(TXTTIPO) = "MATRICULA" Then
         Call PRIMATRICULA(LGJO&, NURECI&, FORMATEADO$, "COBMATRI")
      Else
         Call PRIRECIBO(LGJO&, NURECI&, FORMATEADO$, "RCOB")
      End If
      '
      MsgBox "HECHO"
      Z& = NURECIBO&(FORMATEADO$) 'MUESTRA PROXIMO NUMERO DE RECIBO
      Text24 = FORMATEADO$
      '
      Text1 = ""
      Text1 = LGJO&
      
      'DESCARGO EL FRM DE MODIGRADO
      Unload MODIGRADO

99    Command3.Enabled = True
End Sub

Private Sub Form_Load()
    Call APLICAR_SKIN(Me, App.Path & "\SKINS\dogmax.skn")
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
'    SQLX$ = "alter table CAJABANC_COL drop column REGARGO"
''    BASSSEX.Execute (SQLX$)
    DTP1.Value = Date
    Z& = NURECIBO&(FORMATEADO$)
    Text24 = FORMATEADO$

End Sub

Private Sub MSF_Click()
   '
   'valido por si se hace clic en vacio
   fech% = FECHANUM(TRIM$(MSF.TextMatrix(MSF.Row, 1)))
   DOCU$ = TRIM$(MSF.TextMatrix(MSF.Row, 2))
   If fech% < 1 Or DOCU$ = "" Then Exit Sub
   '
   LGJO& = XVALO(Text1)
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   I& = MSF.MouseRow
   j& = MSF.MouseCol
   '
   Call LIMPIAR_FRAME_VALORES
   If I& < 1 Then Exit Sub
   
   'VALIDO
   If I& = 0 Or I& > MSF.Rows Then Exit Sub
   If j& > MSF.Cols Then Exit Sub
   '
   'RESTAURO AL COLOR ORIGINAL LA COLUMNA 0
   MSF.COL = 0
   For h& = 0 To MSF.Rows - 1
     MSF.Row = h&
     MSF.CellBackColor = &H8000000F
   Next h&
   '
   'PINTO DE AMARILLO LA SELECCIONADA
   MSF.Row = I&
   MSF.CellBackColor = &H80FFFF   ' LA PINTA DE AMARILLA

   If LGJO& < 1 Then Exit Sub
   '
   fech% = FECHANUM(TRIM$(MSF.TextMatrix(MSF.Row, 1)))
   DOCU$ = TRIM$(MSF.TextMatrix(MSF.Row, 2))
   MEX% = XVALO(MSF.TextMatrix(MSF.Row, 3))
   ANOX% = XVALO(TRIM$(MSF.TextMatrix(MSF.Row, 4)))
   VARMOV$ = TRIM$(MSF.TextMatrix(MSF.Row, 5))
   VALOR_CUOTA_ORIGINAL# = XVALO(MSF.TextMatrix(MSF.Row, 6))
   ACUENTA# = XVALO(MSF.TextMatrix(MSF.Row, 7))
   '
   SQLX$ = "SELECT * FROM CAJABANC_COL"
   SQLX$ = SQLX$ + " WHERE LEGAJO = " & LGJO&
   SQLX$ = SQLX$ + " AND (CODICOM_LAR = '" & DOCU$ & "'"
   If TRIM$(DOCU$) = "" Then SQLX$ = SQLX$ + " OR CODICOM_LAR IS NULL "
   SQLX$ = SQLX$ + ") AND MES_CUOTA = " & MEX%
   SQLX$ = SQLX$ + " AND ANO_CUOTA = " & ANOX%
   SQLX$ = SQLX$ + " AND VARIMOVI = '" & VARMOV$ & "'"
     
   Set CTA_CTETMP = BASSSEX.Execute(SQLX$, dbOpenSnapshot)
   With CTA_CTETMP
      Do While Not .EOF
        DEBEX$ = FORMATNUM$(XVALO(!DEBE), "F12.2")
        HABEX$ = FORMATNUM$(XVALO(!HABER), "F12.2")
        SALDOX$ = FORMATNUM$(XVALO(DEBEX$) - ACUENTA#, "F12.2")
        ENSE_PROGRAMADA# = XVALO(!ENSE_PROGRAMADA)
        EMERGENCIA# = XVALO(!EMERGENCIA)
        INGLES# = XVALO(!INGLES)
        COMPUTACION# = XVALO(!COMPUTACION)
        TIMOV$ = SAFETEXT$(!TipoMovim)
       
        Text5 = ENSE_PROGRAMADA#
        Text6 = EMERGENCIA#
        Text7 = INGLES#
        Text20 = COMPUTACION#
        .MoveNext
      Loop
   End With
   '
   Text19 = SALDOX$
   TXTTIPO = VARMOV$
   TXTMES = SAFETEXT$(MEX%)
   TXTANO = SAFETEXT$(ANOX%)
   TXTDCTO = DOCU$
   Text28 = TRIM$(FORMATNUM$(VALOR_CUOTA_ORIGINAL#, "F12.2"))
'   Call SUMARSUBTOTAL
   Call SUMARTOT
   
End Sub
Private Sub MSF_SelChange()
'   Text7 = TRIM$(MSF.TextMatrix(MSF.Row, 0))
   '
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   '
   I& = MSF.Row
   j& = MSF.COL
   '
   Call SELECCION_COLUMNA_0(MSF, I&, j&)
   
End Sub


Private Sub Text1_Change()
     Call LIMPIAR
     MSF.Rows = 1
     LGJO& = XVALO(Text1)
     If LGJO& < 1 Then Exit Sub
     
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
     SQLX$ = SQLX$ + " FROM CAJABANC_COL"
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
           If SALDO < 0.05 Then GoTo 50
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
     NUCOLU% = 3: TEPRUEBA = "12345678900000000000": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Documento"
     NUCOLU% = 4: TEPRUEBA = "ABCD": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Mes"
     NUCOLU% = 5: TEPRUEBA = "CHAPA": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Año"
     NUCOLU% = 6: TEPRUEBA = "123456789012": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = "Concepto"
     NUCOLU% = 7: TEPRUEBA = "1234567890": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Debe"
     NUCOLU% = 8: TEPRUEBA = "1234567890": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Haber"
     NUCOLU% = 9: TEPRUEBA = "1234567890": MSF.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 6: MSF.TextMatrix(0, NUCOLU% - 1) = "Saldo"
     MSF.Redraw = True
End Sub
   '
Sub CARGA_DATOS_ALUMNMOS()

   LEGAJO& = XVALO(Text1)
   If LEGAJO& <= 0 Then Exit Sub
   CONDI$ = "LEGAJO = " & LEGAJO& & "AND STATUS >= 0"
   
   If CantRegistros&("ALUMNOS", CONDI$, "NOMESS") < 1 Then Exit Sub

   SQLX$ = "SELECT * FROM ALUMNOS "
   SQLX$ = SQLX$ + " WHERE " + CONDI$
   Set ALUMTMP = BASSSEX.Execute(SQLX$)
   With ALUMTMP
        Text14 = SAFETEXT$(!NOMBRE_APELLIDO)
        Text2 = SAFETEXT$(!DOMCILIO)
        Text3 = SAFETEXT$(!localidad)
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
             If CONTROLX.Name <> "Text24" Then
              CONTROLX.Text = ""
             End If
           End If
        End If
Next
End Sub

Private Sub TXTBONIFICACION_Change()
    Call SUMARTOT
End Sub
Sub SUMARTOT()
   SUMATO# = (XVALO(Text19) + XVALO(TXTRECARGO)) - XVALO(TXTBONIFICACION)
   Text23 = FORMATNUM$(SUMATO#, "F12.2")
End Sub
'Sub SUMARSUBTOTAL()
'   SUMATO# = XVALO(Text5) + XVALO(Text6) + XVALO(Text7) + XVALO(Text20)
'   Text19 = FORMATNUM$(SUMATO#, "F12.2")
'End Sub
Private Sub TXTBONIFICACION_LostFocus()
   TXTBONIFICACION = FORMATNUM$(TXTBONIFICACION, "F12.2")
End Sub

Private Sub TXTRECARGO_Change()
   Call SUMARTOT
End Sub

Private Sub TXTRECARGO_LostFocus()
    TXTRECARGO = FORMATNUM$(TXTRECARGO, "F12.2")
End Sub
Sub LIMPIAR_FRAME_VALORES()
     Text5 = ""
     Text6 = ""
     Text7 = ""
     Text20 = ""
     Text19 = ""
     TXTBONIFICACION = ""
     TXTRECARGO = ""
     Text23 = ""
     Text28 = ""
     Text29 = ""
     TXTTIPO = ""
     TXTMES = ""
     TXTANO = ""
     TXTDCTO = ""
     
End Sub


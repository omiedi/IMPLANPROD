VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FTCAMBACERES 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FICHA TÉCNICA CAMBACERES"
   ClientHeight    =   10005
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   14670
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10005
   ScaleWidth      =   14670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TXTCODINT 
      Alignment       =   1  'Right Justify
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
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   154
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.TextBox TXTCODINTFTCAMBACERES 
      Alignment       =   1  'Right Justify
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
      Left            =   12840
      Locked          =   -1  'True
      TabIndex        =   153
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.TextBox TXTFTECNICACAMB_TXTVALMAXCINTAS 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3600
      MaxLength       =   2
      TabIndex        =   14
      Top             =   4845
      Width           =   345
   End
   Begin VB.TextBox TXTLABEL 
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
      Height          =   280
      Index           =   29
      Left            =   3960
      TabIndex        =   111
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1430
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
      Index           =   62
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   143
      TabStop         =   0   'False
      Text            =   "Peso rollo (kg)"
      Top             =   9120
      Width           =   1920
   End
   Begin VB.TextBox TXTFTECNICACAMB_PesoRolloKg 
      Alignment       =   2  'Center
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
      Height          =   330
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   142
      Top             =   9120
      Width           =   1005
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
      Index           =   61
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   141
      TabStop         =   0   'False
      Text            =   "Tensión arrollahilo"
      Top             =   8400
      Width           =   1920
   End
   Begin VB.TextBox TXTFTECNICACAMB_TensionarroHilo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   6240
      TabIndex        =   27
      Top             =   8400
      Width           =   1005
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
      Index           =   60
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   140
      TabStop         =   0   'False
      Text            =   "Peso cono vacío (kg)"
      Top             =   8760
      Width           =   1920
   End
   Begin VB.TextBox TXTFTECNICACAMB_PesConVacKg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   6240
      TabIndex        =   28
      Top             =   8760
      Width           =   1005
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
      Index           =   59
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   139
      TabStop         =   0   'False
      Text            =   "Punto frenos"
      Top             =   7680
      Width           =   1920
   End
   Begin VB.TextBox TXTFTECNICACAMB_PuntoFreno 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   6240
      TabIndex        =   25
      Top             =   7680
      Width           =   1005
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
      Index           =   58
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   138
      TabStop         =   0   'False
      Text            =   "Telares posibles"
      Top             =   8040
      Width           =   1920
   End
   Begin VB.TextBox TXTFTECNICACAMB_Telarespos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   6240
      TabIndex        =   26
      Top             =   8040
      Width           =   1005
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
      Index           =   57
      Left            =   12120
      Locked          =   -1  'True
      TabIndex        =   137
      TabStop         =   0   'False
      Text            =   "BOBINA (MTS)"
      Top             =   5280
      Width           =   1440
   End
   Begin VB.TextBox TXTFTECNICACAMB_BOBINASMT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   13560
      TabIndex        =   46
      Top             =   5280
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
      Index           =   56
      Left            =   12120
      Locked          =   -1  'True
      TabIndex        =   136
      TabStop         =   0   'False
      Text            =   "ANCHO (CM)"
      Top             =   4440
      Width           =   1440
   End
   Begin VB.TextBox TXTFTECNICACAMB_TolAnchoCm 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   13560
      TabIndex        =   44
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
      Index           =   55
      Left            =   12120
      Locked          =   -1  'True
      TabIndex        =   135
      TabStop         =   0   'False
      Text            =   "GRAMOS M2 "
      Top             =   4800
      Width           =   1440
   End
   Begin VB.TextBox TXTFTECNICACAMB_TolGramoM2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   13560
      TabIndex        =   45
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
      Index           =   54
      Left            =   12120
      Locked          =   -1  'True
      TabIndex        =   134
      TabStop         =   0   'False
      Text            =   "TOLERANCIA ( + - )"
      Top             =   4080
      Width           =   2440
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
      Index           =   53
      Left            =   12120
      Locked          =   -1  'True
      TabIndex        =   133
      TabStop         =   0   'False
      Text            =   "TRACCION (KG)"
      Top             =   3240
      Width           =   1440
   End
   Begin VB.TextBox TXTFTECNICACAMB_ResTransAnKG 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   13560
      TabIndex        =   42
      Top             =   3240
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
      Index           =   52
      Left            =   12120
      Locked          =   -1  'True
      TabIndex        =   132
      TabStop         =   0   'False
      Text            =   "% ELONGACION "
      Top             =   3600
      Width           =   1440
   End
   Begin VB.TextBox TXTFTECNICACAMB_ResPjeElong 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   13560
      TabIndex        =   43
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
      Index           =   51
      Left            =   12120
      Locked          =   -1  'True
      TabIndex        =   131
      TabStop         =   0   'False
      Text            =   "RESISTENCIA"
      Top             =   2880
      Width           =   2440
   End
   Begin VB.TextBox TXTFTECNICACAMB_PosicionCm10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   13890
      TabIndex        =   40
      Top             =   9120
      Width           =   700
   End
   Begin VB.TextBox TXTFTECNICACAMB_PosicionCm9 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   13185
      TabIndex        =   39
      Top             =   9120
      Width           =   700
   End
   Begin VB.TextBox TXTFTECNICACAMB_PosicionCm8 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   12480
      TabIndex        =   38
      Top             =   9120
      Width           =   700
   End
   Begin VB.TextBox TXTFTECNICACAMB_PosicionCm7 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   11835
      TabIndex        =   37
      Top             =   9120
      Width           =   700
   End
   Begin VB.TextBox TXTFTECNICACAMB_PosicionCm6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   11145
      TabIndex        =   36
      Top             =   9120
      Width           =   700
   End
   Begin VB.TextBox TXTFTECNICACAMB_PosicionCm5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   10440
      TabIndex        =   35
      Top             =   9120
      Width           =   700
   End
   Begin VB.TextBox TXTFTECNICACAMB_PosicionCm4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   9720
      TabIndex        =   34
      Top             =   9120
      Width           =   700
   End
   Begin VB.TextBox TXTFTECNICACAMB_PosicionCm3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   9075
      TabIndex        =   33
      Top             =   9120
      Width           =   700
   End
   Begin VB.TextBox TXTFTECNICACAMB_PosicionCm2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   8385
      TabIndex        =   32
      Top             =   9120
      Width           =   700
   End
   Begin VB.TextBox TXTFTECNICACAMB_PosicionCm1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   7680
      TabIndex        =   31
      Top             =   9120
      Width           =   700
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
      Index           =   50
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   130
      TabStop         =   0   'False
      Text            =   "POSICION (CM)"
      Top             =   8760
      Width           =   6920
   End
   Begin VB.TextBox TXTFTECNICACAMB_CintColCant 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   9120
      TabIndex        =   29
      Top             =   8040
      Width           =   1400
   End
   Begin VB.TextBox TXTFTECNICACAMB_CintColCol 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   9120
      TabIndex        =   30
      Top             =   8400
      Width           =   1400
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
      Index           =   49
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   129
      TabStop         =   0   'False
      Text            =   "Cantidad"
      Top             =   8040
      Width           =   1440
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
      Index           =   47
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   128
      TabStop         =   0   'False
      Text            =   "Color"
      Top             =   8400
      Width           =   1440
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
      Index           =   28
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   127
      TabStop         =   0   'False
      Text            =   "CINTAS DE COLOR"
      Top             =   7680
      Width           =   6920
   End
   Begin VB.TextBox TXTFTECNICACAMB_LubiriTime 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   1560
      TabIndex        =   17
      Top             =   7680
      Width           =   1000
   End
   Begin VB.TextBox TXTFTECNICACAMB_LubriUnMed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   2580
      TabIndex        =   21
      Top             =   7680
      Width           =   1400
   End
   Begin VB.TextBox TXTFTECNICACAMB_BreakHoraVal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   1560
      TabIndex        =   18
      Top             =   8040
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
      Height          =   1050
      Index           =   48
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   126
      TabStop         =   0   'False
      Text            =   "BREAK "
      Top             =   8040
      Width           =   1440
   End
   Begin VB.TextBox TXTFTECNICACAMB_BreakMinVal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   1560
      TabIndex        =   19
      Top             =   8400
      Width           =   1000
   End
   Begin VB.TextBox TXTFTECNICACAMB_BreakHoraUnMed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   2580
      TabIndex        =   22
      Top             =   8040
      Width           =   1400
   End
   Begin VB.TextBox TXTFTECNICACAMB_BreakMinUnMed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   2580
      TabIndex        =   23
      Top             =   8400
      Width           =   1400
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
      Index           =   46
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   125
      TabStop         =   0   'False
      Text            =   "LUBRICACIÓN"
      Top             =   7680
      Width           =   1440
   End
   Begin VB.TextBox TXTFTECNICACAMB_BreakSegUnMed 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   2580
      TabIndex        =   24
      Top             =   8760
      Width           =   1400
   End
   Begin VB.TextBox TXTFTECNICACAMB_BreakSegVal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   1560
      TabIndex        =   20
      Top             =   8760
      Width           =   1000
   End
   Begin VB.TextBox TXTLABEL 
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
      Height          =   280
      Index           =   45
      Left            =   3960
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   7260
      Width           =   1430
   End
   Begin VB.TextBox TXTLABEL 
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
      Height          =   280
      Index           =   44
      Left            =   3960
      TabIndex        =   123
      TabStop         =   0   'False
      Top             =   7005
      Width           =   1430
   End
   Begin VB.TextBox TXTLABEL 
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
      Height          =   280
      Index           =   43
      Left            =   3960
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   6735
      Width           =   1430
   End
   Begin VB.TextBox TXTLABEL 
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
      Height          =   280
      Index           =   42
      Left            =   3960
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1430
   End
   Begin VB.TextBox TXTLABEL 
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
      Height          =   280
      Index           =   41
      Left            =   3960
      TabIndex        =   120
      TabStop         =   0   'False
      Top             =   6190
      Width           =   1430
   End
   Begin VB.TextBox TXTFTECNICACAMB_CantCintAdicRxM 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   3960
      TabIndex        =   16
      Top             =   5925
      Width           =   1430
   End
   Begin VB.TextBox TXTLABEL 
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
      Height          =   280
      Index           =   39
      Left            =   3960
      TabIndex        =   119
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1430
   End
   Begin VB.TextBox TXTLABEL 
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
      Height          =   280
      Index           =   38
      Left            =   3960
      TabIndex        =   118
      TabStop         =   0   'False
      Top             =   5400
      Width           =   1430
   End
   Begin VB.TextBox TXTFTECNICACAMB_CantCintXMod 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   3960
      TabIndex        =   15
      Top             =   5100
      Width           =   1430
   End
   Begin VB.TextBox TXTLABEL 
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
      Height          =   280
      Index           =   36
      Left            =   3960
      TabIndex        =   117
      TabStop         =   0   'False
      Top             =   4845
      Width           =   1430
   End
   Begin VB.TextBox TXTLABEL 
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
      Height          =   280
      Index           =   35
      Left            =   3960
      TabIndex        =   116
      TabStop         =   0   'False
      Top             =   4575
      Width           =   1430
   End
   Begin VB.TextBox TXTLABEL 
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
      Height          =   280
      Index           =   34
      Left            =   3960
      TabIndex        =   115
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1430
   End
   Begin VB.TextBox TXTLABEL 
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
      Height          =   280
      Index           =   33
      Left            =   3960
      TabIndex        =   114
      TabStop         =   0   'False
      Top             =   4020
      Width           =   1430
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   27
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   110
      TabStop         =   0   'False
      Text            =   "% DE CONTRACCION REAL"
      Top             =   7260
      Width           =   3840
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   26
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   109
      TabStop         =   0   'False
      Text            =   "CANTIDAD DE BOBINAS TOTAL"
      Top             =   7005
      Width           =   3840
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   25
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   108
      TabStop         =   0   'False
      Text            =   "CANTIDAD DE BOBINAS X FILETA"
      Top             =   6735
      Width           =   3840
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   24
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   107
      TabStop         =   0   'False
      Text            =   "CANTIDAD DE CINTAS X MITAD DE MAQ."
      Top             =   6480
      Width           =   3840
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   23
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   106
      TabStop         =   0   'False
      Text            =   "MODULOS CON CINTAS ADICIONALES X MITAD"
      Top             =   6190
      Width           =   3840
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   22
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   105
      TabStop         =   0   'False
      Text            =   "CANT. DE CINTAS ADICIONALES REAL X MITAD"
      Top             =   5925
      Width           =   3840
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   21
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   104
      TabStop         =   0   'False
      Text            =   "CANT.DE CINTAS ADICIONALES X MITAD DE MAQ"
      Top             =   5655
      Width           =   3840
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   20
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   103
      TabStop         =   0   'False
      Text            =   "RESTO DE CINTAS A DISTRIBUIR"
      Top             =   5400
      Width           =   3840
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   19
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   102
      TabStop         =   0   'False
      Text            =   "CANT DE CINTAS X MODULO"
      Top             =   5100
      Width           =   3840
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   18
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   101
      TabStop         =   0   'False
      Text            =   "CANT. DE CINTAS X MODULO (MAX.36)     "
      Top             =   4845
      Width           =   3840
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   17
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   100
      TabStop         =   0   'False
      Text            =   "CANT. DE CINTAS X MITAD DE MAQ."
      Top             =   4575
      Width           =   3840
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   16
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   99
      TabStop         =   0   'False
      Text            =   "CANTIDAD DE CINTAS"
      Top             =   4320
      Width           =   3840
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   15
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   98
      TabStop         =   0   'False
      Text            =   "ANCHO CINTA (MM)"
      Top             =   4020
      Width           =   3840
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   14
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   97
      TabStop         =   0   'False
      Text            =   "DIMENSION REAL DE TELA (MM)"
      Top             =   3760
      Width           =   3840
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   13
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   96
      TabStop         =   0   'False
      Text            =   "% DE CONTRACCION"
      Top             =   3500
      Width           =   3840
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   95
      TabStop         =   0   'False
      Text            =   "ANCHO DE TELA (MM)"
      Top             =   3240
      Width           =   3840
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
      TabIndex        =   94
      TabStop         =   0   'False
      Text            =   "DISTRIBUCION DE URDIMBRES"
      Top             =   2880
      Width           =   5280
   End
   Begin VB.TextBox Text45 
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
      Height          =   330
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1000
   End
   Begin VB.TextBox Text42 
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
      Height          =   330
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   90
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1000
   End
   Begin VB.TextBox Text39 
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
      Height          =   330
      Left            =   11640
      Locked          =   -1  'True
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1000
   End
   Begin VB.TextBox Text33 
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
      Height          =   330
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1000
   End
   Begin VB.TextBox TXTFTECNICACAMB_PicsX_MIN 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   9720
      TabIndex        =   7
      Top             =   1800
      Width           =   1000
   End
   Begin VB.TextBox Text27 
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
      Height          =   330
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1000
   End
   Begin VB.TextBox Text25 
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
      Height          =   330
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1000
   End
   Begin VB.TextBox Text24 
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
      Height          =   330
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1000
   End
   Begin VB.TextBox Text23 
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
      Height          =   330
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1000
   End
   Begin VB.TextBox TXTFTECNICACAMB_RefDenier_GRMT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   6600
      TabIndex        =   13
      Top             =   2520
      Width           =   1250
   End
   Begin VB.TextBox TXTFTECNICACAMB_TramaDenier_GRMT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   6600
      TabIndex        =   6
      Top             =   1800
      Width           =   1250
   End
   Begin VB.TextBox TXTFTECNICACAMB_UrdDenier_GRMT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   6600
      TabIndex        =   10
      Top             =   2160
      Width           =   1250
   End
   Begin VB.TextBox Text19 
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
      Height          =   330
      Left            =   5320
      Locked          =   -1  'True
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1300
   End
   Begin VB.TextBox Text18 
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
      Height          =   330
      Left            =   5320
      Locked          =   -1  'True
      TabIndex        =   75
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1300
   End
   Begin VB.TextBox Text17 
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
      Height          =   330
      Left            =   5320
      Locked          =   -1  'True
      TabIndex        =   74
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1300
   End
   Begin VB.TextBox TXTFTECNICACAMB_RefCob_CmMM 
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
      Height          =   330
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1400
   End
   Begin VB.TextBox TXTFTECNICACAMB_TramaCob_CmMM 
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
      Height          =   330
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1400
   End
   Begin VB.TextBox TXTFTECNICACAMB_UrdCob_CmMM 
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
      Height          =   330
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1400
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
      TabIndex        =   69
      TabStop         =   0   'False
      Text            =   "REF."
      Top             =   2520
      Width           =   1440
   End
   Begin VB.TextBox TXTFTECNICACAMB_RefCta_Cm 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   1560
      TabIndex        =   11
      Top             =   2520
      Width           =   1000
   End
   Begin VB.TextBox TXTFTECNICACAMB_RefAnc_Cm 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   2580
      TabIndex        =   12
      Top             =   2520
      Width           =   1400
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
      TabIndex        =   66
      TabStop         =   0   'False
      Text            =   "X"
      Top             =   1440
      Width           =   1440
   End
   Begin VB.TextBox TXTFTECNICACAMB_UrdAnc_Cm 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   2580
      TabIndex        =   9
      Top             =   2160
      Width           =   1400
   End
   Begin VB.TextBox TXTFTECNICACAMB_TramaAnc_Cm 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   2580
      TabIndex        =   5
      Top             =   1800
      Width           =   1400
   End
   Begin VB.TextBox TXTFTECNICACAMB_UrdCta_Cm 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   1560
      TabIndex        =   8
      Top             =   2160
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
      Index           =   8
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   65
      TabStop         =   0   'False
      Text            =   "URD."
      Top             =   2160
      Width           =   1440
   End
   Begin VB.TextBox TXTFTECNICACAMB_TramaCta_Cm 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   1560
      TabIndex        =   4
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
      Index           =   7
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      Text            =   "TRAMA"
      Top             =   1800
      Width           =   1440
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Grabar"
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
      Left            =   13440
      TabIndex        =   63
      Top             =   9600
      Width           =   1155
   End
   Begin VB.Frame FRAME 
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   120
      TabIndex        =   48
      Top             =   0
      Width           =   14490
      Begin VB.TextBox TXTREFILEFACTOR 
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
         Left            =   12960
         Locked          =   -1  'True
         TabIndex        =   157
         TabStop         =   0   'False
         Top             =   960
         Width           =   435
      End
      Begin VB.TextBox TXTREFILEVALOR 
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
         Left            =   12120
         Locked          =   -1  'True
         TabIndex        =   156
         TabStop         =   0   'False
         Top             =   960
         Width           =   795
      End
      Begin VB.TextBox TXTREFILE 
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
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   155
         TabStop         =   0   'False
         Top             =   960
         Width           =   435
      End
      Begin VB.TextBox TXTFTECNICACAMB_GRAMAJE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   10200
         TabIndex        =   152
         Top             =   600
         Width           =   1155
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
         Index           =   40
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   151
         TabStop         =   0   'False
         Text            =   "GRAMAJE"
         Top             =   600
         Width           =   960
      End
      Begin VB.CommandButton Command6 
         Caption         =   "."
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
         Left            =   11400
         TabIndex        =   150
         Top             =   960
         Width           =   175
      End
      Begin VB.TextBox TXTFTECNICACAMB_TIPOTELA 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
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
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   149
         TabStop         =   0   'False
         ToolTipText     =   "Doble Click Para Limpiar Casillero"
         Top             =   960
         Width           =   1200
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
         Index           =   63
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   148
         TabStop         =   0   'False
         Text            =   "TIPO"
         Top             =   960
         Width           =   960
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
         Index           =   37
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   147
         TabStop         =   0   'False
         Text            =   "% Eficiencia"
         Top             =   600
         Width           =   1320
      End
      Begin VB.TextBox TXTFTECNICACAMB_TextEficiencia 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   12720
         TabIndex        =   41
         ToolTipText     =   "Ingresar la Capacidad de la Bolsa Kg. Ej. 50 Kg."
         Top             =   600
         Width           =   675
      End
      Begin VB.TextBox TXTGXMTSLINEAL 
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
         Height          =   330
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   960
         Width           =   1155
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
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   61
         TabStop         =   0   'False
         Text            =   "G x MTS LINEAL"
         Top             =   960
         Width           =   2640
      End
      Begin VB.TextBox TXTFTECNICACAMB_DISTDEDOYARO_MM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   8040
         TabIndex        =   2
         Top             =   600
         Width           =   1155
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
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Text            =   "DIST. ENTRE DEDO Y ARO (MM)"
         Top             =   600
         Width           =   2640
      End
      Begin VB.TextBox TXTFTECNICACAMB_LAMINADA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   4200
         TabIndex        =   3
         ToolTipText     =   "Ingresar la Capacidad de la Bolsa Kg. Ej. 50 Kg."
         Top             =   960
         Width           =   1155
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Text            =   "LAMINADA"
         Top             =   960
         Width           =   1440
      End
      Begin VB.TextBox TXTFTECNICACAMB_ARO_CM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   4200
         TabIndex        =   1
         ToolTipText     =   "Ingresar la Capacidad de la Bolsa Kg. Ej. 50 Kg."
         Top             =   600
         Width           =   1155
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Text            =   "ARO (CM)"
         Top             =   600
         Width           =   1440
      End
      Begin VB.TextBox TXTFTECNICACAMB_GRAMOS_M2 
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
         Height          =   330
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   960
         Width           =   1155
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
         TabIndex        =   56
         TabStop         =   0   'False
         Text            =   "GRAMOS M2"
         Top             =   960
         Width           =   1440
      End
      Begin VB.TextBox TXTFTECNICACAMB_ANCHO_CM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
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
         Left            =   1560
         TabIndex        =   0
         Top             =   600
         Width           =   1155
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
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   "Ancho (Cm)"
         Top             =   600
         Width           =   1440
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
         Left            =   2895
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   "Descripción"
         Top             =   240
         Width           =   1320
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
         Index           =   31
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Text            =   "Código"
         Top             =   240
         Width           =   840
      End
      Begin VB.TextBox TXTFTECNICACAMB_DESCRIP 
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
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   240
         Width           =   9180
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
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
         Left            =   2730
         TabIndex        =   52
         Top             =   240
         Width           =   175
      End
      Begin VB.TextBox TXTFTECNICACAMB_Codigo 
         Alignment       =   1  'Right Justify
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
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   240
         Width           =   1740
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   13680
         Picture         =   "FTCAMBACERES.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   240
         Width           =   650
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "FTCAMBACERES.frx":014A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "FTCAMBACERES.frx":037E
      TabIndex        =   49
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   4695
      Left            =   5520
      TabIndex        =   146
      Top             =   2880
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   8281
      _Version        =   393216
      Cols            =   4
      BackColor       =   14737632
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
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
   Begin VB.TextBox TXTLABEL 
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
      Height          =   280
      Index           =   30
      Left            =   3960
      TabIndex        =   112
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1430
   End
   Begin VB.TextBox TXTLABEL 
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
      Height          =   280
      Index           =   32
      Left            =   3960
      TabIndex        =   113
      TabStop         =   0   'False
      Top             =   3765
      Width           =   1430
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
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
      Left            =   0
      TabIndex        =   145
      Top             =   0
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   144
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "KG X DIA  "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   13560
      TabIndex        =   93
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MTS X DIA "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   12600
      TabIndex        =   91
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MTS X DIA AL 100%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   11640
      TabIndex        =   89
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MTS X MIN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   10680
      TabIndex        =   87
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PICS X MIN (PPM)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   9720
      TabIndex        =   85
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GR/M2 TOTAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   8760
      TabIndex        =   84
      Top             =   1440
      Width           =   1005
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GR/M2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   7800
      TabIndex        =   81
      Top             =   1440
      Width           =   1000
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DENIER GR/9000 MTS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   6600
      TabIndex        =   78
      Top             =   1440
      Width           =   1250
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "% DE CONTRACCION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   5320
      TabIndex        =   76
      Top             =   1440
      Width           =   1300
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "COBERTURA X CM (MM)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3960
      TabIndex        =   72
      Top             =   1440
      Width           =   1395
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ANCHO DE CINTA (CM)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   2580
      TabIndex        =   68
      Top             =   1440
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CINTAS X CM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1560
      TabIndex        =   67
      Top             =   1440
      Width           =   1000
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
   End
   Begin VB.Menu mnuCL 
      Caption         =   "Consultas y Listados"
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
   End
   Begin VB.Menu mnuUt 
      Caption         =   "Utilidades"
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuraciones"
      Begin VB.Menu mnuTablaAux 
         Caption         =   "Tablas Auxiliares"
      End
      Begin VB.Menu NivelAprob 
         Caption         =   "Niveles de Aprobación"
      End
      Begin VB.Menu mnuConfigWaldbot 
         Caption         =   "Configuración General"
      End
   End
   Begin VB.Menu mnuAccesosDirectos 
      Caption         =   "Accesos Directos"
   End
End
Attribute VB_Name = "FTCAMBACERES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Public MODIFIC%
    Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long
    Dim WinRar As String, WinZip As String
    Dim BRTEMP As Recordset
    Public NOPRESENTARMENSAJE%

    Dim ActualizaListaSeleccion As Boolean

Sub CARGARDATOSFICHATECNICA()
   
   SQLX$ = "SELECT * FROM FTECNICACAMB WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE CODINT = " & XVALO(Me.TXTCODINT)
   Set RS = READSET(SQLX$)
   
   With RS
'            TXTFTECNICACAMB_Codigo.TEXT = SAFETEXT$(!Codigo)
'            TXTFTECNICACAMB_DESCRIP.TEXT = SAFETEXT$(!Descrip)
            Me.TXTFTECNICACAMB_ARO_CM.TEXT = TRIM$(FORMATNUM$(XVALO(!ARO_CM), "F12.2"))
            TXTFTECNICACAMB_LAMINADA.TEXT = SAFETEXT$(!LAMINADA)
            TXTFTECNICACAMB_UrdDenier_GRMT.TEXT = TRIM$(FORMATNUM$(XVALO(!UrdDenier_GRMT), "F12.2"))
            TXTFTECNICACAMB_ANCHO_CM.TEXT = TRIM$(FORMATNUM$(XVALO(!ANCHO_CM), "F12.2"))
            TXTFTECNICACAMB_GRAMOS_M2.TEXT = TRIM$(FORMATNUM$(XVALO(!GRAMOS_M2), "F12.2"))
            TXTFTECNICACAMB_PicsX_MIN.TEXT = TRIM$(FORMATNUM$(XVALO(!PicsX_MIN), "F12.2"))
            TXTFTECNICACAMB_DISTDEDOYARO_MM.TEXT = TRIM$(FORMATNUM$(XVALO(!DISTDEDOYARO_MM), "F12.2"))
            TXTFTECNICACAMB_TramaCta_Cm.TEXT = TRIM$(FORMATNUM$(XVALO(!TramaCta_Cm), "F12.2"))
            TXTFTECNICACAMB_TramaAnc_Cm.TEXT = TRIM$(FORMATNUM$(XVALO(!TramaAnc_Cm), "F12.2"))
            TXTFTECNICACAMB_TramaDenier_GRMT.TEXT = TRIM$(FORMATNUM$(XVALO(!TramaDenier_GRMT), "F12.2"))
            TXTFTECNICACAMB_UrdCta_Cm.TEXT = TRIM$(FORMATNUM$(XVALO(!UrdCta_Cm), "F12.2"))
            TXTFTECNICACAMB_UrdAnc_Cm.TEXT = TRIM$(FORMATNUM$(XVALO(!UrdAnc_Cm), "F12.2"))
            
            TXTFTECNICACAMB_UrdCob_CmMM.TEXT = TRIM$(FORMATNUM$(XVALO(!UrdCob_CmMM), "F12.2"))
            TXTFTECNICACAMB_RefCob_CmMM.TEXT = TRIM$(FORMATNUM$(XVALO(!RefCob_CmMM), "F12.2"))
            TXTFTECNICACAMB_RefCta_Cm.TEXT = TRIM$(FORMATNUM$(XVALO(!RefCta_Cm), "F12.2"))
            TXTFTECNICACAMB_RefAnc_Cm.TEXT = TRIM$(FORMATNUM$(XVALO(!RefAnc_Cm), "F12.2"))
            TXTFTECNICACAMB_RefDenier_GRMT.TEXT = TRIM$(FORMATNUM$(XVALO(!RefDenier_GRMT), "F12.2"))
            TXTFTECNICACAMB_CantCintXMod.TEXT = TRIM$(FORMATNUM$(XVALO(!CantCintXMod), "F12.2"))
            TXTFTECNICACAMB_CantCintAdicRxM.TEXT = TRIM$(FORMATNUM$(XVALO(!CantCintAdicRxM), "F12.2"))
            TXTFTECNICACAMB_CantCintAdicRxM.TEXT = TRIM$(FORMATNUM$(XVALO(!CantCintAdicRxM), "F12.2"))
            TXTFTECNICACAMB_CintColCant.TEXT = TRIM$(FORMATNUM$(XVALO(!CintColCant), "F12.2"))
            
            TXTFTECNICACAMB_CintColCol.TEXT = TRIM$(FORMATNUM$(XVALO(!CintColCol), "F12.2"))
            TXTFTECNICACAMB_PosicionCm1.TEXT = TRIM$(FORMATNUM$(XVALO(!PosicionCm1), "F12.2"))
            TXTFTECNICACAMB_PosicionCm2.TEXT = TRIM$(FORMATNUM$(XVALO(!PosicionCm2), "F12.2"))
            
            TXTFTECNICACAMB_PosicionCm3.TEXT = TRIM$(FORMATNUM$(XVALO(!PosicionCm3), "F12.2"))
            TXTFTECNICACAMB_PosicionCm4.TEXT = TRIM$(FORMATNUM$(XVALO(!PosicionCm4), "F12.2"))
            TXTFTECNICACAMB_PosicionCm5.TEXT = TRIM$(FORMATNUM$(XVALO(!PosicionCm5), "F12.2"))
            TXTFTECNICACAMB_PosicionCm6.TEXT = TRIM$(FORMATNUM$(XVALO(!PosicionCm6), "F12.2"))
            TXTFTECNICACAMB_PosicionCm7.TEXT = TRIM$(FORMATNUM$(XVALO(!PosicionCm7), "F12.2"))
            TXTFTECNICACAMB_PosicionCm8.TEXT = TRIM$(FORMATNUM$(XVALO(!PosicionCm8), "F12.2"))
            TXTFTECNICACAMB_PosicionCm9.TEXT = TRIM$(FORMATNUM$(XVALO(!PosicionCm9), "F12.2"))
            TXTFTECNICACAMB_PosicionCm10.TEXT = TRIM$(FORMATNUM$(XVALO(!PosicionCm10), "F12.2"))
            
            TXTFTECNICACAMB_ResTransAnKG.TEXT = TRIM$(FORMATNUM$(XVALO(!ResTransAnKG), "F12.2"))
            TXTFTECNICACAMB_ResPjeElong.TEXT = TRIM$(FORMATNUM$(XVALO(!ResPjeElong), "F12.2"))
            TXTFTECNICACAMB_TolAnchoCm.TEXT = TRIM$(FORMATNUM$(XVALO(!TolAnchoCm), "F12.2"))
            TXTFTECNICACAMB_TolGramoM2.TEXT = TRIM$(FORMATNUM$(XVALO(!TolGramoM2), "F12.2"))
            TXTFTECNICACAMB_BOBINASMT.TEXT = TRIM$(FORMATNUM$(XVALO(!BOBINASMT), "F12.2"))
            TXTFTECNICACAMB_PuntoFreno.TEXT = TRIM$(FORMATNUM$(XVALO(!PuntoFreno), "F12.2"))
            TXTFTECNICACAMB_Telarespos.TEXT = TRIM$(FORMATNUM$(XVALO(!Telarespos), "F12.2"))
            TXTFTECNICACAMB_TensionarroHilo.TEXT = TRIM$(FORMATNUM$(XVALO(!TensionarroHilo), "F12.2"))
            TXTFTECNICACAMB_PesConVacKg.TEXT = TRIM$(FORMATNUM$(XVALO(!PesConVacKg), "F12.2"))
            TXTFTECNICACAMB_PesoRolloKg.TEXT = TRIM$(FORMATNUM$(XVALO(!PesoRolloKg), "F12.2"))

            
            TXTFTECNICACAMB_LubiriTime.TEXT = TRIM$(FORMATNUM$(XVALO(!LubiriTime), "F12.2"))
            TXTFTECNICACAMB_LubriUnMed.TEXT = TRIM$(!LubriUnMed)
            TXTFTECNICACAMB_BreakHoraVal.TEXT = TRIM$(FORMATNUM$(XVALO(!BreakHoraVal), "F12.2"))
            TXTFTECNICACAMB_BreakHoraUnMed.TEXT = TRIM$(!BreakHoraUnMed)
            TXTFTECNICACAMB_BreakMinVal.TEXT = TRIM$(FORMATNUM$(XVALO(!BreakMinVal), "F12.2"))
            
            TXTFTECNICACAMB_BreakMinUnMed.TEXT = TRIM$(!BreakMinUnMed)
            TXTFTECNICACAMB_BreakSegVal.TEXT = TRIM$(FORMATNUM$(XVALO(!BreakSegVal), "F12.2"))
            TXTFTECNICACAMB_BreakSegUnMed.TEXT = TRIM$(!BreakSegUnMed)
            
            TXTFTECNICACAMB_TramaCob_CmMM.TEXT = TRIM$(FORMATNUM$(XVALO(!TramaCob_CmMM), "F12.2"))
            
            TXTFTECNICACAMB_TXTVALMAXCINTAS.TEXT = TRIM$(FORMATNUM$(XVALO(!TramaCob_ValMxCta), "F6"))

            TXTFTECNICACAMB_TextEficiencia.TEXT = TRIM$(FORMATNUM$(XVALO(!TramaCob_Eficien), "F12.2"))
            
            TXTFTECNICACAMB_TIPOTELA.TEXT = TRIM$(!TramaCob_TIPOTELA)
            TXTFTECNICACAMB_GRAMAJE.TEXT = TRIM$(FORMATNUM$(XVALO(!TramaCob_GRAMAJE), "F12.2"))

    
   End With
   
 

End Sub

Sub GRABARGRILLAMODULOS()
    CODI$ = TRIM$(TXTFTECNICACAMB_Codigo)
    DESK$ = TRIM$(TXTFTECNICACAMB_DESCRIP)
    
    SQLX$ = "DELETE FROM FTMODULO WHERE CODINT = " & XVALO(TXTCODINT)
    FLEXCONN.Execute (SQLX$)
    
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    SQLX$ = "SELECT  * FROM FTMODULO WHERE CODINT < 1 "
    RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With RS
      For i& = 1 To MSF1.Rows - 1
         MODULO = Int(XVALO(MSF1.TextMatrix(i&, 1)))
         CANTI = Int(XVALO(MSF1.TextMatrix(i&, 2)))
         REFUERZO = Int(XVALO(MSF1.TextMatrix(i&, 3)))
         .AddNew
         !CODINT = XVALO(TXTCODINT)
         !Codigo = CODI$
         !Descrip = DESK$
         !MODULO = MODULO
         !Cantidad = CANTI
         !REFUERZOS = REFUERZO
         !MarBorra = 0
         .Update
       Next i&
     End With
     RS.Close
     Set RS = Nothing
       
End Sub

Sub MOSTRARGRILLAMODULOS(CI1%)

    i& = 0
    SQLX$ = "SELECT  * FROM FTMODULO WITH(NOLOCK) WHERE CODINT=  " & CI1%
    SQLX$ = SQLX$ + " ORDER BY MODULO"
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
         i& = i& + 1
         MSF1.TextMatrix(i&, 1) = XVALO(!MODULO)
         MSF1.TextMatrix(i&, 2) = XVALO(!Cantidad)
         MSF1.TextMatrix(i&, 3) = XVALO(!REFUERZOS)
         .MoveNext
      Loop
       
    End With
    RS.Close
    Set RS = Nothing
       
End Sub
Sub REFRESCARDATOSXFORMULA()
On Error GoTo ERROR_GUARDAR
Call LIMPIARTEXTSEGUNCOLOR(Me, 1, &HE0E0E0)
K2 = ROUND(XVALO(TXTFTECNICACAMB_GRAMAJE), 2)
K3 = TRIM$(TXTFTECNICACAMB_TIPOTELA)

efic = ROUND(XVALO(TXTFTECNICACAMB_TextEficiencia), 2)
E1 = ROUND(XVALO(TXTFTECNICACAMB_ANCHO_CM), 2)
G1 = ROUND(XVALO(TXTFTECNICACAMB_ARO_CM), 2)
J1 = ROUND(XVALO(TXTFTECNICACAMB_DISTDEDOYARO_MM), 2)
G2 = Me.TXTFTECNICACAMB_LAMINADA
B7 = ROUND(XVALO(TXTFTECNICACAMB_TramaCta_Cm), 2)
C7 = ROUND(XVALO(TXTFTECNICACAMB_TramaAnc_Cm), 2)
F7 = ROUND(XVALO(TXTFTECNICACAMB_TramaDenier_GRMT), 2)
I7 = ROUND(XVALO(TXTFTECNICACAMB_PicsX_MIN), 2)
B8 = ROUND(XVALO(TXTFTECNICACAMB_UrdCta_Cm), 2)
c8 = ROUND(XVALO(TXTFTECNICACAMB_UrdAnc_Cm), 2)
F8 = ROUND(XVALO(TXTFTECNICACAMB_UrdDenier_GRMT), 2)

B9 = ROUND(XVALO(TXTFTECNICACAMB_RefCta_Cm), 2)
C9 = ROUND(XVALO(TXTFTECNICACAMB_RefAnc_Cm), 2)
F9 = ROUND(XVALO(TXTFTECNICACAMB_RefDenier_GRMT), 2)
C18 = ROUND(XVALO(TXTFTECNICACAMB_TXTVALMAXCINTAS), 2)
D19 = ROUND(XVALO(TXTFTECNICACAMB_CantCintXMod), 2)
D22 = ROUND(XVALO(TXTFTECNICACAMB_CantCintAdicRxM), 2)
K28 = ROUND(XVALO(TXTFTECNICACAMB_BOBINASMT), 2)
K35 = ROUND(XVALO(TXTFTECNICACAMB_PesConVacKg), 2)


E2 = ROUND(XVALO(TXTFTECNICACAMB_GRAMOS_M2), 2)
L1 = ROUND(XVALO(TXTGXMTSLINEAL.TEXT), 2)
D7 = ROUND(XVALO(TXTFTECNICACAMB_TramaCob_CmMM), 2)
D8 = ROUND(XVALO(TXTFTECNICACAMB_UrdCob_CmMM), 2)
D9 = ROUND(XVALO(TXTFTECNICACAMB_RefCob_CmMM), 2)
E7 = ROUND(XVALO(Text18), 2)
E8 = ROUND(XVALO(Text17), 2)
E9 = ROUND(XVALO(Text19), 2)
G7 = ROUND(XVALO(Text24), 2)
G8 = ROUND(XVALO(Text23), 2)
G9 = ROUND(XVALO(Text25), 2)
H7 = ROUND(XVALO(Text27), 2)
J7 = ROUND(XVALO(Text33), 2)
K7 = ROUND(XVALO(Text39), 2)
L7 = ROUND(XVALO(Text42), 2)
M7 = ROUND(XVALO(Text45), 2)

D12 = ROUND(XVALO(TXTLABEL(29)), 2)
D13 = ROUND(XVALO(TXTLABEL(30)), 2)
D14 = ROUND(XVALO(TXTLABEL(32)), 2)
D15 = ROUND(XVALO(TXTLABEL(33)), 2)
D16 = ROUND(XVALO(TXTLABEL(34)), 2)
D17 = ROUND(XVALO(TXTLABEL(35)), 2)
D18 = ROUND(XVALO(TXTLABEL(36)), 2)
D20 = ROUND(XVALO(TXTLABEL(38)), 2)
D21 = ROUND(XVALO(TXTLABEL(39)), 2)
D23 = ROUND(XVALO(TXTLABEL(41)), 2)
D24 = ROUND(XVALO(TXTLABEL(42)), 2)
D25 = ROUND(XVALO(TXTLABEL(43)), 2)
D26 = ROUND(XVALO(TXTLABEL(44)), 2)
D26 = ROUND(XVALO(TXTLABEL(45)), 2)
'D28 = XVALO(TXTFTECNICACAMB_ANCHO_CM)
'D29 = XVALO(TXTFTECNICACAMB_ANCHO_CM)

'TXTFTECNICACAMB_GRAMOS_M2 = FORMATNUM$(H7, "F12.2")
'If H7 > 0 Then
'  TXTGXMTSLINEAL = FORMATNUM$(XVALO(((E1 / 100) * 2) * (190 * ((H7 + G9) / H7))), "F12.2")
'Else
'  TXTGXMTSLINEAL = 0
'End If
E2 = ROUND(XVALO(TXTFTECNICACAMB_GRAMOS_M2), 2)

TXTFTECNICACAMB_TramaCob_CmMM = FORMATNUM$(C7 * B7, "F12.2")
D7 = ROUND(XVALO(TXTFTECNICACAMB_TramaCob_CmMM), 2)

TXTFTECNICACAMB_UrdCob_CmMM = FORMATNUM$(c8 * B8, "F12.2")
D8 = ROUND(XVALO(TXTFTECNICACAMB_UrdCob_CmMM), 2)

TXTFTECNICACAMB_RefCob_CmMM = FORMATNUM$(C9 * B9, "F12.2")
D9 = ROUND(XVALO(TXTFTECNICACAMB_RefCob_CmMM), 2)


Text18 = FORMATNUM$((((D7 * 100) / 10) - 100), "F12.2")  'ORIGINAL =(((D7*100)/10)-100)/100 LE SACO EL / 100 POR QUE ES PORCENTAJE, EN EL EXCEL REPRESENTA BIEN POR QUE TIENE FORMATO %
E7 = ROUND(XVALO(Text18), 2)

Text17 = FORMATNUM$((((D8 * 100) / 10) - 100), "F12.2")
E8 = ROUND(XVALO(Text17), 2)
Text19 = FORMATNUM$((((D9 * 100) / 10) - 100), "F12.2")
E9 = ROUND(XVALO(Text19), 2)

If XVALO(C7) > 0 Then
  Text24 = FORMATNUM$(((D7 / 10) / (C7 / 1000)) * F7 / 9000, "F12.2")
Else
  Text24 = ""
End If
G7 = ROUND(XVALO(Text24), 2)

If XVALO(c8) > 0 Then
  Text23 = FORMATNUM$(((D8 / 10) / (c8 / 1000)) * F8 / 9000, "F12.2")
Else
  Text23 = ""
End If
G8 = ROUND(XVALO(Text23), 2)
If D9 > 0 And C9 > 0 Then
  Text25 = FORMATNUM$(((D9 / 10) / (C9 / 1000)) * F9 / 9000, "F12.2")
Else
  Text25 = 0
End If
G9 = ROUND(XVALO(Text25), 2)

Text27 = FORMATNUM$(G7 + G8, "F12.2")
H7 = ROUND(XVALO(Text27), 2)

TXTFTECNICACAMB_GRAMOS_M2 = FORMATNUM$(H7, "F12.2")
'=SI(B9=0,SI(K3="TUBULAR",E1/100*2*K2,E1/100*K2),((E1/100)*2)*(160*((H7+G9)/H7)))
If B9 = 0 Then  'cambias por si es tubular o no
    If K3 = "TUBULAR" Then
       TXTGXMTSLINEAL = E1 / 100 * 2 * K2
    Else
       TXTGXMTSLINEAL = E1 / 100 * K2
    End If
Else
     TXTGXMTSLINEAL = ((E1 / 100) * 2) * (160 * ((H7 + G9) / H7))
End If
L1 = ROUND(XVALO(TXTGXMTSLINEAL.TEXT), 2)

'If H7 > 0 Then
'  TXTGXMTSLINEAL = FORMATNUM$(XVALO(((E1 / 100) * 2) * (190 * ((H7 + G9) / H7))), "F12.2")
'Else
'  TXTGXMTSLINEAL = 0
'End If
'L1 = ROUND(XVALO(TXTGXMTSLINEAL.TEXT), 2)

If XVALO(B7) > 0 Then
   Text33 = FORMATNUM$((I7 / B7) / 100, "F12.2") 'MTS X MIN
Else
   Text33 = ""
End If
J7 = ROUND(XVALO(Text33), 2)
Text39 = FORMATNUM$((J7 * 60) * 24, "F12.2") 'MTS X DIA AL 100%
K7 = ROUND(XVALO(Text39), 2)

If efic > 0 Then efic = efic / 100
Text42 = FORMATNUM$(K7 * efic, "F12.2") 'MTS X DIA AL 80%
L7 = ROUND(XVALO(Text42), 2)

If L7 > 0 Then
   Text45 = FORMATNUM$(L1 / 1000 * L7, "F12.2") 'KG X DIA AL 80%
Else
   Text45 = 0
End If
M7 = ROUND(XVALO(Text45), 2)

'=SI(K3="TUBULAR",E1*2*10,E1*10)
If K3 = "TUBULAR" Then
   TXTLABEL(29) = FORMATNUM$(E1 * 2 * 10, "F12.2") 'ANCHO DE TELA (MM)
Else
   TXTLABEL(29) = FORMATNUM$(E1 * 10, "F12.2") 'ANCHO DE TELA (MM)
End If
D12 = ROUND(XVALO(TXTLABEL(29)), 2)

TXTLABEL(30) = FORMATNUM$(E8, "F12.2") '% DE CONTRACCION
D13 = ROUND(XVALO(TXTLABEL(30)), 2)

TXTLABEL(32) = FORMATNUM$(D12 + D12 * (D13 / 100), "F12.2") 'DIMENSION REAL DE TELA (MM
D14 = ROUND(XVALO(TXTLABEL(32)), 2)

TXTLABEL(33) = FORMATNUM$(c8, "F12.2") 'ANCHO CINTA (MM)
D15 = XVALO(TXTLABEL(33))
If D15 > 0 Then
  TXTLABEL(34) = FORMATNUM$(D14 / D15, "F12.2") 'CANTIDAD DE CINTAS
Else
  TXTLABEL(34) = 0
End If
D16 = ROUND(XVALO(TXTLABEL(34)), 2)

TXTLABEL(35) = FORMATNUM$(D16 / 2, "F12.2") 'CANT. DE CINTAS X MITAD DE MAQ.
D17 = ROUND(XVALO(TXTLABEL(35)), 2)

If XVALO(TXTFTECNICACAMB_TXTVALMAXCINTAS) > 0 Then
   TXTLABEL(36) = FORMATNUM$(D17 / XVALO(TXTFTECNICACAMB_TXTVALMAXCINTAS), "F12.2") 'CANT. DE CINTAS X MODULO (MAX. 36)
Else                                                                'AHORA TIENE CASILLERO PARA INGRESAR VALOR
   TXTLABEL(36) = 0
End If
D18 = ROUND(XVALO(TXTLABEL(36)), 2)

TXTLABEL(38) = FORMATNUM$(D18 - D19, "F12.2") 'RESTO DE CINTAS A DISTRIBUIR
D20 = ROUND(XVALO(TXTLABEL(38)), 2)

TXTLABEL(39) = FORMATNUM$(D20 * C18, "F12.2") 'CANT.DE CINTAS ADICIONALES X MITAD DE MAQ
D21 = ROUND(XVALO(TXTLABEL(39)), 2)

TXTLABEL(41) = FORMATNUM$(D22 / 2, "F12.2") 'MODULOS CON CINTAS ADICIONALES X MITAD
D23 = ROUND(XVALO(TXTLABEL(41)), 2)

TXTLABEL(42) = FORMATNUM$((D19 * C18) + D22, "F12.2") 'CANTIDAD DE CINTAS X MITAD DE MAQ.
D24 = ROUND(XVALO(TXTLABEL(42)), 2)


TXTLABEL(43) = FORMATNUM$((D24 / 2), "F12.2") 'CANTIDAD DE BOBINAS X FILETA
D25 = ROUND(XVALO(TXTLABEL(43)), 2)

TXTLABEL(44) = FORMATNUM$((D24 * 2), "F12.2") 'CANTIDAD DE BOBINAS TOTAL
D26 = ROUND(XVALO(TXTLABEL(44)), 2)

If c8 > 0 Then
   'originalmente (((D26 * 100) / (D12 / c8) - 100) / 100 se le quito el / 100 para que muestre el porcentaje real
   'con / 100 muestra 1.04 sin / 100  es 104.4 asi queda como en el excel
   TXTLABEL(45) = FORMATNUM$(((D26 * 100) / (D12 / c8) - 100), "F12.2")   '% DE CONTRACCION REAL
Else
   TXTLABEL(45) = 0
End If
D26 = ROUND(XVALO(TXTLABEL(45)), 2)

If L1 > 0 Then
  TXTFTECNICACAMB_PesoRolloKg = FORMATNUM$((L1 / 1000 * K28) + K35, "F12.2")
Else
  TXTFTECNICACAMB_PesoRolloKg = 0
  
End If
'=(L1/1000*K28)+K35
ERROR_GUARDAR:

End Sub

Private Sub Command1_Click()
     COD$ = ListaSeleccionArticulos("", ActualizaListaSeleccion)
     ActualizaListaSeleccion = False
     If COD$ <> "" Then
       TXTFTECNICACAMB_Codigo.TEXT = COD$
     End If

End Sub

Private Sub Command13_Click()
  Command13.Enabled = False
'VALIDO PRIMERO QUE TENGA UN CODIGO ACTIVO DEL MASTER
   If CODINT(TXTFTECNICACAMB_Codigo) < 1 Then
      Call COMUNI("Falta Código de Producto o Código Inválido")
      GoTo 99
   End If
    
   'AHORA VALIDO SI ESTA DADO DE ALTA EN LA TABLA FTECNICACAMB
   'VUELVO A VER SI NO TIENE CODIGO INTERNO POR SI SE DIO DE ALTA EN OTRA PC MISMO MOMENTO
   CONDI$ = "Codigo =  '" & TRIM$(Me.TXTFTECNICACAMB_Codigo) & "'"
   Me.TXTCODINTFTCAMBACERES = XVALO(VALOBADA("FTECNICACAMB", "CODINT", CONDI$, "NOMESS"))

   CI1% = XVALO(Me.TXTCODINTFTCAMBACERES.TEXT)
   If CI1% > 0 Then
      ESNUEVO% = 0
   Else
      ESNUEVO% = 1
   End If

   SQLX$ = "SELECT * FROM FTECNICACAMB "
   If ESNUEVO% < 1 Then
     SQLX$ = SQLX$ + " WHERE CODINT = " & XVALO(Me.TXTCODINT)
   Else
     SQLX$ = SQLX$ + " WHERE CODINT < 1 "
   End If
   
   On Error GoTo ERROR_GUARDAR
   FLEXCONN.BeginTrans
   
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset
   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With RS
     If ESNUEVO% > 0 Then
         CODIGOINT = Me.TXTCODINT
         .AddNew
         !CODINT = CODIGOINT
     End If
     
            !Codigo = TXTFTECNICACAMB_Codigo.TEXT
            !Descrip = TXTFTECNICACAMB_DESCRIP.TEXT

            !ARO_CM = FormatNumber(XVALO(Me.TXTFTECNICACAMB_ARO_CM.TEXT), 3)

            !ANCHO_CM = FormatNumber(XVALO(TXTFTECNICACAMB_ANCHO_CM.TEXT), 3)

            !GRAMOS_M2 = FormatNumber(XVALO(TXTFTECNICACAMB_GRAMOS_M2.TEXT), 3)

            !PicsX_MIN = FormatNumber(XVALO(TXTFTECNICACAMB_PicsX_MIN.TEXT), 3)

            !LAMINADA = TXTFTECNICACAMB_LAMINADA.TEXT

            !DISTDEDOYARO_MM = FormatNumber(XVALO(TXTFTECNICACAMB_DISTDEDOYARO_MM.TEXT), 3)

            !TramaCta_Cm = FormatNumber(XVALO(TXTFTECNICACAMB_TramaCta_Cm.TEXT), 3)

            !TramaAnc_Cm = FormatNumber(XVALO(TXTFTECNICACAMB_TramaAnc_Cm.TEXT), 3)

            !TramaDenier_GRMT = FormatNumber(XVALO(TXTFTECNICACAMB_TramaDenier_GRMT.TEXT), 3)

            !UrdCta_Cm = FormatNumber(XVALO(TXTFTECNICACAMB_UrdCta_Cm.TEXT), 3)

            !UrdAnc_Cm = FormatNumber(XVALO(TXTFTECNICACAMB_UrdAnc_Cm.TEXT), 3)

            !UrdDenier_GRMT = FormatNumber(XVALO(TXTFTECNICACAMB_UrdDenier_GRMT.TEXT), 3)

            !UrdCob_CmMM = FormatNumber(XVALO(TXTFTECNICACAMB_UrdCob_CmMM.TEXT), 3)

            !RefCob_CmMM = FormatNumber(XVALO(TXTFTECNICACAMB_RefCob_CmMM.TEXT), 3)

            !RefCta_Cm = FormatNumber(XVALO(TXTFTECNICACAMB_RefCta_Cm.TEXT), 3)

            !RefAnc_Cm = FormatNumber(XVALO(TXTFTECNICACAMB_RefAnc_Cm.TEXT), 3)

            !RefDenier_GRMT = FormatNumber(XVALO(TXTFTECNICACAMB_RefDenier_GRMT.TEXT), 3)

            !CantCintXMod = FormatNumber(XVALO(TXTFTECNICACAMB_CantCintXMod.TEXT), 3)

            !CantCintAdicRxM = FormatNumber(XVALO(TXTFTECNICACAMB_CantCintAdicRxM.TEXT), 3)

            !CintColCant = FormatNumber(XVALO(TXTFTECNICACAMB_CintColCant.TEXT), 3)

            !CintColCol = FormatNumber(XVALO(TXTFTECNICACAMB_CintColCol.TEXT), 3)

            !PosicionCm1 = FormatNumber(XVALO(TXTFTECNICACAMB_PosicionCm1.TEXT), 3)

            !PosicionCm2 = FormatNumber(XVALO(TXTFTECNICACAMB_PosicionCm2.TEXT), 3)

            !PosicionCm3 = FormatNumber(XVALO(TXTFTECNICACAMB_PosicionCm3.TEXT), 3)

            !PosicionCm4 = FormatNumber(XVALO(TXTFTECNICACAMB_PosicionCm4.TEXT), 3)

            !PosicionCm5 = FormatNumber(XVALO(TXTFTECNICACAMB_PosicionCm5.TEXT), 3)

            !PosicionCm6 = FormatNumber(XVALO(TXTFTECNICACAMB_PosicionCm6.TEXT), 3)

            !PosicionCm7 = FormatNumber(XVALO(TXTFTECNICACAMB_PosicionCm7.TEXT), 3)

            !PosicionCm8 = FormatNumber(XVALO(TXTFTECNICACAMB_PosicionCm8.TEXT), 3)

            !PosicionCm9 = FormatNumber(XVALO(TXTFTECNICACAMB_PosicionCm9.TEXT), 3)

            !PosicionCm10 = FormatNumber(XVALO(TXTFTECNICACAMB_PosicionCm10.TEXT), 3)

            !ResTransAnKG = FormatNumber(XVALO(TXTFTECNICACAMB_ResTransAnKG.TEXT), 3)

            !ResPjeElong = FormatNumber(XVALO(TXTFTECNICACAMB_ResPjeElong.TEXT), 3)

            !TolAnchoCm = FormatNumber(XVALO(TXTFTECNICACAMB_TolAnchoCm.TEXT), 3)

            !TolGramoM2 = FormatNumber(XVALO(TXTFTECNICACAMB_TolGramoM2.TEXT), 3)

            !BOBINASMT = FormatNumber(XVALO(TXTFTECNICACAMB_BOBINASMT.TEXT), 3)

            !PuntoFreno = FormatNumber(XVALO(TXTFTECNICACAMB_PuntoFreno.TEXT), 3)

            !Telarespos = FormatNumber(XVALO(TXTFTECNICACAMB_Telarespos.TEXT), 3)

            !TensionarroHilo = FormatNumber(XVALO(TXTFTECNICACAMB_TensionarroHilo.TEXT), 3)

            !PesConVacKg = FormatNumber(XVALO(TXTFTECNICACAMB_PesConVacKg.TEXT), 3)

            !PesoRolloKg = FormatNumber(XVALO(TXTFTECNICACAMB_PesoRolloKg.TEXT), 3)

            !LubiriTime = FormatNumber(XVALO(TXTFTECNICACAMB_LubiriTime.TEXT), 3)

            !LubriUnMed = TXTFTECNICACAMB_LubriUnMed.TEXT

            !BreakHoraVal = FormatNumber(XVALO(TXTFTECNICACAMB_BreakHoraVal.TEXT), 3)

            !BreakHoraUnMed = TXTFTECNICACAMB_BreakHoraUnMed.TEXT

            !BreakMinVal = FormatNumber(XVALO(TXTFTECNICACAMB_BreakMinVal.TEXT), 3)

            !BreakMinUnMed = TXTFTECNICACAMB_BreakMinUnMed.TEXT

            !BreakSegVal = FormatNumber(XVALO(TXTFTECNICACAMB_BreakSegVal.TEXT), 3)

            !BreakSegUnMed = TXTFTECNICACAMB_BreakSegUnMed.TEXT

            !TramaCob_CmMM = FormatNumber(XVALO(TXTFTECNICACAMB_TramaCob_CmMM.TEXT), 3)
            
            !TramaCob_ValMxCta = FormatNumber(XVALO(TXTFTECNICACAMB_TXTVALMAXCINTAS.TEXT), 3)
            
            !TramaCob_Eficien = FormatNumber(XVALO(TXTFTECNICACAMB_TextEficiencia), 3)
            
            !TramaCob_GRAMAJE = FormatNumber(XVALO(TXTFTECNICACAMB_GRAMAJE), 3)
            
            
            !TramaCob_TIPOTELA = TXTFTECNICACAMB_TIPOTELA.TEXT
            
            !MarBorra = 0
    
            .Update
   End With
   Call GRABARGRILLAMODULOS
   TEXT5 = NEWCOTIZ
    
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
       GoTo 99
    Else
      FLEXCONN.CommitTrans
      If NOPRESENTARMENSAJE% < 1 Then
         Call COMUNI("Transacción Realizada")
      End If
    End If
    
99  Command13.Enabled = True

End Sub


Private Sub Command2_Click()
      Call CIERRARCH("*.*")
      Call BAJALDISCO
      Call FINAL("")
End Sub

Private Sub Command6_Click()
    TituloLista$ = "Tabla de Tipo Telas y Refiles"
    Campos$ = "Codigo/A0;Tipo Tela/A18;Refiles/A7;Valor/F12.1;Factor/F6.0"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim obj As New RECORXET
    Set rs1 = obj.RS_TIPOTELAYREFILE()
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 10000
50  FRMZELECHO.Show 1

    TIPOTELA$ = TRIM(RESP_ZELE_VECT(2))
    TXTFTECNICACAMB_TIPOTELA = TIPOTELA$
'    REFILE$ = TRIM(RESP_ZELE_VECT(3))
'    VALORX$ = TRIM(RESP_ZELE_VECT(4))
'    FACTORX$ = TRIM(RESP_ZELE_VECT(5))




End Sub

Private Sub Form_Load()
    If (App.PrevInstance = True) Then
        MsgBox "Yá se Encuentra Abierto el Módulo de Ficha Técnica Cambaceres"
        Call FINAL("")
    End If
    
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    
    Campos$ = "MODULO/F8;CANT./F16;REFUERZOS/F18;ID/A0"
    Call CARGA_ENCABEZADO(MSF1, Campos$, Me)
    MSF1.Rows = 31
    Call numerargrilla(MSF1, 1)
    Call ABRECONEXION
    Call CREAR_TABLA_FTECNICACAMB
    Call GENERAR_ESTRUCTURAS_SQL
    Call MIGRADAT_A_SQL
    
End Sub

Private Sub mnuConfigWaldbot_Click()
   USER1% = USNBR% Mod 100
   If USER1% <> 1 Then
      Call COMUNI("Solo Puede Ingresar el Supervisor del Sistema")
      Exit Sub
   End If

   BkpWaldb.Show 1
   
End Sub

Private Sub mnuTablaAux_Click()
  Call GENERAR_ESTRUCTURAS_SQL
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set rs1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(rs1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.MSF2)
  Set obj = Nothing
  rs1.Close
  Set rs1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  '

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  IDLIST& = XVALO(RESP_ZELE_VECT(1))

  Call ABM_TABLA(IDLIST&)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
      Call CIERRARCH("*.*")
      Call BAJALDISCO
      Call FINAL("")
End Sub

Private Sub MSF1_DblClick()
    Unload VENTABNEW
    Campos$ = "MODULO/F8;CANT./F16;REFUERZOS/F18;ID/A0"

    VENTANA.Inicializar = 0
    VENTANA.Campox = Campos$
    VENTANA.UtilizaFormula = 1
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "Ventana de Edición"
    VENTANA.CampEditables = "2;3"
    VENTANA.Inicializar = 1
    Call CopiarMsfAOtroMsf(MSF1, VENTABNEW.MSF, 1)
    '****************************************************************
    Call T_Espera(1 / 10) ' PARA QUE EL DOBLE CLICK NO APLIQUE SOBRE LA VENTANA QUE SE CARGA
    
    VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
       Call CopiarMsfAOtroMsf(VENTABNEW.MSF, MSF1, 1)
    End If
    Set VENTANA = Nothing
    Unload VENTABNEW
End Sub



Private Sub NivelAprob_Click()
   USER1% = USNBR% Mod 100
   If USER1% <> 1 Then
      Call COMUNI("Solo Puede Ingresar el Supervisor del Sistema")
      Exit Sub
   End If
   
   Call Mostrar_Datusers
End Sub

Private Sub TXTFTECNICACAMB_GRAMAJE_Change()
   Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_TextEficiencia_Change()
   Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_TextEficiencia_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub

Private Sub TXTCODINT_Change()
   CONDI$ = "CODINT = " & XVALO(TXTCODINT)
   Me.TXTCODINTFTCAMBACERES = XVALO(VALOBADA("FTECNICACAMB", "CODINT", CONDI$, "NOMESS"))
   Me.TXTFTECNICACAMB_DESCRIP = VALOBADA("MASTER", "Descripcion", CONDI$, "NOMESS")
End Sub

Private Sub TXTCODINTFTCAMBACERES_Change()
'   Call BLANFORMU
   CI1% = XVALO(Me.TXTCODINTFTCAMBACERES)
   If CI1% > 0 Then
      Call CARGARDATOSFICHATECNICA
      Call MOSTRARGRILLAMODULOS(CI1%)
      Call REFRESCARDATOSXFORMULA
   End If
End Sub

Private Sub TXTFTECNICACAMB_ANCHO_CM_Change()
   Call REFRESCARDATOSXFORMULA
   
End Sub

Private Sub TXTFTECNICACAMB_ANCHO_CM_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)
End Sub


Private Sub TXTFTECNICACAMB_ARO_CM_Change()
Call REFRESCARDATOSXFORMULA
End Sub


Private Sub TXTFTECNICACAMB_ARO_CM_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_BOBINASMT_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_BOBINASMT_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_BreakHoraUnMed_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_BreakHoraVal_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_BreakHoraVal_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_BreakMinUnMed_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_BreakMinVal_Change()
Call REFRESCARDATOSXFORMULA
End Sub


Private Sub TXTFTECNICACAMB_BreakMinVal_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_BreakSegUnMed_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_BreakSegVal_Change()
Call REFRESCARDATOSXFORMULA
End Sub


Private Sub TXTFTECNICACAMB_BreakSegVal_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_CantCintAdicRxM_Change()
Call REFRESCARDATOSXFORMULA

End Sub

Private Sub TXTFTECNICACAMB_CantCintAdicRxM_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_CantCintXMod_Change()
Call REFRESCARDATOSXFORMULA

End Sub

Private Sub TXTFTECNICACAMB_CantCintXMod_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_CintColCant_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_CintColCant_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_CintColCol_Change()
Call REFRESCARDATOSXFORMULA
End Sub


Private Sub TXTFTECNICACAMB_Codigo_Change()
   CI1% = CODINT%(TXTFTECNICACAMB_Codigo.TEXT)
   Call BLANFORMU
   
   TXTCODINT = CI1%
End Sub

Sub BLANFORMU()

            Campos$ = "MODULO/F8;CANT./F16;REFUERZOS/F18;ID/A0"
            Call CARGA_ENCABEZADO(MSF1, Campos$, Me)
            MSF1.Rows = 31
            Call numerargrilla(MSF1, 1)
            
'            TXTFTECNICACAMB_DESCRIP.TEXT = ""

            TXTFTECNICACAMB_ARO_CM.TEXT = ""

            TXTFTECNICACAMB_ANCHO_CM.TEXT = ""

            TXTFTECNICACAMB_GRAMOS_M2.TEXT = ""

            TXTFTECNICACAMB_PicsX_MIN.TEXT = ""

            TXTFTECNICACAMB_LAMINADA.TEXT = ""

            TXTFTECNICACAMB_DISTDEDOYARO_MM.TEXT = ""

            TXTFTECNICACAMB_TramaCta_Cm.TEXT = ""

            TXTFTECNICACAMB_TramaAnc_Cm.TEXT = ""

            TXTFTECNICACAMB_TramaDenier_GRMT.TEXT = ""

            TXTFTECNICACAMB_UrdCta_Cm.TEXT = ""

            TXTFTECNICACAMB_UrdAnc_Cm.TEXT = ""

            TXTFTECNICACAMB_UrdDenier_GRMT.TEXT = ""

            TXTFTECNICACAMB_UrdCob_CmMM.TEXT = ""

            TXTFTECNICACAMB_RefCob_CmMM.TEXT = ""

            TXTFTECNICACAMB_RefCta_Cm.TEXT = ""

            TXTFTECNICACAMB_RefAnc_Cm.TEXT = ""

            TXTFTECNICACAMB_RefDenier_GRMT.TEXT = ""

            TXTFTECNICACAMB_CantCintXMod.TEXT = ""

            TXTFTECNICACAMB_CantCintAdicRxM.TEXT = ""

            TXTFTECNICACAMB_CintColCant.TEXT = ""

            TXTFTECNICACAMB_CintColCol.TEXT = ""

            TXTFTECNICACAMB_PosicionCm1.TEXT = ""

            TXTFTECNICACAMB_PosicionCm2.TEXT = ""

            TXTFTECNICACAMB_PosicionCm3.TEXT = ""

            TXTFTECNICACAMB_PosicionCm4.TEXT = ""

            TXTFTECNICACAMB_PosicionCm5.TEXT = ""

            TXTFTECNICACAMB_PosicionCm6.TEXT = ""

            TXTFTECNICACAMB_PosicionCm7.TEXT = ""

            TXTFTECNICACAMB_PosicionCm8.TEXT = ""

            TXTFTECNICACAMB_PosicionCm9.TEXT = ""

            TXTFTECNICACAMB_PosicionCm10.TEXT = ""

            TXTFTECNICACAMB_ResTransAnKG.TEXT = ""

            TXTFTECNICACAMB_ResPjeElong.TEXT = ""

            TXTFTECNICACAMB_TolAnchoCm.TEXT = ""

            TXTFTECNICACAMB_TolGramoM2.TEXT = ""

            TXTFTECNICACAMB_BOBINASMT.TEXT = ""

            TXTFTECNICACAMB_PuntoFreno.TEXT = ""

            TXTFTECNICACAMB_Telarespos.TEXT = ""

            TXTFTECNICACAMB_TensionarroHilo.TEXT = ""

            TXTFTECNICACAMB_PesConVacKg.TEXT = ""

            TXTFTECNICACAMB_PesoRolloKg.TEXT = ""

            TXTFTECNICACAMB_LubiriTime.TEXT = ""

            TXTFTECNICACAMB_LubriUnMed.TEXT = ""

            TXTFTECNICACAMB_BreakHoraVal.TEXT = ""

            TXTFTECNICACAMB_BreakHoraUnMed.TEXT = ""

            TXTFTECNICACAMB_BreakMinVal.TEXT = ""

            TXTFTECNICACAMB_BreakMinUnMed.TEXT = ""

            TXTFTECNICACAMB_BreakSegVal.TEXT = ""

            TXTFTECNICACAMB_BreakSegUnMed.TEXT = ""

            TXTFTECNICACAMB_TramaCob_CmMM.TEXT = ""
            
            TXTFTECNICACAMB_TextEficiencia = ""
            TXTFTECNICACAMB_TXTVALMAXCINTAS = ""
            TXTFTECNICACAMB_TIPOTELA = ""
            
            Call LIMPIARTEXTSEGUNCOLOR(Me, 1, &HE0E0E0)

End Sub

Private Sub TXTFTECNICACAMB_DISTDEDOYARO_MM_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_DISTDEDOYARO_MM_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_LAMINADA_Change()
Call REFRESCARDATOSXFORMULA
End Sub


Private Sub TXTFTECNICACAMB_LubiriTime_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_LubiriTime_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_LubriUnMed_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_PesConVacKg_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_PesConVacKg_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_PicsX_MIN_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_PicsX_MIN_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_PosicionCm1_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_PosicionCm1_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_PosicionCm10_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_PosicionCm10_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_PosicionCm2_Change()
Call REFRESCARDATOSXFORMULA
End Sub


Private Sub TXTFTECNICACAMB_PosicionCm2_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_PosicionCm3_Change()
Call REFRESCARDATOSXFORMULA
End Sub


Private Sub TXTFTECNICACAMB_PosicionCm3_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_PosicionCm4_Change()
Call REFRESCARDATOSXFORMULA
End Sub


Private Sub TXTFTECNICACAMB_PosicionCm4_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_PosicionCm5_Change()
Call REFRESCARDATOSXFORMULA
End Sub


Private Sub TXTFTECNICACAMB_PosicionCm5_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_PosicionCm6_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_PosicionCm7_Change()
Call REFRESCARDATOSXFORMULA
End Sub


Private Sub TXTFTECNICACAMB_PosicionCm7_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_PosicionCm8_Change()
Call REFRESCARDATOSXFORMULA
End Sub


Private Sub TXTFTECNICACAMB_PosicionCm8_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_PosicionCm9_Change()
Call REFRESCARDATOSXFORMULA
End Sub


Private Sub TXTFTECNICACAMB_PosicionCm9_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_PuntoFreno_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_PuntoFreno_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_RefAnc_Cm_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_RefAnc_Cm_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_RefCta_Cm_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_RefCta_Cm_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_RefDenier_GRMT_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_RefDenier_GRMT_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_ResPjeElong_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_ResPjeElong_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_ResTransAnKG_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_ResTransAnKG_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_Telarespos_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_TensionarroHilo_Change()
Call REFRESCARDATOSXFORMULA
End Sub


Private Sub TXTFTECNICACAMB_TensionarroHilo_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_TIPOTELA_Change()
   CONDI$ = "CAMPO2 = '" & TXTFTECNICACAMB_TIPOTELA & "' AND LISTA = 'TIPOTELACMB'"
   Me.TXTREFILE = VALOBADA("DATASQL", "CAMPO3", CONDI$, "NOMESS")
   Me.TXTREFILEVALOR = FORMATNUM$(XVALO(VALOBADA("DATASQL", "CAMPO4", CONDI$, "NOMESS")), "F10.1")
   Me.TXTREFILEFACTOR = XVALO(VALOBADA("DATASQL", "CAMPO5", CONDI$, "NOMESS"))

   
   
   Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_TolAnchoCm_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_TolAnchoCm_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_TolGramoM2_Change()
Call REFRESCARDATOSXFORMULA
End Sub


Private Sub TXTFTECNICACAMB_TolGramoM2_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_TramaAnc_Cm_Change()
Call REFRESCARDATOSXFORMULA
End Sub


Private Sub TXTFTECNICACAMB_TramaAnc_Cm_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_TramaCta_Cm_Change()
Call REFRESCARDATOSXFORMULA
End Sub


Private Sub TXTFTECNICACAMB_TramaCta_Cm_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_TramaDenier_GRMT_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_TramaDenier_GRMT_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub




Private Sub TXTFTECNICACAMB_UrdAnc_Cm_Change()
Call REFRESCARDATOSXFORMULA
End Sub

Private Sub TXTFTECNICACAMB_UrdAnc_Cm_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_UrdCta_Cm_Change()
Call REFRESCARDATOSXFORMULA
End Sub


Private Sub TXTFTECNICACAMB_UrdCta_Cm_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_UrdDenier_GRMT_Change()
Call REFRESCARDATOSXFORMULA
End Sub


Private Sub TXTFTECNICACAMB_UrdDenier_GRMT_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub


Private Sub TXTFTECNICACAMB_TXTVALMAXCINTAS_Change()
Call REFRESCARDATOSXFORMULA
End Sub


Private Sub TXTFTECNICACAMB_TXTVALMAXCINTAS_KeyPress(KeyAscii As Integer)
   KeyAscii = Solonumeros(KeyAscii)

End Sub
Sub MIGRADAT_A_SQL()

    Call CREA_TABLA_USUARIOSSQL
    Call ACTUALIZARFLEXUSR
    '''''''''''''''''''''''EL IDENTIFICADOR Y LA CLAVE PARA ASIGNARLA MAS ABAJO
    FUSER$ = LOADFILE$("USER.DAT")
    Dim IDENTIFICADOR1$()
    Dim CLAVEUSERDAT1$()
    For II% = 1 To Len(FUSER$) Step 64
       JJ% = JJ% + 1
       F0$ = Mid$(FUSER$, II%, 64)
       CLAV$ = Mid$(F0$, 43, 12): If Left$(CLAV$, 1) = "@" Then CLAV$ = DESENCRI$(Mid$(F0$, 44, 11)) + " "
       ReDim Preserve CLAVEUSERDAT1$(JJ%), IDENTIFICADOR1$(JJ%)
       CLAVEUSERDAT1$(JJ%) = TRIM$(CLAV$)
       Ident$ = Mid$(F0$, 31, 12)
       If Left$(Ident$, 1) = "@" Then Ident$ = DESENCRI$(Mid$(Ident$, 2, 11)) + " "
       IDENTIFICADOR1$(JJ%) = TRIM$(Ident$)
    Next II%
    '
    'tomo el nivel de configuracion y lo cargo en un vector para ganar velocidad ya que la funcion demora mucho
    Dim NICON1%()
    For i& = 1 To ULTREG&("DATUSERS")
       ReDim Preserve NICON1%(i&)
       X$ = REGLEIDO$("DATUSERS", i&)
       NICON1%(i&) = Asc(Mid$(X$, 32, 1))
    Next i&
    '
    'LEO EL IDMAQUINA QUE SE GUARDA EN LA TABLA SQL
    SQLX$ = "SELECT IDMAQUINA FROM USUARIOSSQL WITH(NOLOCK)"
    Set rs1 = READSET(SQLX$)
    Dim CODIMAQUS$()
    L% = 0
    With rs1
      Do While Not .EOF
        L% = L% + 1
        ReDim Preserve CODIMAQUS$(L%)
        CODIMAQUS$(L%) = SAFETEXT$(!IDMAQUINA)
        .MoveNext
      Loop
    End With
    ''''''''''''''''''''''''''''''''' HASTA AQUI
    '
    Call PoneEnCeroNull("USUARIOSSQL", "MARBORRA", REGAF&, "NOMESS")
    Call ACTUREGISTRO("USUARIOSSQL", "MARBORRA", "MARBORRA < 1", 1, REGAF&, "NOMESS")
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
25  On Error Resume Next
    SQLX$ = "SELECT * FROM USUARIOSSQL WHERE NUSER < 1 "
    RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    FE% = HOYSERVER%
    For i& = 1 To ULTREG&("FLEXUSR")
       X$ = REGLEIDO$("FLEXUSR", i&)
       NUS% = Asc(Mid$(X$, 1, 1))
       DENOM$ = Mid$(X$, 2, 29)
       NICON% = NICON1%(NUS%)
       RS.AddNew
       RS!FECHA = CVDAT(FE%)
       RS!NUSER = NUS%
       RS!CLAVE = CLAVEUSERDAT1$(i&) 'CLAVEUSERDAT$(NUS%, RET_IDENTIFICADOR$)
       RS!nombre = TRIM$(DENOM$)
       RS!IDENTIFICADOR = IDENTIFICADOR1$(i&) 'RET_IDENTIFICADOR$ ' VALOR DE RETORNO
       RS!NIVELCONF = NICON%
       RS!IDMAQUINA = CODIMAQUS$(NUS%) ' VECTOR CARGADO MAS ARRIBA
       RS!MarBorra = 0
       RS.Update
    Next i&
    RS.Close
    Set RS = Nothing
    Call BORRAREGISTROS("USUARIOSSQL", "MARBORRA = 1", REGAF&, "NOMESS")
End Sub





VERSION 5.00
Begin VB.Form FRMCOTI2 
   Caption         =   "Cotización (2)"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   14985
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   14985
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CMDDATOSAUXILIARES 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Contiene Cargado Datos Auxiliares"
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
      Left            =   6360
      TabIndex        =   125
      Top             =   5880
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Frame Frame10 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8160
      Index           =   4
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      Begin VB.TextBox TXTPORCENT2 
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
         Index           =   4
         Left            =   11500
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   6550
         Width           =   765
      End
      Begin VB.TextBox TXTLABEL2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   136
         TabStop         =   0   'False
         Text            =   "COMISION REAL"
         Top             =   6550
         Width           =   1815
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   132
         TabStop         =   0   'False
         Text            =   "COSTOS FIJOS %"
         Top             =   4080
         Width           =   2550
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   12255
         Locked          =   -1  'True
         TabIndex        =   135
         TabStop         =   0   'False
         Text            =   "U$S"
         Top             =   4080
         Width           =   460
      End
      Begin VB.TextBox TXTLABEL 
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
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   134
         TabStop         =   0   'False
         ToolTipText     =   "% Costo de Mano de Obra contra Costo Total"
         Top             =   4080
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox TXTCOSTOSFIJOS 
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
         Left            =   12720
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1815
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Index           =   2
         Left            =   9720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   131
         TabStop         =   0   'False
         Text            =   "FRMCOTI2.frx":0000
         Top             =   6960
         Visible         =   0   'False
         Width           =   4920
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   130
         TabStop         =   0   'False
         Text            =   "u$S"
         Top             =   15840
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox TXTVALORLINER 
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
         Left            =   12720
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   7440
         Width           =   1815
      End
      Begin VB.TextBox TXTLABEL2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   128
         TabStop         =   0   'False
         Text            =   "LINER"
         Top             =   7440
         Width           =   2550
      End
      Begin VB.TextBox TXTLABEL2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   9960
         Locked          =   -1  'True
         TabIndex        =   127
         TabStop         =   0   'False
         Text            =   "PRECIO VENTA REAL + LINER"
         Top             =   7752
         Width           =   2550
      End
      Begin VB.TextBox TXTPRECIOFINAL 
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   7752
         Width           =   1815
      End
      Begin VB.TextBox TXTPORCENT2 
         Alignment       =   2  'Center
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
         Left            =   12260
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   5010
         Width           =   460
      End
      Begin VB.TextBox TXTPORCENT2 
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
         Index           =   1
         Left            =   12260
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   5325
         Width           =   460
      End
      Begin VB.TextBox TXTCODESPVPD 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7635
         Locked          =   -1  'True
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   7800
         Width           =   2050
      End
      Begin VB.TextBox Text3 
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
         Left            =   2520
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   6555
         Width           =   5130
      End
      Begin VB.TextBox TXTESPECIFCLIENTE_2 
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
         Left            =   2520
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   7800
         Width           =   5130
      End
      Begin VB.TextBox TXTLABEL2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BeginProperty Font 
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
         Left            =   7635
         Locked          =   -1  'True
         TabIndex        =   122
         TabStop         =   0   'False
         Text            =   "Según Requerimiento"
         Top             =   6555
         Width           =   2055
      End
      Begin VB.TextBox TXTLABEL2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BeginProperty Font 
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
         Left            =   7635
         Locked          =   -1  'True
         TabIndex        =   121
         TabStop         =   0   'False
         Text            =   "30 Días Fecha de Factura"
         Top             =   6855
         Width           =   2050
      End
      Begin VB.TextBox TXTLABEL2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BeginProperty Font 
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
         Left            =   7635
         Locked          =   -1  'True
         TabIndex        =   120
         TabStop         =   0   'False
         Text            =   "15 Días"
         Top             =   7170
         Width           =   2050
      End
      Begin VB.TextBox TXTLABEL2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BeginProperty Font 
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
         Left            =   7635
         Locked          =   -1  'True
         TabIndex        =   119
         TabStop         =   0   'False
         Text            =   "Capital Federal o GBA"
         Top             =   7485
         Width           =   2050
      End
      Begin VB.TextBox TXTROTULOGASTOS 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   118
         TabStop         =   0   'False
         Text            =   "Código de Especif. Interna"
         Top             =   7800
         Width           =   2400
      End
      Begin VB.TextBox TXTROTULOGASTOS 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   116
         TabStop         =   0   'False
         Text            =   "Plazo de Entrega:"
         Top             =   6555
         Width           =   2400
      End
      Begin VB.TextBox TXTMANTENOFERTA 
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
         Left            =   2520
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   7170
         Width           =   5130
      End
      Begin VB.TextBox TXTROTULOGASTOS 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   114
         TabStop         =   0   'False
         Text            =   "Mantenimiento de la Oferta:"
         Top             =   7170
         Width           =   2400
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   113
         TabStop         =   0   'False
         Text            =   "CANTIDAD SOLICITADA"
         Top             =   5400
         Width           =   1920
      End
      Begin VB.TextBox TXTCANTSOLIC 
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
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   112
         TabStop         =   0   'False
         ToolTipText     =   "Doble Clic Para Editar"
         Top             =   5400
         Width           =   960
      End
      Begin VB.TextBox TXTPORCENT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   11160
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   5640
         Width           =   1100
      End
      Begin VB.TextBox TXTCOMISION 
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   5640
         Width           =   1815
      End
      Begin VB.TextBox TXTROTULOGASTOS 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   5640
         Width           =   1440
      End
      Begin VB.ListBox List2 
         Height          =   255
         ItemData        =   "FRMCOTI2.frx":0085
         Left            =   5400
         List            =   "FRMCOTI2.frx":0087
         TabIndex        =   107
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TXTROTULOGASTOS 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   106
         TabStop         =   0   'False
         Text            =   "OBSERVACIONES"
         Top             =   4920
         Width           =   1785
      End
      Begin VB.TextBox Text1 
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
         Height          =   1035
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   5280
         Width           =   5760
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   8880
         TabIndex        =   104
         ToolTipText     =   "Salir del Formulario Activo"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   107
         Left            =   9840
         Locked          =   -1  'True
         TabIndex        =   103
         TabStop         =   0   'False
         Text            =   "N°"
         Top             =   250
         Width           =   315
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   10140
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   102
         Top             =   250
         Width           =   765
      End
      Begin VB.CommandButton Command45 
         Appearance      =   0  'Flat
         Caption         =   "A.Sug."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4680
         TabIndex        =   15
         ToolTipText     =   "ACEPTA SUGERIDO"
         Top             =   4485
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   116
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   101
         TabStop         =   0   'False
         Text            =   "Otros  %"
         Top             =   3400
         Width           =   1800
      End
      Begin VB.TextBox TXTFEENTRECOTI 
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
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   4485
         Width           =   975
      End
      Begin VB.CommandButton Command43 
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
         Height          =   315
         Left            =   9015
         TabIndex        =   99
         Top             =   4485
         Width           =   175
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   102
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   98
         TabStop         =   0   'False
         Text            =   "FECHA DE ENTREGA"
         Top             =   4485
         Width           =   1920
      End
      Begin VB.TextBox TXTROTULOGASTOS 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   97
         TabStop         =   0   'False
         Text            =   "Lugar de Entrega:"
         Top             =   7485
         Width           =   2400
      End
      Begin VB.TextBox TXTROTULOGASTOS 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   96
         TabStop         =   0   'False
         Text            =   "Condiciones de Pago:"
         Top             =   6855
         Width           =   2400
      End
      Begin VB.TextBox TXTUSS2 
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
         Index           =   3
         Left            =   12720
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   6240
         Width           =   1815
      End
      Begin VB.TextBox TXTUSS2 
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   94
         TabStop         =   0   'False
         ToolTipText     =   "COSTO TOTAL + MARGEN DE UTILIDAD"
         Top             =   5925
         Width           =   1815
      End
      Begin VB.TextBox TXTUSS2 
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   93
         TabStop         =   0   'False
         ToolTipText     =   "COSTO TOTAL * PORCENTAJE DE MARGEN DE UTILIDAD"
         Top             =   5325
         Width           =   1815
      End
      Begin VB.TextBox TXTUSS2 
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   92
         TabStop         =   0   'False
         ToolTipText     =   "COSTO INDIRECTO DEL BOLSON + COSTO DIR.BOLSON"
         Top             =   5010
         Width           =   1815
      End
      Begin VB.TextBox TXTLABEL2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   87
         TabStop         =   0   'False
         Text            =   "PRECIO DE VENTA REAL"
         Top             =   6240
         Width           =   2550
      End
      Begin VB.TextBox TXTLABEL2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   85
         TabStop         =   0   'False
         Text            =   "MARGEN UTILIDAD"
         Top             =   5325
         Width           =   2540
      End
      Begin VB.TextBox TXTLABEL2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   84
         TabStop         =   0   'False
         Text            =   "COSTO TOTAL"
         Top             =   5010
         Width           =   2540
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   101
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   83
         TabStop         =   0   'False
         Text            =   "%"
         Top             =   3705
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   100
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   82
         TabStop         =   0   'False
         Text            =   "u$S"
         Top             =   3705
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox TXTSCRAP 
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
         Left            =   12720
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   4695
         Width           =   1815
      End
      Begin VB.TextBox TXTROTULOGASTOS 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   4695
         Width           =   1080
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   98
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   79
         TabStop         =   0   'False
         Text            =   "U$S"
         Top             =   4485
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   108
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   78
         TabStop         =   0   'False
         Text            =   "U$S Sugerido"
         Top             =   4485
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.TextBox TXTUSSUGERIDO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   77
         TabStop         =   0   'False
         Text            =   "U$S"
         ToolTipText     =   "Costo de Materiales segun % de Scrapp"
         Top             =   4695
         Width           =   1080
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   106
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   76
         TabStop         =   0   'False
         Text            =   "%"
         Top             =   4485
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox TXTPORCENT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   10800
         Locked          =   -1  'True
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   4695
         Width           =   840
      End
      Begin VB.TextBox Text13 
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
         Index           =   7
         Left            =   12720
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   3090
         Width           =   1800
      End
      Begin VB.TextBox Text13 
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
         Index           =   6
         Left            =   12720
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   2775
         Width           =   1800
      End
      Begin VB.TextBox Text13 
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
         Index           =   5
         Left            =   12720
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   2460
         Width           =   1800
      End
      Begin VB.TextBox Text13 
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   71
         TabStop         =   0   'False
         ToolTipText     =   "C6 / C1"
         Top             =   2145
         Width           =   1800
      End
      Begin VB.TextBox Text12 
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
         Index           =   4
         Left            =   10920
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   2145
         Width           =   1800
      End
      Begin VB.TextBox Text11 
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
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2145
         Width           =   1800
      End
      Begin VB.TextBox Text10 
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
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   2145
         Width           =   1800
      End
      Begin VB.TextBox Text9 
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   67
         TabStop         =   0   'False
         ToolTipText     =   "Cantidad Según Configuración Big Bag"
         Top             =   2145
         Width           =   1800
      End
      Begin VB.TextBox Text13 
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   66
         TabStop         =   0   'False
         ToolTipText     =   "C5 * C6"
         Top             =   1830
         Width           =   1800
      End
      Begin VB.TextBox Text12 
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
         Index           =   3
         Left            =   10920
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1830
         Width           =   1800
      End
      Begin VB.TextBox Text11 
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
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         ToolTipText     =   "MULTIPLICA LOS 4 CASILLEROS A IZQUIERDA (LINER)  /1000000*0.184"
         Top             =   1830
         Width           =   1800
      End
      Begin VB.TextBox Text10 
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
         Index           =   3
         Left            =   7320
         TabIndex        =   63
         TabStop         =   0   'False
         ToolTipText     =   "LINER.Micronaje"
         Top             =   1830
         Width           =   1800
      End
      Begin VB.TextBox Text9 
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   62
         TabStop         =   0   'False
         ToolTipText     =   "Cantidad Según Configuración Big Bag"
         Top             =   1830
         Width           =   1800
      End
      Begin VB.TextBox Text13 
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
         Index           =   2
         Left            =   12720
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1515
         Width           =   1800
      End
      Begin VB.TextBox Text12 
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
         Left            =   10920
         Locked          =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1515
         Width           =   1800
      End
      Begin VB.TextBox Text11 
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
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1515
         Width           =   1800
      End
      Begin VB.TextBox Text10 
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
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1515
         Width           =   1800
      End
      Begin VB.TextBox Text9 
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1515
         Width           =   1800
      End
      Begin VB.TextBox Text13 
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         ToolTipText     =   "C3 * C6"
         Top             =   885
         Width           =   1800
      End
      Begin VB.TextBox Text12 
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
         Index           =   0
         Left            =   10920
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   885
         Width           =   1800
      End
      Begin VB.TextBox Text11 
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
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   885
         Width           =   1800
      End
      Begin VB.TextBox Text10 
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
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   885
         Width           =   1800
      End
      Begin VB.TextBox Text9 
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         ToolTipText     =   "Cantidad Según Configuración Big Bag"
         Top             =   885
         Width           =   1800
      End
      Begin VB.TextBox Text7 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   885
         Width           =   1800
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   90
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Text            =   "LIBRE"
         Top             =   3090
         Width           =   1800
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   89
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Text            =   "LIBRE"
         Top             =   2775
         Width           =   1800
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   88
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "LIBRE"
         Top             =   2460
         Width           =   1800
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   87
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "ENVOLTORIO"
         Top             =   2145
         Width           =   1800
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   86
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Text            =   "LINER "
         Top             =   1830
         Width           =   1800
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   85
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "IMPRESIÓN"
         Top             =   1515
         Width           =   1800
      End
      Begin VB.TextBox TXTFORMAENTREGA 
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
         Left            =   2520
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   7485
         Width           =   5130
      End
      Begin VB.TextBox TXTFMAPAGO 
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
         Left            =   2520
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   6840
         Width           =   5130
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   84
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "SOBRE [unidades]"
         Top             =   1200
         Width           =   1800
      End
      Begin VB.TextBox Text7 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Porta documento Tipo	"
         Top             =   1200
         Width           =   1800
      End
      Begin VB.TextBox Text9 
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "Porta documento Cantidad"
         Top             =   1200
         Width           =   1800
      End
      Begin VB.TextBox Text10 
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
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1800
      End
      Begin VB.TextBox Text11 
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
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1800
      End
      Begin VB.TextBox Text12 
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
         Index           =   1
         Left            =   10920
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1800
      End
      Begin VB.TextBox Text13 
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "C3 * C6"
         Top             =   1200
         Width           =   1800
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   94
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "COSTO M. DE OBRA %"
         Top             =   3759
         Width           =   1920
      End
      Begin VB.TextBox TXTCOSTODIRBOLSON 
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "COSTO MATERIALES + COSTO M. DE OBRA"
         Top             =   4395
         Width           =   1815
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   96
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "COSTO DIRECTO DEL BOLSON"
         Top             =   4395
         Width           =   2550
      End
      Begin VB.TextBox TXTCOSTOMOBRA 
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
         Left            =   12720
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   3759
         Width           =   1815
      End
      Begin VB.TextBox TXTCOSTOMATERIAL 
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3450
         Width           =   1815
      End
      Begin VB.TextBox Text7 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1515
         Width           =   1800
      End
      Begin VB.TextBox Text7 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "Ancho +   Fuelle  c/u * 2"
         Top             =   1830
         Width           =   1800
      End
      Begin VB.TextBox Text7 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Embalaje U/Fardo"
         Top             =   2145
         Width           =   1800
      End
      Begin VB.TextBox Text8 
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   885
         Width           =   1800
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
         Index           =   1
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         ToolTipText     =   "Porta documento Medida (TEXTO)"
         Top             =   1200
         Width           =   1800
      End
      Begin VB.TextBox Text8 
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1515
         Width           =   1800
      End
      Begin VB.TextBox Text8 
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "LINER.ALTO"
         Top             =   1830
         Width           =   1800
      End
      Begin VB.TextBox Text8 
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2145
         Width           =   1800
      End
      Begin VB.TextBox TEXT57 
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
         Index           =   0
         Left            =   1920
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2460
         Width           =   10800
      End
      Begin VB.TextBox TEXT57 
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
         Index           =   1
         Left            =   1920
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2775
         Width           =   10800
      End
      Begin VB.TextBox TEXT57 
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
         Index           =   2
         Left            =   1920
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3090
         Width           =   10800
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   92
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "COSTO MATERIALES %"
         Top             =   3450
         Width           =   1920
      End
      Begin VB.TextBox TXTLABEL 
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
         Index           =   95
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "% Costo de Materiales contra Costo Total"
         Top             =   3450
         Width           =   630
      End
      Begin VB.TextBox TXTLABEL 
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
         Index           =   97
         Left            =   11640
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "% Costo de Mano de Obra contra Costo Total"
         Top             =   3759
         Width           =   630
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   104
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "ETIQUETA [unidades]"
         Top             =   885
         Width           =   1800
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   103
         Left            =   10920
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "PRECIO (u$s/u)"
         Top             =   250
         Width           =   1800
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   105
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "PRECIO (u$s)"
         Top             =   250
         Width           =   1800
      End
      Begin VB.CommandButton Command47 
         Caption         =   "Guardar Cotización"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   7080
         TabIndex        =   12
         ToolTipText     =   "Guarda Datos de Cotización"
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox TXTTRANSPORTE 
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
         Left            =   12720
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   570
         Width           =   1800
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   109
         Left            =   10920
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "Transp.(Cotización 1)"
         Top             =   570
         Width           =   1800
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   110
         Left            =   12260
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "U$S"
         Top             =   3759
         Width           =   460
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   111
         Left            =   12260
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "U$S"
         Top             =   4395
         Width           =   460
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   112
         Left            =   12260
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "U$S"
         Top             =   3450
         Width           =   460
      End
      Begin VB.TextBox TXTLABEL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   113
         Left            =   6120
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "FECHA DE EMISIÓN"
         Top             =   4920
         Width           =   1920
      End
      Begin VB.TextBox Text15 
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
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton Command48 
         Appearance      =   0  'Flat
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1850
         TabIndex        =   4
         ToolTipText     =   "Refresca % Sugerido"
         Top             =   4485
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.CommandButton Command49 
         Caption         =   "Aprobación de Cotización"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1800
         TabIndex        =   3
         ToolTipText     =   "Aprobación de Cotización Solo Usuarios Habilitados"
         Top             =   240
         Width           =   2415
      End
      Begin VB.CommandButton Command50 
         Caption         =   "Generar x Copia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Se Genera x Copia de la Cotización Activa"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TXTPORCENTAJEHILO 
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
         Left            =   1920
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "% Costo de Materiales contra Costo Total"
         Top             =   3400
         Width           =   720
      End
      Begin VB.TextBox TXTPORCENT2 
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
         Left            =   12260
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   6240
         Width           =   460
      End
      Begin VB.TextBox TXTLABEL2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
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
         Left            =   9720
         Locked          =   -1  'True
         TabIndex        =   86
         TabStop         =   0   'False
         Text            =   "PRECIO DE VENTA"
         Top             =   5925
         Width           =   2550
      End
      Begin VB.TextBox TXTPORCENT2 
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
         Left            =   12260
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   5925
         Width           =   460
      End
      Begin VB.TextBox TXTUSSUGERIDO 
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
         Left            =   12260
         Locked          =   -1  'True
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   5640
         Visible         =   0   'False
         Width           =   460
      End
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuPrint 
         Caption         =   "Imprimir Cotización Activa"
      End
      Begin VB.Menu mnuActuformula 
         Caption         =   "Actualizar Formula Según Cotización Activa"
      End
      Begin VB.Menu mnuHabiltCot 
         Caption         =   "Habiliitar Nuevamente Cotización Activa"
      End
   End
End
Attribute VB_Name = "FRMCOTI2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Sub AGREGA_ROTULO_CONTEPLAST(CI1%)
    'ULTIMA COTIZACION DE ESTE CODIGO
    'ESTA RUTINA LE AGREGA EL ROTULO DE LA OPERACION A LA TABLA FORMULAS PARA CUANDO SABER DE ACUERDO A LA FORMULA DENTRO DE QUE OPERACION ESTA.
    'TAMBIEN EL CORTE Y EL MODUXBB (CANTIDAD DE CARAS)
    ULTIMACOTIZACION& = XVALO(VALOBADA("COTI1", "MAX(NROCOTIZACION)", "CODINTCONJ = " & CI1%, "NOMESS"))
    NUORFORMULA% = 0
    SQLX$ = "SELECT * FROM COTI1 WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NROCOTIZACION  = " & ULTIMACOTIZACION&
    SQLX$ = SQLX$ + " and CODINTELEM > 0 and MODUXBB > 0 and CORTE > 0 "
    SQLX$ = SQLX$ + " ORDER BY COTI1_ID"
    Set RS = READSET(SQLX$)
        With RS
          Do While Not .EOF
            NUORFORMULA% = NUORFORMULA% + 1 'PARA EVITAR QUE SE GRABE + DE UN REGISTRO POR VEZ EN EL CASO QUE LA CONDI DEVUELVA 2
            'O SEA QUE HAYA UN CODIGO REPETIDO EN LA FORMULA
            Codigo$ = SAFETEXT$(!Codigo)
            CI2% = XVALO(!CODINTELEM)
            
            CIRANG = CODIGORANGO&(CI2%)
            MODUXB = XVALO(!moduxbb)
            
            KORTE = TRIM$(FORMATNUM$(XVALO(!corte), "f15.2"))
            If CIRANG > 0 Then CI2% = CIRANG 'LE ASIGNA EL CODIGO RANGO (EL CODIGO RANGO DEBE TENER MISMA U.MEDIDA
            ROTULOGRID$ = SAFETEXT$(!COMPONENTE)
            CONDI$ = "NUORIT = " & NUORFORMULA%
            CONDI$ = CONDI$ + " AND CODICONJ = " & CI1%
            CONDI$ = CONDI$ + " AND CodiElem = " & CI2%

            POS1% = InStr(ROTULOGRID$, "(")
            If POS1% > 0 Then ROTULOGRID$ = TRIM$(Mid$(ROTULOGRID$, 1, POS1% - 1))
            '
            SQLX$ = "UPDATE FORMULAS SET COMPONENTE = '" & ROTULOGRID$ & "'"
            SQLX$ = SQLX$ + " ,MODUXBB = " & MODUXB
            SQLX$ = SQLX$ + " ,CORTE = '" & KORTE & "'"
            SQLX$ = SQLX$ + " WHERE " & CONDI$
            FLEXCONN.Execute SQLX$
            .MoveNext
          Loop
        End With
    RS.Close
    Set RS = Nothing

End Sub

Sub AGREGARLINER(codiconjunto%, NROFT&, NUORIT%)
            'NUORIT% VIENE COMO PARAMETRO PARA SABER CUAL FUE EL ULTIMO REGISTRO ENUMERADO DE ORDEN
            LINER_SI_NO$ = UCase$(TRIM$(VALOBADA("RECETACTP", "LINERSINO", "NUMRECETACTP = " & NROFT&, "NOMESS")))
            If LINER_SI_NO$ = "NO" Or LINER_SI_NO$ = "" Then Exit Sub
            If NROFT& < 1 Then Exit Sub
            
            If LINER_SI_NO$ = "SI" Then
                COD_EXTERANGO$ = Control("CONFLINE", "CORANLIN")
                codielem% = CODINT%(COD_EXTERANGO$)
                USOX$ = "1" 'USO FUERZO A UNO
                DESCRIX$ = DESCRIT0$(codielem%) 'LO LEO DE LA FUNCION POR SI ES RANGO CAMBIA
                Codigo$ = CODEXT$(codielem%)    'LO LEO DE LA FUNCION POR SI ES RANGO CAMBIA
                PDX$ = 0
                COMPONENTEX = DESCRIX$
                MODUXBBX = 1
                CORTEX = 1
                UNXX$ = Unimed$(codielem%)
                
                    SQLX$ = "SELECT * FROM Formulas WHERE CodIConj < 1"
                    'GUARDA EN TABLA CABEZERA COTIZACTP
                    Dim RS1 As ADODB.Recordset
                    Set RS1 = New ADODB.Recordset
                    RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                    With RS1
                      .AddNew
                      !CODICONJ = codiconjunto%
                      !CODIEMPR = CodiEmp%
                      !CODXCONJ = Left(CODEXT$(codiconjunto%), 16)
                      !codielem = codielem%
                      !CodXElem = Codigo$
                      NUORIT% = NUORIT% + 1
                      !NUORIT = NUORIT%
                      !usobruto = XVALO(USOX$)
                      !UNIS_FORM = XVALO(USOX$)
                      !UMED_FORM = UNXX$
                      !LAR_CORTE = 0
                      !PLADISPO = 0
                      !RefeComp = ""
                      !MARBORRA = 0
                      !COMPONENTE = DESCRIX$
                      !moduxbb = 1
                      !corte = CORTEX
                      .Update
                    End With
                    RS1.Close
                    Set RS1 = Nothing
         End If

End Sub

Sub GRABARENPRESUVEN(NCLIE, NROCOTI&, CANTI, CI1%, NVEND%, Codigo$, CODBIG$, NROFICHATEC&, MON%, PREUNIT#, OBSERVAX$, OKX%, MarcarAnterioresComoNoVigentes%, PorComisReal#)

   Call ABREPRESVEN

   OKX% = -1
   '
   
   If NCLIE < 1 Then
      Call COMUNI("N° de Cotización Inexistente o Falta N° de Cliente")
      OKX% = -1
      Exit Sub
   End If
   '
   APLIVA% = 1

   If NCLIE > 0 Then
       PVCLI% = PIVACLI%(NCLIE)
       If PVCLI% = 5 Or PVCLI% = 7 Then APLIVA% = 0
   End If

   HOII% = HOY(HOS$)
   FEMI% = HOII%
   '
   'ATENCION SI PARAMETRO MarcarAnterioresComoNoVigentes% > 0 MARCA COMO NO VIGENTES SI HAY PRESUPUESTOS PARA EL MISMO CLIENTE Y MISMO CODIGO
   'COMO ESTO DEBE ESTAR DENTRO DE UN BEGIN SI LLEGARA A HABER ALGUN PROBLEMA NO SE GRABARIA TAMPOCO ESTAS TRANSACCION
   If MarcarAnterioresComoNoVigentes% > 0 Then
    CONDI$ = "(  MarNoVig Is Null   Or MarNoVig = '' )"
    CONDI$ = CONDI$ + " And Codiint = " & CI1%
    CONDI$ = CONDI$ + " And NROCLIE = " & NCLIE
    cantanteriores% = CantRegistros&("PRESUVEN", CONDI$, "NOMESS")
    If cantanteriores% > 0 Then
       strSQL = "Update Presuven set MarNoVig = '*'"
       strSQL = strSQL + " where " & CONDI$
       FLEXCONN.Execute (FILTSQL$(strSQL))
    End If
   End If

   
   strSQL = "Select * from Presuven"
   strSQL = strSQL + " where NroPres > 0"
   Dim PresuVen As ADODB.Recordset
   Set PresuVen = New ADODB.Recordset
   PresuVen.Open FILTSQL$(strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   NEWPRES& = NUPRESUP&
   With PresuVen


     CODICL$ = "" '
     'no lo encuentra entonces Inserta
        '\\\OT-06-0168-WAR-09/05/06///
        CANTI0 = XVALO(CANTI)
        
        If CANTI0 < 1 Then CANTI0 = 1
        '\\\OT-06-0168-WAR-FIN///

       .AddNew   ' YA CONVERTIDO

       !CODIEMPR = CodiEmp%
        !NroPres = NEWPRES&
        '\\\OT-06-0168-WAR-09/05/06///
         !NumeCRubro = 0
         !Cod_Rubro = ""
         !Cod_Anio = ""
        '\\\OT-06-0168-WAR-FIN///
        !FECHEMIS = CVDAT(FEMI%)
        !NroClie = NCLIE

        RC$ = Left$(rasocli$(NCLIE), 64)
        !RasoClie = RC$
        REFE$ = "Cot: " & NROCOTI& & " F.T: " & NROFICHATEC&
        !NROCOTI = SAFETEXT$(NROCOTI&) 'es de texto
        !Referen = REFE$
        !AteSen = Left$(RC$, 48)
        !LisPre = 0
        !ConPag = 0
        !VENDED = XVALO(NVEND%)

        
        !MonedEmi = MON%
        !NuOrdIt = 1
        !CodiInt = CI1%
        !CODIEXT = Left$(Codigo$, 16)
        !Descrip = Left$(DESCRIT0$(CI1%), 64)
        !Especif = ""
        !Observa = OBSERVAX$
        !PreList = PREUNIT# 'XVALO(Total(0)) / CANTI0 'XVALO(Total(0))'\\\OT-06-0168-WAR-09/05/06///
        !TDESCUEN = ""
           DESC$ = ""
        !PorcDesc = COEFDESCU(DESC$)
        !ImpoDesc = 0
        '
        '

        
        TOT# = PREUNIT# * CANTI
        If GRUPIVA%(CI1%) < 1 Then
            IMPOIVA# = APLIVA% * (XVALO(IMPIVA21) + ROUND#(TOT# * PORCEIVA(CI1%) / 100, 2))
            IMPIVA21 = TRIM$(FORMATNUM$(IMPOIVA#, "F12.2"))
          Else
            IMPOIVA# = APLIVA% * (XVALO(IMPIVA105) + ROUND#(TOT# * PORCEIVA(CI1%) / 100, 2))
            IMPIVA105 = TRIM$(FORMATNUM$(IMPOIVA#, "F12.2"))
        End If



        !PREUNIT = PREUNIT# 'XVALO(Total(3))'\\\OT-06-0168-WAR-09/05/06///
        !Cantit = CANTI '1'\\\OT-06-0168-WAR-09/05/06///
        !ImToNe = PREUNIT# * CANTI '1 * XVALO(Total(3))'\\\OT-06-0168-WAR-09/05/06///
        !FechVent = CVDAT(0)
        !NroPedi = 0
        !NroFac_Cor = ""
        !NroFac_Lar = ""
        !Codi_Cli = CODICL$ ' GRABA CODIGO DE CUENTA DEL CLIENTE / PROSPECT
        !TTIVA21 = XVALO(IMPIVA21)
        !TTIVA105 = XVALO(IMPIVA105)
        !PresComReal = XVALO(PorComisReal#)
'        !TOTPER = XVALO(TOTPER)

       .Update

        End With

      PresuVen.Close
      Set PresuVen = Nothing
      OKX% = 0
End Sub

Function NUPRESUP&()
   '
   NPP& = 1
   Call ABREPRESVEN
   '
   Dim MIRS As ADODB.Recordset
   strSQL = ("SELECT max(NroPres) as maximo FROM PresuVen")
   Set MIRS = FLEXCONN.Execute(FILTSQL$(strSQL))
   If Not IsNull(MIRS!Maximo) Then
      NPP& = 1 + MIRS!Maximo
   End If
   NUPRESUP& = NPP&
   MIRS.Close
   Set MIRS = Nothing
   '
End Function

Private Sub Command1_Click()
   Hide
End Sub

Private Sub Command2_Click()
    Erase RESP_ZELE_VECT
    FTEC& = XVALO(Text5)

    Campos$ = "Rótulo/A48;" '1
    Campos$ = Campos$ + "Texto/A48;"       '2
    Campos$ = Campos$ + "Valor/A12;"       '3
    Campos$ = Campos$ + "Ftecnica/A0"     '4
    
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = "Datos Anexos Cargados"
    Dim obj As New RECORXET
    CAMPOSTABLA$ = "CAMPO1 ,CAMPO2 ,CAMPO3 ,CAMPO4"
    TABLA$ = "DATASQL "
    CONDI$ = " CAMPO4  = '" & CStr(FTEC&) & "' AND LISTA = 'DATOSLIB'"
    '
    Set RS1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
    '
    FRMZELECHO.Width = 15705
50  FRMZELECHO.Show 1

End Sub

 Sub Command47_Click()
        FRMCOTI2.Command47.Enabled = False
        ESNUEVO% = 0
        If COTICONT.lblNuevaCotizacion.Visible = True Then
          ESNUEVO% = 1
        End If
        '
        If XVALO(COTICONT.TXTNUMCOTIZACION) < 1 Then
           Call COMUNI("Debe Seleccionar Una Cotización")
           GoTo 99
        End If

'        If FECHANUM(TXTFEENTRECOTI) < 1 Then
'           Call COMUNI("Falta Fecha de Entrega")
'           GoTo 99
'        End If
        If XVALO(COTICONT.TXTNUMCOTIZACION) < 1 Then
           Call COMUNI("Falta N° de Cotización")
           GoTo 99
        End If
        
        SQLX$ = "SELECT * FROM COTIZACTP "
        If ESNUEVO% < 1 Then
         SQLX$ = SQLX$ + " WHERE NROCOTIZACION  = " & XVALO(COTICONT.TXTNUMCOTIZACION)
        Else
         SQLX$ = SQLX$ + " WHERE NROCOTIZACION  < 1 "
        End If

        CONDI$ = "NROCOTIZACION = " & XVALO(COTICONT.TXTNUMCOTIZACION)
        CONDI$ = CONDI$ + " AND STATUS = 1 "
        If CantRegistros&("COTIZACTP", CONDI$, "NOMES") > 0 Then
           Call COMUNI("Cotización Activa Aprobada\Imposible Aprobar !!")
           GoTo 99
        End If

        On Error GoTo ERROR_GUARDAR
        FLEXCONN.BeginTrans
        'GUARDA EN TABLA CABEZERA COTIZACTP
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With RS
         If ESNUEVO% > 0 Then
             NEWCOTIZ = NumNuevaCotiza
             .AddNew
             !NROCOTIZACION = NEWCOTIZ
             !FECHEMISI = CVDAT(HOY(HO$))
             Else
               NEWCOTIZ = XVALO(COTICONT.TXTNUMCOTIZACION)
         End If
         !rasocli = COTICONT.TXTRAZOCLI
         !Status = 0
         !NUMRECETACTP = XVALO(COTICONT.Text5)
         !NUCLIEN = XVALO(COTICONT.TEXT1)
         !CODIMASTER = COTICONT.TXTCODIGO
         .Update
        End With
        RS.Close
        Set RS = Nothing

        Dim RS1 As ADODB.Recordset
        Set RS1 = New ADODB.Recordset
        '
        
        SQLX$ = "SELECT * FROM COTI2"
        SQLX$ = SQLX$ + " WHERE NROCOTIZACION  < 1 "
        If ESNUEVO% < 1 Then
           Call ACTUREGISTRO("COTI2", "MARBORRA", "NROCOTIZACION = " & NEWCOTIZ, 1, REGAF&, "NOMES")
        End If
        
        RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With RS1
                .AddNew
                !NROCOTIZACION = NEWCOTIZ
                !ETIQ_MODULO = XVALO(FRMCOTI2.Text9(0))
                !ETIQ_COSTO = XVALO(FRMCOTI2.Text12(0))
                !ETIQ_COSTO = XVALO(FRMCOTI2.Text12(0))
                !SOBRE_DOCTIPO = FRMCOTI2.Text7(1)
                !SOBRE_DOCMED = FRMCOTI2.Text8(1)
                !SOBRE_MODULO = XVALO(FRMCOTI2.Text9(1))
                !SOBRE_COSTO = XVALO(FRMCOTI2.Text12(1))
              
                !LINER_ANCHO = XVALO(FRMCOTI2.Text7(3))
                !LINER_ALTO = XVALO(FRMCOTI2.Text8(2))
                !LINER_COSTO = XVALO(FRMCOTI2.Text12(3))
                !LINER_MODULO = XVALO(FRMCOTI2.Text9(3))
                !LINER_KG = XVALO(FRMCOTI2.Text11(3))
                !IMPRE_COSTO = XVALO(FRMCOTI2.Text12(2))
                !IMPRE_MODULO = XVALO(FRMCOTI2.Text9(2))
                !ENVO_EMBF = XVALO(FRMCOTI2.Text7(4))
                !ENVO_MOD = XVALO(FRMCOTI2.Text9(4))
                !ENVO_COSTO = XVALO(FRMCOTI2.Text12(4))
                If FECHANUM(FRMCOTI2.Text15) < 1 Then
                   !FECHEMISI = CVDAT(HOY(HO$))
                End If
                !LIBRE1TXT = FRMCOTI2.TEXT57(0)
                !LIBRE2TXT = FRMCOTI2.TEXT57(1)
                !LIBRE3TXT = FRMCOTI2.TEXT57(2)
                !VALIMPRESION = XVALO(FRMCOTI2.Text13(2))
                !LIBRE1IMPO = XVALO(FRMCOTI2.Text13(5))
                !LIBRE2IMPO = XVALO(FRMCOTI2.Text13(6))
                !LIBRE3IMPO = XVALO(FRMCOTI2.Text13(7))
                !MODOPAGO = Left$(FRMCOTI2.TXTFMAPAGO, 128)
                !FORMAENTREGA = Left$(FRMCOTI2.TXTFORMAENTREGA, 1024)
                !PLAZOENTRE = Left$(FRMCOTI2.Text3, 64)
                !MANTOFERTA = Left(FRMCOTI2.TXTMANTENOFERTA, 128)
                
                !SCRAPP = XVALO(TXTSCRAP)
'                !LOGYFLETE = XVALO(FRMCOTI2.TXTUS(1))
'                !IMPUESTO = XVALO(FRMCOTI2.TXTUS(2))
                !COMISION = XVALO(FRMCOTI2.TXTCOMISION)
'                !ESTRUCTURA = XVALO(FRMCOTI2.TXTUS(4))
'                !GASTFINAN = XVALO(FRMCOTI2.TXTUS(5))
'                !OTROS = XVALO(FRMCOTI2.TXTUS(6))
'                !OTRO1 = XVALO(FRMCOTI2.TXTUS(7))
'                !OTRO2 = XVALO(FRMCOTI2.TXTUS(8))
                !COSTTOT = XVALO(FRMCOTI2.TXTUSS2(0))
                !PJEUTILIDAD = XVALO(FRMCOTI2.TXTPORCENT2(1))
                !COSTMOBRA = XVALO(FRMCOTI2.TXTCOSTOMOBRA)
                !COSTOSFIJOS = XVALO(FRMCOTI2.TXTCOSTOSFIJOS)
                !PREVENTA = XVALO(FRMCOTI2.TXTUSS2(2))
                !PREVENTAREAL = XVALO(FRMCOTI2.TXTUSS2(3))
                !PREVENTAFINAL = XVALO(FRMCOTI2.TXTPRECIOFINAL) 'CON LINER
                !IMPOLINER = XVALO(FRMCOTI2.TXTVALORLINER)  '
                !ComPorReal = XVALO(TXTPORCENT2(4))
                !MICRONALJELINER = XVALO(FRMCOTI2.Text10(3))
                !CODIEMPR = CodiEmp%
                !OBSERVACIONES = TEXT1
                !CANTICONFIRMADA = ROUND(XVALO(TXTCANTSOLIC))
                !CODESPECIFINT = TRIM$(TXTESPECIFCLIENTE_2)
                !PORCENTCOMI = XVALO(TXTPORCENT(1))
                !MARBORRA = 0
                .Update
        End With
        RS1.Close
        Set RS1 = Nothing
        '
        CONDI$ = "SOLICOTI_ID = " & XVALO(COTICONT.LBLNROSOLICITUDCOTIZACION)
        Call ACTUREGISTRO("SOLICOTI", "NROCOTIZACION", CONDI$, NEWCOTIZ, REGAF&, "NOMESS")
ERROR_GUARDAR:
        If Err <> 0 Then
           MsgBox Err.Description
           FLEXCONN.RollbackTrans

           On Error GoTo 0
           Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
           GoTo 99
        Else
           Call BORRAREGISTROS("COTI2", "NROCOTIZACION = " & NEWCOTIZ & " AND MARBORRA = 1", REGAF&, "NOMES")
           FLEXCONN.CommitTrans
           COTICONT.lblNuevaCotizacion.Visible = False
           MODIFICCOT2% = 0
        End If
        '
           
99      FRMCOTI2.Command47.Enabled = True
End Sub

Private Sub TXTCOINDBOLSON_Change()
   Call PORCENTAJECOSTOINDIRECTOBOLSON

End Sub

Private Sub Form_Activate()
    If Command47.Enabled = False Then
       Call BLOQUEAR_CONTROLES(Me, 2, 0, 1)
    Else
       Call BLOQUEAR_CONTROLES(Me, 1, 0, 1)
    End If
End Sub

Private Sub Form_Load()
    If Command47.Enabled = False Then
       Call BLOQUEAR_CONTROLES(Me, 2, 0, 1)
    End If
    
    'SI TIENE CARGADO DATOS AUX MUESTRA EL BOTON PARA VISUALIZARLOS.
    FRMCOTI2.CMDDATOSAUXILIARES.Visible = False
    FTEC& = XVALO(COTICONT.Text5)
    CONDI$ = " CAMPO4  = '" & CStr(FTEC&) & "' AND LISTA = 'DATOSLIB'"
    If CantRegistros&("DATASQL", CONDI$, "NOMESS") > 0 Then
       FRMCOTI2.CMDDATOSAUXILIARES.Visible = True
    End If

End Sub

Private Sub mnuActuformula_Click()
   CI1% = CODINT%(COTICONT.TXTCODIGO)
   If CI1% < 1 Then Exit Sub
'   ACTUALIZA LA ESTRUCTURA BORRA LA ACTUAL Y LA GENERA DEVUELTA
   Call ACTUALIZA_ESTRUCTURA_CONTEPLAST(CI1%)
'   Call AGREGA_ROTULO_CONTEPLAST(CI1%) 'AGREGA EL ROTULO A LA TABLA FORMULAS INDICANDO A QUE TAREA CORRESPONDE FONDO,TAPA ETC
   Call COMUNI("Proceso Realizado")
End Sub

Private Sub mnuHabiltCot_Click()
   '******VALIDAR SI LA MISMA YA ESTA APROBADA
   CONDI$ = "NROCOTIZACION = " & XVALO(Text14)
   If CantRegistros&("COTIZACTP", CONDI$, "NOMES") < 1 Then
      Call COMUNI("La Cotización Activa no ha Sido Grabada\Deberá Grabar la Misma para Luego Aprobarla")
      Exit Sub
   End If
   
   CONDI$ = CONDI$ + " AND STATUS < 1 "
   If CantRegistros&("COTIZACTP", CONDI$, "NOMES") > 0 Then
      Call COMUNI("Cotización Activa No Aprobada\Al Aprobar Se Generará la Habilitación Para Poder Generar el Pedido")
      Exit Sub
   End If
   
   CI1% = CODINT%(COTICONT.TXTCODIGO)
   NVEND% = XVALO(COTICONT.Text4)
   Codigo$ = COTICONT.TXTCODIGO
   CODBIG$ = COTICONT.TXTCODCONFECCION
   NROFICHATEC& = COTICONT.Text5
   NC = XVALO(COTICONT.TEXT1)
   PorComisReal# = XVALO(TXTPORCENT2(4))
                   
   A = rasocli$(1) 'para refrescar el recordset
   
   If TRIM$(rasocli$(NC)) = "" Then
      Call COMUNI("Cliente No Válido")
      Exit Sub
   End If

   If CI1% < 1 Then
      Call COMUNI("Código No Válido")
      Exit Sub
   End If
   If NVEND% < 1 Then
      Call COMUNI("Vendedor No Válido")
      Exit Sub
   End If
    If CODBIG$ = "" Then
      Call COMUNI("Código BigBag No Válido")
      Exit Sub
   End If
   If NROFICHATEC& < 1 Then
      Call COMUNI("Número Ficha Técnica No Válido")
      Exit Sub
   End If
   
   FIRMAELEC.TEXT1 = ""
   FIRMAELEC.Text2 = ""
   FIRMAELEC.Label1 = ""
   If XVALO(Text14) < 1 Then
      Call COMUNI("Debe Haber una Cotización Activa")
      Exit Sub
   End If

   FIRMAELEC.Show 1

   If FIRMAELEC.ESTADO < 1 Then
      Exit Sub
   End If
   
   'TRANSACCIÓN GRABACION
   On Error GoTo ERROR_GUARDAR
   FLEXCONN.BeginTrans
   Call GRABARENPRESUVEN(NC, XVALO(Text14), 1, CI1%, NVEND%, Codigo$, CODBIG$, NROFICHATEC&, 2, XVALO(TXTPRECIOFINAL), TEXT1, OKX%, 1, PorComisReal#)
   If OKX% < 0 Then GoTo 70
    
ERROR_GUARDAR:
   If Err <> 0 Then
70     FLEXCONN.RollbackTrans

       On Error GoTo 0
       MsgBox Err.Description
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       Exit Sub
   Else
      FLEXCONN.CommitTrans
      
      Call COMUNI("Grabación Realizado Exitosamente")
   End If

   
End Sub

Private Sub mnuPrint_Click()
    If DERACCE%("IMPRCOTI", "IMPRIMIR COTIZACIONES") < 2 Then Exit Sub

    NROCOT& = XVALO(Text14)
'    CAMPOSTABLA$ = "FT_NROFT,FT_FECHEMISI, COT_NROCOT,COT_FECHEMISI, FT_CTACLTE,FT_RASOCLI,FT_STAT,FT_CODIGO,FT_ID,COT_ID,FT_NUVEND"
    TABLA$ = "FITECYCOTIZA" 'ES UNA VISTA
    CONDI$ = "FT_NROFT > 0 "
    SQLX$ = "SELECT * FROM FITECYCOTIZA WITH(NOLOCK) WHERE COT_NROCOT = " & NROCOT&
    Set RS = READSET(SQLX$)
    With RS
      NROFT = XVALO(!FT_NROFT)
      NroVend% = XVALO(!FT_NUVEND)
    End With
    FE% = HOY(HO$)
    VENDEDOR$ = TRIM$(ECOARCH$("VENDEDOR", Chr$(NroVend%)))
    FECHALARGA$ = FormatDateTime(FECHATEX$(FE%), vbLongDate)
    'para que en vez de lunes, 3 de febrero quede Lunes, 3 de Febrero (mayuscula en dia y mes)
    Dim S() As String, i As Integer
    S = Split(FECHALARGA$, " ")
    For i = LBound(S) To UBound(S)
       If S(i) <> "de" Then
         Mid$(S(i), 1, 1) = UCase$(Mid$(S(i), 1, 1))
       End If
    Next
    FECHALARGA$ = Join$(S, " ") ' concatena todo el vector
    '
    Call SETPRINPORT("RESTORE")
    PUERTOIMPRE$ = Control$(IDENTER$, "PRINTSER")
    If PUERTOIMPRE$ = "" Then PUERTOIMPRE$ = Control$("*", "PRINTSER")
    If PUERTOIMPRE$ <> "" Then
       Call SETPRINPORT(PUERTOIMPRE$)
    End If
    '
    OPX% = COMALTER%("Selección Empresa\\Cancelar\Conteplast\Envaplast")
    
    FILTX$ = FECHALARGA$ & ";" & SAFETEXT$(NROFT) & ";" & SAFETEXT$(NROCOT&) & ";" & ";" & VENDEDOR$ & ";"
    If OPX% = 2 Then
      Call STDFORM("COIN-ENV", FILTX$)
    ElseIf OPX% = 3 Then
      Call STDFORM("COIN-CTP", FILTX$)
    End If

    Call SETPRINPORT("RESTORE")
    If XVALO(Text14) > 0 Then
    
    Else
        Call COMUNI("No se Detecta una Cotización Activa")
    End If
End Sub

Private Sub PORCEDIT_Change()

End Sub


Private Sub TXTCANTSOLIC_DblClick()
    If TXTCANTSOLIC.BackColor = &HE0E0E0 Then
       TXTCANTSOLIC.BackColor = &HFFFFFF
       TXTCANTSOLIC.Locked = False
      Else
       TXTCANTSOLIC.BackColor = &HE0E0E0
       TXTCANTSOLIC.Locked = True
    End If
    
End Sub


Private Sub TXTCODESPVPD_Change()
   TXTESPECIFCLIENTE_2 = TXTCODESPVPD
End Sub

Private Sub TXTCOMISION_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        Call PINTATEXT(Me.TXTUSS2(3))
   End If

End Sub


Private Sub TXTCOMISION_LostFocus()
   TXTUSS2(0) = FORMATNUM$(XVALO(COSTO_TOTAL#), "F12.4")
   TXTUSS2(2) = FORMATNUM$(SUMARBLOQUETOTALES#, "F12.4")

End Sub


Private Sub TXTCOSTODIRBOLSON_LostFocus()
          TXTUSS2(2) = FORMATNUM$(SUMARBLOQUETOTALES#, "F12.4")

End Sub

Private Sub TXTCOSTOSFIJOS_Change()
MODIFICCOT2% = 1
End Sub

Private Sub TXTCOSTOSFIJOS_LostFocus()
      Call CALCULARCOSTOMATERIALES
      Call CALCULARCOSTODIRECTOBOLSON
      TXTCOSTOMOBRA = FORMATNUM$(XVALO(TXTCOSTOMOBRA), "F12.4")
      Me.TXTCOSTOSFIJOS = FORMATNUM$(XVALO(TXTCOSTOSFIJOS), "F12.4")
      Call porcentajeCostoManoObra
      TXTUSS2(2) = FORMATNUM$(SUMARBLOQUETOTALES#, "F12.4")

End Sub

Private Sub TXTPORCENT_Change(Index As Integer)
   Call CALCUPORCENTAJE_COMISION

End Sub

Private Sub TXTPORCENT_LostFocus(Index As Integer)
      If Index = 1 Then
         PORC = XVALO(TXTPORCENT(Index))
         'TXTCOMISION = FORMATNUM$(XVALO(TXTUSS2(0)) + XVALO(TXTUSS2(1)) * PORC / 100, "F12.4")
'         TXTCOMISION = FORMATNUM$(XVALO(TXTUSS2(1)) * PORC / 100, "F12.4")
         TXTUSS2(2) = FORMATNUM$(SUMARBLOQUETOTALES#, "F12.4")
      End If
End Sub

Private Sub TXTPORCENT2_LostFocus(Index As Integer)
  If Index = 4 Then
    TXTPORCENT2(Index) = FORMATNUM$(TXTPORCENT2(Index), "F5.1")
  End If
End Sub

Private Sub TXTSCRAP_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        Call PINTATEXT(Me.TXTCOMISION)
   End If

End Sub

Private Sub TXTSCRAP_LostFocus()
     TXTSCRAP = FORMATNUM$(XVALO(TXTSCRAP), "F12.4")
     TXTUSS2(0) = FORMATNUM$(XVALO(COSTO_TOTAL#), "F12.4")
     TXTUSS2(2) = FORMATNUM$(SUMARBLOQUETOTALES#, "F12.4")

End Sub


Private Sub TXTSUGCOSTOINDBOLSON_Change()
   Call PORCENTAJESUGERIDOCOSTOINDIRECTOBOLSON
End Sub

Private Sub Command43_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Me.TXTFEENTRECOTI.TEXT = VALACT1$("SELFECHA")
   End If

End Sub


Private Sub TXTFMAPAGO_Change()
MODIFICCOT2% = 1
End Sub

Private Sub Text3_Change()
MODIFICCOT2% = 1
End Sub







Private Sub TXTFORMAENTREGA_Change()
MODIFICCOT2% = 1
End Sub
Sub PORCENTAJESUGERIDOCOSTOINDIRECTOBOLSON()
'       If XVALO(TXTSUGCOSTOINDBOLSON) > 0 Then
'           'B = XVALO(TXTSUGCOSTOINDBOLSON) * XVALO(TXTUSS2(0))
'           B = PORCENT_ENTRE_2_VALORES(XVALO(Me.TXTUSS2(0)), XVALO(Me.TXTSUGCOSTOINDBOLSON))
'           Me.TXTPTJE = FORMATNUM$(B, "F8.2")
'       Else
'          TXTPTJE = 0
'    End If

End Sub






Private Sub TXTUSS2_LostFocus(Index As Integer)
      TXTUSS2(Index) = FORMATNUM$(TXTUSS2(Index), "F12.4")
End Sub

Function PRECIOVENTA#()
   PRECIOVENTA# = XVALO(TXTUSS2(0)) + XVALO(TXTUSS2(1))
End Function

Sub PORCENTAJE_SUG_COSTOINDIRECTOBOLSON()
'    If XVALO(TXTUSS2(0)) > 0 Then
'       PCTJE = XVALO(TXTSUGCOSTOINDBOLSON) / XVALO(TXTUSS2(0))
'       Me.TXTPTJE = FORMATNUM$(PCTJE, "F9.2")
'    Else
'       TXTPTJE = 0
'    End If
    
    
End Sub

Sub PORCENTAJECOSTOINDIRECTOBOLSON()
   TXTUSS2(0) = FORMATNUM$(COSTO_TOTAL#, "F12.4")
   If XVALO(TXTCOINDBOLSON) > 0 Then
'     Me.TXTPJECOSTOINDBOLSONREAL = FORMATNUM$(PORCENT_ENTRE_2_VALORES(XVALO(TXTUSS2(0)), XVALO(TXTCOINDBOLSON)), "F9.2")
   Else
     TXTPJECOSTOINDBOLSONREAL = 0
   End If

    
End Sub

Sub SUMARCOSTOSINDIRECTOS()
   IMPOCOSIND# = 0
'   For i% = 1 To TXTUS.UBound
'      IMPOCOSIND# = IMPOCOSIND# + XVALO(TXTUS(i%))
'   Next i%
'   TXTCOINDBOLSON = FORMATNUM$(IMPOCOSIND#, "F12.2")
End Sub


Sub SUMARTOTAL()
   A = 0
   For i% = 0 To TXTUSS2.UBound - 1
      A = A + TXTUSS2(i%)
   Next i%
   TXTUSS2(3) = FORMATNUM$(XVALO(A), "F12.4")
End Sub

Function STATUSCOTIZA%(NROCOT&)
    '0 PENDIENTE DE APROBACION
    '1 APROBADA
    STATUSCOTIZA% = XVALO(VALOBADA("COTIZACTP", "STATUS", "NROCOTIZACION  = " & NROCOT&, "NOMESS"))
    
End Function

Private Sub TXTUSSUGERIDO_Change(Index As Integer)
   
   Call PORCENTAJESUGERIDOCOSTOINDIRECTOBOLSON
   Call SUMARCOSTOSINDIRECTOSSUGERIDOS
End Sub

Sub SUMARCOSTOSINDIRECTOSSUGERIDOS()
   IMPOCOSINDSUG# = 0
   For i% = 0 To TXTUSSUGERIDO.UBound
      IMPOCOSINDSUG# = IMPOCOSINDSUG# + XVALO(TXTUSSUGERIDO(i%))
   Next i%
   TXTSUGCOSTOINDBOLSON = FORMATNUM$(IMPOCOSINDSUG#, "F12.4")
End Sub

Private Sub Command45_Click()
   Command45.Enabled = False
'   For i% = 1 To TXTUSSUGERIDO.UBound
'      TXTUS(i%) = TXTUSSUGERIDO(i%)
'   Next i%
   Call SUMARCOSTOSINDIRECTOS
99 Command45.Enabled = True
End Sub
Sub CARGAROTULOSGASTOS(INCLUYEPORCENTAJES%)
   'CARGA LOS ROTULOS SEGUN LA CONFIGURACION DE LOS PORCENTAJES DE LOS GASTOS TABLA 'CTP.PCTCOT'"
    CONDI$ = " LISTA =   'CTP.PCTCOT'"
    SQLX$ = "SELECT CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK) WHERE " & CONDI$
    Maximo% = FRMCOTI2.TXTROTULOGASTOS.UBound
    i% = 0
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
         If i% = 2 Then Exit Do
         FRMCOTI2.TXTROTULOGASTOS(i%) = SAFETEXT$(!campo2)
         If INCLUYEPORCENTAJES% > 0 Then
            FRMCOTI2.TXTPORCENT(i%) = FORMATNUM$(XVALO(!CAMPO3), "F12.2")
         End If
         i% = i% + 1
         
'      If i% <= maximo% Then
'         FRMCOTI2.TXTROTULOGASTOS(i%) = SAFETEXT$(!campo2)
'         If INCLUYEPORCENTAJES% > 0 Then
'            FRMCOTI2.TXTPORCENT(i%) = FORMATNUM$(XVALO(!CAMPO3), "F12.2")
'         End If
'         i% = i% + 1
'      End If
     .MoveNext
      Loop
    End With
    
End Sub

Private Sub TXTPORCENT2_Change(Index As Integer)
   Call CALCUPORCENTAJE_MARGENUTILIDAD
   MODIFICCOT2% = 1
End Sub

Private Sub TXTUSS2_Change(Index As Integer)
   If Index = 0 Or Index = 1 Then
      TXTUSS2(2) = FORMATNUM$(PRECIOVENTA#, "F12.4")
      Call porcentajeCostoMaterial
   End If
   
   Call CALCULAPORCENTAJE_ENTRE_PVENTAREALYPRECIOVENTA
   Call porcentajeCostoManoObra
   Call CALCUPORCENTAJE_COMISION
   TXTUSS2(2) = FORMATNUM$(SUMARBLOQUETOTALES#, "F12.4")
   
'   TXTUSS2(1) = XVALO(TXTUSS2(0)) * (1 + XVALO(TXTUSS2(0)) * XVALO(TXTPORCENT2(1)) / 100)
    MODIFICCOT2% = 1
    If Index = 3 Then
       TXTVALORLINER_Change
    End If
End Sub


Sub CALCULAPORCENTAJE_ENTRE_PVENTAREALYPRECIOVENTA()
  If XVALO(TXTUSS2(2)) > 0 Then
     TXTPORCENT2(3) = FORMATNUM$(XVALO(TXTUSS2(3)) / XVALO(TXTUSS2(2)), "F9.4")
     A = XVALO(TXTUSS2(3))
     B = XVALO(TXTUSS2(2))
     TXTPORCENT2(3) = FORMATNUM$(PORCENT_ENTRE_2_VALORES(A, B), "F9.4")
  Else
     TXTPORCENT2(3) = 0
  End If
End Sub

Sub CALCUPORCENTAJE_MARGENUTILIDAD()
   If XVALO(TXTUSS2(0)) > 0 Then
       B = XVALO(TXTUSS2(0)) * XVALO(TXTPORCENT2(1)) / 100
       TXTUSS2(1) = FORMATNUM$(B, "F12.4")
   Else
      TXTUSS2(1) = 0
   End If

End Sub
Sub CALCUPORCENTAJE_COMISION()
   If XVALO(TXTUSS2(0)) > 0 Then
       B = (XVALO(TXTUSS2(0)) + XVALO(TXTUSS2(1))) * XVALO(TXTPORCENT(1)) / 100
       TXTCOMISION = FORMATNUM$(B, "F12.4")
   Else
      TXTCOMISION = 0
   End If

'PRECIOVENTA = CostoStandard / (1 - (PorcComision + PorcUtilidad) / 100)
'ValorComision = PRECIOVENTA * PorcComision / 100
End Sub

Private Sub Command48_Click()
    Call CARGAROTULOSGASTOS(1)
    Call CalcularCostosIndirectos
End Sub

Private Sub TXTCOSTOMOBRA_Change()
MODIFICCOT2% = 1
End Sub

Private Sub TXTCOSTOMOBRA_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'        Call PINTATEXT(TXTUS(1))
'   End If
  
End Sub

Private Sub TXTCOSTOMOBRA_LostFocus()

      Call CALCULARCOSTOMATERIALES
      Call CALCULARCOSTODIRECTOBOLSON
      TXTCOSTOMOBRA = FORMATNUM$(XVALO(TXTCOSTOMOBRA), "F12.4")
      
      Call porcentajeCostoManoObra
      TXTUSS2(2) = FORMATNUM$(SUMARBLOQUETOTALES#, "F12.4")
 
End Sub
Function COSTO_TOTAL#()
 COSTO_TOTAL# = XVALO(TXTCOINDBOLSON) + XVALO(TXTCOSTODIRBOLSON) + XVALO(Me.TXTSCRAP) '+ XVALO(TXTCOMISION)
 
End Function

Private Sub TXTCOSTODIRBOLSON_Change()
   
   Call TXTPORCENT_LostFocus(1)
   TXTUSS2(0) = FORMATNUM$(COSTO_TOTAL#, "F12.4")
  
End Sub
Sub porcentajeCostoManoObra()
    If XVALO(TXTCOSTOMOBRA) > 0 Then
           B = PORCENT_ENTRE_2_VALORES(XVALO(TXTUSS2(0)), XVALO(TXTCOSTOMOBRA))
           TXTLABEL(97) = FORMATNUM$(B, "F9.4")
       Else
          TXTLABEL(97) = 0
    End If

End Sub

Sub CalcularCostosIndirectos()
  COSTOMAT# = XVALO(TXTCOSTOMATERIAL)
     For i% = 0 To TXTPORCENT.UBound
       PORC = XVALO(TXTPORCENT(i%))
       TXTUSSUGERIDO(i%) = FORMATNUM$(COSTOMAT# * PORC / 100, "F12.4")
'       TXTUS(i%) = FORMATNUM$(COSTOMAT# * PORC / 100 , "F12.2") NO ACTUALIZA POR ESTE SE PUEDE EDITAR
     Next i%
End Sub

Sub CALCULARCOSTODIRECTOBOLSON()
    TXTCOSTODIRBOLSON = FORMATNUM$(XVALO(TXTCOSTOMOBRA) + XVALO(TXTCOSTOMATERIAL) + XVALO(Me.TXTCOSTOSFIJOS), "F12.4")
End Sub

Sub TXTCOSTOMATERIAL_Change()
    Call CalcularCostosIndirectos
    Call CALCULARCOSTODIRECTOBOLSON
    Call SUMARCOSTOSINDIRECTOS
    Call PORCENTAJECOSTOINDIRECTOBOLSON
    Call porcentajeCostoMaterial
    Call SUMARCOSTOSINDIRECTOSSUGERIDOS
    
    
End Sub
Function SUMARBLOQUETOTALES#()
   
   SUMARBLOQUETOTALES# = ROUND(XVALO(TXTUSS2(0)), 4) + ROUND(XVALO(TXTUSS2(1)), 4) + ROUND(XVALO(TXTCOMISION), 4)


End Function
Sub porcentajeCostoMaterial()
    If XVALO(TXTCOSTOMATERIAL) > 0 Then
           'B = XVALO(TXTUSS2(0)) / XVALO(TXTCOSTOMATERIAL)
           B = PORCENT_ENTRE_2_VALORES(XVALO(TXTUSS2(0)), XVALO(TXTCOSTOMATERIAL))
           TXTLABEL(95) = FORMATNUM$(B, "F9.4")
       Else
          TXTLABEL(95) = 0
    End If

End Sub
Private Sub TXTPORCENTAJEHILO_Change()
   Call CALCULARCOSTOMATERIALES
End Sub

Private Sub TXTPORCENTAJEHILO_LostFocus()
    TXTPORCENTAJEHILO = FORMATNUM$(XVALO(TXTPORCENTAJEHILO), "F5.2")
End Sub
Private Sub Text8_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Index < Text8.UBound Then
        Call PINTATEXT(Text8(Index + 1))
      Else
        Call PINTATEXT(Text8(0))
      End If
   End If
End Sub

Private Sub Text8_LostFocus(Index As Integer)
      Call VALORESCALUCLADOSBLOQUEETIQUETAS
      Call IMPORTESSBLOQUEETIQUETAS
      Call CALCULARCOSTOMATERIALES
      If Index <> 1 Then
        Text8(Index) = FORMATNUM$(XVALO(Text8(Index)), "F12.4")
      End If

End Sub
Private Sub Text9_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Index < Text9.UBound Then
        Call PINTATEXT(Text9(Index + 1))
      Else
        Call PINTATEXT(Text9(0))
      End If
   End If
End Sub


Private Sub Text9_LostFocus(Index As Integer)
      Call VALORESCALUCLADOSBLOQUEETIQUETAS
      Call IMPORTESSBLOQUEETIQUETAS
      Call CALCULARCOSTOMATERIALES
      Text9(Index) = FORMATNUM$(XVALO(Text9(Index)), "F12.0")


End Sub
Sub CARGATEXTBLOQUEETIQUETAS(CARGARDATOS%)
           
           If CARGARDATOS% > 0 Then
              FRMCOTI2.Text7(1) = COTICONT.TXTPDOCPORTADOCUMENTOTIPO
              FRMCOTI2.Text7(3) = XVALO(COTICONT.Text54) + XVALO(COTICONT.TXT_VPDFUELLELINERCU) * 2
              FRMCOTI2.Text7(4) = COTICONT.TXTEMBALAJEFARDO
              FRMCOTI2.Text8(1) = COTICONT.TXTDOCUMENTOUNIDAD
              FRMCOTI2.Text8(3) = XVALO(COTICONT.Text56)
              FRMCOTI2.Text9(1) = XVALO(COTICONT.TXTPORTADOCUMEDIDA)
              FRMCOTI2.Text10(3) = XVALO(COTICONT.TXTMICRONAJELINER)
              
              Call VALORESCALUCLADOSBLOQUEETIQUETAS
              Call IMPORTESSBLOQUEETIQUETAS
              SQLX$ = "SELECT ETIQUETA,LINER,ENVOLTORIO "
              SQLX$ = SQLX$ + " FROM CONFIGBB WITH(NOLOCK)  WHERE CODBIGBAG = '" & COTICONT.TXTCODCONFECCION & "'"
              Set RS = READSET(SQLX$)
              With RS
                If Not .EOF Then
                    FRMCOTI2.Text9(0) = SAFETEXT$(!ETIQUETA)
                    FRMCOTI2.Text9(3) = SAFETEXT$(!LINER)
                    FRMCOTI2.Text9(4) = SAFETEXT$(!ENVOLTORIO)
                End If
              End With
              RS.Close: Set RS = Nothing
              
           End If
End Sub
Sub VALORESCALUCLADOSBLOQUEETIQUETAS()
   Resultadox# = (XVALO(FRMCOTI2.Text7(3)) * XVALO(FRMCOTI2.Text8(3)) * XVALO(FRMCOTI2.Text9(3)) * XVALO(FRMCOTI2.Text10(3))) / 1000000 * 0.184
   FRMCOTI2.Text11(3) = TRIM$(FORMATNUM$(Resultadox#, "F12.4"))
   
End Sub

Private Sub Text11_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Index < Text11.UBound Then
        Call PINTATEXT(Text11(Index + 1))
      Else
        Call PINTATEXT(Text11(0))
      End If
   End If
End Sub

Private Sub Text11_LostFocus(Index As Integer)
      Call VALORESCALUCLADOSBLOQUEETIQUETAS
      Call IMPORTESSBLOQUEETIQUETAS
      Call CALCULARCOSTOMATERIALES
      If Index <> 3 Then Text11(Index) = FORMATNUM$(XVALO(Text11(Index)), "F12.4")
End Sub
Private Sub Text12_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Index < Text12.UBound Then
        Call PINTATEXT(Text12(Index + 1))
      Else
        Call PINTATEXT(Text12(0))
      End If
   End If
End Sub

Private Sub Text12_LostFocus(Index As Integer)
      Call VALORESCALUCLADOSBLOQUEETIQUETAS
      Call IMPORTESSBLOQUEETIQUETAS
      Call CALCULARCOSTOMATERIALES
      Text12(Index) = FORMATNUM$(XVALO(Text12(Index)), "F12.4")

End Sub

Private Sub Text12_Change(Index As Integer)
   MODIFICCOT2% = 1
End Sub
Private Sub Text13_LostFocus(Index As Integer)
      Call VALORESCALUCLADOSBLOQUEETIQUETAS
      Call IMPORTESSBLOQUEETIQUETAS
      Call CALCULARCOSTOMATERIALES
      
      Text13(Index) = FORMATNUM$(XVALO(Text13(Index)), "F12.4")
  

End Sub
Private Sub Text13_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Index < Text13.UBound Then
        Call PINTATEXT(Text13(Index + 1))
      Else
        Call PINTATEXT(Text13(0))
      End If
   End If
   MODIFICCOT2% = 1
End Sub
Sub IMPORTESSBLOQUEETIQUETAS()
    FRMCOTI2.Text13(0) = TRIM$(FORMATNUM$(XVALO(FRMCOTI2.Text9(0)) * XVALO(FRMCOTI2.Text12(0)), "F12.4"))
    FRMCOTI2.Text13(1) = TRIM$(FORMATNUM$(XVALO(FRMCOTI2.Text9(1)) * XVALO(FRMCOTI2.Text12(1)), "F12.4"))
    FRMCOTI2.Text13(3) = TRIM$(FORMATNUM$(XVALO(FRMCOTI2.Text11(3)) * XVALO(FRMCOTI2.Text12(3)), "F12.4"))
    If XVALO(FRMCOTI2.Text7(4)) > 0 Then
      FRMCOTI2.Text13(4) = TRIM$(FORMATNUM$(XVALO(FRMCOTI2.Text12(4)) / XVALO(FRMCOTI2.Text7(4)), "F12.4"))
    Else
       FRMCOTI2.Text13(4) = 0
    End If
End Sub

Sub CALCULARCOSTOMATERIALES()
    ACUIMPO = 0
    For i% = 1 To COTICONT.MSF1.Rows - 1
       IMPO = XVALO(COTICONT.MSF1.TextMatrix(i%, 9))
       ACUIMPO = ACUIMPO + IMPO
    Next i%
    For i% = 1 To COTICONT.MSF2.Rows - 1
       IMPO = XVALO(COTICONT.MSF2.TextMatrix(i%, 9))
       ACUIMPO = ACUIMPO + IMPO
    Next i%
    For i% = 1 To COTICONT.MSF3.Rows - 1
       IMPO = XVALO(COTICONT.MSF3.TextMatrix(i%, 9))
       ACUIMPO = ACUIMPO + IMPO
    Next i%
    For i% = 1 To COTICONT.MSF4.Rows - 1
       IMPO = XVALO(COTICONT.MSF4.TextMatrix(i%, 9))
       ACUIMPO = ACUIMPO + IMPO
    Next i%
    
    For i% = 1 To COTICONT.MSF5.Rows - 1
       IMPO = XVALO(COTICONT.MSF5.TextMatrix(i%, 9))
       ACUIMPO = ACUIMPO + IMPO
    Next i%
    
    COTICONT.TXTSTOTSOLAPACOTIZACION = FORMATNUM$(XVALO(ACUIMPO), "F12.4") 'este es un subtotal en solapa de cotizacion la primera
    '
    For i% = 0 To FRMCOTI2.Text13.UBound
      IMPO = XVALO(FRMCOTI2.Text13(i%))
      ACUIMPO = ACUIMPO + IMPO
    Next i%
    
    ACUIMPO = ACUIMPO * (1 + XVALO(Me.TXTPORCENTAJEHILO) / 100)
    Me.TXTCOSTOMATERIAL = FORMATNUM$(XVALO(ACUIMPO), "F15.4") '+ ROUND(XVALO(TXTCOSTOMOBRA), 2)
    
End Sub


Private Sub Command49_Click()
   '******VALIDAR SI LA MISMA YA ESTA APROBADA
  Command49.Enabled = False
  CONDI$ = "NROCOTIZACION = " & XVALO(Text14)
   If CantRegistros&("COTIZACTP", CONDI$, "NOMES") < 1 Then
      Call COMUNI("La Cotización Activa no ha Sido Grabada\Deberá Grabar la Misma para Luego Aprobarla")
      GoTo 99
   End If
   CANTCONFIRM& = ROUND(XVALO(TXTCANTSOLIC), 0)
   If CANTCONFIRM& < 1 Then
      CANTSOLIC& = XVALO(VALOBADA("SOLICOTI", "CANTIDAD", "solicoti_id = " & XVALO(COTICONT.LBLNROSOLICITUDCOTIZACION), "NOMESS"))
      If CANTSOLIC& > 0 Then
        OPX% = COMALTER%("Falta Ingresar La Cantidad Confirmada Para Esta Cotización\\Cancelar\Aprobar con Cantidad Solicitada: " & CANTSOLIC&)
        If OPX% < 2 Then
           GoTo 99
        Else
           TXTCANTSOLIC = CANTSOLIC&
           CANTCONFIRM& = TXTCANTSOLIC
        End If
      Else
        Call COMUNI("Falta Ingresar Cantidad")
        GoTo 99
      End If
   End If
   
   CONDI$ = CONDI$ + " AND STATUS = 1 "
   If CantRegistros&("COTIZACTP", CONDI$, "NOMES") > 0 Then
      Call COMUNI("Cotización Activa Aprobada\Imposible Aprobar !!")
      GoTo 99
   End If
   
   FIRMAELEC.TEXT1 = ""
   FIRMAELEC.Text2 = ""
   FIRMAELEC.Label1 = ""
   If XVALO(Text14) < 1 Then
      Call COMUNI("Debe Haber una Cotización Activa")
      GoTo 99
   End If

   FIRMAELEC.Show 1

   If FIRMAELEC.ESTADO < 1 Then
      GoTo 99
   Else
    'TRANSACCIÓN GRABACION
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans
    PorComisReal# = XVALO(TXTPORCENT2(4))
    
    CONDI$ = "NROCOTIZACION = " & XVALO(Text14)
    'ACTUALIZA EL ENCABEZADO DE COTIZACIONES CON STATUS = 1 APROBADA
    Call ACTUREGISTRO("COTIZACTP", "STATUS", CONDI$, 1, REGAF&, "NOMESS")
    Call ACTUREGISTRO("SOLICOTI", "COTIZAR", CONDI$, 3, REGAF&, "NOMESS")
    SQLX$ = "SELECT * FROM SOLICOTI WITH(NOLOCK) WHERE solicoti_id = " & XVALO(COTICONT.LBLNROSOLICITUDCOTIZACION)
    Set RS = READSET(SQLX$)
    With RS
       NC = XVALO(!NUCLIEN)
'       CANTI = XVALO(!CANTIDAD) 'SE GENERA CON LA CANTIDAD CONFIRMADA NO POR LA SOLICITADA POR QUE PUEDE VARIAR
       Codigo$ = SAFETEXT$(!CODIMASTER)
       CODBIG$ = SAFETEXT$(!CODBIGBAG)
       CI1% = XVALO(!CODINT)
       If CI1% < 1 Then CI1% = CODINT(Codigo$)

       NVEND% = XVALO(!NUVEND)
       NROFICHATEC& = XVALO(!NUMRECETACTP)
    End With
    RS.Close
    Set RS = Nothing
    'ESTO ES POR SI ES POR COPIA LA COTIZACION Y NO ENTRA EN EL RECORSET ANTERIOR POR QUE NO VIENE DE UNA SOLICITUD DE COTIZACION
    If XVALO(NC) < 1 Then
        NC = XVALO(COTICONT.TEXT1)
        Codigo$ = TRIM$(COTICONT.TXTCODIGO)
        CODBIG$ = TRIM$(COTICONT.TXTCODCONFECCION)
        CI1% = CODINT(Codigo$)
        NVEND% = XVALO(COTICONT.Text4)
        NROFICHATEC& = XVALO(COTICONT.Text5)
    End If
    
    'Call GRABARENPRESUVEN(NC, XVALO(Text14), CANTI, CI1%, NVEND%, Codigo$, CODBIG$, NROFICHATEC&, 2, XVALO(TXTUSS2(3)), Text1, OKX%)
    'SE CAMBIA XVALO(TXTUSS2(3)) POR XVALO(PREVENTAFINAL) QUE CONSIDERA EL LINER AGREGADO AL FINAL
     Call GRABARENPRESUVEN(NC, XVALO(Text14), 1, CI1%, NVEND%, Codigo$, CODBIG$, NROFICHATEC&, 2, XVALO(TXTPRECIOFINAL), TEXT1, OKX%, 1, PorComisReal#)
    'SIEMPRE SE ACTUALIZA CUANDO SE APRUEBA Y NO ESTA MAL, POR QUE SI SE DECIDE ABRIR LA POSIBILIDAD
    'DE MODIFICAR COTI1 SE ACTUALIZA Y SE QUEDA CON LA ULTIMA.
    Call ACTUALIZA_ESTRUCTURA_CONTEPLAST(CI1%)
'    Call AGREGA_ROTULO_CONTEPLAST(CI1%)
    If OKX% < 0 Then GoTo 70
ERROR_GUARDAR:
    If Err <> 0 Then
70     FLEXCONN.RollbackTrans

       On Error GoTo 0
       MsgBox Err.Description
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       GoTo 99
    Else
      FLEXCONN.CommitTrans
      
      Call COMUNI("Grabación Realizado Exitosamente")
    End If

   End If
   
99 Command49.Enabled = True
End Sub
Sub ACTUALIZA_ESTRUCTURA_CONTEPLAST(CI1%)
'OJO ESTA REPLICADA EN AMASTOCK07
   '
   List2.Clear
   If CI1% < 1 Or CI1% > NUMAS% Then Exit Sub
   '
   If XVALO(COTICONT.TXTNUMCOTIZACION) < 1 Then Exit Sub
   '
   CONDI$ = "CODICONJ = " & CI1%
   NUORIT% = 0
   Call ACTUREGISTRO("FORMULAS", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
    SQLX$ = "SELECT * FROM COTI1 WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NROCOTIZACION  = " & XVALO(COTICONT.TXTNUMCOTIZACION)
    SQLX$ = SQLX$ + " and codigo <> ''"
    SQLX$ = SQLX$ + " and ABS(CODINTELEM) > 0 and MODUXBB > 0 and CORTE > 0 "

    'SQLX$ = SQLX$ + " AND NAMEMSF = '" & MSF.Name & "' "
    SQLX$ = SQLX$ + " ORDER BY COTI1_ID"
    Set RS = READSET(SQLX$)
        With RS
          Do While Not .EOF
            Codigo$ = SAFETEXT$(!Codigo)
            CI2% = CODINT(Codigo$)
            
            EsCintilla% = 0
            CODFAMI$ = TRIM$(CODFAMIL$(CI2%)) 'para cintillas (si es cintilla fabricada son unidades)
            If InStr(CODFAMI$, "CC") > 0 And Unimed$(CI2%) = "U." Then
               EsCintilla% = 1
            End If
               
            CIRANG = CODIGORANGO&(CI2%)
            If CIRANG > 0 Then CI2% = CIRANG 'LE ASIGNA EL CODIGO RANGO (EL CODIGO RANGO DEBE TENER MISMA U.MEDIDA
            CANTIDAD_CARAS = XVALO(!moduxbb)
            CORTEX = XVALO(!corte) 'ES EN CM.
            USO = (CORTEX * CANTIDAD_CARAS)
            
            If EsCintilla% > 0 Then USO = CANTIDAD_CARAS  'Cantidad de Unidades
            
            UNXX$ = Unimed$(CI2%)
            If UNXX$ = "M." Then USO = USO / 100  'LO PASO A MTS
            USOX$ = FORMATNUM$(XVALO(USO), "F12.4")
'            DESCRIX$ = SAFETEXT$(!Descripcion)
'            Codigo$ = SAFETEXT$(!Codigo)
            DESCRIX$ = DESCRIT0$(CI2%) 'LO LEO DE LA FUNCION POR SI ES RANGO CAMBIA
            Codigo$ = CODEXT$(CI2%)    'LO LEO DE LA FUNCION POR SI ES RANGO CAMBIA
            PDX$ = 0
            COMPONENTEX = SAFETEXT$(!COMPONENTE)
            MODUXBBX = SAFETEXT$(!moduxbb)
            CORTEX = XVALO(!corte)
            HastaParentesis% = InStr(COMPONENTEX, "(")
            If HastaParentesis% > 0 Then
               COMPONENTEX = TRIM$(Mid(COMPONENTEX, 1, HastaParentesis% - 1))
            End If
            
            If USO > 0.00005 Then
                SQLX$ = "SELECT * FROM Formulas WHERE CodIConj < 1"
                'GUARDA EN TABLA CABEZERA COTIZACTP
                Dim RS1 As ADODB.Recordset
                Set RS1 = New ADODB.Recordset
                RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                With RS1
                  .AddNew
                  !CODICONJ = CI1%
                  !CODIEMPR = CodiEmp%
                  !CODXCONJ = Left(CODEXT$(CI1%), 16)
                  !codielem = CI2%
                  !CodXElem = Codigo$
                  NUORIT% = NUORIT% + 1
                  !NUORIT = NUORIT%
                  !usobruto = XVALO(USOX$)
                  !UNIS_FORM = XVALO(USOX$)
                  !UMED_FORM = UNXX$
                  !LAR_CORTE = 0
                  !PLADISPO = 0
                  !RefeComp = ""
                  !MARBORRA = 0
                  !COMPONENTE = COMPONENTEX
                  !moduxbb = MODUXBBX
                  !corte = CORTEX
                  .Update
                End With
            RS1.Close
            Set RS1 = Nothing
            End If
            .MoveNext
          Loop
    End With
    RS.Close
    Set RS = Nothing
    '
    Call AGREGARLINER(CI1%, COTICONT.TXTNUMCOTIZACION, NUORIT%)
    '
    
   CONDI$ = CONDI$ + " AND MARBORRA = 1"
   Call BORRAREGISTROS("FORMULAS", CONDI$, REGAF&, "NOMESS")
   
   Dim CTP As New Conteplast
   Call CTP.AgregarOperacionesSegúnCódigo(CI1%, XVALO(COTICONT.TXTNUMCOTIZACION)) 'GENERA LAS OPERACIONES

   '
   
99 Screen.MousePointer = 1
End Sub

Private Sub Command50_Click()
   Command50.Enabled = False
   If XVALO(Text14) > 0 Then
      If VALIDASISEMODIFICO% < 0 Then Exit Sub
      If CantRegistros&("COTIZACTP", "NROCOTIZACION = " & XVALO(Text14)) < 1 Then
         Call COMUNI("Debe Seleccionar Una Cotización Que esté Correctamente Guardada en al Base de Datos")
         GoTo 99
      End If
      
      OPX% = COMALTER%("Confirmación de Grabación de Nueva Cotización en Base a la Cotización Activa\\Cancelar\Si Grabar")
      If OPX% > 1 Then
'         Call Command46_Click
'         Call Command47_Click
         NUEVONUM = NumNuevaCotiza 'EL NRO DE COTIZACION QUE SIGUE
         On Error GoTo ERROR_GUARDAR
         FLEXCONN.BeginTrans
         'DUPLICA EL REGISTRO, PERO ANTES TOMA EL NRO DE ID MAXIMO SEGUEN EL NRO DE COTIZACION
         MAXIMOACTUALID& = XVALO(VALOBADA("COTIZACTP", "MAX(COTIZACTP_ID)", "NROCOTIZACION = " & XVALO(Text14), "NOMESS"))
         
    SQLX$ = "INSERT INTO COTIZACTP " & " (" & CamposdeUnaTabla$("COTIZACTP") & ")"
    SQLX$ = SQLX$ + " SELECT " & CamposdeUnaTabla$("COTIZACTP", 0)
    SQLX$ = SQLX$ + " From COTIZACTP Where NROCOTIZACION = " & XVALO(Text14)
    FLEXCONN.Execute SQLX$
         
'         SQLX$ = "INSERT INTO COTIZACTP(CodiEmpr"
'         SQLX$ = SQLX$ + ",NROCOTIZACION"
'         SQLX$ = SQLX$ + ",FECHEMISI"
'         SQLX$ = SQLX$ + ",CODIMASTER,   NUCLIEN, NUMRECETACTP, RASOCLI, STATUS)"
'         SQLX$ = SQLX$ + " SELECT CodiEmpr,NROCOTIZACION,FECHEMISI, CODIMASTER,   NUCLIEN,"
'         SQLX$ = SQLX$ + "NUMRECETACTP, RASOCLI, STATUS FROM COTIZACTP WHERE NROCOTIZACION = " & XVALO(Text14)
'         FLEXCONN.Execute SQLX$
         '
         'LE CAMBIA SOLO EL NRO DE COTIZACION A LOS QUE COINCIDE EL NROCOTIZCION Y POSEEN ID > AL MAXIMO LEIDO
         SQLX$ = "UPDATE COTIZACTP SET NROCOTIZACION = " & NUEVONUM & ", STATUS = 0 ,FECHEMISI = '" & Date & "' WHERE COTIZACTP_ID > " & MAXIMOACTUALID& & " AND NROCOTIZACION = " & XVALO(Text14)
         FLEXCONN.Execute SQLX$
         '
         
         'DUPLICA EL REGISTRO, PERO ANTES TOMA EL NRO DE ID MAXIMO SEGUEN EL NRO DE COTIZACION
         MAXIMOACTUALID1& = XVALO(VALOBADA("COTI1", "MAX(COTI1_ID)", "NROCOTIZACION = " & XVALO(Text14), "NOMESS"))
    SQLX$ = "INSERT INTO COTI1 " & " (" & CamposdeUnaTabla$("COTI1") & ")"
    SQLX$ = SQLX$ + " SELECT " & CamposdeUnaTabla$("COTI1", 0)
    SQLX$ = SQLX$ + " From COTI1 Where NROCOTIZACION = " & XVALO(Text14)
    FLEXCONN.Execute SQLX$
        
         
'         SQLX$ = "INSERT INTO COTI1(CodiEmpr"
'         SQLX$ = SQLX$ + ",NROCOTIZACION"
'         SQLX$ = SQLX$ + ", CORTE, Codigo, COMPONENTE,  F_TECNICA, FECHEMISI, MARBORRA,MODUXBB, NAMEMSF, ORDEN, STATUS, PRECIOXM)"
'
'         SQLX$ = SQLX$ + " SELECT CodiEmpr, NROCOTIZACION, CORTE, Codigo, COMPONENTE,  F_TECNICA, FECHEMISI, MARBORRA, MODUXBB, NAMEMSF, ORDEN, STATUS, PRECIOXM "
'         SQLX$ = SQLX$ + " FROM COTI1 WHERE NROCOTIZACION = " & XVALO(Text14)
'         FLEXCONN.Execute SQLX$
         'LE CAMBIA SOLO EL NRO DE COTIZACION A LOS QUE COINCIDE EL NROCOTIZCION Y POSEEN ID > AL MAXIMO LEIDO
         SQLX$ = "UPDATE COTI1 SET NROCOTIZACION = " & NUEVONUM & " WHERE COTI1_ID > " & MAXIMOACTUALID1& & " AND NROCOTIZACION = " & XVALO(Text14)
         FLEXCONN.Execute SQLX$

         MAXIMOACTUALID2& = XVALO(VALOBADA("COTI2", "MAX(COTI2_ID)", "NROCOTIZACION = " & XVALO(Text14), "NOMESS"))
    
    SQLX$ = "INSERT INTO COTI2 " & " (" & CamposdeUnaTabla$("COTI2") & ")"
    SQLX$ = SQLX$ + " SELECT " & CamposdeUnaTabla$("COTI2", 0)
    SQLX$ = SQLX$ + " From COTI2  WHERE NROCOTIZACION = " & XVALO(Text14)
    FLEXCONN.Execute SQLX$
         
         'DUPLICA EL REGISTRO, PERO ANTES TOMA EL NRO DE ID MAXIMO SEGUEN EL NRO DE COTIZACION
'         SQLX$ = "INSERT INTO COTI2(CodiEmpr"
'         SQLX$ = SQLX$ + ",NROCOTIZACION"
'         SQLX$ = SQLX$ + ",COMISION, COSTMOBRA, COSTTOT,  ENVO_COSTO, ENVO_EMBF"
'         SQLX$ = SQLX$ + ",ENVO_MOD, ESTRUCTURA, ETIQ_COSTO, ETIQ_MODULO, FECHEMISI, FECHENTRE, FORMAENTREGA, GASTFINAN"
'         SQLX$ = SQLX$ + ",STATUS, SOBRE_PDOCMED, SOBRE_MODULO, SOBRE_DOCTIPO, SOBRE_DOCMED, SOBRE_COSTO, SCRAPP, PREVENTAREAL"
'         SQLX$ = SQLX$ + ",PREVENTA, PJEUTILIDAD, OTROS, OTRO2, OTRO1, MODOPAGO, MICRONALJELINER, MARBORRA, LOGYFLETE, LINER_MODULO"
'         SQLX$ = SQLX$ + ",LINER_KG, LINER_COSTO, LINER_ANCHO, LINER_ALTO, LIBRE3TXT, LIBRE3IMPO, LIBRE2TXT, LIBRE2IMPO, LIBRE1TXT"
'         SQLX$ = SQLX$ + ",LIBRE1IMPO, IMPUESTO, IMPRE_MODULO, IMPRE_COSTO,CANTICONFIRMADA,PORCENTCOMI,IMPOLINER,PREVENTAFINAL)"
'
'         SQLX$ = SQLX$ + " SELECT CodiEmpr, NROCOTIZACION, COMISION, COSTMOBRA, COSTTOT,  ENVO_COSTO, ENVO_EMBF,"
'         SQLX$ = SQLX$ + " ENVO_MOD, ESTRUCTURA, ETIQ_COSTO, ETIQ_MODULO, FECHEMISI, FECHENTRE, FORMAENTREGA, GASTFINAN,"
'         SQLX$ = SQLX$ + " STATUS, SOBRE_PDOCMED, SOBRE_MODULO, SOBRE_DOCTIPO, SOBRE_DOCMED, SOBRE_COSTO, SCRAPP, PREVENTAREAL,"
'         SQLX$ = SQLX$ + " PREVENTA, PJEUTILIDAD, OTROS, OTRO2, OTRO1, MODOPAGO, MICRONALJELINER, MARBORRA, LOGYFLETE, LINER_MODULO,"
'         SQLX$ = SQLX$ + " LINER_KG, LINER_COSTO, LINER_ANCHO, LINER_ALTO, LIBRE3TXT, LIBRE3IMPO, LIBRE2TXT, LIBRE2IMPO, LIBRE1TXT,"
'         SQLX$ = SQLX$ + " LIBRE1IMPO, IMPUESTO, IMPRE_MODULO, IMPRE_COSTO,CANTICONFIRMADA,PORCENTCOMI,IMPOLINER,PREVENTAFINAL FROM COTI2 WHERE NROCOTIZACION = " & XVALO(Text14)
'         FLEXCONN.Execute SQLX$

         'LE CAMBIA SOLO EL NRO DE COTIZACION A LOS QUE COINCIDE EL NROCOTIZCION Y POSEEN ID > AL MAXIMO LEIDO
         SQLX$ = "UPDATE COTI2 SET NROCOTIZACION = " & NUEVONUM & " WHERE COTI2_ID > " & MAXIMOACTUALID2& & " AND NROCOTIZACION = " & XVALO(Text14)
         FLEXCONN.Execute SQLX$
        
ERROR_GUARDAR:
        If Err <> 0 Then
           FLEXCONN.RollbackTrans

           On Error GoTo 0
           Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
           GoTo 99
        Else
          FLEXCONN.CommitTrans
          Call COMUNI("Grabación Realizado Exitosamente")
          Text14 = NUEVONUM
          COTICONT.TXTNUMCOTIZACION = NUEVONUM
        End If
        MODIFIC% = 0
      End If
   End If
99 Command50.Enabled = True
End Sub


Private Sub Text10_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Index < Text10.UBound Then
        Call PINTATEXT(Text10(Index + 1))
      Else
        Call PINTATEXT(Text10(0))
      End If
   End If
   MODIFICCOT2% = 1
End Sub
Private Sub Text10_LostFocus(Index As Integer)
      Call VALORESCALUCLADOSBLOQUEETIQUETAS
      Call IMPORTESSBLOQUEETIQUETAS
      Call CALCULARCOSTOMATERIALES
      Text10(Index) = FORMATNUM$(XVALO(Text10(Index)), "F12.4")


End Sub




Private Sub Text7_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      If Index < Text7.UBound Then
        Call PINTATEXT(Text7(Index + 1))
      Else
        Call PINTATEXT(Text7(0))
      End If
   End If
End Sub

Private Sub Text7_LostFocus(Index As Integer)
      Call VALORESCALUCLADOSBLOQUEETIQUETAS
      Call IMPORTESSBLOQUEETIQUETAS
      Call CALCULARCOSTOMATERIALES
      If Index <> 1 Then
        Text7(Index) = FORMATNUM$(XVALO(Text7(Index)), "F12.4")
      End If
End Sub

Private Sub TXTVALORLINER_Change()
    PREVTAREAL = XVALO(TXTUSS2(3))
    TXTPRECIOFINAL = FORMATNUM$(PREVTAREAL + XVALO(Me.TXTVALORLINER), "F12.4")
End Sub

Private Sub TXTVALORLINER_LostFocus()
   TXTVALORLINER = FORMATNUM$(XVALO(TXTVALORLINER), "F12.4")
End Sub

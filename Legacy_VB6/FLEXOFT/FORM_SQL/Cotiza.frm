VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Cotiza 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cotización "
   ClientHeight    =   9645
   ClientLeft      =   11400
   ClientTop       =   1095
   ClientWidth     =   15720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   15720
   Begin VB.TextBox TXTALERTA 
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
      ForeColor       =   &H000000FF&
      Height          =   1050
      Left            =   7200
      MaxLength       =   14
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   204
      TabStop         =   0   'False
      Text            =   "Cotiza.frx":0000
      Top             =   1560
      Visible         =   0   'False
      Width           =   1620
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   6000
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Observaciones para Panelización"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   15
      Left            =   11520
      TabIndex        =   194
      Top             =   8640
      Width           =   4125
      Begin VB.TextBox TXTObservPenal 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         MaxLength       =   2046
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   44
         Top             =   240
         Width           =   3885
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "OC del Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   14
      Left            =   13200
      TabIndex        =   193
      Top             =   6600
      Width           =   2415
      Begin VB.TextBox TXTNROOCOM 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         MaxLength       =   12
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   360
         Width           =   2180
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Observaciones al Cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   980
      Index           =   13
      Left            =   7080
      TabIndex        =   192
      Top             =   8640
      Width           =   4335
      Begin VB.TextBox TXTObservCliente 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   120
         MaxLength       =   2046
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Top             =   240
         Width           =   4125
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Observaciones a Planta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1040
      Index           =   12
      Left            =   7080
      TabIndex        =   191
      Top             =   7580
      Width           =   8535
      Begin VB.TextBox TXTObservaPlanta 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   120
         MaxLength       =   2046
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   42
         Top             =   240
         Width           =   8325
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Observaciones de Administración"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   11
      Left            =   7080
      TabIndex        =   190
      Top             =   6580
      Width           =   6015
      Begin VB.TextBox TXTObservaAdmin 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   680
         Left            =   120
         MaxLength       =   2046
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   240
         Width           =   5805
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Resumen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5240
      Index           =   10
      Left            =   7080
      TabIndex        =   156
      Top             =   1300
      Width           =   8535
      Begin VB.TextBox TXTTOTALSTENCIL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   5035
         Locked          =   -1  'True
         TabIndex        =   222
         TabStop         =   0   'False
         Top             =   4005
         Width           =   965
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   199
         TabStop         =   0   'False
         Top             =   225
         Width           =   480
      End
      Begin VB.CommandButton Command7 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2920
         TabIndex        =   198
         Top             =   225
         Width           =   175
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   197
         TabStop         =   0   'False
         Text            =   "Costo"
         Top             =   600
         Width           =   2000
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4000
         Locked          =   -1  'True
         TabIndex        =   196
         TabStop         =   0   'False
         Text            =   "Normal"
         Top             =   600
         Width           =   2000
      End
      Begin VB.TextBox TXTStTotalusd 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   4450
         MaxLength       =   4
         TabIndex        =   38
         Top             =   4005
         Width           =   555
      End
      Begin VB.TextBox TXTCantStencil 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   4005
         MaxLength       =   1
         TabIndex        =   37
         Top             =   4005
         Width           =   435
      End
      Begin VB.TextBox Text34 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   4005
         Locked          =   -1  'True
         TabIndex        =   189
         TabStop         =   0   'False
         Top             =   4840
         Width           =   1080
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   10
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   188
         TabStop         =   0   'False
         Text            =   "Saldo"
         Top             =   4840
         Width           =   750
      End
      Begin VB.TextBox Text33 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   4840
         Width           =   1080
      End
      Begin VB.TextBox TXTTtOpcionCosto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   187
         TabStop         =   0   'False
         Top             =   4360
         Width           =   2000
      End
      Begin VB.TextBox TXTPrcTtPlcaSACosto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   186
         TabStop         =   0   'False
         Top             =   3705
         Width           =   2000
      End
      Begin VB.TextBox TXTPreunit_CpteCosto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   185
         TabStop         =   0   'False
         Top             =   2805
         Width           =   2000
      End
      Begin VB.TextBox TXTPreunit_ArmCosto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   184
         TabStop         =   0   'False
         Top             =   2505
         Width           =   2000
      End
      Begin VB.TextBox TXTPrcTtOpcionCosto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   183
         TabStop         =   0   'False
         Top             =   1905
         Width           =   2000
      End
      Begin VB.TextBox TXTtotPlcaCosto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   182
         TabStop         =   0   'False
         Top             =   1605
         Width           =   2000
      End
      Begin VB.TextBox TxtPrcUnitCosto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   181
         TabStop         =   0   'False
         Top             =   1305
         Width           =   2000
      End
      Begin VB.TextBox TxtPrcdm2Costo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   180
         TabStop         =   0   'False
         Top             =   1005
         Width           =   2000
      End
      Begin VB.TextBox TXTTtOpcion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4005
         Locked          =   -1  'True
         TabIndex        =   179
         TabStop         =   0   'False
         Top             =   4360
         Width           =   2000
      End
      Begin VB.TextBox TXTPrcTtPlcaSA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   4000
         Locked          =   -1  'True
         TabIndex        =   178
         TabStop         =   0   'False
         Top             =   3705
         Width           =   2000
      End
      Begin VB.TextBox TXTPreunit_Cpte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   4000
         TabIndex        =   34
         Top             =   2805
         Width           =   2000
      End
      Begin VB.TextBox TXTPreunit_Arm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   4000
         TabIndex        =   33
         Top             =   2505
         Width           =   2000
      End
      Begin VB.TextBox TXTPreunitCto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   4000
         TabIndex        =   32
         Top             =   2205
         Width           =   2000
      End
      Begin VB.TextBox TXTPrcTtOpcion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   4000
         Locked          =   -1  'True
         TabIndex        =   177
         TabStop         =   0   'False
         Top             =   1905
         Width           =   2000
      End
      Begin VB.TextBox TXTtotPlca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   4000
         Locked          =   -1  'True
         TabIndex        =   176
         TabStop         =   0   'False
         Top             =   1605
         Width           =   2000
      End
      Begin VB.TextBox TxtPrcUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   4000
         Locked          =   -1  'True
         TabIndex        =   175
         TabStop         =   0   'False
         Top             =   1305
         Width           =   2000
      End
      Begin VB.TextBox TxtPrcdm2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   4000
         Locked          =   -1  'True
         TabIndex        =   174
         TabStop         =   0   'False
         Top             =   1005
         Width           =   2000
      End
      Begin VB.TextBox TxtSup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   9
         Left            =   100
         Locked          =   -1  'True
         TabIndex        =   173
         TabStop         =   0   'False
         Text            =   "PRECIO TOTAL OPERACIÓN "
         Top             =   4360
         Width           =   3500
      End
      Begin VB.TextBox TxtSup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   8
         Left            =   100
         Locked          =   -1  'True
         TabIndex        =   172
         TabStop         =   0   'False
         Text            =   "Precio Total Placas S/Armado "
         Top             =   3705
         Width           =   3500
      End
      Begin VB.TextBox TxtSup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   7
         Left            =   100
         Locked          =   -1  'True
         TabIndex        =   171
         TabStop         =   0   'False
         Text            =   "Precio Unitario de Componentes "
         Top             =   2805
         Width           =   3500
      End
      Begin VB.TextBox TxtSup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   6
         Left            =   100
         Locked          =   -1  'True
         TabIndex        =   170
         TabStop         =   0   'False
         Text            =   "Precio Unitario de Armado "
         Top             =   2505
         Width           =   3500
      End
      Begin VB.TextBox TxtSup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   5
         Left            =   100
         Locked          =   -1  'True
         TabIndex        =   169
         TabStop         =   0   'False
         Text            =   "Precio Unitario del Circuito "
         Top             =   2205
         Width           =   3500
      End
      Begin VB.TextBox TxtSup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   4
         Left            =   100
         Locked          =   -1  'True
         TabIndex        =   168
         TabStop         =   0   'False
         Text            =   "Precio Total Operación "
         Top             =   1905
         Width           =   3500
      End
      Begin VB.TextBox TxtSup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   3
         Left            =   100
         Locked          =   -1  'True
         TabIndex        =   167
         TabStop         =   0   'False
         Text            =   "Precio Total Placas "
         Top             =   1605
         Width           =   3500
      End
      Begin VB.TextBox TxtSup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   2
         Left            =   100
         Locked          =   -1  'True
         TabIndex        =   166
         TabStop         =   0   'False
         Text            =   "Precio Unitario "
         Top             =   1305
         Width           =   3500
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   9
         Left            =   100
         Locked          =   -1  'True
         TabIndex        =   165
         TabStop         =   0   'False
         Text            =   "Diseño "
         Top             =   3405
         Width           =   3500
      End
      Begin VB.TextBox Text32 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   4000
         TabIndex        =   36
         Top             =   3405
         Width           =   2000
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   8
         Left            =   100
         Locked          =   -1  'True
         TabIndex        =   164
         TabStop         =   0   'False
         Text            =   "STENCIL (Cant./P.Unit./TotalL USD) "
         Top             =   4005
         Width           =   3500
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   7
         Left            =   100
         Locked          =   -1  'True
         TabIndex        =   163
         TabStop         =   0   'False
         Text            =   "Anticipo"
         Top             =   4840
         Width           =   750
      End
      Begin VB.TextBox TxtSup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   1
         Left            =   100
         Locked          =   -1  'True
         TabIndex        =   162
         TabStop         =   0   'False
         Text            =   "Precio dm2 "
         Top             =   1005
         Width           =   3500
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   6
         Left            =   100
         Locked          =   -1  'True
         TabIndex        =   161
         TabStop         =   0   'False
         Text            =   "Precio Fijo x dm2"
         Top             =   600
         Width           =   2200
      End
      Begin VB.TextBox Text31 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   160
         TabStop         =   0   'False
         Top             =   600
         Width           =   1200
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   5
         Left            =   100
         Locked          =   -1  'True
         TabIndex        =   159
         TabStop         =   0   'False
         Text            =   "Gastos Preparatorios "
         Top             =   3105
         Width           =   3500
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   4
         Left            =   100
         Locked          =   -1  'True
         TabIndex        =   158
         TabStop         =   0   'False
         Text            =   "F.Pago "
         Top             =   225
         Width           =   2200
      End
      Begin VB.TextBox Text30 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   157
         TabStop         =   0   'False
         Top             =   225
         Width           =   5160
      End
      Begin VB.TextBox Text26 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   4000
         TabIndex        =   35
         Top             =   3105
         Width           =   2000
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Último Modo Cálculo:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5160
         TabIndex        =   221
         Top             =   4800
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Montaje"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1900
      Index           =   9
      Left            =   2880
      TabIndex        =   147
      Top             =   7710
      Width           =   4095
      Begin VB.TextBox TXTINSTOP3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   3000
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   155
         TabStop         =   0   'False
         Top             =   1140
         Width           =   840
      End
      Begin VB.TextBox TXTTOTALCOMPONENTES 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   3000
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   154
         TabStop         =   0   'False
         Top             =   1440
         Width           =   840
      End
      Begin VB.TextBox TXTSMDTOP3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   3000
         MaxLength       =   5
         TabIndex        =   153
         TabStop         =   0   'False
         Top             =   540
         Width           =   840
      End
      Begin VB.TextBox TXTINSTOP2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   2340
         MaxLength       =   4
         TabIndex        =   31
         Top             =   1140
         Width           =   600
      End
      Begin VB.TextBox TXTSMDBOT2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   2340
         MaxLength       =   4
         TabIndex        =   25
         Top             =   240
         Width           =   600
      End
      Begin VB.TextBox TXTSMDTOP2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   2340
         MaxLength       =   4
         TabIndex        =   27
         Top             =   540
         Width           =   600
      End
      Begin VB.TextBox TXTINSBOT2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   2340
         MaxLength       =   4
         TabIndex        =   29
         Top             =   840
         Width           =   600
      End
      Begin VB.TextBox TXTINSTOP1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   30
         Top             =   1140
         Width           =   600
      End
      Begin VB.TextBox TXTSMDBOT1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   24
         Top             =   240
         Width           =   600
      End
      Begin VB.TextBox TXTSMDTOP1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   26
         Top             =   540
         Width           =   600
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
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
         TabIndex        =   152
         TabStop         =   0   'False
         Text            =   "Comp.SMD TOP"
         Top             =   540
         Width           =   1245
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   151
         TabStop         =   0   'False
         Text            =   "Comp.SMD BOT"
         Top             =   240
         Width           =   1245
      End
      Begin VB.TextBox TXTINSBOT1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1680
         MaxLength       =   4
         TabIndex        =   28
         Top             =   840
         Width           =   600
      End
      Begin VB.TextBox TXTLABEL2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   150
         TabStop         =   0   'False
         Text            =   "Comp. INS.BOT"
         Top             =   840
         Width           =   1245
      End
      Begin VB.TextBox TxtSup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   149
         TabStop         =   0   'False
         Text            =   "TOTAL COMPONENTES "
         Top             =   1440
         Width           =   2150
      End
      Begin VB.TextBox TXTLABEL2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   148
         TabStop         =   0   'False
         Text            =   "Comp.INS TOP"
         Top             =   1140
         Width           =   1245
      End
   End
   Begin VB.Frame Frame6 
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
      Height          =   1900
      Left            =   100
      TabIndex        =   142
      Top             =   7710
      Width           =   2655
      Begin VB.TextBox TXTMASDEFONDO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1560
         MaxLength       =   12
         TabIndex        =   23
         Top             =   855
         Width           =   960
      End
      Begin VB.TextBox TXTPANELIZADO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1560
         MaxLength       =   24
         TabIndex        =   22
         Top             =   540
         Width           =   960
      End
      Begin VB.CheckBox Check20 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Test Eléctrico"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   146
         Top             =   360
         Width           =   2100
      End
      Begin VB.CheckBox Check19 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Panelizado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   145
         Top             =   660
         Width           =   2100
      End
      Begin VB.CheckBox Check18 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Másc.  Fondo."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   144
         Top             =   960
         Width           =   2100
      End
      Begin VB.CheckBox Check17 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Proceso incia en Máscara"
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
         Left            =   120
         TabIndex        =   143
         Top             =   1260
         Width           =   2480
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Entrega"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   8
      Left            =   2880
      TabIndex        =   140
      Top             =   7110
      Width           =   4095
      Begin VB.CheckBox Check13 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Prototipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   141
         Top             =   280
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Corte de Placa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   7
      Left            =   -15
      TabIndex        =   137
      Top             =   7110
      Width           =   2775
      Begin VB.CheckBox Check16 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Router"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   139
         Top             =   280
         Width           =   855
      End
      Begin VB.CheckBox Check15 
         BackColor       =   &H00E2D3C0&
         Caption         =   "V.Scoring"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1485
         TabIndex        =   138
         Top             =   280
         Width           =   1090
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Especiales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   680
      Index           =   6
      Left            =   120
      TabIndex        =   131
      Top             =   6420
      Width           =   6855
      Begin VB.TextBox Text15 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   5760
         MaxLength       =   10
         TabIndex        =   21
         Top             =   240
         Width           =   960
      End
      Begin VB.CheckBox Check12 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Otro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2760
         TabIndex        =   135
         Top             =   280
         Width           =   735
      End
      Begin VB.CheckBox Check11 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Ni.Au."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1880
         TabIndex        =   134
         Top             =   280
         Width           =   855
      End
      Begin VB.CheckBox Check10 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Goma"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1000
         TabIndex        =   133
         Top             =   280
         Width           =   855
      End
      Begin VB.CheckBox Check9 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Grafito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   132
         Top             =   280
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Superficie (mm2)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   136
         Top             =   285
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Leyenda Componentes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Index           =   5
      Left            =   3120
      TabIndex        =   127
      Top             =   5820
      Width           =   3855
      Begin VB.CommandButton Command6 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1815
         TabIndex        =   195
         Top             =   240
         Width           =   175
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1410
         MaxLength       =   3
         TabIndex        =   20
         Top             =   240
         Width           =   390
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   2010
         Locked          =   -1  'True
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   240
         Width           =   1755
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00E2D3C0&
         Caption         =   "L.S."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   129
         Top             =   280
         Width           =   615
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00E2D3C0&
         Caption         =   "L.C."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   780
         TabIndex        =   128
         Top             =   280
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Mascara Antisoldante"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Index           =   4
      Left            =   -15
      TabIndex        =   122
      Top             =   5580
      Width           =   3015
      Begin VB.TextBox Text10 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   19
         Top             =   360
         Width           =   390
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00E2D3C0&
         Caption         =   "L.C."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   730
         TabIndex        =   126
         Top             =   405
         Width           =   615
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00E2D3C0&
         Caption         =   "L.S."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   125
         Top             =   405
         Width           =   615
      End
      Begin VB.CommandButton CMD8 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1720
         TabIndex        =   124
         Top             =   360
         Width           =   175
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   360
         Width           =   1035
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Archivo Asociados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Index           =   2
      Left            =   3120
      TabIndex        =   119
      Top             =   4800
      Width           =   3855
      Begin VB.ListBox List1 
         Height          =   645
         ItemData        =   "Cotiza.frx":000C
         Left            =   120
         List            =   "Cotiza.frx":000E
         TabIndex        =   205
         ToolTipText     =   "Archivos Asociados"
         Top             =   240
         Width           =   3375
      End
      Begin VB.CommandButton CMD9 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   3600
         TabIndex        =   121
         ToolTipText     =   "Borra el Acceso al Archivo Seleccionado"
         Top             =   240
         Width           =   175
      End
      Begin VB.CommandButton CMD10 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   3600
         TabIndex        =   120
         ToolTipText     =   "Seleccionar Archivo Para Asociar a Esta Cotización"
         Top             =   600
         Width           =   175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Programa de PCB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   3
      Left            =   -15
      TabIndex        =   116
      Top             =   4815
      Width           =   3015
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   120
         MaxLength       =   3
         TabIndex        =   18
         Top             =   300
         Width           =   390
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   300
         Width           =   2205
      End
      Begin VB.CommandButton CMD7 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   525
         TabIndex        =   117
         Top             =   300
         Width           =   175
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Plazo de Entrega"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5040
      TabIndex        =   113
      Top             =   3540
      Width           =   1935
      Begin VB.TextBox TXTADICIONALARMADOS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   945
         MaxLength       =   4
         TabIndex        =   17
         ToolTipText     =   "Al modificar este casillero se actualiza la fecha de entrega "
         Top             =   720
         Width           =   840
      End
      Begin VB.TextBox TXTCIRCUITOS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   960
         MaxLength       =   4
         TabIndex        =   16
         ToolTipText     =   "Al modificar este casillero se actualiza la fecha de entrega "
         Top             =   300
         Width           =   840
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Adicional Armados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   120
         TabIndex        =   115
         Top             =   600
         Width           =   750
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Circuitos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   114
         Top             =   330
         Width           =   750
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Excedente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2475
      TabIndex        =   108
      Top             =   3540
      Width           =   2490
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1560
         MaxLength       =   4
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   840
         Width           =   480
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   15
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   111
         TabStop         =   0   'False
         Text            =   "Tolerancia (%) +/-"
         Top             =   840
         Width           =   1245
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Acepta Excedente"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   110
         Top             =   600
         Value           =   -1  'True
         Width           =   1900
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Producción Exacta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   109
         Top             =   240
         Width           =   1900
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Estado Actual"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   103
      Top             =   3540
      Width           =   2415
      Begin VB.CheckBox Check4 
         BackColor       =   &H00E2D3C0&
         Caption         =   "No Standard"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   107
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Prev. Aprob. Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   106
         Top             =   720
         Width           =   2100
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Prev. Aprob. Muestra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   105
         Top             =   480
         Width           =   2100
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Falta Inf. Técnica"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   104
         Top             =   240
         Width           =   2100
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Height          =   2240
      Index           =   1
      Left            =   3840
      TabIndex        =   77
      Top             =   1300
      Width           =   3135
      Begin VB.TextBox TXTCANTICAPAS 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   120
         TabIndex        =   215
         ToolTipText     =   "Cantidad de Capas Cuando es Multicapa"
         Top             =   240
         Width           =   330
      End
      Begin VB.TextBox TXTNUMPISTA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   930
         MaxLength       =   3
         TabIndex        =   15
         Top             =   1810
         Width           =   390
      End
      Begin VB.TextBox TXTNUMPTH 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   930
         MaxLength       =   3
         TabIndex        =   14
         Top             =   1500
         Width           =   390
      End
      Begin VB.TextBox TXTNUMESPCOBRE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   280
         Left            =   930
         MaxLength       =   3
         TabIndex        =   13
         Top             =   1190
         Width           =   390
      End
      Begin VB.TextBox TXTNUMESPMAT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   930
         MaxLength       =   3
         TabIndex        =   12
         Top             =   870
         Width           =   390
      End
      Begin VB.TextBox TXTNUMTMAT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   930
         MaxLength       =   3
         TabIndex        =   11
         Top             =   550
         Width           =   390
      End
      Begin VB.TextBox TXTNUMCARA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   930
         MaxLength       =   3
         TabIndex        =   10
         Top             =   240
         Width           =   390
      End
      Begin VB.TextBox TXTPISTAS 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   1810
         Width           =   1500
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
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
         TabIndex        =   101
         TabStop         =   0   'False
         Text            =   "Pistas"
         Top             =   1810
         Width           =   800
      End
      Begin VB.CommandButton CMD6 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1305
         TabIndex        =   100
         Top             =   1810
         Width           =   175
      End
      Begin VB.TextBox TXTCARAS 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   240
         Width           =   1500
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Index           =   23
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   98
         TabStop         =   0   'False
         Text            =   "Caras"
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox TXTTIPOMATERIAL 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   550
         Width           =   1500
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
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
         TabIndex        =   96
         TabStop         =   0   'False
         Text            =   "Tipo Mat."
         Top             =   550
         Width           =   800
      End
      Begin VB.TextBox TXTESPESOR 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   870
         Width           =   1500
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
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
         TabIndex        =   94
         TabStop         =   0   'False
         Text            =   "Esp.Mat."
         Top             =   870
         Width           =   800
      End
      Begin VB.TextBox TXTESPCOBRE 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   1190
         Width           =   1500
      End
      Begin VB.TextBox TXTPTH 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1500
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
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
         TabIndex        =   91
         TabStop         =   0   'False
         Text            =   "PTH"
         Top             =   1500
         Width           =   800
      End
      Begin VB.TextBox TXT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   90
         TabStop         =   0   'False
         Text            =   "Esp.Cobre"
         Top             =   1190
         Width           =   800
      End
      Begin VB.CommandButton CMD1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1305
         TabIndex        =   89
         Top             =   240
         Width           =   175
      End
      Begin VB.CommandButton CMD2 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1305
         TabIndex        =   88
         Top             =   550
         Width           =   175
      End
      Begin VB.CommandButton CMD3 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1305
         TabIndex        =   87
         Top             =   870
         Width           =   175
      End
      Begin VB.CommandButton CMD4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1305
         TabIndex        =   86
         Top             =   1190
         Width           =   175
      End
      Begin VB.CommandButton CMD5 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1305
         TabIndex        =   85
         Top             =   1500
         Width           =   175
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Medidas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2240
      Index           =   0
      Left            =   0
      TabIndex        =   76
      Top             =   1300
      Width           =   3735
      Begin VB.CommandButton Command12 
         Caption         =   "Cálc.Paneles"
         Enabled         =   0   'False
         Height          =   250
         Left            =   2520
         TabIndex        =   220
         Top             =   1970
         Width           =   1095
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Cálc.Circuito"
         Height          =   250
         Left            =   1200
         TabIndex        =   219
         Top             =   1970
         Width           =   1095
      End
      Begin VB.TextBox TxtSupDM22INV 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   218
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox Check21 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Circuitos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   217
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check14 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Paneles"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2520
         TabIndex        =   216
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtSupLoteM22 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   211
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1100
      End
      Begin VB.TextBox TxtSupDM22 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   210
         TabStop         =   0   'False
         Top             =   1380
         Width           =   1100
      End
      Begin VB.TextBox TxtAncho2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   209
         Top             =   1080
         Width           =   1100
      End
      Begin VB.TextBox TxtLargo2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   208
         Top             =   780
         Width           =   1100
      End
      Begin VB.TextBox TXTCANTIDAD2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   207
         Top             =   480
         Width           =   1100
      End
      Begin VB.TextBox TxtCantidad 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1200
         TabIndex        =   7
         Top             =   480
         Width           =   1100
      End
      Begin VB.TextBox TxtLargo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1200
         TabIndex        =   8
         Top             =   780
         Width           =   1100
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
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
         TabIndex        =   84
         TabStop         =   0   'False
         Text            =   "Largo (mm.)"
         Top             =   780
         Width           =   1000
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
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
         TabIndex        =   83
         TabStop         =   0   'False
         Text            =   "Cantidad"
         Top             =   480
         Width           =   1000
      End
      Begin VB.TextBox TxtAncho 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1200
         TabIndex        =   9
         Top             =   1080
         Width           =   1100
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
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
         TabIndex        =   82
         TabStop         =   0   'False
         Text            =   "Ancho (mm.)"
         Top             =   1080
         Width           =   1000
      End
      Begin VB.TextBox TxtSupDM2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   1380
         Width           =   1100
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
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
         TabIndex        =   80
         TabStop         =   0   'False
         Text            =   "Sup.(dm2)"
         Top             =   1380
         Width           =   1000
      End
      Begin VB.TextBox TxtSupLoteM2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1100
      End
      Begin VB.TextBox TXTLABEL 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
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
         TabIndex        =   78
         TabStop         =   0   'False
         Text            =   "Sup. Lote m2"
         Top             =   1680
         Width           =   1000
      End
      Begin VB.Label LBLMODO 
         Caption         =   "Label5"
         Height          =   255
         Left            =   1320
         TabIndex        =   225
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Height          =   1335
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Width           =   15615
      Begin VB.TextBox TXTLABEL1 
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
         TabIndex        =   226
         TabStop         =   0   'False
         Text            =   "Contacto"
         Top             =   940
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   15360
         TabIndex        =   224
         Top             =   960
         Width           =   175
      End
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   14800
         TabIndex        =   213
         ToolTipText     =   "Generar Nueva Cotización a Partir de la Activa"
         Top             =   150
         Width           =   175
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   15000
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   212
         Top             =   150
         Width           =   495
      End
      Begin VB.CommandButton Command10 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8805
         TabIndex        =   206
         Top             =   150
         Width           =   175
      End
      Begin VB.TextBox TXTSINO 
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
         Left            =   12615
         TabIndex        =   203
         TabStop         =   0   'False
         Top             =   150
         Width           =   495
      End
      Begin VB.TextBox Text19 
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
         Height          =   330
         Left            =   1440
         MaxLength       =   48
         TabIndex        =   1
         Text            =   " "
         Top             =   940
         Width           =   4245
      End
      Begin VB.TextBox Text18 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   150
         Width           =   920
      End
      Begin VB.CommandButton Command8 
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1280
         TabIndex        =   200
         Top             =   555
         Width           =   175
      End
      Begin VB.CommandButton Command5 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7970
         TabIndex        =   75
         Top             =   940
         Width           =   175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10395
         TabIndex        =   74
         Top             =   940
         Width           =   175
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
         Height          =   330
         Left            =   6600
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   5
         Top             =   940
         Width           =   1335
      End
      Begin VB.TextBox TXTLABEL1 
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
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   73
         TabStop         =   0   'False
         Text            =   "Compra"
         Top             =   940
         Width           =   855
      End
      Begin VB.TextBox Text6 
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
         Height          =   330
         Left            =   9000
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   6
         Top             =   940
         Width           =   1335
      End
      Begin VB.TextBox TXTLABEL1 
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
         Left            =   8280
         Locked          =   -1  'True
         TabIndex        =   72
         TabStop         =   0   'False
         Text            =   "Entrega"
         Top             =   940
         Width           =   735
      End
      Begin VB.TextBox Text5 
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
         Height          =   330
         Left            =   12000
         TabIndex        =   4
         Top             =   150
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   12420
         TabIndex        =   71
         Top             =   150
         Width           =   175
      End
      Begin VB.TextBox TXTLABEL1 
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
         Left            =   11280
         Locked          =   -1  'True
         TabIndex        =   70
         TabStop         =   0   'False
         Text            =   "Armado"
         Top             =   150
         Width           =   735
      End
      Begin VB.TextBox Text4 
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
         Left            =   9000
         MaxLength       =   20
         TabIndex        =   3
         Top             =   120
         Width           =   2175
      End
      Begin VB.TextBox TXTLABEL1 
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
         Left            =   8040
         Locked          =   -1  'True
         TabIndex        =   69
         TabStop         =   0   'False
         Text            =   "Código"
         Top             =   150
         Width           =   735
      End
      Begin VB.TextBox Text3 
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
         Height          =   330
         Left            =   5195
         MaxLength       =   128
         TabIndex        =   2
         Top             =   550
         Width           =   10300
      End
      Begin VB.TextBox TXTLABEL1 
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
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   68
         TabStop         =   0   'False
         Text            =   "Modelo"
         Top             =   550
         Width           =   735
      End
      Begin VB.TextBox Text2 
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
         Left            =   1440
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   150
         Width           =   6045
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1100
         TabIndex        =   66
         Top             =   150
         Width           =   175
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
         Height          =   330
         Left            =   13800
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   65
         Top             =   150
         Width           =   975
      End
      Begin VB.TextBox TXTLABEL1 
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
         Left            =   13440
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         Text            =   "N°"
         Top             =   150
         Width           =   375
      End
      Begin VB.TextBox TXTFECHA 
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
         Height          =   330
         Left            =   13980
         MaxLength       =   10
         TabIndex        =   63
         Top             =   940
         Width           =   1335
      End
      Begin VB.TextBox TXTLABEL1 
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
         Left            =   13350
         Locked          =   -1  'True
         TabIndex        =   62
         TabStop         =   0   'False
         Text            =   "Fecha"
         Top             =   940
         Width           =   615
      End
      Begin VB.Label lblestado 
         Alignment       =   2  'Center
         BackColor       =   &H00E2D3C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   330
         Left            =   7515
         TabIndex        =   227
         Top             =   150
         Width           =   495
      End
      Begin VB.Label Label4 
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
         Height          =   330
         Left            =   1470
         TabIndex        =   223
         Top             =   555
         Width           =   2955
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   135
         Left            =   5760
         TabIndex        =   214
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblNuevaCotizacion 
         BackColor       =   &H00E2D3C0&
         Caption         =   "Nueva Cotización"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   10680
         TabIndex        =   202
         Top             =   940
         Visible         =   0   'False
         Width           =   2775
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
         Height          =   330
         Left            =   120
         TabIndex        =   201
         Top             =   555
         Width           =   1140
      End
   End
   Begin VB.TextBox TXTUSO 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   0
      Left            =   7320
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.TextBox TXTUNIMED 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   14400
      Locked          =   -1  'True
      TabIndex        =   59
      TabStop         =   0   'False
      Text            =   "M."
      Top             =   11040
      Width           =   435
   End
   Begin VB.TextBox TXTLABEL1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   170
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Text            =   "CORTE:"
      Top             =   11040
      Width           =   795
   End
   Begin VB.TextBox TXTCORTE 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   13440
      Locked          =   -1  'True
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   11040
      Width           =   885
   End
   Begin VB.Frame FRAME17 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Adjuntar Archivos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   7
      Left            =   11640
      TabIndex        =   53
      Top             =   11400
      Width           =   3975
      Begin VB.CommandButton Command47 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1440
         TabIndex        =   56
         Top             =   1605
         Width           =   1185
      End
      Begin VB.CommandButton Command47 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2700
         TabIndex        =   55
         Top             =   1605
         Width           =   1125
      End
      Begin VB.ListBox List2 
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
         Height          =   1290
         ItemData        =   "Cotiza.frx":0010
         Left            =   240
         List            =   "Cotiza.frx":0012
         TabIndex        =   54
         Top             =   240
         Width           =   3630
      End
   End
   Begin VB.TextBox TXTOBS 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   32
      Left            =   13440
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   10680
      Width           =   2205
   End
   Begin VB.TextBox TXTLABEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   32
      Left            =   11640
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Text            =   "CÓDIGO DE CLIENTE"
      Top             =   10680
      Width           =   1755
   End
   Begin VB.TextBox TXTFECHREVISO 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   14640
      Locked          =   -1  'True
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   11400
      Width           =   1000
   End
   Begin VB.TextBox TXTLABEL1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   35
      Left            =   11640
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Text            =   "PREPARÓ"
      Top             =   11100
      Width           =   1000
   End
   Begin VB.TextBox TXTLABEL1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Index           =   34
      Left            =   11640
      Locked          =   -1  'True
      TabIndex        =   48
      TabStop         =   0   'False
      Text            =   "REVISÓ"
      Top             =   11400
      Width           =   1000
   End
   Begin VB.TextBox TXTFECHPREPARO 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   14640
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   11100
      Width           =   1000
   End
   Begin VB.TextBox TXTPREPARO 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   11100
      Width           =   2000
   End
   Begin VB.TextBox TXTREVISO 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   11400
      Width           =   2000
   End
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuNewCotiz 
         Caption         =   "Nuevo Cotización"
      End
      Begin VB.Menu mnuEditCotización 
         Caption         =   "Abrir Cotización Existente"
      End
      Begin VB.Menu mnuGuardar 
         Caption         =   "Guardar Cotización Activa"
      End
      Begin VB.Menu mnuImprimirCotizActiva 
         Caption         =   "Imprimir Cotización Activa"
         Begin VB.Menu mnuModIndividual 
            Caption         =   "Modo Individual"
         End
         Begin VB.Menu mnuPrintMultiple 
            Caption         =   "Modo Múltiple"
         End
      End
      Begin VB.Menu mnuGuardarComoNuevaCotizacion 
         Caption         =   "Guardar Como Nueva Cotización"
      End
      Begin VB.Menu mnuAproCot 
         Caption         =   "Aprobar Cotización Activa"
      End
   End
   Begin VB.Menu mnuConsultasyllistados 
      Caption         =   "Consultas y Listados"
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuracion"
      Begin VB.Menu mnuTablaAuxiliares 
         Caption         =   "Tablas Auxiliares"
      End
      Begin VB.Menu mnuTablatmat 
         Caption         =   "Tabla Tipo de Material"
      End
      Begin VB.Menu mnuImpLogar 
         Caption         =   "Imortador de Tablas Logaritmica"
      End
      Begin VB.Menu mnusetup 
         Caption         =   "Configuración General"
      End
   End
   Begin VB.Menu mnuAccesosDirectos 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuSegCom 
         Caption         =   "Seguimiento Comercial"
      End
      Begin VB.Menu mnuAbmSec 
         Caption         =   "ABM de Secuencias"
      End
   End
End
Attribute VB_Name = "Cotiza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MODIFIC%
Public PRESENTANDOINFOBD%
Public NOPRESENTARMENSAJE%
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Sub Actualiza_CostoPanel()
   If Check14.Value = 1 Then
        If XVALO(TXTCANTIDAD2) > 0 Then
           If XVALO(TxtCantidad) > 0 Then
             TxtPrcdm2Costo = FORMATNUM$(CALCULACOSTO(TxtSupLoteM22, TXTCARAS, TXTTIPOMATERIAL, TXTESPESOR, TXTESPCOBRE, TXTPTH, RET_COSTONORMAL), "F16.3")
             TxtPrcdm2 = FORMATNUM$(RET_COSTONORMAL, "F16.3") 'VALOR RETORNADO DE LA FUNCION CALCULACOSTO
             TxtPrcUnit = ""
             If XVALO(TxtPrcdm2) > 0 Then
                 If XVALO(TxtSupDM22) > 0 Then
                    TxtPrcUnit = TRIM$(FORMATNUM$(XVALO(TxtPrcdm2) * XVALO(TxtSupDM22INV), "F16.3"))
                 End If
             End If
             
             TxtPrcUnitCosto = ""
             If XVALO(TxtPrcdm2Costo) > 0 Then
                 If XVALO(TxtSupDM22) > 0 Then
                    TxtPrcUnitCosto = TRIM$(FORMATNUM$(XVALO(TxtPrcdm2Costo) * XVALO(TxtSupDM22INV), "F16.3"))
                 End If
             End If
           End If
           TXTtotPlca = FORMATNUM$(XVALO(TxtPrcUnit) * XVALO(TxtCantidad), "F16.3")
           TXTtotPlcaCosto = FORMATNUM$(XVALO(TxtPrcUnitCosto) * XVALO(TxtCantidad), "F16.3")
       End If

   End If
End Sub

Sub Actualizar_CostoCircuito()
   TxtPrcdm2Costo = FORMATNUM$(CALCULACOSTO(TxtSupLoteM2, TXTCARAS, TXTTIPOMATERIAL, TXTESPESOR, TXTESPCOBRE, TXTPTH, RET_COSTONORMAL), "F16.3")
   TxtPrcdm2 = FORMATNUM$(RET_COSTONORMAL, "F16.3") 'VALOR RETORNADO DE LA FUNCION CALCULACOSTO
   TxtPrcUnit = ""
   If XVALO(TxtPrcdm2) > 0 Then
       If XVALO(TxtSupDM2) > 0 Then
          TxtPrcUnit = TRIM$(FORMATNUM$(XVALO(TxtPrcdm2) * XVALO(TxtSupDM2), "F16.3"))
       End If
   End If
   
   TxtPrcUnitCosto = ""
   If XVALO(TxtPrcdm2Costo) > 0 Then
       If XVALO(TxtSupDM2) > 0 Then
          TxtPrcUnitCosto = TRIM$(FORMATNUM$(XVALO(TxtPrcdm2Costo) * XVALO(TxtSupDM2), "F16.3"))
       End If
   End If
   TXTtotPlca = FORMATNUM$(XVALO(TxtPrcUnit) * XVALO(TxtCantidad), "F16.3")
   TXTtotPlcaCosto = FORMATNUM$(XVALO(TxtPrcUnitCosto) * XVALO(TxtCantidad), "F16.3")
End Sub

Function ACTUALIZATOTOALES_MONTAJE_SMD(SMDBOT1, SMDBOT2, SMDTOP1, SMDTOP2)
     ACTUALIZATOTOALES_MONTAJE_SMD = XVALO(SMDBOT1) + XVALO(SMDBOT2) + XVALO(SMDTOP1) + XVALO(SMDTOP2)
End Function

Function ACTUALIZATOTOALES_MONTAJE_INS(INSBOT1, INSBOT2, INSTOP1, INSTOP2)
     ACTUALIZATOTOALES_MONTAJE_INS = XVALO(INSBOT1) + XVALO(INSBOT2) + XVALO(INSTOP1) + XVALO(INSTOP2)
End Function
Sub BLANFORMU()
    Call LIMPIARTEXT(Me, 1, 1)
    Call LIMPIAROPTION(Me)
    Call LIMPIARCHECK(Me)
    Me.Text31 = ""
    List1.Clear
    Option4.Value = True
    Check21.Value = 1
    Text8 = 10
    Me.TXTNUMCARA = 1
    TXTCIRCUITOS = 12
    Check5.Value = 1
    Check7.Value = 1
    Check16.Value = 1
    VALODEFECTO% = 0
    Label1.Caption = ""
    Label3.Visible = False
    '
    TxtPrcdm2 = ""
    TxtPrcUnit = ""
    TXTtotPlca = ""
    TXTPrcTtOpcion = ""
    TXTPreunitCto = ""
    TXTPreunit_Arm = ""
    TXTPreunit_Cpte = ""
    Text26 = ""
    Text32 = ""
    TXTPrcTtPlcaSA = ""
    TXTCantStencil = ""
    TXTStTotalusd = ""
    TXTTOTALSTENCIL = ""
    TXTTtOpcion = ""
    TxtPrcUnitCosto = ""
    TXTtotPlcaCosto = ""
    TXTPrcTtOpcionCosto = ""
    TXTPreunit_ArmCosto = ""
    TXTPreunit_CpteCosto = ""
    TXTPrcTtPlcaSACosto = ""
    TXTTtOpcionCosto = ""
    Text34 = ""
    TxtPrcdm2Costo = ""

    TXTObservaAdmin = ""
    TXTNROOCOM = ""
    TXTObservaPlanta = ""
    TXTObservPenal = ""
    
    TXTObservCliente = ""
    TXTPANELIZADO = ""
    TXTMASDEFONDO = ""
    TXTSMDBOT1 = ""
    TXTSMDTOP1 = ""
    TXTINSBOT1 = ""
    TXTINSTOP1 = ""
    TXTSMDBOT2 = ""
    TXTSMDTOP2 = ""
    TXTINSBOT2 = ""
    TXTINSTOP2 = ""
    TXTSMDTOP3 = ""
    TXTINSTOP3 = ""
    TXTTOTALCOMPONENTES = ""
    LBLMODO = ""
    Label26 = "" ' codigo cliente prospect
    lblestado = ""
    MODIFIC% = 0


  
End Sub

Function CALCULASUPUNITARIA(LARGO As Double, ANCHO As Double) As Double
    
    
      If XVALO(LARGO) > 0 Then
        If XVALO(ANCHO) > 0 Then
           'PASO A DM
           LARGO = LARGO / 100
           ANCHO = ANCHO / 100
           CALCULASUPUNITARIA = LARGO * ANCHO
        End If
      End If
          

End Function

Function CALCULASUPERFICIETOTAL(CANTI, LARGO As Double, ANCHO As Double) As Double
     If CANTI > 0 Then
      'LARGO Y ANCHO VIENEN EN MM LONGITUD
      If XVALO(LARGO) > 0 Then
        If XVALO(ANCHO) > 0 Then
           'PASO A DM
           LARGO = LARGO / 100
           ANCHO = ANCHO / 100
           
           'CALCULO LA SUPERFICIE TOTAL EN M2
           CALCULASUPERFICIETOTAL = CANTI * ((LARGO / 10) * (ANCHO / 10))
        End If
      End If
     End If
End Function

Function ConNroCotizacion%(nrocotiz%)
     ConNroCotizacion% = nrocotiz%
     If XVALO(nrocotiz%) <= 0 Then
        OPX% = COMALTER%("Desea Crear una Nueva Cotización\\Cancelar\Crear Nueva Cotización")
        If OPX% = 2 Then
           mnuNewCotiz_Click
        End If
     End If
End Function



Sub GeneraPedidoDeCotización(NC, RETORNO_PEDIDO, NROCOTI&, VERSIONX&, fentrega%)
       
       OKX% = 0
       RETORNO_PEDIDO = 0
       RSOC$ = TRIM$(rasocli$(1)) 'PARA REFRESCAR EL RECORDSET
       If fentrega% < 1 Then
          fentrega% = HOY(HO$)
       End If
       Call ABREPEDCLI
       '
       On Error GoTo ERROR_GUARDAR
       FLEXCONN.BeginTrans
       
       Screen.MousePointer = 11
       APREC$ = "LIPRE" + "001"
       PRLST# = PRELISTA#(APREC$, CIXI%)
       FEMI% = HOY(HO$)
       FEX$ = FECHATEX$(FEMI%)
       Dim RS As ADODB.Recordset
       Set RS = New ADODB.Recordset
       SQLX$ = "SELECT * FROM  PEDENCA  "
       SQLX$ = SQLX$ + " WHERE NroClie = -1"
25     On Error Resume Next
       RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 25
           End If
           On Error GoTo 0
       With RS
         .AddNew
         !CODIEMPR = CodiEmp%
         NROPEDI1& = NUPEDIDO& ' FUNCION

         !NROPED = NROPEDI1&
         !FeSolEnt = CVDAT(FEMI%)
         RSOC$ = TRIM$(rasocli$(NC))
         REFE$ = " - " + TRIM$(TXTNROOCOM.TEXT)
         If REFE$ = " - " Then REFE$ = ""
         FEX = FEMI%
         REFE$ = REFE$ + " - " + FECHATEX$(FEX)
         LOMA% = 48 - Len(REFE$)
         REFE$ = Left$(RSOC$, LOMA%) + REFE$
         REFE$ = REFE$ & "(Cot. N°: " & NROCOTI& & ")"
         !REFERENCIA = Left$(REFE$, 128)
         !FECHEMIS = CVDAT(FEMI%)
         !NROCLIE = NC
         !RaSoClie = RSOC$
         !NROOCOM = Left$(TRIM$(TXTNROOCOM), 16)
         !DOMIENT = Left$(TRIM$(DOMICLI$(NC)), 48)
         !LOCAENT = Left$(TRIM$(LOCACLI$(NC)), 48)
         !TranspEnt = ""
         !NTranspo = 0
         !NSucent = 0
         !Observa = ""
         !Status = 0
         !ListPre = 0
         !CondPag = 0
         !VENDED = 0
         !TDescApli = ""
         .Update
       End With
       RS.Close
       Set RS = Nothing
       
       ANOTA$ = ""
       TTXYZ$ = Space$(76)
       HOII% = HOY(HOS$)
       Call AGREGA_ANOTAPED_NEW(NROPEDI1&, "Generación de Orden de Venta:")
       IKITEM% = 0
       '
       
       SQLX$ = "SELECT * FROM  PEDDETA  "
       SQLX$ = SQLX$ + " WHERE NROPED = " & NROCOTI&
       
       Set RsPeddeta = New ADODB.Recordset
30     On Error Resume Next
       RsPeddeta.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 30
       End If
       On Error GoTo 0
       
       ConsultaSql$ = "SELECT * FROM COTIDAICHI WITH(NOLOCK) "
       ConsultaSql$ = ConsultaSql$ + " Where NRO_COTIZA = " & NROCOTI&
       ConsultaSql$ = ConsultaSql$ + " AND VRS_COTIZA = " & VERSIONX&
'       If COTIZACIONDIFERENTESCODIGOS%(NROCOTI&) = 0 Then
'       ConsultaSql$ = ConsultaSql$ + " order by VRS_COTIZA"
       IKITEM% = 0
       Set RS1 = READSET(ConsultaSql$)
       With RS1
          Do While Not .EOF
                CIXI% = XVALO(!cod_inte)
                CANTI = XVALO(!CANTIDAD)
                TotoperVers = XVALO(!TtOpcion)
                preunx = TotoperVers / CANTI
                IKITEM% = IKITEM% + 1
                With RsPeddeta
                     .AddNew
                     !CODIEMPR = CodiEmp%
                     IKITEM% = IKITEM% + 1
                     !NROPED = NROPEDI1&
                     !FECHEMIS = CVDAT(FEMI%)
                     !NROCLIE = NC
                     !RaSoClie = Left$(TRIM$(rasocli$(NC)), 64)
                     !NROOCOM = Left$(TRIM$(TXTNROOCOM), 16)
                     !FeSolEnt = CVDAT(fentrega%)
                     !Status = 0
                     !ListPre = 0
                     !CondPag = 0
                     !VENDED = 0
                     !PorComVend = 0
                     !NuOrItem = IKITEM%
                     !CODIINT = CIXI%
                     !CODIEXT = CODEXT$(CIXI%)
                     !SubCod = ""
                     !Descrip = DESCRIT0$(CIXI%)
                     !TLargo = 0
                     !TAncho = 0
                     !TEspesor = 0
                     !CBase = ""
                     !TTALLE = ""
                     !TCOLOR = ""
                     !VenRep = ""   ' Venta o Reparacion
                     !MoneEmis = 2
                     'PREUNI# = PREUNI
                     !TIPOCAM = TICAMONE(2)
                     
                     !PreList = preunx
                     !DescLista = 0
                     !DescCombi = ""
                     !DescRubro = 0
                     !DescTotal = 0
                     !CanAsig = 0
                     !CanPed = CANTI
                     !PreUnid = preunx
                     !ImpoTot = TotoperVers
                     !CANTENT = 0
                     !CanSaldo = 0
                     !FeUltEnt = CVDAT(0)
                     !NroRemi = ""
                     !TPedItem = ""
                     '
                     TTXYZ$ = Space$(76)
                     Call REPLA(TTXYZ$, "- " + CODEXT$(CIXI%), 11, 18)
                     Call REPLA(TTXYZ$, DESCRIT0$(CIXI%), 29, 38)
                     Call REPLA(TTXYZ$, CSTRING$(MKS$(CANTI), 3, 10, 1, 2), 57, 10)
                     Call AGREGA_ANOTAPED_NEW(NROPEDI1&, TTXYZ$)
                    .Update
                   End With
                  
                   RS1.MoveNext
            Loop
          End With
          '
          On Error Resume Next
          RsPeddeta.Close
          Set RsPeddeta = Nothing
'          RS1.Close
'          Set RS1 = Nothing
          On Error GoTo 0
          TTT$ = "Cotización Nro.: " & NROCOTI& & "-" & VERSIONX&
          Call AGREGA_ANOTAPED_NEW(NROPEDI1&, TTT$)
          '
          OKX% = 1
          '

98        Screen.MousePointer = 1
          RETORNO_PEDIDO = NROPEDI1&
          'ACA GRABA EL NRO DE PEDIDO EN LA TABLA DE COTIZACION EL NRO DE PEDIDO, EN TODOS O SOLO EN LA VERSION,
          CONDI$ = " NRO_COTIZA = " & NROCOTI& & " AND VRS_COTIZA = " & VERSIONX&
'          If COTIZACIONDIFERENTESCODIGOS%(NROCOTI&) < 1 Then CONDI$ = CONDI$ + " AND VRS_COTIZA = " & VERSIONX&
          Call ACTUREGISTRO("COTIDAICHI", "NROPED", CONDI$, NROPEDI1&, REGAF&)
          
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
    Else
      FLEXCONN.CommitTrans
    End If


End Sub

Sub GRABARARCHIVOS(NROCOTIZA)
    Call ACTUREGISTRO("DOCURELCOT", "MARBORRA", "NRO_COTIZA = " & NROCOTIZA, 1, REG&)
    '
    For i& = 1 To List1.ListCount
        Z$ = List1.List(i& - 1)
        DESCRI$ = TRIM$(Mid$(Z, 1, 24))
        RUTAX$ = TRIM$(Mid$(Z, 32, 512))
        SQLX$ = "INSERT INTO DOCURELCOT ("
        SQLX$ = SQLX$ + "CODIEMPR,DESCRIPCION,  NRO_COTIZA,MARBORRA,RUTA)"
        SQLX$ = SQLX$ + " VALUES(" & CodiEmp%
        SQLX$ = SQLX$ + " ,'" & DESCRI$ & "'"
        SQLX$ = SQLX$ + " ," & NROCOTIZA
        SQLX$ = SQLX$ + " ," & 0
        SQLX$ = SQLX$ + " ,'" & RUTAX$ & "'"
        SQLX$ = SQLX$ + ")"
        FLEXCONN.Execute (SQLX$)
   Next i&
   
   
    'BORRO LOS REGISTROS MARCADOS CON 1
    Call BORRAREGISTROS("DOCURELCOT", "MARBORRA = 1 AND NRO_COTIZA = " & NROCOTIZA, REGAF&, "NOMESS")
   
   'RECORRO EL LIST PARA GRABAR
'   For i& = 1 To List1.ListCount
'      Z$ = List1.List(i& - 1)
'      DESCRI$ = TRIM$(Mid$(Z, 1, 24))
'      RUTAX$ = TRIM$(Mid$(Z, 32, 512))
'10    RNDX$ = RANDSTRING$(8)
'      If CantRegistros("DOCURELCOT", "DESCRIPCION = '" & RNDX$ & "'", "NOMESS") > 0 Then GoTo 10
'
'      Call CreaRegistro("DOCURELCOT", "DESCRIPCION", RNDX$, "NOMESS")
'      Call ACTUREGISTRO("DOCURELCOT", "NRO_COTIZA", "DESCRIPCION = '" & RNDX$ & "'", nrocotiza, REG&)
'      Call ACTUREGISTRO("DOCURELCOT", "MARBORRA", "DESCRIPCION = '" & RNDX$ & "'", 0, REG&)
'      Call ACTUREGISTRO("DOCURELCOT", "RUTA", "DESCRIPCION = '" & RNDX$ & "'", RUTAX$, REG&)
'      Call ACTUREGISTRO("DOCURELCOT", "DESCRIPCION", "DESCRIPCION = '" & RNDX$ & "'", DESCRI$, REG&)
'   Next i&
'   '

End Sub

Sub printCotiza(NROCOTIZA, NROVERS%)
  'NO ESTA GUARDANDO LAS CANTIDADES DE PANELIZACION
  NROCOTIZACION& = NROCOTIZA
  If CantRegistros&("COTIDAICHI", "NRO_COTIZA = " & XVALO(Text1)) < 1 Then Exit Sub
  NC = XVALO(Text18)
  PROSP$ = TRIM$(Label26)
  CANT_VERS& = CantRegistros&("COTIDAICHI", "NRO_COTIZA = " & XVALO(Text1) & " And MODELO <> '" & TRIM$(Text3) & "'")
  If CANT_VERS& > 0 Then
     'ACA DEBERIA IMPRMIR SI O SI LA OTRA COTIZACIÓN O BIEN SOLO LA QUE ESTA ACTIVA.
     'VER EN TODO CASO DE ARMAR OTRO FRM.
  End If
 
  FORIMPRE$ = TRIM$(Control$("", "COTIDAIC"))
  If FORIMPRE$ = "" Then
     Call COMUNI("Falta Configurar Formulario de Impresión de Cotización")
     Exit Sub
  End If
  '
  If XVALO(Text1) < 1 Then
      Call COMUNI("Falta N° De Cotización")
      Exit Sub
  End If
  '
  If NC < 1 And PROSP$ = "" Then
      Call COMUNI("Esta Cotización No Tiene Asignado Un Número de Cliente")
      Exit Sub
  End If
  '
  'si tiene n cliente real imprime un frm (los 2 son iguales pero uno toma datos de cliente del deudor12 y otro de prospect)
  TEXPANEL$ = "-"
  If Check14.Value = 1 And XVALO(TXTCANTIDAD2) > 0 Then
     TEXPANEL$ = SAFETEXT$(XVALO(TXTCANTIDAD2)) & " x " & SAFETEXT$(XVALO(TxtLargo2)) & " x " & SAFETEXT$(XVALO(TxtAncho2))
  End If

   A = TEL1CLI$(1) ' PARA QUE SE REFRESQUE LA INFO DEL CLIENTE
    NC = XVALO(Text18)
    
    FECHDOC$ = TXTFECHA.TEXT
    NUMEDOC$ = TRIM$(Text1.TEXT)
    '

    TIPODOC$ = "COTIZACION "
    '
    ESDEDIFERENTESCODIGOS% = COTIZACIONDIFERENTESCODIGOS%(NROCOTIZACION&)
    
    CAPARDOC% = 82
    'LIMPIO EL VECTOR
    For rr = 1 To CAPARDOC%
       DOCPARAM$(rr) = ""
    Next rr

    If PROSP$ <> "" Then
        SQLX$ = "SELECT * FROM BSEGCOM WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE CODI_CLI = '" & PROSP$ & "'"
        Dim FICHTMP As ADODB.Recordset
        Set RS = READSET(SQLX$)
        With RS
           If Not .EOF Then
            CODPARAM$(2) = "CUIT-CLI": DOCPARAM$(2) = ""
            CODPARAM$(3) = "RASO-CLI": DOCPARAM$(3) = SAFETEXT$(TRIM$(Mid$(!raso_cli, 1, 59)))
            CODPARAM$(4) = "DOMI-CLI": DOCPARAM$(4) = SAFETEXT$(!domi_cli)
            CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = SAFETEXT$(!domi_cli)
            CODPARAM$(6) = "TELE-CLI": DOCPARAM$(6) = TRIM$(SAFETEXT$(!Tele_Cli) & " " & SAFETEXT$(!TELE_Cli2))
            CODPARAM$(7) = "FAX-CLI": DOCPARAM$(7) = SAFETEXT$(!FAX_CLI)
            CODPARAM$(8) = "MAIL-CLI": DOCPARAM$(8) = SAFETEXT$(!Mail_Cli)
            CODPARAM$(9) = "CTAC-CLI": DOCPARAM$(9) = Left$(Text19, 48)
          End If
        End With
        RS.Close
        Set RS = Nothing
    Else
        CODPARAM$(2) = "CUIT-CLI": DOCPARAM$(2) = CuitCli$(NC)
        CODPARAM$(3) = "RASO-CLI": DOCPARAM$(3) = Text2
        CODPARAM$(4) = "DOMI-CLI": DOCPARAM$(4) = DOMICLI$(NC)
        CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = TRIM$(DOMIENT$(NC))
        CODPARAM$(6) = "TELE-CLI": DOCPARAM$(6) = TRIM$(TEL1CLI$(NC))
        CODPARAM$(7) = "FAX-CLI": DOCPARAM$(7) = TRIM$(FAXCLI$(NC))
        CODPARAM$(8) = "MAIL-CLI": DOCPARAM$(8) = TRIM$(EMAILCLI$(NC))
        CODPARAM$(9) = "CTAC-CLI": DOCPARAM$(9) = TRIM$(CONTACTCLI$(NC))
    End If
    
    CODPARAM$(1) = "FECH-EMI": DOCPARAM$(1) = FECHDOC$
    CODPARAM$(10) = "DES-CONJ": DOCPARAM$(10) = TRIM$(Left(Text3, 24))
    CODPARAM$(11) = "COD-CONJ": DOCPARAM$(11) = TRIM$(Left(Text4, 24))
    CODPARAM$(12) = "MEDIDA-1": DOCPARAM$(12) = FORMATNUM$(XVALO(TxtLargo), "F7.3")
    CODPARAM$(13) = "TALLE": DOCPARAM$(13) = FORMATNUM$(XVALO(TxtAncho), "F7.3")
    CODPARAM$(14) = "MEDIDA-2": DOCPARAM$(14) = TRIM$(FORMATNUM$(TxtSupDM2, "F12.3"))
    CODPARAM$(15) = "MEDIDA-3": DOCPARAM$(15) = TRIM$(TXTTIPOMATERIAL)
    CODPARAM$(16) = "MED-PLCO": DOCPARAM$(16) = TRIM$(TXTESPESOR)
    CARAX$ = TRIM$(TXTCARAS): If XVALO(TXTCANTICAPAS) > 0 Then CARAX$ = CARAX$ + " (" & TXTCANTICAPAS & ")"
    CODPARAM$(17) = "IBRU-EMP": DOCPARAM$(17) = CARAX$
    CODPARAM$(18) = "REF-CONJ": DOCPARAM$(18) = TRIM$(TXTPTH)
    CODPARAM$(19) = "REF-PLCO": DOCPARAM$(19) = TRIM$(TXTESPCOBRE)
         valor$ = "NO": If Check15.Value > 0 Then valor$ = "SI"
    CODPARAM$(20) = "REF-MAT1": DOCPARAM$(20) = valor$
         valor$ = "NO": If Check16.Value > 0 Then valor$ = "SI"
    CODPARAM$(21) = "REF-MAT2": DOCPARAM$(21) = valor$
         PANELIXADO$ = "": If Check19.Value > 0 Then PANELIXADO$ = TXTPANELIZADO
    CODPARAM$(22) = "REF-MAT3": DOCPARAM$(22) = PANELIXADO$
         valor$ = "NO": If Check6.Value > 0 Then valor$ = "SI"
    CODPARAM$(23) = "REF-MAT4": DOCPARAM$(23) = valor$
         valor$ = "NO": If Check5.Value > 0 Then valor$ = "SI"
    CODPARAM$(24) = "REF-MAT5": DOCPARAM$(24) = valor$
         valor$ = "NO": If Check7.Value > 0 Then valor$ = "SI"
    CODPARAM$(25) = "REF-MAT6": DOCPARAM$(25) = valor$
         valor$ = "NO": If Check8.Value > 0 Then valor$ = "SI"
    CODPARAM$(26) = "REF-ASIE": DOCPARAM$(26) = valor$
    CODPARAM$(27) = "DET-PASE": DOCPARAM$(27) = TXTPISTAS
             valor$ = "NO": If Check18.Value > 0 Then valor$ = "SI"
    CODPARAM$(28) = "COL-MPRI": DOCPARAM$(28) = valor$
    CODPARAM$(29) = "ESP-MAT": DOCPARAM$(29) = TRIM$(TXTMASDEFONDO)
             valor$ = "NO": If Check9.Value > 0 Then valor$ = "SI"
    CODPARAM$(30) = "TIP-MPRI": DOCPARAM$(30) = valor$
             valor$ = "NO": If Check10.Value > 0 Then valor$ = "SI"
    CODPARAM$(31) = "PRO-MPRI": DOCPARAM$(31) = valor$
             valor$ = "NO": If Check11.Value > 0 Then valor$ = "SI"
    CODPARAM$(32) = "PAR-MPRI": DOCPARAM$(32) = valor$
             valor$ = "NO": If Check20.Value > 0 Then valor$ = "SI"
    CODPARAM$(33) = "UBI-MPRI": DOCPARAM$(33) = valor$
    
    valor$ = "": If Check6.Value > 0 Then valor$ = TRIM$(Text13)
    CODPARAM$(34) = "LAR-MATE": DOCPARAM$(34) = valor$
    valor$ = "": If Check5.Value > 0 Then valor$ = TRIM$(Text13)
    CODPARAM$(35) = "COD-MAT": DOCPARAM$(35) = valor$
    
    valor$ = "": If Check7.Value > 0 Then valor$ = TRIM$(Text14)
    CODPARAM$(36) = "ACH-MATE": DOCPARAM$(36) = valor$
    valor$ = "": If Check8.Value > 0 Then valor$ = TRIM$(Text14)
    CODPARAM$(37) = "DES-MAT": DOCPARAM$(37) = valor$

    valor$ = "NO": If Check11.Value > 0 Then valor$ = "SI"
    CODPARAM$(38) = "UNI-MPRI": DOCPARAM$(38) = valor$
    

    SQLX$ = "SELECT * FROM COTIDAICHI WITH(NOLOCK) WHERE NRO_COTIZA  = " & NROCOTIZACION
    If ESDEDIFERENTESCODIGOS% <> 0 Then SQLX$ = SQLX$ + "  AND VRS_COTIZA = " & NROVERS%
    SQLX$ = SQLX$ + " ORDER BY  VRS_COTIZA  ASC"
    Set RS = READSET(SQLX$)
    With RS
       Do While Not .EOF
          VERX% = XVALO(!VRS_COTIZA)
          If ESDEDIFERENTESCODIGOS% <> 0 Then
             'si entra aca lo imprime en la primer columna
            CODPARAM$(39) = "CAJU-EMP": DOCPARAM$(39) = XVALO(!CANTIDAD)
            CODPARAM$(40) = "N-PLANO": DOCPARAM$(40) = FORMATNUM$(XVALO(!PreunitCto), "F12.3")
            CODPARAM$(41) = "NUOPERS": DOCPARAM$(41) = FORMATNUM$(XVALO(!GASPREPA), "F12.3")
            CODPARAM$(42) = "NRO-REMI": DOCPARAM$(42) = FORMATNUM$(XVALO(!DISENO), "F12.3")
            CODPARAM$(43) = "NRO-AENT": DOCPARAM$(43) = FORMATNUM$(XVALO(!Preunit_Arm), "F12.3")
            CODPARAM$(44) = "CUIT-TRA": DOCPARAM$(44) = FORMATNUM$(XVALO(!Preunit_Cpte), "F12.3")
            CODPARAM$(45) = "PATE-TRA": DOCPARAM$(45) = FORMATNUM$(XVALO(!CantStencil) * XVALO(!StTotalusd), "F12.3")
            CODPARAM$(46) = "NUBULTOS": DOCPARAM$(46) = FORMATNUM$(XVALO(!TtOpcion), "F12.3")
            Exit Do
          End If
          Select Case VERX%
               Case 0
                 CODPARAM$(39) = "CAJU-EMP": DOCPARAM$(39) = XVALO(!CANTIDAD)
                 CODPARAM$(40) = "N-PLANO": DOCPARAM$(40) = FORMATNUM$(XVALO(!PreunitCto), "F12.3")
                 CODPARAM$(41) = "NUOPERS": DOCPARAM$(41) = FORMATNUM$(XVALO(!GASPREPA), "F12.3")
                 CODPARAM$(42) = "NRO-REMI": DOCPARAM$(42) = FORMATNUM$(XVALO(!DISENO), "F12.3")
                 CODPARAM$(43) = "NRO-AENT": DOCPARAM$(43) = FORMATNUM$(XVALO(!Preunit_Arm), "F12.3")
                 CODPARAM$(44) = "CUIT-TRA": DOCPARAM$(44) = FORMATNUM$(XVALO(!Preunit_Cpte), "F12.3")
                 CODPARAM$(45) = "PATE-TRA": DOCPARAM$(45) = FORMATNUM$(XVALO(!CantStencil) * XVALO(!StTotalusd), "F12.3")
                 
                 CODPARAM$(46) = "NUBULTOS": DOCPARAM$(46) = FORMATNUM$(XVALO(!TtOpcion), "F12.3")
               Case 1
                 CODPARAM$(47) = "N-HPROF": DOCPARAM$(47) = XVALO(!CANTIDAD)
                 CODPARAM$(48) = "N-HINSP": DOCPARAM$(48) = FORMATNUM$(XVALO(!PreunitCto), "F12.3")
                 CODPARAM$(49) = "N-ORIGIN": DOCPARAM$(49) = FORMATNUM$(XVALO(!GASPREPA), "F12.3")
                 CODPARAM$(50) = "N-OCOMPR": DOCPARAM$(50) = FORMATNUM$(XVALO(!DISENO), "F12.3")
                 CODPARAM$(51) = "N-AUTENT": DOCPARAM$(51) = FORMATNUM$(XVALO(!Preunit_Arm), "F12.3")
                 CODPARAM$(52) = "N-BOLREC": DOCPARAM$(52) = FORMATNUM$(XVALO(!Preunit_Cpte), "F12.3")
                 CODPARAM$(53) = "N-REMITO": DOCPARAM$(53) = FORMATNUM$(XVALO(!CantStencil) * XVALO(!StTotalusd), "F12.3")
                 CODPARAM$(54) = "NUME-CHE": DOCPARAM$(54) = FORMATNUM$(XVALO(!TtOpcion), "F12.3")
              Case 2
                 CODPARAM$(55) = "UNI-CONJ": DOCPARAM$(55) = XVALO(!CANTIDAD)
                 CODPARAM$(56) = "CAN-CONJ": DOCPARAM$(56) = FORMATNUM$(XVALO(!PreunitCto), "F12.3")
                 CODPARAM$(57) = "CAN-MPRI": DOCPARAM$(57) = FORMATNUM$(XVALO(!GASPREPA), "F12.3")
                 CODPARAM$(58) = "CAN-PLCO": DOCPARAM$(58) = FORMATNUM$(XVALO(!DISENO), "F12.3")
                 CODPARAM$(59) = "CAN-COMP": DOCPARAM$(59) = FORMATNUM$(XVALO(!Preunit_Arm), "F12.3")
                 CODPARAM$(60) = "CANTIDAD": DOCPARAM$(60) = FORMATNUM$(XVALO(!Preunit_Cpte), "F12.3")
                 CODPARAM$(61) = "IMPO-CAN": DOCPARAM$(61) = FORMATNUM$(XVALO(!CantStencil) * XVALO(!StTotalusd), "F12.3")
                 CODPARAM$(62) = "TOTA-CAN": DOCPARAM$(62) = FORMATNUM$(XVALO(!TtOpcion), "F12.3")
             Case 3
                 CODPARAM$(63) = "COL-CONJ": DOCPARAM$(63) = XVALO(!CANTIDAD)
                 CODPARAM$(64) = "CANT-MIN": DOCPARAM$(64) = FORMATNUM$(XVALO(!PreunitCto), "F12.3")
                 CODPARAM$(65) = "CANT-RES": DOCPARAM$(65) = FORMATNUM$(XVALO(!GASPREPA), "F12.3")
                 CODPARAM$(66) = "CANT-ING": DOCPARAM$(66) = FORMATNUM$(XVALO(!DISENO), "F12.3")
                 CODPARAM$(67) = "CANT-DIS": DOCPARAM$(67) = FORMATNUM$(XVALO(!Preunit_Arm), "F12.3")
                 CODPARAM$(68) = "CANT-FAL": DOCPARAM$(68) = FORMATNUM$(XVALO(!Preunit_Cpte), "F12.3")
                 CODPARAM$(69) = "DESTINO": DOCPARAM$(69) = FORMATNUM$(XVALO(!CantStencil) * XVALO(!StTotalusd), "F12.3")
                 CODPARAM$(70) = "PES-CONJ": DOCPARAM$(70) = FORMATNUM$(XVALO(!TtOpcion), "F12.3")
            Case 4
                 CODPARAM$(71) = "TAL-CONJ": DOCPARAM$(71) = XVALO(!CANTIDAD)
                 CODPARAM$(72) = "COD-MPRI": DOCPARAM$(72) = FORMATNUM$(XVALO(!PreunitCto), "F12.3")
                 CODPARAM$(73) = "DES-MPRI": DOCPARAM$(73) = FORMATNUM$(XVALO(!GASPREPA), "F12.3")
                 CODPARAM$(74) = "USO-MPRI": DOCPARAM$(74) = FORMATNUM$(XVALO(!DISENO), "F12.3")
                 CODPARAM$(75) = "TIT-MPRI": DOCPARAM$(75) = FORMATNUM$(XVALO(!Preunit_Arm), "F12.3")
                 CODPARAM$(76) = "LAR-CUTE": DOCPARAM$(76) = FORMATNUM$(XVALO(!Preunit_Cpte), "F12.3")
                 CODPARAM$(77) = "ACH-CUTE": DOCPARAM$(77) = FORMATNUM$(XVALO(!CantStencil) * XVALO(!StTotalusd), "F12.3")
                 CODPARAM$(78) = "LAR-CUCO": DOCPARAM$(78) = FORMATNUM$(XVALO(!TtOpcion), "F12.3")
            End Select
            .MoveNext
       Loop
    End With
    RS.Close
    Set RS = Nothing
    
    CODPARAM$(79) = "OBSERVA": DOCPARAM$(79) = TXTObservCliente.TEXT
      PLZOENTREGA = XVALO(TXTADICIONALARMADOS) + XVALO(TXTCIRCUITOS)
    CODPARAM$(80) = "LAR-MACO": DOCPARAM$(80) = TRIM$(FORMATNUM$(PLZOENTREGA, "F5"))
     TOLERANXIA = 0
     If Option4.Value = True Then TOLERANXIA = XVALO(Text8)
    CODPARAM$(81) = "ACH-CUCO": DOCPARAM$(81) = TRIM$(FORMATNUM$(XVALO(Left(Text30, 2)), "F5"))
    CODPARAM$(82) = "ACH-MACO": DOCPARAM$(82) = TRIM$(FORMATNUM$(TOLERANXIA, "F5"))
    CAPARDOC% = 82
    CANCELPRINT% = 1
    Call PRINTDOC("@" + FORIMPRE$)   ' PLANILLA DE PRODUCCION

    Screen.MousePointer = 11
    NUDOCU$ = SAFETEXT$(NROCOTIZACION)
    DOCUNU& = DOCID&(TIPODOC$, Str$(NROCOTIZACION), NUDOCU$)
    If DOCUNU& < 1 Then
           Screen.MousePointer = 1
           Call COMUNI("No se Encontró el Documento Impreso")
    End If
    Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NRO), Str$(NRO))

    Call LIMPIATETAB


End Sub

Sub PRINTCOTIZAMULT(NROCOT&)
  NC = XVALO(Text18)
  PROSP$ = TRIM$(Label26)
  VERX% = XVALO(Text12)
  FORIMPRE$ = TRIM$(Control$("", "COTIDAMU"))
  If FORIMPRE$ = "" Then
     Call COMUNI("Falta Configurar Formulario de Impresión de Cotización")
     Exit Sub
  End If
  '
  If XVALO(NROCOT&) < 1 Then
      Call COMUNI("Falta N° De Cotización")
      Exit Sub
  End If
  '
  If NC < 1 And PROSP$ = "" Then
      Call COMUNI("Esta Cotización No Tiene Asignado Un Número de Cliente")
      Exit Sub
  End If
  '
  'si tiene n cliente real imprime un frm (los 2 son iguales pero uno toma datos de cliente del deudor12 y otro de prospect)
  If NC > 0 Then
     FILTX$ = TRIM$(Str$(NROCOT&) & ";" & TRIM$(Str$(NC)) & ";" & TRIM$(Str$(VERX%)))
     Call STDFORM(FORIMPRE$, FILTX$)
     Exit Sub
  End If
  'si tiene asignado un prospect imprime otro (los 2 son iguales pero uno toma datos de cliente del deudor12 y otro de prospect)
  If PROSP$ <> "" Then
     FORIMPRE$ = TRIM$(Control$("", "COTIMUPR"))
     If FORIMPRE$ = "" Then
       Call COMUNI("Falta Configurar Formulario de Impresión de Cotización")
       Exit Sub
     End If

     FILTX$ = TRIM$(Str$(NROCOT&) & ";" & PROSP$ & ";" & TRIM$(Str$(VERX%)))
     Call STDFORM(FORIMPRE$, FILTX$)
     Exit Sub
  End If
  

End Sub

Function VALIDACIONGRABACION%()
   VALIDACIONGRABACION% = 1
   
   NCLIE = XVALO(Text18)
   NPROSPECT = TRIM$(Label26)
   If XVALO(Text1) < 1 Then
      Call COMUNI("Falta N° De Cotización")
      VALIDACIONGRABACION% = -1
   End If
   If Option4.Value = True Then
      If XVALO(Text8) <= 0 Then
         Call COMUNI("Si Hay Marca de Acepta Excedente, Debe Informar Porcentaje Excedente")
         Call PINTATEXT(Text8)
         VALIDACIONGRABACION% = -1
      End If
   End If
   '
   If TRIM$(Text4) <> "" Then ' si tiene ingresado el código se valida
     CI1% = CODINT(Text4)
     If CI1% < 1 Then
         Call COMUNI("El Código DAI-ICHI No Es Válido")
         Call PINTATEXT(Text4)
         VALIDACIONGRABACION% = -1
     End If
   End If
   
   If TRIM$(Text3) = "" Then
         Call COMUNI("Imposible Grabar Cotización Activa\No se Ha Ingresado el Modelo")
         Call PINTATEXT(Text3)
         VALIDACIONGRABACION% = -1
   End If
   '
   If TRIM$(Text30) = "" Then
         Call COMUNI("Imposible Grabar Cotización Activa\No se Ha Ingresado la Condición de Pago")
         Call PINTATEXT(Text30)
         VALIDACIONGRABACION% = -1
   End If
   '
   If NPROSPECT = "" And NCLIE < 1 Then
         Call COMUNI("Imposible Registrar !!!\Falta Ingresar Cliente")
         Call PINTATEXT(Text18)
         VALIDACIONGRABACION% = -1
   End If
   
   
   If NCLIE > 0 Then
    If Abs(STATCLI%(NCLIE)) < 1 Then
         Call COMUNI("Imposible Registrar !!!\Numero de Cliente Inexistente")
         VALIDACIONGRABACION% = -1
    End If
   End If
   '
   If NPROSPECT <> "" Then
         NUFICH& = XVALO(Mid$((VALOBADA("BSEGCOM", "CODI_CLI", "CODI_CLI ='" & NPROSPECT & "'")), 4))
         If NUFICH& < 1 Then
           Call COMUNI("Imposible Registrar !!!\Cliente Prospect Seleccionado Inexistente")
           Call PINTATEXT(Text18)
           VALIDACIONGRABACION% = -1
         End If
   End If
   '
   FE% = FECHANUM(Text6)
   If FE% > 30 Then
      FE1% = FECHANUM(TXTFECHA)
      If FE1% > FE% Then
          Call COMUNI("Imposible Registrar !!!\Fecha de Entrega Anterior a Fecha de Cotización")
          Call PINTATEXT(Text6)
          VALIDACIONGRABACION% = -1
      End If
   End If
   '
   If FE% > 30 Then
      FEENT% = FECHANUM(Text7)
      If FEENT% > FE% Then
          Call COMUNI("Imposible Registrar !!!\Fecha de Entrega Anterior a Fecha de Orden de Compra")
          Call PINTATEXT(Text6)
          VALIDACIONGRABACION% = -1
      End If
   End If
   '
   
   FE1% = FECHANUM(TXTFECHA)
   '
   If FE1% < 1 Then
      Call COMUNI("Falta Fecha o Formato No Válido\(dd/mm/aa) ")
      Call PINTATEXT(TXTFECHA)
      VALIDACIONGRABACION% = -1
   End If

End Function

Private Sub Check1_Click()
   MODIFIC% = 1

End Sub

Private Sub Check10_Click()
    MODIFIC% = 1
End Sub

Private Sub Check11_Click()
    MODIFIC% = 1
End Sub

Private Sub Check12_Click()
    MODIFIC% = 1
End Sub

Private Sub Check13_Click()
  If Check13.Value = 1 Then
     TXTCIRCUITOS = 6
  Else
  
  End If
  MODIFIC% = 1
End Sub

Private Sub Check14_Click()
   If Check14.Value = 1 Then
        TXTCANTIDAD2.Locked = False
        TxtLargo2.Locked = False
        TxtAncho2.Locked = False
        TXTCANTIDAD2.BackColor = &HFFFFFF
        TxtLargo2.BackColor = &HFFFFFF
        TxtAncho2.BackColor = &HFFFFFF

        Call TxtSupLoteM22_Change
        TXTObservaPlanta = " La Cantidad se Refiere a Paneles "
        Me.TXTObservCliente = " Cotizado Por Paneles."
'        TxtCantidad.SetFocus
   End If
End Sub

Private Sub Check15_Click()
   MODIFIC% = 1
End Sub

Private Sub Check16_Click()
    MODIFIC% = 1
End Sub

Private Sub Check17_Click()
MODIFIC% = 1
End Sub

Private Sub Check18_Click()
MODIFIC% = 1
End Sub

Private Sub Check19_Click()
MODIFIC% = 1
End Sub

Private Sub Check2_Click()
   MODIFIC% = 1

End Sub

Private Sub Check20_Click()
  MODIFIC% = 1
End Sub

Private Sub Check21_Click()
     If Check21.Value = 1 Then
'        Call TxtSupLoteM2_Change
        TXTObservaPlanta = " La Cantidad se Refiere a Circuitos "
        Me.TXTObservCliente = " Cotizado Por Circuitos."
      Else
'        Call TxtSupLoteM2_Change
      End If
End Sub

Private Sub Check3_Click()
   MODIFIC% = 1

End Sub

Private Sub Check4_Click()
   MODIFIC% = 1

End Sub


Private Sub Check5_Click()
   If Check5.Value = 0 Then
       Text10 = 3
   End If
   MODIFIC% = 1
End Sub

Private Sub Check6_Click()
   If Check6.Value = 0 Then
       Text10 = ""
   End If
End Sub

Private Sub Check7_Click()
    MODIFIC% = 1
End Sub

Private Sub Check8_Click()
    MODIFIC% = 1
End Sub

Private Sub Check9_Click()
    MODIFIC% = 1
End Sub

Private Sub CMD1_Click()
 SELKARA = SELCARAS
   If SELKARA > 0 Then
      Me.TXTNUMCARA = SELKARA
   End If
End Sub







Private Sub CMD10_Click()
      CD1.Filter = "(*.*)|*.*|"
      CD1.ShowOpen
      Ruta$ = CD1.FileName
      If TRIM(Ruta$) = "" Then Exit Sub
      '
20    DESCRI$ = TRIM$(InputBox$("Ingrese Texto ", "Guardar Como", DESCRI$))
      If DESCRI$ = "" Then Exit Sub
      If Len(DESCRI$) > 24 Then
        Call COMUNI("Longitud de Descripción \Supera La Cantidad Máxima de Caracteres (24)")
        GoTo 20
      End If
      
      'AGREGAR INFORMACION
      Call REPLA(TX$, DESCRI$, 1, 24)
      Call REPLA(TX$, Ruta, 32, 512)
      List1.AddItem TX$

     
End Sub

Private Sub CMD2_Click()
  If EXISTE(LUDAT$ + "\TABTIMB.RFS") > 0 Then
      Archivo$ = "TABTIMB"
     Else
      Archivo = "TATIMAT"
   End If
   '
   Call SELECHO(Archivo$)
   VALONUE$ = VALACT1$(Archivo$)
   If VALONUE$ <> "" Then
      TXTNUMTMAT.TEXT = VALONUE$
   End If
End Sub
Private Sub CMD7_Click()
   PCB = SELEPCB
   If PCB > 0 Then
      Me.Text9 = PCB
   End If
End Sub









Private Sub CMD9_Click()
    If List1.ListIndex >= 0 Then
       List1.RemoveItem (List1.ListIndex)
    End If
End Sub

Private Sub Command1_Click()
   NCLIE = ListaSeleccionClientes("", True)
   If NCLIE < 1 Then Exit Sub
   '
   NC = NCLIE
   Text18.TEXT = TRIM$(Str$(NC))
   Call CARDACLI
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   Text4 = ListaSeleccionArticulos("", True)
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   If Check21.Value < 1 Then
      Call COMUNI("Atención!! \Debe Habilitar el Check Paneles\Para Calcular Cotización Con Paneles")
   End If
   If XVALO(TxtCantidad) <= 0 Then
      Call COMUNI("Atención!! \Falta Ingresar Cantidad de Paneles")
      Call PINTATEXT(Me.TxtCantidad)
      GoTo 99
   End If
   If XVALO(TxtLargo) <= 0 Then
      Call COMUNI("Atención!! \Falta Ingresar el Largo del Panel")
      Call PINTATEXT(Me.TxtLargo)
      GoTo 99
   End If
   If XVALO(TxtAncho) <= 0 Then
      Call COMUNI("Atención!! \Falta Ingresar el Ancho del Panel")
      Call PINTATEXT(Me.TxtAncho)
      GoTo 99
   End If
   '
   If TRIM$(TXTCARAS) = "" Or TRIM$(TXTTIPOMATERIAL) = "" Or TRIM$(TXTESPESOR) = "" Or TRIM$(TXTESPCOBRE) = "" Or TRIM$(TXTPTH) = "" Then
      Call COMUNI("Atención!! \Falta Ingresar Datos")
      If TRIM$(TXTCARAS) = "" Then
         Call PINTATEXT(TXTNUMCARA)
         
      ElseIf TRIM$(TXTTIPOMATERIAL) = "" Then
         Call PINTATEXT(Me.TXTNUMTMAT)
         
      ElseIf TRIM$(TXTESPESOR) = "" Then
         Call PINTATEXT(Me.TXTNUMESPMAT)
         
      ElseIf TRIM$(TXTESPCOBRE) = "" Then
         Call PINTATEXT(Me.TXTNUMESPCOBRE)
         
      ElseIf TRIM$(TXTPTH) = "" Then
         Call PINTATEXT(Me.TXTNUMPTH)
      End If
      GoTo 99
   End If

   Call Actualizar_CostoCircuito
   Call ACTUINGRESOS_MANUALES

'   Label3 = "Último Modo Cálculo: Circuito"
'   Label3.Visible = True
99 Command11.Enabled = True

End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   If Check14.Value < 1 Then
      Call COMUNI("Atención!! \Debe Habilitar el Check Paneles\Para Calcular Cotización Con Paneles")
      GoTo 99
   End If
   If XVALO(TXTCANTIDAD2) <= 0 Then
      Call COMUNI("Atención!! \Fala Ingresar Cantidad de Paneles")
      Call PINTATEXT(Me.TXTCANTIDAD2)
      GoTo 99
   End If
   If XVALO(TxtLargo2) <= 0 Then
      Call COMUNI("Atención!! \Fala Ingresar el Largo del Panel")
      Call PINTATEXT(Me.TxtLargo2)
      GoTo 99
   End If
   If XVALO(TxtAncho2) <= 0 Then
      Call COMUNI("Atención!! \Fala Ingresar el Ancho del Panel")
      Call PINTATEXT(Me.TxtAncho2)
      GoTo 99
   End If
   '
   If TRIM$(TXTCARAS) = "" Or TRIM$(TXTTIPOMATERIAL) = "" Or TRIM$(TXTESPESOR) = "" Or TRIM$(TXTESPCOBRE) = "" Or TRIM$(TXTPTH) = "" Then
      Call COMUNI("Atención!! \Falta Ingresar Datos")
      If TRIM$(TXTCARAS) = "" Then
         Call PINTATEXT(TXTNUMCARA)
         GoTo 99
         
      ElseIf TRIM$(TXTTIPOMATERIAL) = "" Then
         Call PINTATEXT(Me.TXTNUMTMAT)
         GoTo 99
         
      ElseIf TRIM$(TXTESPESOR) = "" Then
         Call PINTATEXT(Me.TXTNUMESPMAT)
         GoTo 99
         
      ElseIf TRIM$(TXTESPCOBRE) = "" Then
         Call PINTATEXT(Me.TXTNUMESPCOBRE)
         GoTo 99
         
      ElseIf TRIM$(TXTPTH) = "" Then
         Call PINTATEXT(Me.TXTNUMPTH)
         GoTo 99
      End If
   End If
   
   If XVALO(TxtSupLoteM22) <= XVALO(TxtSupLoteM2) Then
      Call COMUNI("La Superficie de Paneles debe Ser Superior a la Superficie de Circuitos")
      GoTo 99
   End If
   '
   'AHORA ESTA TOMANDO EL VALOR EN BASE AL CIRCUITO INDIVIDUAL PONDERADO ES DECIR LA SUPERFICIE INDIVIDUAL PONDERADA POR LAS MEDIDAS DEL PANEL
   'CON ESA MEDIDA TRAE EL VALOR DE LA TABLA LOGARITMICA Y EN BASE A ESTE CALCULA MULTIMPLICANDO POR LA CANTIDAD DE CIRCUITOS
   'ESTO LO HABIA APROBADO MARCELO PERO EN EL WORD QUE ENVIO ES CONFUSO POR QUE PARA LOS CALCULOS DEBERIA MULTIPLICAR CONTRA LA CANTIDAD DE PANEELS
   'ENTONCES EN SU MOMENTO PEDIR UNA ESPECIFICACION DE LOS MISMOS, SI EL PRECIO UNITARIO DEBE SER EL DEL PANEL Y LUEGO MULTIMPOICAR POR LOS MISMOS.
   Call Actualiza_CostoPanel
   Call ACTUINGRESOS_MANUALES

   Label3 = "Último Modo Cálculo: Panel"
   Label3.Visible = True
   
99 Command12.Enabled = True
End Sub


Private Sub Command2_Click()
 SELSINOX = SELSINO
   If SELSINOX > 0 Then
      Me.Text5 = SELSINOX
   End If

End Sub

Private Sub Command3_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text6.TEXT = VALACT1$("SELFECHA")
   End If
End Sub


Private Sub Command4_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      TXTFECHA.TEXT = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Command5_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text7.TEXT = VALACT1$("SELFECHA")
   End If
    
End Sub

Private Sub Command7_Click()
   Call SELECHO("CONPAG")
   VDEVU$ = VALACT1$("CONPAG")
   If TRIM$(VDEVU$) <> "" Then
      Text17.TEXT = VDEVU$
   End If
End Sub

Private Sub Command8_Click()
    Call SELECHO("INDIPROS") ' PERMITE CARGA PRESUPUESTO A PARTIR DE PROSPECTOS
    NCLIE$ = TRIM$(VALACT1$("INDIPROS"))
    If NCLIE$ <> "" Then
      Text18.TEXT = "0"
      Label26.Caption = TRIM$(NCLIE$)
      NCLIE = XVALO(VALOBADA("DEUDOR12", "Nume_Cli", "Codi_Cli = '" & NCLIE$ & "'"))
      If NCLIE > 0 Then Text18 = NCLIE
      Call CARDACLI
      Text19.SetFocus
    End If

End Sub

'
Sub CARDACLI()
    '
    Screen.MousePointer = 11
    NC = XVALO(Text18.TEXT)
    If NC > 0 Then
      'If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
      If CantRegistros("deudor12", "nume_cli=" & NC) > 0 Then
        Text2.TEXT = rasocli$(NC)
        Label26 = ""
        Text19 = ""
        '
      End If
    ElseIf NC = 0 Then ' CARGA DATOS DEL PROSPECTO
'        NCLIE$ = TRIM$(Label26)
'        Text2.TEXT = VALOBADA("BSEGCOM", "Raso_Cli", "Codi_Cli = '" & NCLIE$ & "'") 'SELECCIONA LA RAZON SOCIAL DEL PROSPECT
         Text2.TEXT = ""
    End If
99  Screen.MousePointer = 1
    '
End Sub


Private Sub Command9_Click()
   Command9.Enabled = False
   NCOT& = XVALO(Text1)
   If NCOT& < 1 Then
       Call COMUNI("Debe Seleccionar Una Cotización Existente\Para Poder Crear Una Nueva A Partir de la Activa")
       GoTo 99
   End If
   '
   Erase RESP_ZELE_VECT
   NRO = SELEC_COTIZACION(NCOT&)
   If NRO < 1 Then GoTo 99
   
   If XVALO(RESP_ZELE_VECT(1)) < 1 Then GoTo 99
   
   If NRO > 0 Then
     Text1 = ""
     lblNuevaCotizacion.Visible = False
     Text1 = NRO
   End If
   '
   OPX% = COMALTER%("Opciones de Trabajo\\Cotización Sobre Mismo Código (Dif.Cantidad)\Diferente Código")
   MODOX% = OPX%
   LBLMODO = MODOX%
   VERX& = CANT_VERSIONES_EXISTENTES&(NCOT&)
   If VERX& > 1 Then
      MODOCOTI% = MODOCOTIZA%(NCOT&)
      If MODOCOTI% <> MODOX% Then
        If MODOX% = 1 Then
           Call COMUNI("Imposible Agragar Cotización de Mismo Código Diferente Cantidad\Esta Cotización Yá Tiene Agregada Cotización Diferente Código")
        ElseIf MODOX% = 2 Then
           Call COMUNI("Imposible Agragar Cotización de Diferente Código\Esta Cotización Yá Tiene Agregada Cotización Para Mismo Código Diferente Cantidad")
        End If
        GoTo 99
      End If
   End If
   '
   If VERX& > 4 And MODOX% = 1 Then
      Call COMUNI("Imposible Agragar Cotización de Mismo Código Diferente Cantidad\Esta Cotización Yá Tiene Generada el Límite de 5")
      GoTo 99
   End If
   T18 = Text18
   L26 = Label26
   T19 = Text19
   T7 = Text7
   TXTCHA = FECHATEX$(HOY(HO$))
   
   If OPX% < 1 Then GoTo 99
   
   If OPX% > 1 Then Call BLANFORMU
'      CANT_VERS& = CantRegistros&("COTIDAICHI", "NRO_COTIZA = " & XVALO(Text1) & " And MODELO = '" & TRIM$(Text3) & "'")
'      If CANT_VERS& > 1 Then
'         Call COMUNI("Imposible Generar Cotización con Diferente Código a Partir de la Activa\La Misma ya Posee Más de Una Cotización a Mismo Modelo ")
'     'ACA DEBERIA IMPRMIR SI O SI LA OTRA COTIZACIÓN O BIEN SOLO LA QUE ESTA ACTIVA.
'     'VER EN TODO CASO DE ARMAR OTRO FRM.
'      End If

      
'   End If
   '
   
   Text12 = VERX&
   
   '
   Text18 = T18
   Label26 = L26
   Text19 = T19
   Text7 = T7
   TXTFECHA = TXTCHA
   
   
99 Command9.Enabled = True
End Sub

Private Sub Form_Load()
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    NOPRESENTARMENSAJE% = 0
    Call ABRECONEXION
    Call CREATABLA_ENCAPRF
    Call CREATABLA_COTIDAICHI
    Call GENERAR_ESTRUCTURAS_SQL
    If VALIDARINTEGRIDAD_TABLA_CALCULO_LOGARITMICA% > 0 Then
       Call COMUNI("Atención !!! Se ha Detectado Ambiguedad para la Tabla de Configuración")
    End If
    Option4.Value = True
    Check21.Value = 1
    Text8 = 10
    TXTCIRCUITOS = 12
    Check5.Value = 1
    Check7.Value = 1
    Check16.Value = 1
    VALODEFECTO% = 0
   
    lblNuevaCotizacion.Visible = True
    Call BLANFORMU
    Text12 = 0
    Text1 = SAFETEXT$(NumNuevaCotizaDaichi)
    TXTFECHA = FECHATEX$(HOY(HO$))
    MODIFIC% = 0
    PRESENTANDOINFOBD% = 0
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        Call CIERRARCH("*.*")
        Call FINAL("")
End Sub

Private Sub Form_Unload(Cancel As Integer)
        Call CIERRARCH("*.*")
        Call FINAL("")

End Sub

Private Sub Label26_Change()
   COCLI$ = TRIM$(Label26)
   CONDI$ = "CODI_CLI = '" & COCLI$ & "'"
   Label4 = VALOBADA("BSEGCOM", "raso_cli", CONDI$, "NOMESS")
   MODIFIC% = 1

End Sub

Private Sub List1_DblClick()
   Linea$ = List1.List(List1.ListIndex)
   Archi$ = TRIM$(Mid$(Linea$, 25))
   If (EXISTE%(Archi$)) <> 1 Then
       Call COMUNI("Archivo Inexistente")
       Exit Sub
     Else
        Call ShellExecute(Me.hwnd, "Open", Archi$, "", "", 1)
    End If
    MODIFIC% = 1

End Sub


Private Sub mnuAbmSec_Click()
    ABMSECUE.Show 1
End Sub

Private Sub mnuAproCot_Click()
    If DERACCE%("APROCOTI", "Aprobación de Cotización Activa") < 2 Then Exit Sub
    '
    NROCOT& = XVALO(Text1)
    CI1% = CODINT(Text4)
    If CI1% > 0 Then
        If VALIDACIONGRABACION% < 0 Then
          Call COMUNI("Modifique los Valores No Aceptados y Vuelva a Intentar")
          Exit Sub
        End If
        
        A = TRIM$(rasocli$(-1)) ' REFRESCA RECORDSET
        NC = XVALO(Text18)
        If NC > 0 Then
           If TRIM$(rasocli$(NC)) = "" Then
             Call COMUNI("Nro de Cliente No Válido o Inexistente")
             Call PINTATEXT(Text18)
             Exit Sub
           End If
        End If
        ' SELECT * FROM COTIDAICHI WHERE NRO_COTIZA = 122 AND VRS_COTIZA = 0 AND NROPED > 0
        CONDI$ = " NRO_COTIZA = " & NROCOT&
        CONDI$ = CONDI$ + " AND VRS_COTIZA = " & XVALO(Text12)
        TIENPEDIDO& = XVALO(VALOBADA("COTIDAICHI", "NROPED", CONDI$))
        If TIENPEDIDO& > 0 Then
          Call COMUNI("Esta Cotización ya Fué Aprobada y Generó un Pedido (" & TIENPEDIDO& & ")")
          Exit Sub
        End If
        '
        'VALIDACIÓN SI TIENE ES MULTIPLE DE CODIGOS Y SI ALGUNO NO TIENE GENERADO EL CODIGO DAI-CHI
        COTDIFCOD% = COTIZACIONDIFERENTESCODIGOS%(NROCOT&, RET_MODELO$)
        If COTDIFCOD% < 0 Then
          Call COMUNI("Esta Cotización Tiene Modelos: (" & RET_MODELO$ & ") Sin Códigos Dai-Ichi")
          Exit Sub
        End If
        NOPRESENTARMENSAJE% = 1
        Call mnuGuardar_Click
        CONDI$ = "NRO_COTIZA = " & NROCOT& & " AND VRS_COTIZA = " & XVALO(Text12)
        Call ACTUREGISTRO("COTIDAICHI", "ESTADO", CONDI$, "VE", REGAF&, "NOMESS")
        Call GeneraPedidoDeCotización(NC, Ret_NroPedido&, NROCOT&, XVALO(Text12), FECHANUM(Text6))
        NOPRESENTARMENSAJE% = 0
        Call COMUNI("Transacción Realizada")

    Else
        Call COMUNI("Atención!!! Código No Válido o Inexistente")
        Exit Sub
    End If
    
    lblestado = "VE"
End Sub

Private Sub mnuEditCotización_Click()

   NRO = SELEC_COTIZACION
   If NRO > 0 Then
     ESTADOX$ = TRIM$(RESP_ZELE_VECT(12))
     lblestado = ESTADOX$
     If ESTADOX$ = "VE" Then
       Call COMUNI("Atención La Cotización Seleccionada Ha sido Aprobada, Deberá Tener Derechos Para Grabar la misma.")
     End If
     If ESTADOX$ = "O.C." Then
       Call COMUNI("Atención La Cotización Seleccionada Tiene Asignada Orden de Compra , Deberá Tener Derechos Para Grabar la misma.")
     End If
     
     Text1 = ""
     lblNuevaCotizacion.Visible = False
     Text12 = XVALO(RESP_ZELE_VECT(17))
     Text1 = NRO
   End If
End Sub
 Function SELEC_COTIZACION(Optional NROCOT&) As Integer
    Erase RESP_ZELE_VECT
    CAMPOS$ = "N°Cotiz./F9;Versión/F9;Fecha/D8;Cant./F8.1;Pre-Dm2/F12.3;Precio/F12.2;Gasto/F12.2;PTH/A8;Faz/A10"
    CAMPOS$ = CAMPOS$ + ";Term.LC - LS/A12;Sup/F9.3;Est./A4;Código/A20;Modelo/A20;Cta.Clie/F0;Cliente/A24;ID/F9"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, CAMPOS$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, CAMPOS$, GRILLAGRAL, 1)
    FRMZELECHO.Caption = ""
    
    Dim obj As New RECORXET
    Set RS1 = obj.RS_ListaCotización(XVALO(NROCOT&))
    Call Carga_MSF_Recordset(RS1, CAMPOS$, GRILLAGRAL.GRID)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    
    GRILLAGRAL.mnuMenu3.Visible = False
    GRILLAGRAL.Show 1
    SELEC_COTIZACION = XVALO(RESP_ZELE_VECT(1))

End Function
Private Sub mnuGuardar_Click()

   If lblestado.Caption = "VE" Then
      If DERACCE%("GRABCOVE", "Grabación Cotización Aprobado") < 2 Then Exit Sub
   End If
   
   If VALIDACIONGRABACION% < 0 Then Exit Sub
   
   ESNUEVO% = 0
   If lblNuevaCotizacion.Visible = True Then
      If XVALO(Control("", "EDITNUCO")) > 0 Then
        NROCOTIX$ = Text1
        NROCOTICAMBIADO$ = TRIM$(InputBox$("Número de Cotización", "Habilitada la Opción de Edtitar N° de Cotización", NROCOTIX$))
         If NROCOTICAMBIADO$ = "" Then
           Exit Sub
         Else
            Text1 = NROCOTICAMBIADO$
            Text12 = ""
         End If
      End If
      ESNUEVO% = 1
   End If
   
   If XVALO(Text12) > 0 Then
      If CantRegistros&("COTIDAICHI", "NRO_COTIZA = " & XVALO(Text1) & " AND VRS_COTIZA = " & XVALO(Text12)) < 1 Then
         ESNUEVO% = 1
      End If
   End If
   '

   SQLX$ = "SELECT * FROM COTIDAICHI "
   If ESNUEVO% < 1 Then
     SQLX$ = SQLX$ + " WHERE NRO_COTIZA = " & XVALO(Text1)
     If XVALO(Text12) > 0 Then SQLX$ = SQLX$ + " AND VRS_COTIZA = " & XVALO(Text12)
   Else
     SQLX$ = SQLX$ + " WHERE NRO_COTIZA < 1 "
   End If
   
   On Error GoTo ERROR_GUARDAR
   FLEXCONN.BeginTrans
   
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset
   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With RS
     If ESNUEVO% > 0 Then
         If XVALO(Text12) > 0 Then .AddNew: GoTo 30
         If NROCOTICAMBIADO <> "" Then
           'lo hago aqui para tener aprovechar el bloqueo de begin
           If CantRegistros&("COTIDAICHI", "NRO_COTIZA = " & XVALO(NROCOTICAMBIADO), "NOMESS") < 1 Then
             NEWCOTIZ = XVALO(NROCOTICAMBIADO)
           Else
             Call COMUNI("Imposible Grabar Cotización\Número Existente")
             FLEXCONN.RollbackTrans 'rollblakeo para que no de error si ahora guarda
             Exit Sub
           End If
         Else
           NEWCOTIZ = NumNuevaCotizaDaichi
         End If
         
         .AddNew
         !NRO_COTIZA = NEWCOTIZ
         Text1 = NEWCOTIZ
     Else
30       !NRO_COTIZA = XVALO(Text1)
         ESTAT$ = TRIM$(SAFETEXT$(!ESTADO))
     End If
     
     !NUME_CLI = XVALO(Text18)

     !CODI_CLI = Left$(Label26, 12)
     valor$ = "NO": If Check21.Value = 1 Then valor = "SI"
     !CIRCUITO = valor$
     valor$ = "NO": If Check14.Value = 1 Then valor = "SI"
     !PANELIZADO = valor$
     
     !Ctac_Cli = Left(Text19, 48)
     
     !FECHEMISI = CVDAT(CVINT(TXTFECHA))
     
     !FECHCOMPRA = CVDAT(CVINT(Me.Text7))
     
     !FeSolEnt = CVDAT(CVINT(Me.Text6))
     
     !Codigo = TRIM$(Left(Text4, 24))
     
     !cod_inte = CODINT(Text4)
     
     !MODELO = TRIM$(Left$(Text3, 128))
     If Check14.Value = 1 And XVALO(TXTCANTIDAD2) > 0 Then
       !CANTIDAD = XVALO(TXTCANTIDAD2)
       !LARGO = XVALO(Me.TxtLargo2)
       !ANCHO = XVALO(Me.TxtAncho2)
       !SUPERFX1 = XVALO(Me.TxtSupDM22)
       !SUPERFXLOTE = XVALO(Me.TxtSupLoteM22)
     Else
       !CANTIDAD = XVALO(TxtCantidad)
       !LARGO = XVALO(Me.TxtLargo)
       !ANCHO = XVALO(Me.TxtAncho)
       !SUPERFX1 = XVALO(Me.TxtSupDM2)
       !SUPERFXLOTE = XVALO(Me.TxtSupLoteM2)
     End If
     '
     !NUMARMADO = XVALO(Text5)
     !ARMADO = TRIM$(TXTSINO)
     'CARAS
     !CARAS = XVALO(TXTNUMCARA)
     
     !TIMATE = XVALO(TXTNUMTMAT)
     
     !ESPMAT = XVALO(TXTNUMESPMAT)
     
     !ESPCOBRE = XVALO(TXTNUMESPCOBRE)
     
     !PTH = XVALO(TXTNUMPTH)
     
     !PISTAS = XVALO(TXTNUMPISTA)
     
     
     !TEXCARAS = TRIM$(TXTCARAS)
     
     !TEXTIMATE = TRIM$(TXTTIPOMATERIAL)
     
     !TEXESPMAT = TRIM$(TXTESPESOR)
     
     !TEXESPECOBRE = TRIM$(TXTESPCOBRE)
     
     !TEXPTH = TRIM$(TXTPTH)
     
     !TEXPISTAS = TRIM$(TXTPISTAS)
     
     '
     'ESTADO ACATUAL
     valor$ = "COT" 'POR DEFECTO COT DE COTIZACION
     If TRIM$(Text7) <> "" Then valor$ = "O.C." 'SI TIENE FECHA DE COMPRA APLICA O.C.
'     If TRIM$(Text7) = "" and ESTAT$ = = "O.C."  then
     If ESTAT$ <> "" And ESTAT$ <> "COT" And ESTAT$ <> "O.C." Then valor$ = ESTAT$  'SI TENIA UN VALOR DE ESTADO GUARDADO EN  OTRO LUGAR MANTIENE EL MISMO.
     
     !ESTADO = valor$
     
     valor$ = "NO": If Check1.Value > 0 Then valor = "SI"
     
     !FtainfTec = valor$
     valor$ = "NO": If Check2.Value > 0 Then valor = "SI"
     !PrvApMtra = valor$
     
     valor$ = "NO": If Check3.Value > 0 Then valor = "SI"
     !PrvAppgo = valor$
     
     valor$ = "NO": If Check4.Value > 0 Then valor = "SI"
     !NoStdar = valor$
     
     'PROGRAMAPCB
     !PrgPCB = XVALO(Text9)
     
     !TexPrgPCB = Text11
     
     'MASCARA ANTISOLDANTE
     valor$ = "NO": If Check5.Value > 0 Then valor = "SI"
     !LS = valor$
     
     valor$ = "NO": If Check6.Value > 0 Then valor = "SI"
     !lC = valor$
     
     !MASCSOLDANTE = XVALO(Text10)
     
     !TexMASCSOLDANTE = Text13
     
     'LEYENDA COMPONENTES
     valor$ = "NO": If Check8.Value > 0 Then valor = "SI"
     !LCOMP_LS = valor$
     
     valor$ = "NO": If Check7.Value > 0 Then valor = "SI"
     !LCOMP_LC = valor$
     
     !LEYECOM = XVALO(Text16)
     
     !TexLEYECOM = Text14
     
     'EXCEDENTE
     valor$ = "NO": If Option3.Value = True Then valor = "SI"
     !PRDEXACTA = valor$ 'PRODUCCION EXACTA
     valor$ = "NO": If Option4.Value = True Then valor = "SI"
     
     !ACEXCTE = valor$ 'ACEPTA EXCEDENTE
     
     !TOLERANCIA = XVALO(Text8) 'TOLERANCIA
     
     'PLAZOENTREGA Y ADICIONAL
     !PLZ_ENT = XVALO(TXTCIRCUITOS)
     
     !PLZ_ADIC_ARM = XVALO(TXTADICIONALARMADOS)
     
     'ESPECIALES
     valor$ = "NO": If Check9.Value > 0 Then valor$ = "SI"
     !ESP_GRAFITO = valor$ 'GRAFITO
     
     valor$ = "NO": If Check10.Value > 0 Then valor$ = "SI"
     !ESP_GOMA = valor$  'GOMA
     
     valor$ = "NO": If Check11.Value > 0 Then valor$ = "SI"
     !ESP_NI_AU = valor$ 'NIQUEL ORO
     
     valor$ = "NO": If Check12.Value > 0 Then valor$ = "SI"
     !ESP_OTRO = valor$ 'OTRO
     
     !ESP_SUPMM2 = XVALO(Text15) 'OTRO
     
     '
     'CORTE DE PLACA
      valor$ = "NO": If Check16.Value > 0 Then valor$ = "SI"
     !ROUTER = valor$
     
     valor$ = "NO": If Check15.Value > 0 Then valor$ = "SI"
     !SCORING = valor$
     
     'ENTREGA
     valor$ = "NO": If Check13.Value > 0 Then valor$ = "SI"
     !PROTOTIPO = valor$
     
     valor$ = "NO": If Check13.Value > 0 Then valor$ = "SI"
     
     'TEST
     valor$ = "NO": If Check20.Value > 0 Then valor$ = "SI"
     !TEST_ELECT = valor$
     
     valor$ = "NO": If Check19.Value > 0 Then valor$ = "SI"
     !PNLIZADO = valor$
     
     valor$ = "NO": If Check18.Value > 0 Then valor$ = "SI"
     !MASCFONDO = valor$
     
     valor$ = "NO": If Check17.Value > 0 Then valor$ = "SI"
     !PCSO_INI_MASC = valor$
     
     !REFPANELIZADO = Left(TXTPANELIZADO, 24)
     
     !REFMASDEFONDO = Left(TXTMASDEFONDO, 24)
     
     'MONTAJE
     !MTJ_SMDBOT1 = XVALO(TXTSMDBOT1)
     
     !MTJ_SMDTOP1 = XVALO(TXTSMDTOP1)
     
     !MTJ_INSBOT1 = XVALO(TXTINSBOT1)
     
     !MTJ_INSTOP1 = XVALO(TXTINSTOP1)
     
     !MTJ_SMDBOT2 = XVALO(TXTSMDBOT2)
     
     !MTJ_SMDTOP2 = XVALO(TXTSMDTOP2)
     
     !MTJ_INSBOT2 = XVALO(TXTINSBOT2)
     
     !MTJ_INSTOP2 = XVALO(TXTINSTOP2)
     
      
      'RESUMEN
      !GASPREPA = XVALO(Text26)
      
      
      !CONPAG = XVALO(Text17)
      
      !PRC_FJO_DM2 = XVALO(Text31)
      
      !DISENO = XVALO(Text32)
      
      !PRC_DM2 = XVALO(TxtPrcdm2)
      
      !Preunit = XVALO(TxtPrcUnit)
      
      !PrcTtPlca = XVALO(TXTtotPlca)
      
      !PrcTtOpcion = XVALO(TXTPrcTtOpcion)
      
      !PreunitCto = XVALO(TXTPreunitCto)
      
      !Preunit_Arm = XVALO(TXTPreunit_Arm)
      
      !Preunit_Cpte = XVALO(TXTPreunit_Cpte)
      !Prcdm2Costo = XVALO(TxtPrcdm2Costo)
      !PrcUnitCosto = XVALO(TxtPrcUnitCosto)
      !totPlcaCosto = XVALO(TXTtotPlcaCosto)
      !PrcTtOpcionCosto = XVALO(TXTPrcTtOpcionCosto)
      
      !PrcTtPlcaSA = XVALO(TXTPrcTtPlcaSA)
      !TtOpcion = XVALO(TXTTtOpcion)
      !CantStencil = XVALO(TXTCantStencil)
      !StTotalusd = XVALO(TXTStTotalusd)
      !ANTICIPO = XVALO(Text33)
      !SALDO = XVALO(Text34)
      !ObservaAdmin = Me.TXTObservaAdmin
      !NROOCOM = TXTNROOCOM
      !ObservaPlanta = TXTObservaPlanta
      !ObservCliente = TXTObservCliente
      !ObservPenal = TXTObservPenal
      !VRS_COTIZA = XVALO(Text12)
      !CANTICAPAS = XVALO(TXTCANTICAPAS)
      !MODO = XVALO(LBLMODO)
     .Update
   End With
   lblNuevaCotizacion.Visible = False
   
   RS.Close
   Set RS = Nothing
   
   Call GRABARARCHIVOS(XVALO(Text1))
   Call GRABAMODO(XVALO(Text1))
   
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
    Else
      FLEXCONN.CommitTrans
      If NOPRESENTARMENSAJE% < 1 Then
         Call COMUNI("Transacción Realizada")
      End If
    End If

End Sub

Private Sub mnuGuardarComoNuevaCotizacion_Click()
   If XVALO(Text1) > 0 Then
      OPX% = COMALTER%("Confirmación de Grabación de Nueva Cotización en Base a la Cotización Activa\\Cancelar\Si Grabar")
      If OPX% > 1 Then
       lblNuevaCotizacion.Visible = True
        TXTFECHA = FECHATEX$(HOY(HO$))
         TXTNROOCOM = ""
         Text6 = ""
         Text7 = ""
         lblestado.Caption = ""
         Call mnuGuardar_Click
         lblNuevaCotizacion.Visible = False
      End If
   End If
End Sub


Private Sub mnuImpLogar_Click()
   IMPOLOGA.Show 1
End Sub



Private Sub mnuNewCotExistente_Click()

End Sub

Private Sub mnuModIndividual_Click()
   Call printCotiza(XVALO(Text1), XVALO(Text12))
End Sub

Private Sub mnuNewCotiz_Click()
'   OPX% = COMALTER%("Nueva Cotización\\Cotización desde 0\Cotización A Partir de una Existente\Cancelar")
'   If OPX% = 1 Then
        lblNuevaCotizacion.Visible = True
        Text18 = ""
        Text2 = ""
        Call BLANFORMU
        Text12 = 0
        Text1 = SAFETEXT$(NumNuevaCotizaDaichi)
        TXTFECHA = FECHATEX$(HOY(HO$))
        MODIFIC% = 0
'   Else
'        NRO = SELEC_COTIZACION
'        If NRO > 0 Then
'          Text1 = SAFETEXT$(NumNuevaCotizaDaichi)
'          Text12 = XVALO(RESP_ZELE_VECT(17))
'
'        End If
'   End If

End Sub

Private Sub mnuPrintMultiple_Click()
 Call PRINTCOTIZAMULT(XVALO(Text1))
End Sub
Private Sub mnuSegCom_Click()
   If EXISTE%("SEGCOM07.EXE") > 0 Then
       Call CONECRUN("SEGCOM07", "Seguimiento Comercial")
   End If
End Sub

Private Sub mnusetup_Click()
   FSCOTDAICHI.Show 1
End Sub

'Private Sub mnuTablaAuxiliares_Click()




'  Call GENERAR_ESTRUCTURAS_SQL
'  Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
'  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
'  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"
'
'  Dim obj As New RECORXET
'  Set RS1 = obj.RS_ENCAPRFLISTAS()
'  Call Carga_MSF_Recordset(RS1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.msf2)
'  Set obj = Nothing
'  RS1.Close
'  Set RS1 = Nothing
'  EPAC$ = UCase$(TRIM$(EMPREAC$))
'  '
'
'  FRMZELECHO.Width = 8500
'  FRMZELECHO.Show 1
'  IDLIST& = XVALO(RESP_ZELE_VECT(1))
'  If InStr(EPAC$, "ENVAPLAST") > 0 Then
'    LIXTA$ = TRIM$(RESP_ZELE_VECT(2))
'    If LIXTA$ = "MATERIALES ENVAPLAST" Then
'       Call COMUNI("Atención !!! Editar Tabla de Familias Desde Maestro de Artículos")
'       Exit Sub
'    End If
'  End If
'
'  Call ABM_TABLA(IDLIST&)
'
'End Sub
'
'
'
Private Sub mnuvisualiza_Click()
10  NRO = SELEC_COTIZACION

    If NRO < 1 Then Exit Sub
    
    TIDOCUM$ = "COTIZACION "

    Screen.MousePointer = 11
    NUDOCU$ = SAFETEXT$(NRO)
    DOCUNU& = DOCID&(TIDOCUM$, Str$(NRO), NUDOCU$)
    If DOCUNU& < 1 Then
           Screen.MousePointer = 1
           Call COMUNI("No se Encontró el Documento Impreso")
    End If
    Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NRO), Str$(NRO))
      '
    GoTo 10
      '
99
End Sub


Private Sub mnuTablaAuxiliares_Click()
  If DERACCE%("TABAUXGE", "ACCESO A TABLAS AUXILIARES GENERALES") < 2 Then Exit Sub
  
  Call MOSTRARTABLASCONFIGURADAS(Me.Name)
End Sub


Private Sub mnuTablatmat_Click()
   If DERACCE%("STPMAST", "Configuracion de Setup de Articulos") = 2 Then
      Call CARTATIMA
   End If

End Sub

Private Sub Option3_Click()
   If Option3.Value = True Then Text8 = ""
   MODIFIC% = 1

End Sub

Private Sub Option4_Click()
'  If Option4.Value = True Then Text8 = 10
  MODIFIC% = 1

End Sub

Private Sub Text1_Change()
'   Static NROCOTIZA
'   If NROCOTIZA > 0 Then
'    If MODIFIC% = 1 And XVALO(Text1) > 0 Then
'         If COMALTER%("Se han Modificado Datos de La Cotización Activa\\Registrar Cambios\Cancelar") = 1 Then
'           Call mnuGuardar_Click
'           MODIFIC% = 0
'         End If
'    End If
'   End If
'   NROCOTIZA = XVALO(Text1)
'   Call BLANFORMU
'   MODIFIC% = 0
   
   PRESENTANDOINFOBD% = 1
   ID& = XVALO(RESP_ZELE_VECT(17))
   
   If CantRegistros("COTIDAICHI", "NRO_COTIZA = " & XVALO(Text1) & " And COTIDAICHI_ID = " & ID&, "NOMESS") < 1 Then Exit Sub
   '
   If Check14.Value > 0 And XVALO(TXTCANTIDAD2) <= 0 Then
          Call COMUNI("Imposible Guardar Cotización\Está Seleccionada la Opción Panelizado\Y No Se Ingresó Cantidad")
          Exit Sub
   End If
   If Check14.Value > 0 And XVALO(TxtLargo2) <= 0 Then
          Call COMUNI("Imposible Guardar Cotización\Está Seleccionada la Opción Panelizado\Y No Se Ingresó el Largo")
          Exit Sub
   End If
   If Check14.Value > 0 And XVALO(TxtAncho2) <= 0 Then
          Call COMUNI("Imposible Guardar Cotización\Está Seleccionada la Opción Panelizado\Y No Se Ingresó el Ancho")
          Exit Sub
   End If

   Label1.Caption = ID&
   SQLX$ = "SELECT * "
   '
   SQLX$ = SQLX$ + " , Case PANELIZADO "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END PANELIZADO"
   
   SQLX$ = SQLX$ + " , Case CIRCUITO "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END CIRCUITO "
   
   SQLX$ = SQLX$ + " , Case FtainfTec "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END FtainfTec"
   
   SQLX$ = SQLX$ + " , Case PrvApMtra "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END PrvApMtra"
   
   SQLX$ = SQLX$ + " , Case PrvAppgo "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END PrvAppgo"
   
   SQLX$ = SQLX$ + " , Case NoStdar "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END NoStdar "
   
   SQLX$ = SQLX$ + " , Case LS "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END LS "
   
   SQLX$ = SQLX$ + " , Case lC "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END lC "
   
   SQLX$ = SQLX$ + " , Case LCOMP_LS "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END LCOMP_LS "
   
   SQLX$ = SQLX$ + " , Case LCOMP_LC "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END LCOMP_LC "
   
   SQLX$ = SQLX$ + " , Case PRDEXACTA "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END PRDEXACTA"
   
   SQLX$ = SQLX$ + " , Case ACEXCTE "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END ACEXCTE "
   
   SQLX$ = SQLX$ + " , Case ESP_GRAFITO "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END ESP_GRAFITO "
   
   SQLX$ = SQLX$ + " , Case ESP_GOMA "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END ESP_GOMA"
   
   SQLX$ = SQLX$ + " , Case ESP_NI_AU "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END ESP_NI_AU "
   
   
   SQLX$ = SQLX$ + " , Case ESP_OTRO "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END ESP_OTRO "
   
   SQLX$ = SQLX$ + " , Case TEST_ELECT "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END TEST_ELECT "

   SQLX$ = SQLX$ + " , Case MASCFONDO "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END MASCFONDO "
   
   SQLX$ = SQLX$ + " , Case PNLIZADO "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END PNLIZADO "
   
   SQLX$ = SQLX$ + " , Case PCSO_INI_MASC "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END PCSO_INI_MASC "
   
   SQLX$ = SQLX$ + " , Case ROUTER "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END ROUTER "
   
   SQLX$ = SQLX$ + " , Case SCORING "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END SCORING "
   
   SQLX$ = SQLX$ + " , Case PROTOTIPO "
   SQLX$ = SQLX$ + " WHEN 'SI' THEN '1'"
   SQLX$ = SQLX$ + " WHEN 'NO' THEN '0'"
   SQLX$ = SQLX$ + " END PROTOTIPO"
   
   
   SQLX$ = SQLX$ + " FROM COTIDAICHI WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NRO_COTIZA = " & XVALO(Text1)
   SQLX$ = SQLX$ + " And COTIDAICHI_ID = " & ID&

   
   Set RS = READSET(SQLX$)
   
   With RS
      Check21.Value = XVALO(!CIRCUITO)
      Check14.Value = XVALO(!PANELIZADO)
      Text18 = XVALO(!NUME_CLI)
   
      Label26 = SAFETEXT$(!CODI_CLI)
   
      Text19 = SAFETEXT$(!Ctac_Cli)
   
      TXTFECHA = FECHATEX(CVINT(!FECHEMISI))
      
      VALFECH% = CVINT(!FECHCOMPRA)
      If VALFECH% > 0 Then
        FESOL$ = FECHATEX$(CVINT(!FECHCOMPRA))
        If Year(FESOL$) = 1980 Then FESOL$ = ""
      Else
        FESOL$ = ""
      End If
      Text7 = FESOL$
      
      VALFECH% = CVINT(!FeSolEnt)
      If VALFECH% > 0 Then
        FESOL$ = FECHATEX$(CVINT(!FeSolEnt))
        If Year(FESOL$) = 1980 Then FESOL$ = ""
      Else
        FESOL$ = ""
      End If
      
    
     Text4 = SAFETEXT$(!Codigo)
    
     Text3 = SAFETEXT$(!MODELO)
     If Check14.Value = 1 Then
          TXTCANTIDAD2 = FORMATNUM$(XVALO(!CANTIDAD), "F10.0")
          TxtLargo2 = FORMATNUM$(XVALO(!LARGO), "F10.3")
          TxtAncho2 = FORMATNUM$(XVALO(!ANCHO), "F10.3")
          TxtSupDM22 = FORMATNUM$(XVALO(!SUPERFX1), "F10.3")
          TxtSupLoteM22 = FORMATNUM$(XVALO(!SUPERFXLOTE), "F10.3")
      Else
          TxtCantidad = FORMATNUM$(XVALO(!CANTIDAD), "F10.0")
          TxtLargo = FORMATNUM$(XVALO(!LARGO), "F10.3")
          TxtAncho = FORMATNUM$(XVALO(!ANCHO), "F10.3")
          TxtSupDM2 = FORMATNUM$(XVALO(!SUPERFX1), "F10.3")
          TxtSupLoteM2 = FORMATNUM$(XVALO(!SUPERFXLOTE), "F10.3")
      End If
      Text5 = XVALO(!NUMARMADO)
      TXTSINO = SAFETEXT$(!ARMADO)

      TXTNUMCARA = XVALO(!CARAS)
   
      'TXTCARAS = SAFETEXT$(!TEXCARAS)
   
      TXTNUMTMAT = XVALO(!TIMATE)
   
      'TXTTIPOMATERIAL = SAFETEXT$(!TEXTIMATE)
   
      TXTNUMESPMAT = XVALO(!ESPMAT)
   
      'TXTESPESOR = SAFETEXT$(!TEXESPMAT)
      TXTNUMESPCOBRE = XVALO(!ESPCOBRE)
   
      'TXTESPCOBRE = SAFETEXT$(!TEXESPECOBRE)
   
      TXTNUMPTH = XVALO(!PTH)
   
      'TXTPTH = SAFETEXT$(!TEXPTH)
   
      TXTNUMPISTA = XVALO(!PISTAS)
   
      TXTPISTAS = SAFETEXT$(!TEXPISTAS)
   
      
   
      Check1.Value = XVALO(!FtainfTec)
   
      Check2.Value = XVALO(!PrvApMtra)
   
      Check3.Value = XVALO(!PrvAppgo)
   
      Check4.Value = XVALO(!NoStdar)
   
      Text9 = XVALO(!PrgPCB)
   
      Check5.Value = XVALO(!LS)
   
      Check6.Value = XVALO(!lC)
   
      Text10 = XVALO(!MASCSOLDANTE)
   
      Check8.Value = XVALO(!LCOMP_LS)
   
      Check7.Value = XVALO(!LCOMP_LC)
   
      Text16 = XVALO(!LEYECOM)
   
      Option3.Value = XVALO(!PRDEXACTA)    'PRODUCCION EXACTA
   
      Option4.Value = XVALO(!ACEXCTE) 'ACEPTA EXCEDENTE
   
      Text8 = Int(XVALO(!TOLERANCIA))  'TOLERANCIA
   

   
      TXTCIRCUITOS = XVALO(!PLZ_ENT)
      TXTADICIONALARMADOS = XVALO(!PLZ_ADIC_ARM)
   
      Check9.Value = XVALO(!ESP_GRAFITO)
   
      Check10.Value = XVALO(!ESP_GOMA)
   
      Check11.Value = XVALO(!ESP_NI_AU)
   
      Check12.Value = XVALO(!ESP_OTRO)
   
      Text15 = XVALO(!ESP_SUPMM2)
   
      Check20.Value = XVALO(!TEST_ELECT)
   
      Check19.Value = XVALO(!PNLIZADO)
   
      Check18.Value = XVALO(!MASCFONDO)
   
      Check17.Value = XVALO(!PCSO_INI_MASC)
   
      
   
      'CORTE DE PLACA
   
      Check16.Value = XVALO(!ROUTER)
   
      Check15.Value = XVALO(!SCORING)
   
      'ENTREGA
   
      Check13.Value = XVALO(!PROTOTIPO)
   
   
      'MONTAJE
   
      
   
      TXTSMDBOT1 = XVALO(!MTJ_SMDBOT1)
   
      TXTSMDTOP1 = XVALO(!MTJ_SMDTOP1)
   
      TXTINSBOT1 = XVALO(!MTJ_INSBOT1)
   
      TXTINSTOP1 = XVALO(!MTJ_INSTOP1)
   
      
   
      TXTSMDBOT2 = XVALO(!MTJ_SMDBOT2)
   
      TXTSMDTOP2 = XVALO(!MTJ_SMDTOP2)
   
      TXTINSBOT2 = XVALO(!MTJ_INSBOT2)
   
      TXTINSTOP2 = XVALO(!MTJ_INSTOP2)
   
      TXTPANELIZADO = TRIM$(SAFETEXT$(!REFPANELIZADO))
   
      TXTMASDEFONDO = TRIM$(SAFETEXT$(!REFMASDEFONDO))
   

   
      'RESUMEN
   
      Text26 = FORMATNUM$(XVALO(!GASPREPA), "F12.3")
   
      Text17 = XVALO(!CONPAG)
   
      Text31 = FORMATNUM$(XVALO(!PRC_FJO_DM2), "F12.3")
   
      Text32 = FORMATNUM$(XVALO(!DISENO), "F12.3")
   
      TxtPrcdm2 = FORMATNUM$(XVALO(!PRC_DM2), "F12.3")
   
      TxtPrcUnit = FORMATNUM$(XVALO(!Preunit), "F12.3")
   
      TXTtotPlca = FORMATNUM$(XVALO(!PrcTtPlca), "F12.3")
   
      TXTPrcTtOpcion = FORMATNUM$(XVALO(!PrcTtOpcion), "F12.3")
   
      TXTPreunitCto = FORMATNUM$(XVALO(!PreunitCto), "F12.3")
   
      TXTPreunit_Arm = FORMATNUM$(XVALO(!Preunit_Arm), "F12.3")
   
      TXTPreunit_Cpte = FORMATNUM$(XVALO(!Preunit_Cpte), "F12.3")
   
      TXTPrcTtPlcaSA = FORMATNUM$(XVALO(!PrcTtPlcaSA), "F12.3")
   
      TXTTtOpcion = FORMATNUM$(XVALO(!TtOpcion), "F12.3")
      
      TxtPrcdm2Costo = FORMATNUM$(XVALO(!Prcdm2Costo), "F12.3")
      TxtPrcUnitCosto = FORMATNUM$(XVALO(!PrcUnitCosto), "F12.3")
      TXTtotPlcaCosto = FORMATNUM$(XVALO(!totPlcaCosto), "F12.3")
      TXTPrcTtOpcionCosto = FORMATNUM$(XVALO(!PrcTtOpcionCosto), "F12.3")
   
      TXTCantStencil = TRIM$(FORMATNUM$(XVALO(!CantStencil), "F10"))
   
      TXTStTotalusd = TRIM$(FORMATNUM$(XVALO(!StTotalusd), "F10.3"))
      TXTTOTALSTENCIL = TRIM$(FORMATNUM$(XVALO(!CantStencil) * XVALO(!StTotalusd), "F10.3"))

      Text33 = FORMATNUM$(XVALO(!ANTICIPO), "F12.3")
   
      Text34 = FORMATNUM$(XVALO(!SALDO), "F12.3")
   
      
   
      TXTObservaAdmin = SAFETEXT$(!ObservaAdmin)
   
      TXTNROOCOM = SAFETEXT$(!NROOCOM)
   
      TXTObservaPlanta = SAFETEXT$(!ObservaPlanta)
   
      TXTObservCliente = SAFETEXT$(!ObservCliente)
   
      TXTObservPenal = SAFETEXT$(!ObservPenal)
   
      Text12 = XVALO(!VRS_COTIZA)
      TXTCANTICAPAS = CEROIGUALVACIO(XVALO(!CANTICAPAS))
      Text6 = FESOL$

      End With
   RS.Close
   Set RS = Nothing
   
   List1.Clear
   SQLX$ = "SELECT * FROM DOCURELCOT WITH(NOLOCK) WHERE NRO_COTIZA = " & XVALO(Text1)
   Dim DOCTMP As ADODB.Recordset
   Set DOCTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With DOCTMP
      Do While Not .EOF
         DESCRI$ = SAFETEXT$(!Descripcion)
         RUTAX$ = SAFETEXT$(!Ruta$)
         Call REPLA(TX$, DESCRI$, 1, 24)
         Call REPLA(TX$, RUTAX$, 32, 512)
         List1.AddItem TX$
        .MoveNext
      Loop
   End With
   LBLMODO = MODOCOTIZA%(XVALO(Text1))
   PRESENTANDOINFOBD% = 0

   MODIFIC% = 0

End Sub

Sub GRABAR_NUEVA_COTIZACION()

End Sub

Private Sub Text10_GotFocus()
      Call PINTATEXT(Text10)

End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub

Private Sub Text11_Change()
   MODIFIC% = 1

End Sub

Private Sub Text14_Change()
   MODIFIC% = 1
End Sub

Private Sub Text15_GotFocus()
   Call PINTATEXT(Text15)
   MODIFIC% = 1
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub Text16_GotFocus()
   Call PINTATEXT(Text16)
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub

Private Sub Text17_Change()
  Text30 = ""
  ECO$ = TRIM$(ECOARCH$("CONPAG", MKI$(XVALO(Text17))))
  Text30 = ECO$
  MODIFIC% = 1

End Sub

Private Sub Text18_Change()
   NCLIEN = XVALO(Text18)
   Text2.TEXT = rasocli$(NCLIEN)
   Text17 = Str$(CPAGCLI%(NCLIEN))
   MODIFIC% = 0

End Sub

Private Sub Text18_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub Text19_Change()
   MODIFIC% = 1

End Sub

Private Sub Text19_GotFocus()
   Call PINTATEXT(Text19)

End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub Text20_Change()
  Call ACTUINGRESOS_MANUALES
End Sub

Private Sub Text26_Change()
'   Call TXTtotPlca_Change
    Call ACTUINGRESOS_MANUALES
    MODIFIC% = 1

End Sub

Private Sub Text26_GotFocus()
   Call PINTATEXT(Text26)
   
End Sub

Private Sub Text26_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub

Private Sub Text26_LostFocus()
   Text26 = FORMATNUM$(XVALO(Text26), "F12.3")
   Text31 = FORMATNUM$(CalculaPrecioFijoDm2(XVALO(TXTPreunitCto), XVALO(TxtSupDM2)), "f16.3")

   Call ACTUINGRESOS_MANUALES

End Sub


Private Sub Text3_GotFocus()
   Call PINTATEXT(Text3)
   
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If
   
End Sub


Private Sub Text30_Change()
   IMPORTE = FORMATNUM$(XVALO(TXTTtOpcion), "F16.3") 'P
   Text33 = FORMATNUM$(CALCULAPORCENTAJE_ANTICIPO(TRIM$(Text30), IMPORTE), "F16.3") 'P
   Text34 = FORMATNUM$(SALDO(TXTTtOpcion, Text33), "F16.3") ' Q

End Sub

Private Sub Text32_Change()
   Call ACTUINGRESOS_MANUALES
   MODIFIC% = 1

End Sub

Private Sub Text32_GotFocus()
   Call PINTATEXT(Text32)
End Sub

Private Sub Text32_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub Text32_LostFocus()
   Text32 = TRIM$(FORMATNUM$(XVALO(Text32.TEXT), "F12.3"))
   Text31 = FORMATNUM$(CalculaPrecioFijoDm2(XVALO(TXTPreunitCto), XVALO(TxtSupDM2)), "f16.3")
   Call ACTUINGRESOS_MANUALES

End Sub


Private Sub Text33_GotFocus()
   Call PINTATEXT(Text33)
End Sub

Private Sub Text33_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub Text4_Change()
      MODIFIC% = 1

End Sub

Private Sub Text4_DblClick()
   Text4 = ListaSeleccionArticulos("", True)
   MODIFIC% = 1

End Sub

Private Sub Text4_GotFocus()


   Call PINTATEXT(Text4)

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub Text5_Change()
   Me.TXTSINO = ""
   If XVALO(Me.Text5) > 0 Then
     Me.TXTSINO = REFESINO$(XVALO(Me.Text5))
   End If
   MODIFIC% = 1

End Sub

Private Sub Text5_GotFocus()
   Call PINTATEXT(Text5)

End Sub


Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub Text6_Change()
   MODIFIC% = 1

End Sub

Private Sub Text6_GotFocus()
   Call PINTATEXT(Text6)

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub Text7_Change()
      MODIFIC% = 1
      Call ActuFechaEntrega(XVALO(TXTCIRCUITOS), XVALO(TXTADICIONALARMADOS))
End Sub

Private Sub Text7_DblClick()
   Text7 = ""
End Sub


Private Sub Text7_GotFocus()
   Call PINTATEXT(Text7)

End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub Text8_Change()
   If XVALO(Text8) > 0 Then
      Option4.Value = True
   Else
      Option3.Value = True
   End If
End Sub

Private Sub Text8_GotFocus()
Call PINTATEXT(Text8)
   MODIFIC% = 1

End Sub


Private Sub Text9_Change()
   Me.Text11 = ""
   If XVALO(Me.Text9) > 0 Then
     Me.Text11 = REFEPCB(XVALO(Me.Text9))
   End If
End Sub

Private Sub Text9_GotFocus()
   Call PINTATEXT(Text9)
End Sub


Private Sub Text9_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTADICIONALARMADOS_Change()
    Call ActuFechaEntrega(XVALO(TXTCIRCUITOS), XVALO(TXTADICIONALARMADOS))

End Sub

Private Sub TXTADICIONALARMADOS_GotFocus()
      Call PINTATEXT(TXTADICIONALARMADOS)
      MODIFIC% = 1


End Sub


Private Sub TXTADICIONALARMADOS_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTALERTA_DblClick()
   TXTALERTA.Visible = False
End Sub


Private Sub TxtAncho_Change()
    TxtSupDM2 = FORMATNUM$(CALCULASUPUNITARIA(XVALO(Me.TxtLargo), XVALO(Me.TxtAncho)), "F12.3")
    TxtSupLoteM2 = FORMATNUM$(CALCULASUPERFICIETOTAL(XVALO(TxtCantidad), XVALO(TxtLargo), (XVALO(TxtAncho))), "F12.3")  'actualizo la superficie del lote en m2
    MODIFIC% = 1

End Sub


Private Sub TxtAncho_GotFocus()
   Call PINTATEXT(TxtAncho)

End Sub


Private Sub TxtAncho_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TxtAncho_LostFocus()
   TxtAncho = TRIM$(FORMATNUM$(XVALO(TxtAncho), "F12.3"))
End Sub


Private Sub TxtAncho2_Change()
    TxtSupDM22 = FORMATNUM$(CALCULASUPUNITARIA(XVALO(Me.TxtLargo2), XVALO(Me.TxtAncho2)), "F12.3")
    TxtSupLoteM22 = FORMATNUM$(CALCULASUPERFICIETOTAL(XVALO(TXTCANTIDAD2), XVALO(TxtLargo2), (XVALO(TxtAncho2))), "F12.3") 'actualizo la superficie del lote en m2

End Sub

Private Sub TxtAncho2_GotFocus()
    Call PINTATEXT(TxtAncho2)
End Sub


Private Sub TxtAncho2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TxtAncho2_LostFocus()
   TxtAncho2 = TRIM$(FORMATNUM$(XVALO(TxtAncho2), "F12.3"))

End Sub

Private Sub TxtCantidad_Change()

    TxtSupLoteM2 = TRIM$(FORMATNUM$(CALCULASUPERFICIETOTAL(XVALO(TxtCantidad), XVALO(TxtLargo), (XVALO(TxtAncho))), "f12.3")) 'actualizo la superficie del lote en m2)
    'SEGUN INSTRUCCIONES SI SON MENOS DE 5 CIRCUITROS ES PROTOTIPO, PRODUCCION EXACTA Y DEMORA 6 DIAS
    If XVALO(TxtCantidad) > 0 And XVALO(TxtCantidad) <= 5 Then
'       Check13.Value = 1
'       Option3.Value = True
       TXTCIRCUITOS = 6
    Else
'       Check13.Value = 0
'       Option3.Value = False
       TXTCIRCUITOS = 12
    End If
    MODIFIC% = 1

End Sub

Private Sub TxtCantidad_GotFocus()
   Call PINTATEXT(TxtCantidad)
End Sub


Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TxtCantidad_LostFocus()
   TxtCantidad = Int(XVALO(TxtCantidad))
End Sub


Private Sub TXTCANTIDAD2_Change()
    TxtSupLoteM22 = CALCULASUPERFICIETOTAL(XVALO(TXTCANTIDAD2), XVALO(TxtLargo2), (XVALO(TxtAncho2)))  'actualizo la superficie del lote en m2

    If XVALO(TXTCANTIDAD2) <= 5 Then
       Check13.Value = 1
       Option3.Value = True
       TXTCIRCUITOS = 6
    End If

End Sub

Private Sub TXTCANTIDAD2_GotFocus()
    Call PINTATEXT(TXTCANTIDAD2)
End Sub


Private Sub TXTCANTIDAD2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTCANTIDAD2_LostFocus()
   TXTCANTIDAD2 = Int(XVALO(TXTCANTIDAD2))
End Sub

Private Sub TXTCantStencil_Change()
    
    TXTTOTALSTENCIL = TRIM$(FORMATNUM$(XVALO(TXTCantStencil) * XVALO(TXTStTotalusd), "F16.3"))
    MODIFIC% = 1

End Sub

Private Sub TXTCantStencil_GotFocus()
   Call PINTATEXT(TXTCantStencil)
   
End Sub


Private Sub TXTCantStencil_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        TXTStTotalusd.SetFocus
'        KeyAscii = 0        ' Para que no "pite"
'        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If
End Sub

Private Sub TXTCARAS_Change()
   If TXTCARAS Like "MULTICAPA" And XVALO(TXTNUMTMAT) > 0 Then  'VALIDA SI TIENE NRO DE T.MATERIAL PARA QUE NO LO PRESENTE
     If PRESENTANDOINFOBD% < 1 Then
       CANTICAPA$ = TRIM$(InputBox$("Ingreso de Capas", "Ingrese Cantidad de Capas", Mid$(CACOPX$, 2))) 'CUANDO ES LA PRIMERA VEZ
       If CANTICAPA$ = "" Then GoTo 30
       CANTCAPA% = XVALO(CANTICAPA$)
       TXTCANTICAPAS = CANTCAPA%
     End If
   Else
       TXTCANTICAPAS = ""
   End If
30
   Call CargaValoresPorDefecto(TXTCARAS)
   MODIFIC% = 1

End Sub

Private Sub TXTCIRCUITOS_Change()
    Call ActuFechaEntrega(XVALO(TXTCIRCUITOS), XVALO(TXTADICIONALARMADOS))
End Sub
Sub ActuFechaEntrega(PlazoEntregaCircuitos%, PlazoEntregaArmado%)
    Text6 = ""
    'LA FECHA DE ENTREGA SOLO LA ACTUALIZA SI TIENE FECHA DE COMPRA Y SE CONSIDERA ESTA COMO PUNTO DE PARTIDA
    If FECHANUM(Text7) > 0 Then
        If TXTFECHA.TEXT <> "" Then
           CANTDIAX% = PlazoEntregaCircuitos% + PlazoEntregaArmado%
           CANTDIAX = CANT_DIASPOSTHABILES(FECHANUM(Text7), CANTDIAX%)
           Text6 = FECHATEX$(DIASPOST(FECHANUM(Text7), CANTDIAX%))
        End If
    End If
End Sub

Private Sub TXTCIRCUITOS_GotFocus()
      Call PINTATEXT(TXTCIRCUITOS)
      MODIFIC% = 1


End Sub


Private Sub TXTCIRCUITOS_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTESPCOBRE_Change()
   MODIFIC% = 1

End Sub

Private Sub TXTESPESOR_Change()
'   TxtPrcdm2Costo = CALCULACOSTO(TxtSupLoteM2, TXTCARAS, TXTTIPOMATERIAL, TXTESPESOR, TXTESPCOBRE, TXTPTH, RET_COSTONORMAL)
'   TxtPrcdm2 = RET_COSTONORMAL 'VALOR RETORNADO DE LA FUNCION CALCULACOSTO
   MODIFIC% = 1

End Sub

Private Sub TXTFECHA_Change()
   MODIFIC% = 1
   'Call ActuFechaEntrega(XVALO(TXTCIRCUITOS), XVALO(TXTADICIONALARMADOS))
 
End Sub

Private Sub TXTFECHA_GotFocus()
   Call PINTATEXT(TXTFECHA)

End Sub

Private Sub TXTFECHA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTFECHA_LostFocus()
   If FECHANUM(TXTFECHA) < 1 Then
      Call COMUNI("Falta Fecha o Formato No Válido\(dd/mm/aa) ")
      Call PINTATEXT(TXTFECHA)
   End If
End Sub

Private Sub TXTINSBOT1_Change()
   TXTINSTOP3 = TRIM$(FORMATNUM$(ACTUALIZATOTOALES_MONTAJE_INS(TXTINSBOT1, TXTINSBOT2, TXTINSTOP1, TXTINSTOP2), "F12.3"))
End Sub

Private Sub TXTINSBOT1_GotFocus()
  Call PINTATEXT(TXTINSBOT1)

End Sub

Private Sub TXTINSBOT1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTINSBOT2_Change()
   TXTINSTOP3 = TRIM$(FORMATNUM$(ACTUALIZATOTOALES_MONTAJE_INS(TXTINSBOT1, TXTINSBOT2, TXTINSTOP1, TXTINSTOP2), "F12.3"))
End Sub

Private Sub TXTINSBOT2_GotFocus()
  Call PINTATEXT(TXTINSBOT2)

End Sub

Private Sub TXTINSBOT2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTINSTOP1_Change()
   TXTINSTOP3 = TRIM$(FORMATNUM$(ACTUALIZATOTOALES_MONTAJE_INS(TXTINSBOT1, TXTINSBOT2, TXTINSTOP1, TXTINSTOP2), "F12.3"))

End Sub

Private Sub TXTINSTOP1_GotFocus()
  Call PINTATEXT(TXTINSTOP1)
End Sub

Private Sub TXTINSTOP1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTINSTOP2_Change()
   TXTINSTOP3 = TRIM$(FORMATNUM$(ACTUALIZATOTOALES_MONTAJE_INS(TXTINSBOT1, TXTINSBOT2, TXTINSTOP1, TXTINSTOP2), "F12.3"))

End Sub

Private Sub TXTINSTOP2_GotFocus()
  Call PINTATEXT(TXTINSTOP2)

End Sub

Private Sub TXTINSTOP2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTINSTOP3_Change()
   TXTTOTALCOMPONENTES = TRIM$(FORMATNUM$(XVALO(TXTSMDTOP3) + XVALO(TXTINSTOP3), "F12.3"))
   MODIFIC% = 1
End Sub

Private Sub TxtLargo_Change()
   TxtSupDM2 = FORMATNUM$(CALCULASUPUNITARIA(XVALO(Me.TxtLargo), XVALO(Me.TxtAncho)), "F12.3")
   TxtSupLoteM2 = FORMATNUM$(CALCULASUPERFICIETOTAL(XVALO(TxtCantidad), XVALO(TxtLargo), (XVALO(TxtAncho))), "F12.3")  'actualizo la superficie del lote en m2
   MODIFIC% = 1

End Sub

Private Sub TxtLargo_GotFocus()
   Call PINTATEXT(TxtLargo)
End Sub


Private Sub TxtLargo_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TxtLargo_LostFocus()
   TxtLargo = TRIM$(FORMATNUM$(XVALO(TxtLargo), "F12.3"))
End Sub


Private Sub TxtLargo2_Change()
   TxtSupDM22 = FORMATNUM$(CALCULASUPUNITARIA(XVALO(Me.TxtLargo2), XVALO(Me.TxtAncho2)), "F12.3")
   TxtSupLoteM22 = FORMATNUM$(CALCULASUPERFICIETOTAL(XVALO(TXTCANTIDAD2), XVALO(TxtLargo2), (XVALO(TxtAncho2))), "F12.3")  'actualizo la superficie del lote en m2


End Sub

Private Sub TxtLargo2_GotFocus()
   Call PINTATEXT(TxtLargo2)
End Sub


Private Sub TxtLargo2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TxtLargo2_LostFocus()
   TxtLargo2 = TRIM$(FORMATNUM$(XVALO(TxtLargo2), "F12.3"))
End Sub

Private Sub TXTMASDEFONDO_Change()
MODIFIC% = 1
End Sub

Private Sub TXTMASDEFONDO_GotFocus()
  Call PINTATEXT(TXTMASDEFONDO)

End Sub

Private Sub TXTMASDEFONDO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTNROOCOM_GotFocus()
   Call PINTATEXT(TXTNROOCOM)
   MODIFIC% = 1
End Sub

Private Sub TXTNROOCOM_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTNUMCARA_Change()
   Me.TXTCARAS = ""
   If XVALO(Me.TXTNUMCARA) > 0 Then
     Me.TXTCARAS = REFECARAS(XVALO(Me.TXTNUMCARA))
   End If
'   If Check21.Value = 1 Then
'      Call TxtSupLoteM2_Change
'   Else
'      Call TxtSupLoteM22_Change
'   End If
   VALODEFECTO% = 1
End Sub

Private Sub CMD3_Click()
 SELESPEM = SELESPEMATE
   If SELESPEM > 0 Then
      Me.TXTNUMESPMAT = SELESPEM
   End If
End Sub

Private Sub TXTNUMCARA_GotFocus()

   Call PINTATEXT(TXTNUMCARA)

End Sub

Private Sub TXTNUMCARA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub

Private Sub TXTNUMESPCOBRE_GotFocus()
   Call PINTATEXT(TXTNUMESPCOBRE)

End Sub

Private Sub TXTNUMESPCOBRE_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub

Private Sub TXTNUMESPMAT_Change()
   Me.TXTESPESOR = ""
   If XVALO(Me.TXTNUMESPMAT) > 0 Then
     Me.TXTESPESOR = REFEESPEMATE$(XVALO(Me.TXTNUMESPMAT))
   End If
'   If Check21.Value = 1 Then
'      Call TxtSupLoteM2_Change
'   Else
'      Call TxtSupLoteM22_Change
'   End If

End Sub
Private Sub CMD4_Click()
   SELCOBRE = SELEPECOBRE
   If SELCOBRE > 0 Then
      Me.TXTNUMESPCOBRE = SELCOBRE
   End If
End Sub
Private Sub TXTNUMESPCOBRE_Change()
   Me.TXTESPCOBRE = ""
   If XVALO(Me.TXTNUMESPCOBRE) > 0 Then
     Me.TXTESPCOBRE = REFEESPECOBRE(XVALO(Me.TXTNUMESPCOBRE))
   End If
'   If Check21.Value = 1 Then
'      Call TxtSupLoteM2_Change
'   Else
'      Call TxtSupLoteM22_Change
'   End If

End Sub
Private Sub CMD5_Click()
   PTH = SELEPTH
   If PTH > 0 Then
      Me.TXTNUMPTH = PTH
   End If
End Sub



Private Sub TXTNUMESPMAT_GotFocus()
   Call PINTATEXT(TXTNUMESPMAT)

End Sub

Private Sub TXTNUMESPMAT_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub

Private Sub TXTNUMPISTA_GotFocus()
      Call PINTATEXT(TXTNUMPISTA)

End Sub

Private Sub TXTNUMPISTA_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub

Private Sub TXTNUMPTH_Change()
   Me.TXTPTH = ""
   If XVALO(Me.TXTNUMPTH) > 0 Then
     Me.TXTPTH = REFEPTH(XVALO(Me.TXTNUMPTH))
   End If
'   If Check21.Value = 1 Then
'      Call TxtSupLoteM2_Change
'   Else
'      Call TxtSupLoteM22_Change
'   End If

End Sub
Private Sub CMD6_Click()
   PISTA = SELEPISTA
   If PISTA > 0 Then
      Me.TXTNUMPISTA = PISTA
   End If
End Sub
Private Sub TXTNUMPISTA_Change()
   Me.TXTPISTAS = ""
   If XVALO(Me.TXTNUMPISTA) > 0 Then
     Me.TXTPISTAS = REFEPISTA(XVALO(Me.TXTNUMPISTA))
   End If
'   If Check21.Value = 1 Then
'      Call TxtSupLoteM2_Change
'   Else
'      Call TxtSupLoteM22_Change
'   End If

End Sub

Private Sub TXTNUMPTH_GotFocus()
      Call PINTATEXT(TXTNUMPTH)
End Sub

Private Sub TXTNUMPTH_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub

Private Sub TXTNUMTMAT_Change()
   VALONU = XVALO(Me.TXTNUMTMAT.TEXT)
   If VALONU < 0 Or VALONU > 255 Then
      Me.TXTTIPOMATERIAL.TEXT = "": Exit Sub
   End If
   TXTTIPOMATERIAL.TEXT = ECOARCH$("TATIMAT", Chr$(VALONU))
'   If Check21.Value = 1 Then
'      Call TxtSupLoteM2_Change
'   Else
'      Call TxtSupLoteM22_Change
'   End If

End Sub
Private Sub CMD8_Click()
   MASCSOLDANTE = SELEMASCANTSOLDANTE
   If MASCSOLDANTE > 0 Then
      Me.Text10 = MASCSOLDANTE
   End If
End Sub

Private Sub TEXT10_Change()
   VALONU = XVALO(Me.Text10.TEXT)
   Text13 = ""
   If VALONU > 0 Then
     Me.Text13 = REFEMASCANTSOLDANTE(XVALO(Me.Text10))
   End If
   MODIFIC% = 1
   
End Sub
Private Sub Command6_Click()
   LEYECOM = SELELEYECOMP
   If LEYECOM > 0 Then
      Me.Text16 = LEYECOM
   End If
End Sub
Private Sub TEXT16_Change()
   VALONU = XVALO(Me.Text16.TEXT)
   Text14 = ""
   If VALONU > 0 Then
     Me.Text14 = REFELEYECOMP(XVALO(Me.Text16))
   End If
   MODIFIC% = 1
End Sub

Private Sub TXTNUMTMAT_GotFocus()
   Call PINTATEXT(TXTNUMTMAT)

End Sub

Private Sub TXTNUMTMAT_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub

Private Sub TXTObservaAdmin_Change()
    MODIFIC% = 1
End Sub

Private Sub TXTObservaPlanta_Change()
    MODIFIC% = 1
End Sub

Private Sub TXTObservCliente_Change()
    MODIFIC% = 1
End Sub

Private Sub TXTObservPenal_Change()
     MODIFIC% = 1
End Sub


Private Sub TXTPANELIZADO_GotFocus()
  Call PINTATEXT(TXTPANELIZADO)
  MODIFIC% = 1
End Sub

Private Sub TXTPANELIZADO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub

Private Sub TXTPISTAS_Change()
   MODIFIC% = 1

End Sub

Private Sub TxtPrcdm2_Change()
   TxtPrcUnit = ""
End Sub

Private Sub TxtPrcdm2Costo_Change()
  TxtPrcUnitCosto = ""
End Sub


Private Sub TxtPrcUnit_Change()
   
   If Check21.Value = 1 Then
'      TXTtotPlca = FORMATNUM$(XVALO(TxtPrcUnit) * XVALO(TxtCantidad), "F16.3")
   ElseIf Check14.Value = 1 Then
'      CANTIX = (XVALO(TxtSupLoteM22) * XVALO(TXTCANTIDAD2)) / TxtCantidad
'      TXTtotPlca = FORMATNUM$(XVALO(TxtPrcUnit) * XVALO(TXTCANTIDAD2), "F16.3")
   End If
   
   TXTPreunitCto = TxtPrcUnit
End Sub

Private Sub TxtPrcUnit_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
       Call ALERTA(Me.TxtSupLoteM2, "PRECIO UNITARIO = Precio dm2 * sup.(dm2)")
    End If

End Sub

Private Sub TxtPrcUnitCosto_Change()
   If Check21.Value = 1 Then
      TXTtotPlcaCosto = FORMATNUM$(XVALO(TxtPrcUnitCosto) * XVALO(TxtCantidad), "F16.3")
   Else
      TXTtotPlcaCosto = FORMATNUM$(XVALO(TxtPrcUnitCosto) * XVALO(TXTCANTIDAD2), "F16.3")
   End If
End Sub

Private Sub TXTPreunit_Arm_Change()
   Call ACTUINGRESOS_MANUALES
   If XVALO(TXTPreunit_Arm) > 0 Then
      TXTADICIONALARMADOS = "10"
   Else
      TXTADICIONALARMADOS = ""
   End If
   Call ACTUINGRESOS_MANUALES
   MODIFIC% = 1


End Sub

Private Sub TXTPreunit_Arm_GotFocus()
   Call PINTATEXT(TXTPreunit_Arm)

End Sub

Private Sub TXTPreunit_Arm_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub

Private Sub TXTPreunit_Arm_LostFocus()
   TXTPreunit_Arm = FORMATNUM$(XVALO(TXTPreunit_Arm), "F12.3")
End Sub


Private Sub TXTPreunit_Cpte_Change()
   Call ACTUINGRESOS_MANUALES
   MODIFIC% = 1

End Sub

Private Sub TXTPreunit_Cpte_GotFocus()
   Call PINTATEXT(TXTPreunit_Cpte)
End Sub

Private Sub TXTPreunit_Cpte_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub

Private Sub TXTPreunit_Cpte_LostFocus()
   TXTPreunit_Cpte = FORMATNUM$(XVALO(TXTPreunit_Cpte), "F12.3")
End Sub


Private Sub TXTPreunitCto_Change()
   Text31 = FORMATNUM$(CalculaPrecioFijoDm2(XVALO(TXTPreunitCto), XVALO(TxtSupDM2)), "f16.3")
   Call ACTUINGRESOS_MANUALES
   MODIFIC% = 1

End Sub

Private Sub TXTPreunitCto_GotFocus()
   Call PINTATEXT(TXTPreunitCto)

End Sub


Private Sub TXTPreunitCto_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTPreunitCto_LostFocus()
   TXTPreunitCto = FORMATNUM$(XVALO(TXTPreunitCto), "F12.3")
End Sub


Private Sub TXTPTH_Change()
   MODIFIC% = 1

End Sub

Private Sub TXTSMDBOT1_Change()
   TXTSMDTOP3 = TRIM$(FORMATNUM$(ACTUALIZATOTOALES_MONTAJE_SMD(TXTSMDBOT1, TXTSMDBOT2, TXTSMDTOP1, TXTSMDTOP2), "F12.3"))
   MODIFIC% = 1
End Sub

Private Sub TXTSMDBOT1_GotFocus()
  Call PINTATEXT(TXTSMDBOT1)

End Sub

Private Sub TXTSMDBOT1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTSMDBOT2_Change()
   TXTSMDTOP3.TEXT = TRIM$(FORMATNUM$(ACTUALIZATOTOALES_MONTAJE_SMD(TXTSMDBOT1, TXTSMDBOT2, TXTSMDTOP1, TXTSMDTOP2), "F12.3"))

End Sub

Private Sub TXTSMDBOT2_GotFocus()
  Call PINTATEXT(TXTSMDBOT2)

End Sub

Private Sub TXTSMDBOT2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTSMDTOP1_Change()
   TXTSMDTOP3 = TRIM$(FORMATNUM$(ACTUALIZATOTOALES_MONTAJE_SMD(TXTSMDBOT1, TXTSMDBOT2, TXTSMDTOP1, TXTSMDTOP2), "F12.3"))
   MODIFIC% = 1

End Sub

Private Sub TXTSMDTOP1_GotFocus()
  Call PINTATEXT(TXTSMDTOP1)

End Sub

Private Sub TXTSMDTOP1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTSMDTOP2_Change()
   TXTSMDTOP3 = TRIM$(FORMATNUM$(ACTUALIZATOTOALES_MONTAJE_SMD(TXTSMDBOT1, TXTSMDBOT2, TXTSMDTOP1, TXTSMDTOP2), "F12.3"))

End Sub

Private Sub TXTSMDTOP2_GotFocus()
  Call PINTATEXT(TXTSMDTOP2)

End Sub

Private Sub TXTSMDTOP2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub TXTSMDTOP3_Change()
   TXTTOTALCOMPONENTES = TRIM$(FORMATNUM$(TXTSMDTOP3 + TXTINSTOP3, "F12.3"))
   MODIFIC% = 1
End Sub

Private Sub TXTStTotalusd_Change()
    TXTTOTALSTENCIL = TRIM$(FORMATNUM$(XVALO(TXTCantStencil) * XVALO(TXTStTotalusd), "F16.3"))
    MODIFIC% = 1

End Sub

Private Sub TXTStTotalusd_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If
End Sub


Private Sub TXTStTotalusd_LostFocus()
  TXTStTotalusd = TRIM$(FORMATNUM$(XVALO(TXTStTotalusd), "f8"))
  
End Sub


Private Sub TxtSupDM2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then
       TX1$ = " Es la superficie unitaria en dm2 " + vbCrLf
       TX1$ = TX1$ + "FORMULA:" + vbCrLf
       TX1$ = TX1$ + "PASO A DM" + vbCrLf
       TX1$ = TX1$ + "LARGO = LARGO / 100" + vbCrLf
       TX1$ = TX1$ + "  ANCHO = ANCHO / 100" + vbCrLf
       TX1$ = TX1$ + "  RESULTADO = LARGO * ANCHO" + vbCrLf

       Call ALERTA(Me.TxtSupDM2, TX1$)
    End If

End Sub

Private Sub TxtSupDM22_Change()
   TxtSupDM22 = TRIM$(FORMATNUM$(XVALO(TxtSupDM22), "F12.3"))
   If XVALO(TxtCantidad) > 0 Then
      TxtSupDM22INV = (XVALO(TxtSupDM22) * XVALO(TXTCANTIDAD2)) / XVALO(TxtCantidad)
      'EN ESTE CASILLERO GUARDO EL VALOR PONDERADO PARA CALCUALAR LUEGO EL COSTO EN BASE A ESTE.
   End If
End Sub

Private Sub TxtSupLoteM2_LostFocus()
   TxtSupLoteM2 = TRIM$(FORMATNUM$(XVALO(TxtSupLoteM2), "F12.3"))

End Sub


Private Sub TxtSupDM2_LostFocus()
   TxtSupDM2 = TRIM$(FORMATNUM$(XVALO(TxtSupDM2), "F12.3"))

End Sub


Sub ALERTA(Objeto As Object, TEXMEN$)
    OLEFT = Objeto.Left
    OANCHO = Objeto.Width
    DE_IZQ = OLEFT + OANCHO
    MAX_IZQ = DE_IZQ + TXTALERTA.Width
    ANCHOFORM = Me.Width
    If TEXMEN$ <> "" Then TXTALERTA = TEXMEN$
    If MAX_IZQ > ANCHOFORM Then
       TXTALERTA.Visible = True
       TXTALERTA.Left = OLEFT - (Objeto.Width + 300)
       TXTALERTA.Top = Objeto.Top
    Else
       TXTALERTA.Visible = True
       TXTALERTA.Left = DE_IZQ
       TXTALERTA.Top = Objeto.Top
    End If
'    For i% = 1 To 3
'    Call T_Espera(1)
'        RECUERDO$ = TXTALERTA.TEXT
'        TXTALERTA.TEXT = ""
'        Call T_Espera(1 / 10)
'        TXTALERTA.TEXT = RECUERDO$
'        Call T_Espera(1)
'    Next i%
End Sub


Function CalculaPrecioFijoDm2(PrUntCircuito, SupUniDm2)
   CalculaPrecioFijoDm2 = 0
   PrUntCircuito = XVALO(PrUntCircuito)
   SupUniDm2 = XVALO(SupUniDm2)
   If SupUniDm2 > 0 Then
     If PrUntCircuito > 0 Then
       CalculaPrecioFijoDm2 = PrUntCircuito / SupUniDm2
     End If
   End If
End Function

Private Sub TxtSupLoteM2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Button = 2 Then
       TX1$ = " Es la superficie TOTAL en  m2 " + vbCrLf
       TX1$ = TX1$ + "FORMULA:" + vbCrLf
       TX1$ = TX1$ + "PASO A DM" + vbCrLf
       TX1$ = TX1$ + "LARGO = LARGO / 100" + vbCrLf
       TX1$ = TX1$ + "  ANCHO = ANCHO / 100" + vbCrLf
       TX1$ = TX1$ + "CALCULO LA SUPERFICIE TOTAL EN M2" + vbCrLf
       TX1$ = TX1$ + "  RESULTADO = CANTI * ((LARGO / 10) * (ANCHO / 10)) "

       Call ALERTA(Me.TxtSupDM2, TX1$)
    End If
End Sub


Private Sub TxtSupLoteM22_Change()
'   If Check14.Value = 1 Then
'        If XVALO(TXTCANTIDAD2) > 0 Then
'           If XVALO(TxtCantidad) > 0 Then
''             var_TxtSupLoteM22 = (XVALO(Me.TxtSupLoteM22) * XVALO(TXTCANTIDAD2)) / XVALO(TxtCantidad)
''             Text20 = var_TxtSupLoteM22
'             TxtPrcdm2Costo = FORMATNUM$(CALCULACOSTO(TxtSupLoteM22, TXTCARAS, TXTTIPOMATERIAL, TXTESPESOR, TXTESPCOBRE, TXTPTH, RET_COSTONORMAL), "F16.3")
'             'TxtPrcdm2Costo = FORMATNUM$(CALCULACOSTO(var_TxtSupLoteM22, TXTCARAS, TXTTIPOMATERIAL, TXTESPESOR, TXTESPCOBRE, TXTPTH, RET_COSTONORMAL), "F16.3")
'             TxtPrcdm2 = FORMATNUM$(RET_COSTONORMAL, "F16.3") 'VALOR RETORNADO DE LA FUNCION CALCULACOSTO
'             TxtPrcUnit = ""
'             If XVALO(TxtPrcdm2) > 0 Then
'                 If XVALO(TxtSupDM22) > 0 Then
'                    'TxtPrcUnit = TRIM$(FORMATNUM$(XVALO(TxtPrcdm2) * XVALO(TxtSupDM22), "F16.3"))
'                    TxtPrcUnit = TRIM$(FORMATNUM$(XVALO(TxtPrcdm2) * XVALO(TxtSupDM22INV), "F16.3"))
'                 End If
'             End If
'
'             TxtPrcUnitCosto = ""
'             If XVALO(TxtPrcdm2Costo) > 0 Then
'                 If XVALO(TxtSupDM22) > 0 Then
'                    'TxtPrcUnitCosto = TRIM$(FORMATNUM$(XVALO(TxtPrcdm2Costo) * XVALO(TxtSupDM22), "F16.3"))
'                    'CAMBIO SE PASA EL VALOR DE TXTSUPERFICIE PONDERADO
'                    TxtPrcUnitCosto = TRIM$(FORMATNUM$(XVALO(TxtPrcdm2Costo) * XVALO(TxtSupDM22INV), "F16.3"))
'
'                 End If
'             End If
'           End If
'        Else
'
'        End If
'   Else
'      Call TxtCantidad_Change
'
'   End If
   


End Sub

Private Sub TXTTIPOMATERIAL_Change()
'   TxtPrcdm2Costo = CALCULACOSTO(TxtSupLoteM2, TXTCARAS, TXTTIPOMATERIAL, TXTESPESOR, TXTESPCOBRE, TXTPTH, RET_COSTONORMAL)
'   TxtPrcdm2 = RET_COSTONORMAL 'VALOR RETORNADO DE LA FUNCION CALCULACOSTO
   MODIFIC% = 1

End Sub

Private Sub TXTTOTALCOMPONENTES_Change()
   MODIFIC% = 1
End Sub

Private Sub TXTTOTALSTENCIL_Change()
   Call ACTUINGRESOS_MANUALES
End Sub

Private Sub TXTtotPlca_Change()
'   TXTPrcTtOpcion = FORMATNUM$(TXTtotPlca, "F16.3")
'   TXTPrcTtPlcaSA = TXTtotPlca
'   TXTTtOpcion = XVALO(TXTtotPlca) + XVALO(Text26) + XVALO(Text32) + XVALO(TXTStTotalusd)
End Sub
Sub ACTUINGRESOS_MANUALES()
'(PrecTotOperac, PrecUnitCircto, PrecUnitArmado, PrecUnitComponente)
'       TXTPrcTtOpcion = TXTtotPlca
'       TXTPrcTtOpcion = FORMATNUM$(TXTtotPlca, "F16.3")
'       TXTPrcTtPlcaSA = TXTtotPlca
'       TXTTtOpcionCosto = FORMATNUM$(XVALO(TxtCantidad) * (XVALO(PrecUnitCircto) + XVALO(PrecUnitArmado) + XVALO(PrecUnitComponente)) + (XVALO(TXTStTotalusd) * XVALO(TXTCantStencil)), "F12.3")
'
'TXTPreunit_ArmCosto = XVALO(TXTPreunitCto) + XVALO(TXTPreunit_Arm)
'TXTPreunit_CpteCosto = XVALO(TXTPreunitCto) + XVALO(TXTPreunit_Arm) + XVALO(TXTPreunit_Cpte)
''OJO VER EL TEMA DE CANTIDAD CUANDO ES POR PANELES
'TXTPrcTtPlcaSACosto = (XVALO(TXTPreunitCto) + XVALO(TXTPreunit_Arm) + XVALO(TXTPreunit_Cpte)) * XVALO(TxtCantidad)
'TXTTtOpcion = XVALO(TXTtotPlca) + XVALO(Text26) + XVALO(Text32) + XVALO(TXTStTotalusd)
'
'TXTTtOpcionCosto = FORMATNUM$(XVALO(TXTPrcTtPlcaSACosto), "F12.3")


''0''''''''''''''''''''''''''''''''''''''''''''''
'
'TXTPrcTtOpcion = FORMATNUM$(XVALO(TxtPrcUnit) * XVALO(TxtCantidad), "F16.3") 'D
'TXTPreunitCto = FORMATNUM$(XVALO(TXTPreunitCto) * XVALO(TxtCantidad), "F16.3") 'E
'TXTPrcTtPlcaSA = FORMATNUM$(XVALO(TXTPreunitCto) * XVALO(TxtCantidad), "F16.3") 'H
If XVALO(TXTPreunitCto) > 0 Then
   PREUNIX = XVALO(TXTPreunitCto) 'PRECIO UNITARIO DE VALOR INGRESADO MANUALMENTE
   TXTPrcTtPlcaSA = FORMATNUM$(XVALO(PREUNIX) * XVALO(TxtCantidad) + XVALO(Text26) + XVALO(Text32), "F16.3") 'D
'21/01/20  este seria el cambio a realizar comentando desde elsse hasta el endif
'Else solo va a ser el valor que se ingrese a mano el que va a tener en cuenta como solicitaron ellos
'   PREUNIX = XVALO(TxtPrcUnit)     'PRECIO UNITARIO DE TABLA LOGARITIMICA
'   TXTPrcTtPlcaSA = FORMATNUM$(XVALO(PREUNIX) + XVALO(Text26) + XVALO(Text32), "F16.3") 'E
End If
 
    'TXTPrcTtPlcaSA = FORMATNUM$(XVALO(TxtPrcUnit) * XVALO(TxtCantidad), "F16.3") 'H
    TXTPrcTtOpcion = FORMATNUM$(XVALO(TxtPrcUnit) * XVALO(TxtCantidad), "F16.3") 'D
    
    TXTPrcTtOpcionCosto = FORMATNUM$(TXTtotPlcaCosto, "F16.3") 'LL
    TXTPreunit_ArmCosto = FORMATNUM$(XVALO(TXTPreunitCto) + XVALO(TXTPreunit_Arm), "F16.3") ' M
    TXTPreunit_CpteCosto = FORMATNUM$(XVALO(TXTPreunitCto) + XVALO(TXTPreunit_Arm) + XVALO(TXTPreunit_Cpte), "F16.3") 'N
    TXTPrcTtPlcaSACosto = FORMATNUM$(XVALO(TXTPreunit_CpteCosto) * XVALO(TxtCantidad), "F16.3") 'Ñ
    TXTTtOpcionCosto = FORMATNUM$(XVALO(TXTPrcTtPlcaSACosto), "F16.3") 'O  HAY QUE CONFIRMAR SI SUMA LOS STENCIL  (R * S) (XVALO(TXTCantStencil) * XVALO(TXTStTotalusd))
    A = (XVALO(PREUNIX) * XVALO(TxtCantidad))
    AA = ((XVALO(PREUNIX) * XVALO(TxtCantidad)) + XVALO(Text26) + XVALO(Text32) + (XVALO(TXTPreunit_Arm) * XVALO(TxtCantidad)) + (XVALO(TXTPreunit_Cpte) * XVALO(TxtCantidad)) + XVALO(TXTTOTALSTENCIL))

    TXTTtOpcion = FORMATNUM$((XVALO(PREUNIX) * XVALO(TxtCantidad)) + XVALO(Text26) + XVALO(Text32) + (XVALO(TXTPreunit_Arm) * XVALO(TxtCantidad)) + (XVALO(TXTPreunit_Cpte) * XVALO(TxtCantidad)) + XVALO(TXTTOTALSTENCIL), "F16.3")     'I

  
   Text33 = "" 'P
   IMPORTE = FORMATNUM$(XVALO(TXTTtOpcion), "F16.3") 'P
   Text33 = FORMATNUM$(CALCULAPORCENTAJE_ANTICIPO(TRIM$(Text30), IMPORTE), "F16.3") 'P
   Text34 = FORMATNUM$(SALDO(TXTTtOpcion, Text33), "F16.3") ' Q

   

End Sub


Sub CargaValoresPorDefecto(CARAS$)
      TXTNUMTMAT = 11
      TXTNUMESPMAT = 6
      TXTNUMESPCOBRE = 1
      TXTNUMPISTA = 1

      If InStr(CARAS$, "SIMPLE") > 0 Then
           TXTNUMPTH = 2
           TXTNUMPISTA = 1
           Check5.Value = 1
           Check7.Value = 1
           Text10 = 1
           Text16 = 2
      ElseIf InStr(CARAS$, "DOBLE") > 0 Then
           TXTNUMESPCOBRE = 2
           TXTNUMPTH = 1
           Check5.Value = 1
           Check6.Value = 1
           Check7.Value = 1
           Text10 = 1
           Text16 = 2
      End If
   
   
End Sub

Private Sub TXTtotPlcaCosto_Change()
   TXTPrcTtOpcionCosto = TXTtotPlcaCosto
End Sub

Private Sub TXTTtOpcion_Change()
   Text34 = FORMATNUM$(SALDO(TXTTtOpcion, Text33), "F16.3")
   Text33 = ""
   Text33 = CALCULAPORCENTAJE_ANTICIPO(Text30, XVALO(TXTTtOpcion))

End Sub
Function SALDO(TOTOPERACION, ANTICIPO)
   SALDO = XVALO(TOTOPERACION) - XVALO(ANTICIPO)
End Function


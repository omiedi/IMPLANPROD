VERSION 5.00
Begin VB.Form FRMDISENIO 
   Caption         =   "Form1"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17175
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   17175
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
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
      Height          =   315
      Left            =   5280
      TabIndex        =   62
      Text            =   "Prueba nueva fuente ancho fijo"
      Top             =   3720
      Width           =   4605
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E2D3C0&
      Caption         =   "GESTION DE VENTAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3645
      Left            =   12480
      TabIndex        =   48
      Top             =   480
      Width           =   4065
      Begin VB.Frame Frame9 
         BackColor       =   &H00E2D3C0&
         Caption         =   "OPERACIONES COMERCIALES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   240
         TabIndex        =   53
         Top             =   360
         Width           =   3585
         Begin VB.PictureBox Picture14 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   420
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   57
            Top             =   420
            Width           =   2325
            Begin VB.CommandButton Command22 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "FRMDISENIO1.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   58
               ToolTipText     =   "Generar una Factura en Cuenta Corriente"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Facturación C.Cte"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   59
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture12 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   840
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   54
            Top             =   840
            Width           =   2325
            Begin VB.CommandButton Command14 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "FRMDISENIO1.frx":0442
               Style           =   1  'Graphical
               TabIndex        =   55
               ToolTipText     =   "Emitir o Registrar un Recibo de Cobranza"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Recibo Cobranza en Cta. Corriente"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   56
               Top             =   105
               Width           =   1485
            End
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2D3C0&
         Caption         =   "CONTROL DE GESTION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   240
         TabIndex        =   49
         Top             =   2160
         Width           =   3585
         Begin VB.PictureBox Picture11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   630
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   50
            Top             =   420
            Width           =   2325
            Begin VB.CommandButton Command3 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "FRMDISENIO1.frx":058C
               Style           =   1  'Graphical
               TabIndex        =   51
               ToolTipText     =   "Acceso al Control General de Gestión Comercial"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Control General de Gestión Comercial"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   52
               Top             =   105
               Width           =   1620
            End
         End
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      Appearance      =   0  'Flat
      BackColor       =   &H002B9973&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   13680
      ScaleHeight     =   165
      ScaleWidth      =   615
      TabIndex        =   46
      Top             =   5880
      Width           =   650
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H0048CC9D&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00808080&
      Height          =   495
      Left            =   12360
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   43
      Top             =   7080
      Width           =   975
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00CDDADA&
      Height          =   495
      Left            =   11400
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   42
      Top             =   7080
      Width           =   975
   End
   Begin VB.PictureBox Picture6 
      BackColor       =   &H002C3D46&
      Height          =   495
      Left            =   12360
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   41
      Top             =   6600
      Width           =   975
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H0095C5E1&
      Height          =   495
      Left            =   11400
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   40
      Top             =   6600
      Width           =   975
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H0050765A&
      ForeColor       =   &H00404080&
      Height          =   495
      Left            =   12360
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   39
      Top             =   6120
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C7D9CC&
      Height          =   495
      Left            =   11400
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   38
      Top             =   6120
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00624E28&
      Height          =   495
      Left            =   12360
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   37
      Top             =   5640
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E2D3C0&
      Height          =   495
      Left            =   11400
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   36
      Top             =   5640
      Width           =   975
   End
   Begin VB.TextBox Text28 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   5280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   35
      Top             =   960
      Width           =   7050
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Seleccionar Todo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   10680
      TabIndex        =   34
      Top             =   3285
      Width           =   1650
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Pegar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   9825
      TabIndex        =   33
      Top             =   3285
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Copiar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   8970
      TabIndex        =   32
      Top             =   3285
      Width           =   855
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Cortar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   8130
      TabIndex        =   31
      Top             =   3285
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0D0F0&
      Caption         =   "SALIDA"
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
      Left            =   5280
      TabIndex        =   28
      Top             =   240
      Width           =   3165
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0D0F0&
         Caption         =   "Pantalla"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   315
         TabIndex        =   30
         Top             =   250
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0D0F0&
         Caption         =   "Impresión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   1
         Left            =   1785
         TabIndex        =   29
         Top             =   250
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.ListBox List3 
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
      Height          =   2055
      ItemData        =   "FRMDISENIO1.frx":0896
      Left            =   360
      List            =   "FRMDISENIO1.frx":0898
      TabIndex        =   17
      Top             =   5805
      Width           =   9670
   End
   Begin VB.TextBox Text1 
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
      Height          =   300
      Left            =   1560
      TabIndex        =   16
      Text            =   " "
      Top             =   4080
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "CONTROLES OPERATIVOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   4770
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2D3C0&
         Caption         =   "ENTREGA"
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
         Left            =   3120
         TabIndex        =   12
         ToolTipText     =   "Marcado / En Caso de no Tener Domicilio de Entrega Toma el Fiscal"
         Top             =   300
         Width           =   1575
         Begin VB.CheckBox Check6 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "No Replica Domicilio Fiscal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   795
            Left            =   120
            TabIndex        =   13
            ToolTipText     =   "Marcado / En Caso de no Tener Domicilio de Entrega Toma el Fiscal"
            Top             =   320
            Width           =   1290
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00E2D3C0&
         Caption         =   "FORMATO NÚMERO DE CUIT"
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
         Left            =   80
         TabIndex        =   10
         Top             =   1600
         Width           =   4620
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Suprime Validación de Formato (Argentina)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   300
            Width           =   4320
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2D3C0&
         Caption         =   "ALTA DE NUEVA CUENTA"
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
         Left            =   80
         TabIndex        =   5
         Top             =   300
         Width           =   3015
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
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
            Height          =   315
            Left            =   2160
            TabIndex        =   6
            Top             =   500
            Width           =   645
         End
         Begin VB.OptionButton Option3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Número Siguiente al"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   120
            TabIndex        =   7
            Top             =   560
            Width           =   2100
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Razón Social en Mayúscula"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   860
            Width           =   2760
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2D3C0&
            Caption         =   "Número Siguiente al Último"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   405
            Left            =   120
            TabIndex        =   8
            Top             =   210
            Value           =   -1  'True
            Width           =   2595
         End
      End
   End
   Begin VB.TextBox Text25 
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
      Height          =   300
      Left            =   1515
      TabIndex        =   3
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox Text26 
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
      Height          =   300
      Left            =   1515
      TabIndex        =   2
      Text            =   " "
      Top             =   3240
      Width           =   1305
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Baja"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   4320
      Picture         =   "FRMDISENIO1.frx":089A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Baja Lógica de la Cuenta"
      Top             =   3960
      Width           =   650
   End
   Begin VB.CommandButton Command21 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   4320
      Picture         =   "FRMDISENIO1.frx":0BA4
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Alta de Nuevo Cliente"
      Top             =   3240
      Width           =   650
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "INDUSTRIAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10080
      TabIndex        =   61
      Top             =   7320
      Width           =   1185
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CONTABILIDAD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10080
      TabIndex        =   60
      Top             =   6840
      Width           =   1185
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDOES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10080
      TabIndex        =   45
      Top             =   6360
      Width           =   1185
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10080
      TabIndex        =   44
      Top             =   5880
      Width           =   1185
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Editable"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -480
      TabIndex        =   27
      Top             =   4200
      Width           =   1905
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No editable"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -480
      TabIndex        =   26
      Top             =   3840
      Width           =   1905
   End
   Begin VB.Label Label81 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9075
      TabIndex        =   24
      Top             =   5565
      Width           =   840
   End
   Begin VB.Label Label80 
      BackStyle       =   0  'Transparent
      Caption         =   "Haber"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7335
      TabIndex        =   23
      Top             =   5565
      Width           =   840
   End
   Begin VB.Label Label79 
      BackStyle       =   0  'Transparent
      Caption         =   "Debe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5625
      TabIndex        =   22
      Top             =   5565
      Width           =   840
   End
   Begin VB.Label Label78 
      BackStyle       =   0  'Transparent
      Caption         =   "Número"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3810
      TabIndex        =   21
      Top             =   5565
      Width           =   840
   End
   Begin VB.Label Label77 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1635
      TabIndex        =   20
      Top             =   5565
      Width           =   1455
   End
   Begin VB.Label Label76 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   435
      TabIndex        =   19
      Top             =   5565
      Width           =   975
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2745
      TabIndex        =   18
      Top             =   3720
      Width           =   150
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Calle y Número"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   270
      TabIndex        =   15
      Top             =   2970
      Width           =   1185
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código Postal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -480
      TabIndex        =   14
      Top             =   3330
      Width           =   1905
   End
   Begin VB.Label Label37 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   360
      TabIndex        =   25
      Top             =   5520
      Width           =   9675
   End
End
Attribute VB_Name = "FRMDISENIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

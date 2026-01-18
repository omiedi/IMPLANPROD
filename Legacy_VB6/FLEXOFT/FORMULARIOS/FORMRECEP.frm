VERSION 5.00
Begin VB.Form FORMRECEP 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Materiales - Altas de Stock en Depósitos"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9870
   Icon            =   "FORMRECEP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   9870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command10 
      Caption         =   "QC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8820
      Picture         =   "FORMRECEP.frx":030A
      TabIndex        =   47
      ToolTipText     =   "Calificación de Proveedores"
      Top             =   3675
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   8820
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   20
      Top             =   4620
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8820
      Picture         =   "FORMRECEP.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Configuración de Comportamiento y Formularios"
      Top             =   1995
      Width           =   855
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5265
      Left            =   4515
      TabIndex        =   5
      Top             =   210
      Width           =   4110
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Salida"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   210
         TabIndex        =   43
         Top             =   4320
         Width           =   3690
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Pantalla"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   525
            TabIndex        =   45
            Top             =   300
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Impresión"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1995
            TabIndex        =   44
            Top             =   300
            Width           =   1170
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Control de Recepciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   210
         TabIndex        =   39
         Top             =   2880
         Width           =   3690
         Begin VB.PictureBox Picture9 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   360
            ScaleHeight     =   585
            ScaleWidth      =   2955
            TabIndex        =   40
            Top             =   420
            Width           =   3015
            Begin VB.CommandButton Command12 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   17.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   2400
               Picture         =   "FORMRECEP.frx":0A56
               Style           =   1  'Graphical
               TabIndex        =   48
               ToolTipText     =   "Busqueda por Lote del Proveedor."
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command11 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   -20
               Picture         =   "FORMRECEP.frx":1DBC
               Style           =   1  'Graphical
               TabIndex        =   41
               Top             =   0
               Width           =   560
            End
            Begin VB.Label LABEL5 
               BackStyle       =   0  'Transparent
               Caption         =   "Recepciones por Rango de Fechas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   42
               Top             =   105
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Consulta de Pendientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2325
         Left            =   210
         TabIndex        =   29
         Top             =   420
         Width           =   3690
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1155
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   36
            Top             =   1470
            Width           =   2220
            Begin VB.CommandButton Command5 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   -20
               Picture         =   "FORMRECEP.frx":20C6
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Material en Consignación"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   38
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   33
            Top             =   945
            Width           =   2220
            Begin VB.CommandButton Command4 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   -20
               Picture         =   "FORMRECEP.frx":23D0
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Pedidos y Orde- nes de Compra"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   35
               Top             =   0
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   30
            Top             =   420
            Width           =   2220
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
               Height          =   600
               Left            =   -20
               Picture         =   "FORMRECEP.frx":26DA
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Boletas de Recepción"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   32
               Top             =   0
               Width           =   1590
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ingresos a Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5265
      Left            =   210
      TabIndex        =   3
      Top             =   210
      Width           =   4110
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Entradas de Fábrica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   210
         TabIndex        =   22
         Top             =   3255
         Width           =   3690
         Begin VB.PictureBox Picture7 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   23
            Top             =   945
            Width           =   2325
            Begin VB.CommandButton Command9 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   -20
               Picture         =   "FORMRECEP.frx":29E4
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label8 
               BackStyle       =   0  'Transparent
               Caption         =   "Devolución       de Materiales"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   840
               TabIndex        =   25
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   26
            Top             =   420
            Width           =   2325
            Begin VB.CommandButton Command6 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   0
               Picture         =   "FORMRECEP.frx":2B2E
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Propia Produccion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   840
               TabIndex        =   28
               Top             =   50
               Width           =   1005
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Recepciones Externas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2745
         Left            =   210
         TabIndex        =   4
         Top             =   420
         Width           =   3690
         Begin VB.Frame Frame9 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Re-Imprimir Comprobantes"
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
            Left            =   315
            TabIndex        =   16
            Top             =   1995
            Width           =   3060
            Begin VB.CommandButton Command17 
               Caption         =   "F.C.R."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Index           =   2
               Left            =   2100
               TabIndex        =   19
               Top             =   250
               Width           =   750
            End
            Begin VB.CommandButton Command17 
               Caption         =   "C.Insp"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Index           =   1
               Left            =   1155
               TabIndex        =   18
               Top             =   250
               Width           =   750
            End
            Begin VB.CommandButton Command17 
               Caption         =   "B.Rec."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   250
               Index           =   0
               Left            =   210
               TabIndex        =   17
               ToolTipText     =   "Ver y Re-Imprimir Boletas de Recepción"
               Top             =   250
               Width           =   750
            End
         End
         Begin VB.PictureBox Picture11 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1155
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   13
            Top             =   1260
            Width           =   2220
            Begin VB.CommandButton Command16 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   0
               Picture         =   "FORMRECEP.frx":2E38
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Ficha de Control de Recepción"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   15
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   9
            Top             =   840
            Width           =   2220
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
               Height          =   600
               Left            =   -20
               Picture         =   "FORMRECEP.frx":2F82
               Style           =   1  'Graphical
               TabIndex        =   10
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Inspección y Aprobación"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   11
               Top             =   0
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   6
            Top             =   420
            Width           =   2220
            Begin VB.CommandButton Command7 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   0
               Picture         =   "FORMRECEP.frx":328C
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Boleta de Recep- ción (Provisoria)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   8
               Top             =   0
               Width           =   1590
            End
         End
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8820
      Picture         =   "FORMRECEP.frx":3596
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Terminar - Salir de esta Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8820
      Picture         =   "FORMRECEP.frx":36E0
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1155
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Direct"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   8820
      Picture         =   "FORMRECEP.frx":39EA
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Acceso Directo a otras Aplicaciones"
      Top             =   2835
      Width           =   855
   End
   Begin VB.Label NORVENTA 
      Caption         =   "NORVENTA"
      Height          =   225
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   8260
      Picture         =   "FORMRECEP.frx":3CF4
      Top             =   4940
      Width           =   2010
   End
End
Attribute VB_Name = "FORMRECEP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command10_Click()
Command10.Enabled = False
Call CONECRUN("CALPRO02", "Sistema Calificación de Proveedores")
Command10.Enabled = True
End Sub

Private Sub Command11_Click()
    '
    If Option1.Value = True Then
       Call MENSERR(24, "Opcion no Disponible por Pantalla.\  \Emita Listado Impreso y Exporte a Archivo de Texto.")
       Exit Sub
    End If
    '
20  HOI% = HOY(HO$)
    FE$ = "01" + Mid$(HO$, 3)
    FF1% = FECHANUM(FE$): FF2% = HOI%
    XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
    XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0) + MKI$(FF1%) + MKI$(FF2%)
    '
22  RACOD$ = OBJPLA$("DESHASCOD2", VDEF$)
    If RACOD$ = "" Then
        GoTo 99
    End If
    CI1% = CVI(Left$(RACOD$, 2)): cod1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): cod2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TMT% = 0 Or TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(RACOD$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación"
    End If
    FF1% = CVI(Mid$(RACOD$, 6, 2))
    FEX = FF1%: FLI1$ = FECHATEX$(FEX)
    FF2% = CVI(Mid$(RACOD$, 8, 2))
    FEX = FF2%: FLI2$ = FECHATEX$(FEX)
    '
    If cod1$ = "" Or cod2$ = "" Or cod2$ < cod1$ Or FF2% < FF1% Then
        VDEF$ = RACOD$
        GoTo 22
    End If
    '
    RANGO$ = TRIM$(cod1$) + " - " + TRIM$(cod2$)
    RANGO$ = RANGO$ + " / " + TRIM$(TIMAT$)
    RANGO$ = RANGO$ + " / " + FLI1$ + " - " + FLI2$
    '
    Screen.MousePointer = 11
    '
    Call COPYSTRU("BOLRECEP", "BOLSELEC")
    J& = 0
    FIN& = ULTREG&("BOLRECEP")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("BOLRECEP", U&)
       CIIX% = CVI(Mid$(X$, 83, 2))
       FERE% = CVI(Mid$(X$, 81, 2))
       '
       If FERE% >= FF1% Then
         If FERE% <= FF2% Then
           If CVD(Mid$(X$, 89, 8)) > 0.0005 Then ' CANTIDAD NO NULA
             If TMT% = 0 Or TIMATE%(CIIX%) = TMT% Then
               CODE$ = CODEXT$(CIIX%)
               If CODE$ >= cod1$ Then
                 If CODE$ <= cod2$ Then
                   J& = J& + 1
                   Call GRAREG("BOLSELEC", X$, J&)
                 End If
               End If
             End If
           End If
         End If
       End If
       '
    Next U&
    '
    Call SETULTREG("BOLSELEC", J&)
    Call CIERRARCH("BOLSELEC")
    Call CIERRARCH("*.*")
    '
    Call HEADERS("BOLSELEC", "")
    Call HEADERS("BOLSELEC", "Rango: " + RANGO$)
    Call CONECRUN("*BOLSELEC", "Recepciones por Rango de Fechas")
    '
99 Screen.MousePointer = 1
   '
End Sub

'\\\OT-05-0543-WAR-27/07/05///
Private Sub Command12_Click()
    '
    If Option1.Value = True Then
       Call MENSERR(24, "Opcion no Disponible por Pantalla.\  \Emita Listado Impreso y Exporte a Archivo de Texto.")
       Exit Sub
    End If
    '
    LP1$ = TRIM$(UCase$(InputBox("Ingrese el Lote del Proveedor", "Busqueda por Lote del Proveedor")))
    If TRIM$(LP1$) <> "" Then
         '
         Screen.MousePointer = 11
         Call COPYSTRU("BOLRECEP", "BOLSELEC")
         J& = 0
         FIN& = ULTREG&("BOLRECEP")
         For U& = 1 To FIN&
            Call TRACE(20, U&, FIN&)
            X$ = REGLEIDO$("BOLRECEP", U&)
            NOLOPRO$ = TRIM(Mid$(X$, 107, 12))
            '
            If CVD(Mid$(X$, 89, 8)) > 0.0005 Then ' CANTIDAD NO NULA
            NOLOPRO1$ = UCase$(NOLOPRO$)
              If InStr(NOLOPRO1$, LP1$) > 0 Then
                J& = J& + 1
                Call GRAREG("BOLSELEC", X$, J&)
              End If
            End If
            '
         Next U&
         '
         Call SETULTREG("BOLSELEC", J&)
         Call CIERRARCH("BOLSELEC")
         Call CIERRARCH("*.*")
         '
         If J& >= 1 Then
           Call HEADERS("BOLSELEC", "")
           Call HEADERS("BOLSELEC", "Lote Proveedor: " + LP1$)
           Call CONECRUN("*BOLSELEC", "Recepciones por Rango de Fechas")
          Else
           Call MENSERR(24, "No se Encontraron Coincidencias.")
         End If
         '
99      Screen.MousePointer = 1
        '
     Else
      'Call MENSERR(24, "Lote de Proveedor no Válido.")
      Exit Sub
    End If
    '
End Sub
'\\\OT-05-0543-WAR-FIN///


Private Sub Command14_Click()
    '
    OPRECEPMA.Show 1
    
'    On ALTERNA%("Tabla de Tipos de Recepción\Formulario de Impresión de B-R") GoTo 10, 20
'    Exit Sub
'    '
'10  If ULTREG&("TATIREC") < 1 Then
'        X$ = REGBLAN$("TATIREC")
'        Call REPLA(X$, Chr$(1), 1, 1)
'        Call REPLA(X$, "Material General", 2, 31)
'        Call REPLA(X$, "BOLREGEN", 49, 16)
'        Call REPLA(X$, "BOLREGEN", 65, 16)
'        Call REGAPP("TATIREC", X$)
'        Call CIERRARCH("TATIREC")
'    End If
'    '
'    Call BLANARCH("!TATIREC")
'    For I& = 1 To ULTREG&("TATIREC")
'        X$ = REGLEIDO$("TATIREC", I&)
'        Call REGAPP("!TATIREC", X$)
'    Next I&
'    VTB% = VENTABU%("TATIREC")
'    If VTB% >= 0 Then
'        For I& = 1 To ULTREG&("!TATIREC")
'            X$ = REGLEIDO$("!TATIREC", I&)
'            I1% = I& Mod 256
'            Call REPLA(X$, Chr$(I1%), 1, 1)
'            Call GRAREG("TATIREC", X$, I&)
'        Next I&
'    End If
'    Call CIERRARCH("TATIREC")
'    Call CIERRARCH("!TATIREC")
'    Exit Sub
    '
'20  Call CONECRUN("PRIFORMS", "Formularios de Impresion")
    '
End Sub

Private Sub Command16_Click()
    Call CONECRUN("#FICHRED", "Ficha de Control de Recepción")
End Sub

Private Sub Command17_Click(Index As Integer)
    If Index = 0 Then
          Call SHOWBORE
          Exit Sub
       ElseIf Index = 2 Then
          Call CONECRUN("#FICHRED", "Ficha de Control de Recepción")
          Exit Sub
       Else
          Call OPNOIN("")
    End If
End Sub

Private Sub Command2_Click()
    Call APROBOL
End Sub

Private Sub Command3_Click()
    If Option1.Value = True Then
        Call FINAL("*COBOLPE")
      ElseIf Option2.Value = True Then
        Call FINAL("*LIBOLPE")
    End If
End Sub

Private Sub Command4_Click()
    '
    If Option1.Value = True Then
        OPXX% = ALTERNA%("Base de Datos Completa\Ordenes Abiertas\Cerradas\Anuladas")
        Select Case OPXX%
          Case 1
            Call FINAL("*COCOMP0")
          Case 2
            Call OPNOIN("")
          Case 3
            Call OPNOIN("")
          Case 4
            Call OPNOIN("")
          Case Else
            Exit Sub
        End Select
        '
      ElseIf Option2.Value = True Then
        Call FINAL("*FOCOMP0")
    End If
    '
End Sub

Private Sub Command5_Click()
    '
    If Option1.Value = True Then
        Call FINAL("*COMATCO")
      ElseIf Option2.Value = True Then
        Call FINAL("*LIMATCO")
    End If
    '
End Sub

Private Sub Command6_Click()
    Call RECPROP
End Sub

Private Sub Command7_Click()
    Call BOLRECEP
End Sub

Private Sub Command8_Click()
    OPX% = ALTERNA%("Archivos de Transferencia\Otros Accesos Directos")
    If OPX% = 1 Then
         Call ATRANSFE
       ElseIf OPX% = 2 Then
         Call ACCDIR("RECEPMA")
    End If
End Sub

Private Sub Form_Load()
    '
    VERANTER$ = "RECEPMA"
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    If EXISTE%("CALPRO02.EXE") > 0 Then
      Command10.Visible = True
    End If
    '
End Sub

Sub APROBOL() ' aprobacion de boleta de recepcion ************************************
        '
        Dim REGBO&(2048)
        Call GENIBOLPE
        '
1120    'Call CLEARVEN
        X$ = "": Y$ = "": Z$ = "": YY$ = ""
        EBOL$ = "": PEDY$ = "": DESCRI$ = ""
        NBR$ = OBJPLA$("SELBOLPEN", "")
        If NBR$ = "" Then
            Exit Sub
        End If
        '
        NROBOL& = CVS(Left$(NBR$, 4))
        If NROBOL& < 1 Then
            Exit Sub
        End If
        '
        'Call BLIMESS("Seleccionando ...")
        FIN& = ULTREG&("BOLREPEN")
        Call BLANARCH("!BOLREDET")
        CABOPE% = 0
        For I& = 1 To FIN&
           X$ = REGLEIDO$("BOLREPEN", I&)
           If CVS(Left$(X$, 4)) = NROBOL& Then
              CANOM# = CVD(Mid$(X$, 89, 8))
              APRO# = CVD(Mid$(X$, 163, 8))
              RECHA# = CVD(Mid$(X$, 171, 8))
              FALTA# = CVD(Mid$(X$, 179, 8))
              '
              If CANOM# < 0.05 Or APRO# > 0.05 Or RECHA# > 0.05 Or Abs(FALTA#) > 0.05 Then GoTo 1121
              Call REPLA(X$, MKD$(CANOM#), 163, 8)
              '
              If CABOPE% = 0 Then
                  NC% = CVI(Mid$(X$, 105, 2)): NC1% = (-1) * NC%
                  FF% = CVI(Mid$(X$, 81, 2))
                  FEX = FF%: FECHAREC$ = FECHATEX$(FEX)
                  REMI$ = Mid$(X$, 119, 16)
                  DP1% = Asc(Mid$(X$, 97, 1))
                  If ECOARCH$("ECODEP", Chr$(DP1%)) = "" Then
                      DP1% = 0
                  End If
                  '
                  Y$ = String$(128, 0)
                  Call REPLA(Y$, MKS$(NROBOL&), 1, 4)
                  Call REPLA(Y$, MKI$(FF%), 5, 2)
                  Call REPLA(Y$, MKI$(NC%), 7, 2)
                  Call REPLA(Y$, RASOCLI$(NC1%), 9, 32)
                  Call REPLA(Y$, DOMICLI$(NC1%), 41, 32)
                      CLOCA$ = TRIM$(CPOSCLI$(NC1%)): If CLOCA$ <> "" Then CLOCA$ = CLOCA$ + " "
                  Call REPLA(Y$, CLOCA$ + LOCACLI$(NC1%), 73, 32)
                  Call REPLA(Y$, REMI$, 105, 16)
                  Call REPLA(Y$, Chr$(DP1%), 121, 1)
                  AINDI% = AINDI% + 1
              End If
              '
              CABOPE% = CABOPE% + 1: REGBO&(CABOPE%) = I&
              Call REGAPP("!BOLREDET", X$)
              '
           End If
           '
1121    Next I&
        '
        'Call BLIMESS("-")
        If CABOPE% < 1 Then
            Call COMUNI("No hay Recepciones Pendientes\para la Boleta Elegida")
            GoTo 1120
        End If
        '
11215   Z$ = OBJPLA$("BOLREC-MUESTRW", Y$)
        If Z$ = "" Then
            GoTo 1120
        End If
        DP1% = Asc(Mid$(Z$, 121, 1))
        If DP1% < 1 Or ECOARCH$("ECODEP", Chr$(DP1%)) = "" Then
            Call MENSERR(24, "Falta Indicar Depósito de Entrada")
            GoTo 11215
        End If
        '
1122    VTB% = VENTABU%("APROBOL/NC/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(RASOCLI$(NC1%)) + " - " + FECHAREC$)
        ATX% = 0: HAYRECHA% = 0
        For U& = 1 To ULTREG&("!BOLREDET")
           X$ = REGLEIDO$("!BOLREDET", U&)
           CANOM# = CVD(Mid$(X$, 89, 8))
           APRO# = CVD(Mid$(X$, 163, 8))
           RECHA# = CVD(Mid$(X$, 171, 8))
             If RECHA# >= 0.05 Then HAYRECHA% = HAYRECHA% + 1
           FALTA# = CVD(Mid$(X$, 179, 8))
           If Abs(APRO#) < 0.005 Then
             If Abs(RECHA#) < 0.005 Then
               If Abs(FALTA#) < 0.005 Then
                  GoTo 1123                        ' NO SE PROCESO MOVIMIENTO
               End If
             End If
           End If
           '
           FALTB# = CANOM# - APRO# - RECHA#
           If Abs(FALTB# - FALTA#) > 0.005 Then
              ATX% = ATX% + 1
              Call REPLA(X$, MKD$(FALTB#), 179, 8)
              Call GRAREG("!BOLREDET", X$, U&)
           End If
1123    Next U&
        '
        If ATX% > 0 Then
            Call COMUNI("Cantidades Incoherentes.\Revise y Corrija !")
            GoTo 1122
        End If
        '
1124    If VTB% <> 0 Then GoTo 1120
        '
        PORECHLIN% = 0
        CODRECHA$ = "BR"
        REMIRECH$ = "BR." + TRIM$(Str$(NROBOL&))
        If HAYRECHA% > 0 Then
          '\\\OT-05-0619-WAR-27/09/05///
          If EXISTE%("RECHLI02.EXE") > 0 Then
            If HAYRECHA% > 1 Then
              Call MENSERR(24, "Imposible Procesar más de un Rechazo.")
              GoTo 1120
            End If
            PORECHLIN% = 1
            FERECHA% = CVI(Mid$(X$, 81, 2))
            CIRECHA% = CVI(Mid$(X$, 83, 2))
            CARECHA = RECHA#
            GoTo 1125
          End If
          '\\\OT-05-0619-WAR-FIN///
          If TRIM$(CONTROL$("", "REMIRECH")) <> "" Then
            CODRECHA$ = TRIM$(CONTROL$("", "CODRRECH"))
            If CODRECHA$ = "" Then CODRECHA$ = "RD"
            FIN& = ULTREG&("MOVISTO")
            For U& = FIN& To 1 Step -1
               Call TRACE(20, FIN& - U& + 1, FIN&)
               XX$ = REGLEIDO$("MOVISTO", U&)
               If UCase$(Mid$(XX$, 21, 2)) = CODRECHA$ Then
                 UNRORECH& = Val(Mid$(XX$, 26, 7))
                 GoTo 11241
               End If
            Next U&
            UNRORECH& = 0
            '
11241       REMIRECH$ = CODRECHA$ + "-" + TRIM$(Str$(UNRORECH& + 1))
            '
          End If
        End If
        '
1125    Call CIERRARCH("MOVISTO")
        Call CIERRARCH("BOLREPEN")
        Call CIERRARCH("BOLRECEP")
        Call BLOQARCH("MOVISTO")
        Call BLOQARCH("BOLREPEN")
        Call BLOQARCH("BOLRECEP")
        '
        QUEDAPEN% = 0
        HAYRECHA% = 0
        For U& = 1 To ULTREG&("!BOLREDET")
           X$ = REGLEIDO$("!BOLREDET", U&)
           CI% = CVI(Mid$(X$, 83, 2))
           REFX$ = Mid$(X$, 5, 64)
           LOTE$ = Mid$(X$, 69, 12)
           CANOM# = CVD(Mid$(X$, 89, 8))
           APRO# = CVD(Mid$(X$, 163, 8))
           RECHA# = CVD(Mid$(X$, 171, 8))
           FALTA# = CVD(Mid$(X$, 179, 8))
           If Abs(APRO#) < 0.005 Then
             If Abs(RECHA#) < 0.005 Then
               If Abs(FALTA#) < 0.005 Then
                  QUEDAPEN% = 1
                  GoTo 1128                        ' NO SE PROCESO MOVIMIENTO
               End If
             End If
           End If
           '
           UX% = U&
           PRIREG& = ULTREG&("BOLREPEN")
           REG& = REGBO&(UX%)
           If REG& > 0 Then
              If REG& <= ULTREG&("BOLREPEN") Then
                  Y$ = REGLEIDO$("BOLREPEN", REG&)
                  If Left$(Y$, 68) <> Left$(X$, 68) Then
                      Call COMUNI("Punteros Incorrectos.\Consulte Soporte de Sistemas")
                      GoTo 1128
                  End If
                  Call GRAREG("BOLREPEN", X$, REG&)
                  If REG& < PRIREG& Then PRIREG& = REG&
                  '
                  REG1& = CVS(Mid$(X$, 189, 4))
                  If REG1& > 0 Then
                     If REG1& <= ULTREG&("BOLRECEP") Then
                        Y$ = REGLEIDO$("BOLRECEP", REG1&)
                        If Left$(Y$, 68) <> Left$(X$, 68) Then
                           Call COMUNI("Punteros Incorrectos.\Consulte Soporte de Sistemas")
                           GoTo 1127
                        End If
                        Call GRAREG("BOLRECEP", X$, REG1&)
                     End If
                  End If
                  '
1127              DOPRI$ = "BR." + TRIM$(Str$(NROBOL&))
                  DEPX% = DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(CI%)
                  If CANOM# >= 0.05 Then
                      CANTI = CANOM#
                      REFE$ = "Rto." + TRIM$(REMI$) + " - " + RASOCLI$(NC1%)
                      PPX% = InStr(REFX$, "- BC D.")
                      If PPX% > 0 Then
                          REFE$ = Mid$(REFX$, PPX% + 5)
                      End If
                      DOSEC$ = TRIM$(REMI$)
                      While Len(DOSEC$) > 10: DOSEC$ = Mid$(DOSEC$, 2): Wend
                      If Len(DOSEC$) < 8 Then DOSEC$ = "Rt." + DOSEC$
                      Call MOVISTO(FF%, CI%, LOTE$, "BR", DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC1%, DEPX%, CANTI)
                  End If
                  '
                  If RECHA# >= 0.05 Then
                    If PORECHLIN% = 0 Then
                      CANTI = RECHA# * (-1)
                      REFE$ = "RECHAZO - Rto." + TRIM$(REMI$)
                      Call MOVISTO(FF%, CI%, LOTE$, CODRECHA$, REMIRECH$, DOSEC$, REFE$, VUNI#, MONE$, NC1%, DEPX%, CANTI)
                      HAYRECHA% = 1
                    End If
                  End If
                  '
                  If Abs(FALTA#) >= 0.05 Then
                      CANTI = FALTA# * (-1)
                      REFE$ = "FALTANTE - Rto." + TRIM$(REMI$)
                      If CANTI > 0 Then
                         REFE$ = "SOBRANTE - Rto." + TRIM$(REMI$)
                      End If
                      Call MOVISTO(FF%, CI%, LOTE$, "BR", DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC1%, DEPX%, CANTI)
                  End If
                  '
              End If
           End If

1128    Next U&
        '
        Call CIERRARCH("MOVISTO")
        '
        ' DEPURAR BOLREPEN
        '
        J& = 0: K& = 0: NOOBA& = 0
        For U& = 1 To ULTREG&("BOLREPEN")
           X$ = REGLEIDO$("BOLREPEN", U&)
           CANOM# = CVD(Mid$(X$, 89, 8))
           APRO# = CVD(Mid$(X$, 163, 8))
           RECHA# = CVD(Mid$(X$, 171, 8))
           FALTA# = CVD(Mid$(X$, 179, 8))
           If CANOM# > 0.005 Then
             If Abs(APRO#) < 0.005 Then
               If Abs(RECHA#) < 0.005 Then
                 If Abs(FALTA#) < 0.005 Then
                   J& = J& + 1
                   Call GRAREG("BOLREPEN", X$, J&)
                 End If
               End If
             End If
           End If
        Next U&
        '
        Call SETULTREG("BOLREPEN", J&)
        Call CIERRARCH("BOLREPEN")
        Call CIERRARCH("BOLRECEP")
        Call CIERRARCH("*.*")
        Call BAJALDISCO
        '
        NUVALIN& = 1
        If EXISTE%(LU$ + "BRAP-ARC.FRM") > 0 Then
          For KUU& = ULTREG&("PDOCLST") To 1 Step -1
            XYX$ = Mid$(REGLEIDO$("PDOCLST", KUU&), 5)
            If Left$(XYX$, 25) = "Vale Aprobación Recepción" Then
              PPXY% = InStr(26, XYX$, "-")
              If PPXY% < 27 Then PPXY% = 1 + Len(XYX$)
              NUVALIN& = 1 + Val(Mid$(XYX$, 26, PPXY% - 26))
              GoTo 1129
            End If
          Next KUU&
        End If
        '
1129    LU$ = LUDAT$
        If EXISTE%(LU$ + "BRAP-ARC.FRM") > 0 Then
           HOII% = HOY(HOS$)
           FECHDOC$ = HOS$
           NUMEDOC$ = TRIM$(Str$(NUVALIN&))
           TIPODOC$ = "Vale Aprobación Recepción"
           DOCUORIG$ = "BR." + TRIM$(Str$(NROBOL&))
           '
           'CODPARAM$(1) ="TRANSPOR": DOCPARAM$(1) = TRIM$(FOREMIFAC.Text8.TEXT)
           CAPARDOC% = 0
           '
           CODTABU1$(1) = "COD-MAT"
           CODTABU1$(2) = "DES-MAT"
           CODTABU1$(3) = "CANTIDAD"
           CACOLTAB1% = 3
           '
           DESTIDOC% = 0
           CAITAB1% = 0
           '
           For YX& = 1 To ULTREG&("!BOLREDET")
              YY$ = REGLEIDO$("!BOLREDET", YX&)
              CI% = CVI(Mid$(YY$, 83, 2))
              CAN# = CVD(Mid$(YY$, 163, 8))
              If CI% > 0 Then
                If CI% <= NUMAS% Then
                  If CAN# > 0.005 Then
                    '
                    'UNID$ = UNIMED$(CI%)
                    'If UNID$ = "" Then UNID$ = UNIMED$(CI%)
                    CANS$ = CSTRING$(MKD$(CAN#), 3, 12, 1, 2)
                    '
                    CAITAB1% = CAITAB1% + 1
                    CODDX$ = CODEXT$(CI%)
                    TETABU1$(1, CAITAB1%) = CODDX$
                    TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
                    TETABU1$(3, CAITAB1%) = CANS$
                  End If
                End If
              End If
           Next YX&
           '
           Call PRINTDOC("@BRAP-ARC")   ' Remito de Facturacion
           '
        End If
        '
        If HAYRECHA% = 1 Or PORECHLIN% > 0 Then
           '\\\OT-05-0619-WAR-27/09/05///
           If PORECHLIN% = 1 Then
             DOPRI12$ = "BR." + TRIM$(Str$(NROBOL&))
             Call PROCERECHA(CIRECHA%, FERECHA%, CARECHA, DOPRI12$) '\\\OT-06-0191-WAR-09/06/06///
             Call RECONECTA
             Call FINAL("")
             Exit Sub
           End If
           '\\\OT-05-0619-WAR-FIN///
           Call REMIRECHA(NC1%, DEPX%, REMIRECH$)
        End If
        '
        Call GENIBOLPE
        '
99899   Call CIERRARCH("*.*")
        Call BAJALDISCO
        If EXISTE%("CALPRO02.EXE") > 0 Then
          ICALREC.Label5 = TRIM$(Str$(NROBOL&))
          ICALREC.Show 1
        End If
        GoTo 1120
        '
End Sub
'
Sub BOLRECEP()
'
'falta generar ficha de recepcion fervi-air
        '
        If ULTREG&("TATIREC") < 1 Then
            X$ = REGBLAN$("TATIREC")
            Call REPLA(X$, Chr$(1), 1, 1)
            Call REPLA(X$, "Material General", 2, 31)
            Call REPLA(X$, "BOLREGEN", 49, 16)
            Call REPLA(X$, "BOLREGEN", 65, 16)
            Call REGAPP("TATIREC", X$)
            Call CIERRARCH("TATIREC")
        End If
        '
1100    Call CIERRARCH("*.*")
        '
        NROBOL& = 1
        If ULTREG&("BOLRECEP") > 0 Then
            UREG& = ULTREG&("BOLRECEP")
            X$ = REGLEIDO$("BOLRECEP", UREG&)
            If Len(X$) >= 4 Then NROBOL& = 1 + CVS(Left$(X$, 4))
        End If
        '
        If TIMATEI% = 0 Then
            X$ = REGLEIDO$("TATIREC", 1)
            TIMATEI% = Asc(Left$(X$, 1))
        End If
        '
        VDEF$ = MKI$(NC%) + Chr$(TIMATEI%)
        EBOL$ = OBJPLA$("SELPROVREC", VDEF$)
        If EBOL$ = "" Then
            Exit Sub
        End If
        '
        NC% = CVI(EBOL$)
        If NC% < 1 Or ECOARCH$("PROVED1", MKI$(NC%)) = "" Then
            Call MENSERR(24, "Proveedor Inexistente")
            GoTo 1100
        End If
        '
        TIMATEI% = Asc(Mid$(EBOL$, 3, 1))
        If TIMATEI% < 1 Or ECOARCH("TATIREC", Chr$(TIMATEI%)) = "" Then
            Call MENSERR(24, "Falta Tipo de Material")
            GoTo 1100
        End If
        '
        REGT& = TIMATEI%
        X$ = REGLEIDO$("TATIREC", REGT&)
        FORMPLA$ = TRIM$(UCase$(Mid$(X$, 33, 16)))
        FORMTAB$ = TRIM$(UCase$(Mid$(X$, 49, 16)))
        FORIPRE$ = TRIM$(UCase$(Mid$(X$, 65, 16)))
        '
        NC1% = (-1) * NC%
        '
1105    Y$ = PEDY$
        If Y$ = "" Then
            Y$ = Space$(128)
            Call REPLA(Y$, Chr$(0), 121, 1)
        End If
        '
        Call REPLA(Y$, MKS$(NROBOL&), 1, 4)
        Call REPLA(Y$, EBOL$, 7, 2)
        Call REPLA(Y$, MKI$(HOY(HO$)), 5, 2)
        Call REPLA(Y$, RASOCLI$(NC1%), 9, 32)
        Call REPLA(Y$, DOMICLI$(NC1%), 41, 32)
        CLOCA$ = TRIM$(CPOSCLI$(NC1%)): If CLOCA$ <> "" Then CLOCA$ = CLOCA$ + " "
        Call REPLA(Y$, CLOCA$ + LOCACLI$(NC1%), 73, 32)
        Call REPLA(Y$, Chr$(0), 121, 1)
       '
        OBJE$ = "BOLREC-ENCB"
        If CONTROL$("", "NUMEBREC") = "MANUAL" Then
            OBJE$ = "BOLREC-ENCB-M"
        End If
        '
        TICOMPRO$ = TRIM$(CONTROL$("", "COMPBREC"))
        '
        PEDY$ = OBJPLA$(OBJE$, Y$)
        If PEDY$ = "" Then Exit Sub
        '
        FF% = CVI(Mid$(PEDY$, 5, 2))
        FEX = FF%: FECHAREC$ = FECHATEX$(FEX)
        If TRIM$(FECHAREC$) = "" Then
            Call MENSERR(24, "Fecha Incorrecta")
            GoTo 1105
        End If
        '
        REMI$ = TRIM$(Mid$(PEDY$, 105, 16))
        If REMI$ = "" Then
            Call MENSERR(24, "Falta Número de Remito del Proveedor")
            GoTo 1105
        End If
        NFAC$ = TRIM$(Mid$(PEDY$, 129, 16))
        '
        DP1% = Asc(Mid$(PEDY$, 121, 1))
        DEPOSAL$ = TRIM$(ECOARCH$("ECODEP", Chr$(DP1%)))
        If DP1% < 1 Or DEPOSAL$ = "" Then
            Call MENSERR(24, "Falta Depósito de Entrada")
            GoTo 1105
        End If
        '
        NROPRO& = NROBOL&
        NROBOL& = CVS(Left$(PEDY$, 4))
        If NROBOL& < 1 Then
            Call MENSERR(24, "Número Incorrecto de B.R.")
            NROBOL& = NROPRO&
            Call REPLA(PEDY$, MKS$(NROBOL&), 1, 4)
            GoTo 1105
        End If
        '
        If FORMPLA$ = "RECEP-ESPUMA" Then
            CONTA% = 0: CONTB% = 0
            Call BLANARCH("!BOLREDET")
            '
1111        CONTB% = CONTB% + 1
            VDEF$ = MKI$(CONTB%) + String$(46, 0)
            Call REPLA(VDEF$, MKI$(1), 29, 2)
11111       XXX$ = OBJPLA$("RECEP-ESPUMA", VDEF$)
            If XXX$ = "" Then GoTo 1115
            COMAT% = CVI(Mid$(XXX$, 3, 2))
            COLO$ = Mid$(XXX$, 5, 8)
            LARGO = CVS(Mid$(XXX$, 13, 4))
            ANCHO = CVS(Mid$(XXX$, 17, 4))
            ALTO = CVS(Mid$(XXX$, 21, 4))
            KILOS = CVS(Mid$(XXX$, 25, 4))
            CABLO% = CVI(Mid$(XXX$, 29, 2))
            OBS$ = Mid$(XXX$, 31, 18)
            NUOC& = CVS(Mid$(XXX$, 49, 4))
            '
            If COMAT% < 1 Or LARGO < 100 Or ANCHO < 100 Or ALTO < 100 Or KILOS < 1 Or CABLO% < 1 Or CABLO% > 16 Then
                Call MENSERR(24, "Imposible Procesar")
                GoTo 11111
            End If
            '
            PROMKIL# = CDbl(KILOS / CABLO%)
            VOLU = (LARGO / 1000) * (ANCHO / 1000) * (ALTO / 1000)
            DENSI = (Int(10 * PROMKIL# / VOLU + 0.5)) / 10
            DNX$ = TRIM$(Str$(DENSI))
            '
            For CBL% = 1 To CABLO%
               CONTA% = CONTA% + 1
               IDENLO$ = "BR-" + TRIM$(Str$(NROBOL&)) + "-" + TRIM$(Str$(CONTA%))
               '
               TTT$ = REGBLAN$("!BOLREDET")
               Call REPLA(TTT$, MKS$(NROBOL&), 1, 4)
               DESCRI$ = TRIM$(CODEXT$(COMAT%)) + " - BC D." + DNX$ + " "
                  DESCRI$ = DESCRI$ + TRIM$(Str$(LARGO)) + "x" + TRIM$(Str$(ANCHO)) + "x" + TRIM$(Str$(ALTO))
               Call REPLA(TTT$, DESCRI$, 5, 64)
               Call REPLA(TTT$, IDENLO$, 69, 12)
               Call REPLA(TTT$, MKI$(FF%), 81, 2)
               Call REPLA(TTT$, MKI$(COMAT%), 83, 2)
               Call REPLA(TTT$, "KG", 85, 4)
               Call REPLA(TTT$, MKD$(PROMKIL#), 89, 8)
               DEPX% = DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(COMAT%)
               Call REPLA(TTT$, Chr$(DEPX%), 97, 1)
               Call REPLA(TTT$, MKI$(NC%), 105, 2)
               Call REPLA(TTT$, COLO$, 107, 12)
               Call REPLA(TTT$, REMI$, 119, 16)
               Call REPLA(TTT$, MKS$(NUOC&), 151, 4)
               '
               Call REGAPP("!BOLREDET", TTT$)
            Next CBL%
            '
            GoTo 1111
            '
        End If

1112    Call CIERRARCH("*.*")
        '
        JJ& = 0
        '
        DARX$ = TRIM$(CONTROL$("BOLRECEP", "DIANREC"))
        If DARX$ = "" Then
            DIANRE& = 10000
          Else
            DIANRE& = Val(DARX$)
        End If
        HOIX1 = HOY(HOS$)
        '
        ANUCA% = 0
        If CONTROL$("BOLRECEP", "ANUCANTI") = "SI" Then
          ANUCA% = 1
        End If
        '
        Screen.MousePointer = 11
        UREOCO& = ULTREG&("OCOMDETA")
        HAYOC% = 0: HAYLIB& = 0
        Call SETULTREG("!TXOCOLIB", 0)
        RFF$ = RECFILT$("OCOMDETA", 3, MKI$(NC%))
        For UI& = 1 To Len(RFF$) Step 4
          REDET& = CVS(Mid$(RFF$, UI&, 4))
          If REDET& > 0 Then
            If REDET& <= UREOCO& Then
              X$ = REGLEIDO$("OCOMDETA", REDET&)
              If CVI(Mid$(X$, 7, 2)) = NC% Then
                If Asc(Mid$(X$, 124, 1)) < 8 Then ' abierta Y NO CUMPLIDA
                  CAPEN = CVS(Mid$(X$, 55, 4)) - CVS(Mid$(X$, 99, 4))
                  COEFUNI = CVS(Mid$(X$, 63, 4))
                  If COEFUNI < 0.0001 Then COEFUNI = 1
                  If CAPEN * COEFUNI >= 0.05 Then
                    NUOC& = CVS(Left$(X$, 4))
                    FENTRESO% = CVI(Mid$(X$, 95, 2))
                    FESO2 = FENTRESO%
                    If DIENTRE%(HOIX1, FESO2) <= DIANRE& Then
                      CI% = CVI(Mid$(X$, 9, 2))
                      LOTECO$ = Mid$(X$, 105, 12)
                      MONEXO$ = "$": If Asc(Mid$(X$, 122, 1)) > 1 Then MONEXO$ = "U$S"
                      PRUNX = CVD(Mid$(X$, 79, 8))
                      '
                      'PARA DETECTAR ORDENES DE COMPRA LIBRES
                      RGG$ = RECFILT$("TXOCOLIB", 1, MKS$(NUOC&))
                      If Len(RGG$) >= 4 Then
                         HAYLIB& = HAYLIB& + 1
                         RETX& = CVS(Left$(RGG$, 4))
                         XYZ$ = REGLEIDO$("TXOCOLIB", RETX&)
                         Call GRAREG("!TXOCOLIB", XYZ$, HAYLIB&)
                         GoTo 1119
                      End If
                      '
                      Y$ = REGBLAN$("!BOLREDET")
                      Call REPLA(Y$, MKI$(CI%), 83, 2)
                      Call REPLA(Y$, UNIMED$(CI%), 85, 4)
                      Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 89, 8)
                      If ANUCA% = 1 Then
                        Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 163, 8)
                        Call REPLA(Y$, MKD$(0), 89, 8)
                      End If
                      Call REPLA(Y$, LOTECO$, 107, 12)
                      Call REPLA(Y$, MKS$(NUOC&), 151, 4)
                      Call REPLA(Y$, MONEXO$, 155, 4)
                      Call REPLA(Y$, MKS$(PRUNX), 159, 4)
                      Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
                      Call REPLA(Y$, MKS$(REDET&), 189, 4)
                      HAYOC% = 1
                      JJ& = JJ& + 1
                      Call GRAREG("!BOLREDET", Y$, JJ&)
                    End If
                  End If
                End If
              End If
            End If
          End If
1119    Next UI&
        '
        Call SETULTREG("!TXOCOLIB", HAYLIB&)
        If HAYOC% = 0 Then
            Call COMUNI("No hay Ordenes de Compra Pendientes\para este Proveedor.")
        End If
        '
        For I& = 1 To ULTREG&("MACONSIG")
          X$ = REGLEIDO$("MACONSIG", I&)
          If CVI(Mid$(X$, 3, 2)) = NC% Then
            CAPEN = CVD(Mid$(X$, 63, 8)) - CVD(Mid$(X$, 111, 8))
            If CAPEN > 0 Then
              NUOC& = Val(Mid$(X$, 24, 6))
              CI% = CVI(Mid$(X$, 33, 2))
              Y$ = REGBLAN$("!BOLREDET")
              Call REPLA(Y$, MKI$(CI%), 83, 2)
              Call REPLA(Y$, UNIMED$(CI%), 85, 4)
              Call REPLA(Y$, MKD$(0), 89, 8)
              Call REPLA(Y$, "MP " + TRIM$(CSTRING$(MKS$(CAPEN), 3, 12, 1, 2)) + " " + UNIMED$(CI%), 107, 12)
              Call REPLA(Y$, MKS$(NUOC&), 151, 4)
              Call REPLA(Y$, MKS$((-1) * I&), 189, 4)
              For UI& = 1 To JJ&
                If Left$(Y$, 154) = Left$(REGLEIDO$("!BOLREDET", UI&), 154) Then GoTo 11129 ' SUPRIME SI COINCIDE
              Next UI&
              JJ& = JJ& + 1
              Call GRAREG("!BOLREDET", Y$, JJ&)
            End If
          End If
11129   Next I&
        Call SETULTREG("!BOLREDET", JJ&)
        Call CIERRARCH("!BOLREDET")
        '
        'DETECTAR RECEPCIONES LIBRES
        If HAYLIB& > 0 Then
           RESP4% = COMALTER("Elija el Material Recibido:\\Codificado\Texto Libre")
           If RESP4% < 1 Or RESP4% > 2 Then
              Screen.MousePointer = 1
              GoTo 1100
           End If
           '
           If RESP4% = 1 Then
                GoTo 1113
              Else
42              Call RECEOCLI(NC%, FF%, NROBOL&, REMI$, NFAC$)
                Screen.MousePointer = 1
                GoTo 1100
           End If
        End If
        '
1113    AVENTA$ = "BOLREDET"
        'If ANUCA% = 1 Then AVENTA$ = "BOLRENUE"
        '\\\OT-05-0543-WAR-26/07/05///
        If ANUCA% = 1 Then AVENTA$ = "BOLRENU1"
        '\\\OT-05-0543-WAR-FIN///
        VTB% = VENTABU%(AVENTA$ + "/NC/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(RASOCLI$(NC1%)) + " - " + FECHAREC$)
        If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
            X$ = "": Y$ = "": YY$ = ""
            EBOL$ = "": PEDY$ = "": DESCRI$ = ""
            GoTo 1100
        End If
        '
        For YX& = 1 To ULTREG&("!BOLREDET")
          YY$ = REGLEIDO$("!BOLREDET", YX&)
          CI% = CVI(Mid$(YY$, 83, 2))
          CAN# = CVD(Mid$(YY$, 89, 8))
          If CI% > 0 Then
            If CI% <= NUMAS% Then
              If CAN# > 0.005 Then
                GoTo 1114
              End If
            End If
          End If
        Next YX&
        Call MENSERR(24, "Recepción Cancelada !!!\Faltan Cantidades o Códigos.")
        GoTo 1113
        '
1114    Call SETULTREG("!OBSERVOF", 0)
        Call CIERRARCH("!OBSERVOF")
        '
        VTB% = VENTABU%("OBSERVOF/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(RASOCLI$(NC1%)) + " - " + FECHAREC$)
        If VTB% < 0 Then
           If COMALTER%("Realmente Desea Cancelar ?\\Continuar\Cancelar") = 2 Then
              X$ = "": Y$ = "": YY$ = ""
              EBOL$ = "": PEDY$ = "": DESCRI$ = ""
              GoTo 1100
           End If
        End If
        '
11141   CONTA% = ULTREG&("!BOLREDET")
        '
1115    If CONTA% > 0 Then
          '
          For YX& = 1 To ULTREG&("!BOLREDET")
            YY$ = REGLEIDO$("!BOLREDET", YX&)
            CI% = CVI(Mid$(YY$, 83, 2))
            CAN# = CVD(Mid$(YY$, 89, 8))
            If CI% < 1 Or CI% > NUMAS% Then
               Call MENSERR(24, "Código no Válido.")
               GoTo 1113
            End If
            If ANUCA% < 1 Then
              If CAN# < 0.005 Then
                Call MENSERR(24, "Falta Cantidad.")
                GoTo 1113
              End If
            End If
          Next YX&
          '
          Call BLOQARCH("BOLRECEP")
          Call BLOQARCH("BOLREPEN")
          Call BLOQARCH("INDIBOL")
          Call BLOQARCH("IBOLPEN")
          Call BLOQARCH("OBSEBREC")
          '
1125      RFGH$ = RECFILT$("BOLRECEP", 1, MKS$(NROBOL&))
          If RFGH$ <> "" Then
            NROBOL& = NROBOL& + 1
            GoTo 1125
          End If   ' POSIBILIDAD NUMEROS REPETIDOS
          '
          Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
          '
          AINDI% = 0
          NORVENTA = ""
          For YX& = 1 To ULTREG&("!BOLREDET")
            YY$ = REGLEIDO$("!BOLREDET", YX&)
            CI% = CVI(Mid$(YY$, 83, 2))
            CAN# = CVD(Mid$(YY$, 89, 8))
            If CI% > 0 Then
              If CI% <= NUMAS% Then
                If CAN# > 0.005 Then
                  REGOC& = CVS(Mid$(YY$, 189, 4))
                  CARECI = CAN#
                  If REGOC& > 0 Then
                    If REGOC& <= ULTREG&("OCOMDETA") Then
                      ZZ$ = REGLEIDO$("OCOMDETA", REGOC&)
                      COEFUNI = CVS(Mid$(ZZ$, 63, 4))
                      If COEFUNI < 0.0001 Then COEFUNI = 1
                      CATORE = CVS(Mid$(ZZ$, 99, 4)) + CARECI / COEFUNI
                      Call REPLA(ZZ$, MKS$(CATORE), 99, 4)
                      CAPENDIEN = CVS(Mid$(ZZ$, 55, 4)) - CVS(Mid$(ZZ$, 99, 4))
                      NUORSE& = CVS(Mid$(ZZ$, 117, 4))
                      If CAPENDIEN >= 0.005 Then
                        If COMALTER%("Esta O/C por Código " + TRIM$(CODEXT$(CI%)) + "\Queda con un Saldo Pendiente de " + TRIM$(FORMATNUM$(CAPENDIEN, "F10.1")) + " " + TRIM$(UNIMED$(CI%)) + ".\\Mantener Abierta\Cancelar Pendiente") = 2 Then
                          Call REPLA(ZZ$, Chr$(8), 124, 1)
                        End If
                      End If
                    End If
                    Call GRAREG("OCOMDETA", ZZ$, REGOC&)
                  End If
                  '
                  NORVE& = CVS(Mid$(YY$, 151, 4))
                  If NORVE& > 0 Then
                    If NORVE& <> NORVEA& Then
                      NORVEA& = NORVE&
                      If NORVENTA <> "" Then NORVENTA = NORVENTA + " - "
                      NORVENTA = NORVENTA + TRIM$(Str$(NORVE&))
                    End If
                  End If
                  '
                  DESCRI$ = TRIM$(CODEXT$(CI%)) + " - " + TRIM$(DESCRIT$(CI%))
                  If AINDI% = 0 Then
                     X$ = REGBLAN$("INDIBOL")
                     Call REPLA(X$, MKS$(NROBOL&), 1, 4)
                     FEX = FF%
                     Call REPLA(X$, FECHATEX$(FEX), 5, 8)
                     Call REPLA(X$, REMI$, 15, 16)
                     Call REPLA(X$, ECOARCH$("PROVED1", MKI$(NC%)), 32, 29)
                     Call REGAPP("INDIBOL", X$)
                     If NUORSE& < 1 Then Call REGAPP("IBOLPEN", X$)
                     AINDI% = 1
                  End If
                  '
                  NBR$ = TRIM$(Str$(NROBOL&))
                  NORDE% = NORDE% + 1
                  NORDES$ = TRIM$(Str$(NORDE%))
                  IDENLO$ = "BR-" + NBR$ + "-" + NORDES$
                  '
                  Call REPLA(YY$, MKS$(NROBOL&), 1, 4)
                  If TIMATEI% <> 2 Then Call REPLA(YY$, DESCRI$, 5, 64)
                  If TRIM$(Mid$(YY$, 69, 12)) = "" Then
                    Call REPLA(YY$, IDENLO$, 69, 12)
                  End If
                  Call REPLA(YY$, MKI$(FF%), 81, 2)
                  DEPX% = DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(CI%)
                  Call REPLA(YY$, Chr$(DEPX%), 97, 1)
                  Call REPLA(YY$, MKI$(NC%), 105, 2)
                  Call REPLA(YY$, REMI$, 119, 16)
                  Call REPLA(YY$, MKS$(NUORSE&), 189, 4)
                  '
                  Call GRAREG("!BOLREDET", YY$, YX&)
                  Call REPLA(YY$, MKI$(0), 187, 2) ' blanquea fecha entrega solicitada
                  Call REPLA(YY$, MKD$(0), 163, 8)
                  Call REGAPP("BOLRECEP", YY$)
                  REG& = ULTREG&("BOLRECEP")
                  Call REPLA(YY$, MKS$(REG&), 189, 4)
                  If NUORSE& < 1 Then Call REGAPP("BOLREPEN", YY$)
                  '
                  If NUORSE& < 1 Then
                    If REGOC& > 0 Then
                        Call DESCACONSI(NC%, CI%, CAN#)
                      ElseIf REGOC& < 0 Then
                        Call DEVOLCONSI(Abs(REGOC&), NC%, CI%, CAN#)
                    End If
                  End If
                  '
                  If NUORSE& > 0 Then
                    If CAN# > 0.05 Then
                      GoSub BLANTEROP
                    End If
                  End If
                  '
                End If
              End If
            End If
          Next YX&
          '
          For YX& = 1 To ULTREG&("!OBSERVOF")
             YY$ = REGLEIDO$("!OBSERVOF", YX&)
             YY$ = MKS$(NROBOL&) + YY$
             Call REGAPP("OBSEBREC", YY$)
          Next YX&
          '
          Call CIERRARCH("BOLRECEP")
          Call CIERRARCH("BOLREPEN")
          Call CIERRARCH("INDIBOL")
          Call CIERRARCH("IBOLPEN")
          Call CIERRARCH("OBSEBREC")
          '
          Call FRESHECHO("IBOLPEN")
          Screen.MousePointer = 1
          '
          If TRIM$(FORIPRE$) <> "" Then
               '
               CANCELPRINT% = 0
               If CONTROL("BOLRECEP", "SUPRINT") = "SI" Then CANCELPRINT% = 1
               '
               TIPODOC$ = "Boleta de Recepción"
               DESTIDOC% = NC1%
               NUMEDOC$ = Str$(NROBOL&)
               FEX = FF%: FECHDOC$ = FECHATEX$(FEX)
               '
               CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = DEPOSAL$
               CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = REMI$
               CODPARAM$(3) = "N-AUTENT": DOCPARAM$(3) = NORVENTA
               CAPARDOC% = 3
               '
               CODTABU1$(1) = "COD-MAT"
               CODTABU1$(2) = "DES-MAT"
               CODTABU1$(3) = "COLOR"
               CODTABU1$(4) = "UNIMED"
               CODTABU1$(5) = "NRO-LOTE"
               CODTABU1$(6) = "CANTIDAD"
               CODTABU1$(7) = "N-OCOMPR"
               CODTABU1$(8) = "CAKILOS"
               CODTABU1$(9) = "F-PLANO"
               CODTABU1$(10) = "FECH-REC"
               '\\\OT-05-0543-WAR-26/07/05///
               CODTABU1$(11) = "PAR-MPRI"
               '\\\OT-05-0543-WAR-FIN///
               CACOLTAB1% = 11
               '
               CAITAB1% = 0
               For YX& = 1 To ULTREG&("!BOLREDET")
                  YY$ = REGLEIDO$("!BOLREDET", YX&)
                  CI% = CVI(Mid$(YY$, 83, 2))
                  CAN# = CVD(Mid$(YY$, 89, 8))
                  If CI% > 0 Then
                     If CI% <= NUMAS% Then
                        If CAN# > 0.005 Then
                           '
                           UNID$ = TRIM$(Mid$(YY$, 85, 4))
                           If UNID$ = "" Then UNID$ = UNIMED$(CI%)
                           CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
                           KILAJE$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 10, 3, 2)
                           FEX = CVI(Mid$(YY$, 155, 2))
                           FESOLEN$ = FECHATEX$(FEX)
                           FEPLA$ = FEPLANO$(CI%)
                           'NULOTE$ = TRIM$(Mid$(YY$, 107, 12)) '¿Esto estaBA BieN?
                           NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
                           If NULOTE$ = "" Then
                              NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
                           End If
                           NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
                           '\\\OT-05-0543-WAR-26/07/05///
                           LOTPROV$ = TRIM$(Mid$(YY$, 107, 12))
                           '\\\OT-05-0543-WAR-FIN///
                           '
                           CAITAB1% = CAITAB1% + 1
                           TETABU1$(1, CAITAB1%) = TRIM$(CODEXT$(CI%))
                           '
                           If FORMPLA$ = "RECEP-ESPUMA" Then
                                DESMAT$ = Mid$(YY$, 5, 64)
                              Else
                                DESMAT$ = DESCRIT$(CI%)
                           End If
                           TETABU1$(2, CAITAB1%) = DESMAT$
                           TETABU1$(3, CAITAB1%) = COLOR$(CI%)
                           TETABU1$(4, CAITAB1%) = UNID$
                           TETABU1$(5, CAITAB1%) = NULOTE$
                           TETABU1$(6, CAITAB1%) = CANS$
                           TETABU1$(7, CAITAB1%) = NUOCO$
                           TETABU1$(8, CAITAB1%) = KILAJE$
                           TETABU1$(9, CAITAB1%) = FEPLA$
                           TETABU1$(10, CAITAB1%) = FESOLEN$
                           '\\\OT-05-0543-WAR-26/07/05///
                           TETABU1$(11, CAITAB1%) = LOTPROV$
                           '\\\OT-05-0543-WAR-FIN///
                           '
                           If TICOMPRO$ = "FICHA" Then
                               Call PRINTDOC("FORMBREC")
                           End If
                           '
                        End If
                     End If
                  End If
               Next YX&
               '
               If TICOMPRO$ <> "FICHA" Then
                   Call PRINTDOC("FORMBREC")
               End If
               '
          End If
          '
1199    End If
        '
11999   Call CIERRARCH("*.*")
        '
        X$ = "": Y$ = "": YY$ = ""
        EBOL$ = "": PEDY$ = "": DESCRI$ = ""
        GoTo 1100
        '
BLANTEROP:
   '
   FINTER& = ULTREG&("OPERTER")
   For U& = FINTER& To 1 Step -1
     X0$ = REGLEIDO$("OPERTER", U&)
     If CVS(Mid$(X0$, 75, 4)) >= NUORSE& Then
       REGITER& = U&
       GoTo 21
     End If
   Next U&
   GoTo 29
   '
21 ' Marca como "Cumplida" Operación en OPERTER.DAT
   SALRE# = CVD(Mid$(X0$, 79, 8)) - CAN#
   If SALRE# < 0 Then SALRE# = 0
   '
   FEBA% = HOY(HOS$)
   HOBA% = HORANUM%(Time$)
   USUA% = USNBR% Mod 100
   DBAJA$ = MKI$(FEBA%) + MKI$(HOBA%) + Chr$(USUA%) + Chr$(1) + "B.Rec." + TRIM$(Str$(NROBOL&))
   '
   NUOROP% = Asc(Mid$(X0$, 67, 1))
   Call REPLA(X0$, MKD$(CAN#), 87, 8)
   Call REPLA(X0$, MKD$(SALRE#), 95, 8)
   Call REPLA(X0$, "3ROS  ", 151, 6)
   Call REPLA(X0$, MKI$(FEBA%), 157, 2)
   Call REPLA(X0$, MKI$(HORANUM%(Time$)), 159, 2)
   Call REPLA(X0$, MKS$(0), 161, 4)
   Call REPLA(X0$, String$(8, 0), 165, 8)
   Call REPLA(X0$, MKS$(0), 173, 4)
   Call REPLA(X0$, DBAJA$, 177, 22)
   Call GRAREG("OPERTER", X0$, REGITER&)
   '
' Si se recibio un parcial, genera registro en OPERTER.DAT
' por el saldo pendiente
   If SALRE# > 0.01 Then
      NUOROP% = 1 + Asc(Mid$(X0$, 67, 1))
      CIIK% = CVI(Mid$(X0$, 13, 2))
      NOPE% = CVI(Mid$(X0$, 15, 2))
      DALTA$ = Chr$(NUOROP%) + Mid$(DBAJA$, 2, 5)
      INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
      PROGBLA$ = Space$(6) + String$(26, 0)
      MINVARI = TIESTAN(CIIK%, NOPE%, SALRE#, NUOROP%)
      '
      Call REPLA(X0$, DALTA$, 67, 6)
      Call REPLA(X0$, MKD$(SALRE#), 79, 8)
      Call REPLA(X0$, MKD$(0), 87, 8)
      Call REPLA(X0$, MKD$(SALRE#), 95, 8)
      Call REPLA(X0$, String$(32, 0), 103, 32)
      Call REPLA(X0$, MKS$(MINVARI), 111, 4)
      Call REPLA(X0$, MKS$(MINVARI), 131, 4)
      Call REPLA(X0$, String$(16, 0), 135, 16)
      Call REPLA(X0$, INFOBLA$, 151, 48)
      Call REPLA(X0$, PROGBLA$, 225, 32)
      Call REGAPP("OPERTER", X0$)
      AGREGA% = 1
   End If
   '
29 Return
   '
   '
End Sub
'

Sub GENIBOLPE()
        '
        Screen.MousePointer = 11
        '
        K& = 0: NOOBA& = 0
        For U& = 1 To ULTREG&("BOLREPEN")
           X$ = REGLEIDO$("BOLREPEN", U&)
           APRO# = CVD(Mid$(X$, 163, 8))
           RECHA# = CVD(Mid$(X$, 171, 8))
           FALTA# = CVD(Mid$(X$, 179, 8))
           If Abs(APRO#) < 0.005 Then
             If Abs(RECHA#) < 0.005 Then
               If Abs(FALTA#) < 0.005 Then
                  NOOB& = CVS(Left$(X$, 4))
                  If NOOB& <> NOOBA& Then
                     NOOBA& = NOOB&
                     FEBO% = CVI(Mid$(X$, 81, 2))
                     RMTO$ = Mid$(X$, 119, 16)
                     NPBO% = CVI(Mid$(X$, 105, 2))
                     XX$ = REGBLAN$("IBOLPEN")
                     '
                     Call REPLA(XX$, MKS$(NOOB&), 1, 4)
                     FEX = FEBO%
                     Call REPLA(XX$, FECHATEX$(FEX), 5, 8)
                     Call REPLA(XX$, RMTO$, 15, 16)
                     Call REPLA(XX$, ECOARCH$("PROVED1", MKI$(NPBO%)), 32, 29)
                     K& = K& + 1
                     Call GRAREG("IBOLPEN", XX$, K&)
                  End If
               End If
             End If
           End If
        Next U&
        '
        Call SETULTREG("IBOLPEN", K&)
        Call CIERRARCH("IBOLPEN")
        Call FRESHECHO("IBOLPEN")
        '
        Screen.MousePointer = 1
        '
End Sub

Sub REPIBOL()
        '
        Dim REGBO&(2048)
        '
        TICOMPRO$ = TRIM$(CONTROL$("", "COMPBREC"))
        
1120    'Call CLEARVEN
        X$ = "": Y$ = "": Z$ = "": YY$ = ""
        EBOL$ = "": PEDY$ = "": DESCRI$ = ""
        NBR$ = OBJPLA$("SELBOLET", "")
        If NBR$ = "" Then
            Exit Sub
        End If
        '
        NROBOL& = CVS(Left$(NBR$, 4))
        If NROBOL& < 1 Then
            Exit Sub
        End If
        '
        Screen.MousePointer = 11
        FIN& = ULTREG&("BOLRECEP")
        Call SETULTREG("!BOLREDET", 0)
        Call CIERRARCH("!BOLREDET")
        Call SETULTREG("!OBSERVOF", 0)
        Call CIERRARCH("!OBSERVOF")
        '
        For U& = 1 To ULTREG&("BOLRECEP")
            X$ = REGLEIDO$("BOLRECEP", U&)
            If CVS(Left$(X$, 4)) = NROBOL& Then
                NC% = CVI(Mid$(X$, 7, 2))
                Call REGAPP("!BOLREDET", X$)
            End If
        Next U&
        '
        For U& = 1 To ULTREG&("OBSEBREC")
            X$ = REGLEIDO$("OBSEBREC", U&)
            If CVS(Left$(X$, 4)) = NROBOL& Then
                Call REGAPP("!OBSERVOF", Mid$(X$, 5, 72))
            End If
        Next U&
        '
        Screen.MousePointer = 1
        '
        Call CIERRARCH("!BOLREDET")
        Call CIERRARCH("!OBSERVOF")
        '
        If ULTREG&("!BOLREDET") < 1 Then
            Call COMUNI("Boleta de Recepcion Nro." + Str$(NROBOL&) + "\no Figura en la Base de Datos.")
            GoTo 1120
        End If
        '
        X$ = REGLEIDO$("TATIREC", 1)
        FORMPLA$ = TRIM$(UCase$(Mid$(X$, 33, 16)))
        FORMTAB$ = TRIM$(UCase$(Mid$(X$, 49, 16)))
        FORIPRE$ = TRIM$(UCase$(Mid$(X$, 65, 16)))
        '
        If TRIM$(FORIPRE$) = "" Then
            Call COMUNI("No se Definió Formulario de Impresión de B.R.")
            GoTo 1120
        End If
        '
        YY$ = REGLEIDO$("!BOLREDET", 1)
        NC% = CVI(Mid$(YY$, 105, 2))
        NC1% = (-1) * NC%
        FF% = CVI(Mid$(YY$, 81, 2))
        DP1% = Asc(Mid$(YY$, 97, 1))
        DEPOSAL$ = TRIM$(ECOARCH$("ECODEP", Chr$(DP1%)))
        REMI$ = Mid$(YY$, 119, 16)
        '
        If TRIM$(FORIPRE$) <> "" Then
               '
               TIPODOC$ = "Boleta de Recepción"
               DESTIDOC% = NC1%
               NUMEDOC$ = Str$(NROBOL&)
               FEX = FF%: FECHDOC$ = FECHATEX$(FEX)
               '
               CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = DEPOSAL$
               CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = REMI$
               CAPARDOC% = 2
               '
               CODTABU1$(1) = "COD-MAT"
               CODTABU1$(2) = "DES-MAT"
               CODTABU1$(3) = "COLOR"
               CODTABU1$(4) = "UNIMED"
               CODTABU1$(5) = "NRO-LOTE"
               CODTABU1$(6) = "CANTIDAD"
               CODTABU1$(7) = "N-OCOMPR"
               CODTABU1$(8) = "CAKILOS"
               CODTABU1$(9) = "F-PLANO"
               CACOLTAB1% = 9
               '
               CAITAB1% = 0
               For YX& = 1 To ULTREG&("!BOLREDET")
                  YY$ = REGLEIDO$("!BOLREDET", YX&)
                  CI% = CVI(Mid$(YY$, 83, 2))
                  CAN# = CVD(Mid$(YY$, 89, 8))
                  If CI% > 0 Then
                     If CI% <= NUMAS% Then
                        If CAN# > 0.005 Then
                           '
                           UNID$ = TRIM$(Mid$(YY$, 85, 4))
                           If UNID$ = "" Then UNID$ = UNIMED$(CI%)
                           CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
                           KILAJE$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 10, 3, 2)
                           FEPLA$ = FEPLANO$(CI%)
                           NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
                           If NULOTE$ = "" Then
                              NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
                           End If
                           NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
                           '
                           CAITAB1% = CAITAB1% + 1
                           TETABU1$(1, CAITAB1%) = TRIM$(CODEXT$(CI%))
                           '
                           If FORMPLA$ = "RECEP-ESPUMA" Then
                                DESMAT$ = Mid$(YY$, 5, 64)
                              Else
                                DESMAT$ = DESCRIT$(CI%)
                           End If
                           TETABU1$(2, CAITAB1%) = DESMAT$
                           TETABU1$(3, CAITAB1%) = COLOR$(CI%)
                           TETABU1$(4, CAITAB1%) = UNID$
                           TETABU1$(5, CAITAB1%) = NULOTE$
                           TETABU1$(6, CAITAB1%) = CANS$
                           TETABU1$(7, CAITAB1%) = NUOCO$
                           TETABU1$(8, CAITAB1%) = KILAJE$
                           TETABU1$(9, CAITAB1%) = FEPLA$
                           '
                           If TICOMPRO$ = "FICHA" Then
                               Call PRINTDOC("FORMBREC")
                           End If
                           '
                        End If
                     End If
                  End If
               Next YX&
               '
               If TICOMPRO$ <> "FICHA" Then
                   Call PRINTDOC("FORMBREC")
               End If
               '
1199    End If
        '
        Call CIERRARCH("*.*")
        '
        X$ = "": Y$ = "": YY$ = ""
        EBOL$ = "": PEDY$ = "": DESCRI$ = ""
        GoTo 1120
        '
End Sub
'

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Sub SHOWBORE()
    '
3   Call SELECHO("INDIBOL")
    NPX& = Val(VALACT1$("INDIBOL"))
    If NPX& > 0 Then
      GoTo 4
    End If
    Exit Sub
    '
4   TIDOCUM$ = "Boleta de Recepción"
    NUDOCU$ = TRIM$(Str$(NPX&))
    NUDOCU$ = " " + NUDOCU$ + " -"
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        If InStr(XX$, NUDOCU$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    Exit Sub
    '
    '19/03/07 (OT-07-0113)
5   Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
    '19/03/07 (FIN)
'5  For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    '
'    REDOCU$ = REGACT$("PDOCLST")
'    URE& = ULTREG&("PDOCSPL")
'    LI& = 0: LS& = URE&
'6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'    GoTo 8
'    '
'7   Screen.MousePointer = 1
'    Call MENSERR(24, "Documento no Registrado")
'    Exit Sub
'    '
'8   TPSP$ = ""
'    Screen.MousePointer = 11
'    While L& > 1
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then GoTo 9
'      L& = L& - 1
'    Wend
'    '
'9   CAPAGINAS% = 0
'    For KKL& = L& To URE&
'      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'      End If
'    Next KKL&
'    '
'95  SPOOLSTRING$ = TPSP$
'    FIN& = Len(TPSP$)
'    For KKI% = 0 To 5
'      EDITFORM.Picture1(KKI%).Cls
'      EDITFORM.Picture1(KKI%).Top = 0
'      EDITFORM.Picture1(KKI%).Refresh
'      EDITFORM.Picture1(KKI%).Height = 7070
'      EDITFORM.Picture1(KKI%).Width = 11700
'      EDITFORM.Picture1(KKI%).DrawWidth = 3
'      EDITFORM.Picture1(KKI%).DrawStyle = 0
'      EDITFORM.Picture1(KKI%).FillStyle = 1
'    Next KKI%
'    ALTUPAGINA = 7070
'    TOPEPAGINA = 0
'    PAGINACTIVA% = 0
'    CAPAGINAS% = 0
'    '
'10  PPXX% = InStr(TPSP$, Chr$(255))
'    If PPXX% > 0 Then
'      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'      TPSP$ = Mid$(TPSP$, PPXX% + 1)
'      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
'      '
'      If Len(COPRIAT$) > 0 Then
'        COPRI% = Asc(Left$(COPRIAT$, 1))
'        ATRIP$ = Mid$(COPRIAT$, 2)
'        Call SHOWCOMMAND(COPRI%, ATRIP$)
'      End If
'      GoTo 10
'      '
'    End If
'    '
'    UI& = UI& + 1
'    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'    If InStr(XX$, TIDOCUM$) > 4 Then
'      If InStr(XX$, NUDOCU$) > 4 Then
'        DOCUNU& = CVS(Left$(XX$, 4))
'        GoTo 7
'      End If
'    End If
'    '
'    Screen.MousePointer = 1
'    PAGINACTIVA% = 0
'    For KKI% = 0 To 5
'       EDITFORM.Picture1(KKI%).Visible = False
'    Next KKI%
'    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
'    EDITFORM.Picture1(PAGINACTIVA%).Refresh
'    EDITFORM.VScroll1.Min = 0
'    EDITFORM.VScroll1.Value = 0
'    EDITFORM.VScroll1.Max = 1
'    If ALTUPAGINA > EDITFORM.Frame1.Height Then
'       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'    End If
'    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
'    EDITFORM.HScroll1.Min = 0
'    EDITFORM.HScroll1.Value = 0
'    EDITFORM.HScroll1.Max = 1
'    If ANCHPAGINA > EDITFORM.Frame1.Width Then
'      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
'    End If
'    EDITFORM.Command3.Enabled = False
'    EDITFORM.Command5.Enabled = False
'    EDITFORM.Show 1
    '
    GoTo 3
    '
End Sub
'
Sub RECEOCLI(NC%, FF%, NROBOL&, REMI$, NFAC$)
    '
    Call SELECHO("!TXOCOLIB")
    NPX& = Val(VALACT1$("!TXOCOLIB"))
    If NPX& < 1 Then Exit Sub
    '
    ODX$ = FILTERGETRECORD$("OCOMDETA", 1, MKS$(NPX&))
    CAPEN = CVS(Mid$(ODX$, 55, 4)) - CVS(Mid$(ODX$, 99, 4))
    CIXI% = CVI(Mid$(ODX$, 9, 2))
    MONEXO$ = "$": If Asc(Mid$(ODX$, 122, 1)) = 2 Then MONEXO$ = "U$S"
    PRUNX = CVD(Mid$(ODX$, 79, 8))
    '
4   With PRESOCLI
       .Label19 = MONEXO$
       .Text1 = FORMATNUM$(PRUNX, "F12.2")
       .Label4 = TRIM$(Str$(NPX&))
       .Label12 = TRIM$(Str$(NROBOL&))
       .Label2 = FECHATEX$(FF%)
       .Label5 = TRIM$(Str$(NC%))
       .Label11 = RASOCLI$((-1) * NC%)
       .Label16 = CODEXT$(CIXI%)
       .Label15 = VALACT2$("!TXOCOLIB")
       .Label7 = TRIM$(FORMATNUM$(CAPEN, "F8.1"))
       .Label22 = REMI$
       .Label18 = NFAC$
       .Option1.Value = True
       .Show 1
    End With
    '
End Sub

VERSION 5.00
Begin VB.Form FOREMAQ04 
   BackColor       =   &H00D0F0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe de Recepción de Materiales"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   Icon            =   "FOREMAQ04.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command12 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   940
      Left            =   105
      Picture         =   "FOREMAQ04.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Inspección y Aprobación de Recepciones"
      Top             =   8610
      Width           =   750
   End
   Begin VB.CommandButton Command22 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aduana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   770
      Left            =   1575
      Picture         =   "FOREMAQ04.frx":17A8
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Registrar Datos Respacho Aduana"
      Top             =   8610
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Importación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   1960
      Left            =   4410
      TabIndex        =   46
      Top             =   8715
      Width           =   3225
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1050
         MaxLength       =   16
         TabIndex        =   52
         Top             =   920
         Width           =   1995
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   525
         MaxLength       =   24
         TabIndex        =   51
         Top             =   1530
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Temporaria"
         Enabled         =   0   'False
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
         Left            =   210
         TabIndex        =   50
         Top             =   420
         Width           =   1275
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Definitiva"
         Enabled         =   0   'False
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
         Left            =   210
         TabIndex        =   49
         Top             =   630
         Width           =   1170
      End
      Begin VB.CommandButton Command7 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2840
         TabIndex        =   48
         Top             =   525
         Width           =   195
      End
      Begin VB.OptionButton Option0 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Ninguna"
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
         Left            =   1890
         TabIndex        =   47
         Top             =   105
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   10500
         Y1              =   1940
         Y2              =   1940
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Aduana:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   330
         Left            =   -945
         TabIndex        =   56
         Top             =   945
         Width           =   1905
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Despacho Importación:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   330
         Left            =   105
         TabIndex        =   55
         Top             =   1260
         Width           =   2115
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   330
         Left            =   1575
         TabIndex        =   54
         Top             =   300
         Width           =   750
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1890
         TabIndex        =   53
         Top             =   540
         Width           =   960
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Devolución Material en Consignación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   1960
      Left            =   315
      TabIndex        =   37
      Top             =   8925
      Width           =   10580
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   10500
         TabIndex        =   44
         Top             =   315
         Width           =   10565
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Código          Descripción                             U/M      Pend.    Devu.   O/C"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   225
            Left            =   30
            TabIndex        =   45
            Top             =   105
            Width           =   10515
         End
      End
      Begin VB.ListBox List2 
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
         Height          =   1230
         IntegralHeight  =   0   'False
         Left            =   0
         MultiSelect     =   1  'Simple
         TabIndex        =   43
         Top             =   710
         Width           =   10565
      End
      Begin VB.Frame Frame10 
         Caption         =   "Frame10"
         Height          =   1620
         Left            =   1890
         TabIndex        =   42
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame11 
         Caption         =   "Frame11"
         Height          =   1620
         Left            =   6720
         TabIndex        =   41
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame12 
         Caption         =   "Frame12"
         Height          =   1620
         Left            =   7455
         TabIndex        =   40
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame13 
         Caption         =   "Frame13"
         Height          =   1620
         Left            =   8505
         TabIndex        =   39
         Top             =   315
         Width           =   15
      End
      Begin VB.Frame Frame14 
         Caption         =   "Frame14"
         Height          =   1620
         Left            =   9555
         TabIndex        =   38
         Top             =   315
         Width           =   15
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   10500
         Y1              =   1940
         Y2              =   1940
      End
      Begin VB.Line Line2 
         X1              =   10565
         X2              =   10565
         Y1              =   315
         Y2              =   1940
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0F0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8590
      Left            =   0
      TabIndex        =   9
      Top             =   80
      Width           =   12000
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00004000&
         Height          =   9000
         Left            =   11025
         ScaleHeight     =   8940
         ScaleWidth      =   1005
         TabIndex        =   58
         Top             =   0
         Width           =   1065
         Begin VB.CommandButton Command13 
            BackColor       =   &H00E0E0E0&
            Caption         =   "OC"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FOREMAQ04.frx":1AB2
            TabIndex        =   69
            ToolTipText     =   "Visualización de Ordenes de Compra Pendientes"
            Top             =   3850
            Width           =   650
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "List"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FOREMAQ04.frx":1DBC
            Style           =   1  'Graphical
            TabIndex        =   68
            ToolTipText     =   "Recepciones por Rango de Fecha"
            Top             =   5250
            Width           =   650
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ver"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FOREMAQ04.frx":20C6
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
            Top             =   4560
            Width           =   650
         End
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   105
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   65
            Top             =   5835
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   66
               Top             =   105
               Width           =   12000
            End
         End
         Begin VB.CommandButton Command14 
            BackColor       =   &H00E0E0E0&
            Caption         =   "O.K."
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1110
            Left            =   120
            Picture         =   "FOREMAQ04.frx":23D0
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "Aprueba, Graba e Imprime Orden de Compra"
            Top             =   6600
            Width           =   650
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FOREMAQ04.frx":26DA
            Style           =   1  'Graphical
            TabIndex        =   63
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   100
            Width           =   650
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Setup"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FOREMAQ04.frx":2824
            Style           =   1  'Graphical
            TabIndex        =   62
            ToolTipText     =   "Configuración de Funcionamiento de Compras"
            Top             =   3045
            Width           =   650
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Rep."
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FOREMAQ04.frx":2C66
            Style           =   1  'Graphical
            TabIndex        =   61
            ToolTipText     =   "Acceso a Maestro de Repuestos"
            Top             =   2240
            Width           =   650
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Prove."
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   105
            Picture         =   "FOREMAQ04.frx":2F70
            Style           =   1  'Graphical
            TabIndex        =   60
            ToolTipText     =   "Acceso a Base de Datos de Proveedores"
            Top             =   1550
            Width           =   650
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   640
            Left            =   105
            Picture         =   "FOREMAQ04.frx":327A
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   790
            Width           =   650
         End
         Begin VB.Image Image1 
            Height          =   390
            Left            =   -315
            Picture         =   "FOREMAQ04.frx":3584
            Top             =   6075
            Width           =   1515
         End
      End
      Begin VB.Frame Frame20 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Depósito Entrada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   705
         Left            =   6660
         TabIndex        =   32
         Top             =   5850
         Width           =   4080
         Begin VB.CommandButton Command15 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   420
            TabIndex        =   34
            Top             =   315
            Width           =   150
         End
         Begin VB.Label Label16 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   105
            TabIndex        =   35
            Top             =   315
            Width           =   330
         End
         Begin VB.Line Line9 
            X1              =   0
            X2              =   10500
            Y1              =   1940
            Y2              =   1940
         End
         Begin VB.Label Label15 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   630
            TabIndex        =   33
            Top             =   315
            Width           =   3270
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Inspección"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   1180
         Left            =   6660
         TabIndex        =   22
         Top             =   6600
         Width           =   4080
         Begin VB.CheckBox Check1 
            BackColor       =   &H00D0F0F0&
            Caption         =   "Ingresa sin Control"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   315
            TabIndex        =   24
            Top             =   210
            Value           =   2  'Grayed
            Width           =   1935
         End
         Begin VB.Label Label11 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   735
            TabIndex        =   31
            Top             =   735
            Width           =   3165
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Resp:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   330
            Left            =   105
            TabIndex        =   23
            Top             =   750
            Width           =   540
         End
         Begin VB.Line Line8 
            X1              =   0
            X2              =   10500
            Y1              =   1940
            Y2              =   1940
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   1960
         Left            =   160
         TabIndex        =   21
         Top             =   5820
         Width           =   6270
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1590
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   350
            Width           =   6240
         End
         Begin VB.Line Line6 
            X1              =   0
            X2              =   10500
            Y1              =   1940
            Y2              =   1940
         End
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8760
         MaxLength       =   16
         TabIndex        =   19
         Text            =   " "
         Top             =   770
         Width           =   2010
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Detalle de Materiales a Ingresar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   4400
         Left            =   160
         TabIndex        =   15
         Top             =   1260
         Width           =   10580
         Begin VB.Frame Frame19 
            Caption         =   "Frame19"
            Height          =   4530
            Left            =   10290
            TabIndex        =   30
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame18 
            Caption         =   "Frame18"
            Height          =   4530
            Left            =   6300
            TabIndex        =   29
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame17 
            Caption         =   "Frame17"
            Height          =   4530
            Left            =   8780
            TabIndex        =   28
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame16 
            Caption         =   "Frame16"
            Height          =   4530
            Left            =   7245
            TabIndex        =   27
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame7 
            Caption         =   "Frame7"
            Height          =   4530
            Left            =   1890
            TabIndex        =   26
            Top             =   315
            Width           =   15
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00E0E0E0&
            Height          =   435
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   10500
            TabIndex        =   17
            Top             =   315
            Width           =   10565
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "Código          Descripción                          Nº O/C       Saldo     Recibido"
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404000&
               Height          =   225
               Left            =   30
               TabIndex        =   18
               ToolTipText     =   "Doble Click para Cargar Detalle"
               Top             =   105
               Width           =   10515
            End
         End
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3660
            Left            =   0
            TabIndex        =   16
            Top             =   735
            Width           =   10560
         End
         Begin VB.Line Line4 
            X1              =   10560
            X2              =   10560
            Y1              =   315
            Y2              =   4880
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "!"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1680
         TabIndex        =   14
         ToolTipText     =   "Lista de Proveedores con Entregas Programadas"
         Top             =   9520
         Width           =   195
      End
      Begin VB.TextBox Text11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9930
         TabIndex        =   8
         Text            =   " "
         Top             =   105
         Width           =   825
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9615
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   440
         Width           =   960
      End
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10560
         TabIndex        =   7
         Top             =   440
         Width           =   195
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3150
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   770
         Width           =   3120
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2100
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   " "
         Top             =   770
         Width           =   960
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2100
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   440
         Width           =   4170
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2100
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   105
         Width           =   4170
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   770
         Width           =   170
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   690
         MaxLength       =   5
         TabIndex        =   0
         Text            =   " "
         Top             =   780
         Width           =   750
      End
      Begin VB.Line Line7 
         X1              =   195
         X2              =   10745
         Y1              =   5650
         Y2              =   5650
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remito Proveedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   540
         Left            =   7140
         TabIndex        =   20
         Top             =   630
         Width           =   1485
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   750
         Left            =   105
         TabIndex        =   13
         Top             =   60
         Width           =   1695
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   630
         Shape           =   4  'Rounded Rectangle
         Top             =   255
         Width           =   540
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000C000&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   390
         Left            =   525
         Shape           =   4  'Rounded Rectangle
         Top             =   180
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   420
         Shape           =   4  'Rounded Rectangle
         Top             =   105
         Width           =   960
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   330
         Left            =   8670
         TabIndex        =   12
         Top             =   165
         Width           =   1170
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   330
         Left            =   8820
         TabIndex        =   11
         Top             =   480
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Prov:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   330
         Left            =   105
         TabIndex        =   10
         Top             =   840
         Width           =   960
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   11865
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu MnuArch 
      Caption         =   "Archivos"
      Begin VB.Menu MnuNueBr 
         Caption         =   "Nueva B/R"
      End
      Begin VB.Menu php0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEnd 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu MnuEdi 
      Caption         =   "Edición"
   End
   Begin VB.Menu MnuConLis 
      Caption         =   "Consultas y Listados"
   End
   Begin VB.Menu MnuStp 
      Caption         =   "Configuración"
   End
   Begin VB.Menu MnuAccDir 
      Caption         =   "Accesos Directos"
      Begin VB.Menu MnuMasPro 
         Caption         =   "Maestro de Proveedores"
      End
      Begin VB.Menu MnuMasRep 
         Caption         =   "Maestro de Repuestos"
      End
      Begin VB.Menu php1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuOthAcc 
         Caption         =   "Otros Accesos Directos"
      End
   End
   Begin VB.Menu MnuAyu 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "FOREMAQ04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHIPRE$
Dim MUECOPEN%
Dim EMICONSI%
Dim LETRANO%
'
Private Sub Command1_Click()
   Call SELECHO("ACREDOR")
   NCLIE% = Val(TRIM$(VALACT1$("ACREDOR")))
   If NCLIE% > 0 Then
     If REGICLI&((-1) * NCLIE%) > 0 Then
       Call BLANFORMU
       Text1.TEXT = TRIM$(Str$(NCLIE%))
       Call CARDACLI
       Call CARCOPEN
       Call CARLISRE ''**''
'       If ULTREG&("!BOLREDET") < 1 Then
'         Call CARDETCOM
'       End If
     End If
   End If
End Sub

Private Sub Command10_Click()
   Call SELECHO("PROVSEL")
   NCLIE% = Val(TRIM$(VALACT1$("PROVSEL")))
   If NCLIE% > 0 Then
     If REGICLI&((-1) * NCLIE%) > 0 Then
       Call BLANFORMU
       Text1.TEXT = TRIM$(Str$(NCLIE%))
       Call CARDACLI
       Call CARCOPEN
       If ULTREG&("!BOLREDET") < 1 Then
         Call CARDETCOM
       End If
     End If
   End If
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
'   Call CIERRARCH("*.*")
'   ACONEC$ = "AMAPRO"
'   Call CONECRUN(ACONEC$, "Padron Maestro de Proveedores")
   ACRE_MAQ04.Show 1
   Command11.Enabled = True
End Sub
Private Sub Command13_Click()
   Command13.Enabled = False
   Call BLANARCH("!INDIOCOM")
    'Cargo la lista con las O/C Pendientes
    Call ABRECOMPRAS
    SQLX$ = "Select * from OCOMDETA"
    SQLX$ = SQLX$ + " Where Stat_Ocom < 8 "
    SQLX$ = SQLX$ + " AND Stat_Ocom <> -1 "
    SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
    SQLX$ = SQLX$ + "ORDER BY Nume_Ocom ASC"
    '
    Set VRIMPTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    With VRIMPTEMP
        On Error Resume Next ' si hay algun error msgbox err
        .MoveFirst
        If Err Then
          On Error GoTo 0
          Call MENSERR(24, "No Aparecen Ordenes de Compra Emitidas")
          GoTo 99
        End If
        On Error GoTo 0
        JJ& = 0
        Do While (.EOF = False)
            XXY$ = REGBLAN$("INDIOCOM")
            NOC1 = !Nume_Ocom
            If NOC1 <> NOCA Then
              With OcomEnca
                 .FindFirst " NUME_OCOM = " & NOC1 & ""
                 Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4)
                 Call REPLA(XXY$, !Refe_Ocom, 5, 44)
              End With
              JJ& = JJ& + 1
              Call GRAREG("!INDIOCOM", XXY$, JJ&)
              NOCA = NOC1
            End If
            .MoveNext
        Loop
    End With
    Set VRIMPTEMP = Nothing
    Call SETULTREG("!INDIOCOM", JJ&)
    Call CIERRARCH("!INDIOCOM")
    '
    Call FRESHECHO("!INDIOCOM")
3   Call SELECHO("!INDIOCOM/Indice General de Ordenes de Compra Pendientes ")
    NPX& = Val(VALACT1$("!INDIOCOM")) ' ACA LEO DEL INDIOCOM
    If NPX& < 1 Then GoTo 99
    '
    TIDOCUM$ = "Nota de Pedido"
    NUDOCU$ = TRIM$(Str$(NPX&))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU$ = " " + NUDOCU$ + " "
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    GoTo 99
    '
5   For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    GoTo 99
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    FIN& = Len(TPSP$)
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7070
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7070
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    UI& = UI& + 1
    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
    If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.HScroll1.Min = 0
    EDITFORM.HScroll1.Value = 0
    EDITFORM.HScroll1.Max = 1
    If ANCHPAGINA > EDITFORM.Frame1.Width Then
      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command1.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 3
    '
99  Command13.Enabled = True
   '
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   If VALIDATOS% > 0 Then
    OKK% = 0
    Call BOLERECEP
    '
    Call SETULTREG("!BOLREDET", 0)
    Call CIERRARCH("*.*")
    List1.Clear
    Text1.TEXT = ""
    Text2.TEXT = ""
    Text3.TEXT = ""
    Text4.TEXT = ""
    Text5.TEXT = ""
    Text6.TEXT = ""
    Text13.TEXT = ""
    Text11.TEXT = TRIM$(Str$(PRONUBOL&))
    Label16 = TRIM$(CONTROL$("", "DEPOENT"))
    Call CARDACLI
    On Error Resume Next
    Text1.SetFocus
    On Error GoTo 0
    '
   End If
90 Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
   Call SELECHO("ECODEP")
   VDEF$ = TRIM$(VALACT1$("ECODEP"))
   If VDEF$ <> "" Then
     Label16 = VDEF$
   End If
End Sub

Private Sub Command16_Click()
   Command16.Enabled = False
   Command14.Enabled = False
   'If VALIDATOS% > 0 Then
     Call PRIBOLRE(0)
   'End If
   Command14.Enabled = True
   Command16.Enabled = True
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Private Sub Command2_LostFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub


Private Sub Command22_Click()
   Command22.Enabled = False
   '
   If DERACCE%("CADUANA", "Registrar Despacho de Importación") >= 2 Then
     ASELE$ = "IREPRIM"
3    Call SELECHO(ASELE$ + "/Indice General de Informes de Recepción")
     NBOL& = Val(VALACT1$(ASELE$))
     If NBOL& > 0 Then
       RFF$ = RECFILT$("REMAPRI", 1, MKS$(NBOL&))
       If Len(RFF$) < 4 Then
         Call MENSERR(24, "Documento no Registrado")
         GoTo 3
       End If
       '
       REBO& = CVS(Left$(RFF$, 4))
       XX$ = REGLEIDO$("REMAPRI", REBO&)
       '
       ' este bloque es para el caso de que al registrar el IR el proveedor
       ' estaba marcado como local y se lo paso a exterior
       '
       If Mid$(XX$, 193, 1) <> "X" Then
         NPRI% = CVI(Mid$(XX$, 105, 2))
         PVPRX% = PIVACLI%(-1 * NPRI%)
         If PVPRX% = 5 Or PVPRX% = 7 Then
           Call REPLA(XX$, "X", 193, 1)
           Call GRAREG("REMAPRI", XX$, REBO&)
         End If
       End If
       '
       If Mid$(XX$, 193, 1) <> "X" Then
         Call MENSERR(24, "Esta Recepción Corresponde\a Material de Origen LOCAL.")
         GoTo 3
       End If
       '
       VDEF$ = XX$
10     OBX$ = OBJPLA$("CARDESPADUA", VDEF$)
       If OBX$ <> "" Then
         FADUA% = CVI(Mid$(OBX$, 195, 2))
         ADUA$ = TRIM$(Mid$(OBX$, 197, 16))
         NDESP$ = TRIM$(Mid$(OBX$, 213, 24))
         NREMI$ = TRIM$(Mid$(OBX$, 119, 16))
         CARIM$ = UCase$(TRIM$(Mid$(OBX$, 194, 1)))
         If FADUA% < 1 Or ADUA$ = "" Or NDESP$ = "" Or NREMI$ = "" Then
           Call MENSERR(24, "Datos Incompletos")
           VDEF$ = OBX$
           GoTo 10
         End If
         If CARIM$ <> "T" And CARIM$ <> "C" Then
           Call MENSERR(24, "Debe ser T-emporaria o C-onsumo")
           VDEF$ = OBX$
           GoTo 10
         End If
         If FADUA% > CVI(Mid$(XX$, 81, 2)) Then
           Call MENSERR(24, "Fecha Aduana no Válida")
           VDEF$ = OBX$
           GoTo 10
         End If
         NFACT$ = TRIM$(Mid$(OBX$, 135, 16))
         '
         Call BLOQARCH("REMAPRI")
         '
         Screen.MousePointer = 1
         RFF$ = RECFILT$("REMAPRI", 1, MKS$(NBOL&))
         For U& = 1 To Len(RFF$) Step 4
           REBO& = CVS(Mid$(RFF$, U&, 4))
           XX$ = REGLEIDO$("REMAPRI", REBO&)
           Call REPLA(XX$, NREMI$, 119, 16)
           Call REPLA(XX$, NFACT$, 135, 16)
           Call REPLA(XX$, CARIM$, 194, 1)
           Call REPLA(XX$, MKI$(FADUA%), 195, 2)
           Call REPLA(XX$, ADUA$, 197, 16)
           Call REPLA(XX$, NDESP$, 213, 24)
           Call GRAREG("REMAPRI", XX$, REBO&)
         Next U&
         Call CIERRARCH("REMAPRI")
         '
         Screen.MousePointer = 1
         '
       End If
       GoTo 3
       '
     End If
   End If
   '
   Command22.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("SIMPTEMP")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    '
    Command4.Enabled = False
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
    CI1% = CVI(Left$(RACOD$, 2)): COD1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): COD2$ = CODEXT$(CI2%)
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
    If COD1$ = "" Or COD2$ = "" Or COD2$ < COD1$ Or FF2% < FF1% Then
        VDEF$ = RACOD$
        GoTo 22
    End If
    '
    RANGO$ = TRIM$(COD1$) + " - " + TRIM$(COD2$)
    RANGO$ = RANGO$ + " / " + TRIM$(TIMAT$)
    RANGO$ = RANGO$ + " / " + FLI1$ + " - " + FLI2$
    '
    Screen.MousePointer = 11
    '
    Call COPYSTRU("BOLRECEP", "BOLSELEC")
    J& = 0
    '
    Call ABRESTOCKS
    FF1D = CDbl(CVDAT(FF1%))
    FF2H = CDbl(CVDAT(FF2%))
    '
    SQLX$ = " SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + " WHERE cdbl(FechRecep) >= " & FF1D & " "
    SQLX$ = SQLX$ + " AND cdbl(FechRecep) <= " & FF2H & " "
    SQLX$ = SQLX$ + " ORDER BY NumeBoRe "
    '
    Set BRECEPRATMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With BRECEPRATMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Boletas de Recepcion\que Cumplan Condición de Filtro.")
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
        CIXI% = !Cod_Inte
        CODD$ = CODEXT$(CIXI%)
        If CODD$ >= COD1$ Then
          If CODD$ <= COD2$ Then
            If TMT% = 0 Or TIMATE%(CIXI%) = TMT% Then
                Y$ = REGBLAN("BOLSELEC")
                Call REPLA(Y$, MKS$(!NumeBore), 1, 4)
                Call REPLA(Y$, !DESCRIP, 5, 64)
                Call REPLA(Y$, !IdenLote, 69, 12)
                Call REPLA(Y$, MKI$(CVINT(!FechRecep)), 81, 2)
                Call REPLA(Y$, MKI$(!Cod_Inte), 83, 2)
                Call REPLA(Y$, !Unidad, 85, 4)
                Call REPLA(Y$, MKD$(!CantNom), 89, 8)
                Call REPLA(Y$, Chr$(!DepoEnt), 97, 1)
                Call REPLA(Y$, MKI$(!NUMPROV), 105, 2)
                Call REPLA(Y$, !LoteCol_Prov, 107, 12)
                Call REPLA(Y$, !NumRem_Prov, 119, 16)
                  NCX% = !NUMPROV
                Call REPLA(Y$, CODICLI$((-1) * NCX%), 135, 16)
                Call REPLA(Y$, MKS$(!NumOC), 151, 4)
                Call REPLA(Y$, !MoneVaUni, 155, 4)
                Call REPLA(Y$, MKS$(!Valo_Unit), 159, 4)
                Call REPLA(Y$, MKD$(!CanToApro), 163, 8)
                Call REPLA(Y$, MKD$(!CanToRech), 171, 8)
                Call REPLA(Y$, MKD$(!CanToFalta), 179, 8)
                Call REPLA(Y$, MKI$(CVINT(!FechLibe)), 187, 2)
                If IsNull(!Observa) Then
                    USERAPRO$ = " "
                Else
                    USERAPRO$ = !Observa
                End If
                Call REPLA(Y$, USERAPRO$, 193, 48) 'antes estaba el campo observaciones
                '
                JJ& = JJ& + 1
                Call GRAREG("BOLSELEC", Y$, JJ&)
            End If
          End If
        End If
      .MoveNext
      Loop
    End With
    Call SETULTREG("BOLSELEC", JJ&)
    Call CIERRARCH("BOLSELEC")
    Call CIERRARCH("*.*")
    '
    Call HEADERS("BOLSELEC", "")
    Call HEADERS("BOLSELEC", "Rango: " + RANGO$)
    Call CONECRUN("*BOLSELEC", "Recepciones por Rango de Fechas")
    '
99 Screen.MousePointer = 1
   Command4.Enabled = True
   '
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
'    Call CONECRUN("FLSETUP/OPIRMP00", "Configuración de Funcionamiento")
'    Call CIERRARCH("*.*")
'    Call FINAL("")
    OPREMAQ04.Show 1
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    '
    Command6.Enabled = False
3  Call SELECHO("INDIBOL")
   NPX& = Val(VALACT1$("INDIBOL"))
   If NPX& > 0 Then
     Call SHOWBORE(NPX&)
     GoTo 3
   End If
   '
99 Command6.Enabled = True
End Sub
'
'
Sub SHOWBORE(NPX&)
    '
    TIDOCUM$ = "Boleta de Recepción"
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
5  For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    Exit Sub
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    FIN& = Len(TPSP$)
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7070
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7070
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    UI& = UI& + 1
    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
    If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.HScroll1.Min = 0
    EDITFORM.HScroll1.Value = 0
    EDITFORM.HScroll1.Max = 1
    If ANCHPAGINA > EDITFORM.Frame1.Width Then
      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
    End If
    EDITFORM.Command3.Enabled = False
    EDITFORM.Command1.Enabled = False
    EDITFORM.Show 1
    '
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
    End If
    Call SELECHO("SELFECHA")
    VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
    If VDEVU$ <> "" Then
      Label9 = VALACT1$("SELFECHA")
    End If
    Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
'    Call CIERRARCH("*.*")
'    ACONEC$ = "AMASTO01"
'    If EXISTE%("AMASTO02.EXE") > 0 Then ACONEC$ = "AMASTO02"
'    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
    AMAS_MAQ04.Show 1
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text9.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text9.TEXT)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  While WindowState <> 0
    WindowState = 0
    OPENFORMS = DoEvents
    AppActivate Caption
    SendKeys "%" & Chr$(32) & Chr$(13), True
  Wend
End Sub


Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    If EXISTE%("CALPRO02.EXE") > 0 Then
      Command13.Enabled = True
    End If
    '
    HOII% = HOY(HOS$)
    Text9.TEXT = HOS$
    Text11.TEXT = TRIM$(Str$(PRONUBOL&))
    Label11 = USERNAME
    '
    LETRANO% = Val(CONTROL$("BOLRECEP", "LETRANO"))
    '
    DoEvents
    '
    Call SETULTREG("!BOLREDET", 0)
    '
    Label16 = TRIM$(CONTROL$("", "DEPOENT"))
    '
End Sub
'

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call LIBARCH("*.*")
   Hide
End Sub

Private Sub Label16_Change()
   Label15 = ""
   DP1% = Val(Label16)
   If DP1% > 0 Then
     Label15 = ECOARCH$("ECODEP", Chr$(DP1%))
   End If
End Sub

Private Sub Label16_DblClick()
   Call SELECHO("ECODEP")
   VDEF$ = TRIM$(VALACT1$("ECODEP"))
   If VDEF$ <> "" Then
     Label16 = VDEF$
   End If
End Sub

Private Sub Label22_DBLClick()
   Call CARDETCOM
End Sub

Private Sub Label8_Click()
   On Error Resume Next
   Text13.SetFocus
   On Error GoTo 0
End Sub

Private Sub Label9_DblClick()
    If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
    End If
    Call SELECHO("SELFECHA")
    VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
    If VDEVU$ <> "" Then
      Label9 = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub List1_Click()
   List1.ToolTipText = DESCRIT$(CODINT%(Left$(List1.TEXT, 15)))
End Sub

Private Sub List1_DblClick()
   Call CARDETCOM
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
     Call MUESTRADATCOD(Left$(List1.TEXT, 15))
   End If
   '
End Sub

Private Sub List2_Click()
   If MUECOPEN% < 1 Then
     For U& = 1 To List2.ListCount
       List2.Selected(U& - 1) = False
     Next U&
   End If
End Sub

Private Sub MnuConLis_Click()
   '
   Call Command4_Click
   '
End Sub

Private Sub MnuEnd_Click()
   '
   Call Command2_Click
   '
End Sub

Private Sub MnuMasPro_Click()
   '
   Call Command11_Click
   '
End Sub

Private Sub MnuMasRep_Click()
   '
   Call Command8_Click
   '
End Sub

Private Sub MnuNueBr_Click()
   '
   Call BLANFORMU
   Text1.TEXT = ""
   Call Command1_Click
   '
End Sub

Private Sub MnuStp_Click()
   '
   Call Command5_Click
   '
End Sub

Private Sub Text1_Change()
  '
  Call BLANFORMU
  NCLIE% = Val(TRIM$(Text1.TEXT))
  If NCLIE% > 0 Then
    If REGICLI&((-1) * NCLIE%) > 0 Then
      Text1.TEXT = TRIM$(Str$(NCLIE%))
      Call CARDACLI
      Call CARCOPEN
      Call CARLISRE ''**''
   End If
  End If
  '
End Sub

Private Sub Text1_DblClick()
   Call SELECHO("ACREDOR")
   NCLIE% = Val(TRIM$(VALACT1$("ACREDOR")))
   If NCLIE% > 0 Then
     If REGICLI&((-1) * NCLIE%) > 0 Then
       Call BLANFORMU
       Text1.TEXT = TRIM$(Str$(NCLIE%))
       Call CARDACLI
       Call CARCOPEN
       Call CARLISRE ''**''
       If ULTREG&("!BOLREDET") < 1 Then
         Call CARDETCOM
       End If
     End If
   End If
End Sub

Private Sub Text1_GotFocus()
   Text1.TEXT = TRIM$(Text1.TEXT)
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.TEXT)
   Screen.MousePointer = 1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If VALICLIE% > 0 Then
      Call BLANFORMU
      Call CARDACLI
      Call CARCOPEN
      Call CARLISRE ''**''
'      If ULTREG&("!BOLREDET") < 1 Then
'        Call CARDETCOM
'      End If
    End If
  End If
End Sub

Private Sub Text11_Change()
AAA = BBB
End Sub

Private Sub Text11_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
      Exit Sub
   End If
End Sub

Private Sub Text11_LostFocus()
   NUMEPRE$ = ""
   For KK% = 1 To Len(Text11.TEXT)
     CCX% = Asc(Mid$(Text11.TEXT, KK%, 1))
     If CCX% >= 48 Then
       If CCX% <= 57 Then
         NUMEPRE$ = NUMEPRE$ + Chr$(CCX%)
       End If
     End If
   Next KK%
   '
   If NUMEPRE$ <> TRIM$(Text11.TEXT) Then
     Text11.TEXT = NUMEPRE$
     Call MENSERR(24, "Número de Inf.Recepción Ajustado")
   End If
   '
End Sub

Private Sub Text13_GotFocus()
   If VALICLIE% < 1 Then
      On Error Resume Next
      Text1.SetFocus
      On Error GoTo 0
   End If
End Sub

Private Sub Text2_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text3_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Private Sub Text4_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub


Private Sub Text5_GotFocus()
   On Error Resume Next
   Text1.SetFocus
   On Error GoTo 0
End Sub

Function VALICLIE%()
   VCC% = 0
   If Val(Text1.TEXT) > 0 Then
     If Val(Text1.TEXT) <= 32767 Then
       NCLIE% = Val(Text1.TEXT)
       If ECOARCH$("PROVED1", MKI$(NCLIE%)) <> "" Then
         VCC% = 1
         GoTo 99
       End If
     End If
   End If
   Call MENSERR(24, "Falta Número de Proveedor")
   '
99 VALICLIE% = VCC%
   '
End Function

Sub BLANFORMU()
   Text2.TEXT = ""
   Text3.TEXT = ""
   Text4.TEXT = ""
   Text5.TEXT = ""
   Text13.TEXT = ""
   Text6.TEXT = ""
   List1.Clear: List1.ToolTipText = ""
   List2.Clear
   '
   Option0.Value = True
   Label9 = ""
   Text8 = ""
   Text7 = ""
   '
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   Text11.TEXT = TRIM$(Str$(PRONUBOL))
   Label16 = TRIM$(CONTROL$("", "DEPOENT"))
   '
End Sub
'
Sub CARDACLI()
   '
   Screen.MousePointer = 11
   '
   Option0.Value = True
   Option1.Enabled = False
   Option2.Enabled = False
   Label9 = "": Label9.Enabled = False
   Command7.Enabled = False
   Text7 = "": Text7.Enabled = False
   Text8 = "": Text8.Enabled = False
   '
   NC% = (-1) * Val(Text1.TEXT)
   If NC% < 0 Then
     If ECOARCH$("ACREDOR", MKI$((-1) * NC%)) <> "" Then
       Text2.TEXT = RASOCLI$(NC%)
       Text3.TEXT = DOMICLI$(NC%)
       Text4.TEXT = CPOSCLI$(NC%)
       Text5.TEXT = LOCACLI$(NC%)
       '
       Text15 = TRIM$(Str$(MONEVACO%(NC%)))
       ATEN$ = "": PODESC$ = ""
       RFF$ = RECFILT$("OCOMENCA", 4, MKI$(Abs(NC%)))
       For KKK% = Len(RFF$) - 3 To 1 Step -4
         REOCO& = CVS(Mid$(RFF$, KKK%, 4))
         XX$ = REGLEIDO$("OCOMENCA", REOCO&)
         ATEN$ = TRIM$(Mid$(XX$, 163, 30))
         PODESC$ = TRIM$(Mid$(XX$, 97, 16))
         If ATEN$ <> "" Then Exit For
       Next KKK%
       Text6 = ATEN$
       PORDESCU = PODESC$
       '
       PVPRX% = PIVACLI%(NC%)
       If PVPRX% = 5 Or PVPRX% = 7 Then
         Option1.Enabled = True
         Option2.Enabled = True
         Option2.Value = True
         Label9.Enabled = True
         Command7.Enabled = True
         Text7.Enabled = True
         Text8.Enabled = True
       End If
     End If
   End If
   Screen.MousePointer = 1
   '
End Sub
   '
Function PRONUBOL&()
    '
    PRONUBOL& = 1
    Call ABRESTOCKS
    Set BolreTemp = Stocks.OpenRecordset("SELECT DISTINCT NumeBoRe from BOLRECEP ORDER BY NumeBoRe DESC")
    On Error Resume Next
    BolreTemp.MoveFirst
    If Err < 1 Then
      PRONUBOL& = 1 + BolreTemp!NumeBore
    End If
    On Error GoTo 0
    '
End Function

Function ULTIRMP&(LTX$)
   '
   FIN& = ULTREG&("REMAPRI")
   UIRM& = 0
   LTX1$ = UCase$(LTX$)
   '
   For KKU& = FIN& - 512 To FIN&
     If KKU& > 0 Then
       XX$ = REGLEIDO$("REMAPRI", KKU&)
       IDLOT$ = TRIM$(Mid$(XX$, 69, 12))
       If Left$(IDLOT$, 1) = LTX1$ Then
         UIRM1& = Abs(Mid$(IDLOT$, 2))
         If UIRM1& > UIRM& Then
           UIRM& = UIRM1&
         End If
       End If
     End If
   Next KKU&
   '
   ULTIRMP& = UIRM&
   '
End Function

Private Sub Text6_GotFocus()
   On Error Resume Next
   Text6.SelStart = 0
   Text6.SelLength = Len(Text6.TEXT)
   On Error GoTo 0
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
     If TRIM$(Text6) <> "" Then
        Call CARDETCOM
     End If
   End If
End Sub

Private Sub Text9_Change()
  On Error Resume Next
  ANUXI = Val(Mid$(Text9, 7, 2)) Mod 26
  'If LETRANO% > 0 Then
  '  Label13 = Chr$(LETRANO% + ANUXI)
  'End If
  On Error GoTo 0
End Sub

Private Sub Text9_GotFocus()
   On Error Resume Next
   Command9.SetFocus
   On Error GoTo 0
End Sub

Sub GRABOLRE(OKX%)
   '
   OKX% = 0
   If ULTREG&("!BOLREDET") < 1 Then Exit Sub
   '
   Call BLOQARCH("REMAPRI")
   Call BLOQARCH("REPRIPEN")
   Call BLOQARCH("IREPRIM")
   Call BLOQARCH("IREPRIPE")
   Call BLOQARCH("OBSREPRI")
   '
   NPGRABA& = Val(Text11)
   '
   RFB$ = RECFILT$("REMAPRI", 1, MKS$(NPGRABA&))
   If RFB$ <> "" Then
     For KU& = 1 To Len(RFB$) Step 4
       XX$ = REGLEIDO$("REMAPRI", KU&)
       'If UCase$(Mid$(XX$, 101, 2)) = "MP" Then
       Call CIERRARCH("REMAPRI")
       Call CIERRARCH("REPRIPEN")
       Call CIERRARCH("IREPRIM")
       Call CIERRARCH("IREPRIPE")
       Call CIERRARCH("OBSREPRI")
       '
       Text11.TEXT = TRIM$(Str$(PRONUBOL&))
       Call COMUNI("Número Ocupado al Grabar.\FLEXOFT ha Re-Numerado la Recepción\como I/R Número '" + TRIM$(Text11) + "'.\  \Verifique y Vuelva a Grabar.")
       Exit Sub
       'End If
     Next KU&
   End If
   '
   Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
   '
   AINDI% = 0:   NORVENTA = ""
   NROBOL& = NPGRABA&
   FREC% = FECHANUM(Text9)
   '
   ILOTEI& = 0
   URECEP& = ULTREG&("REMAPRI")
   If URECEP& > 0 Then
     YY$ = REGLEIDO$("REMAPRI", URECEP&)
     ILOTEI& = CVS(Mid$(YY$, 237, 4))
   End If
   If ILOTEI& < URECEP& Then ILOTEI& = URECEP&
   '
   If LETRANO% > 0 Then
     NANO% = Val(Mid$(Text9, 7, 2))
     If NANO% > 18 Then NANO% = 18
     LTTX$ = Chr$(LETRANO% + NANO%)
     UIRMP& = ULTIRMP&(LTTX$)
   End If
   For YX& = 1 To ULTREG&("!BOLREDET")
     YY$ = REGLEIDO$("!BOLREDET", YX&)
     CI% = CVI(Mid$(YY$, 83, 2))
     CAN# = CVD(Mid$(YY$, 89, 8))
     If CAN# >= 0.0005 Then
       If TRIM$(Mid$(YY$, 69, 12)) = TRIM$(Mid$(YY$, 107, 12)) Then
         Call REPLA(YY$, Space$(12), 107, 12)
         Call GRAREG("!BOLREDET", YY$, YX&)
       End If
       REGOC& = CVS(Mid$(YY$, 189, 4))
       Call REPLA(YY$, MKS$(NROBOL&), 1, 4)
          DESCRI$ = TRIM$(CODEXT$(CI%)) + " - " + TRIM$(DESCRIT$(CI%))
       Call REPLA(YY$, DESCRI$, 5, 64)
       If TRIM$(Mid$(YY$, 69, 12)) = "" Then
          NBR$ = TRIM$(Text11)
          If LETRANO% > 0 Then
              LTTX$ = Chr$(LETRANO% + NANO%)
              UIRMP& = UIRMP& + 1
              IDENLO$ = LTTX$ + "-" + TRIM$(Str$(UIRMP&))
            Else
              While Len(NBR$) < 6: NBR$ = "0" + NBR$: Wend
              NORDE% = NORDE% + 1
              NORDES$ = TRIM$(Str$(NORDE%))
              IDENLO$ = "BR-" + NBR$ + "-" + NORDES$
          End If
          Call REPLA(YY$, IDENLO$, 69, 12)
       End If
       IDENLO$ = TRIM$(Mid$(YY$, 69, 12))
       If TRIM$(Mid$(YY$, 107, 12)) = "" Then
         Call REPLA(YY$, IDENLO$, 107, 12)
       End If
       Call REPLA(YY$, MKI$(FREC%), 81, 2)
       Call REPLA(YY$, UNIMED$(CI%), 85, 4)
          CARECI = CAN#
          NORVE& = CVS(Mid$(YY$, 151, 4))
          DEPX% = Val(Label16) Mod 256
       Call REPLA(YY$, Chr$(DEPX%), 97, 1)
       Call REPLA(YY$, Chr$(0), 98, 1) ' STATUS = 0
       Call REPLA(YY$, MKI$(0), 99, 2) ' NO ESTA APROBADO
       Call REPLA(YY$, "MP", 101, 2) ' marca materia prima
          NPXX% = Val(Text1)
       Call REPLA(YY$, MKI$(NPXX%), 105, 2)
          REMI$ = TRIM$(Text6)
       Call REPLA(YY$, REMI$, 119, 16)
       If Check1.Value = 1 Then ' aprueba sin Control
          Call REPLA(YY$, MKD$(CAN#), 163, 8)
          Call REPLA(YY$, Chr$(1), 98, 1)  ' marca aprobado
          Call REPLA(YY$, MKI$(FREC%), 99, 2)
       End If
       NUORSE& = CVS(Mid$(YY$, 189, 4))
       ' Call REPLA(YY$, MKS$(NUORSE&), 189, 4)
       '
       Call REPLA(YY$, "L ", 193, 2)
       Call REPLA(YY$, MKI$(0), 195, 2)
       Call REPLA(YY$, Space$(40), 197, 40)
       '
       ILOTEI& = ILOTEI& + 1
       Call REPLA(YY$, MKS$(ILOTEI&), 237, 4)
       '
       If Option1.Value = True Then
         Call REPLA(YY$, "XT", 193, 2)
           FADU% = FECHANUM(Label9)
         Call REPLA(YY$, MKI$(FADU%), 195, 2)
         Call REPLA(YY$, AJUSTI$(Text8, 16), 197, 16)
         Call REPLA(YY$, AJUSTI$(Text7, 24), 213, 24)
       End If
       '
       If Option2.Value = True Then
         Call REPLA(YY$, "XC", 193, 2)
           FADU% = FECHANUM(Label9)
         Call REPLA(YY$, MKI$(FADU%), 195, 2)
         Call REPLA(YY$, AJUSTI$(Text8, 16), 197, 16)
         Call REPLA(YY$, AJUSTI$(Text7, 24), 213, 24)
       End If
       '
       ' agregar a REMAPRI y REPRIPEN
       Call GRAREG("!BOLREDET", YY$, YX&)
       Call REPLA(YY$, MKI$(0), 187, 2) ' blanquea fecha entrega solicitada
       Call REGAPP("REMAPRI", YY$)
       REG& = ULTREG&("REMAPRI")
       '
       If Check1.Value < 1 Then ' NO APRUEBA SIN CONTROL
         Call REPLA(YY$, MKS$(REG&), 189, 4)
         Call REGAPP("REPRIPEN", YY$)
       End If
       '
       ' agregar a IREPRIM e IREPRIPE
       If AINDI% = 0 Then
         X$ = REGBLAN$("IREPRIM")
         Call REPLA(X$, MKS$(NROBOL&), 1, 4)
         Call REPLA(X$, TRIM$(Text9), 5, 8)
         Call REPLA(X$, REMI$, 15, 16)
         Call REPLA(X$, ECOARCH$("PROVED1", MKI$(NPXX%)), 32, 29)
         Call REPLA(X$, MKS$(REG&), 61, 4)
         Call REGAPP("IREPRIM", X$)
         '
         If Check1.Value < 1 Then ' NO APRUEBA SIN CONTROL
           Call REGAPP("IREPRIPE", X$)
         End If
         '
         AINDI% = 1
       End If
       '
       If Check1.Value > 0 Then ' si aprueba sin control then
         '
         DOPRI$ = "BR." + TRIM$(Str$(NROBOL&))
         LOTE$ = TRIM$(Mid$(YY$, 69, 12))
         If InStr(EMPREAC$, "SABO") > 0 Then DOPRI$ = LOTE$
         CANTI = CAN#
         NC1% = (-1) * Abs(NPXX%)
         REFE$ = "Rto." + TRIM$(REMI$) + " - " + RASOCLI$(NC1%)
         DOSEC$ = TRIM$(REMI$)
         While Len(DOSEC$) > 10: DOSEC$ = Mid$(DOSEC$, 2): Wend
         If Len(DOSEC$) < 8 Then DOSEC$ = "Rt." + DOSEC$
         'Call Movisto(FREC%, CI%, LOTE$, "BR", DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC1%, DEPX%, CANTI)
         '
       End If
     End If
   Next YX&
   '
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("REMAPRI")
   Call CIERRARCH("REPRIPEN")
   Call CIERRARCH("IREPRIM")
   Call CIERRARCH("IREPRIPE")
   Call CIERRARCH("OBSREPRI")
   '
   Call FRATEXTO(Text13, 60)
   For YX& = 1 To CARETEX%
     YY$ = MKS$(NROBOL&) + RETEX$(YX&)
     Call REGAPP("OBSREPRI", YY$)
   Next YX&
   Call CIERRARCH("OBSREPRI")
   '
   Call FRESHECHO("IREPRIPE")
   '
   Screen.MousePointer = 1
   OKX% = 1
   Exit Sub
   '
98 Call CIERRARCH("MOVISTO")
   Call CIERRARCH("REMAPRI")
   Call CIERRARCH("REPRIPEN")
   Call CIERRARCH("IREPRIM")
   Call CIERRARCH("IREPRIPE")
   Call CIERRARCH("OBSREPRI")
   '
   Screen.MousePointer = 1
   Call MENSERR(24, MERRO$)
   '
End Sub

Sub CARDETCOM()
   '
   If TRIM$(Text6) = "" Then
     Call COMUNI("Debe Ingresar el Número de \Remito del Proveedor")
     Text6.SetFocus
     Exit Sub
   End If
   '
   FECHAREC$ = REPLACAR$(Text9, "/", ".")
   NROBOL& = Val(Text11)
   NPRO% = Val(Text1)
   NC1% = (-1) * NPRO%
5  VUELTA% = 1
   '
10 ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=Recepción de Materiales "
   ATRI$ = ATRI$ + TRIM$(Str$(NROBOL&)) + " - " + TRIM$(RASOCLI$(NC1%)) + " - " + FECHAREC$
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(140 + Top + Frame1.Top + Frame4.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Frame4.Left))
   '
   AVENTA$ = "RECEMATE"
   VTB% = VENTABU%(AVENTA$ + "/NC/" + ATRI$)
   If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
     If VUELTA% > 1 Then GoTo 5
     Exit Sub
   End If
   '
   If LETRANO% > 0 Then
     NANO% = Val(Mid$(Text9, 7, 2))
     If NANO% > 18 Then NANO% = 18
     LTTX$ = Chr$(LETRANO% + NANO%)
     UIRMP& = ULTIRMP&(LTTX$)
   End If
   '
   NROBOL& = Val(Text11)
   JJ& = 0: ATTX% = 0
   For YX& = 1 To ULTREG&("!BOLREDET")
     YY$ = REGLEIDO$("!BOLREDET", YX&)
     CI% = CVI(Mid$(YY$, 83, 2))
     CAN# = CVD(Mid$(YY$, 89, 8))
     REC# = CVD(Mid$(YY$, 171, 8))
     If REC# < 0.0005 Then 'CAN# < 0.0005
         ATTX% = 1
       Else
         JJ& = JJ& + 1
         Call REPLA(YY$, MKS$(NROBOL&), 1, 4)
         If VUELTA% = 1 Or TRIM$(Mid$(YY$, 69, 12)) = "" Then
            NBR$ = TRIM$(Text11)
            If LETRANO% > 0 Then
                LTTX$ = Chr$(LETRANO% + NANO%)
                UIRMP& = UIRMP& + 1
                IDENLO$ = LTTX$ + "-" + TRIM$(Str$(UIRMP&))
              Else
                While Len(NBR$) < 6: NBR$ = "0" + NBR$: Wend
                NORDE% = NORDE% + 1
                NORDES$ = TRIM$(Str$(NORDE%))
                IDENLO$ = "BR-" + NBR$ + "-" + NORDES$
            End If
            ATXX% = 1
            Call REPLA(YY$, IDENLO$, 69, 12)
         End If
         '
         IDENLO$ = TRIM$(Mid$(YY$, 69, 12))
         If TRIM$(Mid$(YY$, 107, 12)) = "" Or VUELTA% = 1 Then
            Call REPLA(YY$, IDENLO$, 107, 12)
            ATTX% = 1
         End If
         '
         If LETRANO% > 0 Then
           If TRIM$(Mid$(YY$, 107, 12)) <> TRIM$(IDENLO$) Then
             Call REPLA(YY$, IDENLO$, 107, 12)
             ATTX% = 1
           End If
         End If
         Call GRAREG("!BOLREDET", YY$, JJ&)
     End If
   Next YX&
   Call SETULTREG("!BOLREDET", JJ&)
'   If ATTX% > 0 Or VUELTA% = 1 Then
'     VUELTA% = 2
'     GoTo 10
'   End If
   '
   For YX& = 1 To ULTREG&("!BOLREDET")
     YY$ = REGLEIDO$("!BOLREDET", YX&)
     CI% = CVI(Mid$(YY$, 83, 2))
     If CI% <= 0 Or CI% > NUMAS% Then
       Call MENSERR(24, "Revise !!! - Código Inexistente o no Válido.")
       GoTo 10
     End If
   Next YX&
   '
   Call CARLISRE
   Screen.MousePointer = 1
   '
End Sub

Sub CARCOPEN()
   '
   NPRO% = Val(Text1)
   HO% = HOY(HOS$)
   '
   REBLA$ = REGBLAN$("BOLREDET")
   Call SETULTREG("!BOLREDET", 0)
   '
   Screen.MousePointer = 11
   HAYOC% = 0: INCLUPOST% = 0
   '
    Call ABRECOMPRAS
    SQLX$ = "Select * from OCOMDETA"
    SQLX$ = SQLX$ + " Where Stat_Ocom < 8 "
    SQLX$ = SQLX$ + " AND  Stat_Ocom <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
    SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
    SQLX$ = SQLX$ + " AND Npro_Ocom = " & NPRO% & " "
    SQLX$ = SQLX$ + " AND Cod_Inte > 0 " 'EXCLUYE LAS DE 0/C's TEXTO LIBRE
    Set COMPTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With COMPTEMP
      '
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Screen.MousePointer = 1
        HAYOC% = 0
        GoTo 999
      End If
      On Error GoTo 0
      '
      JJ& = 0
      Do While (.EOF = False)
        CAPEN = !Cant_Ocom - !Cant_rec
        COEFUNI = !Coef_Unids
        If COEFUNI < 0.0001 Then COEFUNI = 1
        If CAPEN > 0.005 Then
          NUOC& = !Nume_Ocom
          FENTRESO% = CVINT(!Fentre_Sol)
          If FENTRESO% > HO% Then
            If INCLUPOST% = 0 Then
              OPXX% = COMALTER%("Para este Proveedor Hay Compras en Curso\con Fecha de Entrega Solicitada POSTERIOR a la Actual\\Incluirlas\Excluirlas")
              INCLUPOST% = 1
              If OPXX% = 2 Then INCLUPOST% = (-1)
            End If
          End If
          '
          If FENTRESO% <= HO% Or INCLUPOST% = 1 Then
            CI% = !Cod_Inte
            LOTECO$ = !Doc_Orig
            Y$ = REGBLAN$("!BOLREDET")
            Call REPLA(Y$, MKI$(CI%), 83, 2)
            Call REPLA(Y$, UNIMED$(CI%), 85, 4)
            Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 89, 8)
            Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 163, 8)
            If ANUCA% = 1 Then
              Call REPLA(Y$, MKD$(0), 89, 8)
            End If
            '
            Call REPLA(Y$, MKI$(!NuOrd_Item), 98, 2)
            '
            Call REPLA(Y$, LOTECO$, 107, 12)
            Call REPLA(Y$, MKS$(NUOC&), 151, 4)
            Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
            Call REPLA(Y$, MKS$(!N_OrServ), 189, 4)
            HAYOC% = 1
            JJ& = JJ& + 1
            Call GRAREG("!BOLREDET", Y$, JJ&)
          End If
        End If
       .MoveNext
      Loop
    End With
   '
   If HAYOC% = 0 Then
     Call COMUNI("No hay Ordenes de Compra Pendientes\para este Proveedor.")
   End If
   '
   Call SETULTREG("!BOLREDET", JJ&)
   Call CIERRARCH("!BOLREDET")
   '
999 Screen.MousePointer = 1
   '
End Sub

Sub CARLISRE()
   '
   Screen.MousePointer = 11   ' cargando Lista de Compra
   HOII% = HOY(HOS$)
   List1.Clear: List1.ToolTipText = ""
   FIN& = ULTREG&("!BOLREDET")
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   For U& = 1 To FIN&
     X$ = REGLEIDO$("!BOLREDET", U&)
     CI% = CVI(Mid$(X$, 83, 2))
     FENTRESOL% = CVI(Mid$(X$, 187, 2))
     UNIST$ = UCase$(TRIM$(Mid$(X$, 85, 4)))
       If TRIM$(UNIST$) = "" Then UNIST$ = UNIMED$(CI%)
     CANPED = CVD(Mid$(X$, 163, 8))
     CANREC = CVD(Mid$(X$, 171, 8))
     NUOCO = CVS(Mid$(X$, 151, 4))
     '
     TTXX$ = Space$(128)
     Call REPLA(TTXX$, CODEXT$(CI%), 1, 15)
     Call REPLA(TTXX$, DESCRIT0$(CI%), 17, 35)
     Call REPLA(TTXX$, FORMATNUM$(NUOCO, "F6.0"), 52, 6)
     Call REPLA(TTXX$, FORMATNUM$(CANPED, "F7.1"), 65, 7)
     Call REPLA(TTXX$, FORMATNUM$(CANREC, "F9.2"), 76, 9)
     List1.AddItem TTXX$
     '
   Next U&
   Call CIERRARCH("!BOLREDET")
   '
   List1.Refresh
   Screen.MousePointer = 1
   '
End Sub

Function VALIDATOS%()
   '
   If VALICLIE% < 1 Then GoTo 99
   If ULTREG&("!BOLREDET") < 1 Then
     Call MENSERR(24, "Falta Detalle de Recepción")
     GoTo 99
   End If
   '
   VDT% = 0
   FECHAREC$ = Text9
   If FECHANUM(FECHAREC$) < 1 Then
     Call MENSERR(24, "Fecha de Ingreso no Procesable")
     GoTo 99
   End If
   '
   REMI$ = TRIM$(Text6)
   If REMI$ = "" Then
     Call MENSERR(24, "Falta Número de Remito del Proveedor")
     GoTo 99
   End If
   '
   DP1% = Val(Label16)
   If DP1% < 1 Or ECOARCH$("ECODEP", Chr$(DP1%)) = "" Then
     Call MENSERR(24, "Falta Depósito de Entrada")
     GoTo 99
   End If
   '
   NC1% = (-1) * Val(Text1)
   PIVAPR% = PIVACLI%(NC1%)
   If PIVAPR% = 5 Or PIVAPR% = 7 Then
    ' FEADU$ = TRIM$(Label9)
    ' ADUAN$ = TRIM$(Text8)
    ' DESPAI$ = TRIM$(Text7)
     '
    ' If FECHANUM(FEADU$) < 1 Or ADUAN$ = "" Or DESPAI$ = "" Then
    '    Call MENSERR(24, "Datos de Importación son MANDATORIOS (!!!)\para Proveedores del Exterior.")
    '    GoTo 99
    ' End If
   End If
   '
   For YX& = 1 To ULTREG&("!BOLREDET")
     YY$ = REGLEIDO$("!BOLREDET", YX&)
     CI% = CVI(Mid$(YY$, 83, 2))
     REC# = CVD(Mid$(YY$, 171, 8))
     If CI% < 1 Or CI% > NUMAS% Then
       Call MENSERR(24, "Código no Válido.")
       GoTo 99
     End If
     '
     If REC# < 0.005 Then
       Call MENSERR(24, "Falta Cantidad.")
       GoTo 99
     End If
   Next YX&
   '
   VDT% = 1
   '
99 VALIDATOS% = VDT%
   '
End Function

Sub PRIBOLRE(MODOPRI%)
   '
   FORIPRE$ = CONTROL$("", "FORMIREC")
   If TRIM$(FORIPRE$) <> "" Then
     '
     CANCELPRINT% = 0
     If CONTROL("REMAPRI", "SUPRINT") = "SI" Then CANCELPRINT% = 1
     '
     TIPODOC$ = "Informe de Recepción "
     DESTIDOC% = (-1) * Val(Text1)
     NUMEDOC$ = "IR-" + TRIM$(Text11)
     FECHDOC$ = TRIM$(Text9)
     '
     CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = Label15
     CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = Text6
     CODPARAM$(3) = "N-AUTENT": DOCPARAM$(3) = NORVENTA
     CODPARAM$(4) = "FECH-LAN": DOCPARAM$(4) = Label9
     CODPARAM$(5) = "ZONA-VEN": DOCPARAM$(5) = Text8
     CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = Text7
     CODPARAM$(7) = "COND-TRA": DOCPARAM$(7) = ""
        If Option1.Value = True Then DOCPARAM$(7) = "Temporaria"
        If Option2.Value = True Then DOCPARAM$(7) = "A Consumo"
     CODPARAM$(8) = "REF-MAT1": DOCPARAM$(8) = ""
        If MODOPRI% = 0 Then DOCPARAM$(8) = "VISTA PREVIA"
     CAPARDOC% = 8
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
     CODTABU1$(11) = "NOP-OPER"
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
              NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
              If NULOTE$ = "" Then
                NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
              End If
              NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
              NUOFA$ = CSTRING$(Mid$(YY$, 189, 4), 3, 7, 0, 2)
              '
              CAITAB1% = CAITAB1% + 1
              TETABU1$(1, CAITAB1%) = TRIM$(CODEXT$(CI%))
                DESMAT$ = DESCRIT$(CI%)
              TETABU1$(2, CAITAB1%) = DESMAT$
              TETABU1$(3, CAITAB1%) = COLOR$(CI%)
              TETABU1$(4, CAITAB1%) = UNID$
              TETABU1$(5, CAITAB1%) = NULOTE$
              TETABU1$(6, CAITAB1%) = CANS$
              TETABU1$(7, CAITAB1%) = NUOCO$
              TETABU1$(8, CAITAB1%) = KILAJE$
              TETABU1$(9, CAITAB1%) = FEPLA$
              TETABU1$(10, CAITAB1%) = FESOLEN$
              TETABU1$(11, CAITAB1%) = NUOFA$
              '
            End If
          End If
        End If
     Next YX&
     '
     Call PRINTDOC("FORMIREC")
   End If
   '
End Sub
Sub BOLERECEP()
Dim IMPCONOC%
'
1105 If ULTREG&("TATIREC") < 1 Then
         X$ = REGBLAN$("TATIREC")
         Call REPLA(X$, Chr$(1), 1, 1)
         Call REPLA(X$, "Material General", 2, 31)
         Call REPLA(X$, "BOLREGEN", 49, 16)
         Call REPLA(X$, "BOLREGEN", 65, 16)
         Call REGAPP("TATIREC", X$)
         Call CIERRARCH("TATIREC")
     End If
     '
1100 Call CIERRARCH("*.*")
     '
     NROBOL& = Val(Text11)
     Call ABRESTOCKS
     Set BolreTemp = Stocks.OpenRecordset("SELECT DISTINCT NumeBoRe from BOLRECEP ORDER BY NumeBoRe DESC")
     On Error Resume Next
     BolreTemp.MoveFirst
     If Err < 1 Then
       NROBOL& = 1 + BolreTemp!NumeBore
     End If
     On Error GoTo 0
     '
     If TRIM$(Mid$(EBOLAN$, 4, 12)) <> "" Then
       NC% = 0
     End If
     '
     NC3% = Val(Text1.TEXT) ' CVI(Left$(EBOL$, 2))
     If NC3% > 0 Then
       NC% = NC3%
       GoTo 15
     End If
     '
     NC% = NUMEPRO%(TRIM$(Mid$(EBOL$, 4, 12)))
15   If NC% < 1 Or STATCLI%((-1) * NC%) < 1 Then
       Call MENSERR(24, "Proveedor Inexistente")
       GoTo 1100
     End If
     '
     FF% = FECHANUM(Text9.TEXT) 'TEXTCVI(Mid$(PEDY$, 5, 2))
     FEX = FF%: FECHAREC$ = FECHATEX$(FEX)
     If TRIM$(FECHAREC$) = "" Then
         Call MENSERR(24, "Fecha Incorrecta")
         GoTo 1199
     End If
     '
     REMI$ = Text6.TEXT 'TRIM$(Mid$(PEDY$, 105, 16))
     If REMI$ = "" Then
         Call MENSERR(24, "Falta Número de Remito del Proveedor")
         GoTo 1199
     End If
     '
     DP1% = Val(Label16) 'Asc(Mid$(PEDY$, 121, 1))
     DEPOSAL$ = TRIM$(ECOARCH$("ECODEP", Chr$(DP1%)))
     If DP1% < 1 Or DEPOSAL$ = "" Then
         Call MENSERR(24, "Falta Depósito de Entrada")
         GoTo 1199
     End If
     DEPOSAL$ = DEPOSAL$ + " (" + TRIM$(Str$(DP1%)) + ")"
     '
     'ACA NUEVO PARA GUIDI(COMPRUEBA QUE NO EXISTA EL REMI)
     NC2% = NC%
     NR1$ = REMI$
     R1% = EXIREMI(NR1$, NC2%)
     If R1% = 1 Then
        Call COMUNI("El Número de Remito ya fue Procesado\en Recepción Anterior.")
        GoTo 1199
     End If
     'HASTA ACA NUEVO PARA GUIDI
     '
     For YX& = 1 To ULTREG&("!BOLREDET")
       YY$ = REGLEIDO$("!BOLREDET", YX&)
       CI% = CVI(Mid$(YY$, 83, 2))
       CAN# = CVD(Mid$(YY$, 89, 8))
       REC1# = CVD(Mid$(YY$, 171, 8))
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           If REC1# > 0.005 Then
             GoTo 1141
           End If
         End If
       End If
     Next YX&
     Call MENSERR(24, "Recepción Cancelada !!!\Faltan Cantidades o Códigos.")
     GoTo 1199
     '
1141 CONTA% = ULTREG&("!BOLREDET")
     '
1115 If CONTA% > 0 Then
       '
       For YX& = 1 To ULTREG&("!BOLREDET")
         YY$ = REGLEIDO$("!BOLREDET", YX&)
         CI% = CVI(Mid$(YY$, 83, 2))
         CAN# = CVD(Mid$(YY$, 89, 8))
         REC1# = CVD(Mid$(YY$, 171, 8))
         If CI% < 1 Then
           Call MENSERR(24, "Código no Válido.")
           GoTo 1199
         End If
         If ANUCA% < 1 Then
           If REC1# < 0.005 Then
             Call MENSERR(24, "Falta Cantidad.")
             GoTo 1199
           End If
         End If
1114   Next YX&
       '
1125   Call ABRESTOCKS
       Set BolreTemp = Stocks.OpenRecordset("SELECT DISTINCT NumeBoRe from BOLRECEP ORDER BY NumeBoRe DESC")
       On Error Resume Next
       BolreTemp.MoveFirst
       If Err < 1 Then
         If NROBOL& <= BolreTemp!NumeBore Then
           NROBOL& = 1 + BolreTemp!NumeBore
         End If
       End If
       On Error GoTo 0
       BolreTemp.Close
       '
       Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
       Call ABRESTOCKS
       Call ABRECOMPRAS
       '
       AINDI% = 0
       NORVENTA = ""
       For YX& = 1 To ULTREG&("!BOLREDET")
         YY$ = REGLEIDO$("!BOLREDET", YX&)
         CI% = CVI(Mid$(YY$, 83, 2))
         CAN# = CVD(Mid$(YY$, 89, 8))
         REC1# = CVD(Mid$(YY$, 171, 8))
         NUOC& = Abs(CVS(Mid$(YY$, 151, 4)))
         NUMOIT% = CVI(Mid$(YY$, 98, 2))   'EL Nº DE ORDEN DEL ITEM
         If CI% > 0 Then
           If REC1# > 0.005 Then
             CARECI = REC1#
             LOTECO$ = TRIM$(Mid$(YY$, 107, 12))
             '
             If CVD(Mid$(YY$, 89, 8)) >= 0.05 Then   ' solo si es comprado - no M.P. devuelto
               With OcomDeta
                 .FindFirst ("Nume_Ocom = " & NUOC& & " AND NuOrd_Item = " & NUMOIT% & " AND Stat_Ocom < 8 ")
                 If .NoMatch = False Then
                   .Edit
                   If IsNull(!IContab) Then
                        IMPCONOC% = 0
                      Else
                        IMPCONOC% = !IContab 'ASIGNO LA IMPUTACION CONTABLE
                   End If
                     COEFUNI = !Coef_Unids
                     If COEFUNI < 0.0001 Then COEFUNI = 1
                   !Cant_rec = !Cant_rec + CARECI / COEFUNI
                     CAPENDIEN = !Cant_Ocom - !Cant_rec
                     NUORSE& = !N_OrServ
                   If CAPENDIEN >= 0.005 Then
                       If COMALTER%("Esta O/C por Código " + TRIM$(CODEXT$(CI%)) + "\Queda con un Saldo Pendiente de " + TRIM$(FORMATNUM$(CAPENDIEN, "F10.1")) + " " + TRIM$(UNIMED$(CI%)) + ".\\Mantener Abierta\Cancelar Pendiente") = 2 Then
                         !Stat_Ocom = 8
                       End If
                     Else
                       !Stat_Ocom = 3
                   End If
                   .Update
                 End If
               End With
             End If
             '
             ' aca arma un string con las ordenes de compra recibidas
             NORVE& = Abs(CVS(Mid$(YY$, 151, 4)))
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
             '\\\OT-5-0413-EPB-21/06/05///
             Call CIERRARCH("INDIBOL")
             Call CIERRARCH("IBOLPEN")
             '\\\OT-5-0413-EPB-FIN///
             '
             NBR$ = TRIM$(Str$(NROBOL&))
             NORDE% = NORDE% + 1
             NORDES$ = TRIM$(Str$(NORDE%))
             IDENLO$ = "BR-" + NBR$ + "-" + NORDES$
             '
             Call REPLA(YY$, MKS$(NROBOL&), 1, 4)
             If TIMATEI% <> 2 Then Call REPLA(YY$, DESCRI$, 5, 64)
             '
             If TRIM$(Mid$(YY$, 69, 12)) = "" Then
               Call REPLA(YY$, IDENLO$, 69, 12)
             End If
             Call REPLA(YY$, MKI$(FF%), 81, 2)
             DEPX% = DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(CI%)
             Call REPLA(YY$, Chr$(DEPX%), 97, 1)
             Call REPLA(YY$, MKI$(NC%), 105, 2)
             Call REPLA(YY$, REMI$, 119, 16)
             If CVD(Mid$(YY$, 163, 8)) >= 0.05 Then   ' solo si es comprado - no M.P. devuelto
               Call REPLA(YY$, MKS$(NUORSE&), 189, 4)
             End If
             Call GRAREG("!BOLREDET", YY$, YX&)
             '
             With Bolrecep
              .AddNew
               !NumeBore = NROBOL&
               !DESCRIP = Left$(DESCRI$, 64)
               !IdenLote = Left$(IDENLO$, 12)
               !FechRecep = CVDAT(FF%)
               !Cod_Inte = CI%
               !Cod_Exte = Left$(CODEXT$(CI%), 16)
               !DescripItem = Left$(DESCRIT0$(CI%), 64)
               !Unidad = Left$(UNIMED$(CI%), 4)
               !CantNom = REC1#
               !DepoEnt = DEPX%
               !NUMPROV = NC%
               !RaSo_Prov = Left$(RASOCLI$((-1) * NC%), 64)
                  LOTECO$ = TRIM$(LOTECO$)
                  While Len(LOTECO$) > 12
                    LOTECO$ = Mid$(LOTECO$, 2)
                  Wend
               !LoteCol_Prov = Left$(LOTECO$, 12)
               !NumRem_Prov = Left$(REMI$, 16)
               !NumFac_Prov = ""
               !NumOC = NUOC&
               !IContab = IMPCONOC% 'Imputación contable de la OC.
               !MoneVaUni = ""
               !Valo_Unit = 0
               !CanToApro = 0
               !CanToRech = 0
               !CanToFalta = 0
               !FechLibe = CVDAT(0)
               !NumeOrdSer = NUORSE&
               !Observa = Text13.TEXT
               !Status = 1 'esta marca se pone en la aprobación
               !USUAPRO = Label11
               !FeHoApro = CVDAT(0)
              .Update
                '
                NC1% = NC% * (-1)
                LOTE$ = IDENLO$
                DOPRI$ = "BR." + TRIM$(Str$(NROBOL&))
                DOPRILA$ = DOPRI$
                DOSELA$ = TRIM$(REMI$)
                DOSEC$ = TRIM$(REMI$)
                DEPX% = DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(CI%)
                CANTI = REC1#
                REFE$ = "Rto." + TRIM$(REMI$) + " - " + RASOCLI$(NC1%)
                PPX% = InStr(REFX$, "- BC D.")
                If PPX% > 0 Then
                    REFE$ = Mid$(REFX$, PPX% + 5)
                End If
                While Len(DOSEC$) > 10: DOSEC$ = Mid$(DOSEC$, 2): Wend
                If Len(DOSEC$) < 8 Then DOSEC$ = "Rt." + DOSEC$
                Call MOVISTOK(FF%, CI%, LOTE$, "BR", DOPRILA$, DOPRI$, DOSELA$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CANTI)
                '
             End With
             '
           End If
         End If
       Next YX&
       '
       Bolrecep.Close
       OcomEnca.Close
       OcomDeta.Close
       '
       FORIPRE$ = TRIM$(CONTROL("", "FORMBREC"))
       If TRIM$(FORIPRE$) <> "" Then
         '
         CANCELPRINT% = 0
         If CONTROL("BOLRECEP", "SUPRINT") = "SI" Then CANCELPRINT% = 1
         '
         TIPODOC$ = "Boleta de Recepción"
         DESTIDOC% = (-1) * Abs(NC1%)
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
         CACOLTAB1% = 10
         '
         CAITAB1% = 0
         For YX& = 1 To ULTREG&("!BOLREDET")
           YY$ = REGLEIDO$("!BOLREDET", YX&)
           CI% = CVI(Mid$(YY$, 83, 2))
           CAN# = CVD(Mid$(YY$, 89, 8))
           REC1# = CVD(Mid$(YY$, 171, 8))
           If CI% > 0 Then
             If CI% <= NUMAS% Then
               If REC1# > 0.005 Then
                 '
                 UNID$ = TRIM$(Mid$(YY$, 85, 4))
                 If UNID$ = "" Then UNID$ = UNIMED$(CI%)
                 CANS$ = CSTRING$(MKD$(REC1#), 3, 9, 1, 2)
                 KILAJE$ = CSTRING$(MKD$(REC1# * PESUNI(CI%)), 3, 10, 3, 2)
                 FEX = CVI(Mid$(YY$, 155, 2))
                 FESOLEN$ = FECHATEX$(FEX)
                 FEPLA$ = FEPLANO$(CI%)
                 NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
                 If NULOTE$ = "" Then
                    NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
                 End If
                 NUOCO$ = FORMATNUM$(Abs(CVS(Mid$(YY$, 151, 4))), "F7,0")
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
1199 Screen.MousePointer = 1
     End If
End Sub
Function EXIREMI(NR1$, NC1%)
    '
    Call ABRESTOCKS
    SQLX$ = "SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + " WHERE TRIM(NumRem_Prov) = '" & TRIM(NR1$) & "' "
    SQLX$ = SQLX$ + " AND NumProv = " & NC1% & ""
    Set EXITMP = Stocks.OpenRecordset(SQLX$, 4)
    '
    EXIREMI = 0
    '
    With EXITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Exit Function
      End If
      On Error GoTo 0
      EXIREMI = 1
    End With
    '
End Function

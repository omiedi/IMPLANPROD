VERSION 5.00
Begin VB.Form PEDCOT07 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedido de Cotización"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0F0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8590
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12000
      Begin VB.Frame Frame5 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Solicitudes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6480
         TabIndex        =   41
         Top             =   1560
         Width           =   4280
         Begin VB.CommandButton Command1 
            Caption         =   "Selecc. SIM"
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
            Left            =   360
            TabIndex        =   42
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Identificación"
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
         Height          =   1410
         Left            =   6500
         TabIndex        =   26
         Top             =   105
         Width           =   4280
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
            Height          =   300
            Left            =   3990
            TabIndex        =   44
            ToolTipText     =   "Selección de Destino"
            Top             =   600
            Width           =   195
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   600
            Width           =   2115
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
            Left            =   3990
            TabIndex        =   28
            Top             =   240
            Width           =   195
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
            Left            =   3045
            TabIndex        =   27
            TabStop         =   0   'False
            Text            =   " "
            Top             =   210
            Width           =   960
         End
         Begin VB.Label Label7 
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
            TabIndex        =   36
            Top             =   960
            Width           =   2220
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
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
            Left            =   900
            TabIndex        =   35
            Top             =   240
            Width           =   1170
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Solicitante:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   330
            Left            =   600
            TabIndex        =   34
            Top             =   1005
            Width           =   1230
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Destino:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   300
            Left            =   360
            TabIndex        =   32
            Top             =   675
            Width           =   1485
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
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   330
            Left            =   2040
            TabIndex        =   30
            Top             =   285
            Width           =   750
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "N-C:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   300
            Left            =   -720
            TabIndex        =   29
            Top             =   285
            Width           =   1485
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Invitados"
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
         Height          =   1770
         Left            =   1920
         TabIndex        =   23
         Top             =   120
         Width           =   4140
         Begin VB.ListBox List4 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   960
            Left            =   120
            TabIndex        =   24
            ToolTipText     =   "Doble-Click para Ingresar Proveedores Invitados"
            Top             =   630
            Width           =   3900
         End
         Begin VB.Label Label92 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Pr.            Razon Social       "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   300
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   3900
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00D0F0F0&
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
         Height          =   2990
         Left            =   160
         TabIndex        =   12
         Top             =   2640
         Width           =   10580
         Begin VB.Frame Frame6 
            Caption         =   "Frame18"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4530
            Left            =   7680
            TabIndex        =   22
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame3 
            Caption         =   "Frame18"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4530
            Left            =   6580
            TabIndex        =   21
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame7 
            Caption         =   "Frame7"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4530
            Left            =   1810
            TabIndex        =   15
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame18 
            Caption         =   "Frame18"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4530
            Left            =   5520
            TabIndex        =   14
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame19 
            Caption         =   "Frame19"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4530
            Left            =   10290
            TabIndex        =   13
            Top             =   315
            Width           =   15
         End
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2400
            Left            =   0
            TabIndex        =   18
            ToolTipText     =   "Doble Click para Cargar Detalle"
            Top             =   600
            Width           =   10560
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   10500
            TabIndex        =   16
            Top             =   315
            Width           =   10565
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   $"PEDCOT07.frx":0000
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404000&
               Height          =   225
               Left            =   0
               TabIndex        =   17
               ToolTipText     =   "Doble Click para Cargar Detalle"
               Top             =   0
               Width           =   10515
            End
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
         TabIndex        =   19
         ToolTipText     =   "Lista de Proveedores con Entregas Programadas"
         Top             =   9520
         Width           =   195
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00D0F0F0&
         Caption         =   "Condiciones / Observaciones"
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
         Height          =   2325
         Left            =   160
         TabIndex        =   10
         Top             =   5820
         Width           =   10590
         Begin VB.TextBox Text13 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   320
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            ToolTipText     =   "Click para Ingresar las Condiciones / Observaciones"
            Top             =   330
            Width           =   9960
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00004000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   9000
         Left            =   11040
         ScaleHeight     =   8940
         ScaleWidth      =   1005
         TabIndex        =   1
         Top             =   0
         Width           =   1065
         Begin VB.CommandButton Command27 
            BackColor       =   &H00E0E0E0&
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
            Height          =   690
            Left            =   120
            Picture         =   "PEDCOT07.frx":009D
            Style           =   1  'Graphical
            TabIndex        =   40
            ToolTipText     =   "Otros Accesos Directos"
            Top             =   3840
            Width           =   650
         End
         Begin VB.CommandButton Command6 
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
            Height          =   690
            Left            =   120
            Picture         =   "PEDCOT07.frx":03A7
            Style           =   1  'Graphical
            TabIndex        =   39
            ToolTipText     =   "Acceso a Carga de Cotización"
            Top             =   6075
            Width           =   650
         End
         Begin VB.CommandButton AGenRep 
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
            Height          =   690
            Left            =   120
            Picture         =   "PEDCOT07.frx":04F1
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
            Top             =   5330
            Width           =   650
         End
         Begin VB.CommandButton Command4 
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
            Height          =   690
            Left            =   120
            Picture         =   "PEDCOT07.frx":07FB
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
            Top             =   4560
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
            Picture         =   "PEDCOT07.frx":0B05
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   790
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
            Picture         =   "PEDCOT07.frx":0E0F
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Acceso a Base de Datos de Proveedores"
            Top             =   1550
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
            Picture         =   "PEDCOT07.frx":1119
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Acceso a Maestro de Artículos"
            Top             =   2240
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
            Picture         =   "PEDCOT07.frx":1423
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Configuración de Funcionamiento de Pedido de Cotización"
            Top             =   3085
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
            Picture         =   "PEDCOT07.frx":1865
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   100
            Width           =   650
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
            Height          =   690
            Left            =   120
            Picture         =   "PEDCOT07.frx":19AF
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Aprueba e Imprime Pedido de Cotización"
            Top             =   7560
            Width           =   650
         End
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Left            =   105
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   2
            Top             =   6840
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   500
               Left            =   0
               TabIndex        =   3
               Top             =   105
               Width           =   12000
            End
         End
         Begin VB.Image Image1 
            Height          =   390
            Left            =   -315
            Picture         =   "PEDCOT07.frx":1CB9
            Top             =   7080
            Width           =   1515
         End
      End
      Begin VB.Line Line7 
         X1              =   165
         X2              =   10740
         Y1              =   5640
         Y2              =   5640
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Por el Presente solicitamos la Cotización de los siguientes Items:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   240
         TabIndex        =   31
         Top             =   2280
         Width           =   10455
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PC"
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
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C000&
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   630
         Shape           =   4  'Rounded Rectangle
         Top             =   375
         Width           =   540
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000C000&
         FillColor       =   &H00FFFF80&
         FillStyle       =   0  'Solid
         Height          =   390
         Left            =   525
         Shape           =   4  'Rounded Rectangle
         Top             =   315
         Width           =   750
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000C0C0&
         FillColor       =   &H00C0E0FF&
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   420
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   43
      Top             =   1320
      Width           =   855
   End
End
Attribute VB_Name = "PEDCOT07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Sub PRISOLORF(NUMESOLI&, NORFA$)

   FORIPRE$ = TRIM$(CONTROL$("", "SOLIMATE")) ' FALTA DESIGNAR DONDE SE VA A CAMBIAR
   '
   If FORIPRE$ = "" Then
        Call MENSERR(24, "Falta Definir Formulario Solicitud por O.Fabricaciòn")
        Exit Sub
   End If
   '
   TIPODOC$ = "Solicitud de Materiales por O.F" ' guardar el documento immpreso
   NUMEDOC$ = TRIM$(Str$(NUMESOLI&))
   '
   HOII% = HOY(HO$)
   FEC = FECHATEX$(HOII%)

   CODPARAM$(1) = "REF-MAT1": DOCPARAM$(1) = NORFA$ 'CODIGO
   CODPARAM$(2) = "NUME-COM": DOCPARAM$(2) = NUMESOLI&  'o.f
   CODPARAM$(3) = "FECH-EMI": DOCPARAM$(3) = FEC
   CAPARDOC% = 3
   CAITAB1% = 0
   '
   CODTABU1$(1) = "COD-MPRI"
   CODTABU1$(2) = "DES-MPRI"
   CODTABU1$(3) = "UNI-MPRI"
   CODTABU1$(4) = "CAN-MPRI"
   CACOLTAB1% = 4
   
   For i& = 1 To ULTREG&("!MATEROF")
       x$ = REGLEIDO$("!MATEROF", i&)
       CI% = CVI(Mid$(x$, 1, 2))
       CANTU = CVS(Mid$(x$, 5, 4))
       CAITAB1% = CAITAB1% + 1
       DECRIT$ = TRIM$(DESCRIT0$(CI%))
       '
       TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
       TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
       TETABU1$(3, CAITAB1%) = Unimed$(CI%)
       TETABU1$(4, CAITAB1%) = FORMATNUM$(CANTU, "F9.1")
   Next i&
   Call PRINTDOC("@" + FORIPRE$)   ' Pedido de Cotización

End Sub

Sub SIMXARTICULO()
    
    'SELECCION DE ARTICULOS CON RESERVAS
    CONDI$ = "CANTRES > 0.05 AND (NUME_SOLINT IS NULL OR NUME_SOLINT < 1) ORDER BY COD_INTE"
10  Call CARGALISEL("NUSOLART", "RESERVA", "COD_INTE/C16;DESCRIP/A64", CONDI$, "DISTINCT")
    Call SELECHO("NUSOLART/Lista de Selección de Artículos Reservados")
    CI% = CODINT(VALACT1$("NUSOLART"))
    If CI% = 0 Then Exit Sub '
    If STATITEM%(CI%) < 1 Then
      Call COMUNI("Código: " & CODEXT$(CI%) & " Dado de Baja")
      GoTo 10
    End If
    RESEAR09.Label7 = CODEXT$(CI%)
    RESEAR09.Show 1
    GoTo 10
    '
    
End Sub

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Command1_Click()

    Command1.Enabled = False
    '
    Call GENISXOF
    '
    If ULTREG&("INDISOLI") < 1 Then
       Call COMUNI("No Existen Solicitudes Sin Cotizar")
       GoTo 99
    End If

    Call SELECHO("INDISOLI/Indice de Solicitudes")
    '
    NUSOL& = XVALO(VALACT1$("INDISOLI"))
    If NUSOL& < 1 Then GoTo 99
    '
    SQLX$ = "SELECT * FROM DETASOLINT WITH (NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NUME_SOLINT = " & NUSOL&
    SQLX$ = SQLX$ + " and (COTI_SOLINT < 1 OR COTI_SOLINT IS NULL)"
    Dim DETATMP As ADODB.Recordset
    Set DETATMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    List1.Clear
    With DETATMP
      Do While Not .EOF
         CI% = XVALO(!COD_INTE)
         CAN1 = XVALO(!CANT_SOLINT)
         DESC2$ = DESCRIT0(CI%)
         CAN2$ = FORMATNUM$(CAN1, "F9.1")
         NORFA$ = SAFETEXT$(!DESTI_SOLINT) ' LO TOMO DE ACA POR QUE EL MISMO SE REPITE PARA TODOS LOS ITEMS
        'AGREGO AL LIST
        L1$ = Space(128)
        Call REPLA(L1$, CODEXT$(CI%), 1, 16)
        Call REPLA(L1$, DESC2$, 18, 35)
        Call REPLA(L1$, CAN2$, 54, 10)
        List1.AddItem L1$
      .MoveNext
      Loop
      
    End With
    
    Label8 = NUSOL&
    Text2 = NORFA$ ' PONGO EN DESTINO EL NUMERO DE OF
    DETATMP.Close
    Set DETATMP = Nothing
99  Command1.Enabled = True
'
End Sub

Sub GENISXOF()
    '
    Call ABRECOMPRAS
    CONDI$ = "ENCASOLINT.NUME_SOLINT > 0  AND (ENCASOLINT.STATUS IS NULL OR ENCASOLINT.STATUS < 9) AND  (DETASOLINT.COTI_SOLINT < 1 OR DETASOLINT.COTI_SOLINT IS NULL) ORDER BY DETASOLINT.NUME_SOLINT ASC"
    Call CARGALISEL("INDISOLI", "DETASOLINT+ENCASOLINT/ENCASOLINT.NUME_SOLINT = DETASOLINT.NUME_SOLINT", "DETASOLINT.NUME_SOLINT/F8.0;ENCASOLINT.DESCRI_SOLINT/A54", CONDI$, "DISTINCT")
    Call FRESHECHO("INDISOLI")
    '
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG07/AMAPRO"
   Call CONECRUN(ACONEC$, "Padron Maestro de Proveedores")
   Command11.Enabled = True
End Sub

Private Sub Command14_Click()

    Command14.Enabled = False

    If List4.ListCount <= 0 Then
      Call MENSERR(24, "Debe Elegir Al Menos un Proveedor")
      GoTo 99
    End If
  '
    FECX% = FECHANUM(Text9)
    FF% = HOY(HO$) ' Validacion de que la fecha no sea menor a la del dia
    If FECX% < FF% Then Call MENSERR(24, "La Fecha no Puede ser Menor a la de Hoy"): GoTo 99
  
  '
    If TRIM$(Text2) = "" Then
      Call MENSERR(24, "Falta Ingresar el Destino")
      GoTo 99
    End If
  '
    If List1.ListCount <= 0 Then
      Call MENSERR(24, "Debe Elegir Al Menos un Artículo")
      GoTo 99
    End If
  '
    Screen.MousePointer = 11

    Call ABRECOMPRAS
    '
    CONDI$ = "NUME_SOLINT > 0 "
    NUMSOL& = 1 + XVALO(VALOBADA("ENCASOLINT", "MAX(NUME_SOLINT)", CONDI$, "NOMESS"))
    
    If XVALO(Label8) > 0 Then ' MARCA SI VIENE DE UNA SOLICITUD -- 0  MANUAL Y 1 SOLICITUD x of
       
       For i& = 0 To List1.ListCount - 1
        XX$ = List1.List(i&)
        CODI0$ = Mid$(XX$, 1, 16)
        CI% = CODINT%(CODI0$)
        CAN1 = XVALO(Mid$(XX$, 54, 10))
        FES1 = TRIM$(Mid$(XX$, 65, 10))
        HOY (HOS$)
        If FES1 = "" Then FES1 = HOS$ ' SI NO SE CARGA FECHA SE CARGA LA DEL DIA
        '
        CONDI2$ = "NUME_SOLINT = " & XVALO(Label8) & "  AND COD_INTE = " & CI% & "  "
        SQL1$ = "SELECT FENTRE_SOL, COTI_SOLINT FROM DETASOLINT WHERE " & CONDI2$
        Dim RSTDETASOLINT As ADODB.Recordset
        Set RSTDETASOLINT = New ADODB.Recordset
        RSTDETASOLINT.Open FILTSQL$(SQL1$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With RSTDETASOLINT
            Do While Not .EOF
                !FENTRE_SOL = FES1
                !COTI_SOLINT = Label6
                .Update
                
                .MoveNext
            Loop
        End With
        RSTDETASOLINT.Close
        Set RSTDETASOLINT = Nothing
        
        'Call ACTUREGISTRO("DETASOLINT", "FENTRE_SOL", CONDI2$, FES1, REGAF&, "NOMESS")
        'Call ACTUREGISTRO("DETASOLINT", "COTI_SOLINT", CONDI2$, Label6, REGAF&, "NOMESS")
      '
     Next i&
       'GRABO EL SECTOR EN ENCASOLINT (GRABO FUERA DEL FOR POR QUE ES EN TABLA ENCASOLINT NO DE DETALLE)
        CONDI$ = "NUME_SOLINT = " & XVALO(Label8)
        SQL1$ = "SELECT SECT_SOLINT FROM ENCASOLINT WHERE " & CONDI$
        Dim RSTENCASOLINT As ADODB.Recordset
        Set RSTENCASOLINT = New ADODB.Recordset
        RSTENCASOLINT.Open FILTSQL$(SQL1$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With RSTENCASOLINT
            If Not (.EOF And .BOF) Then
                !Sect_Solint = Left$(TRIM$(Text2), 12)
                .Update
            End If
        End With
        RSTENCASOLINT.Close
        Set RSTENCASOLINT = Nothing
       
        'Call ACTUREGISTRO("ENCASOLINT", "SECT_SOLINT", CONDI$, Left$(TRIM$(Text2), 12), REGAF&, "NOMESS")
       Label8 = 0
       GoTo 50
    End If

    SQLX$ = "SELECT * FROM ENCASOLINT "  'GRABACION EN TABLA ENCASOLINT
    SQLX$ = SQLX$ + " WHERE STATUS = 10" ' CONDICION  INVALIDA FORZADA
    Dim ENCASTMP As ADODB.Recordset
    Set ENCASTMP = New ADODB.Recordset
    ENCASTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With ENCASTMP
       .AddNew   ' YA CONVERTIDO
       On Error Resume Next
       !CodiEmpr = CodiEmp%
       On Error GoTo 0
       !NUME_SOLINT = NUMSOL&
       !Sect_Solint = Left$(TRIM$(Text2), 12)
       !FEMI_SOLINT = Text9
       !Usua_Gen = USNBR% Mod 100
       !Obse_Solint = Text13
       !DESCRI_SOLINT = Text9 & "-" & TRIM$(Text2)
       .Update
    End With
    ENCASTMP.Close
    Set ENCASTMP = Nothing
    '
    SQLX$ = "SELECT * FROM DETASOLINT"  'GRABACION EN TABLA DETASOLINT
    SQLX$ = SQLX$ + " WHERE STATUS = 10" ' CONDICION  INVALIDA FORZADA
    Dim DETASTMP As ADODB.Recordset
    Set DETASTMP = New ADODB.Recordset
    DETASTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With DETASTMP
     For JX% = 0 To List1.ListCount - 1
      L1$ = List1.List(JX%)
      CODI0$ = Mid$(L1$, 1, 16)
      CI% = CODINT%(CODI0$)
      CAN1 = XVALO(Mid$(L1$, 54, 10))
      FES1 = TRIM$(Mid$(L1$, 65, 10))
      HOY (HOS$)
      If FES1 = "" Then FES1 = HOS$ ' SI NO SE CARGA FECHA SE CARGA LA DEL DIA
      
      OBS1$ = Mid$(L1$, 76, 25)
      '
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
      !NUME_SOLINT = NUMSOL&
      !COTI_SOLINT = XVALO(Label6)
      !NUORD_ITEM = JX% + 1
      !COD_INTE = CI%
      !CANT_SOLINT = XVALO(CAN1)
      !UNI_MAS = Unimed$(CI%)
      !FENTRE_SOL = FES1
      !DESCRITEM = DESCRIT$(CI%)
      !DESTI_SOLINT = Left$(TRIM$(Text2), 12) ' AGREGADO POR QUE NO ESTABA GRABANDOSE EL DESTINO
      .Update
      '
     Next JX%
    End With
    DETASTMP.Close
    Set DETASTMP = Nothing
    '
50  Call GENINSOL   ' GENERA INDICE
    '
    'IMPRIME COTIZACION
    '
    FORIPRE$ = CONTROL$("", "FORMPECO")
    '
    Call COPYSTRU("INSOLINT", "INDICOT")

    If TRIM$(FORIPRE$) <> "" Then
    '
     For JX% = 0 To List4.ListCount - 1 'UNO POR CADA PROVEEDOR
      'INFORMACION PARA FORMAR EL STRING DE VARIABLE TIPODOC
      L1$ = List4.List(JX%)
      NPRO% = Val(Mid$(L1$, 1, 5))
      FECHDOC$ = Text9.TEXT
      NPRO% = Val(Mid$(L1$, 1, 5))
      PROVE$ = rasocli((-1) * NPRO%)
      PROV$ = AJUSTI$(PROVE$, 18)
      TELEF$ = TRIM$(TEL1CLI$((-1) * NPRO%))
      NUMEDOC1$ = TRIM$(Label6)
      '
      TIPODOC$ = "Pedido de Cotizacion: " & NUMEDOC1$ & " - " & PROV$ & " - " & FECHDOC$
      '
      'GENERO EL INDICE DE COTIZACIONES
      '
      XXY$ = REGBLAN$("INDICOT")
      REFER$ = TIPODOC$
      NUMED& = XVALO(Label6)
      Call REPLA(XXY$, MKS(NUMED&), 1, 4)
         REFER$ = "Cotizacion: " & NUMEDOC1$ & "-" & PROV$ & "-" & FECHDOC$
      Call REPLA(XXY$, REFER$, 5, 44)
      Call REGAPP("INDICOT", XXY$)
      '
      'IMPRESION
      '
      CODPARAM$(1) = "USUARIO"
      CODPARAM$(2) = "REF-MAT2"
      CODPARAM$(3) = "NOP-OPER"
      CODPARAM$(4) = "TELE-CLI"
      CAPARDOC% = 4
      '
      DOCPARAM$(1) = USERNAME$
      DESTI$ = ECOARCH$("SECPROD", Text2)
      DOCPARAM$(2) = TRIM$(Text2) & "  " & TRIM$(DESTI$)
      DOCPARAM$(3) = TRIM$(Label6)
      DOCPARAM$(4) = TELEF$
      '
      CODTABU1$(1) = "COD-MPRI"
      CODTABU1$(2) = "DES-MPRI"
      CODTABU1$(3) = "CAN-MPRI"
      CODTABU1$(4) = "FECH-REC"
      CODTABU1$(5) = "PAR-MPRI"
      CACOLTAB1% = 5
      '
      DESTIDOC% = NPRO% * (-1)
      CAITAB1% = 0
      CAITCRU% = 0
      '
      NC1% = DESTIDOC%
      For J1% = 0 To List1.ListCount - 1
         x$ = List1.List(J1%)
         CI0% = CODINT(Mid$(x$, 1, 16))
         CANTU = Val(Mid$(x$, 54, 10))
         FEX = FECHANUM(TRIM$(Mid$(x$, 65, 10)))
         OBS1$ = Mid$(x$, 76, 25)
         CAITAB1% = CAITAB1% + 1
         CODDX$ = TRIM$(CODEXT$(CI0%))
         DECRIT$ = TRIM$(DESCRIT0$(CI0%))
         '
         TETABU1$(1, CAITAB1%) = CODDX$
         TETABU1$(2, CAITAB1%) = DECRIT$
         TETABU1$(3, CAITAB1%) = TRIM$(Str$(CANTU))
         TETABU1$(4, CAITAB1%) = FECHATEX$(FEX)
         TETABU1$(5, CAITAB1%) = OBS1$
      Next J1%
      '
      CODTABU2$(1) = "MEDIDA-2" 'DESCRIPCION
      CACOLTAB2% = 1
      '
      CAITAB2% = 0
      DESGRAL$ = Text13.TEXT 'OBSERVACIONES
      Call FRATEXTO(DESGRAL$, 92)
      For U& = 1 To CARETEX%
        Z$ = "." + RETEX$(U&)
        CAITAB2% = CAITAB2% + 1
        TETABU2$(1, CAITAB2%) = Z$
      Next U&
      '
      Call PRINTDOC("FORMPECO")   ' Pedido de Cotización
      '
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
      For KK1% = 1 To CACOLTAB2%
        For KK2% = 1 To CAITAB2%
          TETABU2$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
     Next JX%
     '
     Call FRESHECHO("INDICOT")
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     '
   End If
   Call BLANFORMU
   '
   Call COMUNI("Solicitud Generada Exitosamente")
   '
99 Command14.Enabled = True
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call FINAL("")
End Sub

Private Sub Command27_Click()
    '
    Command27.Enabled = False
    '
5   RTA1% = ALTERNA%("Generar SIM por O/F.\Generar SIM por Artículo\SIM Manual\Anulación de Solicitud\Actualizar Lista de SIM")
    If RTA1% < 1 Or RTA1% > 5 Then GoTo 99
    '
    If RTA1% = 2 Then
      Call SIMXARTICULO ' GENERAR SIM POR ARTÍCULO
      GoTo 5
    ElseIf RTA1% = 3 Then 'MANUAL
      Call GENERASIMANUAL
      GoTo 5
    ElseIf RTA1% = 4 Then
      Call ANUSOLICITUD
      GoTo 5
    ElseIf RTA1% = 5 Then
      Call GENINSOL
      GoTo 5
    ElseIf RTA1% = 1 Then 'Generar Solicitud
    
10      Call SELECHO("ECORDEP/Indice de Ordenes de Fabricación Pendientes")
        NORFA$ = TRIM$(VALACT1$("ECORDEP"))
        If NORFA$ = "" Then GoTo 99
        JJ& = 0
        Call SETULTREG("!MATEROF", 0)
        SQLX$ = " SELECT * FROM RESERVA WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NORFA$ & "'"
        SQLX$ = SQLX$ + " and (NUME_SOLINT IS NULL OR NUME_SOLINT < 1) "
        SQLX$ = SQLX$ + " ORDER BY COD_INTE "
    '
        Call ABRESTOCKS
        CANTI = 0
        Dim RESERTMP As ADODB.Recordset
        Set RESERTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With RESERTMP
          Do While Not .EOF
            CI% = XVALO(!COD_INTE)
            CANTI = XVALO(!CANTRES)
            '
            If CI% > 0 Then
             If CANTI > FRACMIN(CI%) Then
                JJ& = JJ& + 1
                Call REPLA(TX$, MKI$(CI%), 1, 2)
                Call REPLA(TX$, MKS$(CANTI), 5, 4)
                Call GRAREG("!MATEROF", TX$, JJ&)
              End If
            End If
            .MoveNext
          Loop
        End With
        RESERTMP.Close
        Set RESERTMP = Nothing
        If JJ& < 1 Then
           Call COMUNI("Ya se Realizaron Solicitudes para la Orden de Fabricación Elegida")
           GoTo 10
        End If
        '
        Call SETULTREG("!MATEROF", JJ&) 'MUESTRO VENTANA
        
        VTB% = VENTABU%("SOLIMARE/TITUPAN=Solicitud de Material Por Orden de Fabricación" + TITU$)
        If VTB% < 0 Then GoTo 10
        '
        Call ABRECOMPRAS
        CONDI$ = "NUME_SOLINT > 0 "
        NUMSOL& = 1 + XVALO(VALOBADA("ENCASOLINT", "MAX(NUME_SOLINT)", CONDI$, "NOMESS"))
        HOII% = HOY(HO$)

        '
        Call CreaRegistro("ENCASOLINT", "NUME_SOLINT", NUMSOL&)
        REFE$ = "Solicitud de Materiales por O.F " & NORFA$
        CONDI1$ = "NUME_SOLINT = " & NUMSOL&
        Call ACTUREGISTRO("ENCASOLINT", "FEMI_SOLINT", CONDI1$, CVDAT(HOII%), REGAF&)
        Call ACTUREGISTRO("ENCASOLINT", "DESCRI_SOLINT", CONDI1$, REFE$, REGAF&)
        
        '
        SQLX$ = "SELECT * FROM DETASOLINT"  'GRABACION EN TABLA DETASOLINT
        SQLX$ = SQLX$ + " WHERE STATUS = 10" ' CONDICION  INVALIDA FORZADA
        Dim DETASTMP As ADODB.Recordset
        Set DETASTMP = New ADODB.Recordset
        DETASTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With DETASTMP
         For JX& = 1 To ULTREG&("!MATEROF")
          L1$ = REGLEIDO$("!MATEROF", JX&)
          CI% = CVI(Mid$(L1$, 1, 2))
          CANTI = CVS(Mid$(L1$, 5, 4))
          '
          .AddNew   ' YA CONVERTIDO
          On Error Resume Next
          !CodiEmpr = CodiEmp%
          On Error GoTo 0
          !NUME_SOLINT = NUMSOL&
          !NUORD_ITEM = JX&
          !COD_INTE = CI%
          !CANT_SOLINT = XVALO(CANTI)
          !UNI_MAS = Unimed$(CI%)
          !FENTRE_SOL = HO$
          !DESCRITEM = DESCRIT$(CI%)
          !DESTI_SOLINT = NORFA$
          .Update
          CONDI$ = "IDSUBOR = '" & NORFA$ & "' AND COD_INTE = " & CI%
          Call ACTUREGISTRO("RESERVA", "NUME_SOLINT", CONDI$, NUMSOL&, REGAF&, "NOMESS")
          '
         Next JX&
        End With
        DETASTMP.Close
        Set DETASTMP = Nothing
        'Call GENINSOL
        Call PRISOLORF(NUMSOL&, NORFA$)
        Call CIERRARCH("!MATEROF")
        Call BLANFORMU

    End If
    

99  Command27.Enabled = True
    
End Sub
Sub GENERASIMANUAL()

   Call SETULTREG("!MATEROF", JJ&) 'MUESTRO VENTANA
 
   VTB% = VENTABU%("SOLIMANU/TITUPAN=Solicitud de Material Manual")
   If VTB% < 0 Then Exit Sub
   '
   'RECORRO
   FIN& = ULTREG&("!MATEROF")
   ORD% = 0
   For J& = 1 To FIN&
        x$ = REGLEIDO$("!MATEROF", J&)
        CI% = CVI(Mid$(x$, 1, 2))
        CANTI = CVS(Mid$(x$, 5, 4))
        If CI% > 0 Then
            If CANTI > FRACMIN(CI%) Then
              ORD% = ORD% + 1
              If ORD% = 1 Then
                Call GRAENCASOMAN(NUMSOL&) ' vuelve con el numero de solicitud
              End If
              '
              Call GRABAMANUAL(NUMSOL&, CI%, CANTI, ORD%)
            End If
        End If
   Next J&
   '
   HOII% = HOY(HO$)
   FILTX$ = TRIM$(Str$(NUMSOL&)) & ";" & TRIM$(Str$(HOII%))
   Call STDFORM("SOL-MAN", FILTX$, "PRINT")
 
   Call GENISXOF ' ACTUALIZA EL INDICE
   '
   Call COMUNI("Solicitud Generada")
 
 
End Sub
Sub GRABAMANUAL(NUMESOL&, CI%, CAN, NUORD%)
 
    Dim RSTDETA As ADODB.Recordset
    Set RSTDETA = New ADODB.Recordset
    SQLX$ = "SELECT * FROM DETASOLINT WHERE Status < 0"
    RSTDETA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    HOII% = HOY(HO$)
    With RSTDETA
        .AddNew
        !NUME_SOLINT = NUMESOL&
        !NUORD_ITEM = NUORD%
        !COD_INTE = CI%
        !CANT_SOLINT = CAN
        !UNI_MAS = Unimed$(CI%)
        !FENTRE_SOL = FETEXCOR1$(HOII%)
        !DESCRITEM = DESCRIT$(CI%)
        .Update
    End With
    RSTDETA.Close
    Set RSTDETA = Nothing
    
    Call GENINSOL   ' GENERA INDICE

'   'Grabo en Detasolint
'10 RNDX& = RANDSTRING$(-8)
'   If CantRegistros("DETASOLINT", "NUME_SOLINT = '" & RNDX& & "'", "NOMESS") > 0 Then GoTo 10
'
'   Call CreaRegistro("DETASOLINT", "NUME_SOLINT", RNDX&)
'   '
'   CONDI$ = "NUME_SOLINT = " & RNDX&
'   HOII% = HOY(HO$)
'   '
'   Call ACTUREGISTRO("DETASOLINT", "NUORD_ITEM", CONDI$, NUORD%, REGAF&)
'   Call ACTUREGISTRO("DETASOLINT", "COD_INTE", CONDI$, CI%, REGAF&)
'   Call ACTUREGISTRO("DETASOLINT", "CANT_SOLINT", CONDI$, CAN, REGAF&)
'   Call ACTUREGISTRO("DETASOLINT", "Uni_Mas", CONDI$, UNIMED$(CI%), REGAF&)
'   Call ACTUREGISTRO("DETASOLINT", "FENTRE_SOL", CONDI$, FETEXCOR1$(HOII%), REGAF&)
'   Call ACTUREGISTRO("DETASOLINT", "DESCRITEM", CONDI$, DESCRIT$(CI%), REGAF&)
'   '
'   'Este acturegistro siempre ultimo
'   Call ACTUREGISTRO("DETASOLINT", "NUME_SOLINT", CONDI$, NUMESOL&, REGAF&)
'   '
End Sub
 
 
 
Sub GRAENCASOMAN(NUMSOL&)
  '
  'GRABA ENCABEZADO DE ENCASOLINT POR CARGA MANUAL
 
    CONDI$ = "NUME_SOLINT > 0 "
    NUMSOL& = 1 + XVALO(VALOBADA("ENCASOLINT", "MAX(NUME_SOLINT)", CONDI$, "NOMESS"))
    
    Dim RSTDETA As ADODB.Recordset
    Set RSTDETA = New ADODB.Recordset
    SQLX$ = "SELECT * FROM ENCASOLINT WHERE NUME_SOLINT < 1"
    RSTDETA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    HOII% = HOY(HO$)
    REFE$ = "Solicitud de Materiales Carga Manual "
    With RSTDETA
        .AddNew
        !NUME_SOLINT = NUMSOL&
        !FEMI_SOLINT = CVDAT(HOII%)
        !DESCRI_SOLINT = REFE$
        .Update
    End With
    RSTDETA.Close
    Set RSTDETA = Nothing
 
'   Call ABRECOMPRAS
'   CONDI$ = "NUME_SOLINT > 0 "
'   NUMSOL& = 1 + XVALO(VALOBADA("ENCASOLINT", "MAX(NUME_SOLINT)", CONDI$, "NOMESS"))
'   HOII% = HOY(HO$)
'   '
'   Call CreaRegistro("ENCASOLINT", "NUME_SOLINT", NUMSOL&)
'   REFE$ = "Solicitud de Materiales Carga Manual "
'   CONDI$ = "NUME_SOLINT = " & NUMSOL&
'   Call ACTUREGISTRO("ENCASOLINT", "FEMI_SOLINT", CONDI$, CVDAT(HOII%), REGAF&)
'   Call ACTUREGISTRO("ENCASOLINT", "DESCRI_SOLINT", CONDI$, REFE$, REGAF&)
   '
End Sub
 

        

Sub ANUSOLICITUD()
     '
    Call GENISXOF
    '
    If ULTREG&("INDISOLI") < 1 Then
       Call COMUNI("No Existen Solicitudes Para Anular")
       Exit Sub
    End If
    
10  Call SELECHO("INDISOLI/Indice de Solicitudes")
    '
    NUMSOL& = XVALO(VALACT1$("INDISOLI"))
    REFE$ = TRIM$(VALACT2$("INDISOLI"))
    If NUMSOL& < 1 Then Exit Sub
    CONDICION$ = "NUME_SOLINT =  " & NUMSOL& & " AND COTI_SOLINT > 0"
    If CantRegistros("DETASOLINT", CONDICION$) > 0 Then
        Call COMUNI("Imposible Anular Esta Solicitud\Cotizaciones Parciales Realizadas")
        Exit Sub
     End If

    OPX% = COMALTER%("Confirmar Anulación Solicitud Nº: " & NUMSOL& & " \\Cancelar\Confirmar Anulación")
    If OPX% < 2 Then GoTo 10
    If OPX% = 2 Then
        CONDI$ = "NUME_SOLINT = " & NUMSOL&
        Call ACTUREGISTRO("RESERVA", "NUME_SOLINT", CONDI$, 0, REGAF&, "NOMESS")
        Call ACTUREGISTRO("ENCASOLINT", "STATUS", CONDI$, 9, REGAF&, "NOMESS")
        
        Call ACTUREGISTRO("DETASOLINT", "STATUS", CONDI$, 9, REGAF&, "NOMESS")
        Call COMUNI("Solicitud Nº : " & NUMSOL& & " \ " & REFE$ & "\ Anulada")
    End If
     
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call HELPONLINE("PEDICOTI")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    '
4   RTA% = ALTERNA%("Solicitudes Internas de Materiales\Solicitud de Cotizaciones")
    If RTA% < 1 Or RTA% > 2 Then GoTo 99
    If RTA% = 1 Then
       Call VERYREIMPSIM
       GoTo 4
    End If
    '
10  If ULTREG&("INDICOT") > 0 Then ' SI TIENE INFORMACION MUESTRA EL INDICE
     Call SELECHO("INDICOT/Indice General de Concurso de Precios")
    Else
     GoTo 99
    End If
    '
    NUDOCU$ = VALACT1$("INDICOT") ' ASIGNO EL NUMERO DEL FORMULARIO IMPRESO  SELECCIONAD0
    TIDOCUM$ = Left$(VALACT2("INDICOT"), 11)
    pos% = InStr(VALACT2("INDICOT"), "-") ' POSICION DEL PRIMER GUION ANTES DEL VENDED
    PROV$ = Mid$(VALACT2("INDICOT"), pos% + 1, 18)
    If NUDOCU$ = "" Then GoTo 99 ' por si no elige ninguno y cierra el indice
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
             PPX1% = InStr(XX$, PROV$) ' VERIFICA QUE TAMBIEN COINCIDA EL PROVEEDOR
             If PPX1% > 0 Then
               DOCUNU& = CVS(Left$(XX$, 4))
               GoTo 5
             End If
           End If
         End If
    Next UI&
    '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a Esta Solicitud de Material.\ ")
    GoTo 10
       '
5   Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
   
    GoTo 10

99  Command4.Enabled = True


End Sub

Private Sub Command5_Click()
  '
  OPPCMAQ07.Show 1
  '
End Sub

Private Sub Command6_Click()
    CARCOTI07.Show 1
End Sub

Private Sub Command7_Click()
  
    Command7.Enabled = False
    '
    Call SELECHO("SECPROD")
    If XVALO(VALACT1$("SECPROD")) < 1 Then GoTo 99
    '
    Text2 = VALACT1$("SECPROD")
99  Command7.Enabled = True

End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    Call CIERRARCH("*.*")
    ACONEC$ = "AMASTO07"
    If EXISTE%("AMASTO07.EXE") > 0 Then ACONEC$ = "AMASTO07"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text9.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text9.TEXT)
End Sub

Private Sub Form_Load()
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   Call CENTRAFORM(Me)
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   FECHA1$ = Format(Now, "yymmdd")
   ID1$ = CONTROL$("", "IDENPECO")
   If ID1$ = "" Then ID1$ = "PC"
   'Text6.TEXT = ID1$ + "-" + FECHA1$
   '
   HOII% = HOY(HOS$)
   Text9.TEXT = HOS$
   Label7 = Left$(USERNAME, 20)
   Call ABRECOMPRAS
   'Label6 = 1 + XVALO(VALOBADA("ENCASOLINT", "MAX(NUME_SOLINT)", "NUME_SOLINT > 0"))
   '
   Label6 = 1 + XVALO(VALOBADA("DETASOLINT", "MAX(COTI_SOLINT)", "COTI_SOLINT > 0"))  'Numero de Solicitud
  '
   Label8 = 0 ' MARCA PARA SABER SI VIENE DE ORDEN DE FABRICACION O MANULA
End Sub

Sub BLANFORMU()
  '
  FECHA1$ = Format(Now, "yymmdd")
  ID1$ = CONTROL$("", "IDENPECO")
  If ID1$ = "" Then ID1$ = "PC"
  'Text6.TEXT = ID1$ + "-" + FECHA1$
  'incluir el proximo numero de solicitud
  Label6 = 1 + XVALO(VALOBADA("DETASOLINT", "MAX(COTI_SOLINT)", "NUME_SOLINT > 0"))
  '
  HOII% = HOY(HOS$)
  Text9.TEXT = HOS$
  '
  Text2 = "" ' DESTINO
  Text13.TEXT = ""
  List1.Clear
  List4.Clear
  Label7 = ""
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("")
    Call FINAL("")
End Sub

Private Sub Label22_DBLClick()
    '
    Call BLANARCH("!LIPECOTI")
    JJ& = 0
    For JX% = 0 To List1.ListCount - 1
      L1$ = List1.List(JX%)
      CODI0$ = Mid$(L1$, 1, 16)
        CODI1% = CODINT%(CODI0$)
      DESC1$ = Mid$(L1$, 18, 35)
      CAN1 = Val(Mid$(L1$, 54, 10))
      FES1$ = TRIM$(Mid$(L1$, 65, 10))
      OBS1$ = Mid$(L1$, 76, 25)
      '
      YY$ = REGBLAN$("LIPECOTI")
      Call REPLA(YY$, MKI$(CODI1%), 1, 2)
      Call REPLA(YY$, MKS$(CAN1), 3, 4)
      Call REPLA(YY$, MKI$(FECHANUM(FES1$)), 7, 2)
      Call REPLA(YY$, OBS1$, 9, 25)
      JJ& = JJ& + 1
      Call GRAREG("!LIPECOTI", YY$, JJ&)
      '
    Next JX%
    Call SETULTREG("!LIPECOTI", JJ&)
    Call CIERRARCH("!LIPECOTI")
    '
    PRF$ = "LIPECOTI"
    If Label8 > 0 Then PRF$ = "LIPECOT1" ' SI ES POR SIM NO EDITABLE
    ATRI$ = ATRI$ + "/TITUPAN=Listado de Materiales a Cotizar /"
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + PEDCOT07.Left + 170))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + PEDCOT07.Top + 2250))
    VTB% = VENTABU%(PRF$ + ATRI$)
    If VTB% >= 0 Then
      '
      List1.Clear
      '
      FIN& = ULTREG&("!LIPECOTI")
      For J& = 1 To FIN&
        x$ = REGLEIDO$("!LIPECOTI", J&)
        CODI1% = CVI(Mid$(x$, 1, 2))
        CODI2$ = CODEXT$(CODI1%)
        DESC2$ = DESCRIT0(CODI1%)
        CAN1 = CVS(Mid$(x$, 3, 4))
        CAN2$ = FORMATNUM$(CAN1, "F9.1")
        FES2% = CVI(Mid$(x$, 7, 2))
        OBS2$ = Mid$(x$, 9, 55)
        'AGREGO AL LIST
        L1$ = Space(128)
        Call REPLA(L1$, CODI2$, 1, 16)
        Call REPLA(L1$, DESC2$, 18, 35)
        Call REPLA(L1$, CAN2$, 54, 10)
        Call REPLA(L1$, FECHATEX$(FES2%), 65, 10)
        Call REPLA(L1$, OBS2$, 76, 25)
        List1.AddItem L1$
      Next J&
      '
    End If
End Sub

Private Sub Label92_dblClick()
  '
  Call BLANARCH("!INVITA2")
  JJ& = 0
  For JX% = 0 To List4.ListCount - 1
    L1$ = List4.List(JX%)
    NPRO% = Mid$(L1$, 1, 5)
    '
    YY$ = REGBLAN$("INVITA2")
    Call REPLA(YY$, MKI$(NPRO%), 1, 2)
    JJ& = JJ& + 1
    Call GRAREG("!INVITA2", YY$, JJ&)
    '
  Next JX%
  Call SETULTREG("!INVITA2", JJ&)
  Call CIERRARCH("!INVITA2")
  '
  PRF$ = "INVITA2"
  ATRI$ = ATRI$ + "/TITUPAN=Lista de Proveedores Invitados /"
  ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + PEDCOT07.Left + 1000))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + PEDCOT07.Top + 5))
  VTB% = VENTABU%(PRF$ + ATRI$)
  If VTB% >= 0 Then
    '
    List4.Clear
    '
    FIN& = ULTREG&("!INVITA2")
    For J& = 1 To FIN&
      x$ = REGLEIDO$("!INVITA2", J&)
      NPRO1% = CVI(Mid$(x$, 1, 2))
        NPRO2$ = FORMATNUM$(NPRO1%, "F5.0")
      PRO1$ = ECOARCH$("PROVED1", MKI$(NPRO1%))
      'AGREGO AL LIST
      L1$ = Space(128)
      Call REPLA(L1$, NPRO2$, 1, 5)
      Call REPLA(L1$, PRO1$, 8, 20)
      List4.AddItem L1$
    Next J&
    '
  End If
  '
End Sub

Private Sub List1_DblClick()
  '
  Call Label22_DBLClick
  '
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
  '
  CODI3$ = Left(List1.List(List1.ListIndex), 16)
  CODI2% = CODINT%(CODI3$)
  If CODI2% > 0 Then
    List1.ToolTipText = DESCRIT0(CODI2%)
  End If
  '
End Sub

Private Sub List4_Click()
  '
  Call Label92_dblClick
  '
End Sub

Sub GENINSOL() 'CARGA ARCHIVO CENTRAL DE SOLICITUDES INTERNAS
    '
    Screen.MousePointer = 11
    '
    SQLX$ = "SELECT * FROM ENCASOLINT WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NUME_SOLINT > 0 "
    SQLX$ = SQLX$ + " ORDER BY Nume_Solint "
    Dim ENCATMP As ADODB.Recordset
    Set ENCATMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With ENCATMP
        JJ& = 0
       Do While Not .EOF
            XXY$ = REGBLAN$("!INSOLINT")
            REFER$ = SAFETEXT(!DESCRI_SOLINT)
            Call REPLA(XXY$, MKS(!NUME_SOLINT), 1, 4)
            If XVALO(!Status) = 9 Then
                REFER$ = "(Anul) " + REFER$
            End If
            Call REPLA(XXY$, REFER$, 5, 44)
            JJ& = JJ& + 1
            Call GRAREG("!INSOLINT", XXY$, JJ&)
            .MoveNext
       Loop
    End With
    ENCATMP.Close
    Set ENCATMP = Nothing
    Call TRACENTRAL("INSOLINT", "")
    Call SETULTREG("INSOLINT", JJ&)
    Call CIERRARCH("INSOLINT")
    Call FRESHECHO("INSOLINT")
    '
    Screen.MousePointer = 1
    '
End Sub
Sub GENINSOLADD(NUMSOL&) 'CARGA ARCHIVO CENTRAL DE SOLICITUDES INTERNAS
    '
    SQLX$ = "SELECT * FROM ENCASOLINT"
    SQLX$ = SQLX$ + " WHERE NUME_SOLINT = " & NUMSOL&
    Dim ENCATMP As ADODB.Recordset
    Set ENCATMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With ENCATMP
    af = ULTREG("INSOLINT")
      If Not .EOF Then
            XXY$ = REGBLAN$("INSOLINT")
            REFER$ = SAFETEXT(!DESCRI_SOLINT)
            Call REPLA(XXY$, MKS(!NUME_SOLINT), 1, 4)
            If XVALO(!Status) = 9 Then
                REFER$ = "(Anul) " + REFER$
            End If
            Call REPLA(XXY$, REFER$, 5, 44)
            Call REGAPP("INSOLINT", XXY$)
       End If
    End With
    ENCATMP.Close
    Set ENCATMP = Nothing
    Call CIERRARCH("INSOLINT")
    Call FRESHECHO("INSOLINT")
    '
End Sub

     Sub VERYREIMPSIM()
     Call ABRECOMPRAS
     '
     CONDI$ = "NUME_SOLINT > 0 "
     '
10   Call CARGALISEL("INSOPOAR", "ENCASOLINT", "NUME_SOLINT/F6.0;DESCRI_SOLINT/A52", CONDI$, "DISTINCT")
     Call SELECHO("INSOPOAR/Indice de Solicitudes Internas Por Artículos")
   
     NUDOCU$ = VALACT1$("INSOPOAR") ' ASIGNO EL NUMERO DEL FORMULARIO IMPRESO  SELECCIONAD0
     If NUDOCU$ = "" Then Exit Sub ' por si no elige ninguno y cierra el indice
     '
     'para el caso de carga maual muestra stdscreen
     If TRIM$(VALACT2$("INSOPOAR")) = "Solicitud de Materiales Carga Manual" Then
        FILTX$ = TRIM$(NUDOCU$)
        Call STDFORM("SOL-MAN", FILTX$)
        Exit Sub
     End If
     
     FORIPRE$ = TRIM$(CONTROL$("", "SOLIARTI"))
     If EXISTE%(LUDAT$ & FORIPRE$ & ".SQR") > 0 Then
         FILTX$ = TRIM$(NUDOCU$)
         Call STDFORM(FORIPRE$, FILTX$)
         GoTo 10
     End If
     '
     TIDOCUM$ = TRIM$(Left$(VALACT2("INSOPOAR"), 27))
     If RTA% = 2 Then  'POR ARTICULO
        TIDOCUM$ = Left$(VALACT2("INSOPOAR"), 36)
     End If
     '
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
     '
     Screen.MousePointer = 1
     Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a Esta Solicitud de Material.\ ")
     GoTo 10
        '
5    Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
    
    GoTo 10


End Sub


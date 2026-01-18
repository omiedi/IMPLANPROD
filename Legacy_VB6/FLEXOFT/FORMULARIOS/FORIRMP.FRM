VERSION 5.00
Begin VB.Form FORIRMP 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Recepción de Materias Primas"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   Icon            =   "FORIRMP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   Begin VB.Frame Frame1 
      BackColor       =   &H00F0F0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8520
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   12000
      Begin VB.CommandButton Command16 
         Height          =   495
         Left            =   11040
         Picture         =   "FORIRMP.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Imprime Vista Previa"
         Top             =   7200
         Width           =   615
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   10940
         ScaleHeight     =   75
         ScaleWidth      =   795
         TabIndex        =   70
         Top             =   7035
         Width           =   855
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   71
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Frame Frame20 
         BackColor       =   &H00F0F0FF&
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
         Left            =   8240
         TabIndex        =   65
         Top             =   6330
         Width           =   2500
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
            TabIndex        =   67
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
            TabIndex        =   68
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
            TabIndex        =   66
            Top             =   315
            Width           =   1695
         End
      End
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
         Left            =   6615
         Picture         =   "FORIRMP.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Inspección y Aprobación de Recepciones"
         Top             =   105
         Width           =   750
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00F0F0FF&
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
         Left            =   8240
         TabIndex        =   51
         Top             =   7085
         Width           =   2500
         Begin VB.CheckBox Check1 
            BackColor       =   &H00F0F0FF&
            Caption         =   "Ingresa sin Control"
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
            TabIndex        =   53
            Top             =   210
            Value           =   1  'Checked
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
            TabIndex        =   64
            Top             =   730
            Width           =   1590
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
            TabIndex        =   52
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
      Begin VB.Frame Frame3 
         BackColor       =   &H00F0F0FF&
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
         Left            =   4935
         TabIndex        =   42
         Top             =   6300
         Width           =   3225
         Begin VB.OptionButton Option0 
            BackColor       =   &H00F0F0FF&
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
            TabIndex        =   69
            Top             =   105
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1275
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
            TabIndex        =   49
            Top             =   525
            Width           =   195
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00F0F0FF&
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
            TabIndex        =   48
            Top             =   630
            Width           =   1170
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00F0F0FF&
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
            TabIndex        =   47
            Top             =   420
            Width           =   1275
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
            TabIndex        =   44
            Top             =   1530
            Width           =   2535
         End
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
            TabIndex        =   43
            Top             =   920
            Width           =   1995
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
            TabIndex        =   63
            Top             =   540
            Width           =   960
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
            TabIndex        =   50
            Top             =   300
            Width           =   750
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
            TabIndex        =   46
            Top             =   1260
            Width           =   2115
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
            TabIndex        =   45
            Top             =   945
            Width           =   1905
         End
         Begin VB.Line Line5 
            X1              =   0
            X2              =   10500
            Y1              =   1940
            Y2              =   1940
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00F0F0FF&
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
         TabIndex        =   41
         Top             =   6300
         Width           =   4695
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
            TabIndex        =   54
            Top             =   350
            Width           =   4670
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
         TabIndex        =   39
         Text            =   " "
         Top             =   770
         Width           =   2010
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00F0F0FF&
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
         Height          =   2805
         Left            =   160
         TabIndex        =   35
         Top             =   1260
         Width           =   10580
         Begin VB.Frame Frame21 
            Caption         =   "Frame8"
            Height          =   2535
            Left            =   5160
            TabIndex        =   72
            Top             =   360
            Width           =   15
         End
         Begin VB.Frame Frame19 
            Caption         =   "Frame19"
            Height          =   2535
            Left            =   10290
            TabIndex        =   62
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame18 
            Caption         =   "Frame18"
            Height          =   2535
            Left            =   9480
            TabIndex        =   61
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame17 
            Caption         =   "Frame17"
            Height          =   2535
            Left            =   8780
            TabIndex        =   60
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame16 
            Caption         =   "Frame16"
            Height          =   2535
            Left            =   7920
            TabIndex        =   59
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame15 
            Caption         =   "Frame15"
            Height          =   2535
            Left            =   6720
            TabIndex        =   58
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame8 
            Caption         =   "Frame8"
            Height          =   2535
            Left            =   5760
            TabIndex        =   57
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame7 
            Caption         =   "Frame7"
            Height          =   2535
            Left            =   1890
            TabIndex        =   56
            Top             =   315
            Width           =   15
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00E0E0E0&
            Height          =   435
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   10500
            TabIndex        =   37
            Top             =   315
            Width           =   10565
            Begin VB.Label Label22 
               BackStyle       =   0  'Transparent
               Caption         =   "Código          Descripción                U/S    Q.Ped     Recep    O/C Solic    O/F"
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
               TabIndex        =   38
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
            Height          =   2085
            Left            =   0
            TabIndex        =   36
            Top             =   745
            Width           =   10565
         End
         Begin VB.Line Line4 
            X1              =   10560
            X2              =   10560
            Y1              =   315
            Y2              =   2835
         End
      End
      Begin VB.CommandButton Command10 
         Caption         =   "!"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   34
         ToolTipText     =   "Lista de Proveedores con Entregas Programadas"
         Top             =   770
         Width           =   195
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00F0F0FF&
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
         Left            =   160
         TabIndex        =   25
         Top             =   4220
         Width           =   10580
         Begin VB.Frame Frame14 
            Caption         =   "Frame14"
            Height          =   1620
            Left            =   9555
            TabIndex        =   33
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame13 
            Caption         =   "Frame13"
            Height          =   1620
            Left            =   8505
            TabIndex        =   32
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame12 
            Caption         =   "Frame12"
            Height          =   1620
            Left            =   7455
            TabIndex        =   31
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame11 
            Caption         =   "Frame11"
            Height          =   1620
            Left            =   6720
            TabIndex        =   30
            Top             =   315
            Width           =   15
         End
         Begin VB.Frame Frame10 
            Caption         =   "Frame10"
            Height          =   1620
            Left            =   1890
            TabIndex        =   29
            Top             =   315
            Width           =   15
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
            TabIndex        =   28
            Top             =   710
            Width           =   10565
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H00C0C0C0&
            Height          =   435
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   10500
            TabIndex        =   26
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
               TabIndex        =   27
               Top             =   105
               Width           =   10515
            End
         End
         Begin VB.Line Line2 
            X1              =   10565
            X2              =   10565
            Y1              =   315
            Y2              =   1940
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   10500
            Y1              =   1940
            Y2              =   1940
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   6160
         Left            =   10920
         TabIndex        =   15
         Top             =   105
         Width           =   850
         Begin VB.CommandButton Command6 
            BackColor       =   &H00E0E0E0&
            Caption         =   "List"
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
            Left            =   0
            Picture         =   "FORIRMP.frx":1AB2
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Consultas y Listados"
            Top             =   4620
            Width           =   855
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Opciones"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORIRMP.frx":1DBC
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Opciones: Ver por Pantalla, Re-Imprimir, Anular"
            Top             =   3855
            Width           =   855
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
            Left            =   0
            Picture         =   "FORIRMP.frx":1F06
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Registrar Datos Respacho Aduana"
            Top             =   5390
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00E0E0E0&
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
            Height          =   770
            Left            =   0
            Picture         =   "FORIRMP.frx":2210
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00E0E0E0&
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
            Height          =   770
            Left            =   0
            Picture         =   "FORIRMP.frx":235A
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Configuración de Funcionamiento de Recepciones"
            Top             =   1542
            Width           =   855
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Articulos"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Picture         =   "FORIRMP.frx":279C
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Acceso a Maestro de Artículos"
            Top             =   3084
            Width           =   855
         End
         Begin VB.CommandButton Command11 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Proved."
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
            Left            =   0
            Picture         =   "FORIRMP.frx":2AA6
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Acceso a Base de Datos de Proveedores"
            Top             =   2313
            Width           =   855
         End
         Begin VB.CommandButton Command13 
            BackColor       =   &H00E0E0E0&
            Caption         =   "QC"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   770
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Calificación de Proveedores"
            Top             =   6405
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00E0E0E0&
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
            Height          =   770
            Left            =   0
            Picture         =   "FORIRMP.frx":2DB0
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Ayuda FLEXOFT en Línea"
            Top             =   771
            Width           =   855
         End
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00E0E0E0&
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   10920
         Picture         =   "FORIRMP.frx":30BA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Aprueba, Graba e Imprime Informe de Recepción"
         Top             =   7545
         Width           =   855
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
      Begin VB.Image Image2 
         Height          =   510
         Left            =   10360
         Picture         =   "FORIRMP.frx":33C4
         Top             =   6420
         Width           =   2010
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
         TabIndex        =   40
         Top             =   630
         Width           =   1485
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "IRMP"
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
         TabIndex        =   21
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
      Begin VB.Line Line7 
         X1              =   160
         X2              =   10710
         Y1              =   4070
         Y2              =   4070
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   840
         Width           =   960
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
End
Attribute VB_Name = "FORIRMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHIPRE$
Dim MUECOPEN%
Dim EMICONSI%
Dim LETRANO%
'
Private Sub Check1_CLICK()
   If Check1.Value = 1 Then
     If DERACCE%("APROREC", "Aprobación y Recepción Definitiva") < 2 Then
       Check1.Value = 0
       Exit Sub
     End If
   End If
   '
   If Check1.Value = 1 Then
      Label11 = USERNAME
    Else
      Label11 = ""
   End If
End Sub

Private Sub Command1_Click()
   Call SELECHO("ACREDOR")
   NCLIE% = Val(TRIM$(VALACT1$("ACREDOR")))
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
   Call CIERRARCH("*.*")
   ACONEC$ = "AMAPRO"
   Call CONECRUN(ACONEC$, "Padron Maestro de Proveedores")
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call APROBOL
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   Call CONECRUN("CALPRO02", "Sistema Calificación de Proveedores")
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   If VALIDATOS% > 0 Then
     Call GRABOLRE(OKX%)
     If OKX% > 0 Then
       Call PRIBOLRE(1)
       Call BLANFORMU
       Call SETULTREG("!BOLREDET", 0)
       Call CIERRARCH("*.*")
       Text1.TEXT = ""
       On Error Resume Next
       Text1.SetFocus
       On Error GoTo 0
     End If
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
    Command4.Enabled = False
    OPCIBREC.Show 1
99  Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    Call CONECRUN("FLSETUP/OPIRMP00", "Configuración de Funcionamiento")
    Call CIERRARCH("*.*")
    Call FINAL("")
    Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   COLISIMT.Show 1
   Command6.Enabled = True
End Sub
'
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
    Call CIERRARCH("*.*")
    ACONEC$ = "AMASTO01"
    If EXISTE%("AMASTO02.EXE") > 0 Then ACONEC$ = "AMASTO02"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")
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

Private Sub Label22_DblClick()
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

Private Sub Text1_DblClick()
   Call SELECHO("ACREDOR")
   NCLIE% = Val(TRIM$(VALACT1$("ACREDOR")))
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

Private Sub Text1_GotFocus()
   Text1.TEXT = TRIM$(Text1.TEXT)
   Text1.SelStart = 0
   Text1.SelLength = Len(Text1.TEXT)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If VALICLIE% > 0 Then
      Call BLANFORMU
      Call CARDACLI
      Call CARCOPEN
      If ULTREG&("!BOLREDET") < 1 Then
        Call CARDETCOM
      End If
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
   'If LETRANO% > 0 Then
   '  LINFE& = Val(TRIM$(Str$(Asc(Label13))) + "0000")
   '  LSUPE& = LINFE& + 9999
   '  PRONU& = 1 + LINFE&
   '  FIN& = ULTREG&("REMAPRI")
   '  For KKU& = FIN& - 512 To FIN&
   '    If KKU& > 0 Then
   '      XX$ = REGLEIDO$("REMAPRI", KKU&)
   '      If UCase$(Mid$(XX$, 101, 2)) = "MP" Then
   '        NBOLX& = CVS(Left$(XX$, 4))
   '        If NBOLX& <= LSUPE& Then
   '          If NBOLX& >= PRONU& Then
   '            PRONU& = NBOLX& + 1
   '          End If
   '        End If
   '      End If
   '    End If
   '  Next KKU&
   '  PRONUBOL& = PRONU& Mod 10000
   '  Exit Function
   'End If
   '
   NROBOL& = 1
   If ULTREG&("REMAPRI") > 0 Then
     UREG& = ULTREG&("REMAPRI")
     X$ = REGLEIDO$("REMAPRI", UREG&)
     If Len(X$) >= 4 Then NROBOL& = 1 + CVS(Left$(X$, 4))
   End If
   PRONUBOL& = NROBOL&
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
         Call MOVISTO(FREC%, CI%, LOTE$, "BR", DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC1%, DEPX%, CANTI)
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
   ATRI$ = ATRI$ + "/TITUPAN=Informe de Recepción M.Primas "
   ATRI$ = ATRI$ + TRIM$(Str$(NROBOL&)) + " - " + TRIM$(RASOCLI$(NC1%)) + " - " + FECHAREC$
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(140 + Top + Frame1.Top + Frame4.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Frame4.Left))
   '
   AVENTA$ = "IREMAPR" + TRIM$(Str$(VUELTA%))
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
     If CAN# < 0.0005 Then
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
   If ATTX% > 0 Or VUELTA% = 1 Then
     VUELTA% = 2
     GoTo 10
   End If
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
   Screen.MousePointer = 11
   NPRO% = Val(Text1)
   HO% = HOY(HOS$)
   '
   JJ& = 0
   REBLA$ = REGBLAN$("BOLREDET")
   Call SETULTREG("!BOLREDET", 0)
   '
   Screen.MousePointer = 11
   UREOCO& = ULTREG&("OCOMDETA")
   HAYOC% = 0
   '
   RFF$ = RECFILT$("OCOMDETA", 3, MKI$(NPRO%))
   For UI& = 1 To Len(RFF$) Step 4
     REDET& = CVS(Mid$(RFF$, UI&, 4))
     If REDET& > 0 Then
       If REDET& <= UREOCO& Then
         X$ = REGLEIDO$("OCOMDETA", REDET&)
         If CVI(Mid$(X$, 7, 2)) = NPRO% Then
           If Asc(Mid$(X$, 124, 1)) < 9 Then ' abierta
           CAPEN = CVS(Mid$(X$, 55, 4)) - CVS(Mid$(X$, 99, 4))
           COEFUNI = CVS(Mid$(X$, 63, 4))
           If COEFUNI < 0.0001 Then COEFUNI = 1
             If CAPEN > 0 Then
               NUOC& = CVS(Left$(X$, 4))
               FENTRESO% = CVI(Mid$(X$, 95, 2))
               CI% = CVI(Mid$(X$, 9, 2))
               LOTECO$ = Mid$(X$, 105, 12)
               Y$ = REGBLAN$("!BOLREDET")
               Call REPLA(Y$, MKI$(CI%), 83, 2)
               Call REPLA(Y$, UNIMED$(CI%), 85, 4)
               Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 89, 8)
               Call REPLA(Y$, LOTECO$, 107, 12)
               Call REPLA(Y$, MKS$(NUOC&), 151, 4)
               Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
               'Call REPLA(Y$, MKS$(REDET&), 189, 4)
               HAYOC% = 1
               JJ& = JJ& + 1
               Call GRAREG("!BOLREDET", Y$, JJ&)
               '
               TTXX$ = Space$(128)
               Call REPLA(TTXX$, CODEXT$(CI%), 1, 16)
               Call REPLA(TTXX$, DESCRIT0$(CI%), 17, 26)
               Call REPLA(TTXX$, UNIMED$(CI%), 44, 4)
               Call REPLA(TTXX$, FORMATNUM$(CAPEN * COEFUNI, "F8.1"), 48, 8)
               Call REPLA(TTXX$, FORMATNUM$(NUOC&, "F6.0"), 66, 6)
               FEX = FENTRESO%
               Call REPLA(TTXX$, Left$(FECHATEX$(FEX), 5), 74, 5)
               List1.AddItem TTXX$
             End If
           End If
         End If
       End If
     End If
   Next UI&
   '
   If HAYOC% = 0 Then
     Call COMUNI("No hay Ordenes de Compra Pendientes\para este Proveedor.")
   End If
   '
   Call SETULTREG("!BOLREDET", JJ&)
   Call CIERRARCH("!BOLREDET")
   '
   URECON& = ULTREG&("MACONSIG")
   '
   RFF$ = RECFILT$("MACONSIG", 2, MKI$(NPRO%))
   For UI& = 1 To Len(RFF$) Step 4
     I& = CVS(Mid$(RFF$, UI&, 4))
     If I& > 0 Then
       If I& <= URECON& Then
         X$ = REGLEIDO$("MACONSIG", I&)
         If CVI(Mid$(X$, 3, 2)) = NPRO% Then
           CAPEN = CVD(Mid$(X$, 63, 8)) - CVD(Mid$(X$, 111, 8))
           If CAPEN > 0 Then
             NUOC& = Val(Mid$(X$, 24, 6))
             CI% = CVI(Mid$(X$, 33, 2))
             '
             TTXX$ = Space$(128)
             Call REPLA(TTXX$, CODEXT$(CI%), 1, 16)
             Call REPLA(TTXX$, DESCRIT0$(CI%), 17, 39)
             Call REPLA(TTXX$, UNIMED$(CI%), 58, 4)
             Call REPLA(TTXX$, FORMATNUM$(CAPEN, "F8.1"), 62, 8)
             Call REPLA(TTXX$, FORMATNUM$(NUOC&, "F6.0"), 80, 6)
             List2.AddItem TTXX$
   '           Y$ = REGBLAN$("!BOLREDET")
   '           Call REPLA(Y$, MKI$(CI%), 83, 2)
   '           Call REPLA(Y$, UNIMED$(CI%), 85, 4)
   '           Call REPLA(Y$, MKD$(0), 89, 8)
   '           Call REPLA(Y$, "MP " + TRIM$(CSTRING$(MKS$(CAPEN), 3, 12, 1, 2)) + " " + UNIMED$(CI%), 107, 12)
   '           Call REPLA(Y$, MKS$(NUOC&), 151, 4)
   '           Call REPLA(Y$, MKS$((-1) * I&), 189, 4)
   '           For UI& = 1 To JJ&
   '             If Left$(Y$, 154) = Left$(REGLEIDO$("!BOLREDET", UI&), 154) Then GoTo 11129 ' SUPRIME SI COINCIDE
   '           Next UI&
   '           JJ& = JJ& + 1
   '           Call GRAREG("!BOLREDET", Y$, JJ&)
           End If
         End If
       End If
     End If
   Next UI&
11129 '   Next I&
   '     Call SETULTREG("!BOLREDET", JJ&)
   Screen.MousePointer = 1
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
     CANPED = CVS(Mid$(X$, 189, 4))
     CANREC = CVD(Mid$(X$, 89, 8))
     NUOCO = CVS(Mid$(X$, 151, 4))
     NUOFA = CVS(Mid$(X$, 189, 4))
     LOTECO$ = Mid$(X$, 107, 12)
     '
     TTXX$ = Space$(128)
     Call REPLA(TTXX$, CODEXT$(CI%), 1, 15)
     Call REPLA(TTXX$, DESCRIT0$(CI%), 17, 26)
     Call REPLA(TTXX$, UNIST$, 44, 4)
     Call REPLA(TTXX$, FORMATNUM$(CANPED, "F7.1"), 49, 7)
     Call REPLA(TTXX$, FORMATNUM$(CANREC, "F9.2"), 57, 9)
     Call REPLA(TTXX$, FORMATNUM$(NUOCO, "F6.0"), 67, 6)
     Call REPLA(TTXX$, FORMATNUM$(FENTRESOL%, "D8"), 74, 5)
     Call REPLA(TTXX$, FORMATNUM$(NUOFA, "F6.0"), 80, 6)
     Call REPLA(TTXX$, LOTECO$, 101, 12)
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
     CAN# = CVD(Mid$(YY$, 89, 8))
     If CI% < 1 Or CI% > NUMAS% Then
       Call MENSERR(24, "Código no Válido.")
       GoTo 99
     End If
     '
     If CAN# < 0.005 Then
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

Sub APROBOL() ' aprobacion de boleta de recepcion ************************************
   '
   If DERACCE%("APROBOL", "Aprobación Informes de Recepción") < 2 Then
     Exit Sub
   End If
   '
   Dim REGBO&(2048)
   '
20 X$ = "": Y$ = "": Z$ = "": YY$ = ""
   EBOL$ = "": PEDY$ = "": DESCRI$ = ""
   If ULTREG&("IREPRIPE") < 1 Then
     Call COMUNI("No Hay Recepciones\Pendientes de Aprobación")
     Exit Sub
   End If
   '
   Call SELECHO("IREPRIPE")
   'NBR$ = OBJPLA$("SELBOLPEN", "")
   'If NBR$ = "" Then
   '  Exit Sub
   'End If
   '
   NROBOL& = Val(VALACT1$("IREPRIPE"))    'CVS(Left$(NBR$, 4))
   If NROBOL& < 1 Then
     Exit Sub
   End If
   '
   'Call BLIMESS("Seleccionando ...")
   FIN& = ULTREG&("REPRIPEN")
   Call BLANARCH("!BOLREDET")
   CABOPE% = 0
   For I& = 1 To FIN&
     X$ = REGLEIDO$("REPRIPEN", I&)
     If CVS(Left$(X$, 4)) = NROBOL& Then
       REBO& = CVS(Mid$(X$, 189, 4))
       YY$ = REGLEIDO$("REMAPRI", REBO&)
       If Left$(YY$, 186) <> Left$(X$, 186) Then
         Call MENSERR(24, "Datos Inconsistentes - Consulte")
         Exit Sub
       End If
       '
       CANOM# = CVD(Mid$(X$, 89, 8))
       APRO# = CVD(Mid$(X$, 163, 8))
       RECHA# = CVD(Mid$(X$, 171, 8))
       FALTA# = CVD(Mid$(X$, 179, 8))
       '
       If APRO# > 0.05 Or RECHA# > 0.05 Or Abs(FALTA#) > 0.05 Then GoTo 21
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
     End If
     '
21 Next I&
   '
   JJ& = CABOPE%
   Call SETULTREG("!BOLREDET", JJ&)
   '
   If CABOPE% < 1 Then
     Call COMUNI("No hay Recepciones Pendientes\para la Boleta Elegida")
     GoTo 20
   End If
   '
15 Z$ = OBJPLA$("BOLREC-MUESTRW", Y$)
   If Z$ = "" Then
     GoTo 20
   End If
   DP1% = Asc(Mid$(Z$, 121, 1))
   If DP1% < 1 Or ECOARCH$("ECODEP", Chr$(DP1%)) = "" Then
     Call MENSERR(24, "Falta Indicar Depósito de Entrada")
     GoTo 15
   End If
   '
22 'FORVENTB.GRID.Clear
   ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(RASOCLI$(NC1%)) + " - " + FECHAREC$
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(140 + Top + Frame1.Top + Frame4.Top))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Frame4.Left))
   VTB% = VENTABU%("APROIREC" + ATRI$)
   '
   ATX% = 0: HAYRECHA% = 0
   For U& = 1 To ULTREG&("!BOLREDET")
     X$ = REGLEIDO$("!BOLREDET", U&)
     CANOM# = CVD(Mid$(X$, 89, 8))
     APRO# = CVD(Mid$(X$, 163, 8))
     RECHA# = CVD(Mid$(X$, 171, 8))
     If RECHA# >= 0.05 Then HAYRECHA% = 1
     FALTA# = CVD(Mid$(X$, 179, 8))
     If Abs(APRO#) < 0.005 Then
       If Abs(RECHA#) < 0.005 Then
         If Abs(FALTA#) < 0.005 Then
           GoTo 23                        ' NO SE PROCESO MOVIMIENTO
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
23 Next U&
   '
   If ATX% > 0 Then
     Call COMUNI("Cantidades Incoherentes.\Revise y Corrija !")
     GoTo 22
   End If
   '
24 If VTB% <> 0 Then GoTo 20
        '
        'CODRECHA$ = "BR"
        'REMIRECH$ = "BR." + TRIM$(Str$(NROBOL&))
        'If HAYRECHA% > 0 Then
        '  If TRIM$(CONTROL$("", "REMIRECH")) <> "" Then
        '    CODRECHA$ = TRIM$(CONTROL$("", "CODRRECH"))
        '    If CODRECHA$ = "" Then CODRECHA$ = "RD"
        '    FIN& = ULTREG&("MOVISTO")
        '    For U& = FIN& To 1 Step -1
        '       Call TRACE(20, FIN& - U& + 1, FIN&)
        '       XX$ = REGLEIDO$("MOVISTO", U&)
        '       If UCase$(Mid$(XX$, 21, 2)) = CODRECHA$ Then
        '         UNRORECH& = Val(Mid$(XX$, 26, 7))
        '         GoTo 11241
        '       End If
        '    Next U&
        '    UNRORECH& = 0
        '    '
11241   '    REMIRECH$ = CODRECHA$ + "-" + TRIM$(Str$(UNRORECH& + 1))
        '    '
        '  End If
        'End If
        '
25 Call CIERRARCH("MOVISTO")
   Call CIERRARCH("REPRIPEN")
   Call CIERRARCH("REMAPRI")
   '
   Call BLOQARCH("MOVISTO")
   Call BLOQARCH("REPRIPEN")
   Call BLOQARCH("REMAPRI")
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
           GoTo 28                        ' NO SE PROCESO MOVIMIENTO
         End If
       End If
     End If
     '
     UX% = U&
     PRIREG& = ULTREG&("REPRIPEN")
     REG& = REGBO&(UX%)
     If REG& > 0 Then
       If REG& <= ULTREG&("REPRIPEN") Then
         Y$ = REGLEIDO$("REPRIPEN", REG&)
         If Left$(Y$, 68) <> Left$(X$, 68) Then
           Call COMUNI("Punteros Incorrectos.\Consulte Soporte de Sistemas")
           GoTo 28
         End If
         Call GRAREG("REPRIPEN", X$, REG&)
         If REG& < PRIREG& Then PRIREG& = REG&
         '
         REG1& = CVS(Mid$(X$, 189, 4))
         If REG1& > 0 Then
           If REG1& <= ULTREG&("REMAPRI") Then
             Y$ = REGLEIDO$("REMAPRI", REG1&)
             If Left$(Y$, 68) <> Left$(X$, 68) Then
               Call COMUNI("Punteros Incorrectos.\Consulte Soporte de Sistemas")
               GoTo 27
             End If
             Call GRAREG("REMAPRI", X$, REG1&)
           End If
         End If
         '
27       DOPRI$ = "BR." + TRIM$(Str$(NROBOL&))
         If InStr(EMPREAC$, "SABO") > 0 Then DOPRI$ = LOTE$
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
           '
         End If
         '
         If RECHA# >= 0.05 Then
           CANTI = RECHA# * (-1)
           REFE$ = "RECHAZO - Rto." + TRIM$(REMI$)
           DOPRX1$ = REMIRECH$
           If TRIM$(DOPRX1$) = "" Then DOPRX1$ = DOPRI$
           Call MOVISTO(FF%, CI%, LOTE$, CODRECHA$, DOPRX1$, DOSEC$, REFE$, VUNI#, MONE$, NC1%, DEPX%, CANTI)
           HAYRECHA% = 1
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
     '
28 Next U&
   '
   Call CIERRARCH("MOVISTO")
   '
   ' DEPURAR REPRIPEN
   J& = 0: K& = 0: NOOBA& = 0
   For U& = 1 To ULTREG&("REPRIPEN")
     X$ = REGLEIDO$("REPRIPEN", U&)
     APRO# = CVD(Mid$(X$, 163, 8))
     RECHA# = CVD(Mid$(X$, 171, 8))
     FALTA# = CVD(Mid$(X$, 179, 8))
     If Abs(APRO#) < 0.005 Then
       If Abs(RECHA#) < 0.005 Then
         If Abs(FALTA#) < 0.005 Then
           J& = J& + 1
           Call GRAREG("REPRIPEN", X$, J&)
         End If
       End If
     End If
   Next U&
   '
   Call SETULTREG("REPRIPEN", J&)
   Call CIERRARCH("REPRIPEN")
   Call CIERRARCH("REMAPRI")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   LU$ = LUDAT$
   If EXISTE%(LU$ + "BRAP-ARC.FRM") > 0 Then
     HOII% = HOY(HOS$)
     FECHDOC$ = HOS$
     NUMEDOC$ = ""
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
   Call GENIBOLPE
   '
   If HAYRECHA% = 1 Then
     Call REMIRECHA(NC1%, DEPX%, REMIRECH$)
   End If
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   'If EXISTE%("CALPRO02.EXE") > 0 Then
   '  ICALREC.Label5 = TRIM$(Str$(NROBOL&))
   '  ICALREC.Show 1
   'End If
   GoTo 20
   '
End Sub
'
Sub GENIBOLPE()
        '
        Screen.MousePointer = 11
        '
        K& = 0: NOOBA& = 0
        For U& = 1 To ULTREG&("REPRIPEN")
           X$ = REGLEIDO$("REPRIPEN", U&)
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
                     XX$ = REGBLAN$("IREPRIPE")
                     '
                     Call REPLA(XX$, MKS$(NOOB&), 1, 4)
                     FEX = FEBO%
                     Call REPLA(XX$, FECHATEX$(FEX), 5, 8)
                     Call REPLA(XX$, RMTO$, 15, 16)
                     Call REPLA(XX$, ECOARCH$("PROVED1", MKI$(NPBO%)), 32, 29)
                     K& = K& + 1
                     Call GRAREG("IREPRIPE", XX$, K&)
                  End If
               End If
             End If
           End If
        Next U&
        '
        Call SETULTREG("IREPRIPE", K&)
        Call CIERRARCH("IREPRIPE")
        Call FRESHECHO("IREPRIPE")
        '
        Screen.MousePointer = 1
        '
End Sub



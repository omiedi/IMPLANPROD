VERSION 5.00
Begin VB.Form CIMANTEN07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0F0F0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cierre de Orden de Mantenimiento - Gestión de Mantenimiento Industrial"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   Icon            =   "CIMANTEN07.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text28 
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
      Height          =   1125
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   32
      ToolTipText     =   "Descripción General de la Orden de Mantenimiento"
      Top             =   3650
      Width           =   9725
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
      Left            =   7320
      TabIndex        =   68
      Top             =   5040
      Width           =   195
   End
   Begin VB.TextBox Text9 
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
      Left            =   6750
      MaxLength       =   5
      TabIndex        =   69
      Text            =   " "
      Top             =   5040
      Width           =   630
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   1095
      Index           =   1
      Left            =   1200
      ScaleHeight     =   1035
      ScaleWidth      =   9660
      TabIndex        =   61
      Top             =   7080
      Width           =   9720
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   820
         IntegralHeight  =   0   'False
         ItemData        =   "CIMANTEN07.frx":0442
         Left            =   -30
         List            =   "CIMANTEN07.frx":0444
         TabIndex        =   62
         ToolTipText     =   "Doble-Click para Informar Lectura de Contadores"
         Top             =   260
         Width           =   10575
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   330
         Left            =   0
         TabIndex        =   63
         Top             =   0
         Width           =   10515
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "No.   Referencia                            L. Ant.           Valor     L. Act.           Valor    Observaciones"
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
            Left            =   0
            TabIndex        =   64
            Top             =   30
            Width           =   10995
         End
      End
   End
   Begin VB.TextBox Text24 
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
      Left            =   9720
      TabIndex        =   30
      Text            =   " "
      Top             =   600
      Width           =   1210
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Recusos Propios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   120
      TabIndex        =   41
      Top             =   1470
      Width           =   4740
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   42
         ToolTipText     =   "Doble-Click para Informar Horas"
         Top             =   630
         Width           =   4500
      End
      Begin VB.Label Label92 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Legajo  Empleado           Horas    "
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   4500
      End
   End
   Begin VB.TextBox Text1 
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
      Left            =   1205
      MaxLength       =   15
      TabIndex        =   40
      Text            =   " "
      Top             =   240
      Width           =   1200
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
      Height          =   300
      Index           =   1
      Left            =   2355
      Picture         =   "CIMANTEN07.frx":0446
      TabIndex        =   39
      Top             =   240
      Visible         =   0   'False
      Width           =   175
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      Height          =   1350
      Index           =   0
      Left            =   1200
      ScaleHeight     =   1290
      ScaleWidth      =   9660
      TabIndex        =   35
      Top             =   5520
      Width           =   9720
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
         Height          =   1070
         IntegralHeight  =   0   'False
         ItemData        =   "CIMANTEN07.frx":0750
         Left            =   -30
         List            =   "CIMANTEN07.frx":0752
         TabIndex        =   38
         ToolTipText     =   "Doble-Click para Informar Consumos"
         Top             =   260
         Width           =   9720
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   330
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   10515
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   " Código                      Descripción                                                 U/M        Presup.          Utiliz."
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
            Left            =   0
            TabIndex        =   37
            Top             =   30
            Width           =   10900
         End
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00C0C0C0&
      Height          =   350
      Left            =   1200
      ScaleHeight     =   285
      ScaleWidth      =   9645
      TabIndex        =   33
      Top             =   3360
      Width           =   9705
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción General - Especificaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   34
         Top             =   30
         Width           =   9600
      End
   End
   Begin VB.CommandButton Command5 
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
      Left            =   10760
      Picture         =   "CIMANTEN07.frx":0754
      TabIndex        =   31
      Top             =   480
      Visible         =   0   'False
      Width           =   175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Tercerizado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   5040
      TabIndex        =   25
      Top             =   1470
      Width           =   3320
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
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
         Left            =   2240
         TabIndex        =   66
         Text            =   " "
         ToolTipText     =   "Costo de la Tercerización"
         Top             =   1080
         Width           =   920
      End
      Begin VB.CommandButton Command11 
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
         Left            =   2970
         TabIndex        =   27
         Top             =   285
         Width           =   195
      End
      Begin VB.TextBox Text3 
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
         Left            =   2220
         MaxLength       =   5
         TabIndex        =   26
         Text            =   " "
         Top             =   285
         Width           =   750
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Costo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1600
         TabIndex        =   67
         Top             =   1140
         Width           =   960
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1200
         TabIndex        =   29
         Top             =   330
         Width           =   960
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   150
         TabIndex        =   28
         Top             =   630
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0F0F0&
      Caption         =   "Realizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   8490
      TabIndex        =   16
      Top             =   1470
      Width           =   2460
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   160
         TabIndex        =   22
         Text            =   " "
         Top             =   480
         Width           =   1020
      End
      Begin VB.CommandButton Command12 
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
         Left            =   1155
         Picture         =   "CIMANTEN07.frx":0A5E
         TabIndex        =   21
         Top             =   480
         Width           =   175
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   160
         TabIndex        =   20
         Text            =   " "
         Top             =   1080
         Width           =   1020
      End
      Begin VB.CommandButton Command13 
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
         Left            =   1155
         Picture         =   "CIMANTEN07.frx":0D68
         TabIndex        =   19
         Top             =   1080
         Width           =   175
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   18
         Text            =   " "
         Top             =   480
         Width           =   780
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
         Height          =   300
         Left            =   1560
         TabIndex        =   17
         Text            =   " "
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
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
         Left            =   1440
         TabIndex        =   24
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
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
         Left            =   1440
         TabIndex        =   23
         Top             =   840
         Width           =   930
      End
   End
   Begin VB.TextBox Text7 
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
      Left            =   1200
      TabIndex        =   15
      Text            =   " "
      ToolTipText     =   "Referencia de la O/M"
      Top             =   960
      Width           =   6940
   End
   Begin VB.TextBox Text11 
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
      Left            =   9975
      TabIndex        =   8
      Top             =   8595
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004040&
      Height          =   10000
      Left            =   11025
      ScaleHeight     =   9945
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command10 
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
         Height          =   1170
         Left            =   120
         Picture         =   "CIMANTEN07.frx":1072
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Cierra Orden de Mantenimiento"
         Top             =   6980
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         Caption         =   "More"
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
         Left            =   120
         Picture         =   "CIMANTEN07.frx":137C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Opciones Varias"
         Top             =   1800
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command14 
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
         Picture         =   "CIMANTEN07.frx":14C6
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   4095
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command25 
         Caption         =   "Print"
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
         Picture         =   "CIMANTEN07.frx":1908
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Imprime Orden de Trabajo"
         Top             =   3285
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Copy"
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
         Picture         =   "CIMANTEN07.frx":1F72
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Copia Interactiva "
         Top             =   2400
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   6
         Top             =   5565
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   7
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command99 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "CIMANTEN07.frx":20BC
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   8350
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Picture         =   "CIMANTEN07.frx":23C6
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exit"
         Top             =   105
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Help"
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
         Picture         =   "CIMANTEN07.frx":2510
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Ayuda en Línea"
         Top             =   800
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -255
         Picture         =   "CIMANTEN07.frx":281A
         Stretch         =   -1  'True
         Top             =   5805
         Width           =   1410
      End
   End
   Begin VB.Line Line11 
      X1              =   10920
      X2              =   10920
      Y1              =   7080
      Y2              =   8170
   End
   Begin VB.Line Line10 
      X1              =   10920
      X2              =   10920
      Y1              =   5500
      Y2              =   6870
   End
   Begin VB.Line Line9 
      X1              =   1200
      X2              =   10920
      Y1              =   8170
      Y2              =   8170
   End
   Begin VB.Line Line8 
      X1              =   1200
      X2              =   10920
      Y1              =   6870
      Y2              =   6870
   End
   Begin VB.Line Line7 
      X1              =   1200
      X2              =   10920
      Y1              =   4770
      Y2              =   4770
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Referencia:"
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
      Left            =   30
      TabIndex        =   72
      Top             =   1050
      Width           =   1110
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   11025
      Y1              =   8290
      Y2              =   8290
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   11025
      Y1              =   4900
      Y2              =   4900
   End
   Begin VB.Label Label30 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7680
      TabIndex        =   71
      Top             =   5040
      Width           =   3255
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "Deposito de Consumo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   4680
      TabIndex        =   70
      Top             =   5100
      Width           =   1920
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Contadores:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   65
      Top             =   7240
      Width           =   1020
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "F.Emisión:"
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
      Left            =   8760
      TabIndex        =   60
      Top             =   690
      Width           =   930
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de O/M:"
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
      Index           =   0
      Left            =   8955
      TabIndex        =   59
      Top             =   330
      Width           =   1050
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Resp.:"
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
      Left            =   8520
      TabIndex        =   58
      Top             =   1080
      Width           =   585
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   11025
      Y1              =   3150
      Y2              =   3150
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Planta:"
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
      Left            =   30
      TabIndex        =   57
      Top             =   720
      Width           =   1110
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Sección:"
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
      Left            =   1860
      TabIndex        =   56
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Equipo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   39
      Left            =   360
      TabIndex        =   55
      Top             =   315
      Width           =   795
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   2640
      TabIndex        =   54
      Top             =   240
      Width           =   5505
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   1200
      TabIndex        =   53
      Top             =   600
      Width           =   690
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   2805
      TabIndex        =   52
      Top             =   600
      Width           =   930
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Linea:"
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
      Left            =   3240
      TabIndex        =   51
      Top             =   720
      Width           =   1110
   End
   Begin VB.Label Label15 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   4410
      TabIndex        =   50
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Posición:"
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
      Left            =   6240
      TabIndex        =   49
      Top             =   720
      Width           =   1110
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7410
      TabIndex        =   48
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   10080
      TabIndex        =   47
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label24 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
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
      Left            =   9120
      TabIndex        =   46
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Consumidos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   45
      Top             =   5700
      Width           =   1140
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   44
      Top             =   3450
      Width           =   1140
   End
   Begin VB.Line Line6 
      X1              =   10920
      X2              =   10920
      Y1              =   3360
      Y2              =   4750
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   11025
      Y1              =   1370
      Y2              =   1370
   End
   Begin VB.Label Label13 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2280
      TabIndex        =   12
      Top             =   10840
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2040
      TabIndex        =   11
      Top             =   10840
      Visible         =   0   'False
      Width           =   2910
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Resp.:"
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
      Left            =   9450
      TabIndex        =   9
      Top             =   8700
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Line Line1 
      X1              =   -420
      X2              =   11580
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "CIMANTEN07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHEX$(16)
Dim YACAR%, MODIFIC%
Dim GRABATODO%
Dim ARCHIPRE$
Private Sub Command1_Click()
   '
   Call SELECHO("TADEPOSI")
   NDE% = Val(TRIM$(VALACT1$("TADEPOSI")))
   If NDE% > 0 Then
       Text9.TEXT = TRIM$(Str$(NDE%))
       Label30 = ECOARCH$("TADEPOSI", MKI$(Val(Text9.TEXT)))
   End If
   '
End Sub

Private Sub AyuCli_Click()
  Call Command3_Click
End Sub

Private Sub Command10_Click()
  '
  Command10.Enabled = False
  'VALIDACIONES DE CIERRE
  For JX% = 0 To List4.ListCount - 1
    L1$ = List4.List(JX%)
    LEG1% = Val(Left$(L1$, 8))
    EMP1$ = TRIM$(ECOARCH$("PADRON", MKI$(LEG1%)))
    HS1 = Val(Mid$(L1$, 29, 5))
    If HS1 <= 0 Then
      Call MENSERR(24, "Es Necesario Informar Horas Dedicadas\del Empleado '" & EMP1$ & "'")
      GoTo 99
    End If
  Next JX%
  '
  For JX% = 0 To List3.ListCount - 1
    L1$ = List3.List(JX%)
    FUL1% = FECHANUM(Mid$(L1$, 47, 8))
    VAL1 = Val(Mid$(L1$, 55, 10))
    CONT1$ = TRIM$(Mid$(L1$, 5, 22))
    If VAL1 <= 0 Then
      Call MENSERR(24, "Es Necesario Informar Valor\del Contador '" & CONT1$ & "'")
      GoTo 99
    End If
  Next JX%
  '
  For JX% = 0 To List1.ListCount - 1
    L1$ = List1.List(JX%)
    REP1$ = TRIM$(Mid$(L1$, 1, 16))
    CANTINF = Val(Mid$(L1$, 73, 10))
    If CANTINF <= 0 Then
      Call MENSERR(24, "Es Necesario Informar la Cantidad\Consumida del Repuesto '" & REP1$ & "'")
      GoTo 99
    End If
  Next JX%
  '
  DEPOCON% = Val(Text9.TEXT)
  If ECOARCH$("TADEPOSI", MKI$(DEPOCON%)) = "" Then
    Call MENSERR(24, "Es Necesario Asignar Deposito de Consumo.")
    GoTo 99
  End If
  '
  NPRO = Val(Text3.TEXT) * -1
  If REGICLI&(NPRO) <= 0 And List4.ListCount <= 0 Then
    Call MENSERR(24, "Es Necesario Asignar Empleados\o Tercerizar la O/M.")
    GoTo 99
  End If
  '
  HOII% = HOY(HOS$)
  FEMI% = FECHANUM(Text24.TEXT)
  If FEMI% < 1 Or FEMI% > HOII% Then
    Call MENSERR(24, "Fecha de Emisión no Válida")
    GoTo 99
  End If
    '
  FEPROCOM% = FECHANUM(Text2.TEXT)
  If FEPROCOM% < FEMI% Then
    Call MENSERR(24, "Fecha de Comienzo Proyectada no Válida")
    GoTo 99
  End If
  '
  FEPROFIN% = FECHANUM(Text4.TEXT)
    If FEPROFIN% < FEPROCOM% Then
    Call MENSERR(24, "Fecha de Finalización Proyectada NO puede ser\ Inferior a la Fecha de Comienzo Proyectada.")
    GoTo 99
  End If
  '
  '
  NRO& = Val(Label22)
  CONDI$ = " NumeOrMa_OM = " & NRO& & " "
  Call BORRAREGISTROS("RECUPROP", CONDI$, REG&, "NOMESS")
  'LIMPIO LOS EMPLEADO ASIGNADOS
'  With recuprop
'    SQLX$ = " NumeOrMa_OM = " & NRO& & " "
'    .FindFirst SQLX$
'13  If .NoMatch = False Then
'      .Delete
'      .FindNext SQLX$
'      GoTo 13
'      .Update
'    End If
'  End With
'  '
  'CARGO EMPLEADOS ASIGNADOS A LA O/M y HORAS DEDICADAS
  Call ABRECONEXION
  Dim RecuPrTmp As ADODB.Recordset
  Set RecuPrTmp = New ADODB.Recordset
  SQLX$ = "SELECT * FROM RECUPROP "   ' WHERE NumeOrMa_OM = " & NRO& & " "
  RecuPrTmp.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
  If List4.ListCount > 0 Then
    For JX% = 0 To List4.ListCount - 1
      L1$ = List4.List(JX%)
      LEG1% = Val(Left$(L1$, 8))
      HS1 = Val(Mid$(L1$, 29, 5))
      With RecuPrTmp
        .AddNew   ' YA CONVERTIDO
        ON ERROR RESUME NEXT
        !CodiEmpr = Codiemp%
        ON ERROR GOTO 0
          !NumeOrMa_OM = NRO&
          !Descri_OM = Text7.TEXT
          !LegaEmp_OM = LEG1%
          !Empleado_OM = ECOARCH$("PADRON", MKI$(LEG1%))
          !FeCie_OM = Now
          !CaHora_OM = HS1
        .Update
      End With
    Next JX%
  End If
  RecuPrTmp.Close
  Set RecuPrTmp = Nothing
  '
  'CONSUMO MATERIAL
  Dim Reserva As ADODB.Recordset
  Set Reserva = New ADODB.Recordset
  SQLX$ = "SELECT * FROM Reserva where cast(SUBSTRING(IdSubOr,4,6) as integer) = " & NRO& & ""
  Reserva.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

  For UX% = 0 To List1.ListCount - 1
    YY$ = List1.List(UX%)
    NORMA$ = NBRS$ + TRIM$(Str$(UX%))
    NORFX$ = NORMA$
    '
    CIRF0$ = Mid$(YY$, 1, 16)
    CI1% = CODINT%(CIRF0$)
    CANTRX# = Val(Mid$(YY$, 62, 10))
    USOIX = Val(Mid$(YY$, 73, 10))
    OBS2$ = Mid$(YY$, 73, 25)
    IDLOT$ = ""
    '
    With Reserva
      If Not (.EOF And .BOF) Then
        'CONSUMO EL MATERIAL
        CMO$ = TRIM$(CONTROL$("", "CODORMAN"))
        If CMO$ = "" Then CMO$ = "OM"
        DOPRILA$ = !IdSubOr
        DOPRICO$ = DOPRILA$
        DOSELA$ = DOPRILA$
        DOSECO$ = DOPRILA$
        CODICONJ$ = CODEXT$(CI1%)
        REFE$ = "CONSUMO " + CODICONJ$
        NUORIT% = Val(Mid$(!IdSubOr, 11))
        VUNI# = 0
        MONEII% = 1 'PESOS
        DEPO% = DEPOCON% 'DEPO% = DEPOPRE%(CI1%)
        CANTI = USOIX * (-1)
        HOII% = HOY(HOS$)
        'HAGO MOVIMIENTO
        Call MOVISTOK(HOII%, CI1%, " ", CMO$, DOPRILA$, DOPRICO$, DOSELA$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NC1%, DEPO%, CANTI)
        '
        'CANCELO RESERVAS
          '(*)para guardar la cant para el alta por copia
          !UsoUnit = !CantRes '(*)
          !CantRes = 0
          !CantCons = USOIX
        .Update
      End If
    End With
  Next UX%
  Reserva.Close
  Set Reserva = Nothing

  '
' ***LIMPIO RESERVAS NO LAS BORRO LAS RESERVAS***
' ***POR EL ALTA POR COPIA***
'   '
'   CMO$ = TRIM$(CONTROL$("", "CODORMAN"))
'   If CMO$ = "" Then CMO$ = "OM"
'   NBRS$ = TRIM$(Str$(NRO&)): While Len(NBRS$) < 6
'                                NBRS$ = "0" + NBRS$
'                              Wend
'   NBRS$ = CMO$ + "-" + NBRS$ + "-"
'   Call ABRESERVA
'   With Reserva
'     SQLX$ = " VAL(MID(IdSubOr,4,6)) = " & NRO& & ""
'     .FindFirst SQLX$
'15    If .NoMatch = False Then
'         'BORRO RESERVAS
'         .Delete
'       .FindNext SQLX$
'       GoTo 15
'     End If
'  End With
  '

  'CIERRE
   'MARCO LA OM COMO CERRADA
    Call ABREOMAN
    '
    Dim OrdeMant As ADODB.Recordset
    Set OrdeMant = New ADODB.Recordset
    SQLX$ = "SELECT * FROM OrdeMant WHERE NumeOrMa_OM = " & NRO& & " "
    OrdeMant.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With OrdeMant
      If Not (.EOF And .BOF) Then
          !Status_OM = 3 'CERRADA
          !FechCie_OM = Now
          !HoraCie_OM = Now
          !UsuCie_OM = USERNAME$
          !CostoTer_OM = Val(Text8.TEXT)
        .Update
      End If
    End With
  '
  'GRABO LA LECTURA DEL CONTADOR

  For JX% = 0 To List3.ListCount - 1
    L1$ = List3.List(JX%)
    NRO1& = Val(Mid$(L1$, 1, 3))
    FUL1% = FECHANUM(Mid$(L1$, 47, 8))
    VAL1 = Val(Mid$(L1$, 55, 10))
    OBSE$ = Mid$(L1$, 67, 64)
    '
    Dim Contadores As ADODB.Recordset
    Set Contadores = New ADODB.Recordset
    SQLX$ = "SELECT * FROM Contadores WHERE Num_Con = " & NRO1& & " AND "
    SQLX$ = SQLX$ + "Codigo_Equ = '" & Text1.TEXT & "' "
    Contadores.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With Contadores
        !FUL_Con = CVDAT(FUL1%)
        VANT = !Val_Con
        !Val_Con = VAL1
        !Observa_Con = OBSE$
        'CALCULO FECHA DE PROXIMA LECTURA
        FQ = !Freq_Con
        F1 = CDbl(CVDAT(FUL1%))
        FP1 = F1 + FQ
        FPL1% = CVINT%(FP1)
        !FPL_Con = CVDAT(FPL1%)
        'SACO PROMEDIO
        PROM1 = (VANT + VAL1) / 2
        !Prom_Con = PROM1
      .Update
    End With
  Next JX%
  '
  Call INDIMAN 'INDICE DE O/MANTENIMIENTO
  '
  Unload Me
  LISOMAQ07.Command46_Click
  '
99 Command10.Enabled = True
  '
End Sub

Private Sub Command11_Click()
   '
   Call SELECHO("ACREDOR")
   NCLIE% = Val(TRIM$(VALACT1$("ACREDOR")))
   If NCLIE% > 0 Then
     If REGICLI&((-1) * NCLIE%) > 0 Then
       Text3.TEXT = TRIM$(Str$(NCLIE%))
       Label18 = RASOCLI(NCLIE% * (-1))
     End If
   End If
   '
End Sub

Private Sub Command12_Click()
   '
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text2.TEXT = VALACT1$("SELFECHA")
   End If
   '
End Sub

Private Sub Command13_Click()
   '
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text4.TEXT = VALACT1$("SELFECHA")
   End If
   '
End Sub

Private Sub Command14_Click()
  Command14.Enabled = False
    OPOMAQ04.Show 1
  Command14.Enabled = True
End Sub

Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command2_Click()
  Command2.Enabled = False
    LISOMAQ07.Show 1
  Command2.Enabled = True
End Sub
Private Sub Command4_Click()
   Call COMUNI("Ayuda FLEXOFT en Línea.\No Disponible por el Momento.")
End Sub

Private Sub ConPagCli_Click()
With OPMASCL04
  .Option1(6).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub CuCoCli_Click()
With OPMASCL04
  .Option1(8).Value = True
  Call .Command16_Click
End With
End Sub


Private Sub DeLiCreCli_Click()
    Call AMACLIS04.Command3_Click
End Sub

Private Sub EditCli_Click()
   'Call Command9_Click
End Sub

Private Sub ConGenCli_Click()
  '
  Call Command14_Click
  '
End Sub

Private Sub Exit_Click()
   Call Command3_Click
End Sub

Private Sub Command6_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Clipboard.GetText + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
End Sub

Private Sub Command9_Click(Index As Integer)
   '
   Call SELECHO("INDIEQUI")
   COD$ = TRIM$(VALACT1$("INDIEQUI"))
   If COD$ <> "" Then
     Text1.TEXT = COD$
   End If
   '
End Sub

'
Private Sub Form_Load()
  '
  Call CENTRAFORM(Me)
  '
  EPAC$ = TRIM$(EMPREAC$)
  If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
  End If
  Call GRATITFOR(Caption)
  '
  HOII% = HOY(HOS$)
  Text24.TEXT = HOS$
  '
  Label24.Caption = USERNAME$
  '
  '
End Sub
'
Private Sub Form_Activate()
  WindowState = 0
  openforms = DoEvents
End Sub
'
Private Sub Form_Resize()
  If WindowState = 1 Then
    Call CIERRARCH("*.*")
    Call BAJALDISCO
  End If
End Sub
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     openforms = DoEvents
   Wend
End Sub

'\\\OT-05-0396-WAR-13/06/05///
'Private Sub COMMAND15_CLICK()
'    '
'    If DERACCE%("GRACOTCA", "Grabar Cotización") >= 2 Then
'      NUPRESU1$ = TRIM(Label31)
'      'RUBRO =2
'      For I = 0 To List1.ListCount - 1
'          Y$ = List1.List(I)
'          CODEXT1$ = TRIM(Mid$(Y$, 1, 16))
'          CI% = CODINT(CODEXT1$)
'          PREU1# = Val(Mid$(Y$, 65, 12))
'          CANTI1# = Val(Mid$(Y$, 78, 7))
'          IMPITEM# = Val(Mid$(Y$, 86, 12))
'          Call ABREPRECAM
'          With DetaPresu
'            SQLX$ = " Nro_Presu = '" & NUPRESU1$ & "' "
'            SQLX$ = SQLX$ + " AND  Rubro_Det = 3 "
'            'SQLX$ = SQLX$ + " AND  SubRub_Det = 1 "
'            SQLX$ = SQLX$ + " AND CodINte_Det = " & CI% & " "
'            .FindFirst SQLX$
'10              If .NoMatch = False Then
'              .Edit
'                !PreUni_Det = PREU1#
'                !Cantid_Det = CANTI1#
'                !Impo_Det = IMPITEM#
'              .Update
'              .FindNext SQLX$
'              GoTo 10
'            End If
'          End With
'      Next I
'    End If
'End Sub
'\\\OT-05-0396-WAR-FIN///
Sub CAMGRIS()
   '
   COLOBAK = RGB(230, 230, 230)
   Text6.BackColor = COLOBAK
   Text13.BackColor = COLOBAK
   Text38.BackColor = COLOBAK
   TxtLiPre.BackColor = COLOBAK
   TICAMBIS.BackColor = COLOBAK
   '
End Sub

Private Sub Command25_Click()
   'Call PRILISMAT
  PRINTFORM
End Sub

Private Sub Command3_Click()
   '
   List1.Clear
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
   LISOMAQ07.Command46_Click
   'Unload Me
   '
End Sub

Private Sub Command5_Click()
   '
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text24.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text24.TEXT)
   '
End Sub

Private Sub LiCuInCli_Click()
    Call AMACLIS04.Command5_Click
End Sub

Private Sub Label17_dblClick()
  '
  Call BLANARCH("!CONTADOR")
  JJ& = 0
  For JX% = 0 To List3.ListCount - 1
    L1$ = List3.List(JX%)
    NRO1% = Val(Mid$(L1$, 1, 3))
    REFE1$ = Mid$(L1$, 5, 22)
    'FEQ1 = Val(Mid$(L1$, 28, 7))
    FUL1% = FECHANUM(Mid$(L1$, 28, 8))
    VAL1# = Val(Mid$(L1$, 36, 10))
    FPL1% = FECHANUM(Mid$(L1$, 47, 8))
    PROME1# = Val(Mid$(L1$, 55, 10))
    OBSE1$ = Mid$(L1$, 67, 64)
    '
    Y$ = REGBLAN$("CONTADOR")
    Call REPLA(Y$, MKI$(NRO1%), 1, 2)
    Call REPLA(Y$, REFE1$, 3, 32)
    Call REPLA(Y$, MKS(FEQ1), 33, 4)
    Call REPLA(Y$, MKI$(FUL1%), 37, 2)
    Call REPLA(Y$, MKD$(VAL1#), 39, 8)
    Call REPLA(Y$, MKI$(FPL1%), 47, 2)
    Call REPLA(Y$, MKD$(PROME1#), 49, 8)
    Call REPLA(Y$, OBSE1$, 57, 64)
    JJ& = JJ& + 1
    Call GRAREG("!CONTADOR", Y$, JJ&)
  Next JX%
  Call SETULTREG("!CONTADOR", JJ&)
  Call CIERRARCH("!CONTADOR")
  '
  PRF$ = "CARLECIE"
  ATRI$ = ATRI$ + "/TITUPAN=Contadores de Equipo /"
  ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + AMAEQUI07.Left + 170))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + AMAEQUI07.Top + 4250))
  VTB% = VENTABU%(PRF$ + ATRI$)
  If VTB% >= 0 Then
    '
    List3.Clear
    '
    FIN& = ULTREG&("!CONTADOR")
    For J& = 1 To FIN&
      X$ = REGLEIDO$("!CONTADOR", J&)
      NRO% = CVI(Mid$(X$, 1, 2))
        NRO2$ = FORMATNUM$(NRO%, "F3.0")
      REFE$ = Mid$(X$, 3, 32)
      FEQ = CVS(Mid$(X$, 33, 4))
        FEQ2$ = FORMATNUM$(FEQ, "F7.1")
      FUL% = CVI(Mid$(X$, 37, 2))
        FUL2$ = FORMATNUM$(FUL%, "D8")
      VALO = CVD(Mid$(X$, 39, 8))
        VALO1$ = FORMATNUM$(VALO, "F10.2")
      FPL% = CVI(Mid$(X$, 47, 2))
        FPL2$ = FORMATNUM$(FPL%, "D8")
      PROME = CVD(Mid$(X$, 49, 8))
        PROM1$ = FORMATNUM$(PROME, "F10.2")
      OBSE$ = Mid$(X$, 57, 64)
      'AGREGO AL LIST
      L1$ = Space(256)
      Call REPLA(L1$, NRO2$, 1, 3)
      Call REPLA(L1$, REFE$, 5, 22)
      'Call REPLA(L1$, FEQ2$, 28, 7)
      Call REPLA(L1$, FUL2$, 28, 8)
      Call REPLA(L1$, VALO1$, 36, 10)
      Call REPLA(L1$, FPL2$, 47, 8)
      Call REPLA(L1$, PROM1$, 55, 10)
      Call REPLA(L1$, OBSE$, 67, 64)
      List3.AddItem L1$
    Next J&
    '
  End If
  '
End Sub

Private Sub label3_Change()
  '
  Label6 = TRIM$(ECOARCH$("PLANTPRO", Label3.Caption))
  '
End Sub

Private Sub Label7_Change()
  '
  CS1 = Val(Label7)
  Label13.Caption = TRIM$(ECOARCH$("SECPROD", MKS$(CS1)))
  '
End Sub
Private Sub LiGeClaCli_Click()
 Call Command25_Click
End Sub

Private Sub SuspOper_Click()
    Call Command3_Click
End Sub

Sub GENEPLACO() 'GENERAR CAPLACO.DAT
    '
    NUOTRAB$ = TRIM$(Text1.TEXT)
    Call ABREORTRA
    SQLX$ = "SELECT * FROM PlanCob "
    SQLX$ = SQLX$ + " WHERE Nro_Otrab= '" & NUOTRAB$ & "' "
    Set PLACOTMP = OTrabajo.OpenRecordset(SQLX$, 4)
    
    J& = 0
    With PLACOTMP
      '
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do Until .EOF = False
          Y$ = REGBLAN("CAPLACO")
          Call REPLA(Y$, !Nro_Otrab, 1, 16)
          Call REPLA(Y$, !NuOrItem, 17, 2)
          Call REPLA(Y$, !Refe_Pcob, 19, 64)
          Call REPLA(Y$, !Impo_Pcob, 83, 8)
          Call REPLA(Y$, !FeVen_Pcob, 91, 2)
          Call REPLA(Y$, !NroFac_Pcob, 93, 16)
          Call REPLA(Y$, !Moneda, 109, 2)
          Call REPLA(Y$, !Coti_Pcob, 111, 8)
          J& = J& + 1
          Call GRAREG("!CAPLACO", Y$, J&)
        .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    '
    Call SETULTREG("!CAPLACO", J&)
    Call CIERRARCH("!CAPLACO")
    
End Sub

Private Sub Label9_DblClick()
'
  If TRIM(Text1) <> "" Then
    '
    PRF$ = "CAREPUCO"
    ATRI$ = ATRI$ + "/TITUPAN=Lista de Repuestos /"
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + AMAEQUI07.Left + 170))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + AMAEQUI07.Top + 2250))
    VTB% = VENTABU%(PRF$ + ATRI$)
    If VTB% >= 0 Then
      '
      List1.Clear
      '
      FIN& = ULTREG&("!REPUESTO")
      For J& = 1 To FIN&
        X$ = REGLEIDO$("!REPUESTO", J&)
        CODI1% = CVI(Mid$(X$, 1, 2))
        CODI2$ = CODEXT$(CODI1%)
        DESC2$ = DESCRIT0(CODI1%)
        UM2$ = Mid$(X$, 3, 2)
        USO1 = CVS(Mid$(X$, 5, 4))
        USO2$ = FORMATNUM$(USO1, "F9.1")
        OBS2$ = Mid$(X$, 9, 25)
        USOIN1 = CVS(Mid$(X$, 34, 4))
        USOIN2$ = FORMATNUM$(USOIN1, "F9.1")
        
        'AGREGO AL LIST
        L1$ = Space(128)
        Call REPLA(L1$, CODI2$, 1, 16)
        Call REPLA(L1$, DESC2$, 18, 40)
        Call REPLA(L1$, UM2$, 59, 2)
        Call REPLA(L1$, USO2$, 62, 10)
        Call REPLA(L1$, USOIN2$, 73, 10)
        'Call REPLA(L1$, OBS2$, 73, 25)
        List1.AddItem L1$
      Next J&
      '
    End If
  End If
  '
End Sub

Private Sub Label92_dblClick()
'
  If TRIM(Text1) <> "" Then
    '
    Call BLANARCH("!RECPROP")
    JJ& = 0
    For JX% = 0 To List4.ListCount - 1
      L1$ = List4.List(JX%)
      LEG1% = Mid$(L1$, 1, 8)
      HS1 = Val(Mid$(L1$, 29, 5))
      '
      YY$ = REGBLAN$("RECPROP")
      Call REPLA(YY$, MKI$(LEG1%), 1, 2)
      Call REPLA(YY$, MKS$(HS1), 3, 4)
      JJ& = JJ& + 1
      Call GRAREG("!RECPROP", YY$, JJ&)
      '
    Next JX%
    Call SETULTREG("!RECPROP", JJ&)
    Call CIERRARCH("!RECPROP")
    '
    PRF$ = "HORECPRO"
    ATRI$ = ATRI$ + "/TITUPAN=Lista de Empleados /"
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + AMAEQUI07.Left + 170))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + AMAEQUI07.Top + 2250))
    VTB% = VENTABU%(PRF$ + ATRI$)
    If VTB% >= 0 Then
      '
      List4.Clear
      '
      FIN& = ULTREG&("!RECPROP")
      For J& = 1 To FIN&
        X$ = REGLEIDO$("!RECPROP", J&)
        LEG1% = CVI(Mid$(X$, 1, 2))
          LEG2$ = FORMATNUM$(LEG1%, "F7.0")
        EMP1$ = ECOARCH$("PADRON", MKI$(LEG1%))
        HS1 = CVS(Mid$(X$, 3, 4))
        HS2$ = FORMATNUM$(HS1, "F5.1")
        'AGREGO AL LIST
        L1$ = Space(128)
        Call REPLA(L1$, LEG2$, 1, 8)
        Call REPLA(L1$, EMP1$, 10, 18)
        Call REPLA(L1$, HS2$, 29, 5)
        List4.AddItem L1$
      Next J&
      '
    End If
  End If
  '
End Sub

Private Sub List1_DblClick()
  Call Label9_DblClick
End Sub


Private Sub List3_DblClick()
   Call Label17_dblClick
End Sub

Private Sub List4_dblClick()
  '
  Call Label92_dblClick
  '
End Sub

Private Sub Text1_Change()
  '
 ' Call BLANFORMU
  If TRIM$(Text1) <> "" Then
    If ECOARCH$("INDIEQUI", Text1.TEXT) <> "" Then
      Label1 = ECOARCH$("INDIEQUI", Text1.TEXT)
      Call CARGAEQUI(Text1.TEXT)
    End If
  End If
  '
End Sub
Private Sub Text10_GOTFOCUS()
  '
  Command3.SetFocus
  '
End Sub

Private Sub Text12_GotFocus()
  '
  Command3.SetFocus
  '
End Sub
Private Sub Text13_GotFocus()
  '
  Command3.SetFocus
  '
End Sub
Private Sub Text18_GOTFOCUS()
  '
  Command3.SetFocus
  '
End Sub

Private Sub Text2_change()
  '
  Static HOX%
  If HOX% = 0 Then
    HOX% = HOY(HO$)
  End If
  '
  If FECHANUM(Text2.TEXT) > HOX% Then
    Text5.TEXT = "0:00"
  ElseIf FECHANUM(Text2.TEXT) = HOX% Then
    Text5.TEXT = Left$(Time$, 5)
  End If
  '
End Sub

Private Sub Text1_GotFocus()
  '
  On Error Resume Next
    Command3.SetFocus
  On Error GoTo 0
  '
End Sub

Private Sub Text21_Change()
'    '
'    If Val(TEXT21.TEXT) < 0 Or Val(TEXT21.TEXT) > 255 Then
'       TEXT21.TEXT = ""
'    End If
'    '
'    If Val(TEXT21) < 1 Then
'        TxtLiPre.TEXT = "CONCERTADOS FIJOS"
'        ARCHIPRE$ = "MASTER"
'      Else
'        TxtLiPre.TEXT = ECOARCH$("LISTAS", MKI$(Val(TEXT21.TEXT)))
'        ARCHIPRE$ = TRIM$(TEXT21)
'        While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
'        ARCHIPRE$ = "LIPRE" + ARCHIPRE$
'    End If
'    '
End Sub


Sub BLANFORMU()
   '
   Text2 = ""
   Text3 = ""
   '
   HOII% = HOY(HOS$)
   Text24.TEXT = HOS$
   '
   List1.Clear
   Text28.TEXT = ""
   '
   Text23 = ""
   Text35 = ""
   Text36 = ""
   Text37 = ""
   Text3 = ""
   Text18 = ""
   Text5 = ""
   Text7 = ""
   Text8 = ""
   '
   Label1 = ""
   Label3 = ""
   Label6 = ""
   Label7 = ""
   Label13 = ""
   Label15 = ""
   Label20 = ""
   '
   Text9 = ""
   Text10 = ""
   Text25 = ""
   Text11 = ""
   Text13 = ""
   Text12 = ""
   Text17 = ""
   '
   'NUMERO DE PROXIMA OM
   NRO& = 1
   Call ABREOMAN
   '
   SQLX$ = "SELECT NumeOrMa_OM FROM ORDEMANT "
   SQLX$ = SQLX$ + "ORDER BY NumeOrMa_OM DESC;"
   Set ORMANTEMP = Omantin.OpenRecordset(SQLX$, dbOpenSnapshot)
   On Error Resume Next
   With ORMANTEMP
      .MoveFirst
      If Err < 1 Then
        NRO& = 1 + !NumeOrMa_OM
      End If
   End With
   On Error GoTo 0
   '
   Label22 = FORMATNUM$(NRO&, "F9.0")
   '
   YACAR% = 0
   '
End Sub

'
Private Sub Form_Unload(Cancel As Integer)
   '
   Unload Me
   '
End Sub

Private Sub GruCli_Click()
With OPMASCL04
  .Option1(9).Value = True
  Call .Command16_Click
End With
End Sub



Private Sub LiPreCli_Click()
With OPMASCL04
  .Option1(7).Value = True
  Call .Command16_Click
End With
End Sub




Private Sub Print_Click()
   Call Command25_Click
End Sub

Private Sub PrZoVeCli_Click()
With OPMASCL04
  .Option1(5).Value = True
  Call .Command16_Click
End With
End Sub

'Sub PRILISMAT()
'    '
'    'ACA VA CODIGO PARA IMPRIMIR OBSERVACIONES
'    'Call SETULTREG("!OBSERVOF", 0)
'    '
'    ' IMPRESION DE LA LISTA DE MATERIALES
'    '
'    FORIPRE$ = CONTROL$("", "LISTAMAT")
'    '
'    If TRIM$(FORIPRE$) <> "" Then
'      '
'      Screen.MousePointer = 11
'' VER ESTO CON WALTER
'' TODA LA PARTE DE LOS PARAMETROS TENES QUE HACERLA VOS, FLAVIO
'      FECHDOC$ = Text9
'      NUMEDOC$ = TRIM$(Label31) + " / " + TRIM$(Label2)
'      TIPODOC$ = "Lista de Materiales "
'      '
'      DESTIDOC% = Val(Label22) 'CLIENTE
'      '
'      CODPARAM$(1) = "DOCUORIG" 'NUMERO DE OBRA
'      CODPARAM$(2) = "REF-MAT1" 'CODIGO DE GRUPO
'      CODPARAM$(3) = "RASO-CLI" 'CLIENTE
'      CODPARAM$(4) = "FECH-EMI" 'FECHA DE EMISION
'      CODPARAM$(5) = "REF-MAT2" 'DENOMINACION DE LA OBRA
'      CODPARAM$(6) = "REF-MAT3" 'DENOMINACION DEL GRUPO
'      CODPARAM$(7) = "TELE-CLI" 'RESPONSABLE
'
'      CAPARDOC% = 7
'
'      DOCPARAM(1) = Label31
'      DOCPARAM(2) = Label2
'      DOCPARAM(3) = Label5
'      DOCPARAM(4) = Text9
'      DOCPARAM(5) = Label7
'      DOCPARAM(6) = Label8
'      DOCPARAM(7) = Text11
'      '
' 'A PARTIR DE AQUI LO MODIFIQUE YO - EPB
'      CAITAB1% = 0
'      CAITCRU% = 0
'      '
' 'AQUI PONER LOS CODIGOS DE COLUMNAS A LISTAR
'      CODTABU1$(1) = "ORD-ITEM"
'      CODTABU1$(2) = "CANTIDAD"
'      CODTABU1$(3) = "DES-MAT"
'      CODTABU1$(4) = "N-PLANO"
'      CODTABU1$(5) = "MEDIDA-1"
'      CODTABU1$(6) = "MEDIDA-2"
'      CODTABU1$(7) = "COD-MAT"
'      CODTABU1$(8) = "MEDIDA-3"
'
'      CACOLTAB1% = 8
'      '
'      For U& = 1 To List3.ListCount
'        '
'        TTXX$ = List3.List(U& - 1)
'        CAITAB1% = CAITAB1% + 1
'        TETABU1$(1, CAITAB1%) = Left$(TTXX$, 2)
'        TETABU1$(2, CAITAB1%) = Mid$(TTXX$, 4, 6)
'        TETABU1$(3, CAITAB1%) = Mid$(TTXX$, 11, 36)
'        TETABU1$(4, CAITAB1%) = Mid$(TTXX$, 48, 14)
'        TETABU1$(5, CAITAB1%) = Mid$(TTXX$, 63, 5)
'        TETABU1$(6, CAITAB1%) = Mid$(TTXX$, 69, 12)
'        TETABU1$(7, CAITAB1%) = Mid$(TTXX$, 82, 16)
'        TETABU1$(8, CAITAB1%) = Mid$(TTXX$, 98, 2)
'        '
'        'AQUI AGREGA LA DESCRIPCION LARGA
'        DELARGA$ = Mid$(TTXX$, 129)
'        Call FRATEXTO$(DELARGA$, 36)
'        For UI% = 1 To CARETEX%
'          CAITAB1% = CAITAB1% + 1
'          TETABU1$(1, CAITAB1%) = ""
'          TETABU1$(2, CAITAB1%) = ""
'          TETABU1$(3, CAITAB1%) = RETEX$(UI%)
'          TETABU1$(4, CAITAB1%) = ""
'          TETABU1$(5, CAITAB1%) = ""
'          TETABU1$(6, CAITAB1%) = ""
'          TETABU1$(7, CAITAB1%) = ""
'          TETABU1$(8, CAITAB1%) = ""
'        Next UI%
'        '
'        'ESTO ES PARA QUE EL PROXIMO EMPIECE EN UN NUEVO RENGLON TRIPLE
'        While (CAITAB1% Mod 3) <> 0
'          CAITAB1% = CAITAB1% + 1
'          TETABU1$(1, CAITAB1%) = ""
'          TETABU1$(2, CAITAB1%) = ""
'          TETABU1$(3, CAITAB1%) = ""
'          TETABU1$(4, CAITAB1%) = ""
'          TETABU1$(5, CAITAB1%) = ""
'          TETABU1$(6, CAITAB1%) = ""
'          TETABU1$(7, CAITAB1%) = ""
'          TETABU1$(8, CAITAB1%) = ""
'        Wend
'        '
'      Next U&
'      '
'      NOCONFIRM% = 1
'      Call PRINTDOC("LISTAMAT")                   ' LISTA DE MATERIALES
'      '
'      'LIMPIA LA LISTA
'      For KK1% = 1 To CACOLTAB1%
'        For KK2% = 1 To CAITAB1%
'          TETABU1$(KK1%, KK2%) = ""
'        Next KK2%
'      Next KK1%
'      '
'    End If
'    '
'199 Call CIERRARCH("*.*")
'    Call BAJALDISCO
'    Screen.MousePointer = 1
'
'End Sub
Sub CARGAEQUI(EQ1$)
  '
  Call ABREMAQUI
  Dim Equipos As ADODB.Recordset
  SQLX$ = "SELECT * FROM EQUIPOS WHERE Codigo_Equ = '" & EQ1$ & "' "
  Set Equipos =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
  With Equipos
    If Not (.EOF And .BOF) Then
     On Error Resume Next
      Label1 = !Descri_Equ
      SEPRO1 = !Ubica_Equ
      Label7 = FORMATNUM$(SEPRO1, "F7.3")
      '
      Label15 = !Linea_Equ
      Label20 = !Posi_Equ
      Label3 = !Planta_Equ
      '
     On Error GoTo 0
    End If
  End With
  '
End Sub
Sub INDIMAN() 'INDICE DE ORDENES DE MANTENIMIENTO
  '
  Call BLANARCH("!INDIOMAN")
  Call ABREOMAN
  '
  SQLX$ = "SELECT * FROM ORDEMANT "
  SQLX$ = SQLX$ + " ORDER BY NumeOrMa_OM ASC "
  'Set OMTMP = Omantin.OpenRecordset(SQLX$, 4)
  Dim OMTMP As ADODB.Recordset
  Set OMTMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
  
  With OMTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      GoTo 99
    End If
    On Error GoTo 0
    '
    Y$ = REGBLAN$("INDIOMAN")
    JJ& = 0
    Do While .EOF = False
      NRO& = !NumeOrMa_OM
      Call REPLA(Y$, MKS$(NRO&), 1, 4)
      DES1$ = !Codigo_Equ
      DES1$ = DES1$ + " - " + FECHATEX$(CVINT(!FechGen_OM))
      Call REPLA(Y$, DES1$, 5, 44)
      JJ& = JJ& + 1
      Call GRAREG("INDIOMAN", Y$, JJ&)
      .MoveNext
    Loop
    '
  End With
  '
  Call SETULTREG("INDIOMAN", JJ&)
  Call CIERRARCH("INDIOMAN")
  '
99 End Sub
Private Sub Text24_GotFocus()
  '
  On Error Resume Next
    Command3.SetFocus
  On Error GoTo 0
  '
End Sub

Private Sub Text28_gotfocus()
  '
  On Error Resume Next
  Command3.SetFocus
  On Error GoTo 0
  '
End Sub

Private Sub Text3_Change()
  '
  NCLIE% = Val(TRIM$(Text3.TEXT))
  Label18 = RASOCLI(NCLIE% * (-1))
  '
End Sub
Private Sub Text7_GotFocus()
  '
  On Error Resume Next
  Command3.SetFocus
  On Error GoTo 0
  '
End Sub
Private Sub Text5_LOSTFOCUS()
  '
  If Val(Text5.TEXT) > 0 Then
    Text5.TEXT = REPLACAR$(Text5.TEXT, ".", ":")
    If InStr(1, Text5.TEXT, ".") = 0 Then
      If InStr(1, Text5.TEXT, ":") = 0 Then
        Text5.TEXT = Text5.TEXT + ":0"
      End If
    End If
    '
    On Error Resume Next
      D1 = CDate(Text5.TEXT)
    On Error GoTo 0
    Text5.TEXT = Format$(D1, "hh:mm")
  Else
    Text5.TEXT = Format$(Now, "hh:mm")
  End If
  '
End Sub

Private Sub Text6_lostfocus()
  '
  If Val(Text6.TEXT) > 0 Then
    Text6.TEXT = REPLACAR$(Text6.TEXT, ".", ":")
    If InStr(1, Text6.TEXT, ".") = 0 Then
      If InStr(1, Text6.TEXT, ":") = 0 Then
        Text6.TEXT = Text6.TEXT + ":0"
      End If
    End If
    '
    On Error Resume Next
      D2 = CDate(Text6.TEXT)
    On Error GoTo 0
    Text6.TEXT = Format$(D2, "hh:mm")
  Else
    Text6.TEXT = "00.00"
  End If
  '
End Sub
Sub PRIOMAN()
  ' IMPRESION DE LA O/MANTENIMIENTO
    '
    FORIPRE$ = CONTROL$("", "FORMOMAN")
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      FECHDOC$ = Text24.TEXT
      NUMEDOC$ = TRIM$(Label22)
      While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Orden de Mantenimiento"
      '
      CODPARAM$(1) = "COD-MAT"
      CODPARAM$(2) = "DES-MAT"
      CODPARAM$(3) = "REF-MAT1"
      CODPARAM$(4) = "REF-MAT2"
      CODPARAM$(5) = "REF-MAT3"
      CODPARAM$(6) = "MEDIDA-1"
      CODPARAM$(7) = "USUARIO"
      CODPARAM$(8) = "RASO-CLI"
      CODPARAM$(9) = "FECH-LAN"
      CODPARAM$(10) = "TIE-PREP"
      CODPARAM$(11) = "FECH-VEN"
      CODPARAM$(12) = "TIE-PROD"
      CAPARDOC% = 12
      '
      DOCPARAM$(1) = Text1.TEXT
      DOCPARAM$(2) = Label1
      DOCPARAM$(3) = Label3
      DOCPARAM$(4) = Label7
      DOCPARAM$(5) = Label15
      DOCPARAM$(6) = Label20
      DOCPARAM$(7) = Label24
      DOCPARAM$(8) = Label18
      DOCPARAM$(9) = Text2.TEXT
      DOCPARAM$(10) = Text5.TEXT
      DOCPARAM$(11) = Text4.TEXT
      DOCPARAM$(12) = Text6.TEXT
      '
      CODTABU1$(1) = "COD-MPRI"
      CODTABU1$(2) = "DES-MPRI"
      CODTABU1$(3) = "UNI-MPRI"
      CODTABU1$(4) = "CAN-MPRI"
      CODTABU1$(5) = "PAR-MPRI"
      CODTABU1$(6) = "NUOPERS"
      CODTABU1$(7) = "MEDIDA-2"
      '
      CACOLTAB1% = 7
      '
      CAITAB1% = 0
      CAITCRU% = 0
      '
      For J% = 0 To List1.ListCount - 1
         X$ = List1.List(J%)
         '
         CODI2$ = Mid$(X$, 1, 16)
         CI% = CODINT%(CODI2$)
         DESC2$ = DESCRIT0$(CI%)
         UM2$ = UNIMED$(CI%) ' 59, 2)
         USO2$ = Mid$(X$, 62, 10)
         CANTI = Val(USO2$)
         OBS2$ = Mid$(X$, 73, 25)
         '
         CAITAB1% = CAITAB1% + 1
         TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
         TETABU1$(2, CAITAB1%) = DESC2$
         TETABU1$(3, CAITAB1%) = UM2$
         TETABU1$(4, CAITAB1%) = CANTI
         TETABU1$(5, CAITAB1%) = OBS2$
         '
109   Next J%
      '
      CAITAB1% = 0
      For J% = 0 To List4.ListCount - 1
        CAITAB1% = CAITAB1% + 1
        TETABU1$(6, CAITAB1%) = List4.List(J%)
      Next J%
      '
      CAITAB1% = 0
      DESGRAL$ = Text28.TEXT 'DESCRIPCION GENERAL
      Call FRATEXTO(DESGRAL$, 92)
      For u& = 1 To CARETEX%
        Z$ = "." + RETEX$(u&)
        CAITAB1% = CAITAB1% + 1
        TETABU1$(7, CAITAB1%) = Z$
      Next u&
      '
      Call PRINTDOC("FORMOMAN")   ' Orden de Mantenimiento
      '
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
      DOCUORIG$ = "N/Oman-" + NUMEDOC$
      NUMEDOC$ = ""
      '
      '
     End If
199 Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
End Sub
Sub CARCONTA(MAQ1$)
  'CARGO CONTADORES DEL EQUIPO
   '
  Call ABREMAQUI
  '
   SQLX$ = " SELECT * FROM CONTADORES "
   SQLX$ = SQLX$ + " WHERE Codigo_Equ = '" & MAQ1$ & "' "
   'Set REPTMP = Maquinarias.OpenRecordset(SQLX$, 4)
   Set REPTMP =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
   With REPTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      On Error GoTo 0
      Exit Sub
    End If
    On Error GoTo 0
    '
    List3.Clear
    '
    Do While .EOF = False
      NRO% = !Num_Con
        NRO1$ = FORMATNUM$(NRO%, "F3.0")
      REFE$ = !Referen_Con
      FEQ = !Freq_Con
        FEQ1$ = FORMATNUM$(FEQ, "F7.1")
      FUL% = CVINT(!FUL_Con)
        FUL1$ = FORMATNUM$(FUL%, "D8")
      VALO = !Val_Con
        VALO1$ = FORMATNUM$(VALO, "F10.2")
        HOII% = HOY(HOS$)
      FPL% = HOII%
        FPL1$ = FORMATNUM$(FPL%, "D8")
      PROME = 0 '!Prom_Con
        PROM1$ = FORMATNUM$(PROME, "F10.2")
      OBSE$ = "" '!Observa_Con
      'AGREGO AL LIST
      L1$ = Space(256)
      Call REPLA(L1$, NRO1$, 1, 3)
      Call REPLA(L1$, REFE$, 5, 22)
      'Call REPLA(L1$, FEQ2$, 28, 7)
      Call REPLA(L1$, FUL1$, 28, 8)
      Call REPLA(L1$, VALO1$, 36, 10)
      Call REPLA(L1$, FPL1$, 47, 8)
      Call REPLA(L1$, PROM1$, 55, 10)
      Call REPLA(L1$, OBSE$, 67, 64)
      List3.AddItem L1$
      .MoveNext
    Loop
   End With
  '
99 End Sub

Private Sub Text8_LOSTFOCUS()
  '
  CT1 = Val(Text8.TEXT)
  Text8.TEXT = FORMATNUM$(CT1, "F9.2")
  '
End Sub

Private Sub Text9_Change()
  '
  NDE% = Val(Text9.TEXT)
  Label30 = ECOARCH$("TADEPOSI", MKI$(NDE%))
  '
End Sub

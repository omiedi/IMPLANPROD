VERSION 5.00
Begin VB.Form REMUORD05 
   BackColor       =   &H00E0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liquidación Ordinaria"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Barrido"
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
      Left            =   9765
      TabIndex        =   51
      Top             =   810
      Width           =   1110
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0C0&
      Caption         =   "Ctrl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   11025
      TabIndex        =   42
      Top             =   6090
      Width           =   855
      Begin VB.CommandButton Command6 
         Height          =   795
         Left            =   40
         Picture         =   "REMUORD05.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Agrega Concepto al Recibo"
         Top             =   930
         Width           =   765
      End
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
         Height          =   570
         Left            =   105
         Picture         =   "REMUORD05.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Elimina el Concepto Activo del Recibo"
         Top             =   315
         Width           =   660
      End
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
      Height          =   315
      Left            =   2250
      Picture         =   "REMUORD05.frx":0614
      TabIndex        =   32
      Top             =   390
      Width           =   175
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1470
      MaxLength       =   5
      TabIndex        =   31
      Text            =   " "
      Top             =   390
      Width           =   780
   End
   Begin VB.TextBox Text3 
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
      Height          =   315
      Left            =   1470
      TabIndex        =   30
      Top             =   810
      Width           =   5920
   End
   Begin VB.Frame Frame10 
      Caption         =   "Frame10"
      Height          =   4605
      Left            =   6100
      TabIndex        =   27
      Top             =   1300
      Width           =   15
   End
   Begin VB.Frame Frame9 
      Caption         =   "Frame9"
      Height          =   4605
      Left            =   9030
      TabIndex        =   26
      Top             =   1300
      Width           =   15
   End
   Begin VB.Frame Frame8 
      Caption         =   "Frame8"
      Height          =   4605
      Left            =   7455
      TabIndex        =   25
      Top             =   1300
      Width           =   15
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   4605
      Left            =   1050
      TabIndex        =   24
      Top             =   1300
      Width           =   15
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00800000&
      Height          =   5960
      Left            =   11040
      ScaleHeight     =   5895
      ScaleWidth      =   1215
      TabIndex        =   15
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command16 
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
         Height          =   650
         Left            =   105
         Picture         =   "REMUORD05.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   2850
         Width           =   615
      End
      Begin VB.CommandButton Command17 
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
         Height          =   650
         Left            =   105
         Picture         =   "REMUORD05.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   3500
         Width           =   615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "REMUORD05.frx":106A
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   4305
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         Height          =   650
         Left            =   105
         Picture         =   "REMUORD05.frx":1374
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Activa Calculadora del Sistema (Si Disponible)"
         Top             =   1560
         Width           =   615
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   555
         TabIndex        =   28
         Top             =   5040
         Width           =   615
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command19 
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
         Height          =   650
         Left            =   105
         Picture         =   "REMUORD05.frx":167E
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Cierra y abandona la Aplicación"
         Top             =   250
         Width           =   615
      End
      Begin VB.CommandButton Command18 
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
         Height          =   660
         Left            =   105
         Picture         =   "REMUORD05.frx":17C8
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   900
         Width           =   615
      End
      Begin VB.CommandButton Command39 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   105
         Picture         =   "REMUORD05.frx":1AD2
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Imprime Recibo de Haberes"
         Top             =   2200
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "REMUORD05.frx":213C
         Top             =   5355
         Width           =   1515
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   5070
      Left            =   120
      TabIndex        =   5
      Top             =   1400
      Width           =   10995
      Begin VB.PictureBox Picture1 
         Height          =   460
         Left            =   0
         ScaleHeight     =   405
         ScaleWidth      =   10740
         TabIndex        =   11
         Top             =   4535
         Width           =   10805
         Begin VB.Line Line3 
            BorderWidth     =   2
            X1              =   3465
            X2              =   3465
            Y1              =   0
            Y2              =   420
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   1785
            TabIndex        =   52
            Top             =   60
            Width           =   1500
         End
         Begin VB.Label Label10 
            Caption         =   "Neto a Cobrar:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   105
            TabIndex        =   53
            Top             =   90
            Width           =   1905
         End
         Begin VB.Label Label8 
            Caption         =   "Totales:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4935
            TabIndex        =   47
            Top             =   90
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   8925
            TabIndex        =   40
            Top             =   60
            Width           =   1500
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   7355
            TabIndex        =   13
            Top             =   60
            Width           =   1500
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   5985
            TabIndex        =   12
            Top             =   60
            Width           =   1290
         End
      End
      Begin VB.Frame Frame14 
         Caption         =   "Frame10"
         Height          =   4605
         Left            =   10450
         TabIndex        =   39
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame13 
         Caption         =   "Frame10"
         Height          =   3975
         Left            =   0
         TabIndex        =   38
         Top             =   5500
         Width           =   15
      End
      Begin VB.Frame Frame11 
         Caption         =   "Frame10"
         Height          =   4605
         Left            =   5270
         TabIndex        =   37
         Top             =   0
         Width           =   15
      End
      Begin VB.PictureBox Picture5 
         BackColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   0
         ScaleHeight     =   390
         ScaleWidth      =   10740
         TabIndex        =   21
         Top             =   0
         Width           =   10805
         Begin VB.Label Label27 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Código     Descripción                                                     %        Cantidad         Haberes            Deducc."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   25
            TabIndex        =   22
            Top             =   105
            Width           =   11355
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4080
         Left            =   20
         MultiSelect     =   2  'Extended
         TabIndex        =   20
         Top             =   450
         Width           =   10800
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   10815
         Y1              =   5000
         Y2              =   5000
      End
      Begin VB.Line Line2 
         X1              =   10800
         X2              =   10800
         Y1              =   0
         Y2              =   4995
      End
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0C0&
      Caption         =   "Registrar / Editar Concepto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Left            =   105
      TabIndex        =   6
      Top             =   6510
      Width           =   10800
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Deducc."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   9030
         TabIndex        =   45
         Top             =   420
         Width           =   1710
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   105
            MaxLength       =   11
            TabIndex        =   46
            Text            =   " "
            Top             =   360
            Width           =   1500
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Haberes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   7245
         TabIndex        =   10
         Top             =   420
         Width           =   1710
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   105
            MaxLength       =   11
            TabIndex        =   3
            Text            =   " "
            Top             =   360
            Width           =   1500
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Cantidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   5670
         TabIndex        =   9
         Top             =   420
         Width           =   1500
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   105
            MaxLength       =   11
            TabIndex        =   2
            Top             =   360
            Width           =   1290
         End
      End
      Begin VB.CommandButton Command31 
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
         Left            =   1785
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text11 
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
         TabIndex        =   0
         Top             =   420
         Width           =   750
      End
      Begin VB.Label Label4 
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
         Left            =   210
         TabIndex        =   8
         Top             =   855
         Width           =   5280
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Left            =   -30
         TabIndex        =   7
         Top             =   480
         Width           =   960
      End
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   48
      Top             =   390
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro de Legajo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   210
      TabIndex        =   36
      Top             =   315
      Width           =   1140
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido y Nombre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   0
      TabIndex        =   35
      Top             =   750
      Width           =   1350
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo:"
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
      Left            =   2205
      TabIndex        =   34
      Top             =   495
      Width           =   1110
   End
   Begin VB.Label Label20 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4035
      TabIndex        =   33
      Top             =   390
      Width           =   3360
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   11100
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Ltext1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9765
      TabIndex        =   23
      Top             =   435
      Width           =   1110
   End
   Begin VB.Label Label2 
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
      Height          =   330
      Left            =   8340
      TabIndex        =   4
      Top             =   480
      Width           =   1275
   End
End
Attribute VB_Name = "REMUORD05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EMPL%
Dim HUBOCAMBIOS%
' dimensionamiento de vectores
Dim CTA$(16), DEX$(16), CAN(16), IMPO#(16), SIG%(16)
Dim CTAA%(128), CANTI(128), IM#(128), INFO%(128)
Dim FERI1%(20)
'
'CONSTANT RELIMI# = 1000     ' LIMITE PARA EL CAMBIO PORCENTAJE ANSAL

Dim SB(100), QT(100), RM#(100), RT#(100), PI(100), RINTER#(100)
Dim CONOPCI$, ASIGFAM$
Dim ACUE ' AÑOS DE ANTIGUEDAD DE CALCULO
'
Dim NUEMPL%, LEGA%(1024), ACMEN#(1024), RTGAN#(1024), PULIQ$(1024)
Dim VALESC(16)

Private Sub Command11_Click()

End Sub

Private Sub Command1_Click()
    LEGAX% = Val(Text1)
    If ECOARCH$("PADREMP", MKI$(LEGAX%)) <> "" Then
      Call FILTERDELETE("DETLIQUI", 1, MKI$(LEGAX%))
      '
      REBLA$ = REGBLAN$("DETLIQUI")
      For KU& = 1 To List1.ListCount
        TTXX$ = List1.List(KU& - 1)
        CCON% = Val(Left$(TTXX$, 5))
        If CCON% > 0 Then
          If CCON% <= 599 Then
            CANTX = XVALO(Mid$(TTXX$, 52, 8))
            IMPOX# = XVALO(Mid$(TTXX$, 61, 12)) + XVALO(Mid$(TTXX$, 74, 12))
            If Abs(CANTX) >= 0.005 Or Abs(IMPOX#) >= 0.005 Then
              YY$ = REBLA$
              Call REPLA(YY$, MKI$(LEGAX%), 1, 2)
              Call REPLA(YY$, MKI$(CCON%), 3, 2)
              Call REPLA(YY$, MKS$(CANTX), 5, 4)
              Call REPLA(YY$, MKD$(IMPOX#), 9, 8)
              Call REGAPP("DETLIQUI", YY$)
            End If
          End If
        End If
      Next KU&
      Call CIERRARCH("DETLIQUI")
      '
    End If
End Sub
'
Private Sub Command13_Click()
   'Call FRESHECHO("PADREMP")
   Call SELECHO("PADREMP")
   NLEG% = Val(TRIM$(VALACT1$("PADREMP")))
   If NLEG% < 1 Then Exit Sub
   '
   NC% = NLEG%
   Text1.TEXT = TRIM$(Str$(NC%))
   '
   Text11.SetFocus
   '
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   If HUBOCAMBIOS% > 0 Then
     If COMALTER%("Desea Descartar los Cambios Realizados ?\\No, Conservar\Si, Descartar") <> 2 Then
       GoTo 99
     End If
     HUBOCAMBIOS% = 0
   End If
   Call CIERRARCH("*.*")
   Cancel = False
   Unload Me
   '
99 Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    Call BARRELIQ
End Sub

'
Private Sub Command31_Click()
  '
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  Call SELECHO("ECONCEP")
  CONCEP% = Val(TRIM$(VALACT1$("ECONCEP")))
  If CONCEP% < 1 Then Exit Sub
  '
  Text11.TEXT = TRIM$(Str$(CONCEP%))
  '
  If Text14.Enabled = True Then
      Text14.SetFocus
    ElseIf Text15.Enabled = True Then
      Text15.SetFocus
    ElseIf Text2.Enabled = True Then
      Text2.SetFocus
    Else
      Command6.SetFocus
  End If
  '
End Sub

Private Sub Command39_Click()
  '
  Command39.Enabled = False
  If List1.ListCount > 0 Then 'si tiene liquidación
    Call Command1_Click      ' PARA GUARDAR ANTES DE IMPRIMIR
    NULEG1% = XVALO(Text1.TEXT)
    If NULEG1% > 0 Then
      PRINUME& = 1 + XVALO(CONTROL$("NUMERECI", "UNURECI"))
2     NUMERECI& = XVALO(InputBox$("Número Primer Recibo", "Numeración de Recibos de Haberes", 1))
      If NUMERECI& < 1 Or NUMERECI& > 900000 Then GoTo 2
      NUMEDOC$ = TRIM$(FORMATNUM$(NUMERECI&, "F8.0"))
      Call GRACONTROL("NUMERECI", "UNURECI", NUMEDOC$)
      While Len(NUMEDOC$) < 6: NUMEDOC$ = "0" + NUMEDOC$: Wend
      Call INISUEL.PRIRECISUEL(NULEG1%)
    End If
  End If
  Command39.Enabled = True
  '
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   KU& = List1.ListIndex
   If KU& >= 0 Then
     If List1.Selected(KU&) = True Then
       CODII% = XVALO(Left$(List1.List(KU&), 5))
       List1.RemoveItem KU&
       If CODII% < 100 Then
         RM#(CODII%) = 0
         QT(CODII%) = 0
       End If
       Call RECALCONCE
       If KU& < List1.ListCount Then
         List1.ListIndex = KU&
       End If
     End If
   End If
   Command4.Enabled = True
End Sub

Private Sub Command6_Click()
   CODACT% = XVALO(Text11)
   If CODACT% > 0 Then
     For KU& = 1 To List1.ListCount
       If Val(Left$(List1.List(KU& - 1), 5)) = CODACT% Then
         List1.RemoveItem KU& - 1
       End If
     Next KU&
     '
     If CODACT% > 0 Then
       If CODACT% <= 100 Then
         RM#(CODACT%) = Val(Text15)
         QT(CODACT%) = Val(Text14)
       End If
     End If
     '
     If CODACT% > 200 Then
       If CODACT% <= 299 Then
         RT#(CODACT% - 200) = Val(Text2)
       End If
     End If
     '
     ' ASIGNACIONES FAMILIARES
     If CODACT% > 400 Then
       If CODACT% <= 499 Then
         NASI% = CODACT% - 400
         IMPOX# = XVALO(Text14) * VALASI(NASI%) * ESTASI(NASI%)
         Text15 = TRIM$(FORMATNUM$(IMPOX#, "F12.2"))
       End If
     End If
     '
     If Val(Text14) > 0 Or Val(Text15) > 0 Or Val(Text2) > 0 Then
       TTXX$ = Space$(96)
       Call REPLA(TTXX$, FORMATNUM$(CODACT%, "I5"), 1, 5)
       Call REPLA(TTXX$, TRIM$(ECOARCH$("ECONCEP", MKI$(CODACT%))), 9, 36)
       Call REPLA(TTXX$, FORMATNUM$(Val(Text14), "F8.2"), 52, 8)
       Call REPLA(TTXX$, FORMATNUM$(Val(Text15), "F12.2"), 61, 12)
       Call REPLA(TTXX$, FORMATNUM$(Val(Text2), "F12.2"), 74, 12)
       '
       For KU& = List1.ListCount To 1 Step -1
         If Val(Left$(List1.List(KU& - 1), 5)) < CODACT% Then
           List1.AddItem TTXX$, KU&
           GoTo 19
         End If
       Next KU&
       List1.AddItem TTXX$, 0
     End If
     '
19   Call RECALCONCE
     On Error Resume Next
     Text11.SetFocus
     On Error GoTo 0
     '
   End If
   '
End Sub

Private Sub Command7_Click()
   If EXISTE%("C:\WINDOWS\CALC.EXE") > 0 Then
        Shell "C:\WINDOWS\CALC.EXE", 4
      Else
        Call MENSERR(24, "No se Encuentra Utilitario\'CALC.EXE'\en Carpeta 'WINDOWS' del Disco 'C:'.")
   End If
End Sub

Private Sub Form_Load()
   '
   HOII% = HOY(HOS$)
   Ltext1 = HOS$
   '
   ' CARGA DATOS DE LIQUIDACION ACTIVA
   XX$ = REGLEIDO$("ULTILIQ", 1)
   If TRIM$(Left$(XX$, 30)) = "" Then GoTo 90
   '
   GRULI% = Asc(Mid$(XX$, 31, 1))
   TILIQ% = Asc(Mid$(XX$, 34, 1))
   ASIFLI% = Asc(Mid$(XX$, 35, 1))
   '
   If ULTREG&("DETLIQUI") < 1 Then
     Call BARRELIQ
   End If
   Exit Sub
   '
90 Call MENSERR(24, "Falta Inicializar Liquidación")
   Unload Me
  '
End Sub

Sub BARRELIQ()
     '
     Call CARGLIQACT          ' datos de liquidacion activa
     Call CARGAFORMULAS       ' CARGAR CONCEPTOS y RETENCIONES
     JJ& = 0
     REBLA$ = REGBLAN$("LINFHOR")
     '
1000 For KU& = 1 To ULTREG&("PADREMP")
       PP0$ = REGLEIDO$("PADREMP", KU&)
       LEGAX% = CVI(Left$(PP0$, 2))
       Call CARGEMP(LEGAX%)
       If STAT_LEG% > 0 Then  ' ESTADO ACTIVO
         If GRUP_LEG% = GRUP_LIQ% Then
           '
1010       Dim CANTID(1024), IMPOA#(1024)
           For KKU& = 0 To 1024
              CANTID(KKU&) = 0
              IMPOA#(KKU&) = 0
           Next KKU&
           RFF$ = RECFILT$("DETLIQUI", 1, MKI$(LEGAX%))
           For KKU& = 1 To Len(RFF$) Step 4
              RELI& = CVS(Mid$(RFF$, KKU&, 4))
              XX1$ = REGLEIDO$("DETLIQUI", RELI&)
              ICODI% = CVI(Mid$(XX1$, 3, 2))
              CANTID(ICODI%) = CVS(Mid$(XX1$, 5, 4))
              IMPOA#(ICODI%) = CVD(Mid$(XX1$, 9, 8))
           Next KKU&
           '
1020       LIBARRI.Label17 = TRIM$(Str$(LEGAX%))
           LIBARRI.Label16 = APNO_LEG$
           LIBARRI.Label15 = TRIM$(Str$(GRUP_LEG%))
           LIBARRI.Label20 = ECOARCH("DATGREM1", MKI$(GRUP_LEG%))
           NURECON% = 0
           '
           For CONI% = 1 To NUCON%
             If ESTCON%(CONI%) < 1 Then GoTo 8230 ' DESCARTO SUSPENDIDOS
             If FORCON%(CONI%) < 1 Then GoTo 8230 ' DESCARTO CONCEPTOS DE CALCULO
             If QLCON%(CONI%) <> 0 Then If QLCON%(CONI%) <> TIPO_LIQ% Then GoTo 8230
             If CARCON%(CONI%) = 3 Or CARCON%(CONI%) = GRUP_LIQ% Then GoTo 8220 ' MANDATORIO
             For X% = 1 To 96 Step 2
               X1% = CVI(Mid$(COPC_LEG$, X%))
               If X1% = CONI% Then GoTo 8220           ' PREASIGNADO
             Next X%
             GoTo 8230
             '
8220         NURECON% = NURECON% + 1
             If LIBARRI.Label1.Count < NURECON% Then
               Load LIBARRI.Label1(NURECON% - 1)
               Load LIBARRI.Label2(NURECON% - 1)
               Load LIBARRI.Text1(NURECON% - 1)
               Load LIBARRI.Text2(NURECON% - 1)
               Load LIBARRI.Text3(NURECON% - 1)
             End If
             '
             LIBARRI.Label1(NURECON% - 1).Left = LIBARRI.Label1(0).Left
             LIBARRI.Label2(NURECON% - 1).Left = LIBARRI.Label2(0).Left
             LIBARRI.Text1(NURECON% - 1).Left = LIBARRI.Text1(0).Left
             LIBARRI.Text2(NURECON% - 1).Left = LIBARRI.Text2(0).Left
             LIBARRI.Text3(NURECON% - 1).Left = LIBARRI.Text3(0).Left
             TOPE = LIBARRI.Text1(0).Top + 315 * (NURECON% - 1)
             LIBARRI.Label1(NURECON% - 1).Top = TOPE + 30
             LIBARRI.Label2(NURECON% - 1).Top = TOPE + 30
             LIBARRI.Text1(NURECON% - 1).Top = TOPE
             LIBARRI.Text2(NURECON% - 1).Top = TOPE
             LIBARRI.Text3(NURECON% - 1).Top = TOPE
             LIBARRI.Label1(NURECON% - 1).Visible = True
             LIBARRI.Label2(NURECON% - 1).Visible = True
             LIBARRI.Label1(NURECON% - 1).Caption = TRIM$(Str$(CONI%))
             LIBARRI.Label2(NURECON% - 1).Caption = ECOARCH$("ECONCEP", MKI$(CONI%))
               LIBARRI.Text1(NURECON% - 1) = ""
               LIBARRI.Text2(NURECON% - 1) = ""
               LIBARRI.Text3(NURECON% - 1) = ""
             If FORCON%(CONI%) = 2 Then
                 LIBARRI.Text1(NURECON% - 1).Visible = True
                 LIBARRI.Text1(NURECON% - 1).TEXT = FORMATNUM(CANTID(CONI%), "F8.2")
               ElseIf FORCON%(CONI%) = 3 Then
                 LIBARRI.Text2(NURECON% - 1).Visible = True
                 LIBARRI.Text2(NURECON% - 1).TEXT = FORMATNUM(CANTID(CONI%), "F8.2")
               Else
                 LIBARRI.Text3(NURECON% - 1).Visible = True
                 LIBARRI.Text3(NURECON% - 1).TEXT = FORMATNUM(Abs(IMPOA#(CONI%)), "F10.2")
             End If
8230       Next CONI%
           '
           For CONI% = 1 To NURET%
             If ESTRET%(CONI%) < 1 Then GoTo 8270 ' DESCARTO SUSPENDIDOS
             If FORRET%(CONI%) <> 1 Then GoTo 8270 ' DESCARTO CONCEPTOS DE CALCULO
             If QLRET%(CONI%) <> 0 Then If QLRET%(CONI%) <> TIPO_LIQ% Then GoTo 8270
             If CARRET%(CONI%) = 3 Or CARRET%(CONI%) = GRUP_LIQ% Then GoTo 8260 ' MANDATORIO
             For X% = 1 To 96 Step 2
               X1% = CVI(Mid$(COPC_LEG$, X%))
               If X1% = 200 + CONI% Then GoTo 8260         ' PREASIGNADO
             Next X%
             GoTo 8270
             '
8260         NURECON% = NURECON% + 1
             If LIBARRI.Label1.Count < NURECON% Then
               Load LIBARRI.Label1(NURECON% - 1)
               Load LIBARRI.Label2(NURECON% - 1)
               Load LIBARRI.Text1(NURECON% - 1)
               Load LIBARRI.Text2(NURECON% - 1)
               Load LIBARRI.Text3(NURECON% - 1)
             End If
             '
             LIBARRI.Label1(NURECON% - 1).Left = LIBARRI.Label1(0).Left
             LIBARRI.Label2(NURECON% - 1).Left = LIBARRI.Label2(0).Left
             LIBARRI.Text1(NURECON% - 1).Left = LIBARRI.Text1(0).Left
             LIBARRI.Text2(NURECON% - 1).Left = LIBARRI.Text2(0).Left
             LIBARRI.Text3(NURECON% - 1).Left = LIBARRI.Text3(0).Left
             TOPE = LIBARRI.Text1(0).Top + 315 * (NURECON% - 1)
             LIBARRI.Label1(NURECON% - 1).Top = TOPE + 30
             LIBARRI.Label2(NURECON% - 1).Top = TOPE + 30
             LIBARRI.Text1(NURECON% - 1).Top = TOPE
             LIBARRI.Text2(NURECON% - 1).Top = TOPE
             LIBARRI.Text3(NURECON% - 1).Top = TOPE
             LIBARRI.Label1(NURECON% - 1).Visible = True
             LIBARRI.Label2(NURECON% - 1).Visible = True
             LIBARRI.Label1(NURECON% - 1).Caption = TRIM$(Str$(200 + CONI%))
             LIBARRI.Label2(NURECON% - 1).Caption = ECOARCH$("ECONCEP", MKI$(200 + CONI%))
               LIBARRI.Text1(NURECON% - 1) = ""
               LIBARRI.Text2(NURECON% - 1) = ""
               LIBARRI.Text3(NURECON% - 1) = ""
             If FORRET%(CONI%) = 1 Then
               LIBARRI.Text3(NURECON% - 1).Visible = True
               LIBARRI.Text3(NURECON% - 1).TEXT = FORMATNUM(Abs(IMPOA#(200 + CONI%)), "F10.2")
             End If
8270       Next CONI%
           '
           LIBARRI.Height = LIBARRI.Label1(NURECON% - 1).Top + 900
           LIBARRI.Image2.Top = LIBARRI.Height - 650 - LIBARRI.Image2.Height
           LIBARRI.Command9.Top = LIBARRI.Image2.Top - 180 - LIBARRI.Command9.Height
           Call CENTRAFORM(LIBARRI)
           '
           For KU1& = 0 To 100
             QT(KU1&) = 0
             RM#(KU1&) = 0
             RT#(KU1&) = 0
           Next KU1&
           '
           LIBARRI.Show 1
           Call FRESHALL
           OPENFORMS = DoEvents
           If LIBARRI.Label11.Caption <> "OK" Then
             If COMALTER%("¿ Interrumpe Liquidación por Barrido ?\\No, Continuar\Si, Cancelar") = 2 Then
               Exit Sub
             End If
             GoTo 1010
           End If
           '
           ' AQUI SE APROBÓ
           Call FILTERDELETE("DETLIQUI", 1, MKI$(LEGAX%))
           Text1 = TRIM$(Str$(LEGAX%))
           Text3 = APNO_LEG$
           Label15 = TRIM$(Str$(GRUP_LEG%))
           Label20 = ECOARCH("DATGREM1", MKI$(GRUP_LEG%))
           '
           ULI& = 1 + LIBARRI.Label1.UBOUND
           For KU1& = 1 To ULI&
             COCONI% = Val(LIBARRI.Label1(KU1& - 1).Caption)
             QTCONI = Val(LIBARRI.Text1(KU1& - 1).TEXT) + Val(LIBARRI.Text2(KU1& - 1).TEXT)
             IMCONI# = Val(LIBARRI.Text3(KU1& - 1).TEXT)
             If Abs(QTCONI) >= 0.005 Or Abs(IMCONI#) >= 0.005 Then
               Text11 = TRIM$(Str$(COCONI%))
               Text14 = FORMATNUM(QTCONI, "F10.2")
               TTZZ$ = FORMATNUM(IMCONI#, "F12.2")
               If Abs(IMCONI#) >= 0.005 Then
                  If COCONI% <= 200 Or COCONI% > 400 Then
                        Text15 = TTZZ$
                        Text2 = ""
                      Else
                        Text15 = ""
                        Text2 = TTZZ$
                  End If
               End If
               Call Command6_Click
             End If
             LIBARRI.Text1(KU1& - 1).TEXT = ""
             LIBARRI.Text2(KU1& - 1).TEXT = ""
             LIBARRI.Text3(KU1& - 1).TEXT = ""
           Next KU1&
           Call Command1_Click ' PARA GRABAR
           '
           Call RECALCONCE
           Unload LIBARRI
           '
         End If
       End If
29   Next KU&
     '
     Call COMUNI("Barrida Completa")
     '
End Sub
'
Sub BARRELIQVIE()
     '
     Call CARGLIQACT          ' datos de liquidacion activa
     Call CARGAFORMULAS       ' CARGAR CONCEPTOS y RETENCIONES
     JJ& = 0
     REBLA$ = REGBLAN$("LINFHOR")
     '
1000 For KU& = 1 To ULTREG&("PADREMP")
       PP0$ = REGLEIDO$("PADREMP", KU&)
       LEGAX% = CVI(Left$(PP0$, 2))
       Call CARGEMP(LEGAX%)
       If STAT_LEG% > 0 Then  ' ESTADO ACTIVO
         If GRUP_LEG% = GRUP_LIQ% Then
           '
           Text1 = TRIM$(Str$(LEGAX%))
           NOAP$ = APNO_LEG$
           CONTA% = 0
           For CONI% = 1 To NUCON%
             If ESTCON%(CONI%) < 1 Then GoTo 8230 ' DESCARTO SUSPENDIDOS
             If FORCON%(CONI%) < 1 Then GoTo 8230 ' DESCARTO CONCEPTOS DE CALCULO
             If QLCON%(CONI%) <> 0 Then If QLCON%(CONI%) <> TIPO_LIQ% Then GoTo 8230
             If CARCON%(CONI%) = 3 Or CARCON%(CONI%) = GRUP_LIQ% Then GoTo 8220 ' MANDATORIO
             For X% = 1 To 96 Step 2
               X1% = CVI(Mid$(COPC_LEG$, X%))
               If X1% = CONI% Then GoTo 8220           ' PREASIGNADO
             Next X%
             GoTo 8230
             '
8220         CONTA% = CONTA% + 1
             CTAA%(CONTA%) = CONI%: CANTI(CONTA%) = 0: IM#(CONTA%) = 0: RM#(CONI%) = 0: INFO%(CONTA%) = 0
             '
8230       Next CONI%
           '
           For CONI% = 1 To NURET%
             If ESTRET%(CONI%) < 1 Then GoTo 8270 ' DESCARTO SUSPENDIDOS
             If FORRET%(CONI%) <> 1 Then GoTo 8270 ' DESCARTO CONCEPTOS DE CALCULO
             If QLRET%(CONI%) <> 0 Then If QLRET%(CONI%) <> TIPO_LIQ% Then GoTo 8270
             If CARRET%(CONI%) = 3 Or CARRET%(CONI%) = GRUP_LIQ% Then GoTo 8260 ' MANDATORIO
             For X% = 1 To 96 Step 2
               X1% = CVI(Mid$(COPC_LEG$, X%))
               If X1% = 200 + CONI% Then GoTo 8260         ' PREASIGNADO
             Next X%
             GoTo 8270
             '
8260         CONTA% = CONTA% + 1
             CTAA%(CONTA%) = 200 + CONI%: CANTI(CONTA%) = 0: IM#(CONTA%) = 0: INFO%(CONTA%) = 0
             '
8270       Next CONI%
           '
           For KKII% = 1 To 4
             YY$ = REBLA$
             Call REPLA(YY$, MKI$(LEGAI%), 1, 2)
             Call REPLA(YY$, MKI$(0), 3, 2)
             Call REPLA(YY$, String$(30, "."), 5, 30)
             Call REPLA(YY$, String$(10, "."), 37, 10)
             Call REPLA(YY$, String$(10, "."), 49, 10)
             Call REPLA(YY$, String$(10, "."), 61, 12)
             Call REPLA(YY$, String$(64, "."), 75, 64)
             JJ& = JJ& + 1
             Call GRAREG("LINFHOR", YY$, JJ&)
           Next KKII%
           '
           For KU1& = 0 To 100
             QT(KU1&) = 0
             RM#(KU1&) = 0
             RT#(KU1&) = 0
           Next KU1&
           '
           REBLA$ = REGBLAN$("CARGALIQ")
           ULI& = CONTA%
           For KU1& = 1 To ULI&
             YY$ = REBLA$
             Call REPLA(YY$, MKI$(CTAA%(KU1&)), 1, 2)
             Call GRAREG("!CARGALIQ", YY$, KU1&)
           Next KU1&
           Call SETULTREG("!CARGALIQ", ULI&)
           '
           ATRI$ = "/NC"
           ATRI$ = ATRI$ + "/TITUPAN=Liquidación de Horas: " + NOAP$
           ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame4.Top + 40))
           ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame4.Left + 20))
           '
           VTB% = VENTABU%("CARGALIQ" + ATRI$)
           If VTB% < 0 Then
             If COMALTER%("¿ Interrumpe Liquidación por Barrido ?\\No, Continuar\Si, Cancelar") = 2 Then
               Exit Sub
             End If
             GoTo 29
           End If
           '
           If VTB% >= 0 Then
             ULI& = ULTREG&("!CARGALIQ")
             For KU1& = 1 To ULI&
               YY$ = REGLEIDO$("!CARGALIQ", KU1&)
               COCONI% = CVI(Left$(YY$, 2))
               QTCONI = CVS(Mid$(YY$, 3, 4))
               IMCONI# = CVD(Mid$(YY$, 7, 8))
               If Abs(QTCONI) >= 0.005 Or Abs(IMCONI#) >= 0.005 Then
                 Text11 = TRIM$(Str$(COCONI%))
                 TTZZ$ = FORMATNUM(IMCONI#, "F12.2")
                 If COCONI% <= 200 Or COCONI% > 400 Then
                     Text15 = TTZZ$
                     Text2 = ""
                   Else
                     Text15 = ""
                     Text2 = TTZZ$
                 End If
                 Text14 = FORMATNUM(QTCONI, "F10.2")
                 Call Command6_Click
               End If
             Next KU1&
             Call Command1_Click ' PARA GRABAR
           End If
           '
           Call RECALCONCE
           '
         End If
       End If
29   Next KU&
     '
     Call COMUNI("Barrida Completa")
     '
End Sub
'
Sub CARVECLEGA()
    Dim FX1 As String * 64
    N5% = FILEOPEN%(LUDAT$ + "PADRON.DAT", 0, 64)
    '
    For I% = 1 To LOF(N5%) / 64
      Get #N5%, I%, FX1$
      If CVI(FX1$) = 0 Then
        Close #N5%
        Exit Sub
      End If
      If I% > 1024 Then
        Call MENSERR(24, "Vector de Empleados fuera de Rango - Consulte")
        Call FINAL("")
      End If
      NUEMPL% = I%
      LEGA%(I%) = CVI(Left$(FX1$, 2))
    Next I%
    '
    Close #N5%
    '
End Sub

Private Sub Label1_Change()
    Call Label22_Change
End Sub

'
Private Sub Label15_CHANGE()
  '
  If Val(Text1) > 32767 Then
    Call COMUNI("Número de Legajo no Válido.")
    Text1.SelStart = 0
    Text1.SelLength = (Len(Text1))
    Exit Sub
  End If
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  '
  Label20 = ECOARCH("DATGREM1", MKI$(Val(Label15.Caption)))
  MODIFIC% = 1
  '
End Sub

Private Sub Label22_Change()
    NECOBRA# = XVALO(Label22) - XVALO(Label1)
    Label9 = FORMATNUM(NECOBRA#, "F11.2")
End Sub

Private Sub List1_Click()
   Text11 = TRIM$(Left$(List1.TEXT, 5))
End Sub

Private Sub Text1_Change()
  '
  List1.Clear
  Text11 = ""
  Text14 = ""
  Text15 = ""
  Text2 = ""
  '
  Label21 = ""
  Label22 = ""
  Label1 = ""
  '
  RM0# = 0
  For KKI% = 0 To 100
     RM#(KKI%) = 0
     QT(KKI%) = 0
  Next KKI%
  '
  If Val(Text1) > 32767 Then
    Exit Sub
  End If
  '
  LEGAX% = Val(Text1)
  Call CARGEMP(LEGAX%)
  Call CARDATEMP(LEGAX%)
  Call CARDETLIQ(LEGAX%)
  '
End Sub

Sub CARDETLIQ(LEGAX%)
   List1.Clear
   If ECOARCH$("PADREMP", MKI$(LEGAX%)) <> "" Then
     RFF$ = RECFILT$("DETLIQUI", 1, MKI$(LEGAX%))
     For KU& = 1 To Len(RFF$) Step 4
       U& = CVS(Mid$(RFF$, KU&, 4))
       YY$ = REGLEIDO$("DETLIQUI", U&)
       '
       CODACT% = CVI(Mid$(YY$, 3, 2))
       CANTX = CVS(Mid$(YY$, 5, 4))
       IMPOX# = CVD(Mid$(YY$, 9, 8)): IMPOX2# = 0
       If CODACT% >= 200 And CODACT% <= 399 Then
         IMPOX2# = IMPOX#: IMPOX# = 0
       End If
       '
       TTXX$ = Space$(96)
       Call REPLA(TTXX$, FORMATNUM$(CODACT%, "I5"), 1, 5)
       Call REPLA(TTXX$, TRIM$(ECOARCH$("ECONCEP", MKI$(CODACT%))), 9, 36)
       Call REPLA(TTXX$, FORMATNUM$(CANTX, "F8.2"), 52, 8)
       Call REPLA(TTXX$, FORMATNUM$(IMPOX#, "F12.2"), 61, 12)
       Call REPLA(TTXX$, FORMATNUM$(IMPOX2#, "F12.2"), 74, 12)
       List1.AddItem TTXX$
       '
       If CODACT% > 0 Then
         If CODACT% <= 100 Then
           RM#(CODACT%) = IMPOX#
           QT(CODACT%) = CANTX
         End If
       End If
       '
     Next KU&
     If List1.ListCount < 1 Then
        Call RECALCONCE
     End If
     '
     Call SUMAREMU
   End If
End Sub
'
Sub CARDATEMP(LEGAX%) 'CARGA DE EMPLEADOS
    '
    Static N7%
    Dim RBPI As String * 128
    '
    If NUEMPL% < 1 Then
      Call CARVECLEGA
    End If
    Call CARGAFORMULAS
    '
    XX1$ = RECFILT$("PADREMP", 1, MKI$(LEGAX%))
    J1& = CVS(XX1$)
    '
    XX$ = REGLEIDO$("PADREMP", J1&)
    NOAP$ = TRIM$(Mid$(XX$, 3, 30))
    GRU% = Asc(Mid$(XX$, 49, 1))
    Text1 = TRIM$(Str$(LEGAX%))
    Text3.TEXT = NOAP$
    Label15.Caption = Str$(GRU)
    '
    XX$ = REGLEIDO$("ACTIVA", J1&)
    CONOPCI$ = Left$(XX$, 96)
    ASIGFAM$ = Mid$(XX$, 97, 32)
    '
    ENCU% = 0
    For VX% = 1 To NUREM%: SB(VX%) = 0: Next VX%
    For VX% = 1 To NUPAR%: PI(VX%) = 0: Next VX%
    For VX% = 1 To NUPAR%: QT(VX%) = 0: Next VX%
    For VX% = 1 To NUPAR%: RM#(VX%) = 0: Next VX%
    For VX% = 1 To NUPAR%: RT#(VX%) = 0: Next VX%
    For EMPL% = 1 To NUEMPL%
      If LEGA%(EMPL%) = LEGAX% Then ENCU% = 1: GoTo 70
    Next EMPL%
    Exit Sub
    '
70  RBPI$ = REGLEIDO$("REMBAS", J1&)
    RB1$ = Left$(RBPI$, 48)
    PI1$ = Mid$(RBPI$, 49, 80)
    '
    For VX% = 1 To NUREM%
      SB(VX%) = CVS(Mid$(RB1$, 4 * VX% - 3))
    Next VX%
    '
    For VX% = 1 To NUPAR%
      PI(VX%) = CVS(Mid$(PI1$, 4 * VX% - 3))
    Next VX%
    '
    ACAL$ = FILTERGETRECORD$("ANTICAL", 1, MKI$(LEGAX%))
    ACUE = CVI(Mid$(ACAL$, 9, 2))
    '
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
     Text11.SetFocus
   End If
End Sub

Private Sub Text11_Change()
  '
  Text14 = "": Text14.Enabled = False
  Text15 = "": Text15.Enabled = False
  Text2 = "": Text2.Enabled = False
  '
  If ECOARCH$("PADREMP", MKI$(Val(Text1))) = "" Then Exit Sub
  Label4 = TRIM$(ECOARCH$("ECONCEP", MKI$(Val(Text11))))
  If TRIM$(Label4) = "" Then Exit Sub
  '
  COD$ = RTrim(Text11)
  For U& = 1 To List1.ListCount
    TTXX$ = List1.List(U& - 1)
    If Val(Left$(TTXX$, 5)) = Val(COD$) Then
      Text14 = TRIM$(Mid$(TTXX$, 52, 8))
      Text15 = TRIM$(Mid$(TTXX$, 61, 12))
      Text2 = TRIM$(Mid$(TTXX$, 74, 12))
    End If
  Next U&
  '
  CODXX% = Val(COD$)
  TICOX% = 1 + Int(CODXX% / 200)
  jcon% = CODXX% Mod 200
  AJUX% = 0
  If jcon% > 100 Then
    jcon% = jcon% Mod 100
    AJUX% = 1
  End If
  '
  If AJUX% > 0 Then
    Text14.Enabled = True
    Text15.Enabled = True
    Text2.Enabled = True
    Exit Sub
  End If
  '
  ri2# = REMUCALC#(COD$)
  If TICOX% = 1 Then
       Text15 = FORMATNUM$(ri2#, "F10,2")
     ElseIf TICOX% = 2 Then
       Text2 = FORMATNUM$(Abs(ri2#), "F10,2")
  End If
  '
  If TICOX% = 1 Then   ' CONCEPTO REMUNERATIVO
     If FORCON%(jcon%) = 1 Or FORCON%(jcon%) = 4 Then
         Text15.Enabled = True
       ElseIf FORCON%(jcon%) = 2 Or FORCON%(jcon%) = 3 Then
         Text14.Enabled = True
     End If
     Exit Sub
  End If
  '
  If TICOX% = 2 Then
     If FORRET%(jcon%) = 1 Or FORRET%(jcon%) = 3 Then
       Text2.Enabled = True
     End If
     Exit Sub
  End If
  '
  If TICOX% = 3 Then
     Text14.Enabled = True
     Exit Sub
  End If
  '
End Sub

Private Sub Text11_GotFocus()
   Text11.SelStart = 0
   Text11.SelLength = Len(Text11)
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
     If Text14.Enabled = True Then
         Text14.SetFocus
       ElseIf Text15.Enabled = True Then
         Text15.SetFocus
       ElseIf Text2.Enabled = True Then
         Text2.SetFocus
       Else
         Command6.SetFocus
     End If
   ElseIf KEYASII% = 27 Then
     Command1.SetFocus
   End If
End Sub

Private Sub Text14_Change()
   COD$ = RTrim(Text11)
   ICONI% = Val(Text11)
   If ICONI% > 0 Then
     If ICONI% <= NUCON% Then
       QT(ICONI%) = Val(Text14)
       ri2# = REMUCALC#(COD$)
       RM#(ICONI%) = ri2#
       If ri2# > 0 Then
            Text15 = FORMATNUM$(ri2#, "F10,2")
            Text2 = ""
          Else
            Text15 = ""
            Text2 = FORMATNUM$(Abs(ri2#), "F10,2")
       End If
     End If
   End If
End Sub

Private Sub Text14_GotFocus()
   Text14.SelStart = 0
   Text14.SelLength = Len(Text14)
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
     Command6.SetFocus
   End If
End Sub

Private Sub Text14_LOSTFOCUS()
  Text14.TEXT = CSTRING$(MKD$(Val(Text14.TEXT)), 3, 7, 2, 2)
End Sub

Private Sub Text15_GotFocus()
   Text15.SelStart = 0
   Text15.SelLength = Len(Text15)
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
     Command6.SetFocus
   End If
End Sub

'
Private Sub Text15_LOSTFOCUS()
  Text15.TEXT = CSTRING$(MKD$(Val(Text15.TEXT)), 3, 11, 2, 2)
End Sub

Private Sub Text2_GotFocus()
   Text2.SelStart = 0
   Text2.SelLength = Len(Text2)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii% = 13 Then
     Command6.SetFocus
   End If
End Sub

Private Sub Text2_LostFocus()
  Text2.TEXT = CSTRING$(MKD$(Val(Text2.TEXT)), 3, 11, 2, 2)
End Sub

Private Sub Text3_GotFocus()
  '
  On Error Resume Next
    Text1.SetFocus
  On Error GoTo 0
  '
End Sub

Function REMUCALC#(COD$)
      '
      Call CARGAFORMULAS
      '
      RM0# = 0
      For KU& = 1 To List1.ListCount
        TTXX$ = List1.List(KU& - 1)
        CCON% = Val(Left$(TTXX$, 5))
        If CCON% > 0 Then
          If CCON% <= 199 Then
            CCON% = CCON% Mod 100
            If SURETLE%(CCON%) > 0 Then
              RM0# = RM0# + XVALO(Mid$(TTXX$, 61, 12)) - XVALO(Mid$(TTXX$, 74, 12))
            End If
          End If
        End If
      Next KU&
      RM#(0) = RM0#
      '
14100 TICO% = 0: ENCUE% = 0: AJU% = 0: POSEDI% = 1
14110 For I% = 1 To Len(COD$)
14115   B% = Asc(Mid$(COD$, I%)): If B% < 48 Or B% > 57 Then GoTo 14390 ' DESTAJO
14120 Next I%
14125 CODCONX = Val(COD$)
      If CODCONX < 1 Or CODCONX > 599 Then GoTo 14390 ' DESTAJO
14130 '
14145 TICO% = 1 + Int(CODCONX / 200)
      CODCON% = CODCONX Mod 200         '1=REMUN, 2=RETEN, 3=ASFAM
14150 AJU% = Int(CODCON% / 100)
      CODCON% = CODCON% Mod 100        '0=NORMAL, 1=AJUSTE
14155 CDI% = 200 * (TICO% - 1) + 100 * AJU% + CODCON%
14160 On TICO% GoTo 14200, 14230, 14260
      '
14200 If CODCON% < 1 Or CODCON% > NUCON% Then GoTo 14390 ' CONCEPTO NO DEFINIDO
14205 ENCUE% = 1: DENO$ = Left$(CON$(CODCON%), 28)
14210 MCAL% = FORCON%(CODCON%)
      FORMU$ = FORCONCE$(CODCON%)
14215 GoTo 14390
14220 '
14230 If CODCON% < 1 Or CODCON% > NURET% Then GoTo 14390 ' RETENCION NO DEFINIDA
14235 ENCUE% = 1: DENO$ = Left$(RET$(CODCON%), 28)
14240 MCAL% = FORRET%(CODCON%)
      FORMU$ = FORETLE$(CODCON%)
      If RCLRET%(CODCON%) = 2 Then POSEDI% = -1
14245 GoTo 14390
14250 '
14260 If CODCON% < 1 Or CODCON% > NUASI% Then GoTo 14390 ' ASIGNACION NO DEFINIDA
14265 ENCUE% = 1: DENO$ = Left$(ASI$(CODCON%), 28)
14270 MCAL% = 3
14275 GoTo 14390
14280 '
14390 If AJU% > 0 Then DENO$ = "AJUSTE " + Left$(DENO$, 21)
14392 'Return
14395 '
16000 If TOPERET# > 0.5 Then If RM#(0) > TOPERET# Then RM#(0) = TOPERET#
16001 FORMUTRAB$ = FORMU$: RI1% = 0: ri2# = 0
      If MCAL% = 4 Then
        GoSub 16600
        GoTo 16050
      End If
16002 LIMI% = 0: If InStr(FORMUTRAB$, ">") = 0 Then GoTo 16005
16003 LIMI% = InStr(FORMUTRAB$, ">"): FORMU2$ = Mid$(FORMUTRAB$, LIMI% + 1): FORMUTRAB$ = Left$(FORMUTRAB$, LIMI% - 1)
16004 '
16005 PINI% = InStr(FORMUTRAB$, "@"): If PINI% < 1 Then GoTo 16025
16010 GoSub 16060              ' RESOLUCION SUMATORIA
16015 GoTo 16005
16020 '
16025 PINI% = InStr(FORMUTRAB$, "("): If PINI% < 1 Then GoTo 16045
16030 GoSub 16155              ' RESOLUCION PARENTESIS
16035 GoTo 16025
16040 '
16045 CALFIN% = 0: GoSub 16220 ' RESOLUCION FORMULA ALGEBRAICA
16046 If LIMI% = 0 Then GoTo 16050
16047 If LIMI% > 0 Then RIBAS# = ri2#: FORMUTRAB$ = FORMU2$: ri2# = 0: RI1% = 0: LIMI% = -1: GoTo 16005
16048 If RIBAS# > ri2# Then ri2# = RIBAS#
16049 '
16050 If TICO% = 1 Or TICO% = 3 Then
           REMUCALC# = ri2#
         ElseIf TICO% = 2 Then
           REMUCALC# = (-1) * ri2#
      End If
      Exit Function ' Return
16055 '
16060 IZQUI$ = "": If PINI% > 1 Then IZQUI$ = Left$(FORMUTRAB$, PINI% - 1): FORMUTRAB$ = Mid$(FORMUTRAB$, PINI%)
16065 VARI3$ = Mid$(FORMUTRAB$, 2, 2): FORMUTRAB$ = Mid$(FORMUTRAB$, 4)
      If Left$(FORMUTRAB$, 1) = "(" Then FORMUTRAB$ = Mid$(FORMUTRAB$, 2)
16070 LINF% = Val(Left$(FORMUTRAB$, 2)): FORMUTRAB$ = Mid$(FORMUTRAB$, 3)
      If LINF% < 1 Then LINF% = 1
16075 If Left$(FORMUTRAB$, 1) = "." Then FORMUTRAB$ = Mid$(FORMUTRAB$, 2): GoTo 16075
16080 LSUP% = Val(Left$(FORMUTRAB$, 2))
16085 If LSUP% < 10 Then DERE$ = Mid$(FORMUTRAB$, 2) Else DERE$ = Mid$(FORMUTRAB$, 3)
16090 ri2# = 0
16095 For RANGO% = LINF% To LSUP%
16100 VALOR# = 0
16105 If VARI3$ = "SB" Then VALOR# = SB(RANGO%): GoTo 16125
16110 If VARI3$ = "PI" Then VALOR# = PI(RANGO%): GoTo 16125
16115 If VARI3$ = "QT" Then VALOR# = QT(RANGO%): GoTo 16125
16120 If VARI3$ = "RM" Then VALOR# = RM#(RANGO%): GoTo 16125
16125 ri2# = ri2# + (Int(100 * VALOR# + 0.5)) / 100
16130 Next RANGO%
16135 RI1% = RI1% + 1: RINTER#(RI1%) = ri2#: RI3$ = "RI" + Mid$(Str$(RI1%), 2)
16140 FORMUTRAB$ = IZQUI$ + RI3$ + DERE$
16145 Return
16150 '
16155 IZQUI$ = "": If PINI% > 1 Then IZQUI$ = Left$(FORMUTRAB$, PINI% - 1)
16160 FORMUTRAB$ = Mid$(FORMUTRAB$, PINI% + 1)
16165 PINI% = InStr(FORMUTRAB$, "("): If PINI% > 0 Then IZQUI$ = IZQUI$ + "(" + Left$(FORMUTRAB$, PINI% - 1): FORMUTRAB$ = Mid$(FORMUTRAB$, PINI% + 1): GoTo 16165
16170 PFIN% = InStr(FORMUTRAB$, ")"): If PFIN% < 1 Then PFIN% = 1 + Len(FORMUTRAB$)
16175 DERE$ = Mid$(FORMUTRAB$, PFIN% + 1)
16180 FORMUTRAB$ = Left$(FORMUTRAB$, PFIN% - 1)
16185 '
16190 CALFIN% = 1: GoSub 16220       ' RESOLUCION FORMULA SIN PARENTESIS
16195 '
16200 RI1% = RI1% + 1: RINTER#(RI1%) = ri2#: RI3$ = "RI" + Mid$(Str$(RI1%), 2)
16205 FORMUTRAB$ = IZQUI$ + RI3$ + DERE$
16210 Return
16215 '
16220 POPE% = InStr(FORMUTRAB$, "^"): If POPE% > 0 Then OPE% = 5: GoTo 16280
16222 POPE% = InStr(FORMUTRAB$, "x"): If POPE% > 0 Then OPE% = 1: GoTo 16280
16225 POPE% = InStr(FORMUTRAB$, "X"): If POPE% > 0 Then OPE% = 1: GoTo 16280
16230 POPE% = InStr(FORMUTRAB$, "*"): If POPE% > 0 Then OPE% = 1: GoTo 16280
16235 POPE% = InStr(FORMUTRAB$, "/"): If POPE% > 0 Then OPE% = 2: GoTo 16280
16240 '
16245 POPE% = InStr(FORMUTRAB$, "+"): If POPE% > 0 Then OPE% = 3: GoTo 16280
16250 '
16255 POPE% = InStr(FORMUTRAB$, "-"): If POPE% > 0 Then OPE% = 4: GoTo 16280
16260 '
16265 ' NO QUEDAN OPERACIONES ARITMETICAS POR CONSIDERAR
16267 If Abs(ri2#) < 0.005 Then If CALFIN% = 0 Then IZQUIE$ = "": DERIE$ = "": VALO1# = 0: ARGU2$ = FORMUTRAB$: OPE% = 3: CALFIN% = 1: GoTo 16460
16270 Return
16275 '
16280 ARG1$ = "": ARG2$ = ""
16285 For PP% = POPE% - 1 To 1 Step -1
16290 '
16295 CR$ = Mid$(FORMUTRAB$, PP%, 1)
16300 If CR$ = "^" Or CR$ = "x" Or CR$ = "X" Or CR$ = "*" Or CR$ = "/" Or CR$ = "+" Or CR$ = "-" Then GoTo 16320
16305 Next PP%
16310 PP% = 0
16315 '
16320 ARG1$ = Mid$(FORMUTRAB$, PP% + 1, POPE% - PP% - 1): IZQUIE$ = "": If PP% > 0 Then IZQUIE$ = Left$(FORMUTRAB$, PP%)
16325 '
16330 For PP% = POPE% + 1 To Len(FORMUTRAB$)
16335 CR$ = Mid$(FORMUTRAB$, PP%, 1)
16340 If CR$ = "^" Or CR$ = "x" Or CR$ = "X" Or CR$ = "*" Or CR$ = "/" Or CR$ = "+" Or CR$ = "-" Then GoTo 16360
16345 Next PP%
16350 PP% = 1 + Len(FORMUTRAB$)
16355 '
16360 ARG2$ = Mid$(FORMUTRAB$, POPE% + 1, PP% - POPE% - 1): DERIE$ = Mid$(FORMUTRAB$, PP%)
16365 '
16370 ARGU1$ = "": For PP% = 1 To Len(ARG1$): CR1% = Asc(Mid$(ARG1$, PP%)): If (CR1% > 47 Or CR1% = 46) Then ARGU1$ = ARGU1$ + Chr$(CR1%)
16375 Next PP%
16380 '
16385 ARGU2$ = "": For PP% = 1 To Len(ARG2$): CR1% = Asc(Mid$(ARG2$, PP%)): If (CR1% > 47 Or CR1% = 46) Then ARGU2$ = ARGU2$ + Chr$(CR1%)
16390 Next PP%
16395 '
16400 VALO1# = 0
16405 If Left$(ARGU1$, 2) = "AE" Then VALO1# = ACUE: GoTo 16460
16406 If Left$(ARGU1$, 2) = "QM" Then VALO1# = QM#: GoTo 16460
16407 If Left$(ARGU1$, 2) = "RG" Then VALO1# = RETG#: GoTo 16460
16408 If Left$(ARGU1$, 2) = "AC" Then VALO1# = AC: GoTo 16460
16410 If Left$(ARGU1$, 2) = "DJ" Then VALO1# = DJ0#: GoTo 16460
16415 INDI = Val(Mid$(ARGU1$, 3)): INDII% = 0: If INDI >= 0 Then If INDI < 100 Then INDII% = Int(INDI)
16420 If Left$(ARGU1$, 2) = "RI" Then VALO1# = RINTER#(INDII%): GoTo 16460
16425 If Left$(ARGU1$, 2) = "SB" Then VALO1# = SB(INDII%): GoTo 16460
16430 If Left$(ARGU1$, 2) = "QT" Then VALO1# = QT(INDII%): GoTo 16460
16435 If Left$(ARGU1$, 2) = "RM" Then VALO1# = RM#(INDII%): GoTo 16460
16440 If Left$(ARGU1$, 2) = "PI" Then VALO1# = PI(INDII%): GoTo 16460
      If Left$(ARGU1$, 2) = "AM" Then GoSub 16800: VALO1# = VALOACU#: GoTo 16460 ' ACUMULADO MENSUAL DEL CODIGO
16445 '
16450 VALO1# = Val(ARGU1$)
16455 '
16460 VALO2# = 0
16465 If Left$(ARGU2$, 2) = "AE" Then VALO2# = ACUE: GoTo 16520
16466 If Left$(ARGU2$, 2) = "QM" Then VALO2# = QM#: GoTo 16520
16467 If Left$(ARGU2$, 2) = "RG" Then VALO2# = RETG#: GoTo 16520
16468 If Left$(ARGU2$, 2) = "AC" Then VALO2# = AC: GoTo 16520
16470 If Left$(ARGU2$, 2) = "DJ" Then VALO2# = DJ0#: GoTo 16520
16475 INDI = Val(Mid$(ARGU2$, 3)): INDII% = 0: If INDI >= 0 Then If INDI < 100 Then INDII% = Int(INDI)
16480 If Left$(ARGU2$, 2) = "RI" Then VALO2# = RINTER#(INDII%): GoTo 16520
16485 If Left$(ARGU2$, 2) = "SB" Then VALO2# = SB(INDII%): GoTo 16520
16490 If Left$(ARGU2$, 2) = "QT" Then VALO2# = QT(INDII%): GoTo 16520
16495 If Left$(ARGU2$, 2) = "RM" Then VALO2# = RM#(INDII%): GoTo 16520
16500 If Left$(ARGU2$, 2) = "PI" Then VALO2# = PI(INDII%): GoTo 16520
      If Left$(ARGU2$, 2) = "AM" Then GoSub 16800: VALO2# = VALOACU#: GoTo 16520 ' ACUMULADO MENSUAL DEL CODIGO
16505 '
16510 VALO2# = Val(ARGU2$)
16515 '
16520 ri2# = 0
16525 If OPE% = 1 Then ri2# = VALO1# * VALO2#: GoTo 16550
16530 If OPE% = 2 Then If Abs(VALO2#) >= 0.005 Then ri2# = VALO1# / VALO2#: GoTo 16550
16535 If OPE% = 3 Then ri2# = VALO1# + VALO2#: GoTo 16550
16540 If OPE% = 4 Then ri2# = VALO1# - VALO2#: GoTo 16550
16542 If OPE% = 5 Then ri2# = VALO1# ^ VALO2#: GoTo 16550
16545 '
16550 RI1% = RI1% + 1: RINTER#(RI1%) = ri2#: RI3$ = "RI" + Mid$(Str$(RI1%), 2)
16555 FORMUTRAB$ = IZQUIE$ + RI3$ + DERIE$
16560 GoTo 16220
16565 '
16600 For F% = 0 To 10: VALESC(F%) = 0: Next F%
16602 F% = 0
16605 VALESC(F%) = Val(FORMUTRAB$)
16607 F1% = InStr(FORMUTRAB$, "-"): If F1% < 1 Then GoTo 16615
16610 FORMUTRAB$ = Mid$(FORMUTRAB$, F1% + 1): If F% < 10 Then F% = F% + 1: GoTo 16605
16612 '
16615 FMAX% = F%
16617 If QT0 < 0 Or QT0 > FMAX% Then ri2# = 0: GoTo 16625
16620 ri2# = CDbl(VALESC(QT0))
16625 Return
16627 '

16800 ' If N10% < 1 Then GoSub 5960
'      UREDTLQ% = LOF(N10%) / 24
'      VALOACU# = 0
'      For IJKI% = 1 To Len(PULIQ$(EMPL%)) Step 2
' REDTLQ% = CVI(Mid$(PULIQ$(EMPL%), IJKI%, 2))
' If REDTLQ% < 1 Or REDTLQ% > UREDTLQ% Then GoTo 16819
' GET #N10%,REDTLQ%
' If CVI(J0$) = 0 Then GoTo 16819
' If CVI(J0$) = MARCA% Then GoTo 16819 ' ES LIQUIDACION ACTIVA
' LGJ8% = CVI(J1$): If LGJ8% <> LEGA% Then GoTo 16819
' CODI8% = CVI(J2$): If (CODI8% < 1 Or CODI8% > 199) Then GoTo 16819
' If CODI8% Mod 100 <> INDII% Then GoTo 16819
' IMPORT# = CVDMBF(J4$)
' VALOACU# = VALOACU# + IMPORT#
'16819 Next IJKI%
      '
16829 Return
      '

End Function

Sub RECALCONCE()
      '
      Call CARGAFORMULAS       ' CARGAR CONCEPTOS y RETENCIONES
      '
      For jcon% = 1 To NUCON%
        If ESTCON%(jcon%) < 1 Then GoTo 8229 ' DESCARTO SUSPENDIDOS
        If QLCON%(jcon%) <> 0 And QLCON%(jcon%) <> TILIQ% Then GoTo 8229 ' NO CORRESPONDE QUINCENA
        If RCLCON%(jcon%) <> 1 Then GoTo 8229 ' DESCARTO RECALCULO OFF
        '
        If CARCON%(jcon%) = 3 Then GoTo 8220 ' ES MANDATORIO
        If CARCON%(jcon%) = GRULI% Then GoTo 8220 ' CORRESPONDE QUINCENA
        '
        For X% = 1 To 96 Step 2
          X1% = CVI(Mid$(CONOPCI$, X%))
          If X1% = 0 Or X1% > jcon% Then GoTo 8229
          If X1% = jcon% Then GoTo 8220           ' PREASIGNADO
        Next X%
        GoTo 8229
        '
8220    RM#(jcon%) = REMUCALC#(TRIM$(Str$(jcon%)))
        '
        For KU& = 1 To List1.ListCount
          If Val(Left$(List1.List(KU& - 1), 5)) = jcon% Then
            List1.RemoveItem KU& - 1
          End If
        Next KU&
        '
        If Abs(QT(jcon%)) >= 0.005 Or Abs(RM#(jcon%)) >= 0.005 Then
          TTXX$ = Space$(96)
          Call REPLA(TTXX$, FORMATNUM$(jcon%, "I5"), 1, 5)
          Call REPLA(TTXX$, TRIM$(ECOARCH$("ECONCEP", MKI$(jcon%))), 9, 36)
          Call REPLA(TTXX$, FORMATNUM$(QT(jcon%), "F8.2"), 52, 8)
          If RM#(jcon%) > 0 Then
               Call REPLA(TTXX$, FORMATNUM$(RM#(jcon%), "F12.2"), 61, 12)
             Else
               Call REPLA(TTXX$, FORMATNUM$(Abs(RM#(jcon%)), "F12.2"), 74, 12)
          End If
          '
          For KU& = List1.ListCount To 1 Step -1
            If Val(Left$(List1.List(KU& - 1), 5)) < jcon% Then
              List1.AddItem TTXX$, KU&
              GoTo 8229
            End If
          Next KU&
          List1.AddItem TTXX$, 0
        End If
        '
8229  Next jcon%
      '
      RM0# = 0
      For KU& = 1 To List1.ListCount
        TTXX$ = List1.List(KU& - 1)
        CCON% = Val(Left$(TTXX$, 5))
        If CCON% > 0 Then
          If CCON% <= 199 Then
            CCON% = CCON% Mod 100
            If SURETLE%(CCON%) > 0 Then
              RM0# = RM0# + XVALO(Mid$(TTXX$, 61, 12)) - XVALO(Mid$(TTXX$, 74, 12))
            End If
          End If
        End If
      Next KU&
      RM#(0) = RM0#
      '
8240 For jcon% = 1 To NURET%
        If ESTRET%(jcon%) < 1 Then GoTo 8270 ' DESCARTO SUSPENDIDOS
        If RCLRET%(jcon%) <> 1 Then GoTo 8270 ' DESCARTO SI NO RECALCULO ON
        If QLRET%(jcon%) <> 0 Then If QLRET%(jcon%) <> TILIQ% Then GoTo 8270
        '
        If CARRET%(jcon%) = 3 Then GoTo 8260
        If CARRET%(jcon%) = GRULI% Then GoTo 8260 ' MANDATORIO
        '
        For X% = 1 To 96 Step 2
          X1% = CVI(Mid$(CONOPCI$, X%))
          If X1% = 0 Or X1% > 200 + jcon% Then GoTo 8270
          If X1% = 200 + jcon% Then GoTo 8260         ' PREASIGNADO
        Next X%
        GoTo 8270
        '
8260    RT#(jcon%) = REMUCALC#(TRIM$(Str$(200 + jcon%)))
        '
        For KU& = 1 To List1.ListCount
          If Val(Left$(List1.List(KU& - 1), 5)) = 200 + jcon% Then
            List1.RemoveItem KU& - 1
          End If
        Next KU&
        '
        If Abs(RT#(jcon%)) >= 0.005 Then
           TTXX$ = Space$(96)
           Call REPLA(TTXX$, FORMATNUM$(200 + jcon%, "I5"), 1, 5)
           Call REPLA(TTXX$, TRIM$(ECOARCH$("ECONCEP", MKI$(200 + jcon%))), 9, 36)
           If RT#(jcon%) > 0 Then
                Call REPLA(TTXX$, FORMATNUM$(RT#(jcon%), "F12.2"), 61, 12)
              Else
                Call REPLA(TTXX$, FORMATNUM$(Abs(RT#(jcon%)), "F12.2"), 74, 12)
           End If
           '
           For KU& = List1.ListCount To 1 Step -1
             If Val(Left$(List1.List(KU& - 1), 5)) < 200 + jcon% Then
               List1.AddItem TTXX$, KU&
               GoTo 8270
             End If
           Next KU&
           List1.AddItem TTXX$, 0
        End If
        '
8270 Next jcon%
     '
8280 ' solo si se liquidan asignaciones familiares
     If ASIFLI% > 0 Then
        For jcon% = 1 To NUASI%
          If ESTASI%(jcon%) > 0 Then
            CANAS% = Asc(Mid$(AFAM_LEG$, jcon%, 1))
            If CANAS% > 0 Then
              '
8281          For KU& = 1 To List1.ListCount
                If Val(Left$(List1.List(KU& - 1), 5)) = 400 + jcon% Then
                  GoTo 8290      ' PARA DEJAR LO QUE SE INFORMO A MANO
                  'List1.RemoveItem KU& - 1
                  'GoTo 8281
                End If
              Next KU&
              '
              VASI# = CANAS% * VALASI(jcon%)
              If VASI# > 0.005 Then
                TTXX$ = Space$(96)
                Call REPLA(TTXX$, FORMATNUM$(400 + jcon%, "I5"), 1, 5)
                Call REPLA(TTXX$, TRIM$(ECOARCH$("ECONCEP", MKI$(400 + jcon%))), 9, 36)
                Call REPLA(TTXX$, FORMATNUM$(CANAS%, "F8.2"), 52, 5)
                Call REPLA(TTXX$, FORMATNUM$(VASI#, "F12.2"), 61, 12)
                '
                For KU& = List1.ListCount To 1 Step -1
                  If Val(Left$(List1.List(KU& - 1), 5)) < 400 + jcon% Then
                    List1.AddItem TTXX$, KU&
                    GoTo 8290
                  End If
                Next KU&
                List1.AddItem TTXX$, 0
              End If
            End If
          End If
          '
8290    Next jcon%
     End If
     '
     Call SUMAREMU
     '
End Sub

Sub SUMAREMU()
     '
     SUREMU# = 0: SURETE# = 0
     For KU& = 1 To List1.ListCount
        TTXX$ = List1.List(KU& - 1)
        CCON% = Val(Left$(TTXX$, 5))
        REMUU# = XVALO(Mid$(TTXX$, 61, 12)) - XVALO(Mid$(TTXX$, 74, 12))
        If CCON% > 0 And CCON% <= 199 Then
            SUREMU# = SUREMU# + REMUU#
          ElseIf CCON% >= 200 And CCON% <= 399 Then
            SURETE# = SURETE# - REMUU#
          ElseIf CCON% >= 400 And CCON% <= 599 Then
            SUREMU# = SUREMU# + REMUU#
        End If
     Next KU&
     Label22 = TRIM$(FORMATNUM$(SUREMU#, "F12.2"))
     Label1 = TRIM$(FORMATNUM$(SURETE#, "F12.2"))
     '
End Sub

Private Sub Text4_Change()

End Sub

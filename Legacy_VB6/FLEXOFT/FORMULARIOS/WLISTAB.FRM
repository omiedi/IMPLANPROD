VERSION 5.00
Begin VB.Form WLISTAB 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0EEFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generador de Reportes"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   FillStyle       =   0  'Solid
   Icon            =   "WLISTAB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   11910
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   330
      Left            =   6825
      TabIndex        =   50
      Top             =   210
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Impresora:"
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
      Index           =   0
      Left            =   525
      TabIndex        =   40
      Top             =   2600
      Width           =   1250
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Hoja de Cálculo (Planilla Excel)"
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
      Height          =   540
      Index           =   2
      Left            =   6825
      TabIndex        =   39
      Top             =   3700
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Archivo de Texto"
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
      Index           =   1
      Left            =   6825
      TabIndex        =   38
      Top             =   1890
      Width           =   1905
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Exportar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4110
      Left            =   6510
      TabIndex        =   30
      Top             =   1420
      Width           =   5160
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
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
         Left            =   2520
         TabIndex        =   46
         Text            =   " "
         Top             =   3150
         Width           =   2430
      End
      Begin VB.DriveListBox Drive1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   37
         Top             =   315
         Width           =   1275
      End
      Begin VB.DirListBox Dir1 
         Enabled         =   0   'False
         Height          =   2115
         Left            =   2520
         TabIndex        =   36
         Top             =   735
         Width           =   2430
      End
      Begin VB.CommandButton Command4 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4200
         Picture         =   "WLISTAB.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   300
         Width           =   540
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0EEFF&
         Caption         =   "Opciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   160
         TabIndex        =   31
         Top             =   840
         Width           =   2010
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0EEFF&
            Caption         =   "Mostrar Archivo"
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
            TabIndex        =   34
            Top             =   840
            Width           =   1695
         End
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0EEFF&
            Caption         =   "Agregar al Final"
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
            Index           =   0
            Left            =   210
            TabIndex        =   33
            Top             =   315
            Width           =   1695
         End
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0EEFF&
            Caption         =   "Sobre-Escribir"
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
            Index           =   1
            Left            =   210
            TabIndex        =   32
            Top             =   525
            Width           =   1695
         End
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000009&
         X1              =   2325
         X2              =   2325
         Y1              =   105
         Y2              =   2940
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000009&
         X1              =   0
         X2              =   2310
         Y1              =   2115
         Y2              =   2115
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000009&
         X1              =   0
         X2              =   5145
         Y1              =   2955
         Y2              =   2955
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Carpeta:"
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
         Left            =   -210
         TabIndex        =   49
         Top             =   3320
         Width           =   1065
      End
      Begin VB.Label Label11 
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
         Left            =   420
         TabIndex        =   48
         Top             =   3570
         Width           =   4530
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
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
         Left            =   1365
         TabIndex        =   47
         Top             =   3190
         Width           =   1065
      End
      Begin VB.Line Line5 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   5145
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000C&
         X1              =   2310
         X2              =   2310
         Y1              =   105
         Y2              =   2940
      End
      Begin VB.Line Line7 
         BorderColor     =   &H8000000C&
         X1              =   0
         X2              =   2310
         Y1              =   2100
         Y2              =   2100
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   6510
      ScaleHeight     =   75
      ScaleWidth      =   1215
      TabIndex        =   28
      Top             =   1050
      Width           =   1275
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
   Begin VB.CommandButton Command3 
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
      Height          =   855
      Left            =   8925
      Picture         =   "WLISTAB.frx":0784
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Guarda Configuración del Listado"
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Control de Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1720
      Left            =   210
      TabIndex        =   14
      Top             =   210
      Width           =   6105
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         X1              =   3120
         X2              =   3120
         Y1              =   120
         Y2              =   1760
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   3105
         X2              =   3105
         Y1              =   120
         Y2              =   1760
      End
      Begin VB.Label Label9 
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
         Height          =   720
         Left            =   210
         TabIndex        =   45
         Top             =   840
         Width           =   2745
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1365
         TabIndex        =   44
         Top             =   450
         Width           =   1590
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Controlador:"
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
         Left            =   210
         TabIndex        =   43
         Top             =   525
         Width           =   1065
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
         Height          =   720
         Left            =   3255
         TabIndex        =   42
         Top             =   840
         Width           =   2640
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4305
         TabIndex        =   41
         Top             =   450
         Width           =   1590
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
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
         Left            =   3150
         TabIndex        =   15
         Top             =   525
         Width           =   1065
      End
   End
   Begin VB.CommandButton Command2 
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
      Height          =   855
      Left            =   7980
      Picture         =   "WLISTAB.frx":0A8E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Configura Impresora"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9870
      Picture         =   "WLISTAB.frx":0ED0
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Emite Listado"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton BOTEXIT 
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
      Height          =   855
      Left            =   10815
      Picture         =   "WLISTAB.frx":11DA
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Abandona el Listador"
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   210
      TabIndex        =   0
      Top             =   2050
      Width           =   6105
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1575
         TabIndex        =   4
         Text            =   " "
         Top             =   420
         Width           =   2955
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1575
         TabIndex        =   23
         Text            =   " "
         Top             =   840
         Width           =   2955
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0EEFF&
         Caption         =   "Medidas en Milimetros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   1575
         TabIndex        =   16
         Top             =   2100
         Width           =   4320
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   3
            Left            =   3255
            TabIndex        =   25
            Text            =   " "
            Top             =   735
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   2
            Left            =   3255
            TabIndex        =   24
            Text            =   " "
            Top             =   420
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   21
            Text            =   " "
            Top             =   735
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   19
            Text            =   " "
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Margen Sup:"
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
            Index           =   3
            Left            =   1995
            TabIndex        =   27
            Top             =   420
            Width           =   1170
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Izquierdo:"
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
            Index           =   2
            Left            =   2310
            TabIndex        =   26
            Top             =   735
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Largo:"
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
            Index           =   1
            Left            =   -105
            TabIndex        =   20
            Top             =   735
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ancho:"
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
            Index           =   0
            Left            =   -105
            TabIndex        =   18
            Top             =   420
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0EEFF&
         Caption         =   "Tamaño"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   4830
         TabIndex        =   12
         Top             =   340
         Width           =   1065
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   210
            TabIndex        =   13
            Text            =   " "
            Top             =   350
            Width           =   645
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0EEFF&
         Caption         =   "Copias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   4830
         TabIndex        =   10
         Top             =   1260
         Width           =   1065
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   315
            TabIndex        =   11
            Text            =   " "
            Top             =   315
            Width           =   435
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0EEFF&
         Caption         =   "Orientación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   1575
         TabIndex        =   7
         Top             =   1260
         Width           =   2955
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0EEFF&
            Caption         =   "Paisaje"
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
            Index           =   1
            Left            =   1785
            TabIndex        =   9
            Top             =   315
            Width           =   1065
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0EEFF&
            Caption         =   "Retrato"
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
            Index           =   0
            Left            =   210
            TabIndex        =   8
            Top             =   315
            Width           =   1065
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0EEFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   435
         Left            =   630
         TabIndex        =   6
         Top             =   4305
         Width           =   3690
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Papel:"
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
         Left            =   630
         TabIndex        =   17
         Top             =   2205
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fuente:"
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
         Left            =   630
         TabIndex        =   5
         Top             =   900
         Width           =   855
      End
   End
End
Attribute VB_Name = "WLISTAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IMPREDETER As String
'

Private Sub BOTEXIT_Click()
    Call CIERRARCH("*.*")
    If COMMA$ <> "" Then
       Call BLOQEXE("*")
    End If
    WLISTAB.Hide
End Sub

Private Sub Combo1_LostFocus()
    '
    Dim PRX As Printer
    For I% = 1 To Printers.Count
        If Printers(I% - 1).DeviceName = Combo1.TEXT Then
            ' La define como predeterminada del sistema.
            Set Printer = Printers(I% - 1)
            GoTo 10
        End If
    Next
    '
10  Combo1.TEXT = Printer.DeviceName
    '
End Sub

Private Sub Command1_Click()
    '
    If Option1(0) = False Then
       ASALIDA$ = TRIM$(Text2.TEXT)
       If ASALIDA$ = "" Then
          Call MENSERR(24, "Falta Archivo de Exportación")
          Exit Sub
       End If
       PPXY% = InStr(ASALIDA$, ".")
       If PPXY% > 0 Then
          Text2 = Left$(ASALIDA$, PPXY%)
          Call MENSERR(24, "Suprima Extensión de Archivo de Exportación")
          Exit Sub
       End If
       If Len(Label11) > 40 Then
          Call MENSERR(24, "Ruta de Salida Demasiado Larga.")
          Exit Sub
       End If
       ARCHISAL$ = TRIM$(Label11)
       If Len(ARCHISAL$) > 0 Then
         If Right$(ARCHISAL$, 1) <> "\" Then ARCHISAL$ = ARCHISAL$ + "\"
       End If
       ARCHISAL$ = ARCHISAL$ + ASALIDA$
       If Option1(1).Value = True Then
            If InStr(TRIM$(ARCHISAL$), " ") > 0 Then
               Call MENSERR(24, "Archivo o Carpeta de Salida no Válidos.\  \No se Permiten Espacios.\Cambie Archivo o Carpeta de Salida.")
               Exit Sub
            End If
            ARCHISAL$ = ARCHISAL$ + ".TXT"
            Call GRACONTROL("EXPORTEX", IDENTER$, Label11)
          Else
            ARCHISAL$ = ARCHISAL$ + ".XLS"
            Call GRACONTROL("EXPOREXC", IDENTER$, Label11)
       End If
       GoTo 20
    End If
    '
    If TRIM$(IMPREDETER$) <> TRIM$(Combo1.TEXT) Then
        'MENSA$ = "Imposible Aplicar Impresora Seleccionada:\"
        'MENSA$ = MENSA$ + "'" + TRIM$(Combo1.Text) + "'\   \"
        'MENSA$ = MENSA$ + "FLEXOFT ha Restaurado la Impresora Predeterminada:\"
        'MENSA$ = MENSA$ + "'" + TRIM$(IMPREDETER$) + "'\  \"
        'MENSA$ = MENSA$ + "Para Utilizar la Elegida, Cambie la Impresora\"
        'MENSA$ = MENSA$ + "Predeterminada para Windows."
        'Combo1.Text = IMPREDETER$
        'Call COMUNI(MENSA$)
        'Combo1.SetFocus
        'Exit Sub
        Printer.TrackDefault = False
        NPPXX$ = TRIM$(Combo1.TEXT)
        For Each PRX9 In Printers
          If PRX9.DeviceName = NPPXX$ Then
            On Error Resume Next
            Set Printer = PRX9
            On Error GoTo 0
            GoTo 10
          End If
        Next PRX9
        '
10      If Printer.DeviceName <> Combo1.TEXT Then
          MENSA$ = "Imposible Aplicar Impresora Seleccionada:\"
          MENSA$ = MENSA$ + "'" + TRIM$(Combo1.TEXT) + "'\   \"
          MENSA$ = MENSA$ + "FLEXOFT ha Restaurado la Impresora Predeterminada:\"
          MENSA$ = MENSA$ + "'" + TRIM$(IMPREDETER$) + "'\  \"
          MENSA$ = MENSA$ + "Para Utilizar la Elegida, Cambie la Impresora\"
          MENSA$ = MENSA$ + "Predeterminada para Windows."
          Combo1.TEXT = IMPREDETER$
          Call COMUNI(MENSA$)
          Combo1.SetFocus
          Exit Sub
        End If
    End If
    '
    Call GRADATIMPRE
    Call CARDATIMPRE
    '
20  SOBRESCRI% = 0
    If Option3(1) = True Then
        SOBRESCRI% = 1
    End If
    '
    On Error Resume Next
    PORIAN = Printer.Orientation
    If Option2(1).Value = True Then
          Printer.Orientation = 2
        Else
          Printer.Orientation = 1
    End If
    '
    Call VERDATLIS
    INTERLIX = 0
    If REHOJX% > 96 Then INTERLIX = Printer.ScaleHeight / REHOJX%
    '
    If DATLISTOK% <= 0 Then
         On Error Resume Next
         If TRIM$(ARCHISAL$) = "" Then Printer.Orientation = PORIAN
         On Error GoTo 0
         Exit Sub
      ElseIf DATLISTOK% = 1 Then
         If ARCHISAL$ = "" Then
           On Error GoTo 199
         End If
         If SALEXCEL% = 1 Then
           Call ARMAENCA
           Call LISEXCEL
           Call CIERRARCH("*.*")
           WLISTAB.Hide
           Exit Sub
         End If
         'Call SETPRINTER
         'Call CARHEADERS
         'Call ARMAENCA
         If CACOPIAS% < 1 Then CACOPIAS% = 1
         CACOP% = CACOPIAS%
         If ARCHISAL$ <> "" Then CACOP% = 1
         For COP% = 1 To CACOP%
            'SE MUEVE PORQUE SINO DE LA SEGUNDA COPIA EN ADELANTE NO RESPETA LOS SETEOS DE IMPRESION
            Call SETPRINTER
            Call CARHEADERS
            Call ARMAENCA
            '
            Call INITOTA
            If TNC% < 1 Then
                Call ENCAHOJA
            End If
            Call LISTABU
            '
            Printer.Orientation = ORIENTA% 'SE VUELVE A ESTABLECER LA ORIENTACION
            PRINTERSET% = 0                'SE PONE EN CERO PARA QUE SETPRINTER VUELVA A SETEAR LA IMPRESION
         Next COP%
         On Error GoTo 0
    End If
    '
    Call CIERRARCH("*.*")
    PRINTERSET% = 0
    If ARCHISAL$ = "" Then Printer.Orientation = PORIAN
    If COMMA$ <> "" Then
       Call BLOQEXE("*")
    End If
    WLISTAB.Hide
    '
    If ARCHISAL$ <> "" Then
        If WLISTAB.Check1.Value = 1 Then
            LCB$ = LUBAS$
            If EXISTE%(LUCOM$ + "WORDPAD.EXE") = 1 Then
                 Close
                 Shell LUCOM$ + "WORDPAD.EXE " + ARCHISAL$, vbMaximizedFocus
               Else
                 MENSA$ = "Imposible Mostrar Archivo Generado."
                 MENSA$ = MENSA$ + "\  \Falta Utilitario 'WORDPAD.EXE'"
                 MENSA$ = MENSA$ + "\en Directorio 'C:/FLEXOFT'."
                 MENSA$ = MENSA$ + "\   \Instale este Programa Utilitario\y Re-Genere el Listado."
                 Call COMUNI(MENSA$)
            End If
        End If
    End If
    Exit Sub
    '
199 ERNU$ = TRIM$(Str$(Err.Number))
    MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") en la Impresión de\"
    MENSA$ = MENSA$ + "este Listado.\  \"
    '
    MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
    MENSA$ = MENSA$ + "   1.- Impresora Apagada o fuera de Línea.\"
    MENSA$ = MENSA$ + "   2.- Falta Papel o Papel Atorado.\"
    MENSA$ = MENSA$ + "   3.- Problemas en Cable de Señal de Impresora.\"
    MENSA$ = MENSA$ + "   4.- Problemas de Red - No hay Acceso a Impre-\"
    MENSA$ = MENSA$ + "        sora Remota.\"
    MENSA$ = MENSA$ + "   5.- Impresora de Matriz de Puntos Configurada\"
    MENSA$ = MENSA$ + "        para 'Calidad de Imprenta'. Corrija la Con-\"
    MENSA$ = MENSA$ + "        figuración para 'Borrador'.\"
    MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
    '
    c$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(c$, 53, "Error de Impresión de Listado")
    If RESPU% = 4 Then
        Resume
    End If
    Call CIERRARCH("*.*")
    Call FINAL("")
    Close: End
    '
End Sub
'

Private Sub Command2_Click()
    Call SETUPRINTER
End Sub

Private Sub Command3_Click()
    Call GRADATIMPRE
    Call CARDATIMPRE
End Sub


Private Sub Command5_Click()
Dir1.Path = "C:\"
Dir1.Path = "F:\CLIENTES\VDR\DATOS"
End Sub

Private Sub Dir1_Change()
   On Error Resume Next
   RUTASAL$ = Dir1.Path
   Label11.Caption = RUTASAL$
   On Error GoTo 0
End Sub

Private Sub Drive1_Change()
   On Error Resume Next
   Dir1.Path = Drive1.Drive
   On Error GoTo 0
End Sub

Private Sub Form_Load()
    '
    Refresh
    DoEvents
    Screen.MousePointer = 11
    On Error Resume Next
    PORIAN = Printer.Orientation
    'If Err Then
    '  Call MENSERR(24, "Imposible Imprimir.\Error de Comunicacion con la Impresora.")
    '  Resume
    '  On Error GoTo 0
    '  Unload Me
    '  Exit Sub
    'End If
    '
    Printer.TrackDefault = False
    '
    DoEvents
    Call CARDATLIS
    DoEvents
    Call CARDATIMPRE
    DoEvents
    '
    ARCHISAL$ = ""
    Text2.TEXT = ARCHISAL$
    Option1(0).Value = True        ' IMPRESORA
    '
    SOBRESCRI% = 0
    FORMALI$ = FORMA$
    '
    Combo1.Clear
    Dim PRX As Printer
    For Each PRX In Printers
        Combo1.AddItem PRX.DeviceName
        ' La define como predeterminada del sistema.
        'Set Printer = prx
    Next
    Combo1.TEXT = Printer.DeviceName
    IMPREDETER$ = Printer.DeviceName
    '
    If Left$(UCase$(ATRIBU$), 4) = "LPT1" Then
        WLISTAB.Refresh
        Call SETPRINTER
        Call CARHEADERS
        Call ARMAENCA
        Call INITOTA
        If TNC% < 1 Then
            Call ENCAHOJA
        End If
        Call LISTABU
        Call CIERRARCH("*.*")
        PRINTERSET% = 0
        If COMMA$ <> "" Then
           Call BLOQEXE("*")
        End If
        WLISTAB.Hide
    End If
    '
    Printer.Orientation = PORIAN
    If EXISTE%("FLEXCEL.EXE") > 0 Then
      Option1(2).Enabled = True
    End If
    DoEvents
    Screen.MousePointer = 1
    '
End Sub
'

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        ARCHISAL$ = ""
        Text2.TEXT = ""
        Text2.Enabled = False
        Label11 = ""
        Option3(0).Value = False
        Option3(0).Enabled = False
        Option3(1).Value = False
        Option3(1).Enabled = False
        Check1.Enabled = False
        Drive1.Enabled = False
        Dir1.Enabled = False
        Command4.Enabled = False
        SALEXCEL% = 0
      ElseIf Index = 1 Then
        RUTASAL$ = TRIM$(CONTROL$("EXPORTEX", IDENTER$))
        If RUTASAL$ = "" Then RUTASAL$ = Left$(LUDAT$, Len(LUDAT$) - 1)
        On Error Resume Next
        Drive1.Drive = Left$(RUTASAL$, 2)
        Dir1.Path = RUTASAL$
        On Error GoTo 0
        Label11.Caption = RUTASAL$
        ASALIDA$ = TRIM$(ARCHITRA$)
        Text2.TEXT = ASALIDA$
           On Error Resume Next
        Option3(0).Value = True
        Option3(0).Enabled = True
        Option3(1).Enabled = True
        Check1.Enabled = True
        Drive1.Enabled = True
        Dir1.Enabled = True
        Command4.Enabled = True
        SALEXCEL% = 0
        Text2.Enabled = True
        Text2.SetFocus
           On Error GoTo 0
      ElseIf Index = 2 Then
        RUTASAL$ = TRIM$(CONTROL$("EXPOREXC", IDENTER$))
        If RUTASAL$ = "" Then RUTASAL$ = Left$(LUDAT$, Len(LUDAT$) - 1)
        On Error Resume Next
        Drive1.Drive = Left$(RUTASAL$, 2)
        Dir1.Path = RUTASAL$
        On Error GoTo 0
        Label11.Caption = RUTASAL$
        ASALIDA$ = TRIM$(ARCHITRA$)
        Text2.TEXT = ASALIDA$
        Option3(0).Value = False
        Option3(0).Enabled = False
        Option3(1).Value = False
        Option3(1).Enabled = False
        Check1.Enabled = False
        Drive1.Enabled = True
        Dir1.Enabled = True
        Command4.Enabled = True
        SALEXCEL% = 1
        Text2.Enabled = True
        Text2.SetFocus
    End If
End Sub

Private Sub Text3_GotFocus()
    If Option1(0) = True Then
        Option1(0).SetFocus
      Else
        Text2.SetFocus
    End If
End Sub
'

Private Sub Text1_GotFocus()
    If Option1(0) = True Then
        Option1(0).SetFocus
      Else
        Text2.SetFocus
    End If
End Sub


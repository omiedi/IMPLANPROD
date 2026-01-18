VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form WLITAB07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00DAE4EB&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "GENERADOR DE REPORTES"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9570
   FillStyle       =   0  'Solid
   Icon            =   "WLITAB07.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   59
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
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
      Height          =   525
      Left            =   8000
      Picture         =   "WLITAB07.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Configura Impresora"
      Top             =   5190
      Width           =   700
   End
   Begin VB.CommandButton BOTEXIT 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   8745
      Picture         =   "WLITAB07.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Abandona el Listador"
      Top             =   5190
      Width           =   700
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   7250
      Picture         =   "WLITAB07.frx":0896
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Emite Listado"
      Top             =   5190
      Width           =   700
   End
   Begin VB.Timer Timer2 
      Interval        =   350
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00DAE4EB&
      Caption         =   "PROGRESO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   460
      Left            =   7230
      TabIndex        =   48
      Top             =   4725
      Width           =   2220
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   1935
         TabIndex        =   65
         Top             =   190
         Width           =   1965
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   66
            Top             =   0
            Width           =   12000
         End
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00DAE4EB&
      Caption         =   "PRESENTACIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   4710
      TabIndex        =   46
      Top             =   4725
      Width           =   2430
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   540
         Index           =   3
         Left            =   240
         Picture         =   "WLITAB07.frx":0BA0
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Vista Previa en Pantalla"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1050
         TabIndex        =   47
         Top             =   400
         Width           =   1275
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00DAE4EB&
      Caption         =   "EXPORTAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3120
      Left            =   4710
      TabIndex        =   36
      Top             =   1560
      Width           =   4740
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1880
         TabIndex        =   62
         ToolTipText     =   "Cantidad Máxima de Renglones del Detalle por Hoja"
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Enabled         =   0   'False
         Height          =   405
         Left            =   3990
         Picture         =   "WLITAB07.frx":1DE2
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Actualizar Lista"
         Top             =   1455
         Width           =   540
      End
      Begin VB.CommandButton Command5 
         Caption         =   "."
         Enabled         =   0   'False
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
         Left            =   3675
         TabIndex        =   55
         Top             =   1575
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   840
         TabIndex        =   54
         Text            =   " "
         Top             =   1575
         Width           =   2850
      End
      Begin VB.CommandButton Command10 
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
         Left            =   4330
         Picture         =   "WLITAB07.frx":1F2C
         TabIndex        =   49
         ToolTipText     =   "Click para Seleccionar Carpeta de Salida"
         Top             =   2760
         Width           =   175
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00DAE4EB&
         Caption         =   "OPCIONES"
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
         Left            =   2460
         TabIndex        =   40
         Top             =   150
         Width           =   2070
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            BackColor       =   &H00DAE4EB&
            Caption         =   "Mostrar Archivo"
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
            Height          =   435
            Left            =   210
            TabIndex        =   43
            Top             =   720
            Width           =   1800
         End
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00DAE4EB&
            Caption         =   "Agregar al Final"
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
            Left            =   210
            TabIndex        =   42
            Top             =   250
            Width           =   1695
         End
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00DAE4EB&
            Caption         =   "Sobreescribir"
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
            Left            =   210
            TabIndex        =   41
            Top             =   500
            Width           =   1695
         End
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   2100
         TabIndex        =   37
         Text            =   " "
         Top             =   2415
         Width           =   2430
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         BackColor       =   &H00DAE4EB&
         Caption         =   "Cortes Parciales"
         Enabled         =   0   'False
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
         Height          =   435
         Left            =   720
         TabIndex        =   61
         Top             =   600
         Width           =   1200
      End
      Begin VB.Image Image1 
         Height          =   570
         Index           =   1
         Left            =   120
         Picture         =   "WLITAB07.frx":2236
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label18 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   240
         TabIndex        =   58
         ToolTipText     =   "Doble Click para Seleccionar Carpeta de Salida"
         Top             =   1920
         Width           =   4110
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Envio por Mail"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   230
         TabIndex        =   57
         Top             =   1470
         Width           =   660
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Excel Sheet"
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
         Left            =   720
         TabIndex        =   45
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Texto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   44
         Top             =   360
         Width           =   525
      End
      Begin VB.Image Image1 
         Height          =   420
         Index           =   2
         Left            =   210
         Picture         =   "WLITAB07.frx":3280
         Top             =   1020
         Width           =   435
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo / Carpeta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   585
         TabIndex        =   39
         Top             =   2490
         Width           =   1425
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   240
         TabIndex        =   38
         ToolTipText     =   "Doble Click para Seleccionar Carpeta de Salida"
         Top             =   2745
         Width           =   4080
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   480
      Top             =   0
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
      Left            =   3045
      TabIndex        =   22
      Top             =   6720
      Width           =   5160
      Begin VB.DriveListBox Drive1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2520
         TabIndex        =   25
         Top             =   315
         Width           =   1275
      End
      Begin VB.DirListBox Dir1 
         Enabled         =   0   'False
         Height          =   2115
         Left            =   2520
         TabIndex        =   24
         Top             =   735
         Width           =   2430
      End
      Begin VB.CommandButton Command4 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4200
         Picture         =   "WLITAB07.frx":3C62
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   300
         Width           =   540
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
      Left            =   5670
      Picture         =   "WLITAB07.frx":40DC
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Guarda Configuración del Listado"
      Top             =   7560
      Width           =   855
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00DAE4EB&
      Caption         =   "CONTROL DE PROCESO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   90
      TabIndex        =   10
      Top             =   105
      Width           =   9360
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   0
         TabIndex        =   64
         Top             =   0
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   4725
         X2              =   4725
         Y1              =   120
         Y2              =   1380
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   510
         Left            =   240
         TabIndex        =   30
         Top             =   735
         Width           =   4320
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   300
         Left            =   2970
         TabIndex        =   29
         Top             =   345
         Width           =   1590
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Controlador"
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
         Left            =   1785
         TabIndex        =   28
         Top             =   450
         Width           =   1065
      End
      Begin VB.Label Label6 
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
         Height          =   510
         Left            =   4830
         TabIndex        =   27
         Top             =   735
         Width           =   4320
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Height          =   300
         Left            =   7560
         TabIndex        =   26
         Top             =   345
         Width           =   1590
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo"
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
         Left            =   6405
         TabIndex        =   11
         Top             =   450
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DAE4EB&
      Caption         =   "IMPRIMIR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4160
      Left            =   90
      TabIndex        =   0
      Top             =   1560
      Width           =   4530
      Begin VB.Frame Frame13 
         BackColor       =   &H00DAE4EB&
         Caption         =   "ANCHO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   3360
         TabIndex        =   52
         Top             =   720
         Width           =   975
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Height          =   285
            Left            =   120
            TabIndex        =   53
            Top             =   525
            Width           =   735
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00DAE4EB&
         Caption         =   "RENUMERA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   1920
         TabIndex        =   50
         Top             =   1845
         Width           =   1215
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            BackColor       =   &H00DAE4EB&
            Caption         =   "Si"
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
            Left            =   210
            TabIndex        =   51
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00DAE4EB&
         Caption         =   "TIPO DE LETRA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   210
         TabIndex        =   31
         Top             =   720
         Width           =   3080
         Begin VB.TextBox Text7 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2400
            TabIndex        =   34
            Text            =   " "
            Top             =   270
            Width           =   540
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   195
            TabIndex        =   32
            Text            =   " "
            Top             =   630
            Width           =   2745
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tamaño"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   1455
            TabIndex        =   35
            Top             =   340
            Width           =   855
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Fuente"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   200
            TabIndex        =   33
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.ComboBox Combo1 
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
         Left            =   975
         TabIndex        =   4
         Text            =   " "
         Top             =   300
         Width           =   3345
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00DAE4EB&
         Caption         =   "ÁREA DE IMPRESION (MILÍMETROS)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   210
         TabIndex        =   12
         Top             =   2760
         Width           =   4110
         Begin VB.TextBox Text4 
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
            Height          =   285
            Index           =   3
            Left            =   3150
            TabIndex        =   19
            Top             =   735
            Width           =   855
         End
         Begin VB.TextBox Text4 
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
            Height          =   285
            Index           =   2
            Left            =   3150
            TabIndex        =   18
            Top             =   420
            Width           =   855
         End
         Begin VB.TextBox Text4 
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
            Height          =   285
            Index           =   1
            Left            =   720
            TabIndex        =   16
            Text            =   " "
            Top             =   735
            Width           =   855
         End
         Begin VB.TextBox Text4 
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
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   14
            Text            =   " "
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Margen Superior"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   1560
            TabIndex        =   21
            Top             =   500
            Width           =   1530
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Margen Izquierdo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   1605
            TabIndex        =   20
            Top             =   810
            Width           =   1455
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Largo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   -225
            TabIndex        =   15
            Top             =   810
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ancho"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   -225
            TabIndex        =   13
            Top             =   500
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00DAE4EB&
         Caption         =   "COPIAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   3225
         TabIndex        =   8
         Top             =   1845
         Width           =   1095
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   240
            TabIndex        =   9
            Top             =   315
            Width           =   675
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00DAE4EB&
         Caption         =   "ORIENTACIÓN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   210
         TabIndex        =   5
         Top             =   1845
         Width           =   1605
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00DAE4EB&
            Caption         =   "Horizontal"
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
            Left            =   240
            TabIndex        =   7
            Top             =   520
            Width           =   1200
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00DAE4EB&
            Caption         =   "Vertical"
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
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   1065
         End
      End
      Begin VB.Image Image1 
         Height          =   450
         Index           =   0
         Left            =   75
         Picture         =   "WLITAB07.frx":43E6
         Top             =   300
         Width           =   750
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "WLITAB07.frx":55F8
      TabIndex        =   60
      Top             =   120
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "WLITAB07.frx":5672
      Top             =   1200
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
End
Attribute VB_Name = "WLITAB07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IMPREDETER As String
Dim MODOSAL%
'
Private Sub BOTEXIT_Click()
    Call CIERRARCH("*.*")
    If COMMA$ <> "" Then
       Call BLOQEXE("*")
    End If
    WLITAB07.Hide
End Sub






Private Sub Combo1_LostFocus()
    '
    Dim PRX As Printer
    For i% = 1 To Printers.Count
        If Printers(i% - 1).DeviceName = Combo1.TEXT Then
            ' La define como predeterminada del sistema.
            Set Printer = Printers(i% - 1)
            GoTo 10
        End If
    Next
    '
10  Combo1.TEXT = Printer.DeviceName
    '
End Sub

Sub Command1_Click()
    '
    Call REGAUDITRA("LIST", TRIM(Label8) & " - " & TRIM(Label9), "")
    '
    If MODOSAL% = 3 Then
      PREVIEW07.List4.Clear
    End If
    '
    If MODOSAL% <> 0 And MODOSAL% <> 3 Then
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
       If Len(Label11) > 80 Then
          Call MENSERR(24, "Ruta de Salida Demasiado Larga.")
          Exit Sub
       End If
       ARCHISAL$ = TRIM$(Label11)
       If Len(ARCHISAL$) > 0 Then
         If Right$(ARCHISAL$, 1) <> "\" Then ARCHISAL$ = ARCHISAL$ + "\"
       End If
       ARCHISAL$ = ARCHISAL$ + ASALIDA$
       'If Option1(1).Value = True Then
       If MODOSAL% = 1 Then
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
            '18/04/07 (OT-07-0158)
            If Text1 <> "" Then
               If InStr(Text1, "@") > 0 Then
                  WINMAIL07.Text1 = TRIM$(Label18)
                     If WINMAIL07.Text1 = "" Then
                        WINMAIL07.Text1 = Text1
                     End If
                  WINMAIL07.Label2.Caption = Text1
                  WINMAIL07.Text2 = "Envío de Listado"
                  WINMAIL07.Label5.Caption = ARCHISAL$
                     ASALIXX$ = TRIM$(ARCHISAL$)
25                   PPXX% = InStr(ASALIXX$, "\")
                     If PPXX% > 0 Then
                        ASALIXX$ = TRIM$(Mid$(ASALIXX$, PPXX% + 1))
                        GoTo 25
                     End If
                  WINMAIL07.Text5 = "Enviamos Archivo Adjunto: " + ASALIXX$
                  WINMAIL07.Show 1
               Else
                  Call MENSERR(24, "Dirección de Mail No Valida")
                  Text1.SetFocus
                  Exit Sub
               End If
            End If
            '18/04/07 (FIN)
            '
           WLITAB07.Hide
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
            ' INCORPORADO 24/09/15 PARA MANTENER PARCIALES AL EXPORTAR A UN ARCHIVO DE TEXTO
            CorteParcialesCantidadRenglones% = 0
            If ARCHISAL$ <> "" And Check3.Value = 1 Then
                CorteParcialesCantidadRenglones% = 60 ' VALOR POR DEFAULT
                If XVALO(Text3) > 0 Then CorteParcialesCantidadRenglones% = XVALO(Text3)
            End If
            Call LISTABU(CorteParcialesCantidadRenglones%)
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
    '
    If ARCHISAL$ = "SCREEN" Then
      PREVIEW07.Show 1
      Exit Sub
    End If
    '
    On Error Resume Next
    Unload PREVIEW07
    On Error GoTo 0
    '
    '18/04/07 (OT-07-0158)
    If Text1 <> "" Then
       If InStr(Text1, "@") > 0 Then
          WINMAIL07.Text1 = TRIM$(Label18)
             If WINMAIL07.Text1 = "" Then
                WINMAIL07.Text1 = Text1
             End If
          WINMAIL07.Label2.Caption = Text1
          WINMAIL07.Label5.Caption = ARCHISAL$
          WINMAIL07.Text2 = "Envío de Listado"
             ASALIXX$ = TRIM$(ARCHISAL$)
30           PPXX% = InStr(ASALIXX$, "\")
             If PPXX% > 0 Then
                ASALIXX$ = TRIM$(Mid$(ASALIXX$, PPXX% + 1))
                GoTo 30
             End If
          WINMAIL07.Text5 = "Enviamos Archivo Adjunto: " + ASALIXX$
          WINMAIL07.Show 1
       Else
          Call MENSERR(24, "Dirección de Mail No Valida")
          Text1.SetFocus
          Exit Sub
       End If
    End If
    '18/04/07 (FIN)
    '
    WLITAB07.Hide
    '
    If ARCHISAL$ <> "" Then
        If WLITAB07.Check1.Value = 1 Then
            LCB$ = LUBAS$
            ULOWP$ = LUCOM$
            If Left$(JOBID$, 2) = "WF" Then ULOWP$ = App.Path + "\"
            If EXISTE%(ULOWP$ + "WORDPAD.EXE") = 1 Then
                 Close
                 Shell ULOWP$ + "WORDPAD.EXE " + ARCHISAL$, vbMaximizedFocus
               Else
                 MENSA$ = "Imposible Mostrar Archivo Generado."
                 MENSA$ = MENSA$ + "\  \Falta Utilitario 'WORDPAD.EXE'"
                 MENSA$ = MENSA$ + "\en Directorio '" + REPLACAR$(ULOWP$, "\", "/") + "'."
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

Sub Command10_Click()
   Command10.Enabled = False
   Call Label11_DblClick
   Command10.Enabled = True
End Sub

Private Sub Command2_Click()
    '\\\OT-5-622-FG-13/09/05///
    OPX% = COMALTER%("Opciones de Configuración:\\Impresora\Listado\Leyendas Fijas")
    Select Case OPX%
      Case 1
           Call SETUPRINTER
      Case 2
           Call CONECRUN("#PARALIS", "Definicion de Listadores Tabulados")
      Case 3
           Call LEYENLIS
           Call COMUNI("Configuración de Leyendas Completada.\  \Vuelva a Pedir el Listado.")
           Call FINAL("")
    End Select
    '\\\OT-5-622-FG-FIN///
End Sub

Private Sub Command3_Click()
    Call GRADATIMPRE
    Call CARDATIMPRE
End Sub


Private Sub Command5_Click()
   '
   '17/04/07 (OT-07-0157)
   Call SELECHO("MAILIST")
   RAZO1$ = TRIM$(VALACT1$("MAILIST"))
   If RAZO1$ <> "" Then
      Label18 = RAZO1$
      Text1 = TRIM$(VALACT2$("MAILIST"))
   End If
   '17/04/07 (FIN)
   '
End Sub

Sub Command6_Click()
   '
   '17/04/07 (OT-07-0157)
   J& = 0
   Call APERBASDAT("CLIEN")
   'Set CLIEN = Badaclie.OpenRecordset("SELECT * FROM Deudor12", dbOpenSnapshot)
   SQLX$ = "SELECT * FROM Deudor12"
   Dim CLIEN As ADODB.Recordset
   Set CLIEN = New ADODB.Recordset
   CLIEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   With CLIEN
     On Error Resume Next
     .MoveFirst
     If Err Then
        On Error GoTo 0
        Screen.MousePointer = 1
        GoTo 10
     End If
     '
     Do While Not .EOF
        '
        Y$ = REGBLAN$("MAILIST")
        On Error Resume Next
        RAZCLI$ = !raso_cli
        MAILCLI$ = !Mail_Cli
        On Error GoTo 0
        '
        If InStr(MAILCLI$, "@") > 0 Then
           Call REPLA(Y$, RAZCLI$, 1, 32)
           Call REPLA(Y$, LCase$(MAILCLI$), 33, 48)
           J& = J& + 1
           Call GRAREG("MAILIST", Y$, J&)
        End If
        '
        MAILCLI$ = ""
        RAZCLI$ = ""
        '
     .MoveNext
     Loop
     '
   End With
   Set CLIEN = Nothing
   '
10 Call APERBASDAT("PROVE")
   'Set PROVE = Badaprov.OpenRecordset("SELECT * FROM Proved12", dbOpenSnapshot)
   SQLX$ = "SELECT * FROM Proved12"
   Dim PROVE As ADODB.Recordset
   Set PROVE = New ADODB.Recordset
   PROVE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   With PROVE
     On Error Resume Next
     .MoveFirst
     If Err Then
        On Error GoTo 0
        Screen.MousePointer = 1
        GoTo 20
     End If
     '
     Do While Not .EOF
        '
        Y$ = REGBLAN$("MAILIST")
        On Error Resume Next
        RAZPROV$ = !raso_cli
        MAILPROV$ = !Mail_Cli
        On Error GoTo 0
        '
        If InStr(MAILPROV$, "@") > 0 Then
           Call REPLA(Y$, RAZPROV$, 1, 32)
           Call REPLA(Y$, LCase$(MAILPROV$), 33, 48)
           J& = J& + 1
           Call GRAREG("MAILIST", Y$, J&)
        End If
        '
        RAZPROV$ = ""
        MAILPROV$ = ""
        '
     .MoveNext
     Loop
     '
     Call SETULTREG("MAILIST", J&)
     '
20 End With
   Set PROVE = Nothing
   Call COMUNI("Operación Completa")
   Call FILESORT("MAILIST", "", 1, 1, "ASC")
   '17/04/07 (FIN)
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
   Refresh
   DoEvents
   Timer1.Enabled = True
   Timer2.Enabled = True
   MODOSAL% = 3
   DoEvents
   TINI = Timer: While Timer < TINI + 1: Wend
   UTILIZA_SKIN% = 0
   Call APLICACIONSKINS(Me)
End Sub
'

Sub Image1_Click(Index As Integer)
   MODOSAL% = Index
   If Index = 0 Then
        ARCHISAL$ = ""
        PRINTERSET% = 0
        IDL% = IDL0%
        Text2.TEXT = ""
        Text2.Enabled = False
        Label11 = ""
        Option3(0).Value = False
        Option3(0).Enabled = False
        Option3(1).Value = False
        Option3(1).Enabled = False
        Check1.Enabled = False
        Check3.Value = 0: Check3.Enabled = False
        Text3 = "": Text3.Enabled = False
        '17/04/07 (OT-07-0157)
        Text1.TEXT = ""
        '17/04/07 (fin)
        'Drive1.Enabled = False
        'Dir1.Enabled = False
        Command4.Enabled = False
        SALEXCEL% = 0
      ElseIf Index = 1 Then
        RUTASAL$ = TRIM$(Control$("EXPORTEX", IDENTER$))
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
        Check3.Enabled = True
        Text3.Enabled = True
        'Drive1.Enabled = True
        'Dir1.Enabled = True
        Command4.Enabled = True
        SALEXCEL% = 0
        Text2.Enabled = True
        Text2.SetFocus
        '17/04/07 (OT-07-0157)
        Text1.Enabled = True
        Command5.Enabled = True
        Command6.Enabled = True
        '17/04/07 (fin)
        On Error GoTo 0
      ElseIf Index = 2 Then
        RUTASAL$ = TRIM$(Control$("EXPOREXC", IDENTER$))
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
        Check3.Value = 0: Check3.Enabled = False
        Text3 = "": Text3.Enabled = False
        'Drive1.Enabled = True
        'Dir1.Enabled = True
        Command4.Enabled = True
        SALEXCEL% = 1
        Text2.Enabled = True
        Text2.SetFocus
        '17/04/07 (OT-07-0157)
        Text1.Enabled = True
        Command5.Enabled = True
        Command6.Enabled = True
        '17/04/07 (fin)
      ElseIf Index = 3 Then
        ARCHISAL$ = "SCREEN"
        '17/04/07 (OT-07-0157)
        Text1.TEXT = ""
        '17/04/07 (FIN)
        Text2.TEXT = ""
        Text2.Enabled = False
        Label11 = ""
        Option3(0).Value = False
        Option3(0).Enabled = False
        Option3(1).Value = False
        Option3(1).Enabled = False
        Check1.Enabled = False
        Check3.Value = 0: Check3.Enabled = False
        Text3 = "": Text3.Enabled = False
        Command4.Enabled = False
        SALEXCEL% = 0
    End If
End Sub

Private Sub Text1_DblClick()
    Call Command5_Click '17/04/07 (OT-07-0157)
End Sub

Sub CARDATLIS()
     '
5800 GRABACT% = 0
     TOTALI% = 0
     POSITOT% = 0
     '
     LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
     If LONCODI% < 15 Or LONCODI% > 24 Then LONCODI% = 15
     '
     Dim YYY As String * 128
     If N15% > 0 Then Close #N15%
     N15% = FILEOPEN%(FORMA$, 9, 128)
     Get #N15%, 1, YYY$
     '
     If LOF(N15%) <= 128 Then
         On Error Resume Next
         YYY$ = String$(128, 0)
         Put #N15%, 1, YYY$
         Close #N15%
         Kill FORMA$
         On Error GoTo 0
         FORMB$ = UCase$(REPLACAR$(FORMA$, "\", "/"))
         Call MENSERR(24, "Proceso Interrumpido.\Falta Controlador de Procesos '" + FORMB$ + "'.")
         Call FINAL("")
     End If
     '
     Archi$ = Mid$(YYY$, 57, 8)
     ARCHITRA$ = UCase$(TRIM$(Archi$))
     WLITAB07.Label5 = ARCHITRA$
     WLITAB07.Label6 = TITUARCH$(ARCHITRA$)
       If ARCHITRA$ = "RESCUCLI" Then ARCHITRA$ = "!RESCUCLI"
5801 Get #N15%, 1, YYY$
     LOR1% = Asc(Mid$(YYY$, 106, 1)): If LOR1% > 16 Then LOR1% = 0
     LOR% = 256 * LOR1% + Asc(Mid$(YYY$, 107))
     CAMAX& = CVI(Mid$(YYY$, 108))
     PEOF% = 1 + Asc(Mid$(YYY$, 110))
     EOFS$ = Mid$(YYY$, 111, 16)
     WLITAB07.Label8 = PROPRIN$
5803 s1$ = TRIM$(Left$(YYY$, 48))
     WLITAB07.Label9 = TRIM$(s1$)
     '
     EPAC$ = TRIM$(EMPREAC$)
     If EPAC$ <> "" Then
         WLITAB07.Caption = WLITAB07.Caption + "  -  " + EPAC$
     End If
    '
     sali$ = Mid$(YYY$, 49, 8)
     ARCHISAL$ = Mid$(YYY$, 65, 12)
     If Left$(ARCHISAL$, 3) = "LPT" Then ARCHISAL$ = ""
     COD$ = Mid$(YYY$, 81)
     RGI& = Asc(Mid$(COD$, 1))
     RGF& = Asc(Mid$(COD$, 2))
     OFS% = Asc(Mid$(COD$, 3))
5804 BLP% = Asc(Mid$(COD$, 4))
        CXC$ = "": If BLP% = 1 Then CXC$ = "¦"
     VIV% = Asc(Mid$(COD$, 5))
     APQ% = Asc(Mid$(COD$, 6))
     IDL% = REHOJ%: REHOJ1% = REHOJ%
     If Asc(Mid$(COD$, 7)) < REHOJ% Then IDL% = Asc(Mid$(COD$, 7))
     '
     If VIV% = 1 Then IDL% = 8 * IDL% / 6
     IDL0% = IDL%
     REHOJX% = 0
     If Asc(Mid$(COD$, 7)) > 96 Then
       REHOJX% = Asc(Mid$(COD$, 7))
     End If
     SUPRIFE% = 0
     HDR& = Asc(Mid$(COD$, 8))
     If HDR& = 99 Then
        HDR& = 0
        SUPRIFE% = 1
     End If
     RFT% = Asc(Mid$(COD$, 9))
     NCL% = 0: TCL% = 0: TNC% = 0
     RG% = 1: HST% = 0
     HOJINI% = CVI(Mid$(COD$, 10, 2))
5805 RG% = RG% + 1
     If RG% > LOF(N15%) / 128 Or RG% > 49 Then GoTo 5820
5806 Get #N15%, RG%, YYY$
     If YYY$ = String$(128, 0) Then GoTo 5820
     '
5807 TCL% = TCL% + 1
     TVL%(TCL%) = Asc(Mid$(YYY$, 33))
         If RG% = 2 Then TVL%(TCL%) = 1
     LDL%(TCL%) = Asc(Mid$(YYY$, 34))
5808 PR%(TCL%) = Asc(Mid$(YYY$, 35))
     While PR%(TCL%) < PR%(TCL% - 1)
       PR%(TCL%) = PR%(TCL%) + 256
     Wend
     LF%(TCL%) = Asc(Mid$(YYY$, 36))
         If RG% = 2 Then PR%(TCL%) = 0: LF%(TCL%) = 2
         '
     If Asc(Mid$(YYY$, 72)) = 99 Then
        TNC% = TNC% + 1
        CLNC%(TNC%) = TCL%
     End If
     '
5811 If Asc(Mid$(YYY$, 69)) < 1 Then GoTo 5817   ' SUPRIME COLUMNAS LONGITUD 0
5812 NCL% = NCL% + 1
     CFT%(NCL%) = RG% - 1
     CL0%(NCL%) = RG% - 1
     ECL$(NCL%) = TRIM$(Mid$(YYY$, 37, 32))
5813 ACL%(NCL%) = Asc(Mid$(YYY$, 69))
     LK%(NCL%) = Asc(Mid$(YYY$, 70))
     '
     If TVL%(TCL%) = 7 Then   ' ES UN CODIGO
        If LONCODI% > 15 Then
           ACL%(NCL%) = ACL%(NCL%) + LONCODI% - 15
        End If
     End If
     '
5814 MG%(NCL%) = Asc(Mid$(YYY$, 71))
     TT%(NCL%) = Asc(Mid$(YYY$, 73))
        If TT%(NCL%) = 1 Then TT%(NCL%) = TCL%
        If TT%(NCL%) = 255 Then HST% = 1
5815 PE%(NCL%) = 1
        If RG% = 2 Then PE%(NCL%) = 2
5816 JT%(NCL%) = Asc(Mid$(YYY$, 72))
     ae$(NCL%) = Mid$(YYY$, 74, 8)
     CE%(NCL%) = Asc(Mid$(YYY$, 82))
     If TT%(NCL%) > 0 Then
        NTT% = NTT% + 1
        TOTALI% = 1
     End If
     '
     If POSITOT% = 0 Then
        If TT%(NCL%) < 1 Then
           If TVL%(TCL%) = 6 Then
              If ACL%(NCL%) >= 10 Then
                 POSITOT% = NCL%
              End If
           End If
        End If
     End If
     
5817 GoTo 5805
     '
5820 If POSITOT% = 0 Then POSITOT% = 1
     ATOT% = BLP% * (NCL% + 1)
     CLII% = OFS% + 2
     For KCL% = 1 To NCL%
          lC%(KCL%) = ACL%(KCL%)
          ATOT% = ATOT% + ACL%(KCL%) + MG%(KCL%)
          cl%(KCL%) = CLII%
          CLII% = CLII% + ACL%(KCL%) + 1
     Next KCL%
5821 For KCL% = 1 To NCL%
          If CE%(KCL%) <= 0 Then GoTo 5824
5822      For KCL1% = 1 To NCL%
              If CL0%(KCL1%) = CE%(KCL%) Then CE%(KCL%) = KCL1%: PE%(KCL1%) = 3: GoTo 5824
5823      Next KCL1%
5824 Next KCL%
5825 For KCL% = 1 To NCL%
          If TT%(KCL%) <= 0 Then GoTo 5828
5826      For KCL1% = 1 To NCL%
             If CL0%(KCL1%) = TT%(KCL%) Then TT%(KCL%) = KCL1%: GoTo 5828
             If 100 + CL0%(KCL1%) = TT%(KCL%) Then TT%(KCL%) = 100 + KCL1%: GoTo 5828
5827      Next KCL1%
5828 Next KCL%
     '
5830 NFT% = 0    ' NO HAY FILTROS
5850 If N15% > 0 Then
         Close #N15%
         N15% = 0
     End If
     '
     Label15 = TRIM$(Str$(ATOT%))
     '
End Sub
'
        

Private Sub Image1_DblClick(Index As Integer)
   Call Image1_Click(Index)
   Call Command1_Click
End Sub

Private Sub Label11_DblClick()
   If Text2.Enabled = True Then
     SELCARP07.Show 1
   End If
End Sub

Private Sub Timer1_Timer()
   Static YACAR%, VECES
   If YACAR% = 0 Then
       Screen.MousePointer = 11
    On Error Resume Next
    PORIAN = Printer.Orientation
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
    'Option1(0).Value = True        ' IMPRESORA
    MODOSAL% = 0
    Call Image1_Click(0)
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
        Refresh
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
        Hide
    End If
    '
    Printer.Orientation = PORIAN
    'If EXISTE%("FLEXCEL.EXE") > 0 Then
      'Option1(2).Enabled = True
      Image1(2).Enabled = True
    'End If
    DoEvents
    Screen.MousePointer = 1
    YACAR% = 1
    AppActivate Caption
    SendKeys Chr$(1)
  End If
  '
  VECES = VECES + 1
  If VECES > 30 And VECES < 32 Then
    AppActivate Caption
    SendKeys Chr$(1)
    DoEvents
  End If
  '
End Sub

Private Sub Timer2_Timer()
   For KKI% = 0 To 3
     If KKI% <> MODOSAL% Then
         Image1(KKI%).Visible = True
       Else
         Image1(KKI%).Visible = Not Image1(KKI%).Visible
     End If
   Next KKI%
End Sub


Private Sub LEYENLIS()
    '\\\OT-5-622-FG-13/09/05///
    '
    'If TRIM$(CODFOR$) = "" Then
    '    Call MENSERR(24, "Falta Seleccionar Formulario")
    '    PRINFORM.Text1.SetFocus
    '    Exit Sub
    'End If
    '
    CODFOR$ = TRIM$(Label8)
    NX% = FILEOPEN%(LU$ + CODFOR$ + ".PRL", 0, 128)
    Dim RF As String * 128
    Call CIERRARCH("!CARFORMW")
    Call BLANARCH("!CARFORMW")
    For i% = 1 To LOF(NX%) / 128
        Get #NX%, i%, RF$
        If TRIM$(Left$(RF$, 8)) = "TEXTWIN" Then
            RFF$ = RF$
            Call REGAPP("!CARFORMW", RFF$)
        End If
    Next i%
    '
    VTB% = VENTABU%("CARLEYEW")
    If VTB% >= 0 Then
        J% = 0
        For i% = 1 To LOF(NX%) / 128
            Get #NX%, i%, RF$
            If TRIM$(Left$(RF$, 8)) <> "TEXTWIN" Then
                J% = J% + 1
                Put #NX%, J%, RF$
            End If
        Next i%
        '
        For U& = 1 To ULTREG("!CARFORMW")
            RFF$ = REGLEIDO$("!CARFORMW", U&)
            Call REPLA(RFF$, "TEXTWIN", 1, 8)
            RF$ = RFF$
            J% = J% + 1
            Put #NX%, J%, RF$
        Next U&
        '
        For i% = J% + 1 To LOF(NX%) / 128
            RF$ = String$(128, 0)
            J% = J% + 1
            Put #NX%, J%, RF$
        Next i%
        '
    End If
    Close #NX%: NX% = 0
    '
    '\\\OT-5-622-FG-FIN///
End Sub

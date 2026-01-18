VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form REPRFLEX04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D8D8C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuestación Sistemas FLEXOFT"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   11910
   Icon            =   "REPRFLEX04.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   11910
   Begin VB.CommandButton Command7 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   105
      TabIndex        =   187
      Top             =   1155
      Width           =   960
   End
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
      Height          =   285
      Left            =   5145
      TabIndex        =   155
      Top             =   1155
      Width           =   3200
   End
   Begin VB.TextBox Text1 
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
      Left            =   1155
      TabIndex        =   59
      Top             =   210
      Width           =   2010
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
      Height          =   915
      Left            =   1155
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   58
      Top             =   525
      Width           =   2925
   End
   Begin VB.TextBox Text8 
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
      Left            =   6090
      TabIndex        =   57
      Top             =   840
      Width           =   2240
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
      Left            =   9660
      TabIndex        =   56
      Top             =   210
      Width           =   1125
   End
   Begin VB.CommandButton Command16 
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
      Left            =   10755
      Picture         =   "REPRFLEX04.frx":0442
      TabIndex        =   55
      Top             =   525
      Width           =   175
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
      Left            =   8700
      TabIndex        =   54
      Top             =   1155
      Width           =   2070
   End
   Begin VB.TextBox Text10 
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
      Left            =   9660
      TabIndex        =   53
      Top             =   525
      Width           =   1125
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
      Left            =   5145
      TabIndex        =   52
      Top             =   525
      Width           =   3200
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
      Left            =   3150
      Picture         =   "REPRFLEX04.frx":074C
      TabIndex        =   51
      Top             =   200
      Width           =   175
   End
   Begin VB.TextBox Text18 
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
      Left            =   5145
      TabIndex        =   50
      Top             =   210
      Width           =   3200
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
      Left            =   5145
      TabIndex        =   49
      Top             =   840
      Width           =   870
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D8D8C0&
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
      Height          =   880
      Left            =   8160
      TabIndex        =   25
      Top             =   7920
      Visible         =   0   'False
      Width           =   1990
      Begin VB.CommandButton Command20 
         Height          =   500
         Left            =   1280
         Picture         =   "REPRFLEX04.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Búsqueda por CUIT, Domicilio, Nombre de Fantasía, etc"
         Top             =   250
         Width           =   550
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
         Height          =   500
         Left            =   155
         Picture         =   "REPRFLEX04.frx":1DBC
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   250
         Width           =   550
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
         Height          =   500
         Left            =   725
         Picture         =   "REPRFLEX04.frx":20C6
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Restablece los Valores Anteriores - Descarta Cambios Efectuados"
         Top             =   250
         Width           =   550
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   8100
      Left            =   11025
      ScaleHeight     =   8040
      ScaleWidth      =   1215
      TabIndex        =   15
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command8 
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
         Picture         =   "REPRFLEX04.frx":23D0
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Acceso a Base de Datos - Clientes  y Articulos"
         Top             =   1575
         Width           =   640
      End
      Begin VB.CommandButton Command15 
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
         Height          =   690
         Left            =   105
         Picture         =   "REPRFLEX04.frx":26DA
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Guardar "
         Top             =   2310
         Width           =   640
      End
      Begin VB.CommandButton Command22 
         Caption         =   "Copy"
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
         Left            =   105
         Picture         =   "REPRFLEX04.frx":29E4
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Alta de un Nuevo Presupuesto por Copia de uno Existente"
         Top             =   3690
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "REPRFLEX04.frx":2B2E
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Ver y Re-Imprimir"
         Top             =   4380
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   23
         Top             =   7245
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   24
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command25 
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
         Height          =   690
         Left            =   105
         Picture         =   "REPRFLEX04.frx":2E38
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Imprime Presupuesto"
         Top             =   5835
         Width           =   650
      End
      Begin VB.CommandButton Command13 
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
         Height          =   690
         Left            =   105
         Picture         =   "REPRFLEX04.frx":34A2
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Listados Clasificados"
         Top             =   5145
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "REPRFLEX04.frx":37AC
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   6525
         Width           =   650
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
         Picture         =   "REPRFLEX04.frx":3BEE
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   8100
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Height          =   690
         Left            =   105
         Picture         =   "REPRFLEX04.frx":3EF8
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Exit"
         Top             =   105
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "REPRFLEX04.frx":4042
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Ayuda en Línea"
         Top             =   735
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
         Left            =   105
         Picture         =   "REPRFLEX04.frx":434C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Alta de un Nuevo Presupuesto"
         Top             =   3000
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -240
         Picture         =   "REPRFLEX04.frx":4496
         Stretch         =   -1  'True
         Top             =   7455
         Width           =   1410
      End
   End
   Begin VB.ListBox LINDICE 
      Height          =   1035
      Left            =   10920
      Sorted          =   -1  'True
      TabIndex        =   14
      Top             =   3780
      Visible         =   0   'False
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6360
      Left            =   120
      TabIndex        =   5
      Top             =   1600
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11218
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   14211264
      TabCaption(0)   =   "Presupuesto"
      TabPicture(0)   =   "REPRFLEX04.frx":63B8
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(2)=   "Frame6"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Rubros / Cotización"
      TabPicture(1)   =   "REPRFLEX04.frx":63D4
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame20"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame15"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Condiciones"
      TabPicture(2)   =   "REPRFLEX04.frx":63F0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame16"
      Tab(2).Control(1)=   "Frame10"
      Tab(2).Control(2)=   "Frame12"
      Tab(2).Control(3)=   "Frame8"
      Tab(2).Control(4)=   "Frame13"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Archivos Adjuntos"
      TabPicture(3)   =   "REPRFLEX04.frx":640C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame19"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame19 
         BorderStyle     =   0  'None
         Caption         =   "Frame19"
         Height          =   5895
         Left            =   -74895
         TabIndex        =   197
         Top             =   420
         Width           =   10515
         Begin VB.Frame Frame24 
            Caption         =   "Selección Archivo Adjunto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3690
            Left            =   105
            TabIndex        =   202
            Top             =   1995
            Width           =   10260
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
               Height          =   2460
               IntegralHeight  =   0   'False
               Left            =   3570
               TabIndex        =   209
               Top             =   420
               Width           =   6525
            End
            Begin VB.DriveListBox Drive1 
               Height          =   315
               Left            =   1260
               TabIndex        =   213
               Top             =   420
               Width           =   2220
            End
            Begin VB.DirListBox Dir1 
               Height          =   1440
               Left            =   315
               TabIndex        =   212
               Top             =   945
               Width           =   3165
            End
            Begin VB.FileListBox File1 
               Height          =   2040
               Left            =   3780
               TabIndex        =   211
               Top             =   420
               Width           =   1695
            End
            Begin VB.ListBox List4 
               Height          =   2205
               Left            =   5670
               Sorted          =   -1  'True
               TabIndex        =   210
               Top             =   420
               Width           =   1170
            End
            Begin VB.Frame Frame25 
               Caption         =   "Filtros"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1065
               Left            =   315
               TabIndex        =   203
               Top             =   2460
               Width           =   3165
               Begin VB.TextBox Text41 
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
                  Left            =   210
                  TabIndex        =   208
                  Text            =   "*.PDF"
                  Top             =   630
                  Width           =   2745
               End
               Begin VB.OptionButton Option8 
                  Caption         =   "PDF"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   210
                  TabIndex        =   207
                  Top             =   315
                  Value           =   -1  'True
                  Width           =   645
               End
               Begin VB.OptionButton Option7 
                  Caption         =   "DOC"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   915
                  TabIndex        =   206
                  Top             =   315
                  Width           =   645
               End
               Begin VB.OptionButton Option6 
                  Caption         =   "XLS"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   1620
                  TabIndex        =   205
                  Top             =   315
                  Width           =   645
               End
               Begin VB.OptionButton Option5 
                  Caption         =   "ALL"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Left            =   2325
                  TabIndex        =   204
                  Top             =   315
                  Width           =   645
               End
            End
            Begin VB.Label Label87 
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
               Left            =   3255
               TabIndex        =   217
               Top             =   2980
               Width           =   1065
            End
            Begin VB.Label Label86 
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
               Left            =   3990
               TabIndex        =   216
               Top             =   3225
               Width           =   6105
            End
            Begin VB.Label Label85 
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
               Left            =   4305
               TabIndex        =   215
               Top             =   1995
               Width           =   4320
            End
            Begin VB.Label Label84 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Unidad:"
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
               Left            =   105
               TabIndex        =   214
               Top             =   465
               Width           =   1065
            End
         End
         Begin VB.Frame Frame22 
            Caption         =   "Lista de Archivos Adjuntos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   1695
            Left            =   105
            TabIndex        =   200
            Top             =   210
            Width           =   9360
            Begin VB.ListBox List2 
               BeginProperty Font 
                  Name            =   "Lucida Console"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1140
               Left            =   210
               TabIndex        =   201
               Top             =   350
               Width           =   9000
            End
         End
         Begin VB.CommandButton Command12 
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
            Height          =   735
            Left            =   9580
            Picture         =   "REPRFLEX04.frx":6428
            Style           =   1  'Graphical
            TabIndex        =   199
            ToolTipText     =   "Adjuntar / Procesar"
            Top             =   1155
            Width           =   750
         End
         Begin VB.CommandButton Command10 
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
            Height          =   750
            Left            =   9580
            Picture         =   "REPRFLEX04.frx":6732
            Style           =   1  'Graphical
            TabIndex        =   198
            Top             =   315
            Width           =   750
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Valorización en:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   1575
         TabIndex        =   188
         Top             =   6195
         Visible         =   0   'False
         Width           =   3615
         Begin VB.TextBox TICAMBIS 
            Alignment       =   1  'Right Justify
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
            Height          =   300
            Left            =   885
            MaxLength       =   14
            TabIndex        =   189
            Top             =   330
            Width           =   1065
         End
         Begin VB.Label Label81 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Pesos"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   900
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   2100
            TabIndex        =   191
            Top             =   315
            Width           =   1470
         End
         Begin VB.Label Label80 
            BackStyle       =   0  'Transparent
            Caption         =   "T-C:"
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
            Left            =   465
            TabIndex        =   190
            Top             =   405
            Width           =   435
         End
      End
      Begin VB.Frame Frame16 
         Caption         =   "Validez de Oferta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   -69435
         TabIndex        =   168
         Top             =   2835
         Width           =   4935
         Begin VB.TextBox Text32 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Left            =   210
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   169
            Top             =   315
            Width           =   4515
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Forma de Pago"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   -69435
         TabIndex        =   166
         Top             =   600
         Width           =   5040
         Begin VB.TextBox Text42 
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
            Left            =   4200
            TabIndex        =   219
            Text            =   "0"
            Top             =   1680
            Width           =   345
         End
         Begin VB.TextBox Text35 
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
            Left            =   3015
            TabIndex        =   176
            Text            =   "0.0"
            Top             =   1680
            Width           =   450
         End
         Begin VB.TextBox Text34 
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
            Left            =   1995
            TabIndex        =   173
            Text            =   "1"
            Top             =   1680
            Width           =   450
         End
         Begin VB.TextBox Text33 
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
            Left            =   525
            TabIndex        =   170
            Text            =   "25"
            Top             =   1680
            Width           =   555
         End
         Begin VB.TextBox Text16 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1305
            Left            =   210
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   167
            Top             =   315
            Width           =   4620
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ds."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   8
            Left            =   4620
            TabIndex        =   221
            Top             =   1755
            Width           =   255
         End
         Begin VB.Label Label88 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "V-1:"
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
            Left            =   3675
            TabIndex        =   220
            Top             =   1740
            Width           =   540
         End
         Begin VB.Label Label57 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Int:"
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
            Left            =   1995
            TabIndex        =   177
            Top             =   1740
            Width           =   960
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   3405
            TabIndex        =   175
            Top             =   1755
            Width           =   255
         End
         Begin VB.Label Label56 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ctas:"
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
            Left            =   990
            TabIndex        =   174
            Top             =   1740
            Width           =   960
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   1050
            TabIndex        =   172
            Top             =   1755
            Width           =   255
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ant:"
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
            Left            =   -1420
            TabIndex        =   171
            Top             =   1740
            Width           =   1905
         End
      End
      Begin VB.Frame Frame12 
         Caption         =   "Otras Condiciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1755
         Left            =   -74640
         TabIndex        =   164
         Top             =   4410
         Width           =   10100
         Begin VB.TextBox Text29 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1260
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   165
            Top             =   360
            Width           =   9705
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Plazo de Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         Left            =   -74640
         TabIndex        =   162
         Top             =   2415
         Width           =   5040
         Begin VB.VScrollBar VScroll2 
            Height          =   285
            Left            =   3885
            Max             =   365
            TabIndex        =   181
            Top             =   1470
            Value           =   336
            Width           =   175
         End
         Begin VB.TextBox Text39 
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
            Left            =   3465
            TabIndex        =   178
            Text            =   "30"
            Top             =   1470
            Width           =   450
         End
         Begin VB.TextBox Text15 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   210
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   163
            Top             =   315
            Width           =   4605
         End
         Begin VB.Label Label59 
            BackStyle       =   0  'Transparent
            Caption         =   "Días"
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
            Left            =   4200
            TabIndex        =   180
            Top             =   1530
            Width           =   960
         End
         Begin VB.Label Label58 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Entrega / Instalación Primaria a los: "
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
            Left            =   315
            TabIndex        =   179
            Top             =   1530
            Width           =   3165
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Valores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5600
         Left            =   6510
         TabIndex        =   105
         Top             =   580
         Width           =   4005
         Begin VB.Frame Frame21 
            BorderStyle     =   0  'None
            Caption         =   " "
            Height          =   1875
            Left            =   105
            TabIndex        =   125
            Top             =   210
            Width           =   3795
            Begin VB.TextBox Text28 
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
               ForeColor       =   &H000000FF&
               Height          =   285
               Left            =   2505
               TabIndex        =   136
               Top             =   1450
               Width           =   1185
            End
            Begin VB.TextBox Text22 
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
               Left            =   1545
               TabIndex        =   133
               Top             =   1125
               Width           =   660
            End
            Begin VB.TextBox Text19 
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
               ForeColor       =   &H000000FF&
               Height          =   285
               Left            =   2505
               TabIndex        =   132
               Top             =   1125
               Width           =   1185
            End
            Begin VB.TextBox Text20 
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
               Left            =   1545
               TabIndex        =   127
               Top             =   840
               Width           =   660
            End
            Begin VB.TextBox Text26 
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
               ForeColor       =   &H000000FF&
               Height          =   285
               Left            =   2505
               TabIndex        =   126
               Top             =   810
               Width           =   1185
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
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
               Left            =   2505
               TabIndex        =   150
               Top             =   105
               Width           =   1185
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Transporte:"
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
               Left            =   -840
               TabIndex        =   149
               Top             =   180
               Width           =   3165
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Bonificaciones Varias:"
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
               Left            =   -420
               TabIndex        =   137
               Top             =   1540
               Width           =   2850
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Implementación:"
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
               Left            =   -420
               TabIndex        =   135
               Top             =   1215
               Width           =   1905
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   3
               Left            =   2145
               TabIndex        =   134
               Top             =   1215
               Width           =   255
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Licencias:"
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
               Left            =   -420
               TabIndex        =   131
               Top             =   900
               Width           =   1905
            End
            Begin VB.Label Label72 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Descuentos -    Bonificaciones:"
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
               Left            =   -105
               TabIndex        =   130
               Top             =   315
               Width           =   1545
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   2
               Left            =   2145
               TabIndex        =   129
               Top             =   915
               Width           =   255
            End
            Begin VB.Label Label83 
               Alignment       =   1  'Right Justify
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
               Left            =   4680
               TabIndex        =   128
               Top             =   80
               Width           =   1185
            End
         End
         Begin VB.Frame Frame23 
            BorderStyle     =   0  'None
            Height          =   3405
            Left            =   105
            TabIndex        =   106
            Top             =   2100
            Width           =   3795
            Begin VB.CheckBox Check1 
               Caption         =   "Suprimir"
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
               Left            =   2730
               TabIndex        =   218
               Top             =   2520
               Width           =   1485
            End
            Begin VB.TextBox Text23 
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
               Left            =   1140
               TabIndex        =   108
               Top             =   840
               Width           =   1185
            End
            Begin VB.TextBox Text27 
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
               Left            =   1140
               TabIndex        =   107
               Top             =   1140
               Width           =   1185
            End
            Begin VB.Label Label69 
               BackStyle       =   0  'Transparent
               Caption         =   "Mantenimiento Mensual:"
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
               TabIndex        =   196
               Top             =   2520
               Width           =   2430
            End
            Begin VB.Label Label65 
               Alignment       =   1  'Right Justify
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
               Index           =   1
               Left            =   1515
               TabIndex        =   195
               ToolTipText     =   "Doble Click para Abrir Detalle de Módulos"
               Top             =   3045
               Width           =   975
            End
            Begin VB.Label Label65 
               Alignment       =   1  'Right Justify
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
               Index           =   0
               Left            =   315
               TabIndex        =   194
               ToolTipText     =   "Doble Click para Abrir Detalle de Módulos"
               Top             =   3045
               Width           =   975
            End
            Begin VB.Label Label67 
               BackStyle       =   0  'Transparent
               Caption         =   "  Básico:          Kaizen:         Platinum:"
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
               Left            =   0
               TabIndex        =   193
               Top             =   2800
               Width           =   3690
            End
            Begin VB.Label Label65 
               Alignment       =   1  'Right Justify
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
               Index           =   2
               Left            =   2730
               TabIndex        =   192
               ToolTipText     =   "Doble Click para Abrir Detalle de Módulos"
               Top             =   3045
               Width           =   975
            End
            Begin VB.Line Line5 
               X1              =   0
               X2              =   3780
               Y1              =   2415
               Y2              =   2415
            End
            Begin VB.Line Line7 
               X1              =   0
               X2              =   5880
               Y1              =   0
               Y2              =   0
            End
            Begin VB.Label Label46 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Resumen          de Cotización:"
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
               Left            =   -105
               TabIndex        =   124
               Top             =   105
               Width           =   1545
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total Neto:"
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
               Left            =   420
               TabIndex        =   123
               Top             =   600
               Width           =   1905
            End
            Begin VB.Label Label49 
               Alignment       =   1  'Right Justify
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
               Left            =   2505
               TabIndex        =   122
               Top             =   525
               Width           =   1185
            End
            Begin VB.Label Label51 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "IVA"
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
               Left            =   -1575
               TabIndex        =   121
               Top             =   885
               Width           =   1905
            End
            Begin VB.Label Label54 
               Alignment       =   1  'Right Justify
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
               Left            =   2505
               TabIndex        =   120
               Top             =   840
               Width           =   1185
            End
            Begin VB.Label Label71 
               Alignment       =   1  'Right Justify
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
               Left            =   2505
               TabIndex        =   119
               Top             =   1140
               Width           =   1185
            End
            Begin VB.Label Label73 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "IVA"
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
               Left            =   -1575
               TabIndex        =   118
               Top             =   1200
               Width           =   1905
            End
            Begin VB.Label Label74 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Perc. I.B.:"
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
               Left            =   -900
               TabIndex        =   117
               Top             =   1545
               Width           =   1905
            End
            Begin VB.Label Label75 
               Alignment       =   1  'Right Justify
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
               Left            =   1140
               TabIndex        =   116
               Top             =   1440
               Width           =   1185
            End
            Begin VB.Label Label76 
               Alignment       =   1  'Right Justify
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
               Left            =   2505
               TabIndex        =   115
               Top             =   1440
               Width           =   1185
            End
            Begin VB.Label Label77 
               Alignment       =   1  'Right Justify
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
               Left            =   2505
               TabIndex        =   114
               Top             =   1905
               Width           =   1185
            End
            Begin VB.Label Label78 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total General Neto:"
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
               Left            =   450
               TabIndex        =   113
               Top             =   1980
               Width           =   1905
            End
            Begin VB.Label Label79 
               BackStyle       =   0  'Transparent
               Caption         =   "%:"
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
               Left            =   900
               TabIndex        =   112
               Top             =   885
               Width           =   255
            End
            Begin VB.Label Label82 
               BackStyle       =   0  'Transparent
               Caption         =   "%:"
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
               Left            =   900
               TabIndex        =   111
               Top             =   1200
               Width           =   255
            End
            Begin VB.Label PORIVA 
               BackStyle       =   0  'Transparent
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
               Index           =   0
               Left            =   350
               TabIndex        =   110
               Top             =   885
               Width           =   495
            End
            Begin VB.Label PORIVA 
               BackStyle       =   0  'Transparent
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
               Index           =   1
               Left            =   350
               TabIndex        =   109
               Top             =   1200
               Width           =   495
            End
         End
      End
      Begin VB.Frame Frame13 
         Caption         =   "Precios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -74640
         TabIndex        =   76
         Top             =   600
         Width           =   5040
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
            Height          =   1200
            Left            =   210
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   77
            Top             =   315
            Width           =   4620
         End
      End
      Begin VB.Frame Frame20 
         Caption         =   "Resumen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5600
         Left            =   210
         TabIndex        =   75
         Top             =   580
         Width           =   6135
         Begin VB.Frame Frame18 
            BorderStyle     =   0  'None
            Caption         =   "Frame18"
            Height          =   1935
            Left            =   120
            TabIndex        =   91
            Top             =   2520
            Width           =   5895
            Begin VB.TextBox Text31 
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
               Left            =   2190
               TabIndex        =   142
               Top             =   1575
               Width           =   585
            End
            Begin VB.TextBox Text30 
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
               Left            =   3240
               TabIndex        =   141
               Top             =   1155
               Width           =   1185
            End
            Begin VB.TextBox Text14 
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
               Left            =   2145
               TabIndex        =   101
               Top             =   330
               Width           =   585
            End
            Begin VB.TextBox Text44 
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
               Left            =   2145
               TabIndex        =   95
               Top             =   15
               Width           =   585
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Servicios Consultoría:"
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
               Left            =   105
               TabIndex        =   145
               Top             =   1665
               Width           =   2010
            End
            Begin VB.Label Label28 
               Alignment       =   1  'Right Justify
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
               Left            =   3240
               TabIndex        =   144
               Top             =   1575
               Width           =   1185
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hs"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   5
               Left            =   2880
               TabIndex        =   143
               Top             =   1695
               Width           =   255
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Formularios Electronicos:"
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
               Left            =   420
               TabIndex        =   140
               Top             =   1245
               Width           =   2745
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Migracion de Datos:"
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
               Left            =   420
               TabIndex        =   139
               Top             =   720
               Width           =   2745
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
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
               Left            =   3240
               TabIndex        =   138
               Top             =   630
               Width           =   1185
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Entrenamiento:"
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
               Left            =   165
               TabIndex        =   104
               Top             =   420
               Width           =   1905
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
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
               Left            =   3240
               TabIndex        =   103
               Top             =   330
               Width           =   1185
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hs"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   0
               Left            =   2880
               TabIndex        =   102
               Top             =   420
               Width           =   255
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Hs"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   4
               Left            =   2880
               TabIndex        =   96
               Top             =   135
               Width           =   255
            End
            Begin VB.Label Label70 
               Alignment       =   1  'Right Justify
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
               Left            =   4680
               TabIndex        =   94
               Top             =   1575
               Width           =   1185
            End
            Begin VB.Label Label68 
               Alignment       =   1  'Right Justify
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
               Left            =   3240
               TabIndex        =   93
               Top             =   15
               Width           =   1185
            End
            Begin VB.Label Label66 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Instalación:"
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
               Left            =   165
               TabIndex        =   92
               Top             =   105
               Width           =   1905
            End
         End
         Begin VB.Frame Frame14 
            BorderStyle     =   0  'None
            Height          =   780
            Left            =   120
            TabIndex        =   90
            Top             =   1890
            Width           =   5895
            Begin VB.Label Label64 
               BackStyle       =   0  'Transparent
               Caption         =   "b) Implementación:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   105
               TabIndex        =   100
               Top             =   315
               Width           =   2445
            End
            Begin VB.Line Line4 
               X1              =   0
               X2              =   5880
               Y1              =   120
               Y2              =   120
            End
         End
         Begin VB.Frame Frame9 
            BorderStyle     =   0  'None
            Height          =   1750
            Left            =   105
            TabIndex        =   79
            Top             =   210
            Width           =   5895
            Begin VB.VScrollBar VScroll1 
               Height          =   285
               Left            =   1890
               Max             =   99
               TabIndex        =   161
               Top             =   735
               Value           =   99
               Width           =   175
            End
            Begin VB.TextBox Text4 
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
               Left            =   1470
               TabIndex        =   98
               Text            =   "1"
               Top             =   735
               Width           =   435
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   1
               Left            =   2835
               TabIndex        =   160
               Top             =   840
               Width           =   255
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
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
               Left            =   2310
               TabIndex        =   159
               ToolTipText     =   "Doble Click para Abrir Cotización de Accesorios"
               Top             =   735
               Width           =   555
            End
            Begin VB.Label Label53 
               BackStyle       =   0  'Transparent
               Caption         =   "a) Licencias:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   450
               Left            =   105
               TabIndex        =   89
               Top             =   105
               Width           =   2025
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
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
               Left            =   3240
               TabIndex        =   88
               ToolTipText     =   "Doble Click para Abrir Cotización de Varios"
               Top             =   1365
               Width           =   1185
            End
            Begin VB.Label Label44 
               Alignment       =   1  'Right Justify
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
               Left            =   3240
               TabIndex        =   87
               ToolTipText     =   "Doble Click para Abrir Detalle de Desarrollos Especiales"
               Top             =   1050
               Width           =   1185
            End
            Begin VB.Label Label43 
               Alignment       =   1  'Right Justify
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
               Left            =   3240
               TabIndex        =   86
               Top             =   735
               Width           =   1185
            End
            Begin VB.Label Label42 
               Alignment       =   1  'Right Justify
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
               Left            =   3240
               TabIndex        =   85
               ToolTipText     =   "Doble Click para Abrir Detalle de Módulos"
               Top             =   210
               Width           =   1185
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Varios:"
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
               Left            =   1200
               TabIndex        =   84
               Top             =   1440
               Width           =   1905
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " Puestos de Trabajo:"
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
               Left            =   -105
               TabIndex        =   83
               Top             =   630
               Width           =   1485
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Módulos"
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
               Left            =   780
               TabIndex        =   82
               Top             =   300
               Width           =   2325
            End
            Begin VB.Label Label33 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Desarrollos Especiales:"
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
               Left            =   780
               TabIndex        =   81
               Top             =   1140
               Width           =   2325
            End
            Begin VB.Label Label52 
               Alignment       =   1  'Right Justify
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
               Left            =   4680
               TabIndex        =   80
               Top             =   1365
               Width           =   1185
            End
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "c) Cotización Bruta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   210
            TabIndex        =   148
            Top             =   4725
            Width           =   2445
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
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
            Left            =   4785
            TabIndex        =   147
            Top             =   5080
            Width           =   1185
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total General (Bruto sin IVA):"
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
            Left            =   1470
            TabIndex        =   146
            Top             =   5160
            Width           =   3165
         End
         Begin VB.Line Line6 
            X1              =   210
            X2              =   6090
            Y1              =   4515
            Y2              =   4515
         End
      End
      Begin VB.Frame Frame6 
         BorderStyle     =   0  'None
         Caption         =   "Frame6"
         Height          =   5800
         Left            =   -74800
         TabIndex        =   8
         Top             =   420
         Width           =   10320
         Begin VB.Frame Frame11 
            Caption         =   "Representante"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1590
            Left            =   7980
            TabIndex        =   11
            Top             =   4200
            Width           =   2325
            Begin VB.TextBox CUETEXT 
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
               Left            =   420
               TabIndex        =   157
               Text            =   " "
               Top             =   1155
               Width           =   1755
            End
            Begin VB.TextBox Text21 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1440
               TabIndex        =   4
               Text            =   " "
               Top             =   300
               Width           =   720
            End
            Begin VB.CommandButton BOTSEL 
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
               Index           =   5
               Left            =   675
               TabIndex        =   3
               Top             =   315
               Width           =   175
            End
            Begin VB.TextBox COTEXT 
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
               Index           =   5
               Left            =   90
               TabIndex        =   2
               Text            =   " "
               Top             =   315
               Width           =   630
            End
            Begin VB.Label Label18 
               BackStyle       =   0  'Transparent
               Caption         =   "Obs: "
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
               Left            =   105
               TabIndex        =   158
               Top             =   945
               Width           =   1710
            End
            Begin VB.Label DETEXT 
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
               Index           =   5
               Left            =   90
               TabIndex        =   13
               Top             =   630
               Width           =   2070
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "%C:"
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
               Left            =   -345
               TabIndex        =   12
               Top             =   375
               Width           =   1710
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Cronograma de Pagos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1590
            Left            =   3045
            TabIndex        =   35
            Top             =   4200
            Width           =   4800
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
               Height          =   900
               IntegralHeight  =   0   'False
               ItemData        =   "REPRFLEX04.frx":687C
               Left            =   120
               List            =   "REPRFLEX04.frx":687E
               TabIndex        =   47
               Top             =   600
               Width           =   4530
            End
            Begin VB.PictureBox Picture4 
               BackColor       =   &H00FFFFC0&
               Height          =   345
               Left            =   120
               ScaleHeight     =   285
               ScaleWidth      =   4470
               TabIndex        =   45
               Top             =   315
               Width           =   4530
               Begin VB.Label Label5 
                  BackColor       =   &H00FFFFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   "     It.     Descripcion                     Importe         Fecha"
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
                  Left            =   0
                  TabIndex        =   46
                  Top             =   20
                  Width           =   4950
               End
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Sistema Base"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1590
            Left            =   0
            TabIndex        =   34
            Top             =   4200
            Width           =   2955
            Begin VB.OptionButton Option4 
               Caption         =   "Auto- Integrado"
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
               Left            =   315
               TabIndex        =   154
               Top             =   1150
               Width           =   2535
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Power-Up Gestión"
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
               Left            =   315
               TabIndex        =   153
               Top             =   900
               Width           =   2535
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Power-Up Industrial"
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
               Left            =   315
               TabIndex        =   152
               Top             =   650
               Width           =   2535
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Sistema Integrado"
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
               Left            =   315
               TabIndex        =   151
               Top             =   400
               Width           =   2535
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Comercial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   10305
            Begin VB.CommandButton Command6 
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
               Left            =   9840
               Picture         =   "REPRFLEX04.frx":6880
               TabIndex        =   42
               Top             =   420
               Width           =   175
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
               Height          =   285
               Left            =   4800
               Picture         =   "REPRFLEX04.frx":6B8A
               TabIndex        =   41
               Top             =   420
               Width           =   175
            End
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
               Height          =   285
               Left            =   6090
               TabIndex        =   31
               Top             =   420
               Width           =   2040
            End
            Begin VB.TextBox Text12 
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
               Left            =   8820
               TabIndex        =   30
               Text            =   " "
               Top             =   420
               Width           =   1020
            End
            Begin VB.TextBox Text24 
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
               Left            =   3780
               TabIndex        =   1
               Text            =   " "
               Top             =   420
               Width           =   1020
            End
            Begin VB.TextBox Text25 
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
               Left            =   1155
               TabIndex        =   0
               Top             =   420
               Width           =   1935
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Orden de Compra:"
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
               Left            =   5025
               TabIndex        =   33
               Top             =   315
               Width           =   1065
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFF80&
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha"
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
               Left            =   8190
               TabIndex        =   32
               Top             =   465
               Width           =   585
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFF80&
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha"
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
               Left            =   3150
               TabIndex        =   29
               Top             =   465
               Width           =   585
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Pedido de Cotización:"
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
               Left            =   -315
               TabIndex        =   10
               Top             =   315
               Width           =   1380
            End
         End
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Height          =   3795
            Left            =   -945
            TabIndex        =   36
            Top             =   840
            Width           =   11670
            Begin VB.Frame Frame17 
               Caption         =   "Licencias"
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
               Left            =   9660
               TabIndex        =   184
               Top             =   2565
               Width           =   1485
               Begin VB.Label Label63 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " Puestos:"
                  BeginProperty Font 
                     Name            =   "Arial Narrow"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Left            =   -630
                  TabIndex        =   186
                  Top             =   350
                  Width           =   1485
               End
               Begin VB.Label Label62 
                  Alignment       =   1  'Right Justify
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
                  Left            =   945
                  TabIndex        =   185
                  ToolTipText     =   "Doble Click para Abrir Cotización de Accesorios"
                  Top             =   310
                  Width           =   450
               End
            End
            Begin VB.TextBox Text40 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   600
               Left            =   1890
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   182
               Top             =   2670
               Width           =   7680
            End
            Begin VB.PictureBox Picture6 
               BackColor       =   &H00E0E0E0&
               Height          =   2010
               Left            =   9555
               ScaleHeight     =   1950
               ScaleWidth      =   1560
               TabIndex        =   71
               Top             =   525
               Width           =   1620
               Begin VB.TextBox Text38 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   -30
                  TabIndex        =   73
                  Top             =   -30
                  Width           =   1650
               End
               Begin VB.Label Label90 
                  Height          =   435
                  Left            =   315
                  TabIndex        =   97
                  Top             =   945
                  Visible         =   0   'False
                  Width           =   855
               End
               Begin VB.Label Label2 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "Pesos"
                  BeginProperty Font 
                     Name            =   "Arial Black"
                     Size            =   12
                     Charset         =   0
                     Weight          =   900
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   375
                  Left            =   0
                  TabIndex        =   72
                  Top             =   420
                  Width           =   1575
               End
            End
            Begin VB.PictureBox Picture3 
               BackColor       =   &H00E0E0E0&
               Height          =   2010
               Left            =   1020
               ScaleHeight     =   1950
               ScaleWidth      =   795
               TabIndex        =   69
               Top             =   525
               Width           =   855
               Begin VB.TextBox Text36 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "Courier New"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   360
                  Left            =   -30
                  TabIndex        =   70
                  Top             =   -30
                  Width           =   850
               End
            End
            Begin VB.TextBox Text37 
               Alignment       =   2  'Center
               BorderStyle     =   0  'None
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
               Left            =   0
               TabIndex        =   48
               Text            =   "Text36"
               Top             =   0
               Width           =   825
            End
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00FFFFC0&
               Height          =   330
               Left            =   1020
               ScaleHeight     =   270
               ScaleWidth      =   10095
               TabIndex        =   43
               Top             =   210
               Width           =   10155
               Begin VB.Line Line24 
                  X1              =   0
                  X2              =   10080
                  Y1              =   315
                  Y2              =   315
               End
               Begin VB.Line Line23 
                  X1              =   8505
                  X2              =   8505
                  Y1              =   0
                  Y2              =   315
               End
               Begin VB.Line Line2 
                  X1              =   840
                  X2              =   840
                  Y1              =   315
                  Y2              =   0
               End
               Begin VB.Label Label19 
                  BackColor       =   &H00FFFFC0&
                  BackStyle       =   0  'Transparent
                  Caption         =   $"REPRFLEX04.frx":6E94
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
                  Left            =   50
                  TabIndex        =   44
                  Top             =   30
                  Width           =   10200
               End
               Begin VB.Line Line3 
                  X1              =   860
                  X2              =   860
                  Y1              =   0
                  Y2              =   315
               End
               Begin VB.Line Line22 
                  X1              =   8505
                  X2              =   8505
                  Y1              =   0
                  Y2              =   315
               End
            End
            Begin VB.TextBox Text17 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2010
               Left            =   1890
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   40
               Top             =   525
               Width           =   7680
            End
            Begin VB.Label Label61 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFF80&
               BackStyle       =   0  'Transparent
               Caption         =   "Versión:"
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
               Left            =   735
               TabIndex        =   183
               Top             =   2720
               Width           =   1110
            End
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
         Height          =   120
         Left            =   -71205
         TabIndex        =   7
         Top             =   4605
         Width           =   15
      End
      Begin VB.Label Label17 
         Caption         =   "Label8"
         Height          =   120
         Left            =   -70785
         TabIndex        =   6
         Top             =   3660
         Width           =   15
      End
   End
   Begin VB.Label Label37 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "At Sr:"
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
      Left            =   4200
      TabIndex        =   156
      Top             =   1260
      Width           =   900
   End
   Begin VB.Label Label60 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mano de Obra:"
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
      Left            =   8190
      TabIndex        =   99
      Top             =   4830
      Width           =   1905
   End
   Begin VB.Label Label50 
      Caption         =   "Label50"
      Height          =   495
      Left            =   5400
      TabIndex        =   74
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. de Presupuesto:"
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
      Left            =   0
      TabIndex        =   68
      Top             =   110
      Width           =   1140
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
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
      Left            =   -210
      TabIndex        =   67
      Top             =   630
      Width           =   1350
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Domicilio:"
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
      Left            =   4020
      TabIndex        =   66
      Top             =   630
      Width           =   1110
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Localidad:"
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
      Left            =   4200
      TabIndex        =   65
      Top             =   945
      Width           =   900
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
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
      Height          =   330
      Left            =   8925
      TabIndex        =   64
      Top             =   315
      Width           =   690
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Validez:"
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
      Left            =   8805
      TabIndex        =   63
      Top             =   630
      Width           =   810
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Responsable:"
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
      Left            =   8190
      TabIndex        =   62
      Top             =   945
      Width           =   1425
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente:"
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
      Left            =   3990
      TabIndex        =   61
      Top             =   315
      Width           =   1110
   End
   Begin VB.Label Label22 
      BackColor       =   &H00E0E0E0&
      Caption         =   " "
      Height          =   330
      Left            =   3360
      TabIndex        =   60
      Top             =   105
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Line Line1 
      X1              =   -420
      X2              =   11580
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu Archiv 
      Caption         =   "Archivos"
      Begin VB.Menu NueOT 
         Caption         =   "Nuevo Presupuesto"
      End
      Begin VB.Menu EditCli 
         Caption         =   "Abrir Existente"
      End
      Begin VB.Menu plh1 
         Caption         =   "-"
      End
      Begin VB.Menu GuaCli 
         Caption         =   "Guardar"
      End
      Begin VB.Menu php1a 
         Caption         =   "-"
      End
      Begin VB.Menu Print 
         Caption         =   "Imprimir Ficha"
      End
      Begin VB.Menu plh3 
         Caption         =   "-"
      End
      Begin VB.Menu List 
         Caption         =   "Consultas y Listados"
      End
      Begin VB.Menu plh4 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu EdiCli 
      Caption         =   "Edición"
      Begin VB.Menu CoInEv 
         Caption         =   "Copia Interactiva de Presupuesto"
      End
   End
   Begin VB.Menu ConLiCli 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu LiGeClaCli 
         Caption         =   "Listado General y Clasificado"
      End
   End
   Begin VB.Menu ConfCli 
      Caption         =   "Configuración"
      Begin VB.Menu ConGenCli 
         Caption         =   "Configuración General"
      End
   End
   Begin VB.Menu AyuCli 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "REPRFLEX04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHEX$(16)
Dim YACAR%, MODIFIC%
Dim GRABATODO%
Dim AGRECLI%
Dim RECONTANT$
Dim CAJENTRA%
Dim COMPACTO%
Dim VALORA
'
Private Sub AyuCli_Click()
  Call Command4_Click
End Sub

Private Sub Baj_Cli_Click()
End Sub

Private Sub BOTCON_Click()
   '
   NEV1$ = TRIM(Text1)
   If NEV1$ = "" Then
     On Error Resume Next
     Text1.SetFocus
     On Error GoTo 0
     Exit Sub
   End If
   '
   If ULTREG&("ECUECON") < 1 Then
      Call MENSERR(24, "Falta Definir Plan de Cuentas Contables")
      Call CARCUEINGAS
      Exit Sub
   End If
   '
   RG1$ = REGSEL$("ECUECON")
   If Len(RG1$) > 4 Then
      CUETEXT.TEXT = Left$(RG1$, 12)
   End If
   '
End Sub
Private Sub CoInBDCli_Click()
Call Command99_Click
End Sub

Private Sub Check1_Click()
   MODIFIC% = 1
   If Check1.Value = 1 Then
        Label65(0) = ""
        Label65(1) = ""
        Label65(2) = ""
      Else
        Call Label52_CHANGE
   End If
End Sub

Private Sub CoInEv_Click()
  Call Command22_Click
End Sub

Private Sub CoInGr_Click()
    If SSTab1.Tab = 2 Then
        Call Command22_Click
    Else
        Call COMUNI("Para realizar esta Copia debe Situarse en la Solapa de Grupos.")
    End If
End Sub

Private Sub Command10_Click()
    Call CARGOADJUN
End Sub

Private Sub Command11_Click()
    '
    Command11.Enabled = False
    '
3   Call SELECHO("IPRESCAM")
    NUPRESUP$ = TRIM$(VALACT1$("IPRESCAM"))
    If NUPRESUP$ = "" Then GoTo 99
    '
    TIDOCUX$ = "Presupuesto " + NUPRESUP$ + " "
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUX$) > 0 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 5
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    GoTo 99
    '
5   NUPIC% = EDITFORM.Picture1.Count
    For KKI% = 0 To NUPIC% - 1
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
    NUPIC% = EDITFORM.Picture1.Count
    For KKI% = 0 To NUPIC% - 1
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
      ' If InStr(XX$, NUDOCU$) > 4 Then
      If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    NUPIC% = EDITFORM.Picture1.Count
    For KKI% = 0 To NUPIC% - 1
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
99  Command11.Enabled = True
    '
End Sub

Private Sub Command12_Click()
    Command12.Enabled = False
    NUPRESU$ = TRIM$(Text1)
    Call ABREPRECAM
    With Docadjun
20    .FindFirst "Nro_PRESU = '" & NUPRESU$ & "'"
      If .NoMatch = False Then
        .Delete
        GoTo 20
      End If
      '
      For KKIL& = 1 To List2.ListCount
        .AddNew
        !Nro_PRESU$ = NUPRESU$
        !RutArch = List2.List(KKIL& - 1)
        .Update
      Next KKIL&
    End With
    Command12.Enabled = True
End Sub

Private Sub Command16_Click()
  If TRIM(Text1) <> "" Then
     Call SELECHO("SELFECHA")
     VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
     If VDEVU$ <> "" Then
        Text10.TEXT = VALACT1$("SELFECHA")
     End If
     FF% = FECHANUM(Text10.TEXT)
  End If
End Sub


Private Sub Command22_Click()
  '\\\OT-05-0391-WAR-09/06/05///
  Command22.Enabled = False
  If DERACCE%("ANUPRECO", "Generar Nuevo Presupuesto por Copia") < 2 Then
    GoTo 99
  End If
  '
  SSTab1.Tab = 0
9 Screen.MousePointer = 1
  '
  VDEF$ = Space$(64) + String$(6, 0)
  '
  OBX$ = OBJPLA$("DEACOPRE", VDEF$)
  If OBX$ = "" Then GoTo 99
  VDEF$ = OBX$
  '
  PRESU1$ = TRIM$(Mid$(OBX$, 1, 16))
  PRESU2$ = TRIM$(Mid$(OBX$, 17, 48))
  '//////////////VALIDACIONES////////////
  If TRIM$(PRESU1$) = "" Then
    Call COMUNI("Número de Presupuesto no valido")
    GoTo 9
  End If
  '
  If TRIM$(PRESU2$) = "" Then
    Call COMUNI("Número de Presupuesto no valido.")
    GoTo 9
  End If
  '
  If ECOARCH("IPRESCAM", PRESU1$) = "" Then
    Call MENSERR(24, "Presupuesto no valido.")
    GoTo 9
  End If
  '
  If ECOARCH("IPRESCAM", PRESU2$) = "" Then
    Call MENSERR(24, "Presupuesto no valido.")
    GoTo 9
  End If
  '
  If TRIM$(PRESU1$) = TRIM$(PRESU2$) Then
    Call MENSERR(24, "Copia Imposible.\Coinciden Presupuesto Origen y Destino.")
    GoTo 99
  End If
  'validación de que en el presupuesto no haya detalle.
  Call ABREPRECAM
  DetaPresu.FindFirst " Nro_Presu= '" & PRESU2$ & "' "
  If DetaPresu.NoMatch = True Then
    GoTo 17
  Else
    If COMALTER%("El Presupuesto de Cámara " & PRESU2$ & " \ya Contiene Detalle, ¿Desea Sobre-Escribirlo?\\Si,Sobre-escribir.\\No, Cancelar.") = 2 Then
      On Error GoTo 0
      Screen.MousePointer = 1
      Command22.Enabled = True
      Exit Sub
    Else
      DetaPresu.FindFirst " Nro_Presu= '" & PRESU2$ & "' "
      If DetaPresu.NoMatch = False Then
        DetaPresu.Delete
      End If
14    DetaPresu.FindNext " Nro_Presu= '" & PRESU2$ & "' "
      If DetaPresu.NoMatch = True Then
        GoTo 17
      Else
        DetaPresu.Delete
        GoTo 14
      End If
    End If
  End If
    '
17  Call ABREPRECAM
    EncaPresu.FindFirst "Nro_Presu= '" & PRESU1$ & "' "
    With EncaPresu
     USUA$ = !uSUCONF_Presu
     USUC$ = !UsuCont_Presu
     USUAP$ = !UsuApro_Presu
     NOC$ = !NroOcom_Presu
     FPCLIP$ = !FeOcom_Presu
     NUPECOT$ = !PediCoti_Presu
     FEPECOT = !FePedCot_Presu
     CAOPE# = !CanPed_Presu
     ESPECI$ = !Especi_Presu
     MONEOTR% = !Mone_Presu
     COMOTR# = !CotiMone_Presu
     LUENOTR$ = !LENTREGA_Presu
     VEOTR% = !Vended_Presu
     COVEOTR$ = !ComiVend_Presu
     On Error Resume Next
      OBSOTR$ = !Observa_Presu
     On Error GoTo 0
     PREVEOTR# = !PreVeN_PRESU
     CUOTMAN = !Largo_Presu
     ANCHOPRE = !Ancho_Presu
     ALTOPRE = !Alto_Presu
     ESPEPRE = !Espesor_Presu
     On Error Resume Next
      PORGGMAT# = !GGMat_Presu '% MATERIALES
     On Error GoTo 0
     On Error Resume Next
      PORGGOB# = !GGOba_Presu '% MANO DE OBRA
     On Error GoTo 0
     On Error Resume Next
      PORDEMO# = !DesBon_Presu '% DESCUENTOS Y MODIFICACIONES
     On Error GoTo 0
    End With
    '
    Call ABREPRECAM
    EncaPresu.FindFirst "Nro_Presu= '" & PRESU2$ & "' "
    If EncaPresu.NoMatch = False Then
      With EncaPresu
       .Edit
         !uSUCONF_Presu = USUA$ 'usuario de confección
         !UsuCont_Presu = USUAC$ 'usuario que controlo
         !UsuApro_Presu = USUAP$ 'usuario que aprobara
         !NroOcom_Presu = NOC$ 'numero de orden de compra del cliente
         !FeOcom_Presu = FPCLIP$  'fecha de pedido del cliente
         !PediCoti_Presu = NUPECOT$ 'Numero de Oferta
         !FePedCot_Presu = FEPECOT 'fecha de la oferta
         !NuOrItem_Presu = 0
         !CanPed_Presu = CAOPE#
         !Especi_Presu = ESPECI$
         !PreVeN_PRESU = PREVEOTR#
         !Mone_Presu = MONEOTR%
         !CotiMone_Presu = COMOTR#
         !LENTREGA_Presu = LUENOTR$
         !Ajusta_Presu = 0
         !Vended_Presu = VEOTR%
         !ComiVend_Presu = COVEOTR$
         !Observa_Presu = OBSOTR$
         !PreVeN_PRESU = PREVEOTR#
         !Largo_Presu = CUOTMAN
         !Ancho_Presu = ANCHOPRE
         !Alto_Presu = ALTOPRE
         !Espesor_Presu = ESPEPRE
         !GGMat_Presu = PORGGMAT# '% MATERIALES
         !GGOba_Presu = PORGGOB# '% MANO DE OBRA
         !DesBon_Presu = PORDEMO# '% DESCUENTOS Y MODIFICAC
       .Update
      End With
    End If
    'ahora copio la lista de grupos
    SQLX$ = " SELECT * FROM detapresu "
    SQLX$ = SQLX$ + " WHERE Nro_Presu= '" & PRESU1$ & "' "
    Set DETAPRESUTMP = PresuCam.OpenRecordset(SQLX$, 4)
    '
    With DETAPRESUTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       GoTo 36
     Else
       Do While .EOF = False
        NURUB% = !rubro_det
        SUBRUB% = !SUBRUB_DET
        CODET% = !CODINTE_DET
        CODEXTET$ = !CODEXTE_DET
        CANMAT# = !Cantid_Det
        PREUMAT# = !PreUni_Det
        MONEDET% = !MONE_DET
        MED1 = !MED1_DET
        MED2 = !Med2_Det
        MED3 = !Med3_Det
        IMPODEP# = !Impo_Det
        DESCRIDET$ = !DESCRI_DET
        HORADET = !Hora_Det
        OPERADET% = !Opera_Det
        VALOHORDET# = !ValoHor_Det
        OBSGRU$ = !OBSERVA_GRUP
        With DetaPresu
          .AddNew
            !Nro_PRESU = PRESU2$
            !rubro_det = NURUB%
            !SUBRUB_DET = SUBRUB%
            !CODINTE_DET = CODET%
            !CODEXTE_DET = CODEXTET$
            !Cantid_Det = CANMAT#
            !PreUni_Det = PREUMAT#
            !MONE_DET = MONEDET%
            !MED1_DET = MED1
            !Med2_Det = MED2
            !Med3_Det = MED3
            !Impo_Det = IMPODEP#
            !DESCRI_DET = DESCRIDET$
            !Hora_Det = HORADET
            !Opera_Det = OPERADET%
            !ValoHor_Det = VALOHORDET#
            !OBSERVA_GRUP = OBSGRU$
          .Update
        End With
        .MoveNext
       Loop
       End If
    End With
    '
36  Call GENEINDIPRECA
    Text1 = TRIM$(PRESU2$)
    GoTo 99
    '
99  Screen.MousePointer = 1
    Command22.Enabled = True
    '\\\OT-05-0391-WAR-FIN///
End Sub

Private Sub Command22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
    If SSTab1.Tab = 1 Or SSTab1.Tab = 0 Then
        Command22.ToolTipText = "Alta de Nueva Presupuesto de Cámara por Copia de una Existente"
    End If
    If SSTab1.Tab = 2 Then
        Command22.ToolTipText = "Alta de Nueva Lista de Grupos a partir de una Lista Existente"
    End If
    '
End Sub

Private Sub Command4_Click()
   Call HELPONLINE("AMACLI")
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   NUPRECA$ = TRIM$(Text1)
   If NUPRECA$ = "" Or ECOARCH("IPRESCAM", Text1) = "" Then GoTo 99
   '
   TTXX$ = "Confirma Anulación Definitiva\del Presupuesto '" + NUPRECA$ + "' ?\\No, Conservarlo\Si, Anular"
   If COMALTER%(TTXX$) = 2 Then
     Call ABREPRECAM
     With EncaPresu
       .FindFirst "Nro_Presu = '" & NUPRECA$ & "' "
10     If .NoMatch = False Then
         .Delete
         .FindNext "Nro_Presu = '" & NUPRECA$ & "' "
         GoTo 10
       End If
     End With
     '
     With DetaPresu
       .FindFirst "Nro_Presu = '" & NUPRECA$ & "' "
20     If .NoMatch = False Then
         .Delete
         .FindNext "Nro_Presu = '" & NUPRECA$ & "' "
         GoTo 20
       End If
     End With
     '
     With Despecial
       .FindFirst "Nro_Presu = '" & NUPRECA$ & "' "
30     If .NoMatch = False Then
         .Delete
         .FindNext "Nro_Presu = '" & NUPRECA$ & "' "
         GoTo 30
       End If
     End With
     '
     Call GENEINDIPRECA
     Call FRESHECHO("IPRESCAM")
     Text1 = ""
     '
     Call COMUNI("Se ha Borrado Definitivamente\el Presupuesto '" + NUPRECA$ + "'")
   End If
   '
99 Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   '
   Command8.Enabled = False
   Call CIERRARCH("*.*")
   OPX% = ALTERNA%("Maestro de Clientes\Maestro de Items de Stock")
   If OPX% = 1 Then
        ACONEC$ = "ARCHIG04/AMACLI"
        Call CONECRUN(ACONEC$, "Padron Maestro de Clientes")
     ElseIf OPX% = 2 Then
        ACONEC$ = "AMASTO04"
        Call CONECRUN(ACONEC$, "Maestro de Articulos de Stock")
   End If
99 Command8.Enabled = True
   '
End Sub

Private Sub Command99_Click()
   Command99.Enabled = False
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("DEUDOR1")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     R0$ = REGLEIDO$("DEUDOR1", U&)
     NCLIE% = CVI(Left$(R0$, 2))
     R1$ = Mid$(R0$, 3, 30)
     If Left$(R1$, 1) = Chr$(64) Then
       R1$ = AJUSTI$(TRIM$(Mid$(R1$, 2)), 30)
     End If
     R1$ = REPLACAR$(R1$, "@", " ")
     R2$ = AJUSTI$(RASOCLI$(NCLIE%), 30)
     If R1$ <> R2$ Then
       HAYINCO% = HAYINCO% + 1
       Call MENSERR(24, "Revisar Cuenta " + TRIM$(Str$(NCLIE%)))
     End If
   Next U&
   '
   If HAYINCO% < 1 Then
     Call COMUNI("Control Completo.\No se Detectaron Errores.")
   End If
   '
   Screen.MousePointer = 1
   Command99.Enabled = True
End Sub

Private Sub ConGenCli_Click()
    Call Command14_Click
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

Private Sub DeHaCaCli_Click()
Call Command2_Click
End Sub

Private Sub DeLiCreCli_Click()
    Call AMACLIS04.Command3_Click
End Sub

Private Sub COTEXT_LostFocus(Index As Integer)
  '
  If Val(COTEXT(5)) <= 0 Then
    Text21 = ""
  End If
  '
End Sub

Private Sub CUETEXT_Change()
     MODIFIC% = 1
End Sub

Private Sub Dir1_Change()
   Label86.Caption = Dir1.Path
   File1.Path = Dir1.Path
   Label85 = ""
   Call Text41_Change
End Sub

Private Sub Drive1_Change()
   On Error GoTo 20
   Dir1.Path = Drive1.Drive
   GoTo 99
   '
20 Resume 21
21 On Error GoTo 0
   Call MENSERR(24, "Unidad no Disponible.")
   Drive1.Drive = Left$(CurDir$, 2)
   '
99 On Error GoTo 0
End Sub

Private Sub EditCli_Click()
   Call Command9_Click
End Sub

Private Sub Exit_Click()
   Call Command3_Click
End Sub

Private Sub Form_Initialize()
  '
  SSTab1.Tab = 0
  '
End Sub

'
Private Sub Form_Load()
    '
    SSTab1.Tab = 0
    Call CENTRAFORM(Me)
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FUERASYS
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Call GRATITFOR(Caption)
    '
    On Error Resume Next
    Kill LUCOM$ + "GRUORTRA.WKF"
    On Error GoTo 0
    '
    COLOBAK = RGB(230, 230, 230)
    TICAMBIS.BackColor = COLOBAK
    '
    Call LEETASIVA
    CONVERPRECIO = 1
    '
    CLICEN04.Text21 = "1"
    CMIGRA04.Text21 = "2"
    VALORA = 105
    '
    On Error Resume Next
    Drive1.Drive = "K:"
    If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Imposible Acceder a 'K:/DOCU_COM/ESPECIFICACIONES/FORMATO PDF'\   \Unidad 'K.' no Conectada como Unidad de Red.\No se Podrán Adjuntar Especificaciones de Sistemas.")
      Call Option8_Click
      Exit Sub
    End If
    Dir1.Path = "k:\DOCU_COM\ESPECIFICACIONES\FORMATO PDF"
    '
End Sub
'
Private Sub Form_Activate()
   WindowState = 0
   OPENFORMS = DoEvents
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
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub BOTSEL_Click(Index As Integer)
   If TRIM(Text1) <> "" Then
    Call SELECHO("VENDEDOR")
    VDEVU$ = VALACT1$("VENDEDOR")
    If TRIM$(VDEVU$) <> "" Then
         COTEXT(5).TEXT = VDEVU$
    End If
   Else
    Call COMUNI("Debe Seleccionar un Presupuesto de Cámara.")
   End If
End Sub

Private Sub Command1_Click()
   GRABATODO% = 1
   'Call GRAGENCLI
   'Call GRANOTACLI
   GRABATODO% = 0
End Sub

 Sub Command13_Click()
   Command13.Enabled = False
      Call COMUNI("No Disponible por el Momento")
      'LIPRESU04.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   OPPRFLE04.Show 1
   Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
    If TRIM$(Text1.TEXT) <> "" Then
        Call GRABOT
    Else
        Call COMUNI("Debe Seleccionar un Presupuesto.")
    End If
End Sub
 
 Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command2_Click()
   'Call GRAGENCLI
   'Call GRANOTACLI
   'For Index = 0 To 8: YACAR%(Index) = 0: Next Index
   'Call CARDATCLIE
End Sub

'Private Sub Command20_Click()
'   Static ARGU$
'   Command20.Enabled = False
'10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda Cliente por Texto", ARGU$)))
'   If ARGU$ <> "" Then
'     Screen.MousePointer = 11
'     JJ& = 0
'     Call COPYSTRU("DEUDOR1", "SELECLI")
'     '
'     Call ABRECLIEN
'     On Error Resume Next
'     With Deudor12
'       .MoveFirst
'       If Err < 1 Then
'         Do While Not .EOF
'           For KKI% = 1 To .Fields.Count
'             VATEX$ = .Fields(KKI% - 1).Value
'             If InStr(VATEX$, ARGU$) > 0 Then
'               NCLIE% = !Nume_Cli
'               RACLI$ = !Raso_Cli
'               XX$ = MKI$(NCLIE%) + RACLI$
'               JJ& = JJ& + 1
'               Call GRAREG("!SELECLI", XX$, JJ&)
'               Exit For
'             End If
'           Next KKI%
'         .MoveNext
'         Loop
'       End If
'     End With
'     Call SETULTREG("!SELECLI", JJ&)
'     '
'     Screen.MousePointer = 1
'     If JJ& < 1 Then
'         Call MENSERR(24, "No se Encontraron Coincidencias")
'         GoTo 10
'       ElseIf JJ& = 1 Then
'         Text1 = TRIM$(Str$(CVI(Left$(REGLEIDO$("!SELECLI", 1), 2))))
'       Else
'         Call FRESHECHO("!SELECLI")
'         Call SELECHO("!SELECLI")
'         NCLIX$ = VALACT1$("!SELECLI")
'         If NCLIX$ <> "" Then
'           Text1 = NCLIX$
'         End If
'     End If
'   End If
'   Command20.Enabled = True
'End Sub

Private Sub Command21_Click()
   '\\\OT-05-0330-WAR-26/05/05///
   If DERACCE%("ALTANUPR", "Alta de Nuevo Presupuesto FLEXOFT") < 2 Then
      Exit Sub
   End If
   '
   Screen.MousePointer = 1
   '
   'FALTA PONER EL CODIGO PARA QUE TRAIGA LAS FECHAS POR DEFAULT
   '
   HOII% = HOY(HOS$)
   NPX$ = Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Mid$(HOS$, 1, 2)
   Call ABREPRECAM
   SQLX$ = "SELECT NRO_PRESU FROM ENCAPRESU "
   SQLX$ = SQLX$ + "WHERE NRO_PRESU LIKE 'PR-" + NPX$ + "-*' "
   SQLX$ = SQLX$ + "ORDER BY VAL(MID(NRO_PRESU,11,4)) DESC;"
   Set EncaTemp = PresuCam.OpenRecordset(SQLX$, dbOpenSnapshot)
   NROMA& = 0
   On Error Resume Next
   With EncaTemp
     .MoveFirst
     If Err < 1 Then
       NROMA& = Val(Mid$(!Nro_PRESU, 11))
     End If
     On Error GoTo 0
   End With
   '
   NPRESX$ = "PR-" + NPX$ + "-" + TRIM$(Str$(NROMA& + 1))
   VDEF$ = Space$(64) + String$(6, 0)
   Call REPLA(VDEF$, NPRESX$, 1, 16)
   HOII% = HOY(HOS$)
   Call REPLA(VDEF$, MKI$(HOII%), 67, 2)
   FEX0 = HOII%: FEX = FEX0
   While DIENTRE(FEX0, FEX) < 7: FEX = DIPOST(FEX): Wend
   FVEN% = FEX
   Call REPLA(VDEF$, MKI$(FVEN%), 69, 2)
   '
9  OBj$ = OBJPLA$("NUEPREC", VDEF$)
   If OBj$ = "" Then GoTo 99
   VDEF$ = OBj$
   '
   NUPRE$ = TRIM$(Mid$(OBj$, 1, 16))
   DESCRIPRE$ = TRIM$(Mid$(OBj$, 17, 48))
   NUCLIPRE% = CVI(Mid$(OBj$, 65, 2))
   RASOCLIPRE$ = RASOCLI$(NUCLIOTR%)
   FECPRE% = CVI(Mid$(OBj$, 67, 2))
   FEVAPRE% = CVI(Mid$(OBj$, 69, 2))
   
   '//////////////VALIDACIONES////////////
   If TRIM$(NUPRE$) = "" Then
      Call COMUNI("Número de Presupuesto no válido.")
      GoTo 9
   End If
   '
   Call ABREPRECAM
   EncaPresu.FindFirst "Nro_Presu = '" & NUPRE$ & "' "
   If EncaPresu.NoMatch = False Then
      Call COMUNI("Numero de Presupuesto no Válido o Pre-Existente")
      GoTo 9
   End If
   '
   If TRIM$(DESCRIPRE$) = "" Then
      Call COMUNI("Debe Ingresar una Descripcion para el Presupuesto")
      GoTo 9
   End If
   '
   If NUCLIPRE% <= 0 Then
      Call COMUNI("Cliente no Válido.")
      GoTo 9
   End If
   '
   If FEVAPRE% < FECPRE% Then
      Call COMUNI("La Fecha de Validez Debe ser Superior o Igual a la Fecha de Generación")
      GoTo 9
   End If
   '
   jjj = CVINT(Date)
   If FECPRE% > jjj Then
      Call COMUNI("La Fecha de Generación no puede ser Superior al dia de la Fecha.")
      GoTo 9
   End If
   '
   Call ABREPRECAM
   With EncaPresu
      .AddNew
        !Nro_PRESU = NUPRE$
        !REFE_Presu = DESCRIPRE$
        !Nclie_Presu = NUCLIPRE%
        !Femi_Presu = CVDAT(FECPRE%)
        !FVali_Presu = CVDAT(FEVAPRE%)
        !uSUCONF_Presu = USERNAME 'usuario de confección
        !UsuCont_Presu = "" 'usuario que controlo
        !UsuApro_Presu = "" 'usuario que aprobara
        !NroOcom_Presu = "" 'numero de orden de compra del cliente
        !FeOcom_Presu = CVDAT(Int(FECHANUM("01/01/80"))) 'fecha de pedido del cliente
        !PediCoti_Presu = "" 'Numero de Oferta
        !FePedCot_Presu = CVDAT(Int(FECHANUM("01/01/80"))) 'fecha de la oferta
        !NuOrItem_Presu = 0
        !CanPed_Presu = 0
        !Especi_Presu = ""
        !PreVeN_PRESU = 0
        !Mone_Presu = 1     ' MONEDA = PESOS
        !CotiMone_Presu = 1
        !LENTREGA_Presu = CONTACTCLI$(NUCLIPRE%)
        !Ajusta_Presu = 0
        !Vended_Presu = 0
        !ComiVend_Presu = 0
        !Observa_Presu = " "
        !Plazo_Entre = 7
        !Anti_Presu = 25
        !Cuot_Presu = 1
      .Update
   End With
   '
   Call Option1_Click
   '
   Call GENEINDIPRECA
   Text1 = TRIM$(NUPRE$)
   '
99 Screen.MousePointer = 1
   Command21.Enabled = True
   '\\\OT-05-0330-WAR-FIN///
End Sub

Private Sub Command25_Click()
   '
   If TRIM$(Text1) = "" Then
      Call COMUNI("Número de Presupuesto no Válido.")
      Exit Sub
   End If
   '
   If ECOARCH$("IPRESCAM", Text1) = "" Then
      Call COMUNI("Número de Presupuesto no Válido.")
      Exit Sub
   End If
   '
   If Text1 <> "" Then
        NCIX% = Val(Label22)
        EMAI$ = TRIM$(EMAILCLI$(NCIX%))
        If EMAI$ = "" Or InStr(EMAI$, "@") < 1 Then
           GoTo 20 'NO TIENE DIRECCION O DIRECCION NO VALIDA
        End If
        RESP22% = COMALTER%("¿ Formato de Salida ?\\Por E-Mail\Impresora")
        If RESP22% <> 1 Then GoTo 20
        '
15      SENDBYMAIL$ = ""
        pormail% = 1
        SENDBYMAIL$ = EMAI$
        '
        For Each prx1 In Printers
          If InStr(UCase$(prx1.DeviceName), "PDFWRITER") > 0 Then
            For U& = 0 To 256
              MAI_ADJUN$(U&) = ""
            Next U&
            For KU& = 1 To List2.ListCount
              MAI_ADJUN$(KU&) = List2.List(KU& - 1)
            Next KU&
            GoTo 20
          End If
        Next
        Call MENSERR(24, "Imposible Enviar por Mail.\   \Falta Impresora Virtual PDFWRITER.")
        SENDBYMAIL = ""
        '
20      Call PRIDETPRE ' IMPRESION COMPLETA DEL PRESUPUESTO
        pormail% = 0
        SENDBYMAIL$ = ""
     Else
        Call COMUNI("Debe Seleccionar un Presupuesto de Cámara.")
   End If
End Sub

Private Sub Command3_Click()
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
   '
End Sub

Private Sub Command5_Click()
 If TRIM$(Text1) <> "" Then
  Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text24.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text24.TEXT)
 End If
End Sub

Private Sub Command6_Click()
 If TRIM$(Text1) <> "" Then
  Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text12.TEXT = VALACT1$("SELFECHA")
   End If
   FF% = FECHANUM(Text12.TEXT)
 End If
End Sub

Private Sub Command9_Click()
   '
  ' Call GRABOT
   Call SELECHO("IPRESCAM")
   NUPRECAM$ = TRIM$(VALACT1$("IPRESCAM"))
   If NUPRECAM$ <> "" Then
     Text1.TEXT = TRIM$(NUPRECAM$)
   End If
   '
End Sub

Private Sub LiCuInCli_Click()
    Call AMACLIS04.Command5_Click
End Sub



Private Sub Frame3_DBLCLICK()
    Call List1_DBLClick
End Sub

Private Sub Label13_DblClick()
   Call Label66_DblClick
End Sub

Private Sub Label2_dblClick()
   '
Exit Sub
   Static CAMON%, DMONE$(), TCMONE()
   '
   If TRIM$(Text1) <> "" Then
     If MODIFIC% > 0 Then
       Call MENSERR(24, "Presupuesto Modificado.\Imposible Cambiar Moneda")
       Exit Sub
     End If
   End If
   '
   If CAMON% < 1 Then
     CAMON% = ULTREG&("TAMONED")
     If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
          VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
          DMONE$(U&) = DMX$
          TCMONE(U&) = VHX
         Next U&
     End If
   End If
   '
   MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   Label2 = DMONE$(MONEMI%)
   Label2.ForeColor = RGB(0, 0, 255)
   If MONEMI% > 1 Then
      Label2.ForeColor = RGB(255, 0, 0)
   End If
   Label2.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   '
   PRESU2$ = Text1
   Call ABREPRECAM
   EncaPresu.FindFirst " Nro_Presu= '" & PRESU2$ & "' "
   If EncaPresu.NoMatch = False Then
     EncaPresu.Edit
     MONEMIA% = EncaPresu!Mone_Presu
     PRUNI# = EncaPresu!PreVeN_PRESU * TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
     EncaPresu!PreVeN_PRESU = PRUNI#
     EncaPresu!Mone_Presu = MONEMI%
     EncaPresu!CotiMone_Presu = TICAMONE(MONEMI%)
     EncaPresu.Update
   End If
   '
   DetaPresu.FindFirst " Nro_Presu= '" & PRESU2$ & "' "
25 If DetaPresu.NoMatch = False Then
     DetaPresu.Edit
     PRUNI# = DetaPresu!PreUni_Det * TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
     DetaPresu!PreUni_Det = PRUNI#
     IMPOR# = DetaPresu!Impo_Det * TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
     DetaPresu!Impo_Det = IMPOR#
     DetaPresu.Update
     DetaPresu.FindNext " Nro_Presu= '" & PRESU2$ & "' "
     GoTo 25
   End If
   '
   Text1 = ""
   Text1 = PRESU2$
   Call CARGPRE
   
'   Call CARGPRE
'   '
End Sub

Private Sub Label24_Change()
   Call CATOIMP
End Sub

Private Sub Label24_DblClick()
    '
   Call CMIGRA04.CAMGRIS
   CMIGRA04.Label31 = Text1
   '
   CMIGRA04.Show 1
End Sub

Private Sub Label28_Change()
   Call CATOIMP
End Sub

Private Sub LABEL40_CHANGE()
    '
    Call TOTNETO
    '
End Sub
Private Sub Label40_DblClick()
'   '
'   With COTCAM04
'    .BLANFORMU
'    .Label31 = Text1
'    .Label7 = Text3
'    .Text9 = Text9
'    .Text10 = Text10
'    .Text8 = Text11
'    .Text18 = Text18
'    .Text5 = Text5
'    .Text7 = Text7
'    .Text4 = Text8
'    .Label17 = Label2
'    .Label17.ForeColor = Label2.ForeColor
'    .Label17.FontUnderline = Label2.FontUnderline
'    .TICAMBIS = TICAMBIS
'   End With
'   '
'   NUOTR$ = TRIM$(COTCAM04.Label31)
'   NUGRU$ = TRIM$(COTCAM04.Label2)
'   '
'   If NUOTR$ = "" Or ECOARCH("IPRESCAM", NUOTR$) = "" Then Exit Sub
'   If NUGRU$ = "" Then Exit Sub
'   '
'   'Call COTCAM04.CARGOLISMAT
'   Call COTCAM04.CAMGRIS
'   '
'   COTCAM04.Show 1
End Sub
Private Sub LABEL42_CHANGE()
    Call TOTMATCOM
End Sub
Private Sub Label42_DBLClick()
    '
   With CLICEN04
    .Label31 = Text1
    .Label7 = Text3
    .Text9 = Text9
    .Text10 = Text10
    .Text8 = Text11
    .Text18 = Text18
    .Text5 = Text5
    .Text7 = Text7
    .Text4 = Text8
    .Label17 = Label2
    .Text1 = Text2
    .Label17.ForeColor = Label2.ForeColor
    .Label17.FontUnderline = Label2.FontUnderline
    .TICAMBIS = TICAMBIS
   End With
   '
   Call CLICEN04.CAMGRIS
   '
   CLICEN04.Show 1
End Sub

Private Sub Label43_CHANGE()
    Call TOTMATCOM
End Sub

Private Sub Label43_DblClick()
    '
        '
'   With COTACC04
'    .Label31 = Text1
'    .Label7 = Text3
'    .Text9 = Text9
'    .Text10 = Text10
'    .Text8 = Text11
'    .Text18 = Text18
'    .Text5 = Text5
'    .Text7 = Text7
'    .Text4 = Text8
'    .Label17 = Label2
'    .Label17.ForeColor = Label2.ForeColor
'    .Label17.FontUnderline = Label2.FontUnderline
'    .TICAMBIS = TICAMBIS
'   End With
'   '
'   NUOTR$ = TRIM$(COTACC04.Label31)
'   NUGRU$ = TRIM$(COTACC04.Label2)
'   '
'   If NUOTR$ = "" Or ECOARCH("IPRESCAM", NUOTR$) = "" Then Exit Sub
'   If NUGRU$ = "" Then Exit Sub
'   '
'   Call COTACC04.CAMGRIS
'   '
'   COTACC04.Show 1
    '
End Sub

Private Sub Label44_CHANGE()
    Call TOTMATCOM
End Sub

Private Sub Label44_DblClick()
        '
   '
   CDESA04.Show 1
    '
End Sub

Private Sub Label45_Change()
   LICADI# = Val(Label42) * Val(Label45) / 100
   Label43 = FORMATNUM$(LICADI#, "F11.2")
End Sub

Private Sub Label47_CHANGE()
    Call TOTMATCOM
End Sub

Private Sub Label47_DblClick()
        '
   With CVARS04
    .Label31 = Text1
    .Label7 = Text3
    .Text9 = Text9
    .Text10 = Text10
    .Text8 = Text11
    .Text18 = Text18
    .Text5 = Text5
    .Text7 = Text7
    .Text4 = Text8
    .Label17 = Label2
    .Label17.ForeColor = Label2.ForeColor
    .Label17.FontUnderline = Label2.FontUnderline
    .TICAMBIS = TICAMBIS
   End With
   '
   NUOTR$ = TRIM$(CVARS04.Label31)
   NUGRU$ = TRIM$(CVARS04.Label2)
   '
   If NUOTR$ = "" Or ECOARCH("IPRESCAM", NUOTR$) = "" Then Exit Sub
   If NUGRU$ = "" Then Exit Sub
   '
   Call CVARS04.CAMGRIS
   '
   CVARS04.Show 1
   '
End Sub

Private Sub Label49_CHANGE()
    '
    Text23 = FORMATNUM(Val(Label49), "F11.2")
    Text38 = Text23
    Call CALCUOTAS
    '
End Sub

Private Sub Label5_dblClick()
    Call List1_DBLClick
End Sub

Private Sub Label52_CHANGE()
  '
  CAPUES% = Val(Text4): If CAPUES% > 5 Then CAPUES% = 5
  Call CATONET
  Label65(0) = TRIM$(FORMATNUM$(45 + 8 * CAPUES% + 0.009 * Val(Label52), "F9.2"))
  Label65(1) = TRIM$(FORMATNUM$(70 + 8 * CAPUES% + 0.014 * Val(Label52), "F9.2"))
  Label65(2) = TRIM$(FORMATNUM$(85 + 8 * CAPUES% + 0.025 * Val(Label52), "F9.2"))
  '
  ' enjuague para escorial
  If TRIM$(Text1) = "PR-060609-1" Then
    Label65(0) = "207.59"
    Label65(1) = "311.85"
    Label65(2) = "478.97"
  End If
  '
End Sub

Private Sub Label54_CHANGE()
  Call TOTGEN
End Sub

Private Sub LABEL62_CHANGE()
  '
  If XVALO(Label62) <= 1 Then
      TTXX$ = "Versión Monousuario (1 Puesto de Trabajo), apto para ulterior expansión a Multiusuario en Red LAN mediante adquisición de licencias adicionales."
    Else
      TTXX$ = "Versión Multiusuario para procesamiento concurrente (simultaneo) en red LAN, licenciado para un total de " + TRIM$(Label62) + " puestos de trabajo."
  End If
  Text40 = TTXX$
  '
End Sub
Private Sub Label62_DBLClick()
  '
   With COTMOB04
    .Label31 = Text1
    .Label7 = Text3
    .Text9 = Text9
    .Text10 = Text10
    .Text8 = Text11
    .Text18 = Text18
    .Text5 = Text5
    .Text7 = Text7
    .Text4 = Text8
    .Label17 = Label2
    .Label17.ForeColor = Label2.ForeColor
    .Label17.FontUnderline = Label2.FontUnderline
   End With
   '
   NUOTR$ = TRIM$(COTMOB04.Label31)
   NUGRU$ = TRIM$(COTMOB04.Label2)
   '
   If NUOTR$ = "" Or ECOARCH("IPRESCAM", NUOTR$) = "" Then Exit Sub
   If NUGRU$ = "" Then Exit Sub
   '
   Call COTMOB04.CAMGRIS
   '
   COTMOB04.Show 1
End Sub

Sub LABEL63_DBLCLICK()
  '
   '
   With COTRAN04
    .Label31 = Text1
    .Label7 = Text3
    .Text9 = Text9
    .Text10 = Text10
    .Text8 = Text11
    .Text18 = Text18
    .Text5 = Text5
    .Text7 = Text7
    .Text4 = Text8
    .Label17 = Label2
    .Label17.ForeColor = Label2.ForeColor
    .Label17.FontUnderline = Label2.FontUnderline
    .TICAMBIS = TICAMBIS
   End With
   '
   NUOTR$ = TRIM$(COTRAN04.Label31)
   NUGRU$ = TRIM$(COTRAN04.Label2)
   '
   If NUOTR$ = "" Or ECOARCH("IPRESCAM", NUOTR$) = "" Then Exit Sub
   If NUGRU$ = "" Then Exit Sub
   '
   Call COTRAN04.CAMGRIS
   '
   COTRAN04.Show 1
  '
End Sub

Private Sub Label66_DblClick()
   If VALORA = 60 Then
       VALORA = 75
     ElseIf VALORA = 75 Then
       VALORA = 90
   End If
   Call TEXT44_CHANGE
   Call Text14_Change
End Sub

Private Sub Label68_Change()
   Call CATOIMP
End Sub

Private Sub Label7_CHANGE()
   Call CATOIMP
End Sub

Private Sub Label70_CHANGE()
  '
  Call CATONET
  '
  If Option1.Value = True Then
       Call Option1_Click
    ElseIf Option2.Value = True Then
       Call Option2_Click
    ElseIf Option3.Value = True Then
       Call Option3_Click
    ElseIf Option4.Value = True Then
       Call Option4_Click
  End If
  '
End Sub

Private Sub Label71_CHANGE()
  Call TOTGEN
End Sub

Private Sub Label76_CHANGE()
  Call TOTGEN
End Sub

Private Sub Label86_Change()
  If Right$(Label86, 1) = "\" Then
    Label86 = Left$(Label86, Len(Label86) - 1)
  End If
End Sub

Private Sub List2_DblClick()
   If List2.ListIndex >= 0 Then
     List2.RemoveItem List2.ListIndex
   End If
End Sub

Private Sub List3_Click()
   Label85 = TRIM$(List3.TEXT)
End Sub

Private Sub Option8_Click()
   Text41 = "*.PDF"
End Sub

Private Sub Option7_Click()
   Text41 = "*.DOC"
End Sub

Private Sub Option6_Click()
   Text41 = "*.XLS"
End Sub

Private Sub Option5_Click()
   Text41 = "*.*"
End Sub

Private Sub LiGeClaCli_Click()
    Call Command13_Click
End Sub

Private Sub SuspOper_Click()
    Call Command3_Click
End Sub

Private Sub List1_DBLClick()
        'GENERO EL DAT
        If TRIM$(Text1.TEXT) <> "" Then
        NUOTRAB$ = TRIM$(Text1.TEXT)
        Call ABREORTRA
        SQLX$ = "SELECT * FROM PLACOB "
        SQLX$ = SQLX$ + "WHERE Nro_Otrab = '" & NUOTRAB$ & "' "
        SQLX$ = SQLX$ + "ORDER BY FeVen_Pcob ASC "
        Set PLACOTMP = OTrabajo.OpenRecordset(SQLX$, 4)
        J& = 0
        With PLACOTMP
          '
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
            On Error GoTo 0
            Do Until .EOF = True
              Y$ = REGBLAN("!CAPLACO")
              '
              NUTRA$ = TRIM$(!Nro_Otrab)
              Call REPLA(Y$, NUTRA$, 1, 16)
              NUORITOT% = !NuOrItem
              Call REPLA(Y$, MKI(NUORITOT%), 17, 2)
              DESCRIOT$ = TRIM$(!Refe_Pcob)
              Call REPLA(Y$, DESCRIOT$, 19, 64)
              IMPOT# = !Impo_Pcob
              Call REPLA(Y$, MKD(IMPOT#), 83, 8)
              FECOT% = CVINT(!FeVen_Pcob)
              Call REPLA(Y$, MKI(FECOT%), 91, 2)
              If IsNull(!NroFac_Pcob) = True Then
                NROFAOT$ = " "
              Else
                NROFAOT$ = TRIM$(!NroFac_Pcob)
              End If

              Call REPLA(Y$, NROFAOT$, 93, 16)
              MONEOT% = !MONEDA
              Call REPLA(Y$, MKI(MONEOT%), 109, 1)
              COTIOT# = !Coti_Pcob
              Call REPLA(Y$, MKD(COTIOT#), 110, 8)
              '
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
        '
        'Call GENEPLACO 'GENERO EL DAT
        PRF$ = "CAPLACO"
        ATRI$ = "/NC"
        ATRI$ = ATRI$ + "/TITUPAN=Plan de Cobranza"
        ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left))
        ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + SSTab1.Top + Frame18.Top - 500))
        VTB% = VENTABU%(PRF$ + ATRI$)
        'NUEVA VALIDACION
        If VTB% < 0 Then Exit Sub
        '
        If VTB% >= 0 Then
          'GUARDO O EDITO LOS REGISTROS EN TABLA
          'ACA VA RUTINA DE GRABACION
          Call ABREORTRA
          J& = ULTREG("!CAPLACO")
          For KKIL& = 1 To J&
             XX$ = REGLEIDO$("!CAPLACO", KKIL&)
             NUOTRAB$ = TRIM$(Text1.TEXT)
             ITE% = CVI(Mid$(XX$, 17, 2))
             'VALIDACION SI ITEM VALE 0
             If ITE% <= 0 Then GoTo 7
             '
             DESCRI$ = Mid$(XX$, 19, 64)
             IMPOT# = CVD(Mid$(XX$, 83, 8))
             FECOTR% = CVI(Mid$(XX$, 91, 2))
             NUFAC = Mid$(XX$, 93, 16)
             '
             MONE1% = CVI(Mid$(XX$, 109, 1))
             COTI1# = CVD(Mid$(XX$, 110, 8))
             '
             With PlanCob
             '/// Aca FindFirst de Item+NUOTRAB$, Si existe
             ' se edita el registro, sino,
             'se crea registro nuevo////
             SQLX$ = " Nro_Otrab = '" & NUOTRAB$ & "' "
             SQLX$ = SQLX$ + " AND NuOrItem = " & ITE% & " "
             .FindFirst SQLX$
             If .NoMatch = True Then
               .AddNew
                 !Nro_Otrab = NUOTRAB$
                 !NuOrItem = ITE%
                 !Refe_Pcob = DESCRI$
                 !Impo_Pcob = IMPOT#
                 !FeVen_Pcob = CVDAT(FECOTR%)
                 !MONEDA = MONE1%
                 !Coti_Pcob = COTI1#
               .Update
              Else
               .Edit
                 !Nro_Otrab = NUOTRAB$
                 !NuOrItem = ITE%
                 !Refe_Pcob = DESCRI$
                 !Impo_Pcob = IMPOT#
                 !FeVen_Pcob = CVDAT(FECOTR%)
                 !MONEDA = MONE1%
                 !Coti_Pcob = COTI1#
               .Update
              End If
              End With
7         Next KKIL&
        End If
    
'    '/////////*******ACA GENERO NUEVAMENTE EL DAT*******///////////
'        NUOTRAB$ = TRIM$(Text1.TEXT)
'        Call ABREORTRA
'        SQLX$ = "SELECT * FROM PLACOB "
'        SQLX$ = SQLX$ + "WHERE Nro_Otrab = '" & NUOTRAB$ & "' "
'        SQLX$ = SQLX$ + "ORDER BY FeVen_Pcob ASC "
'        Set PLACOTMP = OTrabajo.OpenRecordset(SQLX$, 4)
'        J& = 0
'        With PLACOTMP
'          '
'          On Error Resume Next
'          .MoveFirst
'          xxsd = .RecordCount
'          MsgBox xxsd
'          If Err < 1 Then
'            On Error GoTo 0
'            Do Until .EOF = True
'              Y$ = REGBLAN("!CAPLACO")
'              '
'              NUTRA$ = TRIM$(!Nro_Otrab)
'              Call REPLA(Y$, NUTRA$, 1, 16)
'              NUORITOT% = !NuOrItem
'              Call REPLA(Y$, MKI(NUORITOT%), 17, 2)
'              DESCRIOT$ = TRIM$(!Refe_Pcob)
'              Call REPLA(Y$, DESCRIOT$, 19, 64)
'              IMPOT# = !Impo_Pcob
'              Call REPLA(Y$, MKD(IMPOT#), 83, 8)
'              FECOT% = CVINT(!FeVen_Pcob)
'              Call REPLA(Y$, MKI(FECOT%), 91, 2)
'              If IsNull(!NroFac_Pcob) = True Then
'                NROFAOT$ = " "
'              Else
'                NROFAOT$ = TRIM$(!NroFac_Pcob)
'              End If
'
'              Call REPLA(Y$, NROFAOT$, 93, 16)
'              MONEOT% = !Moneda
'              Call REPLA(Y$, MKI(MONEOT%), 109, 2)
'              COTIOT# = !Coti_Pcob
'              Call REPLA(Y$, MKD(COTIOT#), 111, 8)
'              '
'              J& = J& + 1
'              Call GRAREG("!CAPLACO", Y$, J&)
'            .MoveNext
'            Loop
'          End If
'          On Error GoTo 0
'        End With
'        '
'        Call SETULTREG("!CAPLACO", J&)
'        Call CIERRARCH("!CAPLACO")
        '
    '/////////******ACA GUARDO EL PLAN EDITADO******///////////////
     With PlanCob
20   .FindFirst "Nro_Otrab = '" & NUOTRAB$ & "' "
     If .NoMatch = False Then
       .Delete
       GoTo 20
     End If
     '
     For U& = 1 To ULTREG("!CAPLACO")
        Y$ = REGLEIDO$("!CAPLACO", U&)
        'NUOTR$ = Mid$(Y$, 1, 16)
        ITE1% = CVI(Mid$(Y$, 17, 2))
        REFE1$ = Mid$(Y$, 19, 64)
        IMP1# = CVD(Mid$(Y$, 83, 8))
        FECVEN% = CVI(Mid$(Y$, 91, 2))
        NFAC1$ = Mid$(Y$, 93, 16)
        MON1% = CVI(Mid$(Y$, 109, 2))
        COTMON1# = CVD(Mid$(Y$, 110, 8))
       '
       .AddNew
        !Nro_Otrab = NUOTRAB$
        !NuOrItem = ITE1%
        !Refe_Pcob = REFE1$
        !Impo_Pcob = IMP1#
        !FeVen_Pcob = CVDAT(FECVEN%)
        !NroFac_Pcob = NFAC1$
        !MONEDA = MON1%
        !Coti_Pcob = COTMON1#
       .Update
     Next U&
   End With
   '//////*********/////////////
    List1.Clear
    Call CARGOPLACO
    '
    End If
    '
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
          Call REPLA(Y$, !MONEDA, 109, 2)
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



Private Sub NueOT_Click()
    Call Command21_Click
End Sub



Private Sub Option1_Click()
   Text17 = "Por el otorgamiento de Licencia de Uso "
   If Val(Label70) > 0 Then
        Text17 = Text17 + "y la Implementación "
      Else
        'Text17 = Text17 + "(sin Implementación) "
   End If
   Text17 = Text17 + "de un Software Aplicativo Integrado de Gestión Industrial y Administrativa, marca FLEXOFT, de nuestra manufactura, para operar bajo sistemas operativos Windows 95, 98, 2000, XP o NT, compuesto por los módulos y sub-sistemas que se detallan a continuación:"
End Sub

Private Sub Option2_Click()
   Text17 = "Por el otorgamiento de Licencia de Uso "
   If Val(Label70) > 0 Then
        Text17 = Text17 + "y la Implementación "
      Else
        'Text17 = Text17 + "(sin Implementación) "
   End If
   Text17 = Text17 + "de un Software Aplicativo Integrado de Gestión Industrial, marca FLEXOFT, de nuestra manufactura, para operar bajo sistemas operativos Windows 95, 98, 2000, XP o NT, compuesto por los módulos y sub-sistemas que se detallan a continuación:"
End Sub

Private Sub Option3_Click()
   Text17 = "Por el otorgamiento de Licencia de Uso "
   If Val(Label70) > 0 Then
        Text17 = Text17 + "y la Implementación "
      Else
        'Text17 = Text17 + "(sin Implementación) "
   End If
   Text17 = Text17 + "de un Software Aplicativo Integrado de Gestión Administrativa, marca FLEXOFT, de nuestra manufactura, para operar bajo sistemas operativos Windows 95, 98, 2000, XP o NT, compuesto por los módulos y sub-sistemas que se detallan a continuación:"
End Sub

Private Sub Option4_Click()
   Text17 = "Por el otorgamiento de Licencia de Uso "
   If Val(Label70) > 0 Then
        Text17 = Text17 + "y la Implementación "
      Else
        'Text17 = Text17 + "(sin Implementación) "
   End If
   Text17 = Text17 + "de un Software Aplicativo Auto-Integrado, marca FLEXOFT, de nuestra manufactura, para operar bajo sistemas operativos Windows 95, 98, 2000, XP o NT, compuesto por los módulos y sub-sistemas que se detallan a continuación:"
End Sub

Private Sub SSTab1_Click(PREVIOUS As Integer)
    If SSTab1.Tab = 3 Then
      Call Dir1_Change
    End If
End Sub

Private Sub Text10_Change()
   MODIFIC% = 1
End Sub

Private Sub Text11_GotFocus()
    Text25.SetFocus
End Sub

Sub GRABOT() 'GRABAR PRESUPUESTO DE CAMARA
   '
   If YACAR% < 1 Or MODIFIC% < 1 Then
     Exit Sub
   End If
   '
   If DERACCE%("GRADATPR", "Grabar Datos de Presupuesto FLEXOFT") < 2 Then
       Exit Sub
   End If
   '
   'VALIDACIONES
   If TRIM$(Text1) = "" Then
      Call COMUNI("Número de Presupuesto no Válido.")
      Exit Sub
   End If
   '
   If ECOARCH$("IPRESCAM", Text1) = "" Then
      Call COMUNI("Número de Presupuesto no Válido.")
      Exit Sub
   End If
   '
   If TRIM$(Text3) = "" Then
      Call COMUNI("Falta Referencia / Descripción Corta.")
   End If
   '
   If FECHANUM(Text10) < FECHANUM(Text9) Then
       Call COMUNI("Falta Fecha Validez de Oferta.")
   End If
   '
   If SSTab1.Tab = 0 Then
      If TRIM$(Text17) = "" Then
        Call COMUNI("Falta Especificación / Descripcion Larga.")
      End If
      '
      If Val(TRIM$(Text36)) < 1 Then
          Call COMUNI("Falta Cantidad.")
      End If
      '
      If CDbl(Val(TRIM$(Text38))) <= 0.005 Then
          Call COMUNI("Falta Importe.")
      End If
      '
      'If TRIM$(Text25) = "" Then
      '   Call COMUNI("Falta Pedido de Cotización.")
      'End If
      '
      'If FECHANUM(TRIM$(Text24)) < 1 Then
      '    Call COMUNI("Falta Fecha Pedido de Cotizacion.")
      'End If
      '
'      If TRIM$(Text13) = "" Then
'          Call COMUNI("Orden de Compra no Válida.")
'      End If
'      '
'      If TRIM$(Text12) = "" Then
'          Call COMUNI("Fecha de la Orden de Compra no Válida.")
'      End If
      '
'      If TRIM$(Text21) = "" Then
'          Call COMUNI("Comisión no Válida.")
'      End If
'      '
'      If TRIM$(CUETEXT) = "" Then
'          Call COMUNI("Cuenta de Imputación no Válida.")
'      End If
      '
'      If TRIM$(COTEXT(5)) = "" Then
'          Call COMUNI("Representante no Válido.")
'      End If
      '
'      If List1.ListCount <= 0 Then
'          Call COMUNI("Plan de Cobranza no Válido.")
'          Exit Sub
'      End If
      '
   End If
   NUOTR$ = TRIM$(Text1)
   If NUOTR$ = "" Then Exit Sub
   '
   Call ABREPRECAM
   With EncaPresu
     .FindFirst "  NRO_Presu= '" & NUOTR$ & "' "
     If .NoMatch = True Then
        Exit Sub
     End If
     On Error GoTo 0
     '
     .Edit
       OPCIBAS% = 0
       If Option1.Value = True Then OPCIBAS% = 1
       If Option2.Value = True Then OPCIBAS% = 2
       If Option3.Value = True Then OPCIBAS% = 3
       If Option4.Value = True Then OPCIBAS% = 4
     !NuOrItem_Presu = OPCIBAS%
     !Especi_Presu = Text17
     !Femi_Presu = CVDAT(FECHANUM(Text9))
     !FVali_Presu = CVDAT(FECHANUM(Text10))
     !REFE_Presu = TRIM$(Text3)
     !NroOcom_Presu = TRIM$(Text13)
     !FeOcom_Presu = CVDAT(FECHANUM(Text24))
     !PediCoti_Presu = TRIM$(Text25)
     !FePedCot_Presu = CVDAT(FECHANUM(Text12))
     !CanPed_Presu = Val(Text36)
     !PreVeN_PRESU = Val(Text38)
     !LENTREGA_Presu = Text2
     'medidas
     !Alto_Presu = Val(Text4)
     !Ancho_Presu = Val(Text14)
     !Largo_Presu = Val(Text41)
     !Espesor_Presu = Val(Text22)
     '
     If Val(COTEXT(5)) >= 0 Then
         !Vended_Presu = Val(COTEXT(5))
     End If
     !GGMat_Presu = Val(Text44)
     !GGOba_Presu = Val(Text26)
     !DesBon_Presu = Val(Text20)
     !ComiVend_Presu = Val(Text21)
     '
     '\\\OT-06-0145-WAR-17/04/06///
     !BoniVars_Presu = Val(Text28)
     !FormElec_Presu = Val(Text30)
     !HorasCons_Presu = Val(Text31)
     !ImporCons_Presu = Val(Label28)
     '\\\OT-06-0145-WAR-FIN///
     !Mone_Presu = MONEMI%
     !Precios_Presu = Text6
     !PlaEnt_Presu = Text15
     !ValOfer_Presu = Text32
     !ICONTAB = CUEINT%(CUETEXT) 'imputa. contab. de la EV
     !Observa_Presu = Text29
     !ForPag_Presu = Text16
     '
     !Anti_Presu = XVALO(Text33)
     !Cuot_Presu = XVALO(Text34)
     !Inter_Presu = XVALO(Text35)
     !Plazo_Entre = XVALO(Text39)
     !Plazo_Pago = XVALO(Text42)
     '
     SUPRABON% = 0
     If Check1.Value = 1 Then SUPRABON% = 1
     !Ajusta_Presu = SUPRABON%
     '
     .Update
   End With
   '
   MODIFIC% = 0
   Call GENEINDIPRECA
   '
End Sub

Private Sub Text12_Change()
    MODIFIC% = 1
End Sub
Private Sub Text12_GotFocus()
  '
 ' Text17.SetFocus
  '
End Sub

Private Sub Text13_GOTFOCUS()
  '
  'Text17.SetFocus
  '
End Sub
Private Sub Text13_Change()
    MODIFIC% = 1
End Sub

Private Sub Text14_Change()
   '
   ENTRENA# = VALORA * Val(Text14)
   Label7 = FORMATNUM(ENTRENA#, "F12.2")
   '
End Sub

Private Sub Text14_LOSTFOCUS()
    Text14 = FORMATNUM(Val(Text14), "F5.1")
End Sub

Private Sub Text17_Change()
    MODIFIC% = 1
End Sub

Private Sub Text18_GOTFOCUS()
    Text25.SetFocus
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
        Text14.SetFocus
    End If
    '
End Sub
Private Sub Text14_KeyPress(KeyAscii As Integer)
    '
    If KeyAscii = 13 Then
       Call Text14_LOSTFOCUS
    End If
    '
End Sub
Private Sub Text2_Change()
    MODIFIC% = 1
End Sub

Private Sub Text2_DblClick()
  ' Call Command11_Click
End Sub

Private Sub TEXT21_LOSTFOCUS()
    '
    Text21 = FORMATNUM$(Val(TRIM(Text21)), "F6.2")
    '
End Sub

Private Sub Text22_Change()
   DESCIMP = Val(Text22) * Val(Label70) / 100
   Text19 = FORMATNUM$(DESCIMP, "F11.2")
End Sub

Private Sub Text22_LOSTFOCUS()
  '
   Text22 = FORMATNUM(Val(Text22), "F6.1")
  '
End Sub

Private Sub Text23_GotFocus()
   CAJENTRA% = 21
End Sub

Private Sub Text27_GotFocus()
   CAJENTRA% = 10
End Sub

Private Sub Text33_LostFocus()
   Text33 = FORMATNUM(Val(Text33), "f5.1")
End Sub

Private Sub Text35_Change()
   Call CALCUOTAS
End Sub

Private Sub Text35_LostFocus()
   Call CALCUOTAS
   Text35 = FORMATNUM(Val(Text35), "f3.1")
End Sub

Private Sub Text42_Change()
   Call CALCUOTAS
End Sub

Private Sub Text42_LostFocus()
   Call CALCUOTAS
   Text42 = FORMATNUM(Val(Text42), "I2")
End Sub

Private Sub Text36_GotFocus()
    Text17.SetFocus
End Sub

Private Sub text38_gotfocus()
    Text17.SetFocus
End Sub

Private Sub Text38_lostfocus()
    '
    If TRIM$(Text1.TEXT) <> "" Then
      MODIFIC% = 1
      Text38 = FORMATNUM$(Val(Text38), "F12.2")
    End If
    '
End Sub

Private Sub Text39_Change()
   If Text39 > 365 Then
     Text39 = 365
     Exit Sub
   End If
   On Error Resume Next
   VScroll2.Value = 366 - Val(Text39)
   On Error GoTo 0
   '
   TTXX$ = TRIM$(Text39) + " días corridos de la aprobación de la presente propuesta."
   TTXX$ = TTXX$ + CRLF$ + CRLF$ + "La entrega se considerará cumplimentada con la recepción de la(s) caja(s) de envío conteniendo CD's de instalación, diskettes de habilitación de puestos de trabajo y manuales / documentación respaldatoria de los aplicativos licenciados."
   TTXX$ = TTXX$ + CRLF$ + CRLF$ + "Para el caso de propuestas que incluyan servicios de instalación, sean con o sin cargo adicional, la entrega se considerará cumplimentada una vez instalados los archivos correspondientes a los aplicativos licenciados sobre el disco rígido de vuestro servidor y verificado el funcionamiento general en al menos el 50 % de las estaciones de trabajo licenciadas."
   Text15 = TTXX$
   '
   Call CALCUOTAS
   '
End Sub

Private Sub Text4_Change()
   MODIFIC% = 1
   VALOBA = Val(Label42)
   If Val(Text4) < 1 Then Text4 = 1
   PORCE = 0
   If Text4 < 2 Then
       PORCE = 0
     ElseIf Text4 = 2 Then
       PORCE = 30
     ElseIf Text4 = 3 Then
       PORCE = 45
     ElseIf Text4 = 4 Then
       PORCE = 55
     ElseIf Text4 < 10 Then
       PORCE = 55 + (Val(Text4) - 4) * 8
     ElseIf Text4 < 16 Then
       PORCE = 55 + 48 + (Val(Text4) - 9) * 6
     Else
       PORCE = 55 + 48 + 36 + (Val(Text4) - 15)
   End If
   Label45 = PORCE
   '
   CLICEN04.Label8 = "Total Licencias (" & Text4 & " Puestos):"
   Call CALCUDETLICEN
   '
   Call Label52_CHANGE
   '
   HORAINS = 4 + 1.5 * (Val(Text4) - 1)
   Text44 = FORMATNUM(HORAINS, "F5.1")
   Label62 = TRIM$(Text4)
   'VALOAD = VALOBA * PORCE / 100
   'Label43 = FORMATNUM$(VALOAD, "F12.2")
End Sub

Private Sub Text4_LOSTFOCUS()
    Text4 = FORMATNUM(Val(Text4), "I3")
End Sub

Private Sub Text41_Change()
   File1.Pattern = Text41
   List3.Clear
   For KU& = 1 To File1.ListCount
     List3.AddItem File1.List(KU& - 1)
   Next KU&
End Sub

Private Sub Text44_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Call Text44_LOSTFOCUS
    End If
End Sub

Private Sub Text44_LOSTFOCUS()
  '
  If TRIM$(Text1.TEXT) <> "" Then
    MODIFIC% = 1
    PORCMAT# = Val(Text44)
    Text44 = FORMATNUM(PORCMAT#, "F5.1")
  End If
  '
End Sub
'
Sub TEXT44_CHANGE()
  '
  INSTALA# = VALORA * Val(Text44)
  Label68 = FORMATNUM(INSTALA#, "F12.2")
  '
End Sub
'
Function PORGASGEN$(NUPRESU1$, TIPGAST$)
'DEVUELVE EL PORCENTAJE DE GASTOS GENERALES
  Call ABREPRECAM
  With EncaPresu
    '
    NUPRESU2$ = TRIM$(NUPRESU1$)
    SQLX$ = " Nro_Presu = '" & NUPRESU2$ & "' "
    .FindFirst SQLX$
    If .NoMatch = False Then
     Select Case UCase(TIPGAST$)
       Case "MATERIALES"
         If IsNull(!GGMat_Presu) Then GoTo 10
         PORGASGEN$ = !GGMat_Presu
         Exit Function
       Case "OBRA"
         If IsNull(!GGOba_Presu) Then GoTo 10
         PORGASGEN$ = !GGOba_Presu
         Exit Function
     End Select
    End If
10  PORGASGEN$ = ""
    '
  End With
End Function


Private Sub Text5_GOTFOCUS()
    Text25.SetFocus
End Sub
Private Sub COTEXT_Change(Index As Integer)
ECO$ = ""
If TRIM$(COTEXT(5).TEXT) <> "" Then
    TV11% = TIPFIELD("VENDEDOR", 1)
    LF11% = LENFIELD("VENDEDOR", 1)
    ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
    ECO$ = TRIM$(ECOARCH$("VENDEDOR", ARGU$))
End If
If ECO$ <> "" Then
   DETEXT(5).Caption = ECO$
Else
   DETEXT(5).Caption = "No Definido"
End If
'   '
'   ECO$ = ""
'   If TRIM$(COTEXT(Index).TEXT) <> "" Then
'      TV11% = TIPFIELD(ARCHEX$(Index), 1)
'      LF11% = LENFIELD(ARCHEX$(Index), 1)
'      ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
'      ECO$ = TRIM$(ECOARCH$(ARCHEX$(Index), ARGU$))
'   End If
'   '
'   If ECO$ <> "" Then
'        DETEXT(Index).Caption = ECO$
'      Else
'        DETEXT(Index).Caption = "No Definido"
'   End If
'   '
   MODIFIC% = 1
End Sub

Private Sub COTEXT_DblClick(Index As Integer)
'   Call SELECHO(ARCHEX$(Index))
'   VDEVU$ = VALACT1$(ARCHEX$(Index))
'   If TRIM$(VDEVU$) <> "" Then
'      COTEXT(Index).TEXT = VDEVU$
'   End If
End Sub

Sub BLANFORMU()
   '
   Text2 = ""
   Text3 = ""
   Text24 = ""
   Text25.TEXT = ""
   Text38.TEXT = ""
   Text44.TEXT = ""
   'Text16.TEXT = ""
   Text26.TEXT = ""
   Text20.TEXT = ""
   Text23.TEXT = ""
   Text27.TEXT = ""
   TICAMBIS.TEXT = ""
   Text4.TEXT = ""
   Text14.TEXT = ""
   Text19.TEXT = ""
   Text29.TEXT = ""
   Text19.TEXT = ""
   Text36.TEXT = ""
   Text21.TEXT = ""
   Text23 = ""
   Text37 = ""
   Text3 = ""
   Text18 = ""
   Text5 = ""
   Text7 = ""
   Text8 = ""
   Text9 = ""
   Text10 = ""
   Text25 = ""
   Text11 = ""
   Text13 = ""
   Text12 = ""
   Text17 = ""
   List1.Clear
   Label65(0) = "": Label65(1) = "": Label65(2) = ""
   Text12 = ""
   CLICEN04.List1.Clear
   Text39 = "7"
   Text33 = "25.0"
   Text34 = "1"
   Label40 = ""
   Label52 = ""
   Label70 = ""
   Label54 = ""
   Label77 = ""
   Label49 = ""
   Call CALCUOTAS
   '
   YACAR% = 0
   '
End Sub

'
Private Sub Form_Unload(Cancel As Integer)
   '
'   Call GRAGENCLI
'   Call GRANOTACLI
'   Call CARDATCLIE
'
  Call CIERRARCH("*.*")
'   If AGRECLI% > 0 Then
'      Call SORTCLI(1)
'      AGRECLI% = 0
'   End If
  
  Call BAJALDISCO
  Call FINAL("")
  ' Hide
'   If APLINVO$ <> "" Then
'       Call FINAL("")
'     Else
'       Hide
'   End If
   '
End Sub

Private Sub GruCli_Click()
With OPMASCL04
  .Option1(9).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub GuaCli_Click()
    Call COMMAND15_CLICK
End Sub

Private Sub LiPreCli_Click()
With OPMASCL04
  .Option1(7).Value = True
  Call .Command16_Click
End With
End Sub

Private Sub List_Click()
   Call Command13_Click
End Sub

Private Sub List3_dblClick()
   Call CONECTADOC
End Sub

Private Sub LTEXT36_Change()
   DISPONI = Val(LTEXT36.Caption) - Val(LTEXT37.Caption) - Val(LTEXT35.Caption)
   LTEXT23.ForeColor = &H80000008
   If DISPONI < 0 Then LTEXT23.ForeColor = &HFF&
   LTEXT23.Caption = CSTRING$(MKS$(DISPONI), 4, 11, 2, 2)
   LTEXT23.Refresh
End Sub

Private Sub NueCli_Click()
   Call Command21_Click
End Sub

Private Sub OrDBCli_Click()
Call COMMAND15_CLICK
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

Private Sub Text1_Change()
   'For Index = 0 To 8: YACAR%(Index) = 0: Next Index
   Call BLANFORMU
   If TRIM$(Text1.TEXT) <> "" Then
      NUOTR$ = TRIM$(Text1.TEXT)
      Call CARGPRE
      Call CARTOTRUB
   End If
End Sub
'
Private Sub Text1_DblClick()
   '
   Call Command9_Click
   '
End Sub

Private Sub Text1_GotFocus()
   '
   'Call GRAGENCLI
   'Call GRANOTACLI
   '
End Sub

Private Sub Text19_Change()
    '
    Call TOTNETO
    '
End Sub
Sub text19_lostfocus()
   Text19 = FORMATNUM(Val(Text19), "F11.2")
End Sub
Private Sub text20_change()
  '
  DESCLIC = Val(Text20) * (Val(Label42) + Val(Label43)) / 100
  Text26 = FORMATNUM(DESCLIC, "F11.2")
  '
End Sub
'
Private Sub Text20_LOSTFOCUS()
  '
   Text20 = FORMATNUM(Val(Text20), "F6.1")
  '
End Sub
'

Private Sub Text21_Change()
   MODIFIC% = 1
End Sub
Private Sub text21_gotfocus()
  '
  If Val(COTEXT(5)) <= 0 Then
    COTEXT(5).SetFocus
  End If
  '
End Sub
Private Sub Text23_Change()
   MODIFIC% = 1
   If CAJENTRA% = 21 Then
     Text27 = FORMATNUM$(Val(Label49) - Val(Text23), "F12.2")
   End If
   Call CALCIVA
End Sub
Private Sub text23_lostfocus()
  '
  If TRIM$(Text1.TEXT) <> "" Then
    PORIVA21# = Val(Text23)
    Text23 = FORMATNUM(PORIVA21#, "F12.2")
  End If
  '
End Sub


Private Sub Text24_Change()
   MODIFIC% = 1
End Sub


Private Sub Text25_Change()
   MODIFIC% = 1
End Sub

Private Sub Text26_Change()
    '
    Call TOTNETO
    '
End Sub
Private Sub TEXT26_LOSTFOCUS()
  '
  If TRIM$(Text1.TEXT) <> "" Then
    BONIF1# = Val(Text26)
    Text26 = FORMATNUM(BONIF1#, "F12.2")
  End If
  '
End Sub
Private Sub Text27_Change()
   MODIFIC% = 1
   If CAJENTRA% = 10 Then
     Text23 = FORMATNUM$(Val(Label49) - Val(Text27), "F12.2")
   End If
'   IMPTAS1# = Val(Label49)
'   IMPTAS2# = Val(Text27)
'   If IMPTAS2# > IMPTAS1# Then Text27 = ""
'   IMPTMP# = IMPTAS1# - IMPTAS2#
'   If IMPTMP# = 0 Then Text23 = ""
'   If IMPTAS1# > IMPTAS2# Then
'     Text23 = FORMATNUM(IMPTMP#, "F11.2")
'   End If
   Call CALCIVA2
End Sub
Private Sub TEXT27_LOSTFOCUS()
  '
  If TRIM$(Text1.TEXT) <> "" Then
    PORIVA10# = Val(Text27)
    Text27 = FORMATNUM(PORIVA10#, "F11.2")
  End If
  '
End Sub
Private Sub Text28_Change()
    '
    Call TOTNETO
    '
End Sub
Sub text28_lostfocus()
    Text28 = FORMATNUM(Val(Text28), "F11.2")
End Sub
'
Private Sub Text3_Change()
   MODIFIC% = 1
End Sub

Private Sub Text3_GotFocus()
   If SSTab1.Tab <> 0 Then
      SSTab1.Tab = 0
   End If
End Sub

Private Sub Text30_Change()
   Call CATOIMP
End Sub

'\\\OT-06-0145-WAR-18/04/06///
Private Sub Text31_Change()
   MODIFIC% = 1
   HC1 = Val(Text31)
   IHC1 = HC1 * 90
   Label28 = FORMATNUM$(IHC1, "F11.2")
End Sub
Private Sub text31_lostfocus()
  HS1 = Val(Text31)
  Text31.TEXT = FORMATNUM$(HS1, "F6.1")
End Sub

Private Sub text30_lostfocus()
  QFE1 = Val(Text30)
  Text30.TEXT = FORMATNUM$(QFE1, "F11.2")
End Sub
'\\\OT-06-0145-WAR-FIN///
Private Sub Text33_Change()
   Call CALCUOTAS
   MODIFIC% = 1
End Sub

Private Sub Text34_Change()
   Call CALCUOTAS
   MODIFIC% = 1
End Sub
Sub Text34_LOSTFOCUS()
    Text34 = FORMATNUM(Val(Text34), "F2.0")
End Sub
Private Sub LText35_DblClick()
  Screen.MousePointer = 11
   If Val(Text1.TEXT) > 0 Then
     If Val(Text1.TEXT) <= 32767 Then
       NC% = Val(Text1.TEXT)
       If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
         Call COPYSTRU("CHECAR", "LICHECLI")
         Call HEADERS("LICHECLI", "")
         Call HEADERS("LICHECLI", "Cliente: " + RASOCLI$(NC%))
         '
         JJL& = 0
         RFF$ = RECFILT$("CHECAR", 3, MKI$(NC%))
         HOII% = HOY(HOS$)
         For UI& = 1 To Len(RFF$) Step 4
            RECHE& = CVS(Mid$(RFF$, UI&, 4))
            If RECHE& > 0 Then
               If RECHE& <= ULTREG&("CHECAR") Then
                  XX$ = REGLEIDO$("CHECAR", RECHE&)
                  If CVI(Mid$(XX$, 3, 2)) = NC% Then
                     If CVI(Mid$(XX$, 41, 2)) >= HOII% Then
                        JJL& = JJL& + 1
                        Call GRAREG("LICHECLI", XX$, JJL&)
                     End If
                  End If
               End If
            End If
         Next UI&
         Call SETULTREG("LICHECLI", JJL&)
         Call CIERRARCH("LICHECLI")
         Call CIERRARCH("CHECAR")
         '
         Call CONECRUN("*LICHECLI", "Listado de Cheques por Cliente/WAIT")
         '
       End If
     End If
   End If
   '
   Screen.MousePointer = 1
   '
End Sub


Private Sub Text36_Change()
   MODIFIC% = 1
End Sub
Sub TEXT36_LOSTFOCUS()
    Text36 = FORMATNUM(Val(Text36), "F6.1")
End Sub

Private Sub Text7_GOTFOCUS()
    Text25.SetFocus
End Sub

Private Sub Text8_GOTFOCUS()
    Text25.SetFocus
End Sub

Private Sub Text9_DblClick()
    If COMALTER%("¿ Actualiza Fecha de Presupuesto ?\\No, Conservar\Si, Actualizar") = 2 Then
      HOII% = HOY(HOS$)
      Text9 = HOS$
    End If
End Sub

Private Sub Text9_GOTFOCUS()
    Text25.SetFocus
End Sub

Private Sub VenCli_Click()
With OPMASCL04
  .Option1(4).Value = True
  Call .Command16_Click
End With
End Sub

Sub CARGPRE() 'CARGA LA FICHA DE UN PRESUPUESTO DE Cámara DETERMINADO
   '
   Static CAMON%, DMONE$(), TCMONE()
   '
   If CAMON% < 1 Then
     CAMON% = ULTREG&("TAMONED")
     If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
          VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
          DMONE$(U&) = DMX$
          TCMONE(U&) = VHX
         Next U&
     End If
   End If
   '
   NUPRECA$ = TRIM$(Text1)
   If NUPRECA$ = "" Or ECOARCH("IPRESCAM", Text1) = "" Then Exit Sub
   '
   Call ABREPRECAM
   SQLX$ = "SELECT * FROM ENCAPRESU"
   Set OTTMP = PresuCam.OpenRecordset(SQLX$, 4)
   With OTTMP
     .FindFirst "  NRO_pRESU= '" & NUPRECA$ & "' "
     If .NoMatch = True Then
        On Error GoTo 0
        'Call COMUNI("No Existe la Orden de Trabajo")
        Exit Sub
     End If
     On Error GoTo 0
     '
     Text3 = !REFE_Presu
     '
     NCLIE% = !Nclie_Presu
     Label22 = NCLIE%
     If NCLIE% > 0 Then
       Text18 = RASOCLI(NCLIE%)
       Text5 = DOMICLI$(NCLIE%)
       Text7 = CPOSCLI$(NCLIE%)
       Text8 = LOCACLI$(NCLIE%)
     End If
     '
     On Error Resume Next
     Text9 = FECHATEX(CVINT(!Femi_Presu))
     Text10 = FECHATEX(CVINT(!FVali_Presu))
     Text25 = !PediCoti_Presu
     Text24 = FECHATEX(CVINT(!FeOcom_Presu))
     Text11 = !uSUCONF_Presu
     Text13 = !NroOcom_Presu
     Text12 = FECHATEX(CVINT(!FePedCot_Presu))
     Text36 = !CanPed_Presu
       If Val(Text36) < 1 Then Text36 = "1"
     ALTPRE = !Alto_Presu: If ALTPRE < 1 Or ALTPRE > 99 Then ALTPRE = 1
     ANCPRE = !Ancho_Presu
     CUOTMAN = !Largo_Presu
     ESPRE = !Espesor_Presu
     Text4 = FORMATNUM$(ALTPRE, "I3")
       VScroll1.Value = 100 - Val(Text4)
     Text14 = FORMATNUM$(ANCPRE, "F5.1")
     Text41 = FORMATNUM$(CUOTMAN, "F11.2")
     Text20 = FORMATNUM$(!DesBon_Presu, "F6.1")
     Text22 = FORMATNUM$(ESPRE, "F6.1")
     Text26 = FORMATNUM$(!GGOba_Presu, "F11.2")
     Text44 = FORMATNUM(!GGMat_Presu, "F6.1")
     '\\\OT-06-0145-WAR-17/04/06///
     Text28 = FORMATNUM$(!BoniVars_Presu, "F11.2")
     Text30 = FORMATNUM$(!FormElec_Presu, "F11.2")
     Text31 = FORMATNUM$(!HorasCons_Presu, "F6.1")
     Label28 = FORMATNUM$(!ImporCons_Presu, "F11.2")
     '\\\OT-06-0145-WAR-FIN///
     Text2 = !LENTREGA_Presu
     COTEXT(5) = !Vended_Presu
     COMIVE = Val(!ComiVend_Presu)
     Text21 = FORMATNUM(COMIVE, "F6.2")
     On Error GoTo 0
     If IsNull(!ICONTAB) = True Then
        CUETEXT = ""
      Else
        CUETEXT = CODCUE(!ICONTAB) 'CARGO LA IMPUTACION CONTABLE
     End If
     '
     MONEMI% = !Mone_Presu
     If MONEMI% < 1 Then MONEMI% = 1
     Label2 = DMONE$(MONEMI%)
     Label2.ForeColor = RGB(0, 0, 255)
     If MONEMI% > 1 Then
        Label2.ForeColor = RGB(255, 0, 0)
     End If
     Label2.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
     '
     Label81 = Label2
     Label81.ForeColor = Label2.ForeColor
     Label81.FontUnderline = Label81.FontUnderline
     PVBAS = !PreVeN_PRESU
     Text38 = FORMATNUM$(PVBAS, "F12.2")
     '
     If IsNull(!Anti_Presu) = False Then
         Text33 = FORMATNUM(!Anti_Presu, "f5.1")
       Else
         Text33 = "25.0"
     End If
     If IsNull(!Cuot_Presu) = False Then
         Text34 = !Cuot_Presu
       Else
         Text34 = "1"
     End If
     If IsNull(!Inter_Presu) = False Then
         Text35 = !Inter_Presu
       Else
         Text35 = "0"
     End If
     If IsNull(!Plazo_Pago) = False Then
         Text42 = !Plazo_Pago
       Else
         Text42 = "0"
     End If
     If IsNull(!Plazo_Entre) = False Then
         Text39 = !Plazo_Entre
       Else
         Text39 = "7"
     End If
     '
     If IsNull(!ForPag_Presu) = False Then Text16 = !ForPag_Presu
     If IsNull(!Observa_Presu) = False Then Text29 = !Observa_Presu
     If IsNull(!Precios_Presu) = False Then Text6 = !Precios_Presu
     If IsNull(!PlaEnt_Presu) = False Then Text15 = !PlaEnt_Presu
     If IsNull(!ValOfer_Presu) = False Then Text32 = !ValOfer_Presu
     '
     CMIGRA04.Label31 = Text1  ' MIGRACION DE DATOS
     CDESA04.Label31 = Text1   ' DESARROLLOS ESPECIALES
     '
     OPCIBAS% = !NuOrItem_Presu
     Option1.Value = False: If OPCIBAS% = 1 Then Option1.Value = True
     Option2.Value = False: If OPCIBAS% = 2 Then Option2.Value = True
     Option3.Value = False: If OPCIBAS% = 3 Then Option3.Value = True
     Option4.Value = False: If OPCIBAS% = 4 Then Option4.Value = True
     '
     Call CALCUOTAS
     '
     If TRIM$(Text6) = "" Then
       TTXX$ = "Los valores expresados son en moneda argentina (pesos). Los importes parciales y sub-totales no incluyen el I.V.A. que se detalla por separado. "
       TTXX$ = TTXX$ + CRLF$ + CRLF$ + "Los pagos deberán efectuarse unicamente mediante cheque o transferencia bancaria. Nuestros funcionarios no están autorizados a aceptar pagos en efectivo. "
       TTXX$ = TTXX$ + CRLF$ + CRLF$ + "El importe correspondiente al IVA, en su totalidad, debera abonarse en el momento de la instalación, con cheque propio al día 8 del mes siguiente, sin excepción, según se detalla en el punto correspondiente a 'Forma de Pago'."
       Text6 = TTXX$
     End If
     '
     If TRIM$(Text29) = "" Then
       TTXX$ = "La forma de pago que se detalla en el cuadro correspondiente está referida exclusivamente a los rubros 'Licencias' y 'Desarrollos Especiales', siendo estos últimos exclusivamente los especificados en la presente propuesta."
       TTXX$ = TTXX$ + CRLF$ + CRLF$ + "Los importes correspondientes a actividades de instalación, implementación, capacitación de usuarios, migración de datos, consultoría, desarrollos complementarios no especificados en la presente revisten caracter estimativo, tanto se desarrollen en companía o en gabinete, y serán liquidados en forma semanal en base al siguiente tarifario vigente o el que rija a la fecha de prestación:"
       TTXX$ = TTXX$ + CRLF$ + CRLF$ + " - Analista Junior: Gab: $ 75 / hora ... I.Comp: $ 95 / Hora"
       TTXX$ = TTXX$ + CRLF$ + " - Analista Semi - Senior: Gab: $ 90 / hora ... I.Comp: $ 120 / Hora"
       TTXX$ = TTXX$ + CRLF$ + " - Consultor Senior: Gab: $ 135 / hora ... I.Comp: $ 175 / Hora"
       TTXX$ = TTXX$ + CRLF$ + "Todos los valores más IVA 21 %."
       Text29 = TTXX$
     End If
     '
     If TRIM$(Text32) = "" Then
       Text32 = "7 Dias Corridos. Luego favor consultar."
     End If
     '
     If IsNull(!Especi_Presu) = False Then
       If TRIM$(!Especi_Presu) <> "" Then
         Text17 = !Especi_Presu
       End If
     End If
     '
     SUPRABON% = 0
     If !Ajusta_Presu > 0 Then SUPRABON% = 1
     Check1.Value = SUPRABON%
     '
   End With
   '
   Call CARGOPLACO
   Call CARGOADJUN
   MODIFIC% = 0
   YACAR% = 1
   '
End Sub
'
Sub CARGOADJUN() 'CARGO LISTA DE DETALLE DE UN PRESU
   '
   If TRIM$(Text1) <> "" Then
     NUPRESU$ = TRIM$(Text1)
     Call ABREPRECAM
     SQLX$ = "SELECT * FROM DOCADJUN "
     SQLX$ = SQLX$ + "WHERE TRIM(Nro_presu) = '" & NUPRESU$ & "' "
     SQLX$ = SQLX$ + "ORDER BY RutArch ASC "
     Set DOCADTMP = PresuCam.OpenRecordset(SQLX$, 4)
     '
     List2.Clear
     With DOCADTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do Until .EOF = True
           List2.AddItem !RutArch
           .MoveNext
         Loop
       End If
     End With
   End If
   '
End Sub

Sub GENEINDIPRECA() 'GENERA EL INDICE DE ORDENES DE TRABAJO
   'If UPDATOTR% > 0 Then
      Screen.MousePointer = 11
      Call ABREPRECAM
      Set ORDETRATECO = PresuCam.OpenRecordset("SELECT NRO_Presu, Nclie_Presu, REFE_Presu,FEMI_Presu FROM encapresu  ORDER BY VAL(MID(NRO_PRESU,4,6)) ASC, VAL(MID(NRO_PRESU,11,4)) ASC;", dbOpenDynaset, dbReadOnly)
      '
      JJ& = 0
      On Error Resume Next
      ORDETRATECO.MoveLast
      If Err Then Exit Sub
      '
      FIN& = CAREC%
      REBLB$ = REGBLAN$("IPRESCAM")
      With ORDETRATECO
        .MoveFirst
        '
        Do While Not .EOF
          JJ& = JJ& + 1
          Call TRACE(20, JJ&, FIN&)
          NCLIE% = !Nclie_Presu
          EL2$ = REBLB$
          Call REPLA(EL2$, !Nro_PRESU, 1, 16)
            FEX = CVINT(!Femi_Presu)
            REFEINDI$ = FECHATEX$(FEX) + " - " + TRIM$(FANTCLI$(NCLIE%))
            REFEINDI$ = REFEINDI$ + " - " + !REFE_Presu
          Call REPLA(EL2$, REFEINDI$, 17, 64)
          Call GRAREG("IPRESCAM", EL2$, JJ&)
          '
        .MoveNext
        Loop
      End With
      '
90    ORDETRATECO.Close
      Call SETULTREG("IPRESCAM", JJ&)
      'Call FILESORT("IPRESCAM", "", 1, 1, "ASC")
      Call FRESHECHO("IPRESCAM")
      Call BAJALDISCO
      UPDATOTR% = 0
      Screen.MousePointer = 1
End Sub
'
'\\\OT-05-0394-WAR-FIN///
Sub CARTOTRUB() 'CALCULO LOS TOTALES DE LOS RUBROS
  '
  NUPRECA$ = TRIM$(Text1)
  If NUPRECA$ = "" Or ECOARCH("IPRESCAM", Text1) = "" Then Exit Sub
  '
  Call CARGODETLICEN 'CALCULO DE TOTAL DE LICENCIAS
  Call CATOACC 'CALCULO DE TOTAL DE ACCESORIOS
  'Call CATOEQU 'CALCULO DE TOTAL DE EQUIPOS
  Call CATOVAR 'CALCULO DE TOTAL DE VARIOS
  'Call CATOMOB 'CALCULO DE TOTAL DE MOB
  'Call CATOTRAN 'CALCULO DE TOTAL DE TRANSPORTE
  '
End Sub

Sub CATOLIC() 'CALCULO DE TOTAL DE LICENCIAS
   '
  NUPRESU$ = TRIM(Text1)
  Call ABREPRECAM
  SQLX$ = "SELECT * FROM DETAPRESU "
  SQLX$ = SQLX$ + "WHERE TRIM(Nro_presu) = '" & NUPRESU$ & "' "
  SQLX$ = SQLX$ + "AND RUBRO_DET = 1 "
  SQLX$ = SQLX$ + "ORDER BY CodExte_Det ASC "
  Set DETPRETMP = PresuCam.OpenRecordset(SQLX$, 4)
  J& = 0
  With DETPRETMP
   '
   On Error Resume Next
   .MoveFirst
   If Err < 1 Then
    On Error GoTo 0
    TOT1# = 0
    Do Until .EOF = True
      TOT1# = TOT1# + !Impo_Det
    .MoveNext
    Loop
   End If
   On Error GoTo 0
  End With
  '
  Label42 = FORMATNUM(TOT1#, "F11.2")
  Call Label45_Change
  '
End Sub
Sub CATOACC()
  '
Exit Sub
    NUPRESU$ = TRIM(Text1)
  Call ABREPRECAM
  SQLX$ = "SELECT * FROM DETAPRESU "
  SQLX$ = SQLX$ + "WHERE TRIM(Nro_presu) = '" & NUPRESU$ & "' "
  SQLX$ = SQLX$ + "AND RUBRO_DET = 3 "
  SQLX$ = SQLX$ + "ORDER BY CodExte_Det ASC "
  Set DETPRETMP = PresuCam.OpenRecordset(SQLX$, 4)
  J& = 0
  With DETPRETMP
   '
   On Error Resume Next
   .MoveFirst
   If Err < 1 Then
    On Error GoTo 0
    TOT1# = 0
    Do Until .EOF = True
      TOT1# = TOT1# + !Impo_Det
    .MoveNext
    Loop
   End If
   On Error GoTo 0
  End With
  '
  Label43 = FORMATNUM(TOT1#, "F11.2")
  '
End Sub

Sub CATOVAR()
  '
    NUPRESU$ = TRIM(Text1)
  Call ABREPRECAM
  SQLX$ = "SELECT * FROM DETAPRESU "
  SQLX$ = SQLX$ + "WHERE TRIM(Nro_presu) = '" & NUPRESU$ & "' "
  SQLX$ = SQLX$ + "AND RUBRO_DET = 5 "
  SQLX$ = SQLX$ + "ORDER BY CodExte_Det ASC "
  Set DETPRETMP = PresuCam.OpenRecordset(SQLX$, 4)
  J& = 0
  With DETPRETMP
   '
   On Error Resume Next
   .MoveFirst
   If Err < 1 Then
    On Error GoTo 0
    TOT1# = 0
    Do Until .EOF = True
      TOT1# = TOT1# + !Impo_Det
    .MoveNext
    Loop
   End If
   On Error GoTo 0
  End With
  '
  Label47 = FORMATNUM(TOT1#, "F11.2")
  '
End Sub

Sub CATOTRAN() 'CALCULO DE TOTAL DE TRANSPORTE
 '
  NUPRESU$ = TRIM(Text1)
  Call ABREPRECAM
  SQLX$ = "SELECT * FROM DETAPRESU "
  SQLX$ = SQLX$ + "WHERE TRIM(Nro_presu) = '" & NUPRESU$ & "' "
  SQLX$ = SQLX$ + "AND RUBRO_DET = 12 "
  SQLX$ = SQLX$ + "ORDER BY DESCRI_Det ASC "
  Set DETPRETMP = PresuCam.OpenRecordset(SQLX$, 4)
  J& = 0
  With DETPRETMP
   '
   On Error Resume Next
   .MoveFirst
   If Err < 1 Then
    On Error GoTo 0
    TOT1# = 0
    Do Until .EOF = True
      TOT1# = TOT1# + !Impo_Det
    .MoveNext
    Loop
   End If
   On Error GoTo 0
  End With
  '
  Label63 = FORMATNUM(TOT1#, "F11.2")
End Sub
Sub CATOMOB() 'CALCULO DEL TOTAL DE MANO DE OBRA
  '
  NUPRESU$ = TRIM(Text1)
  Call ABREPRECAM
  SQLX$ = "SELECT * FROM DETAPRESU "
  SQLX$ = SQLX$ + "WHERE TRIM(Nro_presu) = '" & NUPRESU$ & "' "
  SQLX$ = SQLX$ + "AND RUBRO_DET = 11 "
  SQLX$ = SQLX$ + "ORDER BY DESCRI_Det ASC "
  Set DETPRETMP = PresuCam.OpenRecordset(SQLX$, 4)
  J& = 0
  With DETPRETMP
   '
   On Error Resume Next
   .MoveFirst
   If Err < 1 Then
    On Error GoTo 0
    TOT1# = 0
    Do Until .EOF = True
      TOT1# = TOT1# + !Impo_Det
    .MoveNext
    Loop
   End If
   On Error GoTo 0
  End With
  '
  Label62 = FORMATNUM(TOT1#, "F11.2")
End Sub

Sub TOTMATCOM() 'CALCULO DEL TOTAL DE MATERIALES Y COMPONENTES
  '
  TOTPUE1# = Val(Label42)
  TOTACC1# = Val(Label43)
  TOTEQU1# = Val(Label44)
  TOTVAR1# = Val(Label47)
  TOMACO1# = TOTCAM1# + TOTPUE1# + TOTACC1# + TOTEQU1# + TOTVAR1#
  Label52 = FORMATNUM(TOMACO1#, "F11.2")
  '
End Sub
'
Sub CATOGG() 'CALCULO DEL TOTAL DE GASTOS GENERALES
  '
  TOTMAT# = Val(Label68)
  TOTOB# = Val(Label69)
  TOTGG# = TOTMAT# + TOTOB#
  Label70 = FORMATNUM(TOTGG#, "F11.2")
  '
End Sub
'
Sub CATONET() 'CALCULO DEL TOTAL NETO
  '
  TOTMAT# = Val(Label52)
  TOTGAGE# = Val(Label70)
  TOTNET# = TOTMAT# + TOTGAGE#
  Label31 = FORMATNUM(TOTNET#, "F11.2")
  Label40 = FORMATNUM(TOTNET#, "F11.2")
  '
  Call text20_change
  Call Text22_Change
End Sub

Sub CATOIMP() 'CALCULO DEL TOTAL NETO
  '
  TOTIMP# = Val(Label68) + Val(Label7) + Val(Label24) + Val(Text30) + Val(Label28)
  Label70 = FORMATNUM(TOTIMP#, "F11.2")
  '
End Sub

Sub TOTNETO()
  '
  TOTANET# = Val(Label40) - Val(Text26) - Val(Text19) - Val(Text28)
  If TOTANET# < 0 Then TOTANET# = 0
  Label49 = FORMATNUM(TOTANET#, "F11.2")
  '
End Sub
'
Sub CALCIVA()
  '
10 IMPTAS1# = Val(Text23)
   PORTAS1# = Val(PORIVA(0))
   If IMPTAS1# > Val(Label49) Then Text23 = Val(Label49): GoTo 10
   IMPIVA1# = IMPTAS1# * PORTAS1# / 100
   Label54 = FORMATNUM(IMPIVA1#, "F11.2")
  '
End Sub
'
Sub CALCIVA2()
  '
  IMPTAS2# = Val(Text27)
  PORTAS2# = Val(PORIVA(1))
  IMPIVA2# = IMPTAS2# * PORTAS2# / 100
  Label71 = FORMATNUM(IMPIVA2#, "F11.2")
  '
End Sub
'
Sub TOTGEN()
  '
  IMPNET# = Val(Label49)
  IMPIVA1# = Val(Label54)
  IMPIVA2# = Val(Label71)
  IMPPER# = Val(Label76)
  TOTGEN1# = IMPNET# + IMPIVA1# + IMPIVA2# + IMPPER#
  Label77 = FORMATNUM(TOTGEN1#, "F11.2")
  '
End Sub
'
Sub CARGOPLACO() 'CARGO PLAN DE COBRA. DE UNA OT DETERMINADA
   '
   NUOTR$ = TRIM$(Text1)
   If NUOTR$ = "" Or ECOARCH("IPRESCAM", Text1) = "" Then Exit Sub
   '
   Call ABREORTRA
   SQLX$ = "SELECT * FROM PLACOB "
   SQLX$ = SQLX$ + "WHERE Nro_Otrab = '" & NUOTR$ & "' "
   SQLX$ = SQLX$ + "ORDER BY FeVen_Pcob ASC "
   Set OTTMP = OTrabajo.OpenRecordset(SQLX$, 4)
   With OTTMP
      '
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      '
      Do Until .EOF = True
         NOITOT1$ = CStr(!NuOrItem)
         DESCRIOT1$ = CStr(!Refe_Pcob)
         IMPOT1$ = FORMATNUM(!Impo_Pcob, "F11.2")
         FEOT1$ = FECHATEX(CVINT(!FeVen_Pcob))
         Y$ = Space$(64)
         Call REPLA(Y$, AJUSTD$(NOITOT1$, 3), 1, 2)
         Call REPLA(Y$, AJUSTI$(DESCRIOT1$, 15), 6, 15)
         Call REPLA(Y$, IMPOT1$, 20, 11)
         Call REPLA(Y$, FEOT1$, 33, 8)
         List1.AddItem Y$
        .MoveNext
      Loop
      
      '
   End With
End Sub

Sub CARGOLISGRU() 'CARGO LISTA DE GRUPOS DE UNA OT DETERMINADA
   '
   NUOTR$ = TRIM$(Text1)
   If NUOTR$ = "" Or ECOARCH("IPRESCAM", Text1) = "" Then Exit Sub
   '
   Call ABREORTRA
   SQLX$ = "SELECT * FROM GRUPOTRA "
   SQLX$ = SQLX$ + "WHERE Nro_Otrab = '" & NUOTR$ & "' "
   Set OTTMP = OTrabajo.OpenRecordset(SQLX$, 4)
   '
   REBLA$ = REGBLAN$("GRUORTRA")
   JJ& = 0
   List3.Clear
   '
   With OTTMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do Until .EOF = True
         NROGRU$ = !Nro_Grup
         If IsNull(!PROG_GRUP) Then
            PROGRU$ = ""
         Else
            PROGRU$ = !PROG_GRUP
         End If
         If IsNull(!COMP_GRUP) Then
            COMGRU$ = ""
         Else
            COMGRU$ = !COMP_GRUP
         End If
         If IsNull(!DENO_GRUP) Then
            DENGRU$ = ""
         Else
            DENGRU$ = !DENO_GRUP
         End If
         If IsNull(!OBSERVA_GRUP) Then
            OBSGRU$ = ""
         Else
            OBSGRU$ = !OBSERVA_GRUP
         End If
         '
         z$ = REBLA$
         Call REPLA(z$, NROGRU$, 1, 8)
         Call REPLA(z$, PROGRU$, 9, 12)
         Call REPLA(z$, COMGRU$, 21, 12)
         Call REPLA(z$, DENGRU$, 33, 64)
         Call REPLA(z$, OBSGRU$, 97, 64)
         JJ& = JJ& + 1
         Call GRAREG("!GRUORTRA", z$, JJ&)
         '
         Y$ = Space$(128)
         Call REPLA(Y$, NROGRU$, 1, 8)
         Call REPLA(Y$, PROGRU$, 10, 6)
         Call REPLA(Y$, COMGRU$, 17, 6)
         Call REPLA(Y$, DENGRU$, 24, 48)
         Call REPLA(Y$, OBSGRU$, 73, 24)
         List3.AddItem Y$
       .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   Call SETULTREG("!GRUORTRA", JJ&)
   '
End Sub
Sub LEETASIVA()
   '
   URETA& = ULTREG&("GRUPIVA")
   If URETA& < 2 Then
     Call MENSERR(24, "Falta Definir Grupos de I.V.A.")
     Call CIERRARCH("*.*")
     Call FINAL("")
   End If
   '
   For U& = 1 To 2
      X$ = REGLEIDO$("GRUPIVA", U&)
      TAIVA(U& - 1) = CVS(Mid$(X$, 13, 4))
      PORIVA(U& - 1) = TRIM$(CSTRING$(MKS$(TAIVA(U& - 1)), 3, 6, 2, 2))
   Next U&
   'PORIVA(2) = "50.00" ' responsable no inscripto
   Call CIERRARCH("GRUPIVA")
   ALIIVA = ALIVA
   '
End Sub

Private Sub TICAMBIS_GOTFOCUS()
  '
  Command15.SetFocus
  '
End Sub
'\\\OT-05-0394-WAR-10/06/05///
Sub PRIDETPRE() ' IMPRESION DE DETALLE PRESUPUETOS X RUBRO
    '
    '26/03/07 (OT-07-0132)
    OPT% = COMALTER%("Opciones de Formato de Salida:\\Completo\Compacto")
    If OPT% = 1 Then
         COMPACTO% = 0
       ElseIf OPT% = 2 Then
         COMPACTO% = 1
       Else
         Exit Sub
    End If
    '26/03/07 (FIN)
    '
    FORIPRE$ = CONTROL$("", "DETAPRE")
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      AREF1% = Val(Mid$(TRIM$(UCase$(ATRIFORM$(FORIPRE$, "REF-MAT1/FORMATO"))), 2))
         If AREF1% < 24 Then AREF1% = 24
      AREF2% = Val(Mid$(TRIM$(UCase$(ATRIFORM$(FORIPRE$, "REF-MAT2/FORMATO"))), 2))
         If AREF2% < 24 Then AREF3% = 24
      AREF3% = Val(Mid$(TRIM$(UCase$(ATRIFORM$(FORIPRE$, "REF-MAT2/FORMATO"))), 2))
         If AREF3% < 24 Then AREF3% = 24
      Screen.MousePointer = 11
      FECHDOC$ = Text9
      NUMEDOC$ = TRIM$(Text1)
      TIPODOC$ = "Presupuesto "
      '
      DESTIDOC% = Val(Label22) 'CLIENTE
      '
      CODPARAM$(1) = "NUME-COM" 'NUMERO DE PRESU
      CODPARAM$(2) = "FECH-EMI" 'FECHA DE EMISION
      CODPARAM$(3) = "FECH-VEN" 'FECHA DE VALIDEZ
      CODPARAM$(4) = "FECH-LAN" 'FECHA DE PED. COTIZ.
      CODPARAM$(5) = "NRO-AENT" 'PEDIDO COTIZACION
      CODPARAM$(6) = "FECH-REC" 'FECHA DE O/C
      CODPARAM$(7) = "NRO-OCOM" 'O/C
      CODPARAM$(8) = "CTAC-CLI" 'CONTACTO
      
      CAPARDOC% = 8
      
      DOCPARAM(1) = TRIM(Text1)
      DOCPARAM(2) = Text9
      DOCPARAM(3) = Text10
      DOCPARAM(4) = Text24
      DOCPARAM(5) = Text25
      DOCPARAM(6) = Text12
      DOCPARAM(7) = Text13
      DOCPARAM(8) = Text2
      '
 'AQUI PONER LOS CODIGOS DE COLUMNAS A LISTAR
      CODTABU1$(1) = "REF-MAT1" 'RUBRO
      CODTABU1$(2) = "REF-MAT2" 'SUB-RUBRO
      CODTABU1$(3) = "REF-MAT3" 'CODIGO
      CODTABU1$(4) = "PRE-NETO" 'DESCRIPCION
      CODTABU1$(5) = "NET-ITEM" 'METROS
      CODTABU1$(6) = "CANTIDAD" 'CANTIDAD
      CODTABU1$(7) = "REF-ASIE" 'PLAN DE PAGOS
      CACOLTAB1% = 7
      '
      CAITAB1% = 0
      CAITCRU% = 0
      '
      'If COMPACTO% = 1 Then GoTo 10 '26/03/07 (OT-07-0132)
      '
      'DESCRIPCION GENERAL
      TELIBRE$ = Text17
      If TRIM$(TELIBRE$) <> "" Then
        Call FRATEXTO(TELIBRE$, AREF1%)
        For KKU1% = 1 To CARETEX%
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = RETEX$(KKU1%)
        Next KKU1%
        CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
      End If
      '
10    TELIBRE$ = Text40
      If TRIM$(TELIBRE$) <> "" Then
        Call FRATEXTO(TELIBRE$, AREF1%)
        For KKU1% = 1 To CARETEX%
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = RETEX$(KKU1%)
        Next KKU1%
        CAITAB1% = CAITAB1% + 2 ' PARA RENGLON EN BLANCO
      End If
      '
      If COMPACTO% = 1 Then GoTo 11 '26/03/07 (OT-07-0132)
      '
      ICAPI% = 0
      'DETALLE DE LICENCIAS
11    If Val(Label42) > 0.01 Then
        CAITAB1% = CAITAB1% + 1
        ICAPI% = ICAPI% + 1
        TETABU1$(1, CAITAB1%) = TRIM$(Str$(ICAPI%)) + ".- Licenciamiento Módulos Preconfigurados"
        CAITAB1% = CAITAB1% + 1
        '
        CLICEN04.Label31 = Text1
        '
        ILETRA% = Asc("a")
        SUMAPAR# = 0
        For U& = 1 To CLICEN04.List1.ListCount
          TTXX$ = CLICEN04.List1.List(U& - 1)
          CODE$ = TRIM$(Left$(TTXX$, 16))
          CIXI% = CODINT%(CODE$)
          DESX$ = "1." + Chr$(ILETRA%) + ".- " + DESCRIT0$(CIXI%)
          PREUN$ = Mid$(TTXX$, 79, 12)
          '
          Call FRATEXTO(DESX$, AREF2%)
          For KKU1% = 1 To CARETEX%
            CAITAB1% = CAITAB1% + 1
            TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
          Next KKU1%
          TETABU1$(4, CAITAB1%) = PREUN$
          '
          ILETRA% = ILETRA% + 1
          SUMAPAR# = SUMAPAR# + Val(PREUN$)
          If COMPACTO% = 1 Then GoTo 12 '26/03/07 (OT-07-0132)
          '
          CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
          '
          ESPE$ = ESPECIFIT$(CIXI%)
          If TRIM$(ESPE$) <> "" Then
            Call FRATEXTO(ESPE$, AREF3%)
            For KKU1% = 1 To CARETEX%
              CAITAB1% = CAITAB1% + 1
              TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
            Next KKU1%
            CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
          End If
          '
12      Next U&
        '
        If COMPACTO% = 1 Then CAITAB1% = CAITAB1% + 1
        '
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = "Subtotal Licenciamiento Módulos Preconfigurados"
        TETABU1$(5, CAITAB1%) = FORMATNUM$(SUMAPAR#, "F12.2")
        If COMPACTO% = 1 Then
            CAITAB1% = CAITAB1% + 2
          Else
            CAITAB1% = CAITAB1% + 3
        End If
        '
      End If
      '
      'DETALLE DE DESARROLLOS ESPECIALES
13    If Val(Label44) > 0.01 Then
        CAITAB1% = CAITAB1% + 1
        ICAPI% = ICAPI% + 1
        TETABU1$(1, CAITAB1%) = TRIM$(Str$(ICAPI%)) + ".- Desarrollos Especiales"
        CAITAB1% = CAITAB1% + 1
        '
        ILETRA% = Asc("a")
        SUMAPAR# = 0
        For KKI% = 0 To 9
          REFEX3$ = CDESA04.Text2(KKI%)
          ESPE$ = CDESA04.Text3(KKI%)
          PREUN$ = CDESA04.Text1(KKI%)
          '
          If TRIM$(REFEX3$) <> "" Or TRIM$(DESCR$) <> "" Then
            TTXX$ = REFEX3$
            DESX$ = TRIM$(Str$(ICAPI%)) + "." + Chr$(ILETRA%) + ".- " + REFEX3$
            '
            Call FRATEXTO(DESX$, AREF2%)
            For KKU1% = 1 To CARETEX%
              CAITAB1% = CAITAB1% + 1
              TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
            Next KKU1%
            TETABU1$(4, CAITAB1%) = PREUN$
            If COMPACTO% < 1 Then CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
            ILETRA% = ILETRA% + 1
            SUMAPAR# = SUMAPAR# + Val(PREUN$)
            '
            If COMPACTO% = 1 Then GoTo 14 '26/03/07 (OT-07-0132)
            '
            If TRIM$(ESPE$) <> "" Then
              Call FRATEXTO(ESPE$, AREF3%)
              For KKU1% = 1 To CARETEX%
                CAITAB1% = CAITAB1% + 1
                TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
              Next KKU1%
              CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
            End If
          End If
          '
14      Next KKI%
        '
        If COMPACTO% = 1 Then CAITAB1% = CAITAB1% + 1
        '
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = "Subtotal Desarrollos Especiales"
        TETABU1$(5, CAITAB1%) = FORMATNUM$(SUMAPAR#, "F12.2")
        If COMPACTO% = 1 Then
            CAITAB1% = CAITAB1% + 2
          Else
            CAITAB1% = CAITAB1% + 3
        End If
        '
      End If
      '
      'DETALLE DE IMPLEMENTACIóN
      If Val(Label70) > 0.01 Then
        CAITAB1% = CAITAB1% + 1
        ICAPI% = ICAPI% + 1
        TETABU1$(1, CAITAB1%) = TRIM$(Str$(ICAPI%)) + ".- Servicios de Implementación (estimados)"
        CAITAB1% = CAITAB1% + 1
        '
        ILETRA% = Asc("a")
        SUMAPAR# = 0
        '
        'INSTALACION EN VUESTRA SEDE
        If Val(Label68) > 0 Then
          DESX$ = TRIM$(Str$(ICAPI%)) + "." + Chr$(ILETRA%) + ".- " + "Instalación en vuestra sede"
          PREUN$ = Label68
          '
          Call FRATEXTO(DESX$, AREF2%)
          For KKU1% = 1 To CARETEX%
            CAITAB1% = CAITAB1% + 1
            TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
          Next KKU1%
          TETABU1$(4, CAITAB1%) = PREUN$
          ILETRA% = ILETRA% + 1
          SUMAPAR# = SUMAPAR# + Val(PREUN$)
          '
          If COMPACTO% = 1 Then GoTo 16 '26/03/07 (OT-07-0132)
          '
          CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
          ESPE$ = "Comprende la instalación de los archivos correspondientes a módulos ejecutables, controladores de proceso / archivos y formularios electrónicos correspondientes a los productos licenciados sobre el disco rígido de vuestro servidor. "
          ESPE$ = ESPE$ + "Incluye la instalación de los componentes y librerias de Windows necesarias para el correcto funcionamiento de los aplicativos licenciados, unicamente a partir de sistemas operativos originales suministrados por Uds., sin verificación alguna de origen o legitimidad de los mismos."
          ESPE$ = ESPE$ + CRLF$ + CRLF$ + CRLF$ + "Incluye la conexion del disco servidor como unidad de red y la creación de accesos directos al ejecutable FLEXOFT que se establezca como 'principal' en cada uno de las estaciones de trabajo licenciadas."
          If TRIM$(ESPE$) <> "" Then
            Call FRATEXTO(ESPE$, AREF3%)
            For KKU1% = 1 To CARETEX%
              CAITAB1% = CAITAB1% + 1
              TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
            Next KKU1%
            CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
          End If
          '
        End If
        '
       'ENTRENAMIENTO DE USUARIOS
16      If Val(Label7) > 0 Then
          DESX$ = TRIM$(Str$(ICAPI%)) + "." + Chr$(ILETRA%) + ".- " + "Entrenamiento y Capacitación de Usuarios"
          PREUN$ = Label7
          '
          Call FRATEXTO(DESX$, AREF2%)
          For KKU1% = 1 To CARETEX%
            CAITAB1% = CAITAB1% + 1
            TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
          Next KKU1%
          TETABU1$(4, CAITAB1%) = PREUN$
          ILETRA% = ILETRA% + 1
          SUMAPAR# = SUMAPAR# + Val(PREUN$)
          '
          If COMPACTO% = 1 Then GoTo 17 '26/03/07 (OT-07-0132)
          '
          CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
          ESPE$ = "Comprende un total de " + Text14 + " horas de capacitación en vuestra sede, en bloques de 3.0 horas como mínimo cada uno, "
          ESPE$ = ESPE$ + "realizandose capacitación y práctica de los usuarios en las aplicaciones que habitualmente operarán. Grupos de no más de 3 participantes."
          ESPE$ = ESPE$ + CRLF$ + CRLF$ + CRLF$ + "Se recomienda realizar este entrenamiento sobre los datos propios previamente migrados y convertidos. "
          If TRIM$(ESPE$) <> "" Then
            Call FRATEXTO(ESPE$, AREF3%)
            For KKU1% = 1 To CARETEX%
              CAITAB1% = CAITAB1% + 1
              TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
            Next KKU1%
            CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
          End If
          '
        End If
        '
        'MIGRACION DE DATOS
17      If Val(Label24) > 0.01 Then
           DESX$ = TRIM$(Str$(ICAPI%)) + "." + Chr$(ILETRA%) + ".- " + "Migracion de datos pre-existentes"
           PREUN$ = Label24
           SUMAPAR# = SUMAPAR# + Val(PREUN$)
           '
           
           '
           Call FRATEXTO(DESX$, AREF2%)
           For KKU1% = 1 To CARETEX%
             CAITAB1% = CAITAB1% + 1
             TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
           Next KKU1%
           TETABU1$(4, CAITAB1%) = PREUN$
           ILETRA% = ILETRA% + 1
           '
           If COMPACTO% = 1 Then GoTo 18 '26/03/07 (OT-07-0132)
           '
           CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
           ILETRA% = Asc("a")
           For U& = 1 To CMIGRA04.List1.ListCount
             TTXX$ = CMIGRA04.List1.List(U& - 1)
             CODE$ = TRIM$(Left$(TTXX$, 16))
             CIXI% = CODINT%(CODE$)
             DESX$ = TRIM$(Str$(ICAPI%)) + "." + Chr$(ILETRA%) + ".- " + DESCRIT0$(CIXI%)
             '
             Call FRATEXTO(DESX$, AREF2%)
             For KKU1% = 1 To CARETEX%
               CAITAB1% = CAITAB1% + 1
               TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
             Next KKU1%
             ILETRA% = ILETRA% + 1
             '
           Next U&
           CAITAB1% = CAITAB1% + 1
           '
        End If
        '
        'FORMULARIOS ELECTRONICOS
        If Val(Text30) > 0 Then
          DESX$ = TRIM$(Str$(ICAPI%)) + "." + Chr$(ILETRA%) + ".- " + "Diseño y Configuración de Formularios Electronicos"
          PREUN$ = Text30
          '
          Call FRATEXTO(DESX$, AREF2%)
          For KKU1% = 1 To CARETEX%
            CAITAB1% = CAITAB1% + 1
            TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
          Next KKU1%
          TETABU1$(4, CAITAB1%) = PREUN$
          ILETRA% = ILETRA% + 1
          SUMAPAR# = SUMAPAR# + Val(PREUN$)
          '
          If COMPACTO% = 1 Then GoTo 18 '26/03/07 (OT-07-0132)
          '
          CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
          ESPE$ = "Comprende el Diseño y Generación de los archivos controladores de proceso correspondientes a los formularios electrónicos asociados a los módulos licenciados."
          ESPE$ = ESPE$ + CRLF$ + CRLF$ + "El diseño se realizará a vuestra opción en base a vuestras muestras que deberán remitirse en formato impreso o por fax a nuestras oficinas, o bien en base a nuestra biblioteca de formatos preconfigurados."
          ESPE$ = ESPE$ + CRLF$ + CRLF$ + "Comprende en todos los casos la incrustación de vuestro logotipo a partir de vuestros archivos gráficos en formato 'JPEG' o 'BMP'."
          '
          If TRIM$(ESPE$) <> "" Then
            Call FRATEXTO(ESPE$, AREF3%)
            For KKU1% = 1 To CARETEX%
              CAITAB1% = CAITAB1% + 1
              TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
            Next KKU1%
            CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
          End If
          '
        End If
        '
        'SERVICIOS DE CONSULTORÍA
18      If Val(Label28) > 0 Then
          DESX$ = TRIM$(Str$(ICAPI%)) + "." + Chr$(ILETRA%) + ".- " + "Servicios de Consultoría"
          PREUN$ = Label28
          '
          Call FRATEXTO(DESX$, AREF2%)
          For KKU1% = 1 To CARETEX%
            CAITAB1% = CAITAB1% + 1
            TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
          Next KKU1%
          TETABU1$(4, CAITAB1%) = PREUN$
          ILETRA% = ILETRA% + 1
          SUMAPAR# = SUMAPAR# + Val(PREUN$)
          '
          If COMPACTO% = 1 Then GoTo 19 '26/03/07 (OT-07-0132)
          '
          CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
          ESPE$ = "Comprende un total de " + Text31 + " horas de servicios de consultoría en vuestra sede, en bloques de 2,5 horas como mínimo cada uno, "
          ESPE$ = ESPE$ + "a realizar por nuestro Ing. Ernesto Pablo Bauer, Profesor Regular de Economía y Organización de la Facultad de Ingeniería - UBA o profesional senior de nivel equivalente. "
          ESPE$ = ESPE$ + "Se brindará asesoramiento en cuestiones de organización, optimización de procesos, codificación y similares, tendientes a lograr la mejor inserción del sistema en la operatoria cotidiana de la organización."
          If TRIM$(ESPE$) <> "" Then
            Call FRATEXTO(ESPE$, AREF3%)
            For KKU1% = 1 To CARETEX%
              CAITAB1% = CAITAB1% + 1
              TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
            Next KKU1%
            CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
          End If
          '
        End If
        '
19      CAITAB1% = CAITAB1% + 1
        If COMPACTO% = 1 Then CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = "Subtotal Implementación"
        TETABU1$(5, CAITAB1%) = FORMATNUM$(SUMAPAR#, "F12.2")
        If COMPACTO% = 1 Then
             CAITAB1% = CAITAB1% + 2
           Else
             CAITAB1% = CAITAB1% + 3
        End If
        '
      End If
      '
      'AQUI RESUMEN DE COTIZACION
      CAITAB1% = CAITAB1% + 1
      ICAPI% = ICAPI% + 1
      TETABU1$(1, CAITAB1%) = TRIM$(Str$(ICAPI%)) + ".- Resumen de Cotizacion"
      CAITAB1% = CAITAB1% + 1
      '
      If (XVALO(Label42) + XVALO(Label43)) >= 0.005 Then
        CAITAB1% = CAITAB1% + 1
        TETABU1$(2, CAITAB1%) = "Licenciamiento Sistemas"
        TETABU1$(4, CAITAB1%) = TRIM$(FORMATNUM$(XVALO(Label42) + XVALO(Label43), "F12.2"))
        If Val(Text26) >= 0.005 Then
          CAITAB1% = CAITAB1% + 1
          TETABU1$(2, CAITAB1%) = "Bonificación sobre Licencias"
          TETABU1$(4, CAITAB1%) = Text26
        End If
        TETABU1$(5, CAITAB1%) = FORMATNUM$(XVALO(Label42) + XVALO(Label43) - XVALO(Text26), "F11.2")
        If COMPACTO% < 1 Then CAITAB1% = CAITAB1% + 1
      End If
      '
      If Val(Label44) >= 0.005 Then
        CAITAB1% = CAITAB1% + 1
        TETABU1$(2, CAITAB1%) = "Desarrollos Especiales"
        TETABU1$(4, CAITAB1%) = Label44
        TETABU1$(5, CAITAB1%) = Label44
        If COMPACTO% < 1 Then CAITAB1% = CAITAB1% + 1
      End If
      '
      If Val(Label70) >= 0.005 Then
        CAITAB1% = CAITAB1% + 1
        TETABU1$(2, CAITAB1%) = "Estimado Implementación"
        TETABU1$(4, CAITAB1%) = Label70
        If Val(Text19) >= 0.005 Then
          CAITAB1% = CAITAB1% + 1
          TETABU1$(2, CAITAB1%) = "Bonificación sobre Implementación"
          TETABU1$(4, CAITAB1%) = Text19
        End If
        TETABU1$(5, CAITAB1%) = FORMATNUM$(Val(Label70) - Val(Text19), "F11.2")
      End If
      '
      CAITAB1% = CAITAB1% + 1
      TETABU1$(5, CAITAB1%) = "=================="
      CAITAB1% = CAITAB1% + 1
      TETABU1$(2, CAITAB1%) = "Sub-Total"
      TETABU1$(5, CAITAB1%) = FORMATNUM$(Val(Label52) - Val(Text26) + Val(Label70) - Val(Text19), "F11.2")
      '\\\OT-06-0145-WAR-18/04/06///
      If COMPACTO% < 1 Then CAITAB1% = CAITAB1% + 1
      If Val(Text28) > 0.005 Then 'SI HAY BONFIC. VARIAS
        CAITAB1% = CAITAB1% + 1
        TETABU1$(2, CAITAB1%) = "Bonificaciones Varias"
        TETABU1$(5, CAITAB1%) = FORMATNUM$(Val(Text28), "F11.2")
        '
        CAITAB1% = CAITAB1% + 1
        TETABU1$(5, CAITAB1%) = "=================="
        CAITAB1% = CAITAB1% + 1
        TETABU1$(2, CAITAB1%) = "Total Neto"
        TETABU1$(5, CAITAB1%) = FORMATNUM$(Val(Label49), "F11.2")
        '
        If COMPACTO% < 1 Then CAITAB1% = CAITAB1% + 1
      End If
      If Val(Label54) > 0.005 Then 'si tiene IVA TASA 1aria.
        CAITAB1% = CAITAB1% + 1
        TETABU1$(2, CAITAB1%) = "IVA " + FORMATNUM$(Val(PORIVA(0)), "F5.2") + " %"
        TETABU1$(5, CAITAB1%) = FORMATNUM$(Val(Label54), "F11.2")
      End If
      If Val(Label71) > 0.005 Then 'si tiene IVA TASA 2daria.
        CAITAB1% = CAITAB1% + 1
        TETABU1$(2, CAITAB1%) = "IVA " + FORMATNUM$(Val(PORIVA(1)), "F5.2") + " %"
        TETABU1$(5, CAITAB1%) = FORMATNUM$(Val(Label71), "F11.2")
      End If
      '
      CAITAB1% = CAITAB1% + 1
      TETABU1$(5, CAITAB1%) = "=================="
      CAITAB1% = CAITAB1% + 1
      TETABU1$(2, CAITAB1%) = "Total General"
      TETABU1$(5, CAITAB1%) = FORMATNUM$(Val(Label77), "F11.2")
      If COMPACTO% < 1 Then CAITAB1% = CAITAB1% + 1
      '
      If COMPACTO% = 1 Then GoTo 20 '26/03/07 (OT-07-0132)
      '
      DESX$ = "(Corresponde a " + Text40 + ")"
      Call FRATEXTO(DESX$, AREF2%)
      For KKU1% = 1 To CARETEX%
        CAITAB1% = CAITAB1% + 1
        TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
      Next KKU1%
      CAITAB1% = CAITAB1% + 3 ' PARA RENGLON EN BLANCO
      '
      '\\\OT-06-0145-WAR-FIN///
20    'MANTENIMIENTO Y SOPORTE
      If Check1.Value = 1 Then
         CAITAB1% = CAITAB1% + 3
         GoTo 109
      End If
      '
      If Val(Label52) > 0.01 Then
        CAITAB1% = CAITAB1% + 1
        If COMPACTO% = 1 Then CAITAB1% = CAITAB1% + 1
        ICAPI% = ICAPI% + 1
        TETABU1$(1, CAITAB1%) = TRIM$(Str$(ICAPI%)) + ".- Mantenimiento / Soporte / Mesa de Ayuda"
        CAITAB1% = CAITAB1% + 1
        '
        If COMPACTO% = 1 Then GoTo 21 '26/03/07 (OT-07-0132)
        '
        DESX$ = "Comprende los servicios de Mantenimiento de los sistemas licenciados, según el detalle correspondiente a la opción elegida a continuación, entendiendose como parte integrante de la presente propuesta, y por el término de 18 meses a partir de la instalación."
        Call FRATEXTO(DESX$, AREF2%)
        For KKU1% = 1 To CARETEX%
          CAITAB1% = CAITAB1% + 1
          TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
        Next KKU1%
        CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
        '
21      ESPE$ = TRIM$(Str$(ICAPI%)) + ".1.- Servicio Basico:"
        '26/03/07 (OT-07-0132)
        If COMPACTO% = 1 Then ESPE$ = ESPE$ + " $ " + TRIM$(Label65(0)) + "   más I.V.A."
        '26/03/07 (FIN)
        Call FRATEXTO(ESPE$, AREF2%)
        For KKU1% = 1 To CARETEX%
          CAITAB1% = CAITAB1% + 1
          TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
        Next KKU1%
        If COMPACTO% = 1 Then GoTo 22 '26/03/07 (OT-07-0132)
        '
        CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
        '
        ESPE$ = "Comprende:"
        ESPE$ = ESPE$ + CRLF$ + CRLF$ + "Servicio de soporte telefonico de usuarios, sin límite, complementando asistencia verbal con reportes instructivos impresos remitidos por e-mail en formato PDF."
        ESPE$ = ESPE$ + CRLF$ + CRLF$ + "Actualización permanente en correspondencia con cambios de legislación y/o reglamentación pública vigente."
        If TRIM$(ESPE$) <> "" Then
          Call FRATEXTO(ESPE$, AREF3%)
          For KKU1% = 1 To CARETEX%
            CAITAB1% = CAITAB1% + 1
            TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
          Next KKU1%
          CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
        End If
        '
        DESX$ = "Cuota mensual (mes adelantado): $ " + TRIM$(Label65(0)) + "   más I.V.A."
        Call FRATEXTO(DESX$, AREF2%)
        For KKU1% = 1 To CARETEX%
          CAITAB1% = CAITAB1% + 1
          TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
        Next KKU1%
        CAITAB1% = CAITAB1% + 2 ' PARA RENGLON EN BLANCO
        '
22      ESPE$ = TRIM$(Str$(ICAPI%)) + ".2.- Servicio Kaizen:"
        '26/03/07 (OT-07-0132)
        If COMPACTO% = 1 Then ESPE$ = ESPE$ + " $ " + TRIM$(Label65(1)) + "   más I.V.A."
        '26/03/07 (FIN)
        Call FRATEXTO(ESPE$, AREF2%)
        For KKU1% = 1 To CARETEX%
          CAITAB1% = CAITAB1% + 1
          TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
        Next KKU1%
        If COMPACTO% = 1 Then GoTo 23 '26/03/07 (OT-07-0132)
        '
        CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
        '
        ESPE$ = "Comprende el Servicio Básico más:"
        ESPE$ = ESPE$ + CRLF$ + CRLF$ + "Actualizaciones generales en base al principio de la 'Mejora Contínua' o Kaizen."
        ESPE$ = ESPE$ + CRLF$ + CRLF$ + "Servicio de recuperación de datos en caso de necesidad (sujeto a posibilidad técnica)."
        '
        If TRIM$(ESPE$) <> "" Then
          Call FRATEXTO(ESPE$, AREF3%)
          For KKU1% = 1 To CARETEX%
            CAITAB1% = CAITAB1% + 1
            TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
          Next KKU1%
          CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
        End If
        '
        DESX$ = "Cuota mensual (mes adelantado): $ " + TRIM$(Label65(1)) + "   más I.V.A."
        Call FRATEXTO(DESX$, AREF2%)
        For KKU1% = 1 To CARETEX%
          CAITAB1% = CAITAB1% + 1
          TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
        Next KKU1%
        CAITAB1% = CAITAB1% + 2 ' PARA RENGLON EN BLANCO
        '
23      ESPE$ = TRIM$(Str$(ICAPI%)) + ".3.- Servicio Platinum:"
        '26/03/07 (OT-07-0132)
        If COMPACTO% = 1 Then ESPE$ = ESPE$ + " $ " + TRIM$(Label65(2)) + "   más I.V.A."
        '26/03/07 (FIN)
        Call FRATEXTO(ESPE$, AREF2%)
        For KKU1% = 1 To CARETEX%
          CAITAB1% = CAITAB1% + 1
          TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
        Next KKU1%
        If COMPACTO% = 1 Then GoTo 24 '26/03/07 (OT-07-0132)
        '
        CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
        '
        ESPE$ = "Comprende el Servicio Kaizen más:"
        ESPE$ = ESPE$ + CRLF$ + CRLF$ + "Atención domiciliaria 'on demand' en caso de necesidad (solo CABA y GBA)."
        ESPE$ = ESPE$ + CRLF$ + CRLF$ + "Servicios de re-configuración de formularios y controladores de proceso sin límite."
        ESPE$ = ESPE$ + CRLF$ + CRLF$ + "Desarrollo de nuevas consultas y listadores sobre especificación del usuario (con limitaciones)."
        ESPE$ = ESPE$ + CRLF$ + CRLF$ + "Módulo Generador de Reportes en consignación (sin cargo) permitiendo desarrollo de sus propios reportes tabulares y gráficos, incluyendo servicio de soporte telefónico."
        '
        If TRIM$(ESPE$) <> "" Then
          Call FRATEXTO(ESPE$, AREF3%)
          For KKU1% = 1 To CARETEX%
            CAITAB1% = CAITAB1% + 1
            TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
          Next KKU1%
          CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
        End If
        '
        DESX$ = "Cuota mensual (mes adelantado): $ " + TRIM$(Label65(2)) + "   más I.V.A."
        Call FRATEXTO(DESX$, AREF2%)
        For KKU1% = 1 To CARETEX%
          CAITAB1% = CAITAB1% + 1
          TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
        Next KKU1%
        CAITAB1% = CAITAB1% + 2 ' PARA RENGLON EN BLANCO
        '
        DESX$ = "Nivel de Servicio elegido: ...................................................................................."
        DESX$ = Left$(DESX$, AREF2%)
        CAITAB1% = CAITAB1% + 1
        TETABU1$(2, CAITAB1%) = DESX$
24      If COMPACTO% = 1 Then
            CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
          Else
            CAITAB1% = CAITAB1% + 3 ' PARA RENGLON EN BLANCO
        End If '
      End If
      '
      'CONDICIONES COMERCIALES
      CAITAB1% = CAITAB1% + 1
      If COMPACTO% = 1 Then CAITAB1% = CAITAB1% + 1
109   ICAPI% = ICAPI% + 1
      TETABU1$(1, CAITAB1%) = TRIM$(Str$(ICAPI%)) + ".- Condiciones Comerciales y Técnicas:"
      CAITAB1% = CAITAB1% + 1
      '
      ILETRA% = Asc("a")
      SUMAPAR# = 0
      '
      'PRECIOS
      DESX$ = TRIM$(Str$(ICAPI%)) + "." + Chr$(ILETRA%) + ".- " + "Precios:"
      Call FRATEXTO(DESX$, AREF2%)
      For KKU1% = 1 To CARETEX%
        CAITAB1% = CAITAB1% + 1
        TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
      Next KKU1%
      CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
      ILETRA% = ILETRA% + 1
      '
      ESPE$ = Text6
      If TRIM$(ESPE$) <> "" Then
        Call FRATEXTO(ESPE$, AREF3%)
        For KKU1% = 1 To CARETEX%
          CAITAB1% = CAITAB1% + 1
          TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
          If COMPACTO% = 1 Then
            If KKU1% = 3 Then
              CAITAB1% = CAITAB1% + 2
              GoTo 110
            End If
          End If
        Next KKU1%
        CAITAB1% = CAITAB1% + 2 ' PARA RENGLON EN BLANCO
      End If
      '
      'PLAZO DE ENTREGA
110   ESPE$ = Text15
      If TRIM$(ESPE$) <> "" Then
        DESX$ = TRIM$(Str$(ICAPI%)) + "." + Chr$(ILETRA%) + ".- " + "Plazo de Entrega:"
        Call FRATEXTO(DESX$, AREF2%)
        For KKU1% = 1 To CARETEX%
          CAITAB1% = CAITAB1% + 1
          TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
        Next KKU1%
        CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
        ILETRA% = ILETRA% + 1
        '
        If TRIM$(ESPE$) <> "" Then
          Call FRATEXTO(ESPE$, AREF3%)
          For KKU1% = 1 To CARETEX%
            CAITAB1% = CAITAB1% + 1
            TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
            If COMPACTO% = 1 Then
              If KKU1% = 1 Then
                CAITAB1% = CAITAB1% + 2
                GoTo 111
              End If
            End If
          Next KKU1%
          CAITAB1% = CAITAB1% + 2 ' PARA RENGLON EN BLANCO
        End If
      End If
      '
      'FORMA DE PAGO
      
111   ESPE$ = Text16
      If TRIM$(ESPE$) <> "" Then
        DESX$ = TRIM$(Str$(ICAPI%)) + "." + Chr$(ILETRA%) + ".- " + "Forma de Pago:"
        Call FRATEXTO(DESX$, AREF2%)
        For KKU1% = 1 To CARETEX%
          CAITAB1% = CAITAB1% + 1
          TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
        Next KKU1%
        CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
        ILETRA% = ILETRA% + 1
        '
        If COMPACTO% = 1 Then GoTo 112 '26/03/07 (OT-07-0132)
        '
        If TRIM$(ESPE$) <> "" Then
          Call FRATEXTO(ESPE$, AREF3%)
          For KKU1% = 1 To CARETEX%
            CAITAB1% = CAITAB1% + 1
            TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
          Next KKU1%
          CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
        End If
      End If
      '
      ESPE$ = "Cronograma Tentativo de Fecha de Valores a Entregar (dependiendo de fecha efectiva de aceptación y entrega)"
      Call FRATEXTO(ESPE$, AREF3%)
      For KKU1% = 1 To CARETEX%
        CAITAB1% = CAITAB1% + 1
        TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
      Next KKU1%
112   CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
      For KKU1% = 1 To List1.ListCount
         CAITAB1% = CAITAB1% + 1
         TETABU1$(7, CAITAB1%) = List1.List(KKU1% - 1)
         If XVALO(Text34) <= 1 Then
            CAITAB1% = CAITAB1% + 1
            TETABU1$(7, CAITAB1%) = "2  Saldo a convenir."
            GoTo 113
         End If
      Next KKU1%
113   CAITAB1% = CAITAB1% + 2 ' PARA RENGLON EN BLANCO
      '
      If COMPACTO% = 1 Then GoTo 116
      'Garantia
      DESX$ = TRIM$(Str$(ICAPI%)) + "." + Chr$(ILETRA%) + ".- " + "Garantía:"
      Call FRATEXTO(DESX$, AREF2%)
      For KKU1% = 1 To CARETEX%
        CAITAB1% = CAITAB1% + 1
        TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
      Next KKU1%
      CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
      ILETRA% = ILETRA% + 1
      '
      ESPE$ = "Por  el  término  de 18  meses a partir de la entrega, contra todo vicio aparente u oculto "
      ESPE$ = ESPE$ + "que impida  o  dificulte el normal funcionamiento de los sistemas. Se excluyen únicamente "
      ESPE$ = ESPE$ + "fallas atribuibles  al mal  funcionamiento de los equipos,  daños de medios de soporte "
      ESPE$ = ESPE$ + "magnéticos o de otro tipo, funcionamiento  defectuoso  de  software  de  base,  sistemas "
      ESPE$ = ESPE$ + "operativos,   software de red, cableado de alimentación, y otros ajenos al desarrollo aplicativo licenciado."
      ESPE$ = ESPE$ + CRLF$ + CRLF$ + "Las eventuales  prestaciones en garantía consistirán en el suministro sobre medio magnético "
      ESPE$ = ESPE$ + "en versión  autoinstalable  de  los  archivos  correspondientes  a  los módulos ejecutables o controladores "
      ESPE$ = ESPE$ + "de procesos defectuosos. A opción de FLEXOFT dichos archivos correspondientes a los módulos corregidos / reparados "
      ESPE$ = ESPE$ + "podrán ser remitidos via Internet o correo electrónico."
      ESPE$ = ESPE$ + CRLF$ + CRLF$ + "Los servicios de garantía serán prestados en el domicilio de FLEXOFT o de su  Servicio "
      ESPE$ = ESPE$ + "Técnico Autorizado, con excepción de los abonados a Servicio de Mantenimiento nivel 'Platinum' o superior. El suministro incluye los módulos de confección de respaldo de datos, los que el "
      ESPE$ = ESPE$ + "usuario / licenciatario deberá remitir a FLEXOFT  sobre medio removible de soporte o vía Internet. "
      ESPE$ = ESPE$ + "En caso de requerirse  la concurrencia de un técnico a la sede del usuario / licenciatario, se "
      ESPE$ = ESPE$ + "facturarán las horas - hombre de traslado según tarifas vigentes más el kilometraje recorrido (excepto abonados nivel 'Platinum' o superior)."
      Call FRATEXTO(ESPE$, AREF3%)
      For KKU1% = 1 To CARETEX%
        CAITAB1% = CAITAB1% + 1
        TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
      Next KKU1%
      CAITAB1% = CAITAB1% + 2 ' PARA RENGLON EN BLANCO
      '
      'Condicion Técnica
      DESX$ = TRIM$(Str$(ICAPI%)) + "." + Chr$(ILETRA%) + ".- " + "Condiciones Técnicas:"
      Call FRATEXTO(DESX$, AREF2%)
      For KKU1% = 1 To CARETEX%
        CAITAB1% = CAITAB1% + 1
        TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
      Next KKU1%
      CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
      ILETRA% = ILETRA% + 1
      '
      ESPE$ = "Los sistemas cotizados responden a una estructura modular, pudiendo adquirirse  e instalarse "
      ESPE$ = ESPE$ + "en foma parcial. Existen requisitos de pre-existencia entre  las  diversas aplicaciones, "
      ESPE$ = ESPE$ + "que hacen necesario un cierto orden lógico  de  adquisición  e  implementación.  Ante "
      ESPE$ = ESPE$ + "cualquier duda consultar a nuestros representantes técnicos."
      Call FRATEXTO(ESPE$, AREF3%)
      For KKU1% = 1 To CARETEX%
        CAITAB1% = CAITAB1% + 1
        TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
      Next KKU1%
      CAITAB1% = CAITAB1% + 2 ' PARA RENGLON EN BLANCO
      '
      'Suministro
      DESX$ = TRIM$(Str$(ICAPI%)) + "." + Chr$(ILETRA%) + ".- " + "Suministro:"
      Call FRATEXTO(DESX$, AREF2%)
      For KKU1% = 1 To CARETEX%
        CAITAB1% = CAITAB1% + 1
        TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
      Next KKU1%
      CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
      '
      ESPE$ = "El  suministro  se  efectivizará  mediante  otorgamiento  de  licencia de uso por los sistemas aquiridos y la entrega de:"
      ESPE$ = ESPE$ + CRLF$ + CRLF$ + ".   " + TRIM$(Str$(ICAPI%)) + "." + Chr$(ILETRA%) + ".1.- CD  con versión autoinstalable de los módulos licenciados. "
      ESPE$ = ESPE$ + CRLF$ + ".   " + TRIM$(Str$(ICAPI%)) + "." + Chr$(ILETRA%) + ".2.- Un juego de documentación técnica de referencia."
      ESPE$ = ESPE$ + CRLF$ + ".   " + TRIM$(Str$(ICAPI%)) + "." + Chr$(ILETRA%) + ".3.- Diskettes 'llave' (1 por puesto de trabajo habilitado)."
      Call FRATEXTO(ESPE$, AREF3%)
      For KKU1% = 1 To CARETEX%
        CAITAB1% = CAITAB1% + 1
        TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
      Next KKU1%
      ILETRA% = ILETRA% + 1
      CAITAB1% = CAITAB1% + 2 ' PARA RENGLON EN BLANCO
      '
      'Suministro
      ICAPI% = ICAPI% + 1
      DESX$ = TRIM$(Str$(ICAPI%)) + "." + Chr$(ILETRA%) + ".- " + "Propiedad Intelectual:"
      Call FRATEXTO(DESX$, AREF2%)
      For KKU1% = 1 To CARETEX%
        CAITAB1% = CAITAB1% + 1
        TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
      Next KKU1%
      CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
      '
      ESPE$ = "Queda  en  cabeza  de  FLEXOFT,  quien  podrá disponer libremente de los sistemas salvo convenio explícito en contrario. "
      ESPE$ = ESPE$ + "El usuario / licenciatario solo podrá confeccionar copias destinadas a resguardo  de  los  sistemas,  y  deberá "
      ESPE$ = ESPE$ + "abstenerse  de  suministrarlas a terceros sin previa y explícita  autorización  de FLEXOFT.  Lo  mismo  se  aplicará "
      ESPE$ = ESPE$ + "a la documentación técnica suministrada y a toda información acerca de los sistemas aplicativos.  Asimismo el "
      ESPE$ = ESPE$ + "usuario / licenciatario  se deberá abstener de copiar, modificar o reproducir total o parcialmente los sistemas "
      ESPE$ = ESPE$ + "mediante prácticas de ingeniería inversa."
      Call FRATEXTO(ESPE$, AREF3%)
      For KKU1% = 1 To CARETEX%
        CAITAB1% = CAITAB1% + 1
        TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
      Next KKU1%
      ILETRA% = ILETRA% + 1
      CAITAB1% = CAITAB1% + 2 ' PARA RENGLON EN BLANCO
      '
      'Otras Condiciones
      ESPE$ = Text29
      If TRIM$(ESPE$) <> "" Then
        DESX$ = TRIM$(Str$(ICAPI%)) + "." + Chr$(ILETRA%) + ".- " + "Otras Condiciones:"
        Call FRATEXTO(DESX$, AREF2%)
        For KKU1% = 1 To CARETEX%
          CAITAB1% = CAITAB1% + 1
          TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
        Next KKU1%
        CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
        ILETRA% = ILETRA% + 1
        '
        If TRIM$(ESPE$) <> "" Then
          Call FRATEXTO(ESPE$, AREF3%)
          For KKU1% = 1 To CARETEX%
            CAITAB1% = CAITAB1% + 1
            TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
          Next KKU1%
          CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
        End If
      End If
      '
      'VALIDEZ DE OFERTA
116   ESPE$ = Text32
      If TRIM$(ESPE$) <> "" Then
        DESX$ = TRIM$(Str$(ICAPI%)) + "." + Chr$(ILETRA%) + ".- " + "Validez de Oferta:"
        Call FRATEXTO(DESX$, AREF2%)
        For KKU1% = 1 To CARETEX%
          CAITAB1% = CAITAB1% + 1
          TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
        Next KKU1%
        CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
        ILETRA% = ILETRA% + 1
        '
        If TRIM$(ESPE$) <> "" Then
          Call FRATEXTO(ESPE$, AREF3%)
          For KKU1% = 1 To CARETEX%
            CAITAB1% = CAITAB1% + 1
            TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
          Next KKU1%
          CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
        End If
      End If
      If COMPACTO% = 1 Then GoTo 36
      '
      'aceptacion
      ESPE$ = "Mediante firma al pie de la presente propuesta en señal de conformidad."
      If TRIM$(ESPE$) <> "" Then
        DESX$ = TRIM$(Str$(ICAPI%)) + "." + Chr$(ILETRA%) + ".- " + "Aceptación:"
        Call FRATEXTO(DESX$, AREF2%)
        For KKU1% = 1 To CARETEX%
          CAITAB1% = CAITAB1% + 1
          TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
        Next KKU1%
        CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
        ILETRA% = ILETRA% + 1
        '
        If TRIM$(ESPE$) <> "" Then
          Call FRATEXTO(ESPE$, AREF3%)
          For KKU1% = 1 To CARETEX%
            CAITAB1% = CAITAB1% + 1
            TETABU1$(3, CAITAB1%) = RETEX$(KKU1%)
          Next KKU1%
          CAITAB1% = CAITAB1% + 1 ' PARA RENGLON EN BLANCO
        End If
      End If
      '

36    'NOCONFIRM% = 1
      '
      Call PRINTDOC("DETAPRE") ' DETALLE DE PRESUPUESTO
      '
      'LIMPIA LA LISTA
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
    End If
    '
199 Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
End Sub
'
'\\\OT-05-0394-WAR-FIN///
Private Sub VScroll1_Change()
   Text4 = 100 - VScroll1.Value
End Sub

Sub CARGODETLICEN() 'CARGO LISTA DE DETALLE DE UN PRESU
   '
   If TRIM$(Text1) <> "" Then
     NUPRESU$ = TRIM$(Text1)
     Call ABREPRECAM
     SQLX$ = "SELECT * FROM DETAPRESU "
     SQLX$ = SQLX$ + "WHERE TRIM(Nro_presu) = '" & NUPRESU$ & "' "
     SQLX$ = SQLX$ + "AND rubro_det = 1 "
     SQLX$ = SQLX$ + "ORDER BY CodExte_Det ASC "
     Set DETPRETMP = PresuCam.OpenRecordset(SQLX$, 4)
     J& = 0
     '
     With DETPRETMP
       TOTPIS1# = 0: TOTPIS2# = 0
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
        On Error GoTo 0
        CLICEN04.List1.Clear
        TOTPIS# = 0
        Do Until .EOF = True
          CODI1$ = CODEXT(!CODINTE_DET)
          DESCRI1$ = DESCRIT0(!CODINTE_DET)
          PREUNI1# = !PreUni_Det
          CANTI1# = Val(REPRFLEX04.Text4)   ' !Cantid_Det
          PREUNI2$ = FORMATNUM(PREUNI1#, "F12.2")
          CANTI2$ = FORMATNUM(CANTI1#, "F7.1")
          Y$ = Space$(256)
          Call REPLA(Y$, CODI1$, 1, 16)
          Call REPLA(Y$, DESCRI1$, 19, 45)
          Call REPLA(Y$, PREUNI2$, 65, 12)
          'Call REPLA(Y$, CANTI2$, 78, 7)
          IMPOIT1# = PREUNI1# * (1 + Val(REPRFLEX04.Label45) / 100)
15        IMPOIT2$ = FORMATNUM(IMPOIT1#, "F12.2")
          Call REPLA(Y$, IMPOIT2$, 79, 12)
          CLICEN04.List1.AddItem Y$
          TOTPIS1# = TOTPIS1# + PREUNI1#
          TOTPIS2# = TOTPIS2# + IMPOIT1#
          .MoveNext
        Loop
       End If
      End With
    End If
    'CALCULA EL TOTAL DE LA ACCESORIOS.
    CLICEN04.Text13 = FORMATNUM(TOTPIS1#, "F12.2")
    CLICEN04.Text38 = FORMATNUM(TOTPIS2#, "F12.2")
    '
End Sub
'
Sub CALCUDETLICEN()
    '
    TOTPIS1# = 0: TOTPIS2# = 0
    For U& = 1 To CLICEN04.List1.ListCount
      Y$ = CLICEN04.List1.List(U& - 1)
      PREUNI1# = Val(Mid$(Y$, 65, 12))
      CANTI1# = Val(REPRFLEX04.Text4)   ' !Cantid_Det
      'CANTI2$ = FORMATNUM(CANTI1#, "F7.1")
      'Call REPLA(Y$, CANTI2$, 78, 7)
      IMPOIT1# = PREUNI1# * (1 + Val(REPRFLEX04.Label45) / 100)
15    IMPOIT2$ = FORMATNUM(IMPOIT1#, "F12.2")
      Call REPLA(Y$, IMPOIT2$, 79, 12)
      CLICEN04.List1.RemoveItem U& - 1
      CLICEN04.List1.AddItem Y$, U& - 1
      TOTPIS1# = TOTPIS1# + PREUNI1#
      TOTPIS2# = TOTPIS2# + IMPOIT1#
    Next U&
    '
    CLICEN04.Text13 = FORMATNUM(TOTPIS1#, "F12.2")
    CLICEN04.Text38 = FORMATNUM(TOTPIS2#, "F12.2")
    '
End Sub

Sub CALCUOTAS()
    '
    PLAENT = Val(Text39)    ' PLAZO DE ENTREGA
    INCIVA% = 0
    If PLAENT > 30 Then
      PLAENT = 30
      INCIVA% = 1
    End If
    List1.Clear
    ineto# = XVALO(Label52) - XVALO(Text26)
    IMIVA# = ineto# * (XVALO(PORIVA(0))) / 100 ' Val(Label54) + Val(Label71) + Val(Label76)
    If ineto# >= 1 Then PORIVX = 1 + IMIVA# / ineto#
    ANTICI# = ineto# * Val(Text33) / 100
    CANCUO = Val(Text34)
      If CANCUO < 1 Then CANCUO = 1
    VALOCUO# = (ineto# - ANTICI#) * (1 + CANCUO * Val(Text35) / 100) / CANCUO
    If CANCUO < 2 Then
        VALOCUO# = VALOCUO# + IMIVA#
        IMIVA# = 0
      ElseIf INCIVA% = 1 Then
        ANTICI# = ANTICI# * PORIVX
        VALOCUO# = VALOCUO# * PORIVX
        IMIVA# = 0
    End If
    TTXX$ = Space$(60)
    HOII% = HOY(HOS$) '
    IIT% = 0
    '
    If ineto# < 0.01 Then GoTo 80
    '
    'ANTICIPO
    If ANTICI# >= 0.005 Then
      Call REPLA(TTXX$, "    Anticipo", 1, 19)
      IIT% = IIT% + 1
      Call REPLA(TTXX$, FORMATNUM$(IIT%, "I2"), 1, 2)
      Call REPLA(TTXX$, FORMATNUM$(ANTICI#, "F11.2"), 20, 11)
      Call REPLA(TTXX$, HOS$, 33, 8)
      List1.AddItem TTXX$
    End If
    '
    HOIJ = HOII
    FEPRICUO = HOII%
    While DIENTRE(HOIJ, FEPRICUO) < PLAENT
      FEPRICUO = DIPOST(FEPRICUO)
    Wend
    FENTREGA = FEPRICUO
    '
    'IMPORTE IVA
    CUIVA$ = ""
    If IMIVA# >= 0.005 Then
      TTXX$ = Space$(60)
      Call REPLA(TTXX$, "    I.V.A.", 1, 19)
      Call REPLA(TTXX$, FORMATNUM$(IMIVA#, "F11.2"), 20, 11)
      FEPCU$ = FECHATEX$(FEPRICUO)
      MACT% = Val(Mid$(FEPCU$, 4, 2))
      FEX = FEPRICUO
      While Val(Mid$(FECHATEX$(FEX), 4, 2)) = MACT%
        FEX = DIPOST(FEX)
      Wend
      While Val(Left$(FECHATEX$(FEX), 2)) < 8
        FEX = DIPOST(FEX)
      Wend
      Call REPLA(TTXX$, FECHATEX$(FEX), 33, 8)
      FIVA = FEX
      CUIVA$ = TTXX$
    End If
    '
    FEPRICUO = FENTREGA
    While DIENTRE(FENTREGA, FEPRICUO) < XVALO(Text42)
       FEPRICUO = DIPOST(FEPRICUO)
    Wend
    For KU% = 1 To CANCUO
      If CUIVA$ <> "" Then
        If FIVA <= FEPRICUO Then
          IIT% = IIT% + 1
          Call REPLA(CUIVA$, FORMATNUM$(IIT%, "I2"), 1, 2)
          List1.AddItem CUIVA$
          CUIVA$ = ""
        End If
      End If
      '
      IIT% = IIT% + 1
      Call REPLA(TTXX$, FORMATNUM$(IIT%, "I2"), 1, 2)
      Call REPLA(TTXX$, "Cuota Nro " + TRIM$(Str$(KU%)), 5, 19)
      Call REPLA(TTXX$, FORMATNUM$(VALOCUO#, "F11.2"), 20, 11)
      Call REPLA(TTXX$, FECHATEX$(FEPRICUO), 33, 8)
      List1.AddItem TTXX$
      FULCUO = FEPRICUO
      While DIENTRE(FULCUO, FEPRICUO) < 30
        FEPRICUO = DIPOST(FEPRICUO)
      Wend
    Next KU%
    '
    If CUIVA$ <> "" Then
      IIT% = IIT% + 1
      Call REPLA(CUIVA$, FORMATNUM$(IIT%, "I2"), 1, 2)
      List1.AddItem CUIVA$
    End If
    '
80  TTXX$ = ""
    TTXX$ = "Anticipo " + TRIM$(Text33) + " % a la aprobación: $ " + TRIM$(FORMATNUM$(ANTICI#, "F11.2"))
    TTXX$ = TTXX$ + CRLF$ + CRLF$
    If XVALO(Text34) <= 1 Then
       TTXX$ = TTXX$ + "Saldo a convenir."
       GoTo 95
    End If
    '
    If IMIVA# >= 0.005 Then
      TTXX$ = TTXX$ + "Importe correspondiente al I.V.A. en un único pago efectivizado mediante cheque propio de pago diferido al " + FECHATEX$(FIVA) + " contra entrega / instalación: $ " + TRIM$(FORMATNUM$(IMIVA#, "F11.2"))
      TTXX$ = TTXX$ + CRLF$ + CRLF$
    End If
    TTXX$ = TTXX$ + "Saldo en " + TRIM$(Text34) + " cuota(s) mensual(es) igual(es) y consecutiva(s), cada una de $ " + TRIM$(FORMATNUM$(VALOCUO#, "F11.2")) + ", "
    FEPRICUO = FENTREGA
    While DIENTRE(FENTREGA, FEPRICUO) < XVALO(Text42)
       FEPRICUO = DIPOST(FEPRICUO)
    Wend
    If Val(Text35) < 0.05 Then
       TTXX$ = TTXX$ + "sin interés"
      Else
       TTXX$ = TTXX$ + "incluyendo interés del " + TRIM$(Text35) + " % mensual directo"
    End If
    TTXX$ = TTXX$ + ", venciendo la primera "
    If XVALO(Text42) > 0 Then
       TTXX$ = TTXX$ + "a los " + TRIM$(Text42) + " días de la entrega, o sea "
    End If
    TTXX$ = TTXX$ + "el " + FECHATEX$(FEPRICUO) + " y las siguientes cada 30 días a partir de dicha fecha. Las cuotas se instrumentarán mediante cheques propios de pago diferido, debiendo entregarse dichos valores contra entrega y presentación de nuestra factura correspondiente."
95  Text16 = TTXX$
End Sub

Private Sub VScroll2_Change()
   Text39 = 366 - VScroll2.Value
End Sub

Sub CONECTADOC()
   '
   If TRIM$(Label85) <> "" Then
     XX$ = TRIM$(Label86) + "\" + TRIM$(Label85)
     For U& = 1 To List2.ListCount
       If List2.List(U& - 1) = XX$ Then Exit Sub
     Next U&
     List2.AddItem XX$
     List2.Refresh
     DoEvents
   End If
   '
End Sub

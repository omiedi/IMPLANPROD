VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form ENSAMOT1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ensayo de Motores"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   Icon            =   "ENSAMOT1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8310
      Left            =   0
      TabIndex        =   1
      Top             =   105
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   14658
      _Version        =   327681
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Archivo Importación"
      TabPicture(0)   =   "ENSAMOT1.frx":1272
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame3(2)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tabla Valores"
      TabPicture(1)   =   "ENSAMOT1.frx":128E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(0)"
      Tab(1).Control(1)=   "Frame1(1)"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Grafico Inversión"
      TabPicture(2)   =   "ENSAMOT1.frx":12AA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSChart1(0)"
      Tab(2).Control(1)=   "MSChart1(1)"
      Tab(2).Control(2)=   "Label2(1)"
      Tab(2).Control(3)=   "Label2(0)"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "Grafico Arranque"
      TabPicture(3)   =   "ENSAMOT1.frx":12C6
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label2(2)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label2(3)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "MSChart2(1)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "MSChart2(0)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Cupla - rpm"
      TabPicture(4)   =   "ENSAMOT1.frx":12E2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "MSChart3"
      Tab(4).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "Encoder"
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
         Index           =   2
         Left            =   -67020
         TabIndex        =   34
         Top             =   735
         Width           =   2745
         Begin VB.TextBox Text2 
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
            Left            =   1260
            TabIndex        =   39
            Text            =   "11.71875"
            Top             =   945
            Width           =   1275
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
            Left            =   1260
            TabIndex        =   37
            Text            =   "0.005"
            Top             =   525
            Width           =   855
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Seg"
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
            Index           =   3
            Left            =   2205
            TabIndex        =   38
            Top             =   570
            Width           =   810
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Intervalo:"
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
            Index           =   2
            Left            =   420
            TabIndex        =   36
            Top             =   570
            Width           =   810
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Constante:"
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
            Index           =   2
            Left            =   105
            TabIndex        =   35
            Top             =   990
            Width           =   1020
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ensayo de Baja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3585
         Index           =   1
         Left            =   -74630
         TabIndex        =   31
         Top             =   4410
         Width           =   10305
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
            Height          =   2760
            Index           =   1
            Left            =   210
            TabIndex        =   32
            Top             =   650
            Width           =   9885
         End
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "   Ord   Puls        Seg        rpm(i)       rpm(p)     Cupla(i)     Cupla(p)"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   210
            TabIndex        =   33
            Top             =   410
            Width           =   9885
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Ensayo de Alta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3585
         Index           =   0
         Left            =   -74630
         TabIndex        =   28
         Top             =   630
         Width           =   10305
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
            Height          =   2760
            Index           =   0
            Left            =   210
            TabIndex        =   29
            Top             =   650
            Width           =   9885
         End
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "   Ord   Puls        Seg        rpm(i)       rpm(p)     Cupla(i)     Cupla(p)"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   210
            TabIndex        =   30
            Top             =   410
            Width           =   9885
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Archivos de Importación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3165
         Left            =   -74790
         TabIndex        =   5
         Top             =   3255
         Width           =   10515
         Begin VB.CommandButton Command2 
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
            Height          =   750
            Left            =   9450
            Picture         =   "ENSAMOT1.frx":12FE
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Cierra y Abandona la Aplicación"
            Top             =   1155
            Width           =   855
         End
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   9430
            ScaleHeight     =   75
            ScaleWidth      =   825
            TabIndex        =   25
            Top             =   2100
            Width           =   885
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   26
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.PictureBox Picture1 
            Height          =   600
            Left            =   9250
            ScaleHeight     =   540
            ScaleWidth      =   1170
            TabIndex        =   24
            Top             =   2360
            Width           =   1230
            Begin VB.Line Line2 
               X1              =   1160
               X2              =   1160
               Y1              =   630
               Y2              =   0
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   1155
               Y1              =   525
               Y2              =   525
            End
            Begin VB.Image Image2 
               Height          =   510
               Left            =   -400
               Picture         =   "ENSAMOT1.frx":1608
               ToolTipText     =   "Doble Click para Configurar Modalidad de Cálculo"
               Top             =   0
               Width           =   2010
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Ensayo de Baja"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Index           =   1
            Left            =   4830
            TabIndex        =   15
            Top             =   420
            Width           =   4425
            Begin VB.OptionButton Option2 
               Caption         =   "TXT (delimitado por CR-LF)"
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
               Left            =   1050
               TabIndex        =   18
               Top             =   1995
               Width           =   3000
            End
            Begin VB.OptionButton Option1 
               Caption         =   "CSV (delimitado por comas)"
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
               Left            =   1050
               TabIndex        =   17
               Top             =   1680
               Value           =   -1  'True
               Width           =   3060
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
               Height          =   330
               Index           =   1
               Left            =   3990
               TabIndex        =   16
               Top             =   525
               Width           =   175
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
               Index           =   1
               Left            =   105
               TabIndex        =   23
               Top             =   870
               Width           =   810
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
               Height          =   580
               Index           =   1
               Left            =   1050
               TabIndex        =   22
               Top             =   885
               Width           =   3105
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
               Index           =   1
               Left            =   105
               TabIndex        =   21
               Top             =   570
               Width           =   810
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
               Height          =   300
               Index           =   1
               Left            =   1050
               TabIndex        =   20
               Top             =   525
               Width           =   2940
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Formato:"
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
               Index           =   1
               Left            =   105
               TabIndex        =   19
               Top             =   1680
               Width           =   810
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "Ensayo de Alta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Index           =   0
            Left            =   210
            TabIndex        =   6
            Top             =   420
            Width           =   4425
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
               Height          =   330
               Index           =   0
               Left            =   3990
               TabIndex        =   14
               Top             =   525
               Width           =   175
            End
            Begin VB.OptionButton Option1 
               Caption         =   "CSV (delimitado por comas)"
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
               Left            =   1050
               TabIndex        =   12
               Top             =   1680
               Value           =   -1  'True
               Width           =   3060
            End
            Begin VB.OptionButton Option2 
               Caption         =   "TXT (delimitado por CR-LF)"
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
               Left            =   1050
               TabIndex        =   11
               Top             =   1995
               Width           =   3000
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Formato:"
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
               Index           =   0
               Left            =   105
               TabIndex        =   13
               Top             =   1680
               Width           =   810
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
               Height          =   300
               Index           =   0
               Left            =   1050
               TabIndex        =   10
               Top             =   525
               Width           =   2940
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
               Index           =   0
               Left            =   105
               TabIndex        =   9
               Top             =   570
               Width           =   810
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
               Height          =   580
               Index           =   0
               Left            =   1050
               TabIndex        =   8
               Top             =   885
               Width           =   3105
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
               Index           =   0
               Left            =   105
               TabIndex        =   7
               Top             =   870
               Width           =   810
            End
         End
      End
      Begin MSChartLib.MSChart MSChart1 
         Height          =   3750
         Index           =   0
         Left            =   -73950
         OleObjectBlob   =   "ENSAMOT1.frx":4BF2
         TabIndex        =   2
         Top             =   525
         Width           =   9780
      End
      Begin MSChartLib.MSChart MSChart2 
         Height          =   3750
         Index           =   0
         Left            =   1050
         OleObjectBlob   =   "ENSAMOT1.frx":68CF
         TabIndex        =   3
         Top             =   525
         Width           =   9780
      End
      Begin MSChartLib.MSChart MSChart3 
         Height          =   7785
         Left            =   -74895
         OleObjectBlob   =   "ENSAMOT1.frx":858C
         TabIndex        =   4
         Top             =   420
         Width           =   10830
      End
      Begin MSChartLib.MSChart MSChart1 
         Height          =   3750
         Index           =   1
         Left            =   -73950
         OleObjectBlob   =   "ENSAMOT1.frx":A59B
         TabIndex        =   41
         Top             =   4305
         Width           =   9780
      End
      Begin MSChartLib.MSChart MSChart2 
         Height          =   3750
         Index           =   1
         Left            =   1050
         OleObjectBlob   =   "ENSAMOT1.frx":C278
         TabIndex        =   43
         Top             =   4305
         Width           =   9780
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Alta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   315
         TabIndex        =   45
         Top             =   840
         Width           =   645
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Baja:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   315
         TabIndex        =   44
         Top             =   4620
         Width           =   645
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Baja:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   -74685
         TabIndex        =   42
         Top             =   4620
         Width           =   645
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Alta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   -74685
         TabIndex        =   40
         Top             =   840
         Width           =   645
      End
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
      Height          =   645
      Left            =   11130
      Picture         =   "ENSAMOT1.frx":DF35
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   105
      Width           =   645
   End
End
Attribute VB_Name = "ENSAMOT1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CARE&(1), CAREMA&(1), VEPROMA(1)
Dim VEPUL%(1, 32767)
Dim VETIE(1, 32767)
Dim VERPMI(1, 32767)
Dim VERPMP(1, 32767)
Dim VECUP(1, 32767)
Dim CUPROM(1, 32767)
Dim SIGNO%(1, 32767)
Dim VCORTE

Dim ARCHIM$(1)

Private Sub Chart1_GotFocus()

End Sub

Private Sub Command1_Click(Index As Integer)
   '
   Command1(Index).Enabled = False
   '
   CommonDialog1.CancelError = True
   Label11(Index) = "": Label1(Index) = ""
   '
10 CommonDialog1.DialogTitle = "Lectura Datos de Ensayo - " + TRIM$(Mid$(Frame3(Index).Caption, 11))
   CommonDialog1.FILTER = "CSV - Delimitado por Comas|*.csv|"
   If Option2(Index).Value = True Then
     CommonDialog1.FILTER = "TXT - Texto|*.txt|"
   End If
   CommonDialog1.Flags = &H1000 + &H200000 + &H800& + &H4000
   CommonDialog1.FilterIndex = 1
   On Error Resume Next
   CommonDialog1.ShowOpen
   If Err < 1 Then
      ARXXX$ = CommonDialog1.filename
      If ARXXX$ <> "" Then
        For PPX% = Len(ARXXX$) To 1 Step -1
          If Mid$(ARXXX$, PPX%, 1) = Chr$(92) Then
            Label11(Index) = Left$(ARXXX$, PPX% - 1)
            Label1(Index) = Mid$(ARXXX$, PPX% + 1)
            If Label11(Index) = "" Or Label1(Index) = "" Then
              Call MENSERR(24, "Selección no Válida")
              GoTo 10
            End If
            Exit Sub
          End If
        Next PPX%
        Call MENSERR(24, "Selección no Válida")
        GoTo 10
      End If
   End If
   On Error GoTo 0
   '
   Command1(Index).Enabled = True
   '
End Sub

'
Private Sub Command2_Click()
   '
   Command2.Enabled = False
   '
   HAYDATOS% = 0
   For IENSA% = 0 To 1
     ARCHIM$(IENSA%) = TRIM$(Label11(IENSA%)) + "\" + TRIM$(Label1(IENSA%))
     If EXISTE%(ARCHIM$(IENSA%)) > 0 Then
       HAYDATOS% = 1
     End If
   Next IENSA%
   If HAYDATOS% < 1 Then
     Call MENSERR(24, "Falta Seleccionar Archivo(s)")
     GoTo 99
   End If
   '
   INCREM = Val(Text1)
   If INCREM < 0.001 Then
     Call MENSERR(24, "Datos Encoder no Válidos")
     GoTo 99
   End If
   '
   CONSTA = Val(Text2)
   If CONSTA < 1 Then
     Call MENSERR(24, "Datos Encoder no Válidos")
     GoTo 99
   End If
   
   Screen.MousePointer = 11
   For IENSA% = 0 To 1
     If EXISTE%(ARCHIM$(IENSA%)) > 0 Then
       '
       PULMAX% = 1
       TINI = 0
       VETIE(IENSA%, 0) = TINI
       '
       JJ& = 0
       N1% = FILEOPEN%(ARCHIM$(IENSA%), 1, 128)
       FIN& = LOF(N1%): CALE& = 0
       While Not EOF(N1%)
         Line Input #N1%, AAA$
         CALE& = CALE& + Len(AAA$)
         Call TRACE(20, CALE&, FIN&)
         If JJ& < 32767 Then
           If Val(AAA$) > 0 Or JJ& > 0 Then
             JJ& = JJ& + 1
             VEPUL%(IENSA%, JJ&) = Val(AAA$)
             If VEPUL%(IENSA%, JJ&) > PULMAX% Then PULMAX% = VEPUL%(IENSA%, JJ&)
           End If
         End If
       Wend
20     CARE&(IENSA%) = JJ&
       '
       While (CARE&(IENSA%) > 1 And VEPUL%(IENSA%, CARE&(IENSA%)) < 1)
         CARE&(IENSA%) = CARE&(IENSA%) - 1
       Wend
       '
       For U& = 1 To CARE&(IENSA%)
         VETIE(IENSA%, U&) = VETIE(IENSA%, U& - 1) + INCREM
         VERPMI(IENSA%, U&) = VEPUL%(IENSA%, U&) * CONSTA
       Next U&
       '
       '
       ' determinar sentido de giro
       For U& = 1 To CARE&(IENSA%)
         SUANTE = 0: SUPOST = 0
         For JJ& = 1 To 100
           U1& = U& - JJ&: If U1& < 1 Then U1& = 1
           SUANTE = SUANTE + VERPMI(IENSA%, U1&)
           U1& = U& + JJ&: If U1& > CARE&(IENSA%) Then U1& = CARE&(IENSA%)
           SUPOST = SUPOST + VERPMI(IENSA%, U1&)
         Next JJ&
         If SUANTE < 0.0001 Then
            GoTo 25 'PARA EVITAR DESBORDAMIENTO
         End If
         If VEPUL%(IENSA%, U&) > 0 Then BANDA = (PULMAX% / VEPUL%(IENSA%, U&)) / 25 ' 50
         If Abs(SUPOST / SUANTE) > 1 - BANDA And Abs(SUPOST / SUANTE) < 1 + BANDA And SIGNO%(IENSA%, U& - 1) <> 0 Then
            SIGNO%(IENSA%, U&) = SIGNO%(IENSA%, U& - 1)
           Else
25          SIGNO%(IENSA%, U&) = Sgn(SUPOST - SUANTE)
         End If
       Next U&
       '
       'pone signo a la velocidad instantanea
       For U& = 1 To CARE&(IENSA%)
         VERPMI(IENSA%, U&) = SIGNO%(IENSA%, U&) * Abs(VERPMI(IENSA%, U&))
       Next U&
       '
       'calcula velocidad promedio
       For U& = 1 To CARE&(IENSA%)
         KKI% = 0
         SURPM = VERPMI(IENSA%, U&)
         SUMAPUL% = VEPUL%(IENSA%, U&)
         While SUMAPUL% < 3000
           KKI% = KKI% + 1
           LINF& = U& - KKI%: If LINF < 1 Then LINF = 1
           LSUP& = U& + KKI%: If LSUP& > CARE&(IENSA%) Then LSUP& = CARE&(IENSA%)
           SURPM = SURPM + VERPMI(IENSA%, LINF&) + VERPMI(IENSA%, LSUP&)
           SUMAPUL% = SUMAPUL% + VEPUL%(IENSA%, LINF&) + VEPUL%(IENSA%, LSUP&)
         Wend
         VERPMP(IENSA%, U&) = SURPM / (2 * KKI% + 1)
       Next U&
       '
       CAREMA&(IENSA%) = 1: VEPROMA(IENSA%) = 0
       For U& = 1 To CARE&(IENSA%)
         If VERPMP(IENSA%, U&) > VEPROMA(IENSA%) * 1.0001 Then
           VEPROMA(IENSA%) = VERPMP(IENSA%, U&)
           CAREMA&(IENSA%) = U&
         End If
       Next U&
       '
       'calcula cupla instantanea
       For U& = 1 To CARE&(IENSA%)
         LINF& = U& - 2: If LINF < 1 Then LINF = 1
         LSUP& = U& + 2: If LSUP& > CARE&(IENSA%) Then LSUP& = CARE&(IENSA%)
         VECUP(IENSA%, U&) = (VERPMP(IENSA%, LSUP&) - VERPMP(IENSA%, LINF&)) / (VETIE(IENSA%, LSUP&) - VETIE(IENSA%, LINF))
       Next U&
       '
   AMPLI1% = 50: AMPLI2% = 50
   'AMPLIMA% = 200: AMPLIMI% = 15: VCORTE = VEPROMA / 4
   For U& = 1 To CARE&(IENSA%)
     'AMPLI1% = AMPLIMI%: AMPLI2% = AMPLIMI%
     'If Abs(VERPMP(U&)) < VCORTE Then
     '  AMPLI% = AMPLIMA% - (Abs(VERPMP(U&)) - VCORTE / 5) * (AMPLIMA% - AMPLIMI%) / (VCORTE - VCORTE / 5)
     '  If AMPLI% > AMPLIMA% Then AMPLI% = AMPLIMA%
     '  AMPLI1% = AMPLI%: AMPLI2% = AMPLI%
     'End If
     LINF& = U& - AMPLI1%: If LINF < 1 Then LINF = 1
     LSUP& = U& + AMPLI2%: If LSUP& > CARE&(IENSA%) Then LSUP& = CARE&(IENSA%)
     SUCUP = 0
     ' ESTA ES LA FORMA DE CALCULO SEGUN VENUTOLO
     For JJ& = LINF& To LSUP&
       SUCUP = SUCUP + VECUP(IENSA%, JJ&)
     Next JJ&
     CUPROM(IENSA%, U&) = SUCUP / (LSUP& - LINF& + 1)
     '
     'ESTA ES POR CUADRADOS MINIMOS
     'SUMAVE& = 0: SUMATI = 0
     'For JJ& = LINF& To LSUP&
     '  SUMAVE = SUMAVE + Abs(VERPMP(JJ&) - VERPMP(U&))
     '  SUMATI = SUMATI + Abs(VETIE(JJ&) - VETIE(U&))
     'Next JJ&
     'CUPROM(U&) = SUMAVE / SUMATI
   Next U&
   '
90 List1(IENSA%).Clear
   For U& = 1 To CARE&(IENSA%)
     'TTXX$ = FORMATNUM$(VEPUL%(U&), "I6")
     'TTXX$ = TTXX$ + "  " + FORMATNUM$(VETIE(U&), "F10.3")
     'TTXX$ = TTXX$ + " " + FORMATNUM$(VERPMI(U&), "F12.2")
     'TTXX$ = TTXX$ + " " + FORMATNUM$(VERPMP(U&), "F12.2")
     'TTXX$ = TTXX$ + " " + FORMATNUM$(VECUP(U&), "F12.2")
     TTXX$ = AJUSTD$(Format(U&, "#####0"), 6)
     TTXX$ = TTXX$ + " " + AJUSTD$(Format(VEPUL%(IENSA%, U&), "#####0"), 6)
     TTXX$ = TTXX$ + "  " + AJUSTD$(Format(VETIE(IENSA%, U&), "#####0.000"), 10)
     TTXX$ = TTXX$ + " " + AJUSTD$(Format(VERPMI(IENSA%, U&), "########0.00"), 12)
     TTXX$ = TTXX$ + " " + AJUSTD$(Format(VERPMP(IENSA%, U&), "########0.00"), 12)
     TTXX$ = TTXX$ + " " + AJUSTD$(Format(VECUP(IENSA%, U&), "########0.00"), 12)
     TTXX$ = TTXX$ + " " + AJUSTD$(Format(CUPROM(IENSA%, U&), "########0.00"), 12)
     'TTXX$ = TTXX$ + " " + AJUSTD$(Format(SIGNO%(U&), "#####0"), 6)
     List1(IENSA%).AddItem (TTXX$)
   Next U&
   '
   List1(IENSA%).Refresh
  End If
Next IENSA%
   Call GRAFUNCI
   Screen.MousePointer = 1
   MsgBox "Importación Completa"
   '
99 Command2.Enabled = True
End Sub

Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Form_Load()
   Call AUTOEXEC
   ABC$ = UCase$(LUDAT$)
   If InStr(ABC$, " ") > 0 Then ' \ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB") > 0 Then
     LUDAT$ = "F:\CLIENTES\ADELSUR\"
     MsgBox LUDAT$
   End If
   '
   HOII% = HOY(HOS$)
   '
   'Label11.Caption = Dir1.Path
   'File1.Path = Dir1.Path
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Dir1_Change()
   Label11.Caption = Dir1.Path
   File1.Path = Dir1.Path
   Label1 = ""
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

Private Sub File1_Click()
  Label1 = File1.filename
End Sub

Sub GRAFUNCI()
   '
CAREX& = CARE&(0): If CARE&(1) > CAREX& Then CAREX& = CARE&(1)
CAREMAX& = CAREMA&(0): If CAREMA&(1) > CAREMAX& Then CAREMAX& = CAREMA&(1)
'
For IENSA% = 0 To 1
   MSChart1(IENSA%).Repaint = True
   MSChart1(IENSA%).RowCount = 1 + CAREX& / 10
   MSChart2(IENSA%).Repaint = True
   MSChart2(IENSA%).RowCount = 1 + CAREMAX& / 10
   'For KKU& = 1 To CARE& Step CARE& / 10
   '  MSChart1.Row = KKU&
   '  MSChart1.RowLabel = VETIE(KKU&)
   'Next KKU&
   '
   'GRAFUNCI.MSChart1.Plot.Axis(1).AxisTitle.TEXT = "Pesos"
   'If Option10.Value = True Then
   '  GRAFUNCI.MSChart1.Plot.Axis(1).AxisTitle.TEXT = "U$S"
   'End If
   '
   MSChart1(IENSA%).ColumnCount = 3
   PASOPORCERO% = 0
   MSChart2(IENSA%).Column = 1
   For UJU& = 1 To CARE&(IENSA%) Step 10
     MSChart1(IENSA%).Row = 1 + UJU& / 10
     MSChart1(IENSA%).RowLabel = Format$(VETIE(IENSA%, UJU&), "#0")
     MSChart1(IENSA%).Column = 1
     MSChart1(IENSA%).Data = VERPMP(IENSA%, UJU&)
     MSChart1(IENSA%).Column = 3
     MSChart1(IENSA%).Data = 0
     If PASOPORCERO% = 0 Then
       If VERPMP(IENSA%, UJU&) > 0 Then
         PASOPORCERO% = 1
         MSChart1(IENSA%).Data = 1000
         MSChart1(IENSA%).Row = UJU& / 10
         MSChart1(IENSA%).Data = (-1000)
       End If
     End If
     If Int(1 + UJU& / 10) <= MSChart2(IENSA%).RowCount Then
       MSChart2(IENSA%).Row = Int(1 + UJU& / 10)
       MSChart2(IENSA%).Data = VERPMP(IENSA%, UJU&)
     End If
   Next UJU&
   
   '
   MSChart1(IENSA%).Column = 2
   MSChart2(IENSA%).Column = 2
   For UJU& = 1 To CARE&(IENSA%) Step 10
     MSChart1(IENSA%).Row = 1 + UJU& / 10
'CUPROX = 20 * CUPROM(uju&)
'If CUPROX > 2000 Then CUPROX = 2000
'If CUPROX < (-2000) Then CUPROX = (-2000)
     MSChart1(IENSA%).Data = CUPROM(IENSA%, UJU&)
     If Int(1 + UJU& / 10) <= MSChart2(IENSA%).RowCount Then
       MSChart2(IENSA%).Row = Int(1 + UJU& / 10)
       MSChart2(IENSA%).Data = CUPROM(IENSA%, UJU&)
     End If
   Next UJU&
   '
Next IENSA%
Exit Sub
   PASOPORCERO% = 0
   Dim SUCUP(41), CACAS%(41)
   Dim SUCUP2(41), CACAS2%(41)
   For UJU& = 1 To CAREX&
     If UJU& > 10 Then
       If VEPUL%(IENSA%, UJU&) < 1 Then
         PASOPORCERO% = 1
       End If
     End If
     '
     If PASOPORCERO% > 0 Then ' Or Abs(VERPMP(UJU&)) < VEPROMA / 3 Then
       If CUPROM(IENSA%, UJU&) <= 0 Then Exit For
       CLASI% = 1 + Int(Abs(VERPMP(IENSA%, UJU&)) / 50)
       If CLASI% > 0 Then
         If CLASI% <= 40 Then
           SUCUP(CLASI%) = SUCUP(CLASI%) + Abs(CUPROM(IENSA%, UJU&))
           CACAS%(CLASI%) = CACAS%(CLASI%) + 1
 '          SUCUP(CLASI% + 1) = SUCUP(CLASI% + 1) + Abs(CUPROM(UJU&))
 '          CACAS%(CLASI% + 1) = CACAS%(CLASI% + 1) + 1
 '          SUCUP(CLASI% - 1) = SUCUP(CLASI% - 1) + Abs(CUPROM(UJU&))
 '          CACAS%(CLASI% - 1) = CACAS%(CLASI% - 1) + 1
        End If
       End If
     Else
       CLASI% = 1 + Int(Abs(VERPMP(IENSA%, UJU&)) / 50)
       If CLASI% > 0 Then
         If CLASI% <= 40 Then
           SUCUP2(CLASI%) = SUCUP2(CLASI%) + Abs(CUPROM(IENSA%, UJU&))
           CACAS2%(CLASI%) = CACAS2%(CLASI%) + 1
         End If
       End If
       
     End If
19 Next UJU&
   '
   MSChart3.RowCount = 40
   MSChart3.ColumnCount = 2
   VANT = 0
   For CLASI% = 1 To 40
     MSChart3.Column = 1
     MSChart3.Row = CLASI%
     If CACAS%(CLASI%) < 1 Then
          MSChart3.Data = VANT
       Else
          MSChart3.Data = SUCUP(CLASI%) / CACAS%(CLASI%)
          VANT = MSChart3.Data
     End If
     '
     MSChart3.Column = 2
     MSChart3.Row = CLASI%
     If CACAS2%(CLASI%) < 1 Then
          MSChart3.Data = VANT
       Else
          MSChart3.Data = SUCUP2(CLASI%) / CACAS2%(CLASI%)
          VANT = MSChart3.Data
     End If
   Next CLASI%
   '
   
End Sub


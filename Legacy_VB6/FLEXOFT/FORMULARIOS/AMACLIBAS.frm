VERSION 5.00
Begin VB.Form AMACLIBAS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clientes - A-B-M-L Base de Datos de Clientes"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   Icon            =   "AMACLIBAS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   7050
      Left            =   210
      TabIndex        =   26
      Top             =   1155
      Width           =   10530
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Domicilio"
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
         Left            =   0
         TabIndex        =   59
         Top             =   105
         Width           =   10410
         Begin VB.TextBox Text29 
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
            Left            =   7035
            TabIndex        =   65
            Text            =   " "
            Top             =   1155
            Width           =   2880
         End
         Begin VB.TextBox Text27 
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
            Left            =   3885
            TabIndex        =   64
            Text            =   " "
            Top             =   735
            Width           =   6030
         End
         Begin VB.TextBox Text26 
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
            Left            =   1365
            TabIndex        =   63
            Text            =   " "
            Top             =   735
            Width           =   1305
         End
         Begin VB.TextBox Text25 
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
            Left            =   1365
            TabIndex        =   62
            Text            =   " "
            Top             =   315
            Width           =   8550
         End
         Begin VB.TextBox COTEXT 
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
            Index           =   0
            Left            =   1365
            TabIndex        =   61
            Text            =   "  "
            Top             =   1155
            Width           =   465
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
            Height          =   315
            Index           =   0
            Left            =   1785
            TabIndex        =   60
            Top             =   1155
            Width           =   175
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "País:"
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
            Left            =   5040
            TabIndex        =   71
            Top             =   1260
            Width           =   1905
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cod.Postal:"
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
            Left            =   -630
            TabIndex        =   70
            Top             =   840
            Width           =   1905
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Calle y Nro:"
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
            Left            =   -630
            TabIndex        =   69
            Top             =   420
            Width           =   1905
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Provincia:"
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
            TabIndex        =   68
            Top             =   1260
            Width           =   1185
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
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
            Left            =   1890
            TabIndex        =   67
            Top             =   840
            Width           =   1905
         End
         Begin VB.Label DETEXT 
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
            Index           =   0
            Left            =   2100
            TabIndex        =   66
            Top             =   1155
            Width           =   4005
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Agenda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2180
         Left            =   0
         TabIndex        =   48
         Top             =   1840
         Width           =   5160
         Begin VB.TextBox Text32 
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
            Left            =   1365
            TabIndex        =   53
            Text            =   " "
            Top             =   1035
            Width           =   3600
         End
         Begin VB.TextBox Text31 
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
            Left            =   1365
            TabIndex        =   52
            Text            =   " "
            Top             =   675
            Width           =   3600
         End
         Begin VB.TextBox Text30 
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
            Left            =   1365
            TabIndex        =   51
            Text            =   " "
            Top             =   315
            Width           =   3600
         End
         Begin VB.TextBox Text19 
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
            Left            =   1365
            TabIndex        =   50
            Text            =   " "
            Top             =   1395
            Width           =   3600
         End
         Begin VB.TextBox Text36 
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
            Left            =   1365
            TabIndex        =   49
            Text            =   " "
            Top             =   1755
            Width           =   3600
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "E-Mail:"
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
            Left            =   -630
            TabIndex        =   58
            Top             =   1120
            Width           =   1905
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Fax:"
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
            Left            =   -630
            TabIndex        =   57
            Top             =   760
            Width           =   1905
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono:"
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
            Left            =   -630
            TabIndex        =   56
            Top             =   400
            Width           =   1905
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Contacto:"
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
            Left            =   -630
            TabIndex        =   55
            Top             =   1480
            Width           =   1905
         End
         Begin VB.Label Label71 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Horario:"
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
            Left            =   -630
            TabIndex        =   54
            Top             =   1840
            Width           =   1905
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Impuestos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2180
         Left            =   5355
         TabIndex        =   35
         Top             =   1840
         Width           =   5055
         Begin VB.TextBox Text34 
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
            Left            =   2310
            MaxLength       =   15
            TabIndex        =   41
            Text            =   " "
            Top             =   1260
            Width           =   2565
         End
         Begin VB.TextBox COTEXT 
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
            Index           =   2
            Left            =   945
            MaxLength       =   2
            TabIndex        =   40
            Text            =   " "
            Top             =   1260
            Width           =   360
         End
         Begin VB.TextBox Text33 
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
            Left            =   2310
            MaxLength       =   15
            TabIndex        =   39
            Top             =   315
            Width           =   2565
         End
         Begin VB.TextBox COTEXT 
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
            Index           =   1
            Left            =   945
            MaxLength       =   2
            TabIndex        =   38
            Text            =   " "
            Top             =   315
            Width           =   360
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
            Height          =   315
            Index           =   1
            Left            =   1260
            TabIndex        =   37
            Top             =   315
            Width           =   175
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
            Height          =   315
            Index           =   2
            Left            =   1260
            TabIndex        =   36
            Top             =   1260
            Width           =   175
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nro I-B:"
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
            TabIndex        =   47
            Top             =   1365
            Width           =   1905
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cat I-B:"
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
            Left            =   -525
            TabIndex        =   46
            Top             =   1365
            Width           =   1380
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "CUIT:"
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
            TabIndex        =   45
            Top             =   420
            Width           =   1905
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "P-IVA:"
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
            Left            =   -525
            TabIndex        =   44
            Top             =   420
            Width           =   1380
         End
         Begin VB.Label DETEXT 
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
            Index           =   1
            Left            =   945
            TabIndex        =   43
            Top             =   765
            Width           =   3930
         End
         Begin VB.Label DETEXT 
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
            Index           =   2
            Left            =   945
            TabIndex        =   42
            Top             =   1725
            Width           =   3930
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   10335
         TabIndex        =   33
         Top             =   4200
         Width           =   10395
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Anotaciones - Novedades - Comentarios"
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
            TabIndex        =   34
            Top             =   105
            Width           =   9570
         End
      End
      Begin VB.TextBox Text28 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   4620
         Width           =   10395
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Cortar"
         Height          =   260
         Left            =   5475
         TabIndex        =   31
         Top             =   6740
         Width           =   1170
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Copiar"
         Height          =   260
         Left            =   6660
         TabIndex        =   30
         Top             =   6740
         Width           =   1170
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Pegar"
         Height          =   260
         Left            =   7830
         TabIndex        =   29
         Top             =   6740
         Width           =   1275
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Selecc Todo"
         Height          =   260
         Left            =   9120
         TabIndex        =   28
         Top             =   6740
         Width           =   1290
      End
      Begin VB.CommandButton Command7 
         Height          =   260
         Left            =   0
         TabIndex        =   27
         Top             =   6740
         Width           =   5460
      End
      Begin VB.Line Line3 
         X1              =   10395
         X2              =   10395
         Y1              =   4200
         Y2              =   7000
      End
   End
   Begin VB.TextBox Text23 
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
      Left            =   7035
      TabIndex        =   24
      Top             =   735
      Width           =   1425
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   10815
      ScaleHeight     =   75
      ScaleWidth      =   930
      TabIndex        =   21
      Top             =   4670
      Width           =   990
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "A-B-M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3350
      Left            =   10815
      TabIndex        =   17
      Top             =   1260
      Width           =   990
      Begin VB.CommandButton Command15 
         Caption         =   "Sort"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   105
         Picture         =   "AMACLIBAS.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Ordena Base de Datos de Clientes"
         Top             =   1785
         Width           =   765
      End
      Begin VB.CommandButton Command23 
         Caption         =   "Entreg."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   105
         Picture         =   "AMACLIBAS.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Almacenes - Lugares de Entrega"
         Top             =   2520
         Width           =   765
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
         Height          =   700
         Left            =   105
         Picture         =   "AMACLIBAS.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Alta de Nuevo Cliente"
         Top             =   315
         Width           =   750
      End
      Begin VB.CommandButton Command24 
         Caption         =   "Baja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   105
         Picture         =   "AMACLIBAS.frx":0BA0
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Baja Lógica de la Cuenta"
         Top             =   1050
         Width           =   765
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Varios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3320
      Left            =   10815
      TabIndex        =   19
      Top             =   4800
      Width           =   990
      Begin VB.CommandButton Command99 
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
         Height          =   700
         Left            =   105
         Picture         =   "AMACLIBAS.frx":0EAA
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Control de Integridad de la Base de Datos"
         Top             =   2485
         Width           =   765
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
         Height          =   700
         Left            =   105
         Picture         =   "AMACLIBAS.frx":11B4
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   280
         Width           =   765
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
         Height          =   700
         Left            =   105
         Picture         =   "AMACLIBAS.frx":15F6
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Emisión de Listados"
         Top             =   1015
         Width           =   765
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
         Height          =   700
         Left            =   105
         Picture         =   "AMACLIBAS.frx":1900
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Imprime Ficha del Cliente"
         Top             =   1750
         Width           =   765
      End
   End
   Begin VB.ListBox LINDICE 
      Height          =   1035
      Left            =   10710
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   3780
      Visible         =   0   'False
      Width           =   1170
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
      Height          =   770
      Left            =   8610
      Picture         =   "AMACLIBAS.frx":1F6A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Guarda los Cambios Realizados"
      Top             =   300
      Width           =   750
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Restore"
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
      Left            =   9360
      Picture         =   "AMACLIBAS.frx":2274
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Restablece los Valores Anteriores - Descarta Cambios Efectuados"
      Top             =   300
      Width           =   750
   End
   Begin VB.CommandButton Command4 
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
      Left            =   10290
      Picture         =   "AMACLIBAS.frx":257E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Ayuda en Línea"
      Top             =   300
      Width           =   750
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
      Height          =   770
      Left            =   11040
      Picture         =   "AMACLIBAS.frx":2888
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   300
      Width           =   750
   End
   Begin VB.TextBox Text24 
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
      Left            =   3570
      TabIndex        =   2
      Text            =   " "
      Top             =   315
      Width           =   3300
   End
   Begin VB.TextBox Text3 
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
      Left            =   1155
      TabIndex        =   3
      Top             =   735
      Width           =   5710
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
      Left            =   1155
      MaxLength       =   5
      TabIndex        =   0
      Text            =   " "
      Top             =   315
      Width           =   780
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
      Height          =   330
      Left            =   1935
      Picture         =   "AMACLIBAS.frx":29D2
      TabIndex        =   1
      Top             =   315
      Width           =   175
   End
   Begin VB.Label Label46 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor Nro:"
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
      Left            =   7245
      TabIndex        =   23
      Tag             =   " "
      Top             =   290
      Width           =   1005
   End
   Begin VB.Label Label38 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre de Fantasía:"
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
      Left            =   2310
      TabIndex        =   18
      Top             =   240
      Width           =   1110
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "R.Social:"
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
      Left            =   -315
      TabIndex        =   16
      Top             =   840
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Nro de Cuenta:"
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
      TabIndex        =   15
      Top             =   240
      Width           =   1140
   End
End
Attribute VB_Name = "AMACLIBAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ARCHEX$(16)
Dim YACAR%(8), MODIFIC%(8)
Dim GRABATODO%
Dim AGRECLI%
Dim RECONTANT$

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

'
Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    ARCHEX$(0) = "PROVINC"
    ARCHEX$(1) = "CAPOSIV"
    ARCHEX$(2) = "CARETIB"
    ARCHEX$(3) = "LISTAS"
    ARCHEX$(4) = "CONPAG"
    ARCHEX$(5) = "VENDEDOR"
    ARCHEX$(6) = "ECUECON"
    ARCHEX$(7) = "GRUCOCLI"
    '
    Call CARTAVALI
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
   Call SELECHO(ARCHEX$(Index))
   VDEVU$ = VALACT1$(ARCHEX$(Index))
   If TRIM$(VDEVU$) <> "" Then
      COTEXT(Index).TEXT = VDEVU$
   End If
End Sub

Private Sub Command1_Click()
   GRABATODO% = 1
   Call GRAGENCLI
   Call GRANOTACLI
   GRABATODO% = 0
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   FORLICLI.Show 1
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call COMUNI("Sin Opciones de Configuración\en la Presente Versión.")
   Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
   Command15.Enabled = False
   Call GRAGENCLI
   Call GRANOTACLI
   Call CIERRARCH("DEUDOR1")
   Call CIERRARCH("DEUDOR2")
   Call SORTCLI(1)
   Call CIERRARCH("DEUDOR1")
   Call CIERRARCH("DEUDOR2")
   Command15.Enabled = True
End Sub

Private Sub Command17_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command2_Click()
   Call GRAGENCLI
   Call GRANOTACLI
   For Index = 0 To 8: YACAR%(Index) = 0: Next Index
   Call CARDATCLIE
End Sub

Private Sub Command21_Click()
   '
   If DERACCE%("ALTACLI", "Alta de Nuevo Cliente") < 2 Then
      Exit Sub
   End If
   '
   Command21.Enabled = False
   Call GRAGENCLI
   Call GRANOTACLI
   '
   Screen.MousePointer = 11
   PRONUCLI& = Val(CONTROL$("", "PRONUCLI"))
   If PRONUCLI& > 9999 Then PRONUCLI& = 0
   NROPRO& = 0
   PCLI$ = String$(32767, 0)
   For U& = 1 To ULTREG&("DEUDOR1")
      X$ = REGLEIDO$("DEUDOR1", U&)
      NCXI% = CVI(Left$(X$, 2))
      If NCXI% > 0 Then
        If NCXI% < 9998 Then
          Mid$(PCLI$, NCXI%, 1) = Chr$(1)
          If NCXI% > NROPRO& Then
            NROPRO& = NCXI%
          End If
        End If
      End If
   Next U&
   NCLIE% = NROPRO& + 1
   '
   If PRONUCLI& >= 0 Then
      For KKII% = PRONUCLI& + 1 To 9999
        If Asc(Mid$(PCLI$, KKII%, 1)) < 1 Then
           NCLIE% = KKII%
           NROPRO& = NCLIE% - 1
           GoTo 5
        End If
      Next KKII%
   End If
   '
5  Screen.MousePointer = 1
   '
   VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, MKI$(NCLIE%), 1, 2)
   Call REPLA(VDEF$, Space$(62), 3, 62)
   Call REPLA(VDEF$, Chr$(1), 65, 1)
   Call REPLA(VDEF$, Space$(15), 66, 15)
   '
10 Screen.MousePointer = 1
   Call CIERRARCH("DEUDOR1")
   Call CIERRARCH("DEUDOR2")
   Call CIERRARCH("BADACLIE")
   OBX$ = OBJPLA$("ALTACLIEN", VDEF$)
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   NCLINUE% = CVI(Left$(OBX$, 2))
   '
   If NCLINUE% < 1 Or NCLINUE% > 9999 Then
      Call MENSERR(24, "Número de Cuenta no Válido.")
      GoTo 10
   End If
   '
   If RECFILT$("DEUDOR1", 1, MKI$(NCLINUE%)) <> "" Then
      Call MENSERR(24, "Ingreso no Válido.\Número de Cuenta Pre-existente.")
      GoTo 10
   End If
   '
   DCLINUE$ = TRIM$(Mid$(OBX$, 3, 62))
   If Len(DCLINUE$) < 1 Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Razón Social del Cliente")
      GoTo 10
   End If
   '
   PIVCLI% = Asc(Mid$(OBX$, 65, 1))
   If PIVCLI% < 0 Or PIVCLI% > 6 Then
      Call MENSERR(24, "Ingreso no Válido.\Posición IVA no Figura en Tablas.")
      GoTo 10
   End If
   '
   If PIVCLI% <> 0 Then
     If PIVCLI% <> 5 Then
       CUTT$ = TRIM$(Mid$(OBX$, 66, 15))
       If VALICUIT%(CUTT$) <> 1 Then
         Call MENSERR(24, "Ingreso no Válido.\Número de C.U.I.T. Incorrecto.")
         GoTo 10
       End If
     End If
   End If
   '
   Screen.MousePointer = 11
   ARX1$ = "DEUDOR1": ARX2$ = "DEUDOR2"
   Call BLOQARCH(ARX1$)
   Call BLOQARCH(ARX2$)
   Call BLOQARCH("BADACLIE")
   '
   CUTT$ = TRIM$(CUTT$)
   For U& = 1 To ULTREG&(ARX1$)
     XX$ = REGLEIDO$(ARX1$, U&) + REGLEIDO$(ARX2$, U&)
     If CVI(Left$(XX$, 2)) = NCLINUE% Then
       Call MENSERR(24, "Ingreso no Válido.\Número de Cuenta Pre-existente.")
       GoTo 10
     End If
     '
     If CUTT$ <> "" Then
       If CVI(Left$(XX$, 2)) > 0 Then
         If TRIM$(Mid$(XX$, 107, 15)) = CUTT$ Then
           NCLIX% = CVI(Left$(XX$, 2))
           If COMALTER%("Existe Registrada una Cuenta con el Mismo C.U.I.T.\  \'" + TRIM$(RASOCLI$(NCLIX%)) + "' (" + TRIM$(Str$(NCLIX%)) + ")\\Cancelar\Registrar Igual") <> 2 Then
             GoTo 10
           End If
         End If
       End If
     End If
   Next U&
   '
   REGX& = 1 + ULTREG&(ARX1$)
   XX$ = REGBLAN$(ARX1$) + String$(128, 0)
   YY$ = REGBLAN$("BADACLIE")
   Call REPLA(XX$, MKI$(NCLINUE%), 1, 2)
   Call REPLA(XX$, DCLINUE$, 3, 30)
   Call REPLA(XX$, TRIM$(CUTT$), 107, 15)
   Call REPLA(XX$, Chr$(PIVCLI%), 136, 1)
   '
   Call REPLA(YY$, MKI$(NCLINUE%), 1, 2)
   Call REPLA(YY$, DCLINUE$, 3, 62)
   '
   Call REGAPP(ARX1$, Left$(XX$, 32))
   REGX& = ULTREG&(ARX1$)
   RGBLX$ = REGBLAN$(ARX2$)
   While ULTREG&(ARX2$) < REGX&
     Call REGAPP(ARX2$, RGBLX$)
   Wend
   '
   Call GRAREG(ARX1$, Left$(XX$, 32), REGX&)
   Call GRAREG(ARX2$, Mid$(XX$, 33), REGX&)
   Call CIERRARCH(ARX1$)
   Call CIERRARCH(ARX2$)
   '
   RFF$ = RECFILT$("BADACLIE", 1, MKI$(NCLINUE%))
   If Len(RFF$) >= 4 Then
     REBADA& = CVS(Left$(RFF$, 4))
     If REBADA& > 0 Then
       If REBADA& <= ULTREG&("BADACLIE") Then
         Call GRAREG("BADACLIE", YY$, REBADA&)
         GoTo 29
       End If
     End If
   End If
   Call REGAPP("BADACLIE", YY$)
   '
29 Call CIERRARCH("BADACLIE")
   Call BAJALDISCO
   '
   AGRECLI% = 1
   Text1.TEXT = TRIM$(Str$(NCLINUE%))
   Call CARDATCLIE
   Call CARTAVALI
   '
             For Index = 0 To 2
             If Index <> 1 Then
             VAPRE$ = TRIM$(CONTROL$("AMACLI", ARCHE$(Index)))
             COTEXT(Index) = "": DETEXT(Index) = ""
             If VAPRE$ <> "" Then
                TV11% = TIPFIELD(ARCHE$(Index), 1)
                LF11% = LENFIELD(ARCHE$(Index), 1)
                ARGU$ = CBIN$(VAPRE$, TV11%, LF11%)
                ECO$ = TRIM$(ECOARCH$(ARCHE$(Index), ARGU$))
                If ECO$ <> "" Then
                    COTEXT(Index).TEXT = VAPRE$
                    DETEXT(Index).Caption = ECO$
                  Else
                    COTEXT(Index).TEXT = ""
                    DETEXT(Index).Caption = ""
                End If
             End If
             End If
             'DETEXT(Index).TabStop = False
          Next Index
          '

   Screen.MousePointer = 1
   '
99 Command21.Enabled = True
   '
End Sub

Private Sub Command23_Click()
   Command23.Enabled = False
   NCXI% = Val(Text1)
   If ECOARCH("DEUDOR1", MKI$(NCXI%)) = "" Then
      Call MENSERR(24, "Ingrese Número de Cliente")
      GoTo 99
   End If
   '
   RFF$ = RECFILT$("ALMAENT", 3, MKI$(NCXI%))
   JJ& = 0
   For U& = 1 To Len(RFF$) Step 4
     REGE& = CVS(Mid$(RFF$, U&, 4))
     X$ = REGLEIDO$("ALMAENT", REGE&)
     JJ& = JJ& + 1
     Call GRAREG("!ALMAENT", X$, JJ&)
   Next U&
   Call SETULTREG("!ALMAENT", JJ&)
   '
   VTB% = VENTABU%("ALMAENT")
   If VTB% >= 0 Then
     URN& = ULTREG&("ALMAENT")
     RNC$ = RECFILT$("ALMAENT", 3, MKI$(NCXI%))
     KK& = (-3)
     '
     For U& = 1 To ULTREG&("!ALMAENT")
      TTYY$ = AJUSTI$(REGLEIDO$("!ALMAENT", U&), 30) + MKI$(NCXI%)
      KK& = KK& + 4
      If KK& <= Len(RNC$) - 3 Then
           RENOTA& = CVS(Mid$(RNC$, KK&, 4))
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG("ALMAENT", TTYY$, RENOTA&)
     Next U&
   End If
   '
20 While KK& < Len(RNC$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RNC$, KK&, 4))
      Call GRAREG("ALMAENT", Space$(30) + MKI$(0), RENOTA&)
   Wend
   '
   Call CIERRARCH("ALMAENT")
99 Command23.Enabled = True
   '
End Sub

Private Sub Command24_Click()
   '
   Command24.Enabled = False
   Screen.MousePointer = 11
   '
   Call GRAGENCLI
   Call GRANOTACLI
   '
   If DERACCE%("BAJACLI", "Baja de Cuenta de Cliente") < 2 Then
     GoTo 99
   End If
   '
   If Val(Text1.TEXT) > 0 Then
     If Val(Text1.TEXT) <= 32767 Then
       NCLIE% = Val(Text1.TEXT)
       If REGICLI&(NCLIE%) > 0 Then
         NC% = NCLIE%
         '
         If Abs(SALDCLI#(NC%)) >= 0.005 Then
           Call MENSERR(24, "Baja Imposible - Cuenta con Saldo.")
           GoTo 99
         End If
         '
         RFF$ = RECFILT$("FACTUR", 1, MKI$(NC%))
         If RFF$ <> "" Then
           Call MENSERR(24, "Baja Imposible - Cuenta con Movimiento.")
           GoTo 99
         End If
         '
         RFF$ = RECFILT$("COBRAN", 1, MKI$(NC%))
         If RFF$ <> "" Then
           Call MENSERR(24, "Baja Imposible - Cuenta con Movimiento.")
           GoTo 99
         End If
         
         TTXX$ = "La Baja de un Cliente\Puede Producir Inconsistencias en la Base de Datos\en Caso de que la Cuenta Registre\Facturación, Cobranza o Pedidos."
         TTXX$ = TTXX$ + "\  \Puede Ahora Cancelar la Operación\o Confirmar la Baja de la Cuenta:\'" + TRIM$(Text3.TEXT) + "'"
         If COMALTER%(TTXX$ + "\\Cancelar Operación\Confirmar Baja") = 2 Then
           '
           YACAR%(1) = 1
           Text28.TEXT = ""
           MODIFIC%(1) = 1
           GRABATODO% = 1
           Call GRANOTACLI
           '
           REG& = REGICLI&(NCLIE%)
           XX$ = REGLEIDO$("DEUDOR1", REG&)
           Call COPYSTRU("DEUDOR1", "DEUDOBAJ")
           Call REGAPP("DEUDOBAJ", XX$)
           Call CIERRARCH("DEUDOBAJ")
           '
           Call REPLA(XX$, MKI$(0), 1, 2)
           Call REPLA(XX$, Space$(30), 3, 30)
           Call GRAREG("DEUDOR1", XX$, REG&)
           Call CIERRARCH("DEUDOR1")
           '
           AGRECLI% = 1
           Text1.TEXT = ""
           Call BLANFORMU
           For KKI% = 0 To 4: MODIFIC%(KKI%) = 0: YACAR%(KKI%) = 0: Next KKI%
           GRABATODO% = 0
           Call BAJALDISCO
           '
         End If
         '
       End If
     End If
   End If
   '
99 Screen.MousePointer = 1
   Command24.Enabled = True
   '
End Sub

Private Sub Command25_Click()
   PRINTFORM
End Sub

Private Sub Command3_Click()
   '
   Call GRAGENCLI
   Call GRANOTACLI
   Call CARDATCLIE
   '
   Call CIERRARCH("*.*")
   If AGRECLI% > 0 Then
      Call SORTCLI(1)
      AGRECLI% = 0
   End If
   Call BAJALDISCO
   Hide
   '
End Sub

Private Sub Command5_Click()
   Text28.SelStart = 0
   Text28.SelLength = Len(Text28.TEXT)
   Text28.SetFocus
End Sub

Private Sub Command6_Click()
   COMIE = Text28.SelStart
   LOSEL = Text28.SelLength
   Text28.TEXT = Left$(Text28.TEXT, COMIE) + Clipboard.GetText + Mid$(Text28.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text28.TEXT, Text28.SelStart + 1, Text28.SelLength)
End Sub

Private Sub Command9_Click()
   '
   Call GRAGENCLI
   Call GRANOTACLI
   '
   Call SELECHO("DEUDOR1")
   NCLIE% = Val(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE% < 1 Then Exit Sub
   '
   NC% = NCLIE%
   Text1.TEXT = TRIM$(Str$(NC%))
   'Call CARDATCLIE
   '
End Sub

Private Sub COTEXT_Change(Index As Integer)
   '
   ECO$ = ""
   If TRIM$(COTEXT(Index).TEXT) <> "" Then
      TV11% = TIPFIELD(ARCHEX$(Index), 1)
      LF11% = LENFIELD(ARCHEX$(Index), 1)
      ARGU$ = CBIN$(TRIM$(COTEXT(Index).TEXT), TV11%, LF11%)
      ECO$ = TRIM$(ECOARCH$(ARCHEX$(Index), ARGU$))
   End If
   '
   If ECO$ <> "" Then
        DETEXT(Index).Caption = ECO$
      Else
        DETEXT(Index).Caption = "No Definido"
   End If
   '
   MODIFIC%(0) = 1
End Sub

Private Sub COTEXT_DblClick(Index As Integer)
   Call SELECHO(ARCHEX$(Index))
   VDEVU$ = VALACT1$(ARCHEX$(Index))
   If TRIM$(VDEVU$) <> "" Then
      COTEXT(Index).TEXT = VDEVU$
   End If
End Sub

Sub CARDATCLIE()
    '
    NCLIE% = Val(TRIM$(Text1.TEXT))
    If NCLIE% < 1 Or ECOARCH$("DEUDOR1", MKI$(NCLIE%)) = "" Then
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Text3.TEXT = RASOCLI$(NCLIE%)
    Text24.TEXT = TEL2CLI$(NCLIE%)
    Text23.TEXT = NPROCLI$(NCLIE%)
    MODIFIC(0) = 0
    '
    Call CARGENCLI
    Call CARNOTACLI
    Screen.MousePointer = 1
    '
End Sub
'
Sub CARGENCLI()
    '
    If YACAR%(0) > 0 Then Exit Sub
    NC% = Val(TRIM$(Text1.TEXT))
    '
    Text25.TEXT = DOMICLI$(NC%)
    Text25.Refresh
    Text26.TEXT = CPOSCLI$(NC%)
    Text26.Refresh
    Text27.TEXT = LOCACLI$(NC%)
    Text27.Refresh
    Text23.TEXT = NPROCLI$(NC%)
    Text23.Refresh
    Text29.TEXT = PAISCLI$(NC%)
    Text29.Refresh
    Text30.TEXT = TEL1CLI$(NC%)
    Text30.Refresh
    Text31.TEXT = FAXCLI$(NC%)
    Text31.Refresh
    Text32.TEXT = EMAILCLI$(NC%)
    Text32.Refresh
    Text19.TEXT = CONTACTCLI$(NC%)
    Text19.Refresh
    Text36.TEXT = HORARCLI$(NC%)
    Text36.Refresh
    Text33.TEXT = CUITCLI$(NC%)
    Text33.Refresh
    Text34.TEXT = NIBRCLI$(NC%)
    Text34.Refresh
    '
    If ULTREG&("CAPOSIV") < 8 Then
      IQ& = 0
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(0) + AJUSTI$("IVA Consumidor Final            ", 31) + AJUSTI$("Cons.Fin", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(1) + AJUSTI$("IVA Responsable Inscripto   ", 31) + AJUSTI$("Resp.Ins", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(2) + AJUSTI$("IVA Responsable no Inscripto", 31) + AJUSTI$("R.No Ins", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(3) + AJUSTI$("IVA No Responsable          ", 31) + AJUSTI$("No Resp.", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(4) + AJUSTI$("IVA Responsable Exento      ", 31) + AJUSTI$("Exento  ", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(5) + AJUSTI$("IVA Operac.Comercio Exterior", 31) + AJUSTI$("Com.Ext. ", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(6) + AJUSTI$("IVA Contrib.Monotributo     ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(7) + AJUSTI$("IVA Exento Zona Promovida   ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
      Call SETULTREG("CAPOSIV", IQ&)
      Call CIERRARCH("CAPOSIV")
    End If
    '
    COTEXT(0).TEXT = TRIM$(PROVCLI$(NC%))
    COTEXT(0).Refresh
    COTEXT(1).TEXT = TRIM$(Str$(PIVACLI%(NC%)))
    COTEXT(1).Refresh
    COTEXT(2).TEXT = TRIM$(CATIBRU$(NC%))
    COTEXT(2).Refresh
    '
    YACAR%(0) = 1: MODIFIC%(0) = 0
    '
End Sub
'
Sub GRAGENCLI()
    '
    NCLIE% = Val(TRIM$(Text1.TEXT))
    If NCLIE% < 1 Then
      If GRABATODO% = 1 Then
        Call MENSERR(24, "Imposible Registrar !!!\Número de Cliente no Válido.")
        Exit Sub
      End If
    End If
    '
    If MODIFIC%(0) < 1 Then
      If GRABATODO% <> 1 Then
        Exit Sub
      End If
    End If
    '
    If DERACCE%("MODATCLI", "Modificar Datos de Clientes") < 2 Then
       Exit Sub
    End If
    '
    If GRABATODO% <> 1 Then
      Command9.SetFocus
      If COMALTER%("Se han Modificado Datos del Cliente\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
        YACAR%(0) = 0
        Screen.MousePointer = 11
        Text3.TEXT = RASOCLI$(NCLIE%)
        Text24.TEXT = TEL2CLI$(NCLIE%)
        MODIFIC(0) = 0
        Call CARGENCLI
        Screen.MousePointer = 1
        Exit Sub
      End If
    End If
    '
    Screen.MousePointer = 11
    '
    CUIT$ = TRIM$(Text33.TEXT)
    If CUIT$ <> "" Then
      If VALICUIT%(CUIT$) <> 1 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar !!!\Número de C.U.I.T. no Válido")
        SSTab1.Tab = 0
        Text33.SetFocus
        Exit Sub
      End If
    End If
    '
    If NCLIE% = 0 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar Cuenta Número '0'")
        Exit Sub
    End If
    '
    If TRIM$(Text3.TEXT) = "" Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Imposible Registrar !!!\Falta Razón Social del Cliente.")
        Exit Sub
    End If
    '
    ARX1$ = "DEUDOR1": ARX2$ = "DEUDOR2"
    REGX& = REGICLI&(NCLIE%)
    XX11$ = REGLEIDO$("DEUDOR1", REGX&)
    If Abs(CVI(Left$(XX11$, 2))) <> Abs(NCLIE%) Then
       REGX& = 0
    End If
    '
    If REGX& < 1 Or REGX& > ULTREG&(ARX1$) Then
       REGX& = 0
       For UUL& = 1 To ULTREG&("DEUDOR1")
         XX11$ = REGLEIDO$("DEUDOR1", UUL&)
         If Abs(CVI(Left$(XX11$, 2))) = NCLIE% Then
           REGX& = UUL&
           GoTo 20
         End If
       Next UUL&
       REGX& = 1 + ULTREG&("DEUDOR1")
       XX11$ = REGBLAN$(ARX1$)
       Call REPLA(XX11$, MKI$(NCLIE%), 1, 2)
       Call REPLA(XX11$, TRIM$(Text3.TEXT), 3, 30)
       Call GRAREG(ARX1$, Left$(XX11$, 32), REGX&)
       RGBLX$ = REGBLAN$(ARX2$)
       While ULTREG&(ARX2$) < REGX&
         Call REGAPP(ARX2$, RGBLX$)
       Wend
       Call CIERRARCH(ARX1$): Call CIERRARCH(ARX2$)
       Call BAJALDISCO
    End If
    '
20  XX$ = REGLEIDO$(ARX1$, REGX&) + REGLEIDO$(ARX2$, REGX&)
    If Abs(CVI(Left$(XX$, 2))) <> Abs(NCLIE%) Then
       Screen.MousePointer = 1
       Call MENSERR(24, "Imposible Registrar.\Consulte Soporte de Sistemas.")
       Exit Sub
    End If
    '
    CARASO% = 0
    If TRIM$(Text3.TEXT) <> TRIM$(RASOCLI$(NCLIE%)) Then
       CARASO% = 1
    End If
    '
    If TRIM$(Left$(Text3.TEXT, 30)) <> TRIM$(Mid$(XX$, 3, 30)) Then
       AGRECLI% = 1
    End If
    '
    Call REPLA(XX$, TRIM$(Text3.TEXT), 3, 30)
    Call REPLA(XX$, TRIM$(Text25.TEXT), 33, 25)
    CPOST% = 0
    If Val(Text26.TEXT) > 0 Then
      If Val(Text26.TEXT) < 32768 Then
        CPOST% = Val(Text26.TEXT)
      End If
    End If
    Call REPLA(XX$, MKI$(CPOST%), 60, 2)
    Call REPLA(XX$, TRIM$(Text27.TEXT), 62, 25)
    Call REPLA(XX$, TRIM$(Text30.TEXT), 87, 10)
    Call REPLA(XX$, TRIM$(Text24.TEXT), 97, 10)
    Call REPLA(XX$, TRIM$(CUIT$), 107, 15)
    Call REPLA(XX$, TRIM$(Text34.TEXT), 122, 13)
    Call REPLA(XX$, TRIM$(COTEXT(2).TEXT), 135, 1)
    Call REPLA(XX$, TRIM$(COTEXT(0).TEXT), 149, 1)
    '
    PIVA% = VALIMI(COTEXT(1).TEXT, 0, 7)
    Call REPLA(XX$, Chr$(PIVA%), 136, 1)
    '
    CATIB% = VALIMI(COTEXT(2).TEXT, 0, 2)
    '
    Call GRAREG(ARX1$, Left$(XX$, 32), REGX&)
    Call GRAREG(ARX2$, Mid$(XX$, 33), REGX&)
    Call CIERRARCH(ARX1$)
    Call CIERRARCH(ARX2$)
    '
    VDL& = CVI(Mid$(XX$, 159, 2))
    If TRIM$(Mid$(XX$, 33, 25)) <> "" Then
      If VDL& > 0 Then
        For KK& = 1 To ULTREG&("BKDATCLI")
          XXYY$ = REGLEIDO$("BKDATCLI", KK&)
          If CVI(Left$(XXYY$, 2)) = NCLIE% Then
            Call GRAREG("BKDATCLI", XX$, KK&)
            GoTo 23
          End If
        Next KK&
        Call REGAPP("BKDATCLI", XX$)
23      Call CIERRARCH("BKDATCLI")
      End If
    End If
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    If NCLIE% > 0 Then
       RFF$ = RECFILT$("BADACLIE", 1, MKI$(NCLIE%))
       REBACLI& = 0
       If Len(RFF$) >= 4 Then
          REBACLI& = CVS(Left$(RFF$, 4))
       End If
       If REBACLI& < 1 Or REBACLI& > ULTREG&("BADACLIE") Then
         REBACLI& = 0
         For UUL& = 1 To ULTREG&("BADACLIE")
           XX11$ = REGLEIDO$("BADACLIE", UUL&)
           If Abs(CVI(Left$(XX11$, 2))) = NCLIE% Then
             REBACLI& = UUL&
             GoTo 120
           End If
         Next UUL&
         REBACLI& = 1 + ULTREG&("BADACLIE")
       End If
       '
120    XX$ = REGBLAN$("BADACLIE")
       Call REPLA(XX$, MKI$(NCLIE%), 1, 2)
       Call REPLA(XX$, Text3.TEXT, 3, 62)
       Call REPLA(XX$, Text24.TEXT, 65, 32)
       Call REPLA(XX$, Text25.TEXT, 97, 64)
       Call REPLA(XX$, Text26.TEXT, 161, 16)
       Call REPLA(XX$, Text27.TEXT, 177, 48)
       Call REPLA(XX$, Text29.TEXT, 225, 32)
       Call REPLA(XX$, Text30.TEXT, 257, 24)
       Call REPLA(XX$, Text31.TEXT, 281, 24)
       Call REPLA(XX$, Text32.TEXT, 305, 48)
       Call REPLA(XX$, Text19.TEXT, 385, 48)
       Call REPLA(XX$, Text23.TEXT, 433, 16)
       Call REPLA(XX$, Text36.TEXT, 449, 32)
       Call GRAREG("BADACLIE", XX$, REBACLI&)
       Call CIERRARCH("BADACLIE")
    End If
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    MODIFIC%(0) = 0
    Screen.MousePointer = 1
    '
End Sub
'
Sub CARNOTACLI()
   '
   If YACAR%(1) > 0 Then Exit Sub
   NCLIE% = Val(Text1.TEXT)
   Text28.TEXT = ""
   '
   If NCLIE% > 0 Then
      If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
         Screen.MousePointer = 11
         RNC$ = RECFILT$("NOTACLI", 1, MKI$(NCLIE%))
         URN& = ULTREG&("NOTACLI")
         '
         For U& = 1 To Len(RNC$) Step 4
            RENOTA& = CVS(Mid$(RNC$, U&, 4))
            If RENOTA& > 0 Then
               If RENOTA& <= URN& Then
                  TTYY$ = RTrim$(Mid$(REGLEIDO$("NOTACLI", RENOTA&), 3))
                  If Text28.TEXT <> "" Then
                     Text28.TEXT = Text28.TEXT + Chr$(13) + Chr$(10)
                  End If
                  Text28.TEXT = Text28.TEXT + RTrim$(TTYY$)
               End If
            End If
         Next U&
         Screen.MousePointer = 1
      End If
   End If
   Text28.Refresh
   YACAR%(1) = 1: MODIFIC%(1) = 0
   '
End Sub
'
Sub GRANOTACLI()
   '
   NCLIE% = Val(TRIM$(Text1.TEXT))
   If NCLIE% < 1 Then Exit Sub
   If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) = "" Then Exit Sub
   If YACAR%(1) < 1 Then Exit Sub
   '
   If NCLIE% > 0 Then
      If ECOARCH$("DEUDOR1", MKI$(NCLIE%)) <> "" Then
         '
         If MODIFIC%(1) < 1 Then
           Exit Sub
         End If
         '
         If DERACCE%("ANOTACLI", "Anotaciones de de Clientes") < 2 Then
            Exit Sub
         End If
         '
         If GRABATODO% <> 1 Then
           Command9.SetFocus
           If COMALTER%("Se han Modificado las Notas del Cliente\'" + TRIM$(Text3.TEXT) + "'\\Registrar Cambios\Cancelar") <> 1 Then
             YACAR%(1) = 0
             Call CARNOTACLI
             Exit Sub
           End If
           Command9.SetFocus
         End If
         '
         Screen.MousePointer = 11
         URN& = ULTREG&("NOTACLI")
         RNC$ = RECFILT$("NOTACLI", 1, MKI$(NCLIE%))
         KK& = (-3)
         '
         TEXTO$ = Text28.TEXT + Chr$(13) + Chr$(10)
         Call FRATEXTO(TEXTO$, 78)
         For LL% = 1 To CARETEX%
            TTYY$ = MKI$(NCLIE%) + RETEX$(LL%)
            KK& = KK& + 4
            If KK& <= Len(RNC$) - 3 Then
                RENOTA& = CVS(Mid$(RNC$, KK&, 4))
              Else
                URN& = URN& + 1
                RENOTA& = URN&
            End If
            Call GRAREG("NOTACLI", TTYY$, RENOTA&)
         Next LL%
         '
20       While KK& < Len(RNC$) - 3
            KK& = KK& + 4
            RENOTA& = CVS(Mid$(RNC$, KK&, 4))
            Call GRAREG("NOTACLI", MKI$(0) + Space$(78), RENOTA&)
         Wend
         '
         Call CIERRARCH("NOTACLI")
         MODIFIC%(1) = 0
         Screen.MousePointer = 1
      End If
   End If

End Sub
'
Sub BLANFORMU()
   '
   Text3 = ""
   Text24 = ""
   Text25.TEXT = ""
   Text26.TEXT = ""
   Text27.TEXT = ""
   Text29.TEXT = ""
   Text30.TEXT = ""
   Text31.TEXT = ""
   Text32.TEXT = ""
   Text19.TEXT = ""
   Text33.TEXT = ""
   Text34.TEXT = ""
   Text36.TEXT = ""
   '
   For KK% = 0 To 2
     COTEXT(KK%).TEXT = ""
   Next KK%
   '
   Text28.TEXT = ""
   '
   Text23 = ""
   Text35 = ""
   Text36 = ""
   Text37 = ""
   '
End Sub

'
Private Sub Form_Unload(Cancel As Integer)
   '
   Call GRAGENCLI
   Call GRANOTACLI
   Call CARDATCLIE
   '
   Call CIERRARCH("*.*")
   If AGRECLI% > 0 Then
      Call SORTCLI(1)
      AGRECLI% = 0
   End If
   Call BAJALDISCO
   On Error Resume Next
   Hide
   On Error GoTo 0
   '
End Sub

'
Private Sub Text1_Change()
   For Index = 0 To 8: YACAR%(Index) = 0: Next Index
   Call BLANFORMU
   If Val(Text1.TEXT) > 0 Then
     If Val(Text1.TEXT) <= 32767 Then
       NCLIE% = Val(Text1.TEXT)
       If REGICLI&(NCLIE%) > 0 Then
         NC% = NCLIE%
         Call CARDATCLIE
       End If
     End If
   End If
End Sub
'
Private Sub Text1_DblClick()
   '
   Call GRAGENCLI
   Call GRANOTACLI
   '
   Call SELECHO("DEUDOR1")
   NCLIE% = Val(TRIM$(VALACT1$("DEUDOR1")))
   If NCLIE% < 1 Then Exit Sub
   '
   NC% = NCLIE%
   Text1.TEXT = TRIM$(Str$(NC%))
   '
End Sub

Private Sub Text1_GotFocus()
   '
   Call GRAGENCLI
   Call GRANOTACLI
   '
End Sub

Private Sub Text19_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text23_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text23_GotFocus()
   'If SSTab1.Tab <> 0 Then
   '   SSTab1.Tab = 0
   'End If
End Sub

Private Sub Text24_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text25_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text26_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text27_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text28_Change()
   MODIFIC(1) = 1
End Sub

Private Sub Text29_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text3_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text3_GotFocus()
   'If SSTab1.Tab <> 0 Then
   '   SSTab1.Tab = 0
   'End If
End Sub

Private Sub Text30_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text31_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text32_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text33_Change()
   MODIFIC%(0) = 1
End Sub

Private Sub Text34_Change()
   MODIFIC%(0) = 1
End Sub


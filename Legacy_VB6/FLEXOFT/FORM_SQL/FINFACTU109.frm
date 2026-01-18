VERSION 5.00
Begin VB.Form FINFACTU109 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aprobación de Comprobante"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8130
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Aprobar"
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
      Left            =   5895
      Picture         =   "FINFACTU109.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5355
      Width           =   1000
   End
   Begin VB.CommandButton Command4 
      Caption         =   "C.A.E."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   5880
      Picture         =   "FINFACTU109.frx":030A
      TabIndex        =   63
      Top             =   5355
      Width           =   1000
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Consumo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   5880
      TabIndex        =   54
      Top             =   3885
      Width           =   2115
      Begin VB.CommandButton Command3 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1785
         TabIndex        =   57
         Top             =   420
         Width           =   225
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1260
         TabIndex        =   55
         Top             =   420
         Width           =   540
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Depósito:"
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
         Left            =   120
         TabIndex        =   58
         Top             =   525
         Width           =   1005
      End
      Begin VB.Label Label8 
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
         Height          =   315
         Left            =   105
         TabIndex        =   56
         Top             =   840
         Width           =   1905
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Detalle de Exportación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   160
      TabIndex        =   40
      Top             =   6240
      Width           =   7820
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   6165
         MaxLength       =   25
         TabIndex        =   66
         Top             =   1680
         Width           =   1395
      End
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1995
         TabIndex        =   65
         Top             =   1680
         Width           =   3300
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   1995
         MaxLength       =   25
         TabIndex        =   60
         Top             =   1350
         Width           =   1995
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   5565
         MaxLength       =   25
         TabIndex        =   59
         Top             =   1350
         Width           =   1995
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   5565
         MaxLength       =   25
         TabIndex        =   51
         Top             =   1050
         Width           =   1995
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   1995
         MaxLength       =   25
         TabIndex        =   49
         Top             =   1050
         Width           =   1995
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   5565
         MaxLength       =   20
         TabIndex        =   47
         Top             =   735
         Width           =   1995
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   1995
         MaxLength       =   20
         TabIndex        =   45
         Top             =   735
         Width           =   1995
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   5565
         MaxLength       =   20
         TabIndex        =   42
         Top             =   420
         Width           =   1995
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   1995
         MaxLength       =   20
         TabIndex        =   41
         Top             =   420
         Width           =   1995
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Importe:"
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
         Index           =   12
         Left            =   4440
         TabIndex        =   68
         Top             =   1740
         Width           =   1680
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Gastos Exportacion:"
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
         Index           =   13
         Left            =   120
         TabIndex        =   67
         Top             =   1740
         Width           =   1860
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Seguro:"
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
         Index           =   11
         Left            =   3885
         TabIndex        =   62
         Top             =   1410
         Width           =   1680
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Costo CIF - Flete:"
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
         Index           =   10
         Left            =   105
         TabIndex        =   61
         Top             =   1410
         Width           =   1860
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Guia de Importación:"
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
         Index           =   9
         Left            =   105
         TabIndex        =   52
         Top             =   1110
         Width           =   1860
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Credito Nro.:"
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
         Index           =   8
         Left            =   3885
         TabIndex        =   50
         Top             =   1110
         Width           =   1680
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Banco Local:"
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
         Index           =   7
         Left            =   3885
         TabIndex        =   48
         Top             =   795
         Width           =   1680
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Banco Exterior:"
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
         Index           =   6
         Left            =   105
         TabIndex        =   46
         Top             =   795
         Width           =   1860
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Puerto de Embarque:"
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
         Index           =   5
         Left            =   105
         TabIndex        =   44
         Top             =   475
         Width           =   1860
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
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
         Height          =   225
         Index           =   4
         Left            =   3885
         TabIndex        =   43
         Top             =   475
         Width           =   1680
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Detalle de Remito"
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
      Left            =   945
      TabIndex        =   31
      Top             =   1680
      Visible         =   0   'False
      Width           =   5580
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   420
         Width           =   5160
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Observaciones en Remito"
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
      Left            =   160
      TabIndex        =   36
      ToolTipText     =   "Doble Click Para Quitar Observaciones"
      Top             =   2310
      Width           =   7820
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Top             =   315
         Width           =   7365
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Bultos - Valor Declarado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   160
      TabIndex        =   26
      Top             =   3885
      Width           =   5580
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   1575
         MaxLength       =   10
         TabIndex        =   53
         Top             =   840
         Width           =   1365
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   4130
         MaxLength       =   10
         TabIndex        =   39
         Top             =   840
         Width           =   1365
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   1575
         MaxLength       =   10
         TabIndex        =   9
         Top             =   420
         Width           =   1365
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4725
         MaxLength       =   3
         TabIndex        =   10
         Top             =   420
         Width           =   750
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Bultos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   4050
         TabIndex        =   27
         Top             =   525
         Width           =   600
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Kilos Neto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   585
         TabIndex        =   33
         Top             =   945
         Width           =   945
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Valor Declarado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   28
         Top             =   525
         Width           =   1440
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Kilos Bruto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3030
         TabIndex        =   38
         Top             =   945
         Width           =   990
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Observaciones en Remito"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   7245
      TabIndex        =   24
      ToolTipText     =   "Doble Click Para Quitar Observaciones"
      Top             =   3045
      Visible         =   0   'False
      Width           =   7680
      Begin VB.PictureBox Picture2 
         Height          =   795
         Left            =   105
         ScaleHeight     =   735
         ScaleWidth      =   7365
         TabIndex        =   25
         Top             =   315
         Width           =   7420
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
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
            Index           =   8
            Left            =   0
            MaxLength       =   48
            TabIndex        =   8
            Text            =   " "
            Top             =   460
            Width           =   7420
         End
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
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
            Index           =   7
            Left            =   0
            MaxLength       =   48
            TabIndex        =   7
            Text            =   " "
            Top             =   230
            Width           =   7420
         End
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
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
            Index           =   6
            Left            =   0
            MaxLength       =   48
            TabIndex        =   6
            Text            =   " "
            Top             =   0
            Width           =   7420
         End
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
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
      Left            =   6970
      Picture         =   "FINFACTU109.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5355
      Width           =   1000
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Copias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   160
      TabIndex        =   17
      Top             =   5250
      Width           =   5580
      Begin VB.CheckBox Check8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Replica en Expedición"
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
         Left            =   3795
         TabIndex        =   69
         Top             =   315
         Width           =   1155
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   1
         Left            =   2835
         MaxLength       =   1
         TabIndex        =   12
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   0
         Left            =   1110
         MaxLength       =   1
         TabIndex        =   11
         Top             =   240
         Width           =   435
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   5040
         Picture         =   "FINFACTU109.frx":091E
         ToolTipText     =   "Doble-Click para Ingresar Texto Resumen del Remito"
         Top             =   315
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Remito:"
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
         Left            =   1890
         TabIndex        =   19
         Top             =   480
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
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
         Left            =   105
         TabIndex        =   18
         Top             =   450
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Observaciones en Factura"
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
      Left            =   7245
      TabIndex        =   15
      ToolTipText     =   "Doble Click Para Quitar Observaciones"
      Top             =   1050
      Visible         =   0   'False
      Width           =   7680
      Begin VB.PictureBox Picture1 
         Height          =   1470
         Left            =   105
         ScaleHeight     =   1410
         ScaleWidth      =   7365
         TabIndex        =   16
         Top             =   315
         Width           =   7420
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
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
            Index           =   5
            Left            =   0
            MaxLength       =   54
            TabIndex        =   5
            Text            =   " "
            Top             =   1150
            Width           =   7420
         End
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
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
            Index           =   4
            Left            =   0
            MaxLength       =   54
            TabIndex        =   4
            Text            =   " "
            Top             =   920
            Width           =   7420
         End
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
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
            Index           =   3
            Left            =   0
            MaxLength       =   54
            TabIndex        =   3
            Text            =   " "
            Top             =   690
            Width           =   7420
         End
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
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
            Index           =   2
            Left            =   0
            MaxLength       =   54
            TabIndex        =   2
            Text            =   " "
            Top             =   460
            Width           =   7420
         End
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
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
            Index           =   1
            Left            =   0
            MaxLength       =   54
            TabIndex        =   1
            Text            =   " "
            Top             =   230
            Width           =   7420
         End
         Begin VB.TextBox Text1 
            BorderStyle     =   0  'None
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
            Index           =   0
            Left            =   0
            MaxLength       =   54
            TabIndex        =   0
            Top             =   0
            Width           =   7420
         End
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Observaciones en Factura"
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
      Left            =   160
      TabIndex        =   34
      ToolTipText     =   "Doble Click Para Quitar Observaciones"
      Top             =   735
      Width           =   7820
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   315
         Width           =   7365
      End
   End
   Begin VB.Label Label13 
      Caption         =   "Label13"
      Height          =   255
      Left            =   2400
      TabIndex        =   71
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0E0FF&
      Caption         =   " codigo de barras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3240
      TabIndex        =   70
      Top             =   360
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label11 
      Height          =   225
      Left            =   7035
      TabIndex        =   64
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label CLAREDOL 
      Height          =   225
      Left            =   1365
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "Reajuste por Dolar"
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
      Left            =   6360
      TabIndex        =   29
      Top             =   4245
      Width           =   1170
   End
   Begin VB.Label Label5 
      Height          =   225
      Left            =   105
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   3255
      TabIndex        =   22
      Top             =   105
      Width           =   540
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   540
      Left            =   4095
      TabIndex        =   21
      Top             =   105
      Width           =   3795
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   " "
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
      Left            =   210
      TabIndex        =   20
      Top             =   315
      Width           =   3375
   End
End
Attribute VB_Name = "FINFACTU109"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub SOLICITARCAE_WSBFE()
    FLAG% = 0
    cont% = 0
    '
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    If EXISTE%(CARPETOK$ + "token_WSBFE.txt") = 0 Then
      TXT$ = REPLACAR$(CARPETOK$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'Token_WSBFE.txt'\en Carpeta " & TXT$)
      Exit Sub
    End If
    If EXISTE%(LUDAT$ + "consulta_cae_WSBFE_original.txt") = 0 Then
      TXT$ = REPLACAR$(LUDAT$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'consulta_cae_WSBFE_original.txt'\en Carpeta " & TXT$)
      Exit Sub
    End If
    NARCHI1% = FILEOPEN%(LUDAT$ + "consulta_cae_WSBFE_original.txt", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
    NARCHI2% = FILEOPEN%(CARPETOK$ + "consulta_cae_WSBFE.txt", 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA

    Do While Not EOF(NARCHI1%) ' RECORRE
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
       cont% = cont% + 1
       pos% = InStr(XX$, ">")
       ETIQUETACIERRE$ = TRIM$(Mid$(XX$, pos% + 1))
       TX$ = Mid$(XX$, 1, pos%)
       
          Select Case cont%
           
           Case 2: DATO$ = TIPO_DOC$      ' Tipo Documento (Cuit,Dni, etc) Identificador del comprador
           Case 3: DATO$ = NRO_DOC$      ' Numero Dni o Cuit
           Case 4: DATO$ = "1"   ' Numero de zona
           Case 5: DATO$ = TIPO_CBTE$        ' Tipo de Comprobante(Factura A o B ,Nota de Credito,Nota de Debito)
           Case 6: DATO$ = TRIM$(Str(XVALO(FORFACPRINT.Label12))) ' Punto de Venta
           Case 7: DATO$ = TRIM$(FORFACPRINT.Text11)            ' Numero de comprobante inicial
           Case 8: DATO$ = TRIM$(FORFACPRINT.Total(6))          ' Importe Total
           Case 9: DATO$ = IMP_TOT_CONC$            ' Importe no gravados = C/Iva, entre ellos neto no gravado, las percepciones de iva e IIBB
           Case 10: DATO$ = IMP_NETO$                ' Importe Neto (IMPORTE NETO = IMPORTE NETO - IVA NO GRAVADO)
           Case 11: DATO$ = IMPTO_LIQ$            ' Impuesto Liquidado(la suma de los importes de iva  aplicados)
           Case 12: DATO$ = "0"           ' Impto_Liq_rni - siempre es cero  o percepcion a no categorizados
           Case 13: DATO$ = "0"           ' importe operaciones externas (No dieron informacion de que se trata)
           Case 14: DATO$ = "0"
           Case 15: DATO$ = "0"
           Case 16: DATO$ = "0"
           Case 17: DATO$ = "0"
           Case 18: DATO$ = MONEDA_CAE$ ' QUEDA PARA AGREGAR UNA FUNCION POR EL TIPO DE CAMBIO
           Case 19: DATO$ = TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4"))
           
           Case 20:           '1)Fecha Comprobante,2)Fecha de comienzo de Servicio, 3)Fecha Fin de Servicio, Fecha de Vencimento
                    FECHA$ = FECHATEXLAR$(FECHANUM(FORFACPRINT.Text9))
                    DIA$ = Mid(FECHA$, 1, 2): MES$ = Mid$(FECHA$, 4, 2): ANIO$ = Mid$(FECHA$, 7, 4)
                    DATO$ = ANIO$ + MES$ + DIA$
          End Select
        TX$ = TX$ + DATO$ + ETIQUETACIERRE$ ' CONCATENA EL RENGLON A IMPRIMIR
        DATO$ = ""
        Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
     Loop
    '
    Print #NARCHI2%, "<Items>"

    For I& = 1 To ULTREG&("!CARDEFAC")
      XX$ = REGLEIDO$("!CARDEFAC", I&)
      CI% = CVI(Mid$(XX$, 5, 2))
      DESCRI$ = DESCRIT$(CI%)
      CAN = CVS(Mid$(XX$, 1, 4))
      PRECI# = CVS(Mid$(XX$, 429, 4))
      '
      PORDESCUEN = PORCEDESCU(FORFASQL.PORDESCU) 'MKS$(Mid$(XX$, 425, 4))
      IMPDESC = (PRECI# * PORDESCUEN) / 100
      PRECI# = ROUND(PRECI# - IMPDESC)
      '
      PRCI$ = FORMATNUM$(PRECI#, "F12.2")
      IMPORTE# = CAN * PRECI#
      IMPRT$ = FORMATNUM$(IMPORTE#, "F12.2")
      CONDI$ = "Codint = " & CI%
      NOMENC$ = VALOBADA("MASTER", "CODIGO_NCM", CONDI$, "NOMESS")
      COSEC$ = VALOBADA("MASTER", "CODIGO_SEC", CONDI$, "NOMESS")
      UNIMEX$ = VALOBADA("MASTER", "PRO_UMED", CONDI$, "NOMESS")
      COTIVA$ = VALOBADA("MASTER", "IVA_ID", CONDI$, "NOMESS")
      '
      Print #NARCHI2%, "<item>"
      Print #NARCHI2%, "<pro_codigo_ncm>" & TRIM$(NOMENC$) & "</pro_codigo_ncm>"
      Print #NARCHI2%, "<pro_codigo_sec>" & TRIM$(COSEC$) & "</pro_codigo_sec>"
      Print #NARCHI2%, "<pro_ds>" & TRIM$(DESCRI$) & "</pro_ds>"
      Print #NARCHI2%, "<pro_qty>" & CAN & "</pro_qty>"
      Print #NARCHI2%, "<pro_umed>" & TRIM$(UNIMEX$) & "</pro_umed>"
      Print #NARCHI2%, "<pro_precio_uni>" & TRIM$(PRCI$) & "</pro_precio_uni>"
      Print #NARCHI2%, "<imp_bonif>0</imp_bonif>"
      Print #NARCHI2%, "<imp_total>" & TRIM$(IMPRT$) & "</imp_total>"
      Print #NARCHI2%, "<iva_id>" & TRIM$(COTIVA$) & "</iva_id>"
      Print #NARCHI2%, "</item>"
    Next I&
    Print #NARCHI2%, "</Items>"
    Print #NARCHI2%, "</Cmp>"
    
    Close #NARCHI1%
    Close #NARCHI2%
    '
    On Error Resume Next
    Kill CARPETOK$ + "CAE.TXT" ' ELIMINAMOS EL ARCHIVO
    On Error GoTo 0
    '
    If EXISTE%(CARPETOK$ + "CAE.txt") = 0 Then
      If EXISTE%(CARPETOK$ + "wsbfe_cliente.exe") = 1 Then
         XXX = Shell(CARPETOK$ + "wsbfe_cliente.exe AUTO", 4)
         FLAG% = 1
      Else
         Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPETOK$, "\", "/") & "wsbfe_cliente.exe")
         GoTo 99
      End If
      
    End If

99 End Sub

Sub SOLICITARCAE_WSFEX()
    FLAG% = 0
    cont% = 0
    '
    'TOMO LOS DATOS DEL FRM DE INGRESO DE DATOS
    
    TIPO_EXPO$ = Left$(TRIM$(DATFAEX07.Combo1.TEXT), 1)
    DEST_COMP$ = Left$(TRIM$(DATFAEX07.Combo2.TEXT), 3)
    CUIT_PAIS$ = Left$(TRIM$(DATFAEX07.Combo4.TEXT), 11)
    INCOTS$ = TRIM$(DATFAEX07.Combo5.TEXT)
    IDIOMA$ = TRIM$(Left(DATFAEX07.Combo6.TEXT, 1))
    
    COMP_ASOCIAD% = DATFAEX07.List1.ListCount
    S_N$ = "N" 'VARIABLE QUE DEFINE SI TIENE PERMIOS O NO
    
    PERMISO% = DATFAEX07.List2.ListCount: If PERMISO% > 0 Then S_N$ = "S"
    ' EN EL CASO QUE SEA NOTA DE CREDITO O DEBITO ESTE DEBE ESTAR VACIO
    If FORFACPRINT.Label7.Caption = "N.Débito" Then S_N$ = ""
    If FORFACPRINT.Label7.Caption = "N.Crédito:" Then S_N$ = ""
    ' EN EL CASO QUE EL TIPO DE EXPORTACION NO SEA DE BIENES ESTE DEBE ESTAR VACIO
    If TIPO_EXPO$ <> "1" Then S_N$ = ""
    '
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    If EXISTE%(CARPETOK$ + "token_WSFEX.txt") = 0 Then
      TXT$ = REPLACAR$(CARPETOK$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'token_WSFEX.txt'\en Carpeta " & TXT$)
      Exit Sub
    End If
    If EXISTE%(LUDAT$ + "consulta_cae_WSFEX_original.txt") = 0 Then
      TXT$ = REPLACAR$(LUDAT$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'consulta_cae_WSFEX_original.txt'\en Carpeta " & TXT$)
      Exit Sub
    End If
    NARCHI1% = FILEOPEN%(LUDAT$ + "consulta_cae_WSFEX_original.txt", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
    NARCHI2% = FILEOPEN%(CARPETOK$ + "consulta_cae_WSFEX.txt", 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
    NC% = XVALO(FORFACPRINT.Text1)
    Do While Not EOF(NARCHI1%) ' RECORRE
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
       cont% = cont% + 1
       pos% = InStr(XX$, ">")
       ETIQUETACIERRE$ = TRIM$(Mid$(XX$, pos% + 1))
       TX$ = Mid$(XX$, 1, pos%)
       
          Select Case cont%
           Case 2
            FECHA$ = FECHATEXLAR$(FECHANUM(FORFACPRINT.Text9))
            DIA$ = Mid(FECHA$, 1, 2): MES$ = Mid$(FECHA$, 4, 2): ANIO$ = Mid$(FECHA$, 7, 4)
            DATO$ = ANIO$ + MES$ + DIA$
           Case 3: DATO$ = TIPOX_CBTEX$ ' TIPO DE COMPROBANTE
           Case 4: DATO$ = TRIM$(Str(XVALO(FORFACPRINT.Label12))) ' Punto de Venta
           Case 5: DATO$ = TRIM$(FORFACPRINT.Text11)            ' Numero de comprobante inicial
           Case 6: DATO$ = TIPO_EXPO$ ' TIPO DE EXPORTACION
           Case 7: DATO$ = S_N$ ' PERMISO EXISTENTE
           Case 9:
           If S_N$ = "S" Then
              For J% = 0 To DATFAEX07.List2.ListCount - 1
                PERMI$ = DATFAEX07.List2.List(J%)
                Print #NARCHI2%, "<permiso>"
                Print #NARCHI2%, "<id_permiso>" & TRIM$(PERMI$) & "</id_permiso>"
                Print #NARCHI2%, "<dst_merc>" & TRIM$(DEST_COMP$) & "</dst_merc>"
                Print #NARCHI2%, "</permiso>"
              Next J%
           End If
           Case 10: DATO$ = DEST_COMP$ ' DESTINO PAIS
           Case 11: DATO$ = TRIM$(RASOCLI$(NC%))
           Case 12: DATO$ = CUIT_PAIS$ ' CUIT DE PAIS
           Case 13: DATO$ = TRIM$(DOMICLI$(NC%)) 'DOMICILIO
           Case 14: DATO$ = TRIM$(CUITCLI$(NC%)) 'CUIT
           Case 15: DATO$ = MONEDA_CAE$ ' TIPO DE MONEDA
           Case 16: DATO$ = TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) 'COTIZACION
           Case 17: DATO$ = Left$(Text4, 2000) ' OBSERVACIONES
           Case 18: DATO$ = TRIM$(FORFACPRINT.Total(6)) ' Importe Total
           Case 19: DATO$ = Left$(Text4, 2000) ' OBSERVACIONES
           Case 21:
              If COMP_ASOCIAD% > 0 Then
                For J% = 0 To DATFAEX07.List1.ListCount - 1
                 RELEI$ = DATFAEX07.List1.List(J%)
                 CBTE$ = Mid$(TRIM$(RELEI$), 1, 2)
                 PVTA$ = Mid$(RELEI$, 20, 4)
                 NURO$ = Mid$(RELEI$, 32, 8)
                 Print #NARCHI2%, "<cmp_asoc>"
                 Print #NARCHI2%, "<cbte_tipo>" & TRIM$(CBTE$) & "</cbte_tipo>"
                 Print #NARCHI2%, "<cbte_punto_vta>" & TRIM$(PVTA$) & "</cbte_punto_vta>"
                 Print #NARCHI2%, "<cbte_nro>" & TRIM$(NURO$) & "</cbte_nro>"
                 Print #NARCHI2%, "</cmp_asoc>"
                Next J%
              End If
           Case 22: DATO$ = TRIM$(Left$(FORFACPRINT.Eco21(1), 50)) ' FORMA DE PAGO
           Case 23: DATO$ = INCOTS$ ' TIPO DE VENTA FOB,CIF,ETC
           FORFACPRINT.Text18 = INCOTS$ ' ACTUALIZO EL FRM
           Case 25: DATO$ = IDIOMA$ '1-ESPAÑOL 2-INGLES 3-PORTUGUES
           
          End Select
        TX$ = TX$ + DATO$ + ETIQUETACIERRE$ ' CONCATENA EL RENGLON A IMPRIMIR
        DATO$ = ""
        Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
        
     Loop
    '
    Print #NARCHI2%, "<Items>"

    For I& = 1 To ULTREG&("!CARDEFAC")
      XX$ = REGLEIDO$("!CARDEFAC", I&)
      CI% = CVI(Mid$(XX$, 5, 2))
      DESCRI$ = DESCRIT$(CI%)
      CAN = CVS(Mid$(XX$, 1, 4))
      PRECI# = CVS(Mid$(XX$, 429, 4))
      '
      PORDESCUEN = PORCEDESCU(FORFACPRINT.PORDESCU) 'MKS$(Mid$(XX$, 425, 4))
      IMPDESC = (PRECI# * PORDESCUEN) / 100
      PRECI# = ROUND(PRECI# - IMPDESC, 3)
      '
      PRCI$ = FORMATNUM$(PRECI#, "F12.3")
      IMPORTE# = CAN * PRECI#
      IMPRT$ = FORMATNUM$(IMPORTE#, "F12.3")
      CONDI$ = "Codint = " & CI%
      NOMENC$ = VALOBADA("MASTER", "CODIGO_NCM", CONDI$, "NOMESS")
'      COSEC$ = VALOBADA("MASTER", "CODIGO_SEC", CONDI$, "NOMESS")
      UNIMEX$ = TRIM$(VALOBADA("MASTER", "PRO_UMED", CONDI$, "NOMESS")) 'REVISA SI ESTA EN BASE DE DATOS
      If UNIMEX$ = "" Then UNIMEX$ = Str$(UNIMEAFIP%(CI%)) ' SI NO LO TRAE DEL ARCHIVO DE UNIMED
'      COTIVA$ = VALOBADA("MASTER", "IVA_ID", CONDI$, "NOMESS")
      '
      Print #NARCHI2%, "<item>"
      Print #NARCHI2%, "<pro_codigo>" & TRIM$(NOMENC$) & "</pro_codigo>"
      Print #NARCHI2%, "<pro_ds>" & TRIM$(DESCRI$) & "</pro_ds>"
      Print #NARCHI2%, "<pro_qty>" & CAN & "</pro_qty>"
      Print #NARCHI2%, "<pro_umed>" & TRIM$(UNIMEX$) & "</pro_umed>"
      Print #NARCHI2%, "<pro_precio_uni>" & TRIM$(PRCI$) & "</pro_precio_uni>"
      Print #NARCHI2%, "<imp_total>" & TRIM$(IMPRT$) & "</imp_total>"
      Print #NARCHI2%, "</item>"
    Next I&
    Print #NARCHI2%, "</Items>"
    Print #NARCHI2%, "</Cmp>"
    
    Do While Not EOF(NARCHI2%) ' RECORRE
       Line Input #NARCHI2%, TTXX$ ' LEE CADA REGISTRO
       TTX1$ = TTXX$ + vbCrLf
    Loop
    Call ABRECABAN
    Call ACTUREGISTRO("CAJABANC", "TEXTOCAE", "CODICOM_LAR = '" & DOCUM$ & "'", TTX1$, REGF&, "NOMESS")
    
    Close #NARCHI1%
    Close #NARCHI2%
    '
    On Error Resume Next
    Kill CARPETOK$ + "CAE.TXT" ' ELIMINAMOS EL ARCHIVO
    On Error GoTo 0
    '
'    If EXISTE%(CARPETOK$ + "CAE.txt") = 0 Then
'      If EXISTE%(CARPETOK$ + "wsfex_cliente.exe") = 1 Then
'         XXX = Shell(CARPETOK$ + "wsfex_cliente.exe AUTO", 4)
'         FLAG% = 1
'      Else
'         Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPETOK$, "\", "/") & "wsfex_cliente.exe")
'         GoTo 99
'      End If
'    End If


99 End Sub


Private Sub Command1_Click()
   '\\\OT-07-0023-RG-24/01/07///
   If XVALO(Text2(15)) > 0 Then
      If CONTROL$("SVD451S", "CUEVE-13") = "" Then
          Call MENSERR(24, "Para Imputar Gastos de Exportación Debe Primero\Configurar la Cuenta Contable.")
          Exit Sub
      End If
      If Text10 = "" Then
          Call MENSERR(24, "Debe Ingresar Referencia Para los Gastos de Exportación.")
          Exit Sub
      End If
   End If
   If Text10 <> "" And XVALO(Text2(15)) < 0.005 Then
      Call MENSERR(24, "Importe no Valido para Gastos de Exportación.")
      Exit Sub
   End If
   '
   DEP% = XVALO(Text7)
   If ECOARCH$("TADEPOSI", Chr(DEP%)) <> "" Then
      FORFINFAC.Label5.Caption = "APRO"
      Hide
   Else
      Call MENSERR(24, "Falta Ingresar el Depósito de Consumo")
      Text7 = TRIM$(Str$(DEPOEXPE%))
      Text7.SetFocus
   End If
   '
   
   '\\\OT-07-0023-RG-FIN///
End Sub

Private Sub Command2_Click()
   FORFINFAC.Label5.Caption = "CANCE"
   Hide
End Sub

Private Sub Command3_Click()
    '\\\OT-07-0023-RG-24/01/07///
   If DERACCE%("MODEFACT", "Modificacion de Deposito de Facturacion") >= 2 Then
        Call SELECHO("TADEPOSI")
        DEP$ = VALACT1$("TADEPOSI")
        If DEP$ <> "" Then
           Text7.TEXT = DEP$
        End If
    Else
        Text4.SetFocus
    End If

    '\\\OT-07-0023-RG-FIN///
End Sub

Private Sub Form_Load()
   If CONTROL$("FACTUREM", "FORMAREM") = "COMPACTO" Then
     Label1(1).Left = 1890
     Text2(1).Left = 2835
     Image1.Visible = True
     Check8.Left = 3795
   End If
   '
   Check8.Value = 0 ' CONTROL PARA IMPRIMIR COPIA DE REMITO EN IMPRESORA DE DESPACHO
   If CONTROL("FACTURA", "REPLIREM") = "SI" Then Check8.Value = 1
   '
End Sub

Private Sub Frame1_DblClick()
   For U& = 0 To 5
     Text1(U&) = ""
   Next U&
End Sub

Private Sub Frame3_DblClick()
   For U& = 6 To 8
     Text1(U&) = ""
   Next U&
End Sub

Private Sub Image1_DblClick()
   Frame5.Visible = True
   Text3 = "Bultos: " + TRIM$(Text2(2)) + Chr$(13) + Chr$(10) + "Valor Declarado: " + TRIM$(Text2(3))
   Text3.SetFocus
End Sub



Private Sub Label6_DblClick()
Exit Sub
   If CLAREDOL = "0" Then
      TTXX$ = "LOS PRECIOS DE LA PRESENTE FACTURA EQUIVALEN"
      TTXX$ = TTXX$ + " A LA PARIDAD DE U$$ 1.00 = $ " + TRIM$(FORFACLI.TICAMBIS) + "."
      TTXX$ = TTXX$ + " SERA CANCELADA SEGUN COTIZACION DOLAR VENDEDOR BNA, CIERRE 24 HS."
      TTXX$ = TTXX$ + " ANTES DE LA FECHA DE PAGO."
      '
      Call FRATEXTO(TTXX$, 54)
      For KKII% = 1 To CARETEX%
        FORFINFAC.Text1(KKII% - 1) = RETEX$(KKII%)
      Next KKII%
      CLAREDOL = "1"
    Else
      For KKII% = 1 To 8
        FORFINFAC.Text1(KKII% - 1) = ""
      Next KKII%
      CLAREDOL = "0"
   End If
End Sub

Private Sub Text2_gotfocus(Index As Integer)
   Text2(Index).SelStart = 0
   Text2(Index).SelLength = Len(Text2(Index).TEXT)
End Sub
'\\\OT-06-0125-WAR-05/04/06///
Private Sub Text2_LostFocus(Index As Integer)
  If Index = 4 Or Index = 11 Then
    NT2 = XVALO(Text2(Index))
    Text2(Index) = FORMATNUM$(NT2, "F9.1")
  End If
  If Index = 12 Or Index = 13 Or Index = 15 Then
    NT2 = XVALO(Text2(Index))
    Text2(Index) = FORMATNUM$(NT2, "F10.2")
  End If
End Sub
'\\\OT-06-0125-WAR-FIN///
Private Sub Text7_Change()
  '\\\OT-07-0023-RG-24/01/07///
  If Text7 = "" Then Label8.Caption = ""
     On Error Resume Next
     If Text7 <> "" Then
        DEP% = XVALO(Text7)
        Label8.Caption = Mid$(ECOARCH$("TADEPOSI", Chr(DEP%)), 1, 13)
     End If
  '\\\OT-07-0023-RG-FIN///
End Sub

Private Sub Text7_DblClick()
   '\\\OT-07-0023-RG-24/01/07///
   Call Command3_Click
   '\\\OT-07-0023-RG-FIN///
End Sub

Private Sub Text7_GotFocus()
   If DERACCE%("MODEFACT", "Modificacion de Deposito de Facturacion") < 2 Then
     Text4.SetFocus
   End If
End Sub

Private Sub Timer2_Timer()
   Static CONTA%
   CONTA% = CONTA% + 1 ' SI SUPERA LOS 10 SEGUNDOS DE DEMORA EN TRAER EL CAE PRESENTA MENSAJE
   If CONTA% > 15 Then
      If COMALTER%("No Se Ha Podido Obtener el CAE \\Seguir Intentando\Cancelar Factura") <> 1 Then
         Call FINAL("")
      End If
      CONTA% = 0
   End If
   KAE$ = LEECAE$(FVENCIM$)
   IDX$ = LeeId$
'   If KAE$ <> "" Then
'      Timer2.Enabled = False
'      Label11 = KAE$
'      Command2.Enabled = False
'      Command1.Enabled = True
'      Text4 = "CAE: " + KAE$ + "................Id.Tran: " + IDX$ + " F.Vto.: " + FVENCIM$ + vbCrLf + Text4
'      Text2(0).Enabled = True
'      Command2.Enabled = True
'      Call Command1_Click    ' PARA APROBARLO DIRECTAMENTE
'      Command1.Top = 5355
'      Command1.Height = 790
'      Command4.Enabled = True
'      Screen.MousePointer = 1
'   End If
End Sub

Function LEECAE$(FVENCIM$)
    '
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    LEECAE$ = ""
    CAE$ = ""
    If EXISTE%(CARPETOK$ + "CAE.TXT") > 0 Then
        NARCHIVO% = FILEOPEN%(CARPETOK$ + "CAE.TXT", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
        '
        Do While Not EOF(NARCHIVO%) ' RECORRE
          Line Input #NARCHIVO%, XXXXX$ ' LEE CADA REGISTRO
          If InStr(UCase$(XXXXX$), UCase$("<cae>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<cae>")) + 5
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             CAE$ = Mid$(TXT$, 1, POS2% - 1)
             'Se comenta para que no salga una vez obtenido el CAE
             'Exit Do
          End If
          ' Obtiene fecha de vencimiento
          'Esto es para obtener la fecha de vencimiento de la factura de exportacion o bienes de Capital
          If InStr(UCase$(XXXXX$), UCase$("<Fch_venc_Cae>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<Fch_venc_Cae>")) + 14
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             VENCIMIENTO$ = Mid$(TXT$, 1, POS2% - 1)
             FVENCIM$ = VENCIMIENTO$
             Exit Do
          End If
          '
          'Esto es para obtener la fecha de vencimiento de la factura comun
          If InStr(UCase$(XXXXX$), UCase$("<fecha_vto>")) > 0 Then
             POS1% = InStr(UCase$(XXXXX$), UCase$("<fecha_vto>")) + 11
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             VENCIMIENTO$ = Mid$(TXT$, 1, POS2% - 1)
             FVENCIM$ = VENCIMIENTO$
             Exit Do
          End If
          '
          

        Loop
        '
        Close NARCHIVO%
        LEECAE$ = CAE$
        NumeroCae$ = LEECAE$
        Timer2.Enabled = False
        '
        If XVALO(CAE$) < 1 Then
          MOTIVO1$ = ""
          MOTIVO2$ = ""
          VUELTA% = 0 ' LECTURA DE MOTIVO DEL ARCHIVO CAE.TXT
          NARCHIVO% = FILEOPEN%(CARPETOK$ + "CAE.TXT", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
          Do While Not EOF(NARCHIVO%) ' RECORRE
              Line Input #NARCHIVO%, XX$ ' LEE CADA REGISTRO
              If InStr(XX$, "<motivo>") > 0 Or InStr(XX$, "<Obs>") Then
                 POS1% = InStr(XX$, ">")
                 POS2% = InStr(XX$, "</")
                 HASTA% = POS2% - (POS1% + 1)
                 If VUELTA% < 1 Then ' SI ES LA PRIMERA VEZ ASIGNA MOTIVO1
                  MOTIVO1$ = Mid$(XX$, POS1% + 1, HASTA%)
                  VUELTA% = VUELTA% + 1
                 ElseIf VUELTA% > 0 Then ' LA SEGUNDA VEZ QUE PASA CARGA MOTIVO2
                  MOTIVO2$ = "-" & Mid$(XX$, POS1% + 1, HASTA%)
                 End If
              End If
          Loop
          Close NARCHIVO%
          MOTIVO$ = MOTIVO1$ & MOTIVO2$ ' ASIGNA LOS DOS MOTIVOS
          '
          Screen.MousePointer = 1
          Call COMUNI("CAE Devuelto No Válido.\Motivo: " & MOTIVO$ & "\Debe Re-Lanzar el Módulo de Facturación.")
          LEECAE$ = ""
          Hide
          Unload Me
          On Error Resume Next
          Call FINAL("")
          On Error GoTo 0
          Close: End
        End If
        '
    End If
    '
    ' Armar codigo de barras
    '
    ' Obtiene CUIT
    N1% = FILEOPEN(LUDAT$ + "SVD147.EXE", 0, 41)
    Dim X As String * 41
    Get #N1%, 6, X
    CUITT$ = TRIM$(REPLACAR$(TRIM$(DESENCRI$(X)), "-", ""))
    Close #N1%
    ' Tipo de comrpobante
    COMPROBANTE$ = TRIM$(TIPOX_CBTEX$)
    ' Punto de venta
    PUNTOVENTA$ = TRIM$(Mid$(Label3, 1, 4))
    ' CAE
    CODIGOAUTORIZACION$ = TRIM$(CAE$)
    ' Fecha de vencimiento
    FECHAVENCIMIENTO$ = TRIM$(VENCIMIENTO$)
    ' Arma cadena
    XX$ = CUITT$ + COMPROBANTE$ + PUNTOVENTA$ + CODIGOAUTORIZACION$ + FECHAVENCIMIENTO$
    ' Obtiene digito verificador
    DIGITOVERIFICADOR$ = OBTIENEDIGITO$(XX$)
    ' Arma cadena final
    XX$ = XX$ + DIGITOVERIFICADOR$
    ' Convierte al formato del codigo de barras
    YY$ = CONVIERTOCODIGOBARRAS$(XX$)
    '
    Label12 = YY$
    '
End Function




Sub Command4_Click()
    '
    Command4.Enabled = False
    If XVALO(Text2(0)) < 1 Then
         Call MENSERR(24, "No Corresponde C.A.E. Para Solo Remito.")
         Timer2.Enabled = False
         Label11 = ""
         Command1.Enabled = True
         Command2.Enabled = True
         Command4.Enabled = True
       Else
         Text2(0).Enabled = False
         Command2.Enabled = False
         Command4.Enabled = False
         Screen.MousePointer = 11
         '
         SALPUVE% = FORFACPRINT.VERIFSALIDA%
         If SALPUVE% = 1 Then
             Call SOLICITARCAE
             Timer2.Enabled = True
           ElseIf SALPUVE% = 2 Then
             Call SOLICITARCAE_WSBFE
             Timer2.Enabled = True
           ElseIf SALPUVE% = 3 Then
             Screen.MousePointer = 1
             DATFAEX07.Show 1
             If DATFAEX07.Label17 = "OK" Then
               Call SOLICITARCAE_WSFEX
               Timer2.Enabled = True
             Else
               Command2.Enabled = True
               Command4.Enabled = True
             End If
               'Command4.Enabled = True
           Else
             Exit Sub
         End If
    End If
    'Command4.Enabled = True   ' que no se pueda hacer dos veces
End Sub

Sub SOLICITARCAE()
    FLAG% = 0
    cont% = 0
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    If EXISTE%(CARPETOK$ + "token.txt") = 0 Then
      TXT$ = REPLACAR$(CARPETOK$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'Token.txt'\en Carpeta " & TXT$)
      GoTo 99
    End If
    If EXISTE%(LUDAT$ + "consulta_cae_original.txt") = 0 Then
      TXT$ = REPLACAR$(LUDAT$, "\", "/")
      Call MENSERR(24, "Falta Archivo 'consulta_cae_original.txt'\en Carpeta " & TXT$)
      GoTo 99
    End If
    NARCHI1% = FILEOPEN%(LUDAT$ + "consulta_cae_original.txt", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
    NARCHI2% = FILEOPEN%(CARPETOK$ + "consulta_cae.txt", 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA

    Do While Not EOF(NARCHI1%) ' RECORRE
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
       cont% = cont% + 1
       pos% = InStr(XX$, ">")
       ETIQUETACIERRE$ = TRIM$(Mid$(XX$, pos% + 1))
       TX$ = Mid$(XX$, 1, pos%)
       
          Select Case cont%
           'Case 4: DATO$ = TRIM$(FORFACPRINT.Label18)          ' Cantidad de registros del detalle de la factura
           Case 4: DATO$ = "1"            ' corregido porque debe ser siempre 1
           Case 5: DATO$ = "1"            ' Tipo de Presentacion dada (0 servicios , 1 Bienes de Capital)
           Case 9: DATO$ = TIPO_DOC$      ' Tipo Documento (Cuit,Dni, etc) Identificador del comprador
           Case 10: DATO$ = NRO_DOC$      ' Numero Dni o Cuit
           Case 11: DATO$ = TIPO_CBTE$        ' Tipo de Comprobante(Factura A o B ,Nota de Credito,Nota de Debito)
           Case 12: DATO$ = TRIM$(Str(XVALO(FORFACPRINT.Label12))) ' Punto de Venta
           Case 13: DATO$ = TRIM$(FORFACPRINT.Text11)            ' Numero de comprobante inicial
           Case 14: DATO$ = TRIM$(FORFACPRINT.Text11)            ' Numero de comprobante hasta
           Case 15: DATO$ = TRIM$(FORFACPRINT.Total(6))          ' Importe Total
           Case 16: DATO$ = IMP_TOT_CONC$            ' Importe no gravados = C/Iva, entre ellos neto no gravado, las percepciones de iva e IIBB
           Case 17: DATO$ = IMP_NETO$                ' Importe Neto (IMPORTE NETO = IMPORTE NETO - IVA NO GRAVADO)
           Case 18: DATO$ = IMPTO_LIQ$            ' Impuesto Liquidado(la suma de los importes de iva  aplicados)
           Case 19: DATO$ = "0"           ' Impto_Liq_rni - siempre es cero  o percepcion a no categorizados
           Case 20: DATO$ = "0"           ' importe operaciones externas (No dieron informacion de que se trata)
           Case 21, 22, 23, 24:           '1)Fecha Comprobante,2)Fecha de comienzo de Servicio, 3)Fecha Fin de Servicio, Fecha de Vencimento
                    FECHA$ = FECHATEXLAR$(FECHANUM(FORFACPRINT.Text9))
                    DIA$ = Mid(FECHA$, 1, 2): MES$ = Mid$(FECHA$, 4, 2): ANIO$ = Mid$(FECHA$, 7, 4)
                    DATO$ = ANIO$ + MES$ + DIA$
          End Select
          ' NOTA !!!!************************************************
          ' EN LA FILA 24 SE DEBE PONER LA FECHA DE VENCIMIENTO DEL COMPROBANTE
          ' PROVISORIAMENTE SE PUSO LA MISMA DEL DIA
          '*********************************************************************
        TX$ = TX$ + DATO$ + ETIQUETACIERRE$ ' CONCATENA EL RENGLON A IMPRIMIR
        DATO$ = ""
        Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
     Loop
    '
    Close #NARCHI1%
    Close #NARCHI2%
    '
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    On Error Resume Next
    Kill CARPETOK$ + "CAE.TXT" ' ELIMINAMOS EL ARCHIVO
    On Error GoTo 0
    '
    If EXISTE%(CARPETOK$ + "CAE.txt") = 0 Then
      If EXISTE%(CARPETOK$ + "wsfe_cliente.exe") = 1 Then
         XXX = Shell(CARPETOK$ + "wsfe_cliente.exe AUTO", 4)
         FLAG% = 1
      Else
         Call COMUNI("Falta Ejecutable: " & REPLACAR$(CARPETOK$, "\", "/") & "wsfe_cliente.exe")
         GoTo 99
      End If
    End If
    '
99 End Sub

Function TIPO_DOC$()
    ' TIPO DE DOCUMENTO
    NC% = XVALO(FORFACPRINT.Text1)
    POSDEIVA% = PIVACLI%(NC)
    Select Case POSDEIVA%
        Case 0: VALOR$ = 96
        Case 1: VALOR$ = 80
        Case 2: VALOR$ = 80
        Case 3: VALOR$ = 80
        Case 4: VALOR$ = 80
        Case 5: VALOR$ = 80
        Case 6: VALOR$ = 80
        Case 7: VALOR$ = 80
    End Select
    TIPO_DOC$ = VALOR$
End Function

Function MONEDA_CAE$()
    '
    MONE$ = ""
    If MONEMI% = 1 Then
         MONE$ = "PES"
      ElseIf MONEMI% = 2 Then
         MONE$ = "DOL"
      Else
         DEMONE$ = UCase(DEMONECO$(MONEMI%))
         If Left$(DEMONE$, 4) = "EURO" Then
              MONE$ = "060"
           ElseIf Left$(DEMONE$, 4) = "REAL" Then
              MONE$ = "012"
         End If
    End If
    '
99  MONEDA_CAE$ = MONE$
    '
End Function

Function TIPO_CBTE$()

    TIPOCOMP$ = UCase$(FORFACPRINT.Label7 + FORFACPRINT.Label13)
    
    Select Case TIPOCOMP$
        Case "FACTURA:A": VALOR$ = "1"
        Case "N.DÉBITO:A": VALOR$ = "2"
        Case "N.CRÉDITO:A": VALOR$ = "3"
        Case "FACTURA:B": VALOR$ = "6"
        Case "N.DÉBITO:B": VALOR$ = "7"
        Case "N.CRÉDITO:B": VALOR$ = "8"
    End Select
    '
    TIPO_CBTE$ = VALOR$
    '
End Function
Function TIPOX_CBTEX$()

    TIPOCOMP$ = UCase$(FORFACPRINT.Label7 + FORFACPRINT.Label13)
    
    Select Case TIPOCOMP$
        Case "FACTURA:E": VALOR$ = "19"
        Case "N.DÉBITO:E": VALOR$ = "20"
        Case "N.CRÉDITO:E": VALOR$ = "21"
    End Select
    '
    TIPOX_CBTEX$ = VALOR$
    '
End Function





Function NRO_DOC$()
    NRO_DOC$ = ""
    POXIVA% = PIVACLI%(TRIM$(FORFACPRINT.Text1))
    '
    NUMERO_CUIT_DOC1$ = CUITCLI$(TRIM$(FORFACPRINT.Text1))
    NUMERO_CUIT_DOC$ = REPLACAR(NUMERO_CUIT_DOC1$, "-", "")
    '
    If POXIVA% < 1 And TRIM$(NUMERO_CUIT_DOC$) = "" Then NUMERO_CUIT_DOC$ = "0"
    '
    If Not IsNumeric(NUMERO_CUIT_DOC$) Then
      Call MENSERR(24, "Documento Nro.: " & NUMERO_CUIT_DOC$ & " No Válido\El Dato debe ser Numérico\Modifique y Vuelva a Cargar la Factura ")
      Call FINAL("")
    End If
    '
    NRO_DOC$ = NUMERO_CUIT_DOC$
    '
End Function

Function IMPTO_LIQ$()
  IMPTO_LIQ$ = ""
  IMPTOLIQUID$ = FORMATNUM$(XVALO(FORFACPRINT.Total(3)) + XVALO(FORFACPRINT.Total(4)) + XVALO(FORFACPRINT.Total(5)), "F12.2")
  IMPTO_LIQ$ = TRIM$(IMPTOLIQUID$)
End Function

Function IMP_TOT_CONC$()

   IMP_TOT_CONC$ = "" ' SUMA DE IVA NO GRAVADO + PERCEPCION DE IVA + PERCEPCION DE INGRESOS BRUTOS
   IMPTOTACONC$ = FORMATNUM$(XVALO(FORFACPRINT.TLABEL28) + XVALO(FORFACPRINT.Text17) + XVALO(FORFACPRINT.Text15), "F12.2")
   If TRIM$(IMPTOTACONC$) = "" Then IMPTOTACONC$ = "0"
   IMP_TOT_CONC$ = TRIM$(IMPTOTACONC$)
   
End Function

Function IMP_NETO$()
   IMP_NETO$ = ""
   IMPORNETO$ = FORMATNUM$(XVALO(FORFACPRINT.Total(2)) - XVALO(FORFACPRINT.TLABEL28), "F12.2")
   IMP_NETO$ = TRIM$(IMPORNETO$)
End Function

Function LeeId$()
    '
    CARPETOK$ = "C:\FLEXOFT\FACELEC\"
    LeeId$ = ""
    If EXISTE%(CARPETOK$ + "CAE.TXT") > 0 Then
        NARCHIVO% = FILEOPEN%(CARPETOK$ + "CAE.TXT", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
        '
        Do While Not EOF(NARCHIVO%)       ' RECORRE
          Line Input #NARCHIVO%, XXXXX$   ' LEE CADA REGISTRO
          If InStr(UCase$(XXXXX$), "<ID>") > 0 Then
             POS1% = InStr(UCase$(XXXXX$), "<ID>") + 4
             TXT$ = Mid$(XXXXX$, POS1%)
             POS2% = InStr(TXT$, "</")
             ID1$ = Mid$(TXT$, 1, POS2% - 1)
          End If
        Loop
        '
        Close NARCHIVO%
        LeeId$ = ID1$
   End If
   IdentiCae$ = LeeId$
   '
End Function

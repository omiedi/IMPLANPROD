VERSION 5.00
Begin VB.Form FINFACTURA 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aprobación de Comprobante"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8130
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFC0&
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
         BackColor       =   &H00FFFFC0&
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
         Left            =   210
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
      BackColor       =   &H00FFFFC0&
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
      Height          =   1860
      Left            =   160
      TabIndex        =   40
      Top             =   6300
      Width           =   7820
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
         Top             =   1470
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
         Top             =   1470
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
         BackColor       =   &H00FFFFC0&
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
         Top             =   1530
         Width           =   1680
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Top             =   1530
         Width           =   1860
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         BackColor       =   &H00FFFFC0&
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
         BackColor       =   &H00FFFFC0&
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
         BackColor       =   &H00FFFFC0&
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
         BackColor       =   &H00FFFFC0&
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
         BackColor       =   &H00FFFFC0&
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
      BackColor       =   &H00FFFFC0&
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
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFC0&
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
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
         BackColor       =   &H00FFFFC0&
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
         Height          =   225
         Index           =   2
         Left            =   3165
         TabIndex        =   27
         Top             =   525
         Width           =   1485
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Height          =   225
         Left            =   255
         TabIndex        =   33
         Top             =   945
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
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
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   28
         Top             =   525
         Width           =   1440
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Height          =   225
         Left            =   2535
         TabIndex        =   38
         Top             =   945
         Width           =   1485
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
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
      Height          =   840
      Left            =   6970
      Picture         =   "FINFACTURA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5355
      Width           =   1000
   End
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
      Height          =   840
      Left            =   5880
      Picture         =   "FINFACTURA.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5355
      Width           =   1000
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
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
      Height          =   960
      Left            =   160
      TabIndex        =   17
      Top             =   5250
      Width           =   5580
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
         Left            =   4725
         MaxLength       =   1
         TabIndex        =   12
         Top             =   315
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
         Left            =   2415
         MaxLength       =   1
         TabIndex        =   11
         Top             =   315
         Width           =   435
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   5040
         Picture         =   "FINFACTURA.frx":0614
         ToolTipText     =   "Doble-Click para Ingresar Texto Resumen del Remito"
         Top             =   315
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   3465
         TabIndex        =   19
         Top             =   525
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Top             =   525
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
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
      BackColor       =   &H00FFFFC0&
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
Attribute VB_Name = "FINFACTURA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '\\\OT-07-0023-RG-24/01/07///
   DEP% = XVALO(Text7)
   If ECOARCH$("TADEPOSI", Chr(DEP%)) <> "" Then
      FINFACTURA.Label5.Caption = "APRO"
      Hide
   Else
      Call MENSERR(24, "Falta Ingresar el Depósito de Consumo")
      Text7 = TRIM$(Str$(DEPOEXPE%))
      Text7.SetFocus
   End If
   '\\\OT-07-0023-RG-FIN///
End Sub

Private Sub Command2_Click()
   FINFACTURA.Label5.Caption = "CANCE"
   Hide
End Sub

Private Sub Command3_Click()
    '\\\OT-07-0023-RG-24/01/07///
    Call SELECHO("TADEPOSI")
    DEP$ = VALACT1$("TADEPOSI")
    If DEP$ <> "" Then
       Text7.TEXT = DEP$
    End If
    '\\\OT-07-0023-RG-FIN///
End Sub


Private Sub Form_Load()

    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Frame7.BackColor = &HFCD7B6
      Me.Frame8.BackColor = &HFCD7B6
      Me.Frame9.BackColor = &HFCD7B6
      Me.Label1(0).BackColor = &HFCD7B6
      Me.Label1(1).BackColor = &HFCD7B6
      Me.Label1(2).BackColor = &HFCD7B6
      Me.Label1(3).BackColor = &HFCD7B6
      Me.Label1(4).BackColor = &HFCD7B6
      Me.Label1(5).BackColor = &HFCD7B6
      Me.Label1(6).BackColor = &HFCD7B6
      Me.Label1(7).BackColor = &HFCD7B6
      Me.Label1(8).BackColor = &HFCD7B6
      Me.Label1(9).BackColor = &HFCD7B6
      Me.Label1(10).BackColor = &HFCD7B6
      Me.Label1(11).BackColor = &HFCD7B6
      Me.Label2.BackColor = &HFCD7B6
      Me.Label3.BackColor = &HFCD7B6
      Me.Label4.BackColor = &HFCD7B6
      Me.Label5.BackColor = &HFCD7B6
      Me.Label6.BackColor = &HFCD7B6
      Me.Label7.BackColor = &HFCD7B6
      Me.Label8.BackColor = &HFCD7B6
      Me.Label9.BackColor = &HFCD7B6
      Me.Label10.BackColor = &HFCD7B6
    End If

   If CONTROL$("FACTUREM", "FORMAREM") = "COMPACTO" Then
     Label1(1).Left = 3165
     Text2(1).Left = 4425
     Image1.Visible = True
   End If
   Text7.TEXT = CONTROL$("", "DEPOEXPE")
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

Private Sub image1_DblClick()
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
        FINFACTURA.Text1(KKII% - 1) = RETEX$(KKII%)
      Next KKII%
      CLAREDOL = "1"
    Else
      For KKII% = 1 To 8
        FINFACTURA.Text1(KKII% - 1) = ""
      Next KKII%
      CLAREDOL = "0"
   End If
End Sub

Private Sub Text2_DblClick(Index As Integer)
   On Error Resume Next
   Call FORFACCS.PRIFACTURA
   If Not (Err) Then
      On Error GoTo 0
      Call COMUNI("Se ha Emitido Copia Impresa\SIN GRABAR LA FACTURA.")
   End If
   On Error GoTo 0
End Sub

Private Sub Text2_gotfocus(Index As Integer)
   Text2(Index).SelStart = 0
   Text2(Index).SelLength = Len(Text2(Index).TEXT)
End Sub
'\\\OT-06-0125-WAR-05/04/06///
Private Sub TEXT2_LOSTFOCUS(Index As Integer)
  If Index = 4 Or Index = 11 Then
    NT2 = XVALO(Text2(Index))
    Text2(Index) = FORMATNUM$(NT2, "F9.1")
  End If
  If Index = 12 Or Index = 13 Then
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

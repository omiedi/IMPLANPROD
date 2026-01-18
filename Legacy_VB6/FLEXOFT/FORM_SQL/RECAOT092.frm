VERSION 5.00
Begin VB.Form RECAOT092 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción y Etiquetado"
   ClientHeight    =   6255
   ClientLeft      =   4620
   ClientTop       =   330
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "!"
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
      Left            =   3855
      TabIndex        =   66
      Top             =   1695
      Width           =   175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "!"
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
      Left            =   6180
      TabIndex        =   20
      Top             =   5415
      Width           =   175
   End
   Begin VB.TextBox Text13 
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
      Left            =   2955
      MaxLength       =   11
      TabIndex        =   62
      Top             =   5430
      Width           =   975
   End
   Begin VB.TextBox Text12 
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
      Left            =   5130
      MaxLength       =   16
      TabIndex        =   17
      Top             =   4905
      Width           =   1215
   End
   Begin VB.TextBox Text11 
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
      Left            =   1350
      MaxLength       =   13
      TabIndex        =   6
      Top             =   4470
      Width           =   1215
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
      Left            =   1365
      MaxLength       =   11
      TabIndex        =   8
      Top             =   5415
      Width           =   855
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
      Height          =   780
      Left            =   6615
      Picture         =   "RECAOT092.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Graba - Guardar Modificaciones"
      Top             =   5385
      Width           =   915
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
      Left            =   4920
      TabIndex        =   15
      Top             =   4050
      Width           =   1215
   End
   Begin VB.CommandButton Command18 
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
      Left            =   6135
      TabIndex        =   16
      Top             =   4050
      Width           =   175
   End
   Begin VB.CommandButton Command17 
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
      Left            =   5985
      TabIndex        =   19
      Top             =   5415
      Width           =   175
   End
   Begin VB.TextBox Text17 
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
      IMEMode         =   3  'DISABLE
      Left            =   4980
      TabIndex        =   18
      Top             =   5415
      Width           =   1005
   End
   Begin VB.TextBox Text3 
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
      Left            =   1335
      MaxLength       =   13
      TabIndex        =   7
      Top             =   4905
      Width           =   1215
   End
   Begin VB.TextBox Text1 
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
      Left            =   1335
      TabIndex        =   5
      Top             =   4050
      Width           =   1215
   End
   Begin VB.TextBox Text10 
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
      Left            =   1350
      MaxLength       =   3
      TabIndex        =   4
      Top             =   3540
      Width           =   510
   End
   Begin VB.TextBox Text6 
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
      Left            =   1350
      TabIndex        =   3
      Top             =   3135
      Width           =   1230
   End
   Begin VB.TextBox Text14 
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
      Left            =   1350
      MaxLength       =   20
      TabIndex        =   2
      Top             =   2745
      Width           =   1875
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Fechas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   3780
      TabIndex        =   37
      ToolTipText     =   "Doble-Click para Fechas Automáticas"
      Top             =   2415
      Width           =   2760
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
         Left            =   1350
         TabIndex        =   13
         Top             =   1100
         Width           =   1020
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
         Height          =   285
         Left            =   1350
         TabIndex        =   11
         Top             =   700
         Width           =   1020
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
         Left            =   1365
         TabIndex        =   9
         Top             =   300
         Width           =   1020
      End
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2370
         TabIndex        =   14
         Top             =   1100
         Width           =   175
      End
      Begin VB.CommandButton Command6 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2370
         TabIndex        =   12
         Top             =   700
         Width           =   175
      End
      Begin VB.CommandButton Command7 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2370
         TabIndex        =   10
         Top             =   300
         Width           =   175
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Vencimiento:"
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
         Left            =   90
         TabIndex        =   40
         Top             =   1170
         Width           =   1125
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Re-Test:"
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
         Left            =   90
         TabIndex        =   39
         Top             =   765
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "Fabricación:"
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
         Left            =   90
         TabIndex        =   38
         Top             =   375
         Width           =   1095
      End
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
      Left            =   1350
      TabIndex        =   0
      Top             =   1700
      Width           =   2325
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
      Height          =   300
      Left            =   3675
      TabIndex        =   1
      Top             =   1700
      Width           =   175
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   5340
      Left            =   6600
      ScaleHeight     =   5280
      ScaleWidth      =   915
      TabIndex        =   22
      Top             =   0
      Width           =   975
      Begin VB.CommandButton Command20 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   105
         Picture         =   "RECAOT092.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Ficha Kardex"
         Top             =   1800
         Width           =   645
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   105
         Picture         =   "RECAOT092.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Alta Nuevo Item de Stock"
         Top             =   870
         Width           =   645
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
         Height          =   675
         Left            =   105
         Picture         =   "RECAOT092.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Salir"
         Top             =   135
         Width           =   645
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   100
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   24
         Top             =   4455
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   255
            TabIndex        =   25
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "RECAOT092.frx":0A68
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4725
         Width           =   1515
      End
   End
   Begin VB.Label label38 
      Caption         =   "Label38"
      Height          =   255
      Left            =   120
      TabIndex        =   68
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label37 
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
      Left            =   2580
      TabIndex        =   65
      Top             =   4950
      Width           =   420
   End
   Begin VB.Label Label36 
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
      Left            =   2595
      TabIndex        =   64
      Top             =   4515
      Width           =   420
   End
   Begin VB.Label Label32 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "P.Prom(gr):"
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
      Left            =   2400
      TabIndex        =   63
      Top             =   5220
      Width           =   1155
   End
   Begin VB.Label Label30 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Precinto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3960
      TabIndex        =   61
      Top             =   4965
      Width           =   1110
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Kg Bruto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   105
      TabIndex        =   60
      Top             =   4545
      Width           =   1155
   End
   Begin VB.Label Label20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Potencia:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   90
      TabIndex        =   58
      Top             =   5490
      Width           =   1155
   End
   Begin VB.Label Label19 
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
      Left            =   3600
      TabIndex        =   57
      ToolTipText     =   "Cantidad nominal"
      Top             =   210
      Width           =   630
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Orden:"
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
      Left            =   2700
      TabIndex        =   56
      Top             =   280
      Width           =   780
   End
   Begin VB.Label Label23 
      Alignment       =   1  'Right Justify
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
      Left            =   1350
      TabIndex        =   55
      Top             =   5760
      Width           =   5010
   End
   Begin VB.Label Label29 
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
      Left            =   3030
      TabIndex        =   54
      ToolTipText     =   "Cantidad nominal"
      Top             =   4470
      Width           =   3300
   End
   Begin VB.Label Label15 
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
      Left            =   2580
      TabIndex        =   53
      Top             =   4095
      Width           =   420
   End
   Begin VB.Label Label13 
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
      Left            =   2595
      TabIndex        =   52
      Top             =   3180
      Width           =   540
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Factura:"
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
      Left            =   3570
      TabIndex        =   51
      Top             =   1060
      Width           =   780
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
      Left            =   4410
      TabIndex        =   50
      Top             =   990
      Width           =   2070
   End
   Begin VB.Label Label7 
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
      Left            =   1350
      TabIndex        =   49
      Top             =   990
      Width           =   2040
   End
   Begin VB.Label Label27 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   4065
      TabIndex        =   48
      Top             =   4120
      Width           =   735
   End
   Begin VB.Label Label18 
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
      Height          =   300
      Left            =   90
      TabIndex        =   47
      Top             =   2170
      Width           =   1155
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Posicion:"
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
      Left            =   3825
      TabIndex        =   46
      Top             =   5505
      Width           =   1110
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tara / Bulto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   90
      TabIndex        =   45
      Top             =   4980
      Width           =   1155
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cant./Bulto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   90
      TabIndex        =   44
      ToolTipText     =   "Doble Click para Ingresar Valores Individuales"
      Top             =   4125
      Width           =   1155
   End
   Begin VB.Label Label35 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Height          =   180
      Left            =   90
      TabIndex        =   43
      Top             =   3615
      Width           =   1155
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cant.Recep:"
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
      Left            =   90
      TabIndex        =   42
      Top             =   3210
      Width           =   1155
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Lote de Proveedor:"
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
      Left            =   90
      TabIndex        =   41
      Top             =   2640
      Width           =   1155
   End
   Begin VB.Label Label8 
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
      Left            =   5880
      TabIndex        =   36
      ToolTipText     =   "Cantidad nominal"
      Top             =   1700
      Width           =   630
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad:"
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
      Left            =   4665
      TabIndex        =   35
      Top             =   1770
      Width           =   1125
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   15
      X2              =   6630
      Y1              =   1485
      Y2              =   1470
   End
   Begin VB.Label Label6 
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
      Left            =   1350
      TabIndex        =   34
      ToolTipText     =   "Descripción"
      Top             =   2100
      Width           =   5160
   End
   Begin VB.Label Label26 
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
      Height          =   300
      Left            =   90
      TabIndex        =   33
      Top             =   1770
      Width           =   1155
   End
   Begin VB.Label Label34 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Numero:"
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
      Left            =   90
      TabIndex        =   32
      Top             =   280
      Width           =   1150
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
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
      Left            =   1350
      TabIndex        =   31
      Top             =   210
      Width           =   1215
   End
   Begin VB.Label Label24 
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
      Left            =   1350
      TabIndex        =   30
      Top             =   600
      Width           =   5130
   End
   Begin VB.Label Label31 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Height          =   195
      Left            =   90
      TabIndex        =   29
      Top             =   1060
      Width           =   1150
   End
   Begin VB.Label Label28 
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
      Left            =   5430
      TabIndex        =   28
      Top             =   210
      Width           =   1035
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
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
      Height          =   225
      Left            =   4650
      TabIndex        =   27
      Top             =   280
      Width           =   645
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor:"
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
      Left            =   90
      TabIndex        =   26
      Top             =   670
      Width           =   1150
   End
End
Attribute VB_Name = "RECAOT092"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MODIFICANDO As Boolean
Private ImprimirBoleta As Boolean
Private NUMBR$
Private IDREQUE$
Dim CASIACTIVO%

Private Sub Command1_Click()
    Command1.Enabled = False
    '
    Screen.MousePointer = 11
    If ImprimirBoleta Then
       'Call PRIBOLRE1
       Call PRINTETIGEN
       ImprimirBoleta = False
    End If
    '
    NORFA$ = TRIM$(RECAOT091.Text9)
    NUOCO$ = TRIM$(RECAOT091.Text2)
    If NORFA$ <> "" Then
         Call REEMPLAZACOMPRAS(NORFA$, 1)
       ElseIf NUOCO$ <> "" Then
         Dim REQUESEL As ADODB.Recordset
         SQLX$ = "SELECT * FROM BALAREQMO WITH(NOLOCK)"
         SQLX$ = SQLX$ + " WHERE IDENLOTE = '" & NUOCO$
         SQLX$ = SQLX$ + " ' ORDER BY IDENREQUE ASC"
         Set REQUESEL = FLEXCONN.Execute(FILTSQL$(SQLX$))
         Do While Not REQUESEL.EOF
            NORFA$ = REQUESEL!Destino
            Call REEMPLAZACOMPRAS(NORFA$, 1)
         REQUESEL.MoveNext
         Loop
         REQUESEL.Close
         Set REQUESEL = Nothing
    End If
    Screen.MousePointer = 1
    '
'    If CODCARPE$ = "" Then
'        RECAOT091.LIMPTEXT
'        Unload Me
'    Else
'        Unload Me
'    End If
'    If CODCARPE$ <> "" Then    COMENTADO POR QUE SI NO NO SE CERRABA
     Unload Me
'    End If
    '
End Sub

Sub PRIBOLRE1()
     
               nroprov% = -(XVALO(RECAOT091.Text1))
               TIPODOC$ = "Boleta de Recepción"
               DESTIDOC% = (-1) * Abs(nroprov%)
               NUMEDOC$ = Label33
        
               ' Impresion boleta de recepcion
               CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = ECOARCH("TADEPOSI", Chr$(1)) 'deposito
               CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = TRIM$(Label7) 'remito
               CODPARAM$(3) = "N-AUTENT": DOCPARAM$(3) = NORVENTA
               CODPARAM$(4) = "NRO-PASE": DOCPARAM$(4) = NUCONVI$
               CODPARAM$(5) = "DOCUORIG": DOCPARAM$(5) = Label9 ' FACTURA
               CODPARAM$(6) = "FECH-EMI": DOCPARAM$(6) = Label28 ' FECHA EMISION
               '
               CAPARDOC% = 6
               '
               CODTABU1$(1) = "COD-MAT" ' codigo
               CODTABU1$(2) = "DES-MAT" ' descripcion
               CODTABU1$(3) = "UNIMED" ' unidad
               CODTABU1$(4) = "NRO-LOTE" ' lote
               CODTABU1$(5) = "CANTIDAD" 'CANTIDAD
               CODTABU1$(6) = "CAKILOS"  'UBICACION
               '
               CACOLTAB1% = 6
               '
               CONDI$ = "SELECT * FROM LOTINGRE WHERE IDENLOTE LIKE '%BR." & Label33 & "%'"
               Set RS = FLEXCONN.Execute(FILTSQL$(CONDI$))
               On Error Resume Next
               RS.MoveFirst
               On Error GoTo 0
               If Err = 0 Then
                 CAITAB1% = 0
                 Do While Not RS.EOF
                     CAITAB1% = CAITAB1% + 1
                     CI% = XVALO(RS!Cod_Inte)
                     TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
                     ORDITEM% = ORDITEM% + 1
                     TETABU1$(2, CAITAB1%) = TRIM$(DESCRIT0$(CI%))  'DESCRIPCION
                     TETABU1$(3, CAITAB1%) = TRIM$(UNIMED$(CI%))   'UNIDAD
                     TETABU1$(4, CAITAB1%) = TRIM$(RS!LOTECOL_PROV) 'LOTE PROV
                     TETABU1$(5, CAITAB1%) = TRIM$(FORMATNUM(XVALO(RS!CANSALDO), "F13.1")) '"F12.1"))  'CANTIDAD REC.
                     TETABU1$(6, CAITAB1%) = TRIM$(SAFETEXT$(RS!UBICFIS)) 'UBICACION
                     RS.MoveNext
                 Loop
                 Call PRINTDOC("BOLERECE")
               End If
            
End Sub
Private Sub Command10_Click()
   '
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Label28 = VDEV$
   End If
   '
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    '
    Call SELECHO("MASTER")
    aux$ = VALACT1("MASTER")
    If aux$ <> "" Then
       LIMPTEXT
       Text8 = aux$
       Call Text8_KeyPress(13)
    End If
    '
    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   
   If TRIM$(Text1) = "" Then
      Exit Sub
   End If
   CI% = Str(CODINT(Text8))
   CantRem% = CantRegistros("MASTER", "codint = '" & Str(CI%) & "'")
   If CantRem% = 0 Then
      COMUNI ("Codigo Inexistente")
      Text8.SetFocus
      Exit Sub
   End If
   
   If XVALO(Text10) <= 0 Then
      Text10 = 1
   End If
   
  If Validacion = False Then
     Exit Sub
  End If
  If MODIFICANDO = True Then
     MODIFICANDO = False
  End If
  List1.AddItem LINEA$
  LIMPTEXT
      
End Sub

Private Sub Command13_Click()
   If List1 = "" Then
      Exit Sub
   End If
   
   If MODIFICANDO = True Then
       COMUNI ("Se encuentra Editando " & Text8 & ", Guarde los Cambios Antes de Continuar")
       Exit Sub
   End If
   
   If List1.Selected(List1.ListIndex) = False Then
      Exit Sub
   End If
      
   List1.RemoveItem (List1.ListIndex)
  ' TIMPTEXT
   
End Sub


Private Sub COMMAND15_CLICK()
   If List1 = "" Then
      Exit Sub
   End If
   If List1.Selected(List1.ListIndex) = False Then
      COMUNI ("Seleccionar Posicion en la Lista, para Insertar")
      Exit Sub
   End If
   If Validacion = False Then
      Exit Sub
   End If
   
   If MODIFICANDO = True Then
      MODIFICANDO = False
   End If
   
   Dim vector() As String
   ReDim vector(List1.ListCount + 1)
   
   pos% = List1.ListIndex
   vi% = 0
   'Recorre el list1, se almacena en un Vector de String
   For AX% = 0 To List1.ListCount - 1
       If AX% = pos% Then
         vector(pos%) = LINEA
         vi% = vi% + 1
       End If
       vector(vi%) = List1.List(AX%)
       vi% = vi% + 1
   Next AX%
   
   List1.AddItem ""
   For AX% = 0 To List1.ListCount - 1
      List1.List(AX%) = vector(AX%)
   Next AX%
   LIMPTEXT
   
End Sub

Private Sub Command16_Click()
    Call FINAL("?RECEPLIST")
End Sub

Private Sub Command17_Click()
    Command17.Enabled = False
    '
    DEPP% = XVALO(Label38)
    Call GENTAUBI(DEPP%)
    If ULTREG&("!TAUBICA") < 1 Then GoTo 99
    Call SELECHO("!TAUBICA")
    
    UBISELE$ = VALACT1$("!TAUBICA")
    If UBISELE$ <> "" Then
       Text17 = UBISELE$
    End If
    '
99  Command17.Enabled = True
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   '
   Call SELECHO("ECORDEF")
   VDEV$ = VALACT1$("ECORDEF")
   
   If VDEV$ <> "" Then
      Text9 = VDEV$
   End If
   '
   Command18.Enabled = True
End Sub

Private Sub Command2_Click()
     Command2.Enabled = False
     Screen.MousePointer = 11
     '
     COD$ = TRIM$(Text8)                    'Codigo ext
     CI% = CODINT(COD$)                     'Cod Interno
     DEPP% = XVALO(Label38)
     '
     If CI% = 0 Then
        COMUNI ("Codigo de Articulo Invalido")
        GoTo 99
     End If
     'Valida si esta dentro de la lista de requerimientos
     If Text9 <> "" Then
        If ECOARCH$("ECORDEF", TRIM$(Text9)) = "" Then
           COMUNI ("Destino Inexistente")
           Text9.SetFocus
           GoTo 99
        End If
     End If
     '
     If TRIM(Text9) <> "" Then
        CONDI$ = "COD_INTE = " & CI% & " AND DESTINO ='" & TRIM$(Text9) & "'"
        CAN% = CantRegistros("BALAREQMO", CONDI$)
        If CAN% < 1 Then
           Call COMUNI("Codigo No es Requerido en " & Text9)
           GoTo 99
        End If
     End If
     '
     'CONDI$ = "codint =" & CI% & "and Trazable = 1"
     'Traza% = XVALO(VALOBADA("Master", "Trazable", CONDI$))
     If ESTRAZABLE%(CI%) > 0 Then
        If Text14 = "" Then
           Call COMUNI("Falta Ingresar Lote")
           Text14.SetFocus
           GoTo 99
        End If
     End If
     '
     'Verifica si esta Bloqueado Sobre-embarque
     NPOC$ = TRIM(RECAOT091.Text2)
     If CantRegistros("OCOMENCA", "Id_Comprob ='" & NPOC$ & "'") > 0 Then
        If CONTROL$("BOLRECEP", "BLOSOBRE") = "SI" Then
           NUMPED% = XVALO(VALOBADA("OCOMENCA", "Nume_Pedi", "ID_COMPROB = '" & NPOC$ & "' AND STAT_OCOM < 8"))
           NUMOCO% = XVALO(VALOBADA("OCOMENCA", "Nume_Ocom", "ID_COMPROB = '" & NPOC$ & "' AND STAT_OCOM < 8"))
           CONDI$ = "(Nume_Pedi =" & NUMPED% & " OR Nume_Ocom = " & NUMOCO% & ") AND COD_INTE = " & CI% & " AND (CANT_OCOM - CANT_REC) > 0.005 AND STAT_OCOM < 8"
           CANOCO# = XVALO(VALOBADA("OCOMDETA", "SUM(CANT_OCOM)", CONDI$))
           CANREC# = XVALO(VALOBADA("LOTINGRE", "SUM(CANALTA)", "DOCUORIG = '" & CODOCO$ & "' AND COD_INTE = " & CI%))
           RESUL# = CANOCO# - CANREC#      'Cantidad Pendiente de Recepcion
           PORSOBRE# = XVALO(CONTROL("BOLRECEP", "PORSOBRE")) 'Porcentaje Permitido
           PENDMASPORC# = RESUL# + (RESUL# * PORSOBRE# / 100) 'Cant Pend. + el Porcentaje Permitido
           If XVALO(Text6.TEXT) > PENDMASPORC# Then
              COMUNI ("Cantidad Recibida Es Mayor a la Cantidad Pendiente\Para Comprobante " & NPOC$)
              Text6.SetFocus
              GoTo 99
           End If
        End If
     End If
     '
     If InStr(Text1, ";") < 1 Then
        If Not (XVALO(Text6) <= XVALO(Text10) * XVALO(Text1)) Then
           COMUNI ("Cantidad Recibida Menor a los Bultos Ingresados")
           Text6.SetFocus
           GoTo 99
        End If
        '
        If Not (XVALO(Text6) > (XVALO(Text10) - 1) * XVALO(Text1)) Then
            COMUNI ("Cantidad de Bultos Excede en Mas de un Bulto la Cantidad Recibida")
            Text6.SetFocus
            GoTo 99
        End If
     End If
     '
     DEPP% = XVALO(Label38)
     CONDI$ = "DEPOUBIC = " & DEPP%
     If CantRegistros("UBIPRE", CONDI$) > 0 Then
        If Label23 = "" Then
           COMUNI ("Falta Ingresar Posición")
           Text17.SetFocus
           GoTo 99
        End If
     End If
     '
     POSI$ = TRIM$(Text17)                  'Ubicacion Propuesta
     Destino$ = TRIM$(Text9)                'Destino
     '
     If PERMITEUBI%(CI%, Destino$, POSI$) <> 1 Then
        Text17.SetFocus
        GoTo 99
     End If
     '
     If XVALO(Text3) = 0 Then
        Text3 = "0.000"
     End If
     '
     If COMALTER%("Revise los Datos Ingresados.\¿ Confirma la Recepción ?\\Si, Confirmar\No, Cancelar") <> 1 Then GoTo 99
        Dim Ok As Boolean
        Ok = False
        COMP$ = CONTROL("", "CODMBREC")        'Tipo de Ingreso BR o IR
        If TRIM$(COMP$) = "" Then COMP$ = "BR"
        FEC% = FECHANUM(TRIM$(Label28))        'Fecha
        DOPRILAR$ = COMP$ & "." & Label33           'Numero
        DoPriCor$ = COMP$ & "." & Label33          'Numero
        DOSELAR$ = "RT." + Label7
        DOSECOR$ = "RT." + Label7 ': While Len(DOSECOR$) > 12: DOSECOR$ = Mid$(DOSECOR$, 2): Wend
        NROCLIE% = XVALO(RECAOT091.Text5)   ' Nro CLIENTE - LABORATORIO
        If NROCLIE% < 1 Then
           NROCLIE% = (-1) * (XVALO(RECAOT091.Text1))   'Nro Proveedor
        End If
        FEFAB% = FECHANUM(TRIM$(Text2))
        FEVEN% = FECHANUM(TRIM$(Text7))
        LOTEPROV$ = TRIM$(Text14)
        Ubic$ = TRIM$(Text17)    'Ubica
        REMITO$ = TRIM$(Label7)
        DESTIN$ = TRIM$(Text9)
     
        CRec# = XVALO(Text6)     'Cantidad total recibida
        '
        'Calculo cant recibida
        DEPORBU$ = TRIM$(Text1)
        If Right$(DEPORBU$, 1) = ";" Then DEPORBU$ = Left$(DEPORBU$, Len(DEPORBU$) - 1)
        RESTO# = CRec#
        RESTO# = CRec#
        INICIO% = XVALO(VALOBADA("MOVISTO", "max(Nuoritem)", "DoPriCor ='" & DoPriCor$ & "'")) + 1
        Do While RESTO# > 0
          '
          PPXX% = InStr(DEPORBU$, ";")
          If PPXX% > 0 Then
               CXBUL# = XVALO(Left$(DEPORBU$, PPXX% - 1))
               DEPORBU$ = Mid$(DEPORBU$, PPXX% + 1)
             Else
               CXBUL# = XVALO(DEPORBU$)    'cantidad x bulto en Text
          End If
          If CXBUL# > RESTO# Then
             CXBUL# = RESTO#
          End If
          '
          If CXBUL# < FRACMIN(CI%) Then GoTo 39 ' PARA QUE SUPRIMA LAS RECEPCIONES MENORES QUE FRACMIN
          '                                     ' RECLAMO DE A.PORTEIRO SEGUN OT-10-0736
          NORDEN% = XVALO(VALOBADA("MOVISTO", "max(Nuoritem)", "DoPriCor ='" & DoPriCor$ & "'"))
          NORDEN% = NORDEN% + 1
          IBOLX$ = TRIM$(Str$(NORDEN%)): While Len(IBOLX$) < 2: IBOLX$ = "0" + IBOLX$: Wend
          LOTE$ = DoPriCor$ & "-" & IBOLX$      ' Nro Lote Propio
          LOTEPROV$ = TRIM$(Text14)             ' Lote
          NUMBR$ = DoPriCor$
             '
             DOCORI$ = TRIM$(RECAOT091.Text2)
          Call MOVISTOK(FEC%, CI%, LOTE$, COMP$, DOPRILAR$, DoPriCor$, DOSELAR$, DOSECOR$, "Recep." + Label24, NORDEN%, 0, 1, NROCLIE%, DEPP%, CXBUL#, , , , DOCORI$)
          Call CREALOTIN(FEC%, FEFAB%, FEVEN%, CI%, LOTE$, "", LOTEPROV$, CXBUL#, 0, 0, DEPP%, Ubic$, RASOCLI$(nroprov%), REMITO$, "", "CUAR", 0, DESTIN$)
          Call ACTUREGISTRO("UBIPRE", "Status", "codiUbic = '" & Ubic$ & "'", 1, nroREG&)
             '
             CONDI$ = "IdenLote ='" & LOTE$ & "'"
          Call ACTUREGISTRO("LOTINGRE", "tara", CONDI$, Text3, REG&)
             FF% = FECHANUM(TRIM$(Text5))
          Call ACTUREGISTRO("LOTINGRE", "Fechco_rep", CONDI$, CVDAT(FF%), REG&)
          Call ACTUREGISTRO("LOTINGRE", "Potencia", CONDI$, Text4, REG&)
          Call ACTUREGISTRO("LOTINGRE", "Precinto", CONDI$, Text12, REG&)
          Call ACTUREGISTRO("LOTINGRE", "DocuOrig", CONDI$, DOCORI$, REG&)
          '
          'Grabar Carpeta de Importacion.
          CODICARP$ = RECAOT091.Label7
          CONDI2$ = "COD_INTE = " & CI% & " AND IDENLOTE='" & LOTE$ & "'"
          Call ACTUREGISTRO("MOVISTO", "CODCARPIMP", CONDI2$, CODICARP$, REG&, "NOMESS")
          Call ACTUREGISTRO("LOTINGRE", "CODCARPIMP", CONDI2$, CODICARP$, REG&, "NOMESS")
          '
          'Actualizar cantidad recibida en Detalle de Pedido
          If TRIM(Text9) <> "" Then
            NORFA$ = TRIM(Text9)
            CONDI$ = "ANEX_OCOM ='" & NORFA$ & "'"
            'PREGUNTO SI EXISTEN PEDIDOS PARA ESA ORDEN DE FABRICACION
            If CantRegistros("OCOMENCA", CONDI$) > 0 Then
               NUMSUM$ = SAFETEXT(VALOBADA("OCOMENCA", "Id_Comprob", "ANEX_OCOM='" & NORFA$ & "'"))
               Call ACSALOC(NUMSUM$)  ' ESTA RUTINA ACTULIZA LOS SALDOS DE LS ORDENES DE COMPRAS Y NOTAS DE PEDIDO
               Call Pedido_Pendiente(NUMSUM$)
            End If
          End If
          '
39        RESTO# = RESTO# - CXBUL#
          Ok = True
          '
       Loop
       '
       DOCUO$ = TRIM$(RECAOT091.Text2)
       If DOCUO$ <> "" Then
           Call Pedido_Pendiente(DOCUO$)
           Call FRESHECHO("!ARPENRE")
       End If
       '
       If Ok Then
         'Call PRIBOLRE
         'Call PRIETIQ
         For AX% = INICIO% To NORDEN%
           'Se modifica para poder obtener el total de los bultos al imprimir etiquetas
           IBOLX$ = TRIM$(Str$(AX%)): While Len(IBOLX$) < 2: IBOLX$ = "0" + IBOLX$: Wend
           LOTE$ = DoPriCor$ & "-" & IBOLX$
           NOCONFIRM% = 1
           Call PRINSTIKAPRO1(CI%, LOTE$, DEPP%)
           NOCONFIRM% = 0
         Next AX%
         ImprimirBoleta = True
       End If
    
       Text1 = ""
       Call SETULTREG("!DETPORBU", 0)
       LIMPTEXT2
       Text6.SetFocus
       
99 Screen.MousePointer = 1
   Command2.Enabled = True
   
End Sub

Private Sub Command20_Click()
    Command20.Enabled = False
    FIKARDEX09.Text1 = Text8
    FIKARDEX09.Show 1
    Command20.Enabled = True
End Sub

Private Sub Command3_Click()
  Command3.Enabled = False
  '
  If DERACCE%("AMASTO07", "Maestro Item de Stock") > 1 Then
   '
   If EXISTE%("AMASTO07.EXE") > 0 Then
      Call CONECRUN("AMASTO07", "Maestro de Items de Stock")
     Else
      AMAS_GES07.Show 1
   End If
   '
  End If
  '
  Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   '
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text7 = VDEV$
   End If
   '
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   '
   CIXI% = CODINT(TRIM$(Text8))
   LOTEPROV$ = TRIM$(Text14)
   LOTE$ = VALOBADA("LOTINGRE", "IDENLOTE", "COD_INTE =" & CIXI% & " AND LOTECOL_PROV ='" & LOTEPROV$ & "'")
   DEPP% = XVALO(Label38)
   If XVALO(CIXI%) > 0 Then
      Call UBIRECOM(CIXI%, LOTE$, DEPP%)
      Call CommSel("!TAUBICA", Command5, Text17)
   End If
   '
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   '
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text5 = VDEV$
   End If
   '
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   '
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Text2 = VDEV$
   End If
   '
   Command7.Enabled = True
End Sub
'
Private Sub Command8_Click()
   '
   Call SELECHO("!ARPENRE")
   COD$ = VALACT1("!ARPENRE")
   IDREQUE$ = Mid(VALACT2("!ARPENRE"), 34, 13)
   '
   If COD$ <> "" Then
      CI% = CODINT(COD$)
      NORFA$ = TRIM(Text9)
      CODOCO$ = TRIM$(RECAOT091.Text2)
      NUMPED% = XVALO(VALOBADA("OCOMENCA", "Nume_Pedi", "ID_COMPROB = '" & CODOCO$ & "' AND STAT_OCOM < 8"))
      NUMOCO% = XVALO(VALOBADA("OCOMENCA", "Nume_Ocom", "ID_COMPROB = '" & CODOCO$ & "' AND STAT_OCOM < 8"))
      CONDI$ = "(Nume_Pedi =" & NUMPED% & " OR Nume_Ocom = " & NUMOCO% & ") AND COD_INTE = " & CI% & " AND (CANT_OCOM - CANT_REC) > 0.005 AND STAT_OCOM < 8"
      CANOCO# = XVALO(VALOBADA("OCOMDETA", "SUM(CANT_OCOM)", CONDI$))
      CANREC# = XVALO(VALOBADA("LOTINGRE", "SUM(CANALTA)", "DOCUORIG = '" & CODOCO$ & "' AND COD_INTE = " & CI%))
      RESUL# = CANOCO# - CANREC#
      Text8 = COD$
      If RESUL# > 0 Then
          CIXI% = CODINT(Text8)
          Text6 = FORMATNUM(RESUL#, "F13." & CNTDC$(CIXI%))
          Text1 = TRIM$(Text6)
         Else
          Text6 = ""
       End If
     End If

End Sub
'
Private Sub Form_Activate()
    'Text8.SetFocus
End Sub
'
Private Sub Form_Load()
    MODIFICANDO = False
    ImprimirBoleta = False
    
    'Copia Informacion de RECAOT091 a RECAOT092
    Label33 = RECAOT091.Label33
    Label28 = RECAOT091.Label28
    Label24 = RECAOT091.Label24
      If RECAOT091.Label24 = "" Then Label24 = RECAOT091.Label3
    Label7 = RECAOT091.Text4
    Label9 = RECAOT091.Text3
    Text9 = RECAOT091.Text9
    Label29 = RECAOT091.Label29
    Label38 = XVALO(RECAOT091.Text6)   ' Asigna el deposito de entrada
    '
    'Si viene con Destino, No se puede modificar
    If TRIM(Text9) <> "" Then
       Text9.Enabled = False
       Command18.Enabled = False
      Else
       Text9.Enabled = True
       Command18.Enabled = True
    End If
    '
    Command11.Enabled = True
    If TRIM$(RECAOT091.Text2) <> "" Then
       Command11.Enabled = False
    End If
    '
    Label13 = "KG"
    Label36 = "KG"
    Label37 = "KG"
    Label15 = "KG"
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    'NUBORE$ = TRIM$(Str$(ProNroComprobante("BR")))  'Proximo Nro Comprobante
    'While Len(NUBORE$) < 6: NUBORE$ = "0" + NUBORE$: Wend
    'NUBORE$ = "BR." + NUBORE$
    'Label33 = NUBORE$
    '
    '
    Command8.Visible = False
    CONDI$ = ""
    NORFA$ = TRIM(Text9)
    IDSUM$ = TRIM(RECAOT091.Text2.TEXT)
    
    If NORFA$ <> "" Then
       ID_COMP$ = SAFETEXT(VALOBADA("OCOMENCA", "ID_COMPROB", "ANEX_OCOM='" & NORFA$ & "'"))
    End If
    '
    If IDSUM$ <> "" Then
       ID_COMP$ = IDSUM$
    End If
    '
    If ID_COMP$ <> "" Then
       Call Pedido_Pendiente(ID_COMP$)
       Command8.Visible = True
    End If
    
    'If NORFA$ <> ""  Then
    '   CONDI$ = "ANEX_OCOM ='" & NORFA$ & "'"
       'PREGUNTO SI EXISTEN PEDIDOS PARA ESA ORDEN DE FABRICACION
     
       
    
99
    '
    If Left$(APLINVO$, 8) = "FRABOBIN" Then
       FRABOBI07.Show 1
       Call FINAL("")
    End If
    '
    If Left$(APLINVO$, 8) = "PREPATIN" Then
       PREPATI07.Show 1
       Call FINAL("")
    End If
    '
    If Left$(APLINVO$, 8) = "TRAZORFA" Or Left$(APLINVO$, 7) = "OFABRIC" Then
       TRAZORF07.Show 1
       Call FINAL("")
    End If
    '
    If Left$(APLINVO$, 8) = "ORFCIE" Then   ' Or INTERPRE = 1 Then
       CIERRAOF07.Show 1
       Call FINAL("")
    End If
    '
    If Left$(APLINVO$, 7) = "BORECEP" Then ' Or INTERPRE% = 1 Then
       Text1 = Mid$(APLINVO$, 8)
       openforms = DoEvents
    End If
    '
    If Left$(APLINVO$, 7) = "GENUMSE" Then
       GENUMSE07.Show 1
       Call FINAL("")
    End If
    '
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call Command1_Click
End Sub

'
Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    On Error Resume Next
    Hide
    On Error GoTo 0
End Sub
'
Private Sub Frame5_DblClick()
   HOII% = HOY(HOS$)
   '
   If ESTRAZABLE%(CODINT%(Text8)) > 0 Then
     If TRIM$(Text2) = "" Then
          Text2 = HOS$
          ANX% = XVALO(Mid$(HOS$, 7)) + 2
          ANXX$ = TRIM$(Str$(ANX%)): While Len(ANXX$) < 2: ANXX$ = "0" + ANXX$: Wend
          VENX$ = Left$(HOS$, 6) + ANXX$
          If TRIM$(Text5) = "" Then Text5 = VENX$
          If TRIM$(Text7) = "" Then Text7 = VENX$
       Else
          Call MENSERR(24, "Imposible Aplicar Vencimientos Automáticos")
     End If
     '
     If XVALO(Label13) < 1 Then
         'Command12.SetFocus
        Else
        Command15.SetFocus
     End If
   End If
   '
End Sub



Private Sub Label14_DblClick()
     Call CARGAPORDETALLE
End Sub

Private Sub Label8_Change()
     'Tratamiento de unidad de medida para todos los labels
     If UCase(TRIM(Label8)) = "KG" Or UCase(TRIM(Label8)) = "GR" Then
        Label13 = Label8   ' Cant. Recibida
        Label15 = Label8   ' Cant / Bulto
        Label36 = "KG"   ' Total Bruto
        Label37 = Label8   ' tara / Bulto
       Else
        Label13 = Label8   ' Cant. Recibida
        Label15 = Label8   ' Cant / Bulto
        Label36 = "KG"   ' Total Bruto
        Label37 = "KG"   ' tara / Bulto
     End If
    
End Sub

Private Sub Text1_Change()
   Text1.ToolTipText = Text1.TEXT
End Sub
'
Private Sub Text1_GotFocus()
   Text1.SelStart = 0
   Text1.SelLength = 13
End Sub
'
Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text11.SetFocus
    End If
End Sub
'
Private Sub Text1_LostFocus()
   If InStr(Text1, ";") < 1 Then
      CIXI% = CODINT(Text8)
      Text1 = FORMATNUM(Text1, "F13." & CNTDC$(CIXI%))
      If TRIM(Text1) = "*************" Then
         COMUNI ("Cantidad por Bulto Supera Maximo Admitido")
         Text1.SetFocus
         Exit Sub
      End If
   End If
End Sub

Private Sub Text10_GotFocus()
     Text10.SelStart = 0
     Text10.SelLength = Text10.MaxLength
End Sub
'
Private Sub Text10_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text1.SetFocus
    End If
End Sub

Private Sub Text10_LostFocus()
   If XVALO(Text10) = 1 Then
      Text1 = Text6
   End If
   Text10 = FORMATNUM(Text10, "F3.0")
   
End Sub

Private Sub Text11_GotFocus()
   Text11.SelStart = 0
   Text11.SelLength = 13
   CASIACTIVO% = 2
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text3.SetFocus
    End If
End Sub

Private Sub Text11_LostFocus()
    '
    CIXI% = CODINT(Text8)
    If XVALO(Text11) = 0 Then
       Text11 = FORMATNUM(0, "F13.3")
       Exit Sub
    End If
    '
    Text11 = FORMATNUM(Text11, "F13.3")
    If TRIM(Text11) = "*************" Then
       COMUNI ("Total Bruto Supera Maximo Admitido")
       Text11.SetFocus
       Exit Sub
    End If
    '
    'Text3 = XVALO(Text11) - XVALO(Text6)
    'Text3 = FORMATNUM(Text3, "F11.3")
    '
    '
    If (UCase(Label8) <> "KG" And UCase(Label8) <> "GR") Then Exit Sub
    '
    'Calcula Tara si Cant es > 0
    If XVALO(Text11) < XVALO(Text6) Then
       COMUNI ("Total Bruto no Puede ser Menor a Cantidad Recibida")
       Text11.SetFocus
       Exit Sub
    End If
    If XVALO(Text6) > 0 Then
       If XVALO(Text10) > 0.00005 Then
          AA# = (XVALO(Text11) - XVALO(Text6)) / XVALO(Text10)
          Text3 = FORMATNUM(AA#, "F13.3")
          'Se modifica el calculo, por inconveniente en la asignacion del resultado en Text3
          'Text3 = (XVALO(Text11) - XVALO(Text6)) / XVALO(Text10)
          'Text3 = FORMATNUM(Text3, "F11." & CNTDC$)
       End If
    End If
     
    'Calcula cant Recibida si tara > 0 y cant rec = 0
    If XVALO(Text3) > 0 And XVALO(Text6) = 0 Then
       XX# = XVALO(Text11) - XVALO(Text3)
       If XX# < 0 Then  'Si Calculo de Cant Rec < 0, SALE
          Exit Sub
       End If
       Text6 = XX#
       Text6 = FORMATNUM(Text6, "F13." & CNTDC$(CIXI%))
    End If
     
End Sub

Private Sub Text12_GotFocus()
    Text12.SelStart = 0
    Text12.SelLength = Text12.MaxLength
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text17.SetFocus
    End If
End Sub

Private Sub Text13_LostFocus()
       If XVALO(Text13) > 99.99999 Then
          COMUNI ("Maximo Admitido 99.99999")
          Text13.SetFocus
          Exit Sub
        Else
          Text13 = FORMATNUM(Text13, "F13.5")
       End If
End Sub

'
Private Sub Text14_GotFocus()
    Text14.SelStart = 0
    Text14.SelLength = Text14.MaxLength
 
End Sub
'
Private Sub Text14_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text6.SetFocus
    End If
End Sub
'
Private Sub Text17_Change()
    If Text17 <> "" Then
       CONDI$ = "codiUbic ='" & TRIM$(Text17) & "'"
       DES$ = VALOBADA("UBIPRE", "Referencia", CONDI$)
       'DES$ = ECOARCH$("TAUBICA", Text17)
       Label23 = DES$
    End If
    
End Sub
'
Private Sub Text17_DblClick()
    Call Command17_Click
End Sub
'
Private Sub Text17_GotFocus()
    Text17.SelStart = 0
    Text17.SelLength = Text2.MaxLength
End Sub
'
Private Sub Text17_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
       Call Command17_Click
       Command2.SetFocus
    End If
End Sub
'
Private Sub Text17_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Command2.SetFocus
    End If
End Sub
'

'
Private Sub Text2_GotFocus()
 Text2.SelStart = 0
 Text2.SelLength = Text2.MaxLength
End Sub
'
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
       Call Command7_Click
       Text5.SetFocus
    End If
End Sub
'
Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text5.SetFocus
    End If
End Sub
'
Private Sub Text2_LostFocus()
   FEC = FECHANUM(TRIM(Text2))
   Text2 = FECHATEX(FEC)
   
End Sub
'
Private Sub Text3_GotFocus()
   Text3.SelStart = 0
   Text3.SelLength = Text3.MaxLength
   CASIACTIVO% = 3
End Sub
'
Private Sub Text3_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     If Text4.Enabled = True Then
        Text4.SetFocus
     Else
        Text12.SetFocus
     End If
  End If
End Sub
'
Sub CALCUTARA()
' FALTA COMPLETAR
   CIXI% = CODINT%(Text8)
   If CIXI% > 0 Then
     UME$ = TRIM$(UNIMED$(CIXI%))
     If UME$ = "KG" Then
       KINETO = XVALO(Text6)
       CABUL = XVALO(Text10)
       If CABUL >= 1 Then
          KIBRUTO = XVALO(Text11)
          KITARA = CABUL * XVALO(Text3)
          '
          If CASIACTIVO% = 1 Then
            If XVALO(Text11) < KINETO Then
               Text11 = FORMATNUM(KINETO, "F13." & CNTDC$(CIXI%))
            End If
            KITARA = (XVALO(Text11) - XVALO(Text6)) / CABUL
            Text3 = KITARA
            Call Text3_LostFocus
            Exit Sub
          ElseIf CASIACTIVO% = 2 Then
            If XVALO(Text6) > KIBRUTO - KITARA Then
               Text6 = FORMATNUM(KINETO, "F13." & CNTDC$(CIXI%))
            End If
            KITARA = (XVALO(Text11) - XVALO(Text6)) / CABUL
            Text3 = KITARA
            Call Text3_LostFocus
            Exit Sub
          End If
       End If
     End If
   End If
   '
End Sub
'
Private Sub Text3_LostFocus()
       CIXI% = CODINT(Text8)
       Text3 = FORMATNUM(Text3, "F13.3")
       If TRIM(Text3) = "*************" Then
          COMUNI ("Tara por Bulto Supera Maximo Admitido")
          Text3.SetFocus
          Exit Sub
       End If
       If XVALO(Text3) = 0 Then
          Text3 = FORMATNUM(0, "F13.3")
       End If
     '
     If (UCase(Label8) <> "KG" And UCase(Label8) <> "GR") Then Exit Sub
     '
     'Calcula Bruto si CantRec es > 0
     If XVALO(Text6) > 0 Then
        Text11 = XVALO(Text6) + (XVALO(Text3) * XVALO(Text10))
        Text11 = FORMATNUM(Text11, "F13.3")
        Exit Sub
     End If
      
     'Calcula  Cant Rec, si Bruto > 0 y cant rec = 0
      If XVALO(Text11) > 0 And XVALO(Text6) = 0 Then
         XX# = XVALO(Text11) - XVALO(Text3)
         If XX# < 0 Then
            Exit Sub
         End If
         Text6 = XX#
         Text6 = FORMATNUM(Text3, "F13." & CNTDC$(CIXI%))
      End If
       
End Sub
'
Private Sub Text4_GotFocus()
    Text4.SelStart = 0
    Text4.SelLength = Text4.MaxLength
End Sub
'
Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text2.SetFocus
    End If
End Sub
'
Private Sub Text4_LostFocus()
      '
      If XVALO(Text4) > 999.99 Then
         COMUNI ("Maximo Admitido 999.99")
         Text4.SetFocus
         Exit Sub
        Else
         Text4 = TRIM$(FORMATNUM(Text4, "F13.2"))
      End If
      '
      If XVALO(Text4) = 0 Then
          Text4 = "0.00"
      End If
      '
End Sub
'
Private Sub Text5_GotFocus()
    Text5.SelStart = 0
    Text5.SelLength = Text2.MaxLength
End Sub
'
Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
       Call Command6_Click
       Text7.SetFocus
    End If
End Sub
'
Private Sub Text5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text7.SetFocus
    End If
End Sub
'
Private Sub Text5_LostFocus()
   FEC = FECHANUM(TRIM(Text5))
   Text5 = FECHATEX(FEC)
End Sub
'
Private Sub Text6_GotFocus()
   Text6.SelStart = 0
   Text6.SelLength = 13
   CASIACTIVO% = 1
End Sub
'
Private Sub Text6_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text10.SetFocus
    End If
    If KeyAscii = 27 Then
       LIMPTEXT
       Text8.SetFocus
    End If
End Sub

'
Private Sub Text7_GotFocus()
   Text7.SelStart = 0
   Text7.SelLength = Text2.MaxLength
End Sub
'
Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
       Call Command4_Click
       Text9.SetFocus
    End If
End Sub
'
Private Sub Text7_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Text9.Enabled = True Then
          Text9.SetFocus
         Else
          Text12.SetFocus
       End If
    End If
End Sub
'
Private Sub Text7_LostFocus()
   FEC = FECHANUM(TRIM(Text7))
   Text7 = FECHATEX(FEC)
End Sub
'
Private Sub Text8_DblClick()
   If Command11.Enabled = True Then
      Call Command11_Click
    Else
      Call Command8_Click
   End If
End Sub
'
'
Sub LIMPTEXT()
    Text2 = ""
    Text5 = ""
    Text6 = ""
    Text8 = ""
    Text7 = ""
    Text11 = ""
    Text12 = ""
    Text14 = ""
    Text16 = ""
    Text17 = ""
    Label6 = ""
    Label8 = ""
    Text8 = ""
    
    Label13 = ""
    Text10 = ""
   
    Label23 = ""
    Text3 = ""
    Text4 = ""
    
End Sub
Sub LIMPTEXT2()
    Text6 = ""       'cant recibida
    Text10 = "1"     'bultos
    Text1 = ""       'cant x bulto
    Call SETULTREG("!DETPORBU", 0)
    Text11 = ""      'total bruto
    Text3 = ""       'tara x bulto
    Text12 = ""      'precinto
End Sub

'
Function LINEA$()
     Dim Codigo As String * 16
     Dim Descrip As String * 41
     Dim Unidad As String * 4
     Dim CantPend As String * 9
     Dim CantRec As String * 9
     Dim FentreSol As String * 10
     Dim NumOcom As String * 12
     Dim LOTE As String * 17
     Dim Ubic As String * 33
     Dim Fabri As String * 9
     Dim ReTest As String * 9
     Dim Venc As String * 9
     Dim BULTO As String * 3
     Dim Destino As String * 15
     Codigo = ""
     Descrip = ""
     Unidad = ""
     CantPend = ""
     CantRec = ""
     FentreSol = ""
     NumOcom = ""
     LOTE = ""
     Ubic = ""
     Fabri = ""
     ReTest = ""
     Venc = ""
     BULTO = ""
     '
     Codigo = AJUSTI$(TRIM$(Text8), 16)
     Descrip = AJUSTI$(TRIM$(Label6), 41)
     Unidad = AJUSTI$(TRIM$(Label8), 4)
     CantPend = AJUSTD$("", 9)
     CantRec = AJUSTD$(FORMATNUM(XVALO(Text6), "F9.1"), 9)
     'FentreSol = AJUSTD$(TRIM$(Label32), 10)
     NumOcom = AJUSTI$("", 12)
     LOTE = AJUSTI$(TRIM$(Text14), 17)
     Ubic = AJUSTI$(TRIM$(Text17), 33)
     Fabri = AJUSTI$(TRIM$(Text2), 9)
     ReTest = AJUSTI$(TRIM$(Text5), 9)
     Venc = AJUSTI$(TRIM$(Text7), 9)
     'NORD = AJUSTI$(TRIM$(Label13), 3)
     BULTO = AJUSTI$(TRIM$(Text10), 3)
     Destino = AJUSTI$(TRIM$(Text9), 15)
     '
     LINEA$ = Codigo & Descrip & Unidad & CantPend & CantRec & FentreSol & NumOcom & LOTE & Ubic & Fabri & ReTest & Venc & BULTO & Destino
End Function
'
Function ListFecha$()
     Call SELECHO("SELFECHA")
     ListFecha$ = VALACT1$("SELFECHA")
End Function
'
Private Sub Text2_DblClick()
     Call Command7_Click
End Sub
'
Private Sub Text5_DblClick()
     Call Command6_Click
End Sub
'
Private Sub Text6_LostFocus()
     '
     CIXI% = CODINT(Text8)
     Text6 = FORMATNUM(Text6, "F13." & CNTDC$(CIXI%))
     If TRIM(Text6) = "*************" Then
        COMUNI ("Cantidad Recibida Supera Maximo Admitido")
        Text6.SetFocus
        Exit Sub
     End If
     '
     If (UCase(Label8) <> "KG" And UCase(Label8) <> "GR") Then Exit Sub
     '
     'Calcula Bruto si Tara es > 0
     If XVALO(Text3) > 0 Then
        Text11 = XVALO(Text6) + (XVALO(Text3) * XVALO(Text10))
        Text11 = FORMATNUM(Text11, "F13." & CNTDC$(CIXI%))
     End If
     'Calcula Tara si Bruto es > 0
     If XVALO(Text11) > 0 Then
        If XVALO(Text10) > 0.00005 Then
           Text3 = (XVALO(Text11) - XVALO(Text6)) / XVALO(Text10)
           Text3 = FORMATNUM(Text3, "F13." & CNTDC$(CIXI%))
        End If
     End If
     '
End Sub
'
Private Sub Text7_DblClick()
   Call Command4_Click
End Sub
'
Function Validacion()
    Validacion = False
    If CODINT(Text8) < 0 Then
       Call COMUNI("Codigo Inexistente")
       Validacion = False
       Exit Function
    End If
    '
    If Text14 = "" Then
       Call COMUNI("Ingresar Lote")
       Text14.SetFocus
       Validacion = False
       Exit Function
    End If
    '
    If Text17 = "" Then
       Call COMUNI("Ingresar Ubicación")
       Text17.SetFocus
       Validacion = False
       Exit Function
    End If
    '
    If XVALO(Text6) <= 0 Then
       Call COMUNI("Ingresar Cantidad Recibida")
       Text6.SetFocus
       Validacion = False
       Exit Function
    End If
    '
    If FECHANUM(TRIM$(Text2)) = 0 Then
       COMUNI ("Ingresar Fecha de Fabricación")
       Text2.SetFocus
       Validacion = False
       Exit Function
    End If
    '
    If FECHANUM(TRIM$(Text5)) = 0 Then
       Call COMUNI("Ingresar Fecha de ReTest")
       Text5.SetFocus
       Validacion = False
       Exit Function
    End If
    '
    If FECHANUM(TRIM$(Text7)) = 0 Then
       Call COMUNI("Ingresar Fecha de Vencimiento")
       Text7.SetFocus
       Validacion = False
       Exit Function
    End If
    '
    If (FECHANUM(TRIM$(Text5)) < FECHANUM(Text2)) Then
       COMUNI ("Fecha de Fabricación es Mayor a Fecha de Retest")
       Text2.SetFocus
       Validacion = False
       Exit Function
    End If
    '
    If (FECHANUM(TRIM$(Text7)) < FECHANUM(Text5)) Then
       COMUNI ("Fecha de Re-test es Mayor a Fecha de Vencimiento")
       Text5.SetFocus
       Validacion = False
       Exit Function
    End If
    '
    Validacion = True
    '
End Function

Private Sub Text8_Change()
     
         'LABEL15 = ""
         CI% = CODINT%(Text8)
         aux$ = TRIM$(Text8)
         Text14 = ""
         
       If (CI%) > 0 Then
          CONDI$ = "COD_INTE = " & Str(CI%)
          Text8 = aux$
          Label6 = VALOBADA("MASTER", "Descripcion", "Codint = " & Str(CI%))
          Label8 = VALOBADA("MASTER", "Umedida", "Codint = " & Str(CI%))
          Text10 = 1
          If ESTRAZABLE%(CI%) < 1 Then
             Text2 = "": Text2.Enabled = False: Command7.Enabled = False
             Text5 = "": Text5.Enabled = False: Command6.Enabled = False
             Text7 = "": Text7.Enabled = False: Command4.Enabled = False
             Text4 = "": Text4.Enabled = False:
          Else
             Text2.Enabled = True: Command7.Enabled = True
             Text5.Enabled = True: Command6.Enabled = True
             Text7.Enabled = True: Command4.Enabled = True
             Text4 = 100: Text4.Enabled = True
          End If
       End If
     '
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
       Call Command11_Click
    End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Text14.SetFocus
   End If
End Sub

Private Sub Text8_LostFocus()
    If Text8 <> "" Then
       CI% = CODINT(Text8)
       Text8 = CODEXT(CI%)
    End If
End Sub

Private Sub Text9_Change()
    If Text9 <> "" Then
       DES$ = ECOARCH$("ECORDEF", Text9)
       Label29 = DES$
    End If
End Sub

Private Sub Text9_DblClick()
  Call Command18_Click
End Sub

Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
       Call Command18_Click
       Text17.SetFocus
    End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       Text12.SetFocus
    End If
End Sub

Sub PRIBOLRE()
               nroprov% = -(XVALO(Text1))
               TIPODOC$ = "Boleta de Recepción"
               DESTIDOC% = (-1) * Abs(nroprov%)
               NUMEDOC$ = Label33
        
               ' Impresion boleta de recepcion
               CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = ECOARCH("TADEPOSI", Chr$(1)) 'deposito
               CODPARAM$(2) = "NRO-REMI":  DOCPARAM$(2) = TRIM$(Label7) 'remito
               CODPARAM$(3) = "N-AUTENT": DOCPARAM$(3) = NORVENTA
               CODPARAM$(4) = "NRO-PASE": DOCPARAM$(4) = NUCONVI$
               CODPARAM$(5) = "DOCUORIG": DOCPARAM$(5) = Label9 ' FACTURA
               CODPARAM$(6) = "FECH-EMI": DOCPARAM$(6) = Label28 ' FECHA EMISION
               '
               CAPARDOC% = 6
               '
               CODTABU1$(1) = "COD-MAT" ' codigo
               CODTABU1$(2) = "DES-MAT" ' descripcion
               CODTABU1$(3) = "UNIMED" ' unidad
               CODTABU1$(4) = "NRO-LOTE" ' lote
               CODTABU1$(5) = "CANTIDAD" 'CANTIDAD
               CODTABU1$(6) = "CAKILOS"  'UBICACION
               '
               CACOLTAB1% = 6
               '
               CAITAB1% = 0
               For I = 0 To List1.ListCount - 1
                 CAITAB1% = CAITAB1% + 1
                 TETABU1$(1, CAITAB1%) = Left$(List1.List(I), 15) 'Codigo
                 ORDITEM% = ORDITEM% + 1
                 TETABU1$(2, CAITAB1%) = TRIM$(Mid$(List1.List(I), 16, 41))  'DESCRIPCION
                 TETABU1$(3, CAITAB1%) = TRIM$(Mid$(List1.List(I), 58, 4))   'UNIDAD
                 TETABU1$(4, CAITAB1%) = TRIM$(Mid$(List1.List(I), 102, 17)) 'LOTE PROV
                 TETABU1$(5, CAITAB1%) = TRIM$(Mid$(List1.List(I), 71, 9))   'CANTIDAD REC.
                 TETABU1$(6, CAITAB1%) = TRIM$(Mid$(List1.List(I), 119, 33)) 'UBICACION
               Next I
            
            Call PRINTDOC("BOLERECE")
            
End Sub

Sub PRIETIQ()

20          CON$ = CONTROL$("", "ETIQUETA")
            If CON$ = "" Then
               COMUNI ("DEBE CONFIGURAR EL FORMULARIO PREVIAMENTE...")
               CFGRECEP.Show 1
               GoTo 20
            End If
            
            CACOLTAB1% = 0
            'Impresion de etiqueta
            
            TIPODOC$ = "Etiqueta de Recepcion"
            
            CODPARAM$(1) = "NUME-COM"
            CODPARAM$(7) = "USUARIO"
            CODPARAM$(8) = "FECH-EMI"
            CODPARAM$(2) = "NRO-LOTE"
            CODPARAM$(3) = "DES-MAT"
            CODPARAM$(4) = "COD-MAT"
            CODPARAM$(5) = "FECH-LAN"
            CODPARAM$(6) = "CANTIDAD"
            CODPARAM$(9) = "REF-MAT2"
            CODPARAM$(10) = "UNIMED"
            CODPARAM$(11) = "MEDIDA-1"
            CODPARAM$(12) = "REF-MAT1"
            CAPARDOC% = 12
            
            DOCPARAM$(1) = Label7            'REMITO
            DOCPARAM$(7) = TRIM$(USERNAME)  'USUARIO
            DOCPARAM$(8) = TRIM$(Label28)   'FECHA EMISION
             
            '
            'Recorre todo el List
            For I = 0 To List1.ListCount - 1
              CantBul% = XVALO(Mid(List1.List(I), 178, 4))
               'Bucle por cantidad de bultos
               For AX% = 1 To CantBul%
                
                DOCPARAM$(3) = TRIM$(Mid$(List1.List(I), 16, 41)) ' Descripcion
                DOCPARAM$(4) = Left$(List1.List(I), 15) 'Codigo
                DOCPARAM$(5) = TRIM$(Mid$(List1.List(I), 161, 9))   'Retest
                DOCPARAM$(6) = TRIM$(Mid$(List1.List(I), 71, 9))   ' Cantidad Recibida
                DOCPARAM$(9) = AX% & " De " & CantBul%
                DOCPARAM$(10) = TRIM$(Mid$(List1.List(I), 58, 4)) ' UNIDAD
                DOCPARAM$(11) = TRIM$(Mid$(List1.List(I), 102, 17)) 'LOTE PROV
                IBOLX$ = TRIM$(Str(I + 1)): While Len(IBOLX$) < 2: IBOLX$ = "0" + IBOLX$: Wend
                DOCPARAM$(12) = AJUSTI(Left$(List1.List(I), 15), 15) + AJUSTI(Label33, 9) & "-" & IBOLX$
                DOCPARAM$(2) = TRIM$(Label33) & "-" & IBOLX$  'LOTE PROPIO

                Call PRINTDOC("ETIQUETA")
               Next AX%
             Next I
  
End Sub

Function NroComp()
    'Crea el Proximo numero de comprobante,
    'Completa con 6 Ceros y agrega leyenda BR. al principio
    '
    NUBORE$ = TRIM$(Str$(ProNroComprobante("BR")))  'Proximo Nro Comprobante
    While Len(NUBORE$) < 6: NUBORE$ = "0" + NUBORE$: Wend
    NUBORE$ = "BR." + NUBORE$
    NroComp = NUBORE$
    RECAOT091.Label33 = NUBORE$
    '
End Function

Sub PRINTETIGEN()

  NUMBR$ = NUMBR$ & "-00"
  'Impresion Etiqueta de conjunto
  CODPARAM$(1) = "NRO-LOTE": DOCPARAM$(1) = NUMBR$
  CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = TRIM$(Label7)
  CODPARAM$(3) = "FECH-REC": DOCPARAM$(3) = TRIM$(Label28)
  
  CAPARDOC% = 3
  '
  Call PRINTDOC("@" & "ETIN-BLP")
     
End Sub
'
Sub Pedido_Pendiente(ID_COMP$)
  '
  ID_COMP$ = TRIM$(RECAOT091.Text2)
  If ID_COMP$ = "" Then
     SOFA$ = TRIM$(RECAOT091.Text9)
     ID_COMP$ = SAFETEXT(VALOBADA("OCOMENCA", "ID_COMPROB", "ANEX_OCOM='" & SOFA$ & "'"))
  End If
  '
  If ID_COMP$ <> "" Then
     Call ACSALOC(ID_COMP$)
  End If
  '
  NUMPED% = XVALO(VALOBADA("OCOMENCA", "Nume_Pedi", "ID_COMPROB='" & ID_COMP$ & "'"))
  NUMOCO% = XVALO(VALOBADA("OCOMENCA", "Nume_Ocom", "ID_COMPROB='" & ID_COMP$ & "'"))
  SQLX$ = ""
  If NUMPED% > 0 Then
     SQLX$ = "SELECT Nume_Pedi, COD_INTE, CANT_OCOM, CANT_REC, IDENREQUE FROM OCOMDETA WHERE Nume_Pedi =" & NUMPED% & " AND (CANT_OCOM - CANT_REC) > 0.00005"
    ElseIf NUMOCO% > 0 Then
     SQLX$ = "SELECT Nume_Pedi, COD_INTE, CANT_OCOM, CANT_REC, IDENREQUE FROM OCOMDETA WHERE Nume_Ocom =" & NUMOCO% & " AND (CANT_OCOM - CANT_REC) > 0.00005 AND STAT_OCOM < 3"
     'Se agrego Stat_Ocom < 3 para que en la lista de selccion, muestre solamente el ultimo pedido de la Orden de Compra elegida.
  End If
  '
  Call COPYSTRU("ECOMAST", "ARPENRE")
  Call SETULTREG("!ARPENRE", 0)
  Call BORRALISEL("!ARPENRE")
  LF1% = LENFIELD("ARPENRE", 1)
  PF2% = POSFIELD("ARPENRE", 2)
  Call FRESHECHO("!ARPENRE")
  'SELECCIONO LOS ITEMS QUE COMPONEN EL DETALLE DEL PEDIDO
  Dim Pedi As ADODB.Recordset
  Set Pedi = New ADODB.Recordset
  Set Pedi = FLEXCONN.Execute(FILTSQL$(SQLX$))
  If Pedi.BOF = True And Pedi.EOF = True Then GoTo 99
     With Pedi
        II& = 0
        Do While Not (.EOF)
           YY$ = REGBLAN("ARPENRE")
           CICI% = XVALO(!Cod_Inte)
           DESCRI$ = Left(DESCRIT0(CICI%), 46)
           Call REPLA(YY$, CODEXT(!Cod_Inte), 1, LF1%)
           Call REPLA(YY$, DESCRI$, PF2%, 34)
           SALPE$ = FORMATNUM$(XVALO(!CANT_OCOM) - XVALO(!Cant_Rec), "F12." + CNTDC$)
           If XVALO(SALPE$) >= FRACMIN(CICI%) Then
              II& = II& + 1
              SELPEX$ = FORMATSTAN(SALPE$, "F12", CICI%)
              Call REPLA(YY$, SELPEX$, PF2% + 34, 12)
              Call GRAREG("!ARPENRE", YY$, II&)
           End If
           .MoveNext
        Loop
     End With
     Call SETULTREG("!ARPENRE", II&)
     Call CIERRARCH("!ARPENRE")
     Call FRESHECHO("!ARPENRE")
99    '
End Sub

Sub CARGAPORDETALLE()
    '
    CIXI% = CODINT(Text8)
    CANBU& = XVALO(Text10)
    Call SETULTREG("!DETPORBU", CANBU&)
    '
    REBLA$ = REGBLAN$("DETPORBU")
    COMP$ = CONTROL("", "CODMBREC")        'Tipo de Ingreso BR o IR
    If TRIM$(COMP$) = "" Then COMP$ = "BR"
    DoPriCor$ = COMP$ & "." & Label33          'Numero
    NORDEN% = XVALO(VALOBADA("MOVISTO", "max(Nuoritem)", "DoPriCor ='" & DoPriCor$ & "'"))
    For U& = 1 To CANBU&
       NORDEN% = NORDEN% + 1
       IBOLX$ = TRIM$(Str$(NORDEN%)): While Len(IBOLX$) < 2: IBOLX$ = "0" + IBOLX$: Wend
       LOTE$ = DoPriCor$ & "-" & IBOLX$ 'Nro Lote Propio
       If U& < ULTREG&("!DETPORBU") Then
            XX$ = REGLEIDO$("!DETPORBU", U&)
          Else
            XX$ = REBLA$
       End If
       Call REPLA(XX$, LOTE$, 1, 16)
       Call GRAREG("!DETPORBU", XX$, U&)
    Next U&
    '
10  VTB% = VENTABU%("DETPORBU")
    If VTB% >= 0 Then
       CANTO# = 0: TCAN$ = ""
       For U& = 1 To CANBU&
          XX$ = REGLEIDO$("!DETPORBU", U&)
          CAPAR# = CVD(Mid$(REGLEIDO$("!DETPORBU", U&), 17, 8))
          If CAPAR# < 0.000005 Then
             Call MENSERR(24, "No Puede Haber Bultos Sin Cantidad.")
             GoTo 10
          End If
          CANTO# = CANTO# + CAPAR#
          TCAN$ = TCAN$ + TRIM$(FORMATNUM$(CAPAR#, "F13." & CNTDC$(CIXI%))) + " ; "
       Next U&
       '
       Text6 = FORMATNUM$(CANTO#, "F13." & CNTDC$(CIXI%))
       Call Text6_LostFocus
       Text1 = TCAN$
    End If
    '
End Sub

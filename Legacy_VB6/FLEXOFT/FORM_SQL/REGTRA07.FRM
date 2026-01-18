VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form REGTRA07 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Trazabilidad de Componentes"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Reservas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   135
      TabIndex        =   26
      Top             =   3450
      Width           =   8400
      Begin VB.Frame Frame5 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   6600
         TabIndex        =   27
         Top             =   480
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   5565
         TabIndex        =   29
         Top             =   480
         Width           =   15
      End
      Begin VB.Frame Frame6 
         Caption         =   "                "
         Height          =   1800
         Left            =   2000
         TabIndex        =   28
         Top             =   480
         Width           =   15
      End
      Begin VB.ListBox List2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         ItemData        =   "REGTRA07.frx":0000
         Left            =   190
         List            =   "REGTRA07.frx":0002
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   600
         Width           =   8025
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Código                    Descripción                                                  Cant.   Lote"
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
         Left            =   195
         TabIndex        =   31
         Top             =   360
         Width           =   8025
      End
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
      Height          =   300
      Left            =   8190
      TabIndex        =   0
      Top             =   180
      Width           =   175
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Actualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   135
      TabIndex        =   14
      Top             =   5925
      Width           =   8400
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
         Index           =   2
         Left            =   5085
         TabIndex        =   46
         Top             =   1470
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
         Height          =   300
         Index           =   1
         Left            =   5085
         TabIndex        =   41
         Top             =   1155
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
         Height          =   300
         Index           =   0
         Left            =   5085
         TabIndex        =   36
         Top             =   840
         Width           =   175
      End
      Begin VB.CommandButton Command3 
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
         Height          =   1350
         Left            =   7560
         Picture         =   "REGTRA07.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   420
         Width           =   660
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Asignado:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1365
         TabIndex        =   54
         Top             =   1220
         Width           =   1080
      End
      Begin VB.Label Label27 
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
         Left            =   1680
         TabIndex        =   53
         Top             =   1470
         Width           =   1140
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Requerido:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   52
         Top             =   1220
         Width           =   1080
      End
      Begin VB.Label Label24 
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
         Left            =   420
         TabIndex        =   51
         Top             =   1470
         Width           =   1140
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   5355
         TabIndex        =   50
         Top             =   1485
         Width           =   810
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3045
         TabIndex        =   49
         Top             =   1485
         Width           =   1770
      End
      Begin VB.Label Label5 
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
         Index           =   2
         Left            =   6285
         TabIndex        =   48
         Top             =   1485
         Width           =   1155
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
         Index           =   2
         Left            =   3570
         TabIndex        =   47
         Top             =   1470
         Width           =   1575
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   5355
         TabIndex        =   45
         Top             =   1170
         Width           =   810
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3045
         TabIndex        =   44
         Top             =   1170
         Width           =   1770
      End
      Begin VB.Label Label5 
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
         Index           =   1
         Left            =   6285
         TabIndex        =   43
         Top             =   1170
         Width           =   1155
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
         Index           =   1
         Left            =   3570
         TabIndex        =   42
         Top             =   1155
         Width           =   1575
      End
      Begin VB.Label Label6 
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
         Left            =   2415
         TabIndex        =   38
         Top             =   840
         Width           =   420
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
         Index           =   0
         Left            =   3570
         TabIndex        =   37
         ToolTipText     =   "Doble Click para Eliminar Asignacion"
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Origen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -940
         TabIndex        =   32
         Top             =   840
         Width           =   1770
      End
      Begin VB.Label Label10 
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
         Left            =   945
         TabIndex        =   21
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -940
         TabIndex        =   20
         Top             =   420
         Width           =   1770
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
         Left            =   945
         TabIndex        =   19
         Top             =   420
         Width           =   1905
      End
      Begin VB.Label Label5 
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
         Index           =   0
         Left            =   6285
         TabIndex        =   18
         ToolTipText     =   "Doble Click para Eliminar Asignacion"
         Top             =   855
         Width           =   1155
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3045
         TabIndex        =   17
         Top             =   855
         Width           =   1770
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   5355
         TabIndex        =   16
         Top             =   855
         Width           =   810
      End
      Begin VB.Label Label1 
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
         Left            =   3045
         TabIndex        =   3
         Top             =   420
         Width           =   4380
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consumos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   135
      TabIndex        =   8
      Top             =   945
      Width           =   8400
      Begin VB.Frame Frame1 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   6600
         TabIndex        =   25
         Top             =   480
         Width           =   15
      End
      Begin VB.Frame Frame14 
         Caption         =   "                "
         Height          =   1800
         Left            =   2000
         TabIndex        =   13
         Top             =   480
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "                                                                                                                  "
         Height          =   1800
         Left            =   5565
         TabIndex        =   10
         Top             =   480
         Width           =   15
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1680
         ItemData        =   "REGTRA07.frx":030E
         Left            =   190
         List            =   "REGTRA07.frx":0310
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   8025
      End
      Begin VB.Label Label25 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Código                    Descripción                                                  Cant.   Lote"
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
         Left            =   195
         TabIndex        =   9
         Top             =   360
         Width           =   8025
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   8400
      Left            =   8640
      ScaleHeight     =   8340
      ScaleWidth      =   1320
      TabIndex        =   4
      Top             =   -20
      Width           =   1380
      Begin VB.CommandButton Command8 
         Caption         =   "Rec. Reser."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   55
         Top             =   5160
         Width           =   690
      End
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "REGTRA07.frx":0312
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2850
         Width           =   690
      End
      Begin VB.CommandButton Command6 
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
         Picture         =   "REGTRA07.frx":061C
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Importacion de Articulos Trazables"
         Top             =   7800
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Excel 5.0;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   0
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5760
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   5880
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
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
         Picture         =   "REGTRA07.frx":0766
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Acceso a Maestro de Items de Stock"
         Top             =   2100
         Width           =   690
      End
      Begin VB.CommandButton Command16 
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
         Picture         =   "REGTRA07.frx":0A70
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Consulta de Trazabilidad de Componentes"
         Top             =   4350
         Width           =   690
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Kdx"
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
         Left            =   120
         Picture         =   "REGTRA07.frx":0BBA
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   3600
         Width           =   690
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
         Left            =   120
         Picture         =   "REGTRA07.frx":0EC4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Salir"
         Top             =   225
         Width           =   690
      End
      Begin VB.CommandButton Command10 
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
         Height          =   690
         Left            =   120
         Picture         =   "REGTRA07.frx":100E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   975
         Width           =   690
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   6
         Top             =   7155
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -285
         Picture         =   "REGTRA07.frx":1318
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   7440
         Width           =   1515
      End
   End
   Begin VB.Label Label13 
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
      Left            =   2700
      TabIndex        =   24
      Top             =   585
      Width           =   5655
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -705
      TabIndex        =   23
      Top             =   315
      Width           =   1545
   End
   Begin VB.Label Label11 
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
      Left            =   525
      TabIndex        =   22
      Top             =   585
      Width           =   1980
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
      Left            =   6735
      TabIndex        =   12
      Top             =   180
      Width           =   1500
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Número de O/F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4950
      TabIndex        =   11
      Top             =   225
      Width           =   1545
   End
End
Attribute VB_Name = "REGTRA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Command1.Enabled = False
3   Call SELECHO("ECORDEF/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
    NORFA$ = TRIM$(VALACT1$("ECORDEF"))
    If NORFA$ = "" Then GoTo 99
    '
    STATORF% = XVALO(VALOBADA("ORDEFABR", "StatuOrf", "IdSubOr = '" & NORFA$ & "'"))
    If STATORF% >= 8 Then
       Call MENSERR(24, "Orden '" + NORFA$ + "' Anulada")
       GoTo 99
    End If
    '
    Label28 = TRIM$(NORFA$)
    '
99  Command1.Enabled = True

End Sub

Private Sub Command11_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")

End Sub

Private Sub Command16_Click()
    '
    If DERACCE("CONTRAZA", "Consulta de Trazabilidad") > 1 Then
       CORETRA07.Show 1
    End If
    '
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   CODILLA$ = TRIM$(Label8)
   If CODILLA$ = "" Then CODILLA$ = TRIM$(Label11)
   FIKARDEX07.Text1 = CODILLA$
   FIKARDEX07.Show 1
   Command2.Enabled = True

End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    If DERACCE("REGTRAZA", "Registro de Trazabilidad") > 1 Then
       '
       Ci% = CODINT(Label8)
       If Ci% < 1 Then
          Call MENSERR(24, "Falta seleccionar Item")
          GoTo 99
       End If
       '
       SUMASI = 0
       For KKI% = 1 To Label4.Count
          If XVALO(Label5(KKI% - 1)) > 0 Then
             SUMASI = SUMASI + XVALO(Label5(KKI% - 1))
          End If
       Next KKI%
       '
       If SUMASI > XVALO(Label24) + FRACMIN(Ci%) Then
          Call MENSERR(24, "Suma de Asignados supera Requerido")
          GoTo 99
       End If
       Label27 = TRIM$(FORMATNUM(SUMASI, "F12." & CNTDC$))
       '
       NORFA$ = Label28
       If Label10 = "Consumos" Then
          CONDI$ = "DopriCor = '" & NORFA$ & "' AND CodiMovi = 'CF' AND Cod_Inte = " & Ci% & " AND OrdConsum = " & TRIM$(Label6)
          Dim MOVISE As ADODB.Recordset
          Set MOVISE = New ADODB.Recordset
          Set MOVISE = FLEXCONN.Execute(FILTSQL$("SELECT * FROM MOVISTO WITH(NOLOCK) WHERE " + CONDI$))
          CACAM& = MOVISE.Fields.Count
          ReDim VARS(CACAM&)
          For U& = 1 To CACAM&
             VARS(U&) = MOVISE.Fields(U& - 1).Value
          Next U&
          '
          LOTEVIE$ = TRIM$(VALOBADA("MOVISTO", "IdenLote", CONDI$))
          Call ACTUREGISTRO("Movisto", "CantIngre", CONDI$, 0, REGAF&, "NOMESS")
          Call ACTUREGISTRO("Movisto", "CantSalid", CONDI$, 0, REGAF&, "NOMESS")
          If TRIM$(LOTEVIE$) <> "" Then
             Call ACSALOTE(Ci%, LOTEVIE$, "NOMESS")
          End If
          Call BORRAREGISTROS("MOVISTO", CONDI$, RAFE&, "NOMESS")
          '
          Dim Movisto As ADODB.Recordset
          Set Movisto = New ADODB.Recordset
          strSQL = "SELECT * FROM MOVISTO WHERE " + CONDI$
25        On Error Resume Next
          Movisto.Open strSQL, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 25
          End If
          On Error GoTo 0
          '
          SUMASI = 0
          For KKI% = 1 To Label5.Count
             CANTASI = XVALO(Label5(KKI% - 1).Caption)
             If CANTASI >= FRACMIN(Ci%) Then
                With Movisto
                   .AddNew
                   For U& = 1 To CACAM&
                      Movisto.Fields(U& - 1).Value = VARS(U&)
                   Next U&
                   !CantSalid = CANTASI: SUMASI = SUMASI + CANTASI
                   !IDENLOTE = TRIM$(Label4(KKI% - 1))
                   .Update
                   Call ACSALOTE(Ci%, TRIM$(Label4(KKI% - 1)), "NOMESS")
                End With
             End If
          Next KKI%
          '
          RESTO = XVALO(Label24) - SUMASI
          If RESTO >= FRACMIN(Ci%) Then
            With Movisto
               .AddNew
               For U& = 1 To CACAM&
                  Movisto.Fields(U& - 1).Value = VARS(U&)
               Next U&
               !CantSalid = RESTO
               !IDENLOTE = ""
               .Update
            End With
          End If
          '
          Movisto.Close: Set Movisto = Nothing
          MOVISE.Close: Set MOVISE = Nothing
          Call Label28_Change
       End If
       '
       If Label10 = "Reservas" Then
          CONDI$ = " IdSubOr = '" & NORFA$ & "' AND Cod_Inte = " & Ci% & " AND OrdConsum = " & TRIM$(Label6)
          Dim RESERSE As ADODB.Recordset
          Set RESERSE = New ADODB.Recordset
          LISTADECAMPOS$ = "COD_INTE, COD_EXTE, DESCRIP, IDSUBOR, NUORITEM, STATUS, FECHRES, CANTRES, CANTCONS, SALDO_ENTRE, IDENLOTE, DEPOSITO, UBICALOT, USOUNIT, CODIEMPR, NUME_SOLINT, ORDCONSUM"
          Set RESERSE = FLEXCONN.Execute(FILTSQL$("SELECT " & LISTADECAMPOS$ & " FROM RESERVA WITH(NOLOCK) WHERE " + CONDI$))
          CACAM& = RESERSE.Fields.Count
          ReDim VARS(CACAM&)
          For U& = 1 To CACAM&
             VARS(U&) = RESERSE.Fields(U& - 1).Value
          Next U&
          '
          CantidadReservada# = VALOBADA("RESERVA", "SUM(CANTRES)", CONDI$)
          CantidadConsumida# = VALOBADA("RESERVA", "SUM(CantCons)", CONDI$)
          '
          Call ACTUREGISTRO("RESERVA", "CantRes", CONDI$, 0, REGAF&, "NOMESS")
          Call ACTUREGISTRO("RESERVA", "CantCons", CONDI$, 0, REGAF&, "NOMESS")
          If TRIM$(LOTEVIE$) <> "" Then
             Call ACSALOTE(Ci%, LOTEVIE$, "NOMESS")
          End If
          Call BORRAREGISTROS("RESERVA", CONDI$, RAFE&, "NOMESS")
          '
          Dim Reserva As ADODB.Recordset
          Set Reserva = New ADODB.Recordset
          strSQL = "SELECT " & LISTADECAMPOS$ & " FROM RESERVA WHERE " + CONDI$
35        On Error Resume Next
          Reserva.Open strSQL, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 35
          End If
          On Error GoTo 0
          '
          SUMASI = 0
          For KKI% = 1 To Label5.Count
             CANTASI = XVALO(Label5(KKI% - 1).Caption)
             If CANTASI >= FRACMIN(Ci%) Then
                With Reserva
                   .AddNew
                   For U& = 1 To CACAM&
                      Reserva.Fields(U& - 1).Value = VARS(U&)
                   Next U&
                   If KKI% > 1 Then !CANTCONS = 0
                   !CANTRES = XVALO(!CANTCONS) + CANTASI: SUMASI = SUMASI + CANTASI
                   !IDENLOTE = TRIM$(Label4(KKI% - 1))
                   .Update
                   Call ACSALOTE(Ci%, TRIM$(Label4(KKI% - 1)), "NOMESS")
                End With
             End If
          Next KKI%
          '
          RESTO = XVALO(Label24) - SUMASI
          If RESTO >= FRACMIN(Ci%) Then
            With Reserva
               .AddNew
               For U& = 1 To CACAM&
                  Reserva.Fields(U& - 1).Value = VARS(U&)
               Next U&
               !CANTRES = RESTO
               If SUMASI > 0 Then !CANTCONS = 0
               !IDENLOTE = ""
               .Update
            End With
          End If
          '
          Reserva.Close: Set Reserva = Nothing
          RESERSE.Close: Set RESERSE = Nothing
          Call Label28_Change
       End If
    End If
    '
99  Command3.Enabled = True
    '
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    If DERACCE%("AMASTO07", "Maestro de Items de Stock") > 1 Then
       AMASTOK07.Text1 = Label11
       AMASTOK07.Show 1
       Call Label28_Change
    End If
    Command4.Enabled = True
End Sub

Private Sub Command5_Click(Index As Integer)
    '
    CIXI% = CODINT%(Label8)
    If Label8 = "" Or CIXI% < 1 Then Exit Sub
    '
    Call SELECHO("!LILOTSE")
    If VALACT1$("!LILOTSE") <> "" Then
      'Valida si ya se Reservo lo demandado
       If XVALO(Label27) >= XVALO(Label24) Then
          Call COMUNI("La Demanda ya esta Completa")
          Exit Sub
       End If
       'Valida Seleccion Repetida
       XX$ = TRIM$(VALACT1$("!LILOTSE"))
       For KKI% = 1 To Label4.Count
         If KKI% - 1 <> Index Then
           If TRIM$(Label4(KKI% - 1)) = VALACT1$("!LILOTSE") Then
             COMUNI ("No Válido Seleccionar dos Veces Mismo Lote " & XX$ & ".")
             Label4(Index) = ""
             Exit Sub
           End If
         End If
       Next KKI%
       Label4(Index) = VALACT1$("!LILOTSE")
       '
       'Asigna Cantidad a Reservar del lote Seleccionado
       CONDI$ = "IDENLOTE ='" & Label4(Index) & "'"
       Call ACSALOTE(CIXI%, Label4(Index), "NOMESS")
       SALDO# = VALOBADA("LOTINGRE", "CANSALDO", CONDI$)
       DEPO% = XVALO(CONTROL$("", "DEPOCEN")) ' TOMA SALDOS DE LOTE DEL DEPOSITO PAÑOL UNICAMENTE
       If DEPO% > 0 Then SALDO# = SaldoLote(CIXI%, TRIM$(Label4(Index)), DEPO%)
       SALDO# = SALDO# - CANTRESER(CIXI%, Label4(Index))
       If SALDO# >= XVALO(Label24) Then 'Si el lote alcanza, se asigna lo demandado menos lo ya asignado
            Label5(Index) = TRIM$(FORMATNUM$(XVALO(Label24) - XVALO(Label27), "F11." + CNTDC$))
          Else
            Label5(Index) = TRIM$(FORMATNUM(SALDO#, "F11." + CNTDC$))
       End If
       Call SumarReservado
       '
    End If
    '
End Sub

Function CANTRESER(CIXI%, LOTE$)
'
DEPO% = XVALO(CONTROL$("", "DEPOCEN"))
SQLX$ = "SELECT * FROM RESERVA WITH(NOLOCK) WHERE COD_INTE=" & CIXI% & " AND IDENLOTE='" & LOTE$ & "'"
Set rs = READSET(SQLX$)
CRTOTAL = 0
'
With rs
    Do While Not .EOF
        ORFA$ = SAFETEXT$(!idsubor)
        CR = XVALO(!CANTRES)
        CONDI$ = "COD_INTE=" & CIXI% & " AND DOSECCOR='" & ORFA$ & "' AND IDENLOTE='" & LOTE$ & "' AND CODIMOVI='TR' "
        If DEPO% > 0 Then CONDI$ = CONDI$ & " AND DEPOMOVI=" & DEPO%
        CANTR = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$))
        CRL = CR - CANTR: If CRL < 0 Then CRL = 0
        CRTOTAL = CRTOTAL + CRL
    .MoveNext
    Loop
End With
'
CANTRESER = CRTOTAL
'
End Function

Private Sub Command6_Click()
    '
    Command6.Enabled = False
    '
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Archivo Origen de Datos"
    '     cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
    CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0

    DoEvents
    Screen.MousePointer = 11
    Data1.DatabaseName = UCase$(CommonDialog1.FileName)
    On Error Resume Next
    Data1.Refresh
    Data1.RecordSource = Data1.Database.TableDefs(0).Name
    Data1.Refresh
    On Error GoTo 0
    
    With Data1
     .Recordset.MoveFirst
     Do While Not .Recordset.EOF
        CODEXTT$ = ""
        CODEXTT$ = .Recordset.Fields(0)
        Ci% = CODINT%(CODEXTT$)
        CONDI$ = "Codint = " & Ci% & ""
        If Ci% <= 0 Then
            Call MENSERR(24, "Código '" & CODEXTT$ & "' Inexistente o Dado de Baja.")
            GoTo 20
        End If
        TRAZABL% = XVALO(VALOBADA("MASTER", "Trazable", CONDI$))
        If TRAZABL% = 2 Then
            Call MENSERR(24, "Código " & CODEXTT$ & " Trazable por Número de Serie.")
            GoTo 20
        End If
        '
        ESTRAZA% = 1
        '
        Call ACTUREGISTRO("MASTER", "Trazable", CONDI$, ESTRAZA%, REGAF&)
20      Screen.MousePointer = 11
       .Recordset.MoveNext
     Loop
    End With
    Screen.MousePointer = 1
    Call MENSERR(24, "Proceso de Actualización Masiva Completo")
    '
99  Command6.Enabled = True
    '
End Sub

Private Sub Command7_Click()
  '
  If DERACCE%("TRANEQUI", "Tranferencia de Equipos") < 2 Then
    GoTo 99
  End If
  '
  Screen.MousePointer = 11
  Command7.Enabled = False
  '
  Dim TANUM As ADODB.Recordset
  Set TANUM = New ADODB.Recordset
  SQLX$ = "SELECT idsubor, sum(deponumse) as SxDepoNumse from Tanumse WITH(NOLOCK) where deponumse < 0 group BY IDSUBOR"
  '
  Set TANUM = FLEXCONN.Execute(FILTSQL$(SQLX$))
  If TANUM.BOF = True And TANUM.EOF = True Then GoTo 99
  With TANUM
    If .BOF = True And .EOF = True Then GoTo 99
    .MoveFirst
    Do While Not (.EOF)
      OF$ = SAFETEXT(!idsubor)
      SALDO# = VALOBADA("ORDEFABR", "CANPROY - CANAPRO", "IdSubOr = '" & OF$ & "'")
      EQUIENPRO = Abs(XVALO(!SXDEPONUMSE))
      If Abs(!SXDEPONUMSE) <> SALDO# Then
         TTXX$ = "Orden de Fabricacion '" & OF$ & "'\no Coinciden los Equipos en Proceso (" & EQUIENPRO & ")\con el Saldo de la O/Fabricacion (" & SALDO# & ")"
         If SALDO# > 0 Then
             Call COMUNI(TTXX$)
           Else
             TTXX$ = TTXX$ + "\   \¿ Desea Ingresar Equipos a Stock\para Reparar esta Diferencia ?\\No, Dejar como está\Si, ingresar a Stock"
             If COMALTER%(TTXX$) = 2 Then
                Call SELECHO("TADEPOSI")
                DEPO% = XVALO(VALACT1$("TADEPOSI")) '  SELECCIONAR EL DEPOSITO DE ENTRADA
                If DEPO% > 0 Then
                  Dim Equipos As ADODB.Recordset
                  Set Equipos = New ADODB.Recordset
                  SQLX$ = "Select * from Tanumse WITH(NOLOCK) where deponumse < 0 AND IDSUBOR = '" & OF$ & "'"
                 
                  Set Equipos = FLEXCONN.Execute(FILTSQL$(SQLX$))
                  With Equipos
                    .MoveFirst
                    Do While Not (.EOF)
                      Serie$ = SAFETEXT(!NUMEROSERIE)
                      Call ANOTANUSE(Serie$, "Transferencia Manual de Proceso a Producto Terminado") 'ESCRIBIR REGISTRO EN HISTORIAL DE EQUIPOS
                      .MoveNext
                    Loop
                  End With
                  CONDI$ = "deponumse < 0 AND IDSUBOR='" & OF$ & "' "
                  Call ACTUREGISTRO("TANUMSE", "DEPONUMSE", CONDI$, DEPO%, REG&)
 '                HACER ACTUREGISTRO EN EL TANUMSE CON EL DEPOSITO DE DESTINO
                End If
             End If
         End If
      End If
      .MoveNext
    Loop
  End With
  
99  Command7.Enabled = True
    Screen.MousePointer = 1

End Sub

Private Sub Command8_Click()
    '
    ' PROCESO DE CORRECCION DE RESERVAS
    ' 1) VERIFICA SI LA RESERVA DE LA ORDEN ACTIVA ES CORRECTA
    ' 2) CORRIGE NO TRAZABLES Y AJUSTA CANTIDAD DE LO TRAZABLES SEGUN EL CONSUMO
    ' 3) AGREGAR UN NUEVO REGISTRO PARA DEJAR LA RESERVA A LO NECESARIO SEGUN LA FORMULA DE PRODUCTO
    '
    Command8.Enabled = False
    '
    If DERACCE%("RECURESE", "Proceso Recuperacion de Reservas") < 2 Then
       GoTo 99
    End If
    '
    'VALIDA QUE HAYA UNA O-FABRICACION ACTIVA
    NROOF$ = TRIM$(Label28)
    If NROOF$ = "" Then GoTo 99
    '
    CONDI$ = "IDSUBOR ='" & NROOF$ & "'"
    If CantRegistros("ORDEFABR", CONDI$, "NOMESS") < 1 Then GoTo 99
    '
    'VERIFICA SI HAY ERROR EN LA RESERVA DE LA O-FABRICACION ACTIVA
    If VERIFICARESERVA(NROOF$) = 1 Then
        Call COMUNI("Reserva de '" & NROOF$ & "' Correcta")
        GoTo 99
    End If
    '
    OPX% = COMALTER%("Se Han Encontrado Diferencias en la Reserva de esta O-Fabricacion\¿Desea Corregirlas?\\No, Continuar\Si, Corregir")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    If OPX% = 1 Then GoTo 99
    '
    'CORRIGE LAS RESERVAS DE LOS NO TRAZABLES Y CON LOS TRAZABLES RESERVADO = CONSUMIDO
    Call ABRESTOCKS
    SQLX$ = "SELECT RESERVA.COD_INTE AS COD1, ORDEFABR.COD_INTE AS COD2, ORDEFABR.IDSUBOR, ORDEFABR.CanProy, RESERVA.CANTRES, RESERVA.CANTCONS, RESERVA.SALDO_ENTRE, RESERVA.IDENLOTE "
    SQLX$ = SQLX$ + " FROM RESERVA INNER JOIN ORDEFABR ON RESERVA.IDSUBOR = ORDEFABR.IDSUBOR "
    SQLX$ = SQLX$ + " WHERE ORDEFABR.IDSUBOR='" & NROOF$ & "'"
    
    Dim RST As ADODB.Recordset
    Set RST = New ADODB.Recordset
    RST.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

    With RST
        Do While Not .EOF
            CICOMP% = XVALO(!COD1) 'CODIGO INTERNO DEL COMPONENTE
            CICONJ% = XVALO(!COD2) 'CODIGO INTERNO DEL CONJUNTO
            '
            'OBTIENE CANTIDAD NECESARIA DEL COMPONENTE SEGUN FORMULA
            CONDI$ = "CODICONJ=" & CICONJ% & " AND CODIELEM=" & CICOMP%
            CANNEC = XVALO(VALOBADA("FORMULAS", "USOBRUTO", CONDI$, "NOMESS")) * XVALO(!CANPROY)
            '
            'OBTIENE TOTAL CONSUMIDO DEL ITEM PARA LA OF
            DOPRI$ = SAFETEXT$(!idsubor)
            DOSEC$ = "PF-" + TRIM$(Mid$(DOPRI$, 4))
            CONDI$ = "CODIMOVI = 'CF' AND COD_INTE = " & CICOMP%
            CONDI$ = CONDI$ & " AND DOPRICOR = '" & DOPRI$ & "'"
            CONDI$ = CONDI$ & " AND DOSECCOR = '" & DOSEC$ & "'"
            IDLO$ = TRIM$(SAFETEXT$(!IDENLOTE))
            If IDLO$ <> "" Then CONDI$ = CONDI$ & " AND IDENLOTE = '" & IDLO$ & "'"
            CanConsu = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$, "NOMESS"))
            '
            'CORRIGE RESERVA
            'PARA LOS ITEMS NO TRAZABLES LA CANTIDAD RESERVADA ES LA QUE PROVIENE DE LA FORMULA MULTIPLICADA POR LA CANTIDAD A FABRICAR
            CANRES = CANNEC
            If ESTRAZABLE%(CICOMP%) = 1 And CanConsu > 0 Then CANRES = CanConsu 'SI ES TRAZABLE LA CANTIDAD RESERVADA ES IGUAL A LA CONSUMIDA
            !CANTRES = CANRES                 'CORRIGE CANTIDAD RESERVADA
            !CANTCONS = CanConsu              'CORRIGE CANTIDAD CONSUMIDA
            SALDORES = !CANTRES - !CANTCONS
            If SALDORES < 0 Then SALDORES = 0 'ATAJA SOBRECONSUMO
            !Saldo_Entre = SALDORES           'CORRIGE SALDO
            .Update                           'ACTUALIZA
    
            .MoveNext
        Loop
    End With
    On Error Resume Next
    RST.Close
    Set RST = Nothing
    On Error GoTo 0
    '
    'CORRIGE LAS RESERVAS AGREGANDO REGISTROS PARA COMPENSAR SEGUN FORMULA DE PRODUCTO
    SQLX$ = "SELECT * FROM ORDEFABR INNER JOIN FORMULAS ON ORDEFABR.COD_INTE = FORMULAS.CODICONJ "
    SQLX$ = SQLX$ + " WHERE IDSUBOR='" & NROOF$ & "'"
    
    Dim RST1 As ADODB.Recordset
    Set RST1 = New ADODB.Recordset
    RST1.Open (SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

    With RST1
        Do While Not .EOF
            CICONJ% = XVALO(!CODICONJ) 'CODIGO INTERNO DEL CONJUNTO
            CICOMP% = XVALO(!CODIELEM) 'CODIGO INTERNO DEL COMPONENTE
            STAT% = XVALO(VALOBADA("MASTER", "STATUS", "CODINT=" & CICOMP%, "NOMESS")) 'VERIFICA SI ESTA DE BAJA
            Codigo$ = CODEXT$(CICOMP%)
            If STAT% > 0 Then
                USOUNI = XVALO(!USOBRUTO)   'USO UNITARIO
                CANTFAB = XVALO(!CANPROY)   'CANTIDAD PROYECTADA
                CANNEC = USOUNI * CANTFAB   'CANTIDAD NECESARIA
                '
                'OBTIENE CANTIDAD RESERVADA EN TABLA RESERVA
                IDLO$ = SAFETEXT$(!idsubor) 'ORDEN DE FABRICACION
                CONDI$ = "COD_INTE=" & CICOMP% & " AND IDSUBOR='" & IDLO$ & "'"
                CANRES = XVALO(VALOBADA("RESERVA", "SUM(CANTRES)", CONDI$, "NOMESS")) 'CANTIDAD TOTAL RESERVADA
                '
                'GENERA NUEVO REGISTRO POR FALTANTE DE RESERVA
                If CANRES < CANNEC Then
                    FALTANTE = CANNEC - CANRES
10                  RNDX$ = RANDSTRING$(-8)
                    If CantRegistros("RESERVA", "IDSUBOR = '" & RNDX$ & "'", "NOMESS") > 0 Then GoTo 10
                    Call CreaRegistro("RESERVA", "IDSUBOR", RNDX$, "NOMESS")
                    Call ACTUREGISTRO("RESERVA", "COD_INTE", "IDSUBOR = '" & RNDX$ & "'", CICOMP%, RAFE&, "NOMESS")
                       Codigo$ = CODEXT$(CICOMP%)
                    Call ACTUREGISTRO("RESERVA", "COD_EXTE", "IDSUBOR = '" & RNDX$ & "'", Codigo$, RAFE&, "NOMESS")
                       DESCRI$ = DESCRIT0$(CICOMP%)
                    Call ACTUREGISTRO("RESERVA", "DESCRIP", "IDSUBOR = '" & RNDX$ & "'", DESCRI$, RAFE&, "NOMESS")
                       NUOR = XVALO(VALOBADA("RESERVA", "Max(NUORITEM)", "IDSUBOR='" & IDLO$ & "'", "NOMESS"))
                       NUOR = NUOR + 1
                    Call ACTUREGISTRO("RESERVA", "NUORITEM", "IDSUBOR = '" & RNDX$ & "'", NUOR, RAFE&, "NOMESS")
                    Call ACTUREGISTRO("RESERVA", "STATUS", "IDSUBOR = '" & RNDX$ & "'", 0, RAFE&, "NOMESS")
                    Call ACTUREGISTRO("RESERVA", "FECHRES", "IDSUBOR = '" & RNDX$ & "'", HOY(HO$), RAFE&, "NOMESS")
                    Call ACTUREGISTRO("RESERVA", "CANTRES", "IDSUBOR = '" & RNDX$ & "'", FALTANTE, RAFE&, "NOMESS")
                    Call ACTUREGISTRO("RESERVA", "CANTCONS", "IDSUBOR = '" & RNDX$ & "'", 0, RAFE&, "NOMESS")
                    Call ACTUREGISTRO("RESERVA", "SALDO_ENTRE", "IDSUBOR = '" & RNDX$ & "'", FALTANTE, RAFE&, "NOMESS")
                    Call ACTUREGISTRO("RESERVA", "IDENLOTE", "IDSUBOR = '" & RNDX$ & "'", "", RAFE&, "NOMESS")
                    Call ACTUREGISTRO("RESERVA", "DEPOSITO", "IDSUBOR = '" & RNDX$ & "'", 0, RAFE&, "NOMESS")
                    Call ACTUREGISTRO("RESERVA", "IDENLOTE", "IDSUBOR = '" & RNDX$ & "'", "", RAFE&, "NOMESS")
                    Call ACTUREGISTRO("RESERVA", "UBICALOT", "IDSUBOR = '" & RNDX$ & "'", "", RAFE&, "NOMESS")
                    Call ACTUREGISTRO("RESERVA", "USOUNIT", "IDSUBOR = '" & RNDX$ & "'", USOUNI, RAFE&, "NOMESS")
                    Call ACTUREGISTRO("RESERVA", "USOUNIT", "IDSUBOR = '" & RNDX$ & "'", USOUNI, RAFE&, "NOMESS")
                    Call ACTUREGISTRO("RESERVA", "OBSERVA", "IDSUBOR = '" & RNDX$ & "'", "", RAFE&, "NOMESS")
                    Call ACTUREGISTRO("RESERVA", "MARBORRA", "IDSUBOR = '" & RNDX$ & "'", 0, RAFE&, "NOMESS")
                    '
                    Call ACTUREGISTRO("RESERVA", "IDSUBOR", "IDSUBOR = '" & RNDX$ & "'", IDLO$, RAFE&, "NOMESS")
                End If
            End If
            .MoveNext
        Loop
    End With
    On Error Resume Next
    RST1.Close
    Set RST1 = Nothing
    On Error GoTo 0
    '
    'ACTUALIZA SALDO DE LOTES
    SQLX$ = "SELECT * FROM RESERVA WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE IDSUBOR='" & NROOF$ & "'"
    
    Dim RST2 As ADODB.Recordset
    Set RST2 = READSET(SQLX$)

    With RST2
        Do While Not .EOF
            CICOMP% = XVALO(!COD_INTE) 'CODIGO INTERNO DEL COMPONENTE
            IDLO$ = SAFETEXT$(!IDENLOTE) 'LOTE DEL COMPONENTE
            If ESTRAZABLE%(CICOMP%) = 1 Then
                Call ACSALOTE(CICOMP%, IDLO$) 'ACTUALIZA SALDOS DEL LOTE
            End If
            .MoveNext
        Loop
    End With
    On Error Resume Next
    RST2.Close
    Set RST2 = Nothing
    On Error GoTo 0
    '
    Call COMUNI("Proceso de Recuperacion de Reserva Finalizado")
    '
    'PARA QUE ACTUALICE LA INFORMACION EN PANTALLA
    Label28 = ""
    Label28 = NROOF$
    '
99  Command8.Enabled = True
End Sub

Function VERIFICARESERVA(NROOF$)
    '
    valor = 1
    NROOF1$ = TRIM$(NROOF$)
    If NROOF1$ = "" Then Exit Function
    '
    SQLX$ = "SELECT * FROM ORDEFABR INNER JOIN FORMULAS ON ORDEFABR.COD_INTE = FORMULAS.CODICONJ "
    SQLX$ = SQLX$ + " WHERE ORDEFABR.IDSUBOR='" & NROOF1$ & "'"
    
    Dim RST As ADODB.Recordset
    Set RST = FLEXCONN.Execute(FILTSQL$(SQLX$))

    With RST
        Do While Not .EOF
            CICONJ% = XVALO(!CODICONJ) 'CODIGO INTERNO DEL CONJUNTO
            CICOMP% = XVALO(!CODIELEM) 'CODIGO INTERNO DEL COMPONENTE
            Codigo$ = CODEXT$(CICOMP%)
            STAT% = XVALO(VALOBADA("MASTER", "STATUS", "CODINT=" & CICOMP%, "NOMESS")) 'VERIFICA SI ESTA DE BAJA
            If STAT% > 0 Then
                'CALCULA NECESIDAD DE RESERVA SEGUN FORMULA DE PRODUCTO
                USOUNI = XVALO(!USOBRUTO)   'USO UNITARIO
                CANTFAB = XVALO(!CANPROY)   'CANTIDAD PROYECTADA
                CANNEC = USOUNI * CANTFAB   'CANTIDAD NECESARIA SEGUN FORMULA
                '
                'OBTIENE CANTIDAD RESERVADA EN TABLA RESERVA
                IDLO$ = SAFETEXT$(!idsubor) 'ORDEN DE FABRICACION
                CONDI$ = "COD_INTE=" & CICOMP% & " AND IDSUBOR='" & IDLO$ & "'"
                CANRES = XVALO(VALOBADA("RESERVA", "SUM(CANTRES)", CONDI$, "NOMESS")) 'CANTIDAD TOTAL RESERVADA
                '
                If CANRES < CANNEC Then
                    valor = 0 'ERROR EN LA RESERVA
                    Exit Do
                End If
            End If
            .MoveNext
        Loop
    End With
    On Error Resume Next
    RST.Close
    Set RST = Nothing
    On Error GoTo 0
    '
    VERIFICARESERVA = valor
    '
End Function

Private Sub Form_Load()
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    Call ABRECONEXION
    '
    If APLINVO$ = "CONSUTRA" Then
       CORETRA07.Show 1
       Call FINAL("")
    End If
End Sub


Private Sub Label28_Change()
10 Call LIMPIAR
   NORFA$ = Label28
   Ci% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", "IdSubOr = '" & NORFA$ & "'"))
   If Ci% > 0 Then
     Label11 = CODEXT(Ci%)
     Label13 = DESCRIT$(Ci%)
     CONDI$ = " DopriCor = '" & NORFA$ & "' AND CodiMovi = 'CF'"
     Call ENUMERA("MOVISTO", "OrdConsum", CONDI$)
     Call CARGALISTA(List1, "MOVISTO", "Cod_Exte/A16; Descrip/A32; CantSalid/F10.2;IdenLote/A16; OrdConsum/F12", CONDI$)
     i& = 1
     '
     ' ELIMINA LOS ITEMS NO TRAZABLES
     Do While i& <= List1.ListCount
       CIXI% = CODINT%(TRIM$(Left$(List1.LIST(i& - 1), 16)))
       If ESTRAZABLE%(CIXI%) > 0 Then
         i& = i& + 1
       Else
         List1.RemoveItem (i& - 1)
       End If
     Loop
     '
     'CONDI$ = " IdSubOr = '" & NORFA$ & "'"
     'SE MODIFICA PARA QUE NO TRAIGA LAS RESERVAS YA CONSUMIDAS
     CONDI$ = " IdSubOr = '" & NORFA$ & "' AND (CANTRES-CANTCONS > 0.05)"
     Call ENUMERA("RESERVA", "OrdConsum", CONDI$)
     Call CARGALISTA(List2, "RESERVA", "Cod_Exte/A16; Descrip/A32; (CantRes - CantCons)/F10.2;IdenLote/A16; OrdConsum/F12", CONDI$)
20   ' ELIMINA LOS ITEMS NO TRAZABLES
     For i& = 1 To List2.ListCount
        CIXI% = CODINT%(TRIM$(Left$(List2.LIST(i& - 1), 16)))
        If ESTRAZABLE%(CIXI%) < 1 Then
           List2.RemoveItem (i& - 1)
           GoTo 20
        End If
     Next i&
     '
     If List1.ListCount + List2.ListCount < 1 Then
        Call MENSERR(24, "Este Producto no Registra Componentes Trazables.")
        Call Command4_Click
     End If
     '
     ' VERIFICA SOBRE-CONSUMO
     For i& = 1 To List2.ListCount
       SALRE = XVALO(Mid$(List2.LIST(i& - 1), 51, 10))
       If SALRE < 0 Then
          Call RECUPERACONSUMO(Ci%, NORFA$)
          Call COMUNI("MOVIMIENTOS RECUPERADOS")
          GoTo 10
       End If
     Next i&
   End If

End Sub

Sub RECUPERACONSUMO(CIXI%, NORFAX$)
   '
   Dim CANCONSUMI(128), FECONSUMI%(128), DEPCONSUMI%(128)
   CAMORE% = 0
   '
   CMOV$ = "CF"
   DOPRI$ = NORFAX$
   DOSEC$ = "PF-" + TRIM$(Mid$(NORFAX$, 4))
   '
   CONDI$ = "CODIMOVI = '" & CMOV$ & "' "
   CONDI$ = CONDI$ & " AND COD_INTE = " & CIXI%
   CONDI$ = CONDI$ & " AND DOPRICOR = '" & DOPRI$ & "'"
   CONDI$ = CONDI$ & " AND DOSECCOR = '" & DOSEC$ & "'"
   SQLX$ = "SELECT * FROM MOVISTO WITH(NOLOCK) WHERE " + CONDI$
   Set RSMOVI = READSET(SQLX$)
   '
   With RSMOVI
      Do While Not .EOF
        CAMORE% = CAMORE% + 1
        CANCONSUMI(CAMORE%) = XVALO(!CantSalid)
        FECONSUMI%(CAMORE%) = CVINT(!FECHMOV)
        DEPCONSUMI(CAMORE%) = XVALO(!DepoMovi)
      .MoveNext
      Loop
   End With
   RSMOVI.Close
   Set RSMOVI = Nothing
   '
   ' BORRA EL CONSUMIDO DE TODAS LAS RESERVAS
   CONDI2$ = " IdSubOr = '" & NORFAX$ & "' AND COD_INTE = " & CIXI%
   strSQL = "Select * from RESERVA WHERE " + CONDI2$
   '
   ' BORRA LOS MOVIMIENTOS DE STOCK
   Call BORRAREGISTROS("MOVISTO", CONDI$, RAFE&)
   '
   ' RE-GENERA LOS MOVIMIENTOS
   For KI% = 1 To CAMORE%
      Call CONSUMECOMPO(CIXI%, NORFAX$, CANCONSUMI(KI%), FECONSUMI%(KI%), DEPCONSUMI%(KI%))
   Next KI%
   '
End Sub

Private Sub Label5_Click(Index As Integer)
     Label4(Index) = ""
     Call SumarReservado
End Sub

Private Sub Label4_Click(Index As Integer)
     Label4(Index) = ""
     Label5(Index) = ""
     Call SumarReservado
End Sub

Private Sub Label8_Change()
     CIXI% = CODINT%(Label8)
     DEPO% = XVALO(CONTROL$("", "DEPOCEN")) ' TOMA SALDOS DE LOTE DEL DEPOSITO PAÑOL UNICAMENTE
     If CIXI% > 0 Then
        Call CREALOTES(CIXI%)
        Call ARMALISLOTES(CIXI%, DEPO%, , , 1)
     End If
End Sub

Private Sub List1_DblClick()
     Call LIMPIAR
     TTXX$ = List1.LIST(List1.ListIndex)
     Label8 = Mid$(TTXX$, 1, 16)
     Label10 = "Consumos"
     Label1 = TRIM$(Mid$(TTXX$, 17, 32))
     Label5(0) = Mid$(TTXX$, 50, 11)
     Label4(0) = Mid$(TTXX$, 62, 16)
     Label24 = TRIM$(Label5(0))
     Label6 = TRIM$(Mid$(TTXX$, 79, 12))
     Call SumarReservado
End Sub

Private Sub List2_DblClick()
     Call LIMPIAR
     TTXX$ = List2.LIST(List2.ListIndex)
     Label8 = Mid$(TTXX$, 1, 16)
     Label10 = "Reservas"
     Label1 = TRIM$(Mid$(TTXX$, 17, 32))
     Label5(0) = Mid$(TTXX$, 50, 11)
     Label4(0) = Mid$(TTXX$, 62, 16)
     Label24 = Mid$(TTXX$, 50, 11) 'El Total Demandado
     Label6 = TRIM$(Mid$(TTXX$, 79, 12))
     Call SumarReservado
End Sub

Sub LIMPIAR()
     Label8 = ""
     Label10 = ""
     Label1 = ""
     Label24 = ""
     Label27 = ""
     Label6 = ""
     For KKI% = 1 To Label5.Count
       Label4(KKI% - 1) = "": Label5(KKI% - 1) = ""
     Next KKI%
End Sub

Sub SumarReservado()
    SUMASI = 0
    For KKI% = 1 To Label5.Count
      If TRIM$(Label4(KKI% - 1)) <> "" Then
         SUMASI = SUMASI + XVALO(Label5(KKI% - 1))
      End If
    Next KKI%
    Label27 = FORMATNUM(SUMASI, "F11." + CNTDC$)
End Sub

'Sub ACOCONSUM(NORFA$)
'   NORFA$ = Label28
'   CI% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", "IdSubOr = '" & NORFA$ & "'"))
'   If CI% > 0 Then
'     '
'     ' OBTIENE CANTIDADES CONSUMIDAS POR CODIGO Y LOTE
'     SQLX$ = "SELECT COD_INTE,IDENLOTE,SUM(CANTSALID) AS CONSUM FROM MOVISTO WITH(NOLOCK) "
'     SQLX$ = SQLX$ + " INNER JOIN MASTER ON COD_INTE = CODINT "
'     SQLX$ = SQLX$ + " WHERE MASTER.TRAZABLE = 1 "
'     SQLX$ = SQLX$ + " AND MOVISTO.DopriCor = '" & NORFA$ & "' AND MOVISTO.CodiMovi = 'CF' "
'     SQLX$ = SQLX$ + " GROUP BY MOVISTO.COD_INTE,MOVISTO.IDENLOTE "
'     RSMOVI = READSET(SQLX$)
'     '
'     With RSMOVI
'        Do While Not .EOF
'           CIXI% = !COD_INTE
'           IDLO$ = !IDENLOTE
'           CONDI$ = " IdSubOr = '" & NORFA$ & "' AND COD_INTE = " & CIXI% & " AND IDENLOTE = '" & IDLO$ & "'"
'           CANTIRES = XVALO(VALOBADA("RESERVA", "CANTRES", CONDI$))
'           If !CONSUM > CANTIRES + FRACMIN(CIXI%) / 2 Then
'              ' HAY QUE RECUPERAR LOS CONSUMOS
'           End If
'        .MoveNext
'        Loop
'     End With
'     '
'
'     ' AQUI COMIENZA LA RECUPERACION DEL CODIGO CIXI% DE LA OF NORFA$
'     CONDI$ = "DopriCor = '" & NORFA$ & "' AND CodiMovi = 'CF'  And COD_INTE = " & CIXI%
'     TOCONSUMOVI = XVALO(VALOBADA("MOVISTO", "SUM(CANTSALID)", CONDI$))
'     SALCONSU = TOCONSUMOVI
'     If SALCONSU < 0 Then SALCONSU = 0
'     '
'     CONDI2$ = " IdSubOr = '" & NORFA$ & "' AND COD_INTE = " & CIXI%   '  & " AND IDENLOTE = '" & IDLO$ & "'"
'     STRSQL = "Select * from RESERVA WHERE " + CONDI2$
'     Dim RRES As ADODB.Recordset
'     Set RRES = New ADODB.Recordset
'25   On Error Resume Next
'     RRES.Open (STRSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'        On Error GoTo 0
'        Call CapturaErrorOpen
'        GoTo 25
'     End If
'     On Error GoTo 0
'     '
'     U& = 0
'     With RRES
'       Do While Not .EOF
'         U& = U& + 1
'         Call TRACE(20, U&, FIN&)
'         QRES = XVALO(!CANTRES)
'         QCON = QRES: If QCON > SALCONSU Then QCON = SALCONSU
'         !CANTCONSU = QCON
'         .Update
'         SALCONSU = SALCONSU - QCON
'       .MoveNext
'       Loop
'     End With
'     RRES.Close
'     Set RRES = Nothing
'     CONSUMONORESERVADO = SALCONSU: If CONSUMONORESERVADO < 0 Then CONSUMONORESERVADO = 0
'     '
'     ' AQUI ESTA ARREGLADO EL TEMA DEL SOBRE-CONSUMO DE LA TABLA RESERVA. FALTA TRASLADAR ESTO AL MOVISTO AJUSTANDO
'     ' LOS LOTES A LO EFECTIVAMENTE CONSUMIDO
'     '
'     CONDI3$ = "COD_INTE = " & CIXI% & " AND DopriCor = '" & NORFA$ & "' AND CodiMovi = 'CF' "
'     STRSQL = "Select * from MOVISTO WHERE " + CONDI3$
'     Dim RMOV As ADODB.Recordset
'     Set RMOV = New ADODB.Recordset
'35   On Error Resume Next
'     RMOV.Open (STRSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'        On Error GoTo 0
'        Call CapturaErrorOpen
'        GoTo 35
'     End If
'     On Error GoTo 0
'     '
'     U& = 0
'     With RMOV
'       Do While Not .EOF
'         U& = U& + 1
'         Call TRACE(20, U&, FIN&)
'         QMOVI = XVALO(!CANTSALID)
'         QCON = QRES: If QCON > SALCONSU Then QCON = SALCONSU
'         !CANTCONSU = QCON
'         .Update
'         SALCONSU = SALCONSU - QCON
'       .MoveNext
'       Loop
'     End With
'     RRES.Close
'     Set RRES = Nothing
'     CONSUMONORESERVADO = SALCONSU: If CONSUMONORESERVADO < 0 Then CONSUMONORESERVADO = 0
'     '
'     Call ENUMERA("MOVISTO", "OrdConsum", CONDI$)
'
'End Sub
'

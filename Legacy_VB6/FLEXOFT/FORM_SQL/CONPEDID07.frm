VERSION 5.00
Begin VB.Form CONPEDID07 
   BackColor       =   &H00E0FFD0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Pedidos por Número"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command8 
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
      Height          =   330
      Left            =   10970
      TabIndex        =   51
      Top             =   3550
      Width           =   750
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   10970
      ScaleHeight     =   75
      ScaleWidth      =   705
      TabIndex        =   49
      Top             =   4020
      Width           =   770
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   50
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command2 
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
      Left            =   10970
      Picture         =   "CONPEDID07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1590
      Width           =   750
   End
   Begin VB.CommandButton Command1 
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
      Height          =   650
      Left            =   10970
      Picture         =   "CONPEDID07.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Graba las Modificaciones"
      Top             =   950
      Width           =   750
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Opcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3350
      Left            =   10920
      TabIndex        =   43
      Top             =   4250
      Width           =   875
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "CONPEDID07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Preparación de Pedidos"
         Top             =   2050
         Width           =   650
      End
      Begin VB.CommandButton Command5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "CONPEDID07.frx":0EDE
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Generar Orden de Fabricación"
         Top             =   1450
         Width           =   645
      End
      Begin VB.CommandButton Command24 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   95
         Picture         =   "CONPEDID07.frx":11E8
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Despacho - Entrega al Cliente"
         Top             =   2650
         Width           =   650
      End
      Begin VB.CommandButton Command21 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   95
         Picture         =   "CONPEDID07.frx":14F2
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Consultar Pendientes de Despacho"
         Top             =   250
         Width           =   650
      End
      Begin VB.CommandButton Command23 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   95
         Picture         =   "CONPEDID07.frx":17FC
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Listado Pendientes de Despacho"
         Top             =   850
         Width           =   650
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "View"
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
      Left            =   10970
      Picture         =   "CONPEDID07.frx":1B06
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Consultas y Re-Impresión"
      Top             =   2900
      Width           =   750
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Pedido / Orden de Venta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1550
      Left            =   105
      TabIndex        =   21
      Top             =   210
      Width           =   5370
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2205
         TabIndex        =   24
         Top             =   315
         Width           =   180
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4830
         TabIndex        =   37
         ToolTipText     =   "Doble Click para Restaurar Situación de Preparacion y Despacho"
         Top             =   315
         Width           =   330
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "St:"
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
         Left            =   3465
         TabIndex        =   36
         Top             =   380
         Width           =   1275
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000016&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3360
         TabIndex        =   35
         Top             =   1125
         Width           =   1800
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "O/C:"
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
         TabIndex        =   34
         Top             =   1170
         Width           =   1275
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000016&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1155
         TabIndex        =   30
         Top             =   1125
         Width           =   1605
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente::"
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
         TabIndex        =   29
         Top             =   765
         Width           =   750
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000016&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1155
         TabIndex        =   28
         Top             =   720
         Width           =   4020
      End
      Begin VB.Label Label3 
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
         Height          =   330
         Left            =   1890
         TabIndex        =   27
         Top             =   380
         Width           =   1275
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
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
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1155
         TabIndex        =   26
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000016&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3255
         TabIndex        =   25
         Top             =   315
         Width           =   1065
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Venta/Rep:"
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
         Left            =   -840
         TabIndex        =   23
         Top             =   1170
         Width           =   1905
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
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
         Left            =   -840
         TabIndex        =   22
         Top             =   380
         Width           =   1905
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Entrega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1550
      Left            =   5670
      TabIndex        =   18
      Top             =   210
      Width           =   5055
      Begin VB.Label LENTREGA 
         BackColor       =   &H80000016&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   1155
         TabIndex        =   40
         Top             =   1125
         Width           =   3705
      End
      Begin VB.Label LENTREGA 
         BackColor       =   &H80000016&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   1155
         TabIndex        =   39
         Top             =   720
         Width           =   3705
      End
      Begin VB.Label LENTREGA 
         BackColor       =   &H80000016&
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
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   1155
         TabIndex        =   38
         Top             =   315
         Width           =   3705
      End
      Begin VB.Label Ltota 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Index           =   2
         Left            =   -840
         TabIndex        =   20
         Top             =   380
         Width           =   1905
      End
      Begin VB.Label Ltota 
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
         Index           =   5
         Left            =   -840
         TabIndex        =   19
         Top             =   1170
         Width           =   1905
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Novedades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   105
      TabIndex        =   7
      Top             =   6300
      Width           =   10620
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1175
         Left            =   1155
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   735
         Width           =   7155
      End
      Begin VB.CommandButton Command3 
         Caption         =   "O.K"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   745
         Left            =   8360
         Picture         =   "CONPEDID07.frx":1E10
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1140
         Width           =   750
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0C0&
         Height          =   425
         Left            =   0
         ScaleHeight     =   360
         ScaleWidth      =   10515
         TabIndex        =   9
         Top             =   315
         Width           =   10580
         Begin VB.Line Line2 
            X1              =   0
            X2              =   1155
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Label Label5 
            Caption         =   "Label5"
            Height          =   15
            Left            =   0
            TabIndex        =   11
            Top             =   420
            Width           =   1380
         End
         Begin VB.Label Label4 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   10
            Top             =   75
            Width           =   10200
         End
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   10665
         Y1              =   1900
         Y2              =   1900
      End
      Begin VB.Label Label12 
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
         Height          =   225
         Left            =   8400
         TabIndex        =   33
         Top             =   840
         Width           =   4000
      End
      Begin VB.Line Line7 
         X1              =   10580
         X2              =   10580
         Y1              =   315
         Y2              =   1900
      End
      Begin VB.Line Line6 
         X1              =   8295
         X2              =   10605
         Y1              =   1130
         Y2              =   1130
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Registrar Anotación"
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
         Left            =   9345
         TabIndex        =   32
         Top             =   1310
         Width           =   960
      End
      Begin VB.Line Line5 
         X1              =   8295
         X2              =   8295
         Y1              =   735
         Y2              =   1995
      End
      Begin VB.Label Label9 
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
         Height          =   1165
         Left            =   0
         TabIndex        =   12
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label10 
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
         Height          =   1170
         Left            =   8310
         TabIndex        =   13
         Top             =   735
         Width           =   2260
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0FFD0&
      Caption         =   "Anotaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4450
      Left            =   105
      TabIndex        =   5
      Top             =   1800
      Width           =   10620
      Begin VB.Frame Frame7 
         Caption         =   "Frame7"
         Height          =   3795
         Left            =   10290
         TabIndex        =   42
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame4"
         Height          =   4005
         Left            =   8085
         TabIndex        =   17
         Top             =   630
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "Frame3"
         Height          =   4110
         Left            =   1155
         TabIndex        =   16
         Top             =   630
         Width           =   15
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3660
         Left            =   0
         TabIndex        =   15
         Top             =   735
         Width           =   10580
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   10515
         TabIndex        =   6
         Top             =   315
         Width           =   10580
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha     Detalle                                                   Usuario"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   105
            TabIndex        =   31
            Top             =   105
            Width           =   10200
         End
      End
      Begin VB.Line Line4 
         X1              =   0
         X2              =   10605
         Y1              =   4410
         Y2              =   4410
      End
      Begin VB.Line Line1 
         X1              =   10580
         X2              =   10580
         Y1              =   320
         Y2              =   6225
      End
   End
   Begin VB.CommandButton Command26 
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
      Left            =   10970
      Picture         =   "CONPEDID07.frx":211A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Configurar Comprobante de Despacho"
      Top             =   2250
      Width           =   750
   End
   Begin VB.CommandButton Command9 
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
      Left            =   10970
      Picture         =   "CONPEDID07.frx":255C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Sale de la Aplicación"
      Top             =   300
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10340
      Picture         =   "CONPEDID07.frx":26A6
      Top             =   7710
      Width           =   2010
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -105
      TabIndex        =   4
      Top             =   2625
      Width           =   1695
   End
End
Attribute VB_Name = "CONPEDID07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command12_Click()

End Sub

Private Sub Command21_Click()
   Command21.Enabled = False
   Call COPYSTRU("INDIPEDI", "PEPENDES")
   RFF$ = RECFILT$("DESPAPED", 7, Chr$(1))
   NPXA& = 0: JJ& = 0: FIN& = Len(RFF$)
   For U& = 1 To FIN& Step 4
     Call TRACE(20, U&, FIN&)
     REDE& = CVS(Mid$(RFF$, U&, 4))
     XX$ = REGLEIDO$("DESPAPED", REDE&)
     NPX& = CVS(Left$(XX$, 4))
     If NPX& <> NPXA& Then
       ' AQUI REVISAR POR LOS MOVIMIENTOS DE MOVISTO SI ESE PEDIDO ESTA AUN PENDIENTE.
       ' SI NO ESTA PENDIENTE, PONER EN EL REGISTRO LA MARCA EN '2'
       ' E IR A NEXT U&
       NPXA& = NPX&
       NCLIE% = CVI(Mid$(XX$, 7, 2))
       NUOCO$ = TRIM$(Mid$(XX$, 9, 16))
       FEPE = CVI(Mid$(XX$, 5, 2))
       REPAR$ = "": If Mid$(XX$, 51, 1) = "R" Then REPAR$ = " (Reparación)"
       YY$ = MKS$(NPX&) + TRIM$(RASOCLI$(NCLIE%))
       If NUOCO$ <> "" Then YY$ = YY$ + " - " + NUOCO$
       YY$ = YY$ + " - " + FECHATEX$(FEPE) + REPAR$
       JJ& = JJ& + 1
       Call GRAREG("PEPENDES", YY$, JJ&)
     End If
   Next U&
   '
   Call SETULTREG("PEPENDES", JJ&)
   If JJ& < 1 Then
     Call COMUNI("No hay Pedidos\Pendientes de Despacho.")
     GoTo 99
   End If
   '
3  Call SELECHO("PEPENDES/Indice de Pedidos Pendientes de Despacho)")
   NPX& = Val(VALACT1$("PEPENDES"))
   If NPX& < 1 Then GoTo 99
   '
   Label7 = TRIM$(Str$(NPX&))
   '
99 Command21.Enabled = True
End Sub

Private Sub Command23_Click()
   Command23.Enabled = False
   Dim CIKU%(4196), NPKU&(4196), CAYADES(4196)
   '
   OPLIDE07.Show 1
   Call FRESHALL
   DoEvents
   If OPLIDESP.OPIMPRE = "" Then GoTo 99
   '
   Screen.MousePointer = 11
   NEPREPA% = 0
   If OPLIDESP.Check1(1).Value = 1 Then NEPREPA% = 1
   '
   Dim ACTI%(32767), STACT(32767), OFPEN(32767)
   Dim NROPEPE&(8192), REGPEPE&(8192), SALPEPE(8192)
   Dim ITPEDI$(), CODIT%(), CAPREPA(), CAYAEN(), NUPEDI&()
   MAXPED& = 4096
   ReDim ITPEDI$(MAXPED&), CODIT%(MAXPED&), CAPREPA(MAXPED&), CAYAEN(MAXPED&), NUPEDI&(MAXPED&)
   '
   DEPEX% = Int(Val(CONTROL$("", "DEPOEXPE")))
   DETER% = Int(Val(CONTROL$("", "DEPOTERM")))
   If DEPEX% < 1 Or DETER% < 1 Or DEPEX% = DETER% Then
     Call MENSERR(24, "Proceso Imposible.\  \Falta Configurar Depósitos\de Expedición\y Productos Terminados.")
     GoTo 99
   End If
   '
   RFF$ = RECFILT$("DESPAPED", 7, Chr$(1))
   FIN& = Len(RFF$)
   RFF1$ = RECFILT$("DESPAPED", 7, Chr$(11))
   FIN1& = Len(RFF1$)
   '
   Screen.MousePointer = 11
   HURECU% = 0
   For U& = 1 To FIN& Step 4
     Call TRACE(20, U&, FIN&)
     REDE& = CVS(Mid$(RFF$, U&, 4))
     ' AQUI VERIFICAR ESTADO DE PREPARACION DEL REGISTRO Y PASARLO A STATUS 3 SI ESTA CUMPLIDO
     Call VERIDESPRE(REDE&, HUBORECU%)
   Next U&
   Call CIERRARCH("DESPAPED")
   '
   JJ1& = 0
   RFF2$ = RECFILT$("DESPAPED", 7, Chr$(1))
   RFF1$ = RECFILT$("DESPAPED", 7, Chr$(11))
   RFF$ = RFF1$ + RFF2$
   FIN& = Len(RFF$)
   '
   For U& = 1 To FIN& Step 4
      Call TRACE(20, U&, FIN&)
      REDE& = CVS(Mid$(RFF$, U&, 4))
      X$ = REGLEIDO$("DESPAPED", REDE&)
      If Asc(Mid$(X$, 27, 1)) = 11 Then
         JJ1& = JJ1& + 1
         Call GRAREG("DESPAPEX", X$, JJ1&)
      End If
   Next U&
   Call SETULTREG("DESPAPEX", JJ1&)
   Call CIERRARCH("DESPAPEX")
   '
   ' ESTO QUE SIGUE ES PARA FORZAR LA REINDEXACION SI HUBO REGISTROS RECUPERADOS
   'If FIN& > 0 Then
   '  Call CIERRARCH("*.*")
   '  Call BAJALDISCO
   'End If
   '
   'Call FILTERFILE("DESPAPED", "DESPAPEX", 7, Chr$(11))
   Call FILESORT("DESPAPEX", "", 1, 1)
   RFF$ = RECFILT$("DESPAPEX", 7, Chr$(11))
   FIN& = Len(RFF$)
   JK& = 0: NPEDXA& = 0
   REBLA$ = REGBLAN$("CONPEDIP")
   For KU& = 1 To FIN& Step 4
     Call TRACE(20, KU&, FIN&)
     U& = CVS(Mid$(RFF$, KU&, 4))
     X$ = REGLEIDO$("DESPAPEX", U&)
     NPEDXX& = CVS(Left$(X$, 4))
     STATX% = Asc(Mid$(X$, 27, 1))
     REPARX% = 0: If Mid$(X$, 51, 1) = "R" Then REPARX% = 1
     '
     If REPARX% <> 1 Then
       NC% = CVI(Mid$(X$, 7, 2))
       CIXI% = CVI(Mid$(X$, 33, 2))
       CAPEDID = CVS(Mid$(X$, 75, 4))  ' CANTIDAD PEDIDA
       FEMIPED% = CVI(Mid$(X$, 5, 2))
       If CIXI% > 0 And CIXI% <= NUMAS% Then
         CAPRE = CVS(Mid$(X$, 79, 4))
         CAYEN = CVS(Mid$(X$, 95, 4))
         '
         ' ESTO ES PARA COMPENSAR LA APARICION DE PEDIDOS VIEJOS - EPB 28/02/06
         If CAPRE < CAYEN Then
           CAPRE = CAYEN
           NPEDX6& = CVS(Left$(X$, 4))
           RGH$ = RECFILT$("DESPAPED", 1, MKS$(NPEDX6&))
           For KKUU& = 1 To Len(RGH$) Step 4
             KKU6& = CVS(Mid$(RGH$, KKUU&, 4))
             XY6$ = REGLEIDO$("DESPAPED", KKU6&)
             Call REPLA(XY6$, Chr$(1), 27, 1)
             Call GRAREG("DESPAPED", XY6$, KKU6&)
           Next KKUU&
           Call CIERRARCH("DESPAPED")
         End If
         '
         CASAL = CAPEDID - CAYEN
26       If NEPREPA% > 0 Then CASAL = CAPEDID - CAPRE
         CAPEN = CASAL
         '
         If CASAL > 0.05 Then
           XX$ = REGLEIDO$("DESPAPEX", U&)
           If Mid$(XX$, 51, 1) <> "R" Then
             YY$ = REBLA$
             Call REPLA(YY$, Left$(XX$, 70), 1, 70)
             Call REPLA(YY$, MKS$(CASAL), 71, 4)
             Call REPLA(YY$, String$(54, 0), 75, 54)
             NPEDX& = CVS(Left$(YY$, 4))
             STAPEA% = 1
             If NPEDX& <> NPEDXA& Then STAPEA% = 1
             '
             If ACTI%(CIXI%) < 1 Then
               ACTI%(CIXI%) = 1
               STACT(CIXI%) = STOCKACT(CIXI%)
               '
               If NEPREPA% = 1 Then
                 ' MsgBox "OJO Instruccion Deshabilitada"
                 ' STACT(CIXI%) = STACT(CIXI%) - STOPREP(CIXI%)
                 ' ESTO QUE SIGUE ES PARA NO CONSIDERAR LO DEL DEPOSITO DE EXPEDICION
                 If DEPEX% > 0 Then STACT(CIXI%) = STACT(CIXI%) - STODEPOS(CIXI%, DEPEX%)
               End If
               '
               If STACT(CIXI%) < 0 Then STACT(CIXI%) = 0
               OFPEN(CIXI%) = STORFAB(CIXI%)
             End If
             '
             Call REPLA(YY$, MKS$(STACT(CIXI%)), 75, 4)
             Call REPLA(YY$, MKS$(OFPEN(CIXI%)), 79, 4)
             FALTA = CAPEN - STACT(CIXI%) - OFPEN(CIXI%)
             If FALTA < 0 Then FALTA = 0
             Call REPLA(YY$, MKS$(FALTA), 83, 4)
             JJ& = JJ& + 1
             Call GRAREG("CONPEDIP", YY$, JJ&)
             '
             STAPE% = 1
             If CAPEN > CVS(Mid$(YY$, 75, 4)) Then STAPE% = 2
             If FALTA > 0 Then STAPE% = 3
             If STAPE% > STAPEA% Then STAPEA% = STAPE%
             '
             For KK& = JJ& To 1 Step -1
               ZZ$ = REGLEIDO$("CONPEDIP", KK&)
               If CVS(Left$(ZZ$, 4)) <> NPEDX& Then GoTo 15
               Call REPLA(ZZ$, MKI$(STAPEA%), 127, 2)
               Call GRAREG("CONPEDIP", ZZ$, KK&)
             Next KK&
             '
15           NPEDXA& = NPEDX&
             If CAPEN <= STACT(CIXI%) Then
                 STACT(CIXI%) = STACT(CIXI%) - CAPEN
                 CAPEN = 0
               Else
                 CAPEN = CAPEN - STACT(CIXI%)
                 STACT(CIXI%) = 0
             End If
             '
             If CAPEN <= OFPEN(CIXI%) Then
                 OFPEN(CIXI%) = OFPEN(CIXI%) - CAPEN
                 CAPEN = 0
               Else
                 CAPEN = CAPEN - OFPEN(CIXI%)
                 OFPEN(CIXI%) = 0
             End If
             '
           End If
         End If
       End If
     End If
   Next KU&
   '
   Call SETULTREG("CONPEDIP", JJ&)
   Screen.MousePointer = 1
   If JJ& < 1 Then
     Call COMUNI("No hay Pedidos\Pendientes de Despacho.")
     GoTo 99
   End If
   '
   Call HEADERS("CONPEDIP", "")
   MAGRU% = Val(OPLIDESP.AGRUPA)
   If MAGRU% = 1 Then
       OPX% = Val(OPLIDESP.OPIMPRE) - 1
       If OPX% > 3 Then GoTo 99
       '
       If OPX% < 1 Then
            SITU$ = "Todos los Pendientes": If NEPREPA% = 1 Then SITU$ = SITU$ + " - Netos de Preparados"
            Call HEADERS("CONPEDIP", "Situación: " + SITU$)
          Else
            KK& = 0
            For U& = 1 To JJ&
              YY$ = REGLEIDO$("CONPEDIP", U&)
              If CVI(Mid$(YY$, 127, 2)) = OPX% Then
                KK& = KK& + 1
                Call GRAREG("CONPEDIP", YY$, KK&)
              End If
            Next U&
            Call SETULTREG("CONPEDIP", KK&)
            Call CIERRARCH("CONPEDIP")
            If OPX% = 1 Then
                 SITU$ = "Listos para Despachar": If NEPREPA% = 1 Then SITU$ = SITU$ + " - Netos de Preparados"
                 Call HEADERS("CONPEDIP", "Situación: " + SITU$)
              ElseIf OPX% = 2 Then
                 SITU$ = "En Preparación / Fabricación": If NEPREPA% = 1 Then SITU$ = SITU$ + " - Netos de Preparados"
                 Call HEADERS("CONPEDIP", "Situación: " + SITU$)
              ElseIf OPX% = 3 Then
                 SITU$ = "Para Fabricar": If NEPREPA% = 1 Then SITU$ = SITU$ + " - Netos de Preparados"
                 Call HEADERS("CONPEDIP", "Situación: " + SITU$)
            End If
       End If
       Call FINAL("*LIDESPE")
     ElseIf MAGRU% = 2 Then
       Call FILESORT("CONPEDIP", "", 7, 7, "ASC")
       Call FINAL("*LIDESQU")
   End If
   '
99 Command23.Enabled = True
   '
End Sub

Private Sub Command24_Click()
   Command24.Enabled = False
   If DERACCE%("DESPAMAT", "Despachar Pedidos de Clientes") > 1 Then
     Call DESPAMAT
   End If
   Command24.Enabled = True
End Sub

Private Sub Command26_Click()
   Command26.Enabled = False
   CFGDESPA.Show 1
   Command26.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   NPX& = Val(Label7)
   If NPX& > 0 Then
     If TRIM$(Text1) <> "" Then
       Call FRATEXTO(Text1, 56)
       For KKI% = 1 To CARETEX%
         TTXX$ = Space$(76)
         Call REPLA(TTXX$, RETEX$(KKI%), 11, 56)
         If KKI% = 1 Then
           Call REPLA(TTXX$, HOS$, 1, 8)
           Call REPLA(TTXX$, USERTER$, 69, 24)
         End If
         Call REGAPP("ANOTAPED", Left$(MKS$(NPX&) + TTXX$, 96))
       Next KKI%
       Call REGAPP("ANOTAPED", MKS$(NPX&) + String$(88, "="))
       Call CIERRARCH("ANOTAPED")
       Text1 = ""
       '
       List1.Clear
       RFF$ = RECFILT$("ANOTAPED", 1, MKS$(NPX&))
       '
       For U& = 1 To Len(RFF$) Step 4
         RENOTA& = CVS(Mid$(RFF$, U&, 4))
         XX$ = REGLEIDO$("ANOTAPED", RENOTA&)
         List1.AddItem Mid$(XX$, 5, 92)
       Next U&
       '
     End If
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command1.Enabled = False
'3  Call SELECHO("INDIPEDI/Indice General de Pedidos (Pendientes-Cumplidos-Anulados)")
'   NPX& = Val(VALACT1$("INDIPEDI"))
3  Call SELECHO("PEDENCA/Indice General de Pedidos (Pendientes-Cumplidos-Anulados)")
   NPX& = Val(VALACT1$("PEDENCA"))
   If NPX& < 1 Then GoTo 99
   '
   Label7 = TRIM$(Str$(NPX&))
   '
99 Command1.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   'Call ORFNUE(1)
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   NPX& = Val(Label7)
   '
   If NPX& > 0 Then
      TIDOCUM$ = "Pedido de Cliente"
      If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
        TIDOCUM$ = "Orden de Despacho"
      End If
      NUDOCU$ = TRIM$(Str$(NPX&))
      While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
      NUDOCU$ = NUDOCU$ + " "
      Screen.MousePointer = 11
      '
      For UI& = ULTREG&("PDOCLST") To 1 Step -1
        XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
        If InStr(XX$, TIDOCUM$) > 4 Then
          PPXX% = InStr(XX$, "Nro.")
          If PPXX% > 0 Then
            XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
          End If
          If InStr(XX$, NUDOCU$) > 4 Then
            DOCUNU& = CVS(Left$(XX$, 4))
            GoTo 5
          End If
        End If
      Next UI&
      '
      Screen.MousePointer = 1
      Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a este Pedido.\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
      On Error Resume Next
      OPPEDIDO.Command46.SetFocus
      On Error GoTo 0
      OPPEDIDO.Show 1
      GoTo 99
      '
      '16/03/07 (OT-07-0113)
5     Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
      '16/03/07 (FIN)
'5     For KKI% = 0 To 5
'        EDITFORM.Picture1(KKI%).Cls
'        EDITFORM.Picture1(KKI%).DrawWidth = 3
'        EDITFORM.Picture1(KKI%).DrawStyle = 0
'        EDITFORM.Picture1(KKI%).FillStyle = 1
'      Next KKI%
'      '
'      REDOCU$ = REGACT$("PDOCLST")
'      URE& = ULTREG&("PDOCSPL")
'      LI& = 0: LS& = URE&
'6     E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
'      L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
'      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'      If DCC& < DOCUNU& Then LI& = L&: GoTo 6
'      If DCC& > DOCUNU& Then LS& = L&: GoTo 6
'      GoTo 8
'      '
'7     Screen.MousePointer = 1
'      Call MENSERR(24, "Documento no Registrado")
'      GoTo 99
'      '
'8     TPSP$ = ""
'      Screen.MousePointer = 11
'      While L& > 1
'        DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
'        If DCC& < DOCUNU& Then GoTo 9
'        L& = L& - 1
'      Wend
'      '
'9     CAPAGINAS% = 0
'      For KKL& = L& To URE&
'        TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
'        If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
'        If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
'          TPSP$ = TPSP$ + Mid$(TBUF$, 5)
'        End If
'      Next KKL&
'      '
'95    SPOOLSTRING$ = TPSP$
'      For KKI% = 0 To 5
'        EDITFORM.Picture1(KKI%).Cls
'        EDITFORM.Picture1(KKI%).Top = 0
'        EDITFORM.Picture1(KKI%).Refresh
'        EDITFORM.Picture1(KKI%).Height = 7170
'        EDITFORM.Picture1(KKI%).Width = 11700
'        EDITFORM.Picture1(KKI%).DrawWidth = 3
'        EDITFORM.Picture1(KKI%).DrawStyle = 0
'        EDITFORM.Picture1(KKI%).FillStyle = 1
'      Next KKI%
'      ALTUPAGINA = 7170
'      TOPEPAGINA = 0
'      PAGINACTIVA% = 0
'      CAPAGINAS% = 0
'      '
'10    PPXX% = InStr(TPSP$, Chr$(255))
'      If PPXX% > 0 Then
'        COPRIAT$ = Left$(TPSP$, PPXX% - 1)
'        TPSP$ = Mid$(TPSP$, PPXX% + 1)
'        Call TRACE(20, FIN& - Len(TPSP$), FIN&)
'        '
'        If Len(COPRIAT$) > 0 Then
'          COPRI% = Asc(Left$(COPRIAT$, 1))
'          ATRIP$ = Mid$(COPRIAT$, 2)
'          Call SHOWCOMMAND(COPRI%, ATRIP$)
'        End If
'        GoTo 10
'        '
'      End If
'      '
'      Screen.MousePointer = 1
'      PAGINACTIVA% = 0
'      For KKI% = 0 To 5
'         EDITFORM.Picture1(KKI%).Visible = False
'      Next KKI%
'      EDITFORM.Picture1(PAGINACTIVA%).Visible = True
'      EDITFORM.Picture1(PAGINACTIVA%).Refresh
'      EDITFORM.VScroll1.Min = 0
'      EDITFORM.VScroll1.Value = 0
'      EDITFORM.VScroll1.Max = 1
'      If ALTUPAGINA > EDITFORM.Frame1.Height Then
'         EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'      End If
'      EDITFORM.Command3.Enabled = False
'      EDITFORM.Command5.Enabled = False
'      EDITFORM.Show 1
    End If
    '
99  Command6.Enabled = True
    '
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   PREPAP07.Show 1
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   '
   Screen.MousePointer = 11
   FIN& = ULTREG("PEDDETA")
   NPEDA& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     PDC$ = REGLEIDO$("PEDDETA", U&)
     FEMIPED% = CVI(Mid$(PDC$, 5, 2))
     If FEMIPED% >= FECHANUM("01/11/05") Then
       If Asc(Mid$(PDC$, 27, 1)) <> 9 Then
         If Mid$(PDC$, 51, 1) <> "R" Then
           NPED& = CVS(PDC$)
           CAN = CVS(Mid$(PDC$, 79, 4))
           ABUS1$ = Left$(PDC$, 26)
           ABUS2$ = Mid$(PDC$, 28, 43)
           '
           RFF$ = RECFILT$("DESPAPED", 1, MKS$(NPED&))
           For KU& = 1 To Len(RFF$) Step 4
             REDE& = CVS(Mid$(RFF$, KU&))
             XX$ = REGLEIDO$("DESPAPED", REDE&)
             If Left$(XX$, 26) = ABUS1$ Then
               If Mid$(XX$, 28, 43) = ABUS2$ Then
                 If CVS(Mid$(XX$, 75, 4)) = CAN Then
                   GoTo 90
                 End If
               End If
             End If
           Next KU&
           If NPED& <> NPEDA& Then
             Call MENSERR(24, "Pedido " & NPED& & " Inconsistente en DESPAPED")
             NPEDA& = NPED&
           End If
         End If
       End If
     End If
90 Next U&
   '
   Screen.MousePointer = 1
   Call COMUNI("Fin del Control")
   '
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub Form_Load()
    Call CENTRAFORM(Me)
    '
    DoEvents
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Screen.MousePointer = 1
    HOII% = HOY(HOS$)
    Label9 = " " + HOS$
    Label12 = Left$(USERNAME, 17)
    '
End Sub


Private Sub Label18_DblClick()
   NPX& = Val(Label7)
   If NPX& > 0 Then
     If DERACCE%("RESITPED", "Restaurar Situación de Preparacion y Despacho") = 2 Then
       If UCase$(Label14) = "VENTA" Then
         If COMALTER%("Realmente Desea Restaurar este Pedido\para Preparación y Despacho,\Revirtiendo TODOS los Movimientos ?\\No, Conservar\Si, Restaurar") = 2 Then
           '
           Screen.MousePointer = 11
           RFF$ = RECFILT$("DESPAPED", 1, MKS$(NPX&))
           If RFF$ = "" Then
             Call PREPAPED.RECUDESPA(NPX&)
             RFF$ = RECFILT$("DESPAPED", 1, MKS$(NPX&))
           End If
           '
           Screen.MousePointer = 11
           For U& = 1 To Len(RFF$) Step 4
             REPED& = CVS(Mid$(RFF$, U&, 4))
             XX$ = REGLEIDO$("DESPAPED", REPED&)
             Call REPLA(XX$, Chr$(1), 27, 1)
             Call GRAREG("DESPAPED", XX$, REPED&)
           Next U&
           Call CIERRARCH("DESPAPED")
           '
           REBLA$ = REGBLAN$("MOVISTO")
           ABUSBLA$ = "VP." + TRIM$(Str$(NPX&))
           ABUSBLB$ = "DM." + TRIM$(Str$(NPX&))
           FIN& = ULTREG&("MOVISTO")
           For KKU& = 1 To FIN
             Call TRACE(20, KKU&, FIN&)
             XX$ = REGLEIDO$("MOVISTO", KKU&)
             ABUSX$ = TRIM$(Mid$(XX$, 23, 10))
             If ABUSX$ = ABUSBLA$ Or ABUSX$ = ABUSBLB$ Then
               Call GRAREG("MOVISTO", REBLA$, KKU&)
               GoTo 19
             End If
             '
             ' ESTO ES PARA RESOLVER EL TEMA DEL PEDIDO 13260 DE DOSIVAC
             If ((Mid$(XX$, 21, 2) = "PP" Or Mid$(XX$, 21, 2) = "DM") And CVS(Mid$(XX$, 112, 4)) = NPX&) Then
               Call GRAREG("MOVISTO", REBLA$, KKU&)
               GoTo 19
             End If
19         Next KKU&
           Call CIERRARCH("MOVISTO")
           '
           Screen.MousePointer = 11
           TTXYZ$ = Space$(76)
           HOII% = HOY(HOS$)
           Call REPLA(TTXYZ$, HOS$, 1, 8)
           Call REPLA(TTXYZ$, "Restauracion de Preparados y Despacho ", 11, 56)
           Call REPLA(TTXYZ$, USERTER$, 69, 24)
           Call REGAPP("ANOTAPED", MKS$(NPXX&) + TTXYZ$)
           Call REGAPP("ANOTAPED", MKS$(NPXX&) + Space$(88))
           Call REGAPP("ANOTAPED", MKS$(NPXX&) + String$(88, "="))
           Call CIERRARCH("ANOTAPED")
           Call FRESHECHO("ANOTAPED")
           Call CIERRARCH("*.*")
           Call BAJALDISCO
           Screen.MousePointer = 1
           Call COMUNI("Pedido " + TRIM$(Str$(NPX&)) + " Restaurado.\  \Recuerde que Debe Volver a Preparar\y Despachar los Materiales\Componentes de Este Pedido.")
         End If
       End If
     End If
   End If
End Sub

Private Sub Label7_Change()
   List1.Clear
   If Val(Label7) < 1 Then Exit Sub
   '
   NPX& = Val(Label7)
   RFF$ = RECFILT$("PEDENCA", 1, MKS$(NPX&))
   If Len(RFF$) >= 4 Then
     REPED& = CVS(Left$(RFF$, 4))
     XX$ = REGLEIDO$("PEDENCA", REPED&)
     FEX = CVI(Mid$(XX$, 53, 2))
     Label8 = FECHATEX$(FEX)
     NCLIE% = CVI(Mid$(XX$, 55, 2))
     Label11 = RASOCLI$(NCLIE%)
     Label16 = Mid$(XX$, 57, 16)
     Label18 = TRIM$(Str$(Asc(Mid$(XX$, 267, 1))))
     '
     LENTREGA(0) = Mid$(XX$, 73, 32)
     LENTREGA(1) = Mid$(XX$, 105, 32)
     LENTREGA(2) = Mid$(XX$, 137, 32)
   End If
   '
   Label14.ForeColor = RGB(0, 0, 255)
   Label14 = "Venta": REPARX% = 0
   RFQ$ = RECFILT$("PEDDETA", 1, MKS$(NPX&))
   For UKKK& = 1 To Len(RFQ$) Step 4
     UU& = CVS(Mid$(RFQ$, UKKK&, 4))
     XX$ = REGLEIDO$("PEDDETA", UU&)
     If CVS(Left$(XX$, 4)) = NPX& Then
       If UCase$(Mid$(XX$, 51, 1)) = "R" Then
         Label14.ForeColor = RGB(255, 0, 0)
         Label14 = "Reparación"
         REPARX% = 1
       End If
       GoTo 20
     End If
   Next UKKK&
   '
'20 RFF$ = RECFILT$("ANOTAPED", 1, MKS$(NPX&))
'   For U& = 1 To Len(RFF$) Step 4
'     RENOTA& = CVS(Mid$(RFF$, U&, 4))
'     XX$ = REGLEIDO$("ANOTAPED", RENOTA&)
'     List1.AddItem Mid$(XX$, 5, 92)
'   Next U&
20 Call APERBASDAT("PEDCLI")
   SQLX$ = "SELECT ANOTAPED FROM PEDENCA"
   SQLX$ = SQLX$ + " WHERE NroPed = " & NPX& & ""
   Dim PEDDDENCA As New ADODB.Recordset
   Set PEDDDENCA = New ADODB.Recordset
   PEDDDENCA.OPEN FILTSQL$ (SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
   With PEDDDENCA
    On Error Resume Next
    .MoveFirst
    If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
         XX$ = !ANOTAPED
         List1.AddItem Mid$(XX$, 5, 92)
        Loop
    End If
    On Error GoTo 0
   End With
   PEDDDENCA.Close
   Set PEDDDENCA = Nothing
   ' 21/01/2006: al consultar un pedido que recupere
   '             la consistencia de preparacion y despacho
   If REPARX% <> 1 Then ' NO ES UNA REPARACIóN
     Screen.MousePointer = 11
     HUBOREC% = 0
     RFF$ = RECFILT$("DESPAPED", 1, MKS$(NPX&))
     For U& = 1 To Len(RFF$) Step 4
       REDES& = CVS(Mid$(RFF$, U&, 4))
       Call VERIDESPRE(REDES&, HUBORECU%)
       If HUBORECU% > 0 Then HUBOREC% = 1
     Next U&
     Call CIERRARCH("DESPAPED")
     Screen.MousePointer = 1
     If HUBOREC% > 0 Then
       Call COMUNI("Se Han Recuperado Acumulados\de Preparación y Despacho.")
     End If
   End If
   '
End Sub

Sub DESPAMAT()
   '
   Static DEPEX%
   '
   Screen.MousePointer = 11
   If DEPEX% < 1 Then
     DEPEX% = Int(Val(CONTROL$("", "DEPOEXPE")))
     If DEPEX% < 1 Then
       Call MENSERR(24, "Proceso Imposible.\  \Falta Configurar Depósitos\de Expedición\y Productos Terminados.")
       GoTo 99
     End If
   End If
   '
   Screen.MousePointer = 11
   Call COPYSTRU("INDIPEDI", "PEPENDES")
   RFF$ = RECFILT$("DESPAPED", 7, Chr$(1))
   FIN& = Len(RFF$)
   '
   Screen.MousePointer = 11
   For U& = 1 To FIN& Step 4
     Call TRACE(20, U&, FIN&)
     REDE& = CVS(Mid$(RFF$, U&, 4))
     ' AQUI VERIFICAR ESTADO DE PREPARACION DEL REGISTRO Y PASARLO A STATUS 3 SI ESTA CUMPLIDO
     'Call VERIDESPRE(REDE&, HUBORECU%)
   Next U&
   '
   RFF$ = RECFILT$("DESPAPED", 7, Chr$(11))
   NPXA& = 0: JJ& = 0: FIN& = Len(RFF$)
   For U& = 1 To FIN& Step 4
     REDE& = CVS(Mid$(RFF$, U&, 4))
     XX$ = REGLEIDO$("DESPAPED", REDE&)
     NPX& = CVS(Left$(XX$, 4))
     CAPEDIT = CVS(Mid$(XX$, 75, 4))
     ACUPRE = CVS(Mid$(XX$, 79, 4))
     ACUDES = CVS(Mid$(XX$, 95, 4))
     SALPENDES = ACUPRE - ACUDES
     If CAPEDIT < ACUPRE Then
       SALPENDES = CAPEDIT - ACUDES
     End If
     '
     If SALPENDES >= 0.05 Then
       If NPX& <> NPXA& Then
         NPXA& = NPX&
         NCLIE% = CVI(Mid$(XX$, 7, 2))
         NUOCO$ = TRIM$(Mid$(XX$, 9, 16))
         FEPE = CVI(Mid$(XX$, 5, 2))
         'REPAR$ = "": If Mid$(XX$, 51, 1) = "R" Then REPAR$ = " (Reparación)"
         If Mid$(XX$, 51, 1) <> "R" Then
           YY$ = MKS$(NPX&) + TRIM$(RASOCLI$(NCLIE%))
           If NUOCO$ <> "" Then YY$ = YY$ + " - " + NUOCO$
           YY$ = YY$ + " - " + FECHATEX$(FEPE) + REPAR$
           JJ& = JJ& + 1
           Call GRAREG("PEPENDES", YY$, JJ&)
         End If
       End If
     End If
29 Next U&
   '
   Call SETULTREG("PEPENDES", JJ&)
   Call FILESORT("PEPENDES", "", 1, 1)
   Call FRESHECHO("PEPENDES")
   If JJ& < 1 Then
     Call COMUNI("No hay Pedidos Preparados\Pendientes de Despacho.")
     GoTo 99
   End If
   '
3  Screen.MousePointer = 1
   Call SELECHO("PEPENDES/Indice de Pedidos Pendientes de Despacho")
   NPX& = Val(VALACT1$("PEPENDES"))
   If NPX& < 1 Then GoTo 99
   '
5  ' Label7 = TRIM$(Str$(NPX&))
   '
   FORIPRE$ = TRIM$(CONTROL$("", "CDESPACH"))
   If FORIPRE$ <> "" Then
     LU$ = LUDAT$
     If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
       RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
       Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito en Consignación:\  \'" + RMCC$ + "'.")
       Exit Sub
     End If
   End If
   '
   DEPOCON% = DEPOEXPE%
   DEPORIG% = DEPOCON%
   DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
   If DEPOCON% < 1 Or TRIM$(DEPOSI$) = "" Then
     Call MENSERR(24, "Falta Configurar Depósito de Expedición.")
     Exit Sub
   End If
   '
   Call CIERRARCH("*.*")
   Call BLANARCH("!BOLREDET")
   '
   Screen.MousePointer = 11
   NPED& = NPX&
   REBLA$ = REGBLAN$("DESPAPED")
   RFF$ = RECFILT$("DESPAPED", 1, MKS$(NPED&))
   '
   For U& = 1 To Len(RFF$) Step 4
     REDE& = CVS(Mid$(RFF$, U&, 4))
     XX$ = REGLEIDO$("DESPAPED", REDE&)
     STATU% = Asc(Mid$(XX$, 27, 1))
     If STATU% < 2 Or STATU = 11 Then
       If Mid$(XX$, 51, 1) <> "R" Then
         CI% = CVI(Mid$(XX$, 33, 2))
         NCLIE% = CVI(Mid$(XX$, 7, 2))
         ACUPREPA = CVS(Mid$(XX$, 79, 4))
           If CVS(Mid$(XX$, 75, 4)) < ACUPREPA Then ACUPREPA = CVS(Mid$(XX$, 75, 4))
         CAPENE# = ACUPREPA - CVS(Mid$(XX$, 95, 4))
         FENTRESO% = CVI(Mid$(XX$, 25, 2))
         '
         If CAPENE# >= 0.05 Then
           Y$ = REGBLAN$("!BOLREDET")
           Call REPLA(Y$, MKI$(CI%), 83, 2)
           Call REPLA(Y$, UNIMED$(CI%), 85, 4)
           Call REPLA(Y$, MKD$(CAPENE#), 89, 8)
           Call REPLA(Y$, MKS$(NPED&), 151, 4)
           Call REPLA(Y$, MKI$(FENTRESO%), 155, 2)
           Call REPLA(Y$, MKD$(CAPENE#), 163, 8)
           Call REPLA(Y$, MKS$(REDE&), 189, 4)
           Call REGAPP("!BOLREDET", Y$)
         End If
       End If
     End If
   Next U&
   '
12 VTB% = VENTABU%("DESPAPED/NC/TITUPAN=Despacho Materiales Vendidos - " + TRIM$(RASOCLI$(NCLIE%)))
   If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
     Exit Sub
   End If
   '
   For YX& = 1 To ULTREG&("!BOLREDET")
     YY$ = REGLEIDO$("!BOLREDET", YX&)
     CI% = CVI(Mid$(YY$, 83, 2))
     CAN# = CVD(Mid$(YY$, 163, 8))
     If CI% > 0 Then
       If CI% <= NUMAS% Then
         If CAN# > 0.005 Then
           GoTo 13
         End If
       End If
     End If
   Next YX&
   Call MENSERR(24, "Despacho Cancelado !!!\Faltan Cantidades o Códigos.")
   Exit Sub
   '
13 QUEDAPEN% = 0
   For YX& = 1 To ULTREG&("!BOLREDET")
     YY$ = REGLEIDO$("!BOLREDET", YX&)
     CI% = CVI(Mid$(YY$, 83, 2))
     CAPREP# = CVD(Mid$(YY$, 163, 8))
     MAXX# = CVD(Mid$(YY$, 89, 8))
     If CAPREP# > MAXX# + 0.05 Then
       Call MENSERR(24, "Imposible Despachar Cantidad Mayor\a Preparada Pendiente de Despacho.")
       GoTo 12
     End If
     If CAPREP# < MAXX# - 0.05 Then QUEDAPEN% = 1
   Next YX&
   '
15 CONTA% = ULTREG&("!BOLREDET")
   If CONTA% > 0 Then
     Call CIERRARCH("MOVISTO")
     Call BLOQARCH("MOVISTO")
     Call BLOQARCH("ANOTAPED")
     Call BLOQARCH("ANOTAREP")
     Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
     '
     TTXXX$ = "Materiales Vendidos"
     If REPARX% = 1 Then
        TTXXX$ = "Reparación"
        NREP& = 0
        For KU& = ULTREG&("PRESREPA") To 1 Step -1
          XYX$ = REGLEIDO$("PRESREPA", KU&)
          If CVS(Mid$(XYX$, 125, 4)) = NPED& Then
            NREP& = CVS(Left$(XYX$, 4))
            GoTo 16
          End If
        Next KU&
     End If
     '
16   TTXYZ$ = Space$(76)
     HOII% = HOY(HOS$)
     Call REPLA(TTXYZ$, HOS$, 1, 8)
     Call REPLA(TTXYZ$, "Despacho de " + TTXXX$ + ":", 11, 56)
     Call REPLA(TTXYZ$, USERTER$, 69, 24)
     Call REGAPP("ANOTAPED", MKS$(NPED&) + TTXYZ$)
     Call REGAPP("ANOTAPED", MKS$(NPED&) + Space$(88))
     If REPARX% = 1 Then
       Call REGAPP("ANOTAREP", MKS$(NREP&) + TTXYZ$)
       Call REGAPP("ANOTAREP", MKS$(NREP&) + Space$(88))
     End If
     '
     AINDI% = 0
     For YX& = 1 To ULTREG&("!BOLREDET")
       YY$ = REGLEIDO$("!BOLREDET", YX&)
       CI% = CVI(Mid$(YY$, 83, 2))
       CAN# = CVD(Mid$(YY$, 163, 8))
       LOTE$ = Mid$(YY$, 69, 12)
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           If CAN# > 0.005 Then
             NUOC& = CVS(Mid$(YY$, 151, 4))
             REDE& = CVS(Mid$(YY$, 189, 4))
             ZY$ = REGLEIDO$("DESPAPED", REDE&)
             CAVERI$ = Mid$(ZY$, 113, 10)
             DOSEC$ = CAVERI$
             CMO$ = "DM"
             If REPARX% = 1 Then CMO$ = "*D"
             REMI$ = "DM." + TRIM$(Str$(NPED&))
             FF% = HOY(HOS$)
             '
             CANTI = (-1) * CAN#
             'Call MOVISTO(FF%, CI%, LOTE$, CMO$, REMI$, DOSEC$, RASOCLI$(NCLIE%), 0, "$", NC1%, DEPORIG%, CANTI)
             '
             If CAN# >= 0.05 Then
               ACUDESPA = CVS(Mid$(ZY$, 95, 4)) + CAN#
               Call REPLA(ZY$, MKS$(ACUDESPA), 95, 4)
               Call REPLA(ZY$, Chr$(1), 27, 1) ' MARCA QUE ESTA DESPACHADO
               Call GRAREG("DESPAPED", ZY$, REDE&)
             End If
             '
             TTXX$ = Space$(76)
             CIXI% = CI%
             CAENT = Abs(CANTI)
             Call REPLA(TTXX$, "- " + CODEXT$(CIXI%), 11, 18)
             Call REPLA(TTXX$, DESCRIT0$(CIXI%), 29, 38)
             Call REPLA(TTXX$, CSTRING$(MKS$(CAENT), 3, 10, 1, 2), 57, 10)
             Call REGAPP("ANOTAPED", Left$(MKS$(NPED&) + TTXX$, 96))
             If REPARX% = 1 Then
               Call REGAPP("ANOTAREP", MKS$(NREP&) + TTXX$)
             End If
             '
           End If
         End If
       End If
     Next YX&
     Call CIERRARCH("MOVISTO")
     '
     Call REGAPP("ANOTAPED", Left$(MKS$(NPED&) + Space$(10) + String$(56, "-"), 96))
     Call REGAPP("ANOTAPED", MKS$(NPED&) + String$(88, "="))
     If REPARX% = 1 Then
       Call REGAPP("ANOTAREP", Left$(MKS$(NREP&) + Space$(10) + String$(56, "-"), 96))
       Call REGAPP("ANOTAREP", MKS$(NREP&) + String$(88, "="))
     End If
     '
     Call CIERRARCH("ANOTAPED")
     Call CIERRARCH("ANOTAREP")
     Call CIERRARCH("MOVIREPA")
     Call CIERRARCH("DESPAPED")
     '
     Label7 = ""
     Label7 = TRIM$(Str$(NPED&))
     '
     Screen.MousePointer = 1
     '
     If TRIM$(FORIPRE$) <> "" Then
       '
       FEX = FF%
       FECHDOC$ = FECHATEX$(FEX)
       NUMEDOC$ = TRIM$(REMI$)
       TIPODOC$ = "Despacho de Materiales"
       DESTIDOC% = NCLIE%
       '
       CODPARAM$(1) = "UBI-MPRI": DOCPARAM$(1) = TRIM$(Label14)
       CAPARDOC% = 1
       '
       CODTABU1$(1) = "COD-MAT"
       CODTABU1$(2) = "DES-MAT"
       CODTABU1$(3) = "CANTIDAD"
       CODTABU1$(4) = "CAKILOS"
       CODTABU1$(5) = "UNIMED"
       CACOLTAB1% = 5
       '
       CAITAB1% = 0
       '
       For YX& = 1 To ULTREG&("!BOLREDET")
         YY$ = REGLEIDO$("!BOLREDET", YX&)
         CI% = CVI(Mid$(YY$, 83, 2))
         CAN# = CVD(Mid$(YY$, 89, 8))
         If CI% > 0 Then
           If CI% <= NUMAS% Then
             If CAN# > 0.005 Then
               UNID$ = TRIM$(Mid$(YY$, 85, 4))
               If UNID$ = "" Then UNID$ = UNIMED$(CI%)
               CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
               KILO$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 12, 3, 2)
               NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
               NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
               '
               CAITAB1% = CAITAB1% + 1
               TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
               TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
               TETABU1$(3, CAITAB1%) = CANS$
               TETABU1$(4, CAITAB1%) = KILO$
               TETABU1$(5, CAITAB1%) = UNID$
               '
             End If
           End If
         End If
       Next YX&
       '
       CANCELPRINT% = 0
       If CONTROL$("CDESPACH", "SUPRINT") = "SI" Then CANCELPRINT% = 1
       Call PRINTDOC("CDESPACH")   ' Comprobante de Despacho
       '
     End If
   End If
   '
   If QUEDAPEN% < 1 Then
     Call FILTERDELETE("PEPENDES", 1, MKS$(NPED&))
     Call FRESHECHO("PEPENDES")
   End If
   GoTo 3
   '
99 Call CIERRARCH("*.*")
   '
End Sub


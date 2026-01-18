VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form OPCICOM07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ORDENES DE COMPRA - PROVEEDORES EXTERNOS"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   9285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   9120
      TabIndex        =   57
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   9000
      Sorted          =   -1  'True
      TabIndex        =   56
      Top             =   3960
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Consignación"
      Height          =   255
      Left            =   2880
      TabIndex        =   55
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   53
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   4740
      Left            =   8400
      ScaleHeight     =   4710
      ScaleWidth      =   1035
      TabIndex        =   45
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   58
         Top             =   3840
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   59
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Excel 5.0;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   -360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3120
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Co-Pr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "OPCICOM07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Carga de Articulo por Proveedor"
         Top             =   2955
         Width           =   650
      End
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   480
         Sorted          =   -1  'True
         TabIndex        =   50
         Top             =   3720
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Direct"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "OPCICOM07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Acceso Directo a Otros Procesos"
         Top             =   885
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "OPCICOM07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   195
         Width           =   650
      End
      Begin VB.CommandButton Command18 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "OPCICOM07.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   1575
         Width           =   650
      End
      Begin VB.CommandButton Command19 
         Caption         =   "N.Orig"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "OPCICOM07.frx":09E0
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Códigos Equivalentes de Proveedor"
         Top             =   2265
         Width           =   650
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   600
         Top             =   3240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "GESTION DE COMPRAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4095
      Begin VB.Frame Frame7 
         BackColor       =   &H00CDDADA&
         Caption         =   "MATERIALES CONSIGNADOS"
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
         Left            =   240
         TabIndex        =   9
         Top             =   2880
         Width           =   3585
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   270
            ScaleHeight     =   615
            ScaleWidth      =   3015
            TabIndex        =   10
            Top             =   315
            Width           =   3050
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
               Height          =   630
               Left            =   2460
               Picture         =   "OPCICOM07.frx":0B2A
               Style           =   1  'Graphical
               TabIndex        =   51
               ToolTipText     =   "Listado de Materiales Pendientes de Consignacion"
               Top             =   0
               Width           =   560
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
               Height          =   630
               Left            =   0
               Picture         =   "OPCICOM07.frx":0E34
               Style           =   1  'Graphical
               TabIndex        =   3
               ToolTipText     =   "Entrega en Consignación "
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command15 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   1900
               Picture         =   "OPCICOM07.frx":113E
               Style           =   1  'Graphical
               TabIndex        =   4
               ToolTipText     =   "Consumo Materiales Consignados /Ajustes de Materiales Consignados"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Consignación    de Materiales"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   750
               TabIndex        =   11
               Top             =   105
               Width           =   1280
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00CDDADA&
         Caption         =   "ORDENES DE COMPRAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2475
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   3585
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   915
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   12
            Top             =   1650
            Width           =   2325
            Begin VB.CommandButton Command10 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   1740
               Picture         =   "OPCICOM07.frx":1288
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "Revertir Anulaciones de Ordenes de Compra"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command14 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "OPCICOM07.frx":13D2
               Style           =   1  'Graphical
               TabIndex        =   2
               ToolTipText     =   "Anulación de Orden de Compra"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Anulación de O/Compras"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   700
               TabIndex        =   13
               Top             =   100
               Width           =   1155
            End
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   715
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   42
            Top             =   1250
            Width           =   2325
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
               Height          =   630
               Left            =   0
               Picture         =   "OPCICOM07.frx":16DC
               Style           =   1  'Graphical
               TabIndex        =   43
               ToolTipText     =   "Marcar Items como Cumplidos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Items Cumplidos"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   44
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture14 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   515
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   14
            Top             =   850
            Width           =   2325
            Begin VB.CommandButton Command22 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "OPCICOM07.frx":19E6
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Ver y Re-Imprimir"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   15
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture15 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   315
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   16
            Top             =   450
            Width           =   2325
            Begin VB.CommandButton Command13 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "OPCICOM07.frx":1CF0
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Modificar Orden de Compra no Anulada"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label16 
               BackStyle       =   0  'Transparent
               Caption         =   "Modificar O/C"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   17
               Top             =   105
               Width           =   1590
            End
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Pedidos de Suministro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   840
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   3165
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   29
         Top             =   315
         Width           =   2325
         Begin VB.CommandButton Command6 
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
            Height          =   600
            Left            =   -20
            Picture         =   "OPCICOM07.frx":1E3A
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Generar un Pedido de Suministro"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar Pedido de Suministro"
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
            Left            =   735
            TabIndex        =   31
            Top             =   45
            Width           =   1425
         End
      End
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   26
         Top             =   795
         Width           =   2325
         Begin VB.CommandButton Command17 
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
            Height          =   600
            Left            =   0
            Picture         =   "OPCICOM07.frx":2144
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Adjudicación de Pedido de Suministro a Proveedor"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Asignacion de Proveedor"
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
            Left            =   840
            TabIndex        =   28
            Top             =   90
            Width           =   1170
         End
      End
      Begin VB.PictureBox Picture13 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   735
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   23
         Top             =   1260
         Width           =   2325
         Begin VB.CommandButton Command21 
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
            Height          =   600
            Left            =   0
            Picture         =   "OPCICOM07.frx":244E
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Pedidos de Suministro Pendientes"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "PS's Pendientes"
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
            Left            =   735
            TabIndex        =   25
            Top             =   105
            Width           =   1485
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "OPCIONES DE CONTROL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      Left            =   4440
      TabIndex        =   7
      Top             =   120
      Width           =   3885
      Begin VB.Frame Frame3 
         BackColor       =   &H00CDDADA&
         Caption         =   "RESUMENES DE COMPRAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   3375
         Begin VB.PictureBox Picture12 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   735
            ScaleHeight     =   615
            ScaleWidth      =   2190
            TabIndex        =   36
            Top             =   1020
            Width           =   2220
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
               Height          =   630
               Left            =   -20
               Picture         =   "OPCICOM07.frx":2758
               Style           =   1  'Graphical
               TabIndex        =   37
               ToolTipText     =   "Estadísticas de Compras"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Resúmenes Estadísiticos"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   38
               Top             =   120
               Width           =   1170
            End
         End
         Begin VB.PictureBox Picture10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   420
            ScaleHeight     =   615
            ScaleWidth      =   2190
            TabIndex        =   33
            Top             =   540
            Width           =   2220
            Begin VB.CommandButton Command12 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   -20
               Picture         =   "OPCICOM07.frx":2A62
               Style           =   1  'Graphical
               TabIndex        =   34
               ToolTipText     =   "Resúmenes de Ultimas Compras"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "Detalle de         Ultimas Compras"
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
               Height          =   540
               Left            =   750
               TabIndex        =   35
               Top             =   30
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00CDDADA&
         Caption         =   "COMPLEMENTOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   240
         TabIndex        =   8
         Top             =   2280
         Width           =   3375
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   19
            Top             =   1560
            Visible         =   0   'False
            Width           =   2220
            Begin VB.CommandButton Command23 
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
               Height          =   600
               Left            =   0
               Picture         =   "OPCICOM07.frx":2D6C
               Style           =   1  'Graphical
               TabIndex        =   20
               ToolTipText     =   "Análisis Estadístico de Consumo"
               Top             =   0
               Visible         =   0   'False
               Width           =   560
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Análisis Estadís- tico de Consumo"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   21
               Top             =   105
               Width           =   1695
            End
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   600
            ScaleHeight     =   615
            ScaleWidth      =   2190
            TabIndex        =   39
            Top             =   600
            Width           =   2220
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
               Height          =   630
               Left            =   0
               Picture         =   "OPCICOM07.frx":3076
               Style           =   1  'Graphical
               TabIndex        =   40
               ToolTipText     =   "Control de Costos de Reposición"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Costo Reposición"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   700
               TabIndex        =   41
               Top             =   240
               Width           =   1695
            End
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "OPCICOM07.frx":31C0
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "OPCICOM07.frx":33F4
      TabIndex        =   54
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu RecibeArchStockMYM 
         Caption         =   "Recibir Archivo de Stocks Máximos y Mínimos"
      End
      Begin VB.Menu RecibirExcepcion 
         Caption         =   "Recibir Archivo de Excepciones"
      End
      Begin VB.Menu RECIBOPEDICOMPRA 
         Caption         =   "Recibir Archivo Generación de Pedidos Asociado a Compra"
      End
      Begin VB.Menu mnuAjusCons 
         Caption         =   "Ajuste de Consignaciones"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "OPCICOM07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub AJUSARMADOR()
   '
   If CantRegistros("RENCONSIG") < 1 Then
      Call MENSERR(24, "Proceso Imposible.\   \Debe Previamente Realizar una Recepcion de Armadores.")
      Exit Sub
   End If
   '
   'PRESENTA OBJETO PARA SELECCIONAR EL CODIGO Y EL PROVEEDOR
   VDEF$ = ""
10 OBX$ = OBJPLA$("AJUARMAD", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   'TOMO LOS VALORES RECIBIDOS
   NPRO% = CVI(Left$(OBX$, 2))
   Ci% = CVI(Mid$(OBX$, 3, 2))
   '
   Call ABRESTOCKS
   Call PoneEnCeroNull("MACONSIG", "Cant_Devu", REGAF&, "NOMESS")
   '
   'CALCULO LA CANTIDAD ENTREGADA MENOS LA RENDIDA
   CONDI$ = "NPRO_OCOM = " & NPRO% & " AND COD_INTE =  " & Ci%
   CANTIENTRE = XVALO(VALOBADA("MACONSIG", "SUM(CANT_ENTREG)", CONDI$, "NOMESS"))
   CANTIRENDI = XVALO(VALOBADA("MACONSIG", "SUM(CANT_RENDID)", CONDI$, "NOMESS"))
   CANTIDVAJ = XVALO(VALOBADA("MACONSIG", "SUM(Cant_Devu)", CONDI$, "NOMESS"))
   CANTISALDO# = CANTIENTRE - (CANTIRENDI + CANTIDVAJ)
   '
12 'LLAMO AL OBJETO PARA MOSTRAR LAS CANTIDADES
   Call REPLA(VDEF$, MKD$(CANTISALDO#), 1, 8)
   Call REPLA(VDEF$, MKD$(CANTISALDO#), 9, 8)
   OBX$ = OBJPLA$("CANTARMA", VDEF$)
   If OBX$ = "" Then GoTo 10
   '
   'CALCULO LA DIFERENCIA ENTRE EL SALDO ANTERIOR Y EL SALDO AJUSTADO
   SALAJUSTADO = CVD(Mid$(OBX$, 9, 8))
   If SALAJUSTADO < 0 Then
      Call MENSERR(24, "Imposible Procesar Saldos Negativos")
      GoTo 12
   End If
   '
   DIFERAJSTE = CANTISALDO# - SALAJUSTADO
   '
   REMI$ = "AR-" & PRONUREMIAJ&

   HO% = HOY(HOS$)
   'GENERO NUEVO REGISTRO
   SQLX$ = "SELECT * From RENCONSIG where Nu_Prove = " & NPRO% & " "
   Dim RenConsig As ADODB.Recordset
   Set RenConsig = New ADODB.Recordset
25 On Error Resume Next
   RenConsig.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   'FECH = FETEXCOR1$(HO%) 'COMENTADO POR QUE DA ERROR AL ASIGNAR AL CAMPO
   With RenConsig
        ' AQUÍ DEVOLUCION DIRECTA DEL MATERIAL RECIBIDO
        .AddNew
        !Fe_Recep = CVDAT(HO%)
        !NUMEBORE = 0
        !nro_ocom = REMI$
        !Nu_Prove = NPRO%
        !Codi_Conj = 0
        !Cant_Conj = 0
        !Codi_Elem = Ci%
        !CANT_RENDID = DIFERAJSTE
        !TipoMovim = "AJ"
        !MarProce = 0
        .Update
    End With
    '
    RenConsig.Close
    Set RenConsig = Nothing
    '
    Call AJUMACON
   
   '
'   'PROCESO DE CONSOLIDACION DE CONSIGNACIONES
'   Call CONSOCONSI(NPRO%, CI%)
   '
   Call COMUNI("Ajuste de Saldo Procesado")
   '
End Sub
'
Sub CONSOCONSI(NPRO%, Ci%)
    ' CONCILIACION DE CONSIGNACIONES
    
    SQLX$ = "SELECT * FROM MACONSIG "
    SQLX$ = SQLX$ + " WHERE Cod_Inte = " & Ci%
    SQLX$ = SQLX$ + " AND NPRO_OCOM = " & NPRO%
    SQLX$ = SQLX$ + " AND Cant_Rendid > Cant_Entreg "
    '
    Call ABRESTOCKS
    ACUMULADO = 0
    '
    Dim MACONTMP As ADODB.Recordset
    Set MACONTMP = New ADODB.Recordset
25  On Error Resume Next
    MACONTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    With MACONTMP
      Do While Not .EOF
        CANTENTRE = XVALO(!cant_entreg)
        CANTRENDI = XVALO(!CANT_RENDID)
        Diferencia = CANTRENDI - CANTENTRE
        'ASINGNO LA DIFERENCIA
        !CANT_RENDID = CANTENTRE
        ACUMULADO = ACUMULADO + Diferencia
        'GRABO
        .Update
     .MoveNext
     Loop
    End With
    MACONTMP.Close
    Set MACONTMP = Nothing
    '
    'SI EL ACUMULADO ES 0 SE VA
    If ACUMULADO <= 0 Then Exit Sub
    '
    SQLX$ = "SELECT * FROM MACONSIG "
    SQLX$ = SQLX$ + " WHERE COD_INTE = " & Ci%
    SQLX$ = SQLX$ + " AND NPRO_OCOM = " & NPRO%
    SQLX$ = SQLX$ + " AND CANT_RENDID < CANT_ENTREG "
    Set MACONTMP = New ADODB.Recordset
    '
26  On Error Resume Next
    MACONTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 26
   End If
   On Error GoTo 0
   '
    With MACONTMP
      Do While Not .EOF
        CANTENTRE = XVALO(!cant_entreg)
        CANTRENDI = XVALO(!CANT_RENDID)
        'ASINGNO LA DIFERENCIA
        Diferencia = CANTENTRE - CANTRENDI
        If Diferencia > ACUMULADO Then Diferencia = ACUMULADO
        !CANT_RENDID = Diferencia + XVALO(!CANT_RENDID)
        ACUMULADO = ACUMULADO - Diferencia
        'GRABO
        .Update
        If ACUMULADO <= 0 Then Exit Do
     .MoveNext
     Loop
    End With
    MACONTMP.Close
    Set MACONTMP = Nothing
    '
    'SI ACUMULADO > 0  GENERA UN NUEVO REGISTRO POR EL ACUMULADO
    If ACUMULADO <= 0 Then Exit Sub
    '
100 ABUS$ = RANDSTRING$(-10)
    CONDI$ = "NRO_OCOM = '" & ABUS$ & "'"
    If CantRegistros("MACONSIG", CONDI$) > 0 Then GoTo 100
    Call CreaRegistro("MACONSIG", "NRO_OCOM", ABUS$, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "NPRO_OCOM", CONDI$, NPRO%, REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "RASO_PROV", CONDI$, rasocli(-1 * NPRO%), REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "SECC_EJEC", CONDI$, 0, REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "Femi_Ocom", CONDI$, fech, REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "Deno_Secc", CONDI$, "", REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "Nro_Orfa", CONDI$, "", REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "Cod_Inte", CONDI$, Ci%, REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "Cod_Exte", CONDI$, CODEXT$(Ci%), REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "Descritem", CONDI$, DESCRIT0$(Ci%), REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "Lote_Ocom", CONDI$, "", REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "Unidad", CONDI$, Unimed(Ci%), REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "Cant_Asig", CONDI$, 0, REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "Cant_Entreg", CONDI$, 0, REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "Fech_Entreg", CONDI$, fech, REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "NRemi_Consig", CONDI$, "", REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "Valo_Unit", CONDI$, 0, REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "Moneda", CONDI$, MONEMI%, REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "Cant_Rendid", CONDI$, ACUMULADO, REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "Saldo_Pend", CONDI$, 0, REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "Observa", CONDI$, "", REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "Valo_Tot", CONDI$, 0, REGAF&, "NOMESS")
    Call ACTUREGISTRO("MACONSIG", "NRO_OCOM", CONDI$, "OC-000000-00", REGAF&, "NOMESS")
End Sub

Sub CONTROLCOMPRASANULADAS()
    
    Call DESHASFECHA(Des%, Has%, PERIO$)
    '
    If PERIO$ = "" Then Exit Sub
    '
    Desde$ = FETEXCOR1$(Des%)
    Hasta$ = FETEXCOR1$(Has%)
    '
    Dim Rst  As ADODB.Recordset
    Set Rst = New ADODB.Recordset
    SQLX$ = "select * from ocomenca where stat_ocom<3 and femi_ocom>='" & Desde$ & "' and femi_ocom<='" & Hasta$ & "'"
    '
    Rst.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, acmdtext
    '
    With Rst
        Do While Not .EOF
            OC$ = SAFETEXT$(!nume_ocom)
            CONDI$ = "nume_ocom='" & OC$ & "' and stat_ocom=9"
            CANTREG& = CantRegistros("ocomdeta", CONDI$)
            If CANTREG& > 0 Then
                !Stat_Ocom = 9
                .Update
            End If
            .MoveNext
        Loop
    End With
    
    Rst.Close
    Set Rst = Nothing
    
    Call COMUNI("Proceso de Control Finalizado")
    
End Sub

Sub CONTROLRECIBIDOCONBR()
    '
    ' BALANCEA ASIGNADO, ENTREGADO Y DEVUELTO. O SEA, EL SALDO CONSIGNADO
    SQLX$ = "Select * from OCOMDETA "
    SQLX$ = SQLX$ + " Where Stat_Ocom < 8 "
    SQLX$ = SQLX$ + " AND Stat_Ocom <> -1 "
    SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.001 "

    Dim RST1 As ADODB.Recordset
    Set RST1 = New ADODB.Recordset
24  On Error Resume Next
    RST1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 24
    End If
    On Error GoTo 0

    With RST1
      Do While Not .EOF
         CARECI0 = XVALO(!Cant_Rec)
         CARECI1 = XVALO(VALOBADA("BOLRECEP", "SUM(CANTOAPRO)", "COD_INTE=" & CStr(XVALO(!cod_inte)) & " AND NUMOC=" & CStr(XVALO(!nume_ocom))))
         If Abs(CARECI0 - CARECI1) > 0.05 Then
           OPX% = COMALTER%("Se Ha Detectado Una Diferencia en el Material Recibido en OC " & SAFETEXT$(!nume_ocom) & "con las Siguientes Cantidades\BR= " & TRIM$(FORMATNUM$(CARECI1, "F12.1")) & _
           "\OC=" & TRIM$(FORMATNUM$(CARECI0, "F12.1")) & " ¿Desea Corregir?\\No, Continuar\Si, Actualizar")
           If OPX% = 2 Then
               !Cant_Rec = CARECI1
               .Update
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
End Sub

Sub CUMPLECONS_ORD_CERRADAS()
'1. Recorrer OCOMDETA, ubicar los ítems con status cerrado o anulado (8 o 9), poner los materiales consignados (de ese ítem) con cantidad entregada = cantidad rendida más cantidad devuelta
    Screen.MousePointer = 11
    'PRESENTO LOS ITEMS DE LA ORDEN DE COMPRA SELECCIONADA
    '*****************************************************
    '
    If CLACONTRA% <> 1 Then Screen.MousePointer = 1: Exit Sub
    '
    SQLX$ = "SELECT * FROM OCOMDETA WITH(NOLOCK) WHERE Stat_Ocom >= 3 "
    Dim Rst As ADODB.Recordset
    Set Rst = READSET(SQLX$)
    With Rst
      Do While Not .EOF
         CIXI% = XVALO(!cod_inte)
         NUMECO& = XVALO(!nume_ocom)
         NUOR% = XVALO(!NuOrd_Item)
         NROCOM$ = "OC-" & Right$("000000" & SAFETEXT$(NUMECO&), 6) & "-" & Right$("00" & SAFETEXT$(NUOR%), 2)

         CONDI$ = " Nro_Ocom =  '" & NROCOM$ & "' "
         CONDI$ = CONDI$ + " AND COD_INTE = " & CIXI%
         'IGUALA CANTIDADES
         STRSQL = "update MACONSIG set Cant_Entreg = (CANT_RENDID + Cant_DEVU) where " & CONDI$
         FLEXCONN.Execute (FILTSQL$(STRSQL))
         .MoveNext
      Loop
   End With
   Rst.Close
   Set Rst = Nothing

'2. Recorrer OCOMDETA, ubicar los ítems con status cerrado o anulado (9), poner los materiales consignados (de todos los ítems) con cantidad asignada = cantidad rendida
    'PRESENTO LOS ITEMS DE LA ORDEN DE COMPRA SELECCIONADA
    '*****************************************************
    SQLX$ = "SELECT * FROM OCOMENCA WITH(NOLOCK) WHERE Stat_Ocom >= 3 "
    Dim RST1 As ADODB.Recordset
    Set RST1 = READSET(SQLX$)
    With RST1
      Do While Not .EOF
         NUMECO& = XVALO(!nume_ocom)
         NROCOM$ = "OC-" & Right$("000000" & SAFETEXT$(NUMECO&), 6) & "-"

         CONDI$ = " Nro_Ocom LIKE  '" & NROCOM$ & "%' "
         'IGUALA CANTIDADES
         STRSQL = "update MACONSIG set CANT_ASIG = CANT_RENDID where " & CONDI$
         FLEXCONN.Execute (FILTSQL$(STRSQL))
         .MoveNext
      Loop
   End With

  Screen.MousePointer = 1
  RST1.Close
  Set RST1 = Nothing
End Sub

Sub CUMPLECONSIGXART()

    If DERACCE%("CUMCONAR", "Dar Cumplidas Consignaciones por Articulo") < 2 Then Exit Sub
    
    'PRESENTO ORDENES DE COMPRA PENDIENTES
    '*************************************
    CONDI$ = " OCOMDETA.Stat_Ocom < 8 "  ' Excluyo las que estan anuladas O COMPLIDAS
'    CONDI$ = CONDI$ + " AND OCOMDETA.Cant_Rec < OCOMDETA.Cant_Ocom "
    CONDI$ = CONDI$ + " AND OCOMENCA.Stat_Ocom = 0 "
    CONDI$ = CONDI$ & " AND OCOMENCA.Nume_Ocom > 0 "
    CONDI$ = CONDI$ + " AND OCOMENCA.CODIEMPR = " & CodiEmp%
    CONDI$ = CONDI$ + " AND OCOMDETA.CODIEMPR = " & CodiEmp%
    CONDI$ = CONDI$ + " ORDER BY OCOMENCA.Nume_Ocom  DESC"
    Call CARGALISEL("ORCOPEN", "OCOMENCA+OCOMDETA/OCOMENCA.Nume_Ocom = OCOMDETA.Nume_Ocom", "OCOMENCA.Nume_Ocom/F7;OCOMENCA.Refe_Ocom/A44", CONDI$, "DISTINCT")
10  Call SELECHO("ORCOPEN/Indice General de Ordenes de Compra (Pendientes)")
    '
    NUMECO& = XVALO(VALACT1$("ORCOPEN"))
    If NUMECO& < 1 Then Exit Sub
    
    'PRESENTO LOS ITEMS DE LA ORDEN DE COMPRA SELECCIONADA
    '*****************************************************
    CONDI$ = "Nume_Ocom = " & NUMECO&
    CONDI$ = CONDI$ + " AND Stat_Ocom < 8"
    CONDI$ = CONDI$ + " ORDER BY Nuord_Item"
    Call CARGALISEL("DETACOMP", "OCOMDETA", "COD_INTE/C16+MASTER/A48;CANT_OCOM/F9.1;Nuord_Item/F5", CONDI$)
20  Call SELECHO("DETACOMP/Código                            Descripción                                                                                                            Cant.Comp.   Orden")
    '
    CIXI% = CODINT%(VALACT1$("DETACOMP"))
    If CIXI% < 1 Then GoTo 10
     NUOR% = XVALO(Right(VALACT2$("DETACOMP"), 5))

    'PRESENTO LOS ITEMS DE LA ORDEN DE COMPRA SELECCIONADA
    '*****************************************************
    NROCOM$ = "OC-" & Right$("000000" & SAFETEXT$(NUMECO&), 6) & "-" & Right$("00" & SAFETEXT$(NUOR%), 2)

30  CONDI$ = " Nro_Ocom =  '" & NROCOM$ & "' "
    CONDI$ = CONDI$ + " AND CANT_ASIG  > (CANT_ENTREG + 0.05)"
'    CONDI$ = CONDI$ + " ORDER BY COD_EXTE"
    Call CARGALISEL("DETAMACO", "MACONSIG", "COD_INTE/C16+MASTER/A48;", CONDI$, "DISTINCT")
    Call SELECHO("DETAMACO/Detalle de Consignaciones Pendienes: " & SAFETEXT$(NUMECO&) & ")")
    '
    CIXI% = CODINT%(VALACT1$("DETAMACO"))
    If CIXI% < 1 Then GoTo 20
    '
    CONDI$ = "Nro_OCOM = '" & NROCOM$ & "'"
    CONDI$ = CONDI$ + " AND COD_INTE = " & CIXI%
    
    'IGUALA CANTIDADES
    STRSQL = "update MACONSIG set Cant_Asig = Cant_Entreg where " & CONDI$
    FLEXCONN.Execute (FILTSQL$(STRSQL))
    '---------------------------------------------------------------------------------------
    Call COMUNI("Transacción Finalizada")
    GoTo 30
    
End Sub



Function PRONUREMIAJ&()
    'PROXIMO NUMERO DE MATERIALES EN CONSIGNACION CON
    Call ABRESTOCKS
    UNUM& = 0
    '
    SQLX$ = "SELECT max(cast(substring(Nro_Ocom , 4,12)as int))as Maximo FROM RENCONSIG WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE Nro_Ocom LIKE 'AR-%'"
    Dim MIRS As ADODB.Recordset
    Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
        UNUM& = XVALO(MIRS!Maximo)
    MIRS.Close
    PRONUREMIAJ& = 1 + UNUM&
    Set MIRS = Nothing

End Function




Sub ANULACOTIZA(NUOCO&, RTA%)
    '
    RT% = 0
    SQLX$ = "SELECT * FROM OCOMDETA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE STAT_OCOM < 9 "
    SQLX$ = SQLX$ + " AND NUME_OCOM = " & NUOCO&
    SQLX$ = SQLX$ + " AND CANT_REC > 0"
    Call ABRECOMPRAS
    Dim OCOMTMP As ADODB.Recordset
    Set OCOMTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With OCOMTMP
      Do While Not .EOF
        Ci% = XVALO(!cod_inte)
        CONDI$ = "COD_INTE = " & Ci% & " AND NUME_OCOM = " & NUOCO&
        If CantRegistros("RESERVA", CONDI$) > 0 Then
           Call COMUNI("Imposible Anular Esta Orden de Compra\Artículos con Solicitud de Cotización Realizada")
           RT% = 1
           GoTo 99
        End If
      .MoveNext
      Loop
    End With
    OCOMTMP.Close
    Set OCOMTMP = Nothing
    CONDI$ = "NUME_OCOM = " & NUOCO&
    '
    Call ACTUREGISTRO("RESERVA", "NUME_OCOM", CONDI$, 0, REGAF&, "NOMESS")
    '
99  RTA% = RT%
End Sub

Sub DEPUCONSIG()
    '
    CONDI$ = "COD_INTE>0"
    FIN& = CantRegistros("MACONSIG", CONDI$)
    SQLX$ = "SELECT * FROM MACONSIG WHERE " & CONDI$
    Dim RST_CONSIGNA As ADODB.Recordset
    Set RST_CONSIGNA = New ADODB.Recordset
    RST_CONSIGNA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    JJ& = 1
    With RST_CONSIGNA
        Do While Not .EOF
            Call TRACE(24, JJ&, FIN&)
            NROOCOM& = XVALO(Mid$(SAFETEXT$(!nro_ocom), 4, 6))
            CONDI$ = "NUME_OCOM = " & NROOCOM&
            Stat_Ocom = XVALO(VALOBADA("OCOMENCA", "STAT_OCOM", CONDI$))
            If Stat_Ocom > 0 Then
                NORCOM$ = Left$(SAFETEXT$(!nro_ocom), 9) + "%"
                CIX% = XVALO(!cod_inte)
                NCX% = XVALO(!Npro_Ocom)
                '
                !cant_entreg = !cant_asig
                !CANT_RENDID = !cant_entreg
                !cant_devu = 0
                .Update
                '
                Call BORRAREGISTROS("RENCONSIG", "nro_ocom like'" & NORCOM$ & "'", REGAF&)
            Else
                ORD% = XVALO(Mid$(SAFETEXT$(!nro_ocom), 11, 2))
                CONDI$ = "NUME_OCOM= " & NROOCOM& & " AND NUORD_ITEM=" & ORD%
                Stat_Ocom = XVALO(VALOBADA("OCOMDETA", "STAT_OCOM", CONDI$))
                CCOMP = XVALO(VALOBADA("OCOMDETA", "cant_ocom", CONDI$))
                CRECI = XVALO(VALOBADA("OCOMDETA", "cant_rec", CONDI$))
                SALDO = CCOMP - CRECI
                If Stat_Ocom > 0 Or SALDO <= 0.005 Then
                If SALDO <= 0.005 Then
                A = B
                End If
                    NORCOM$ = Left$(SAFETEXT$(!nro_ocom), 9) + "%"
                    CIX% = XVALO(!cod_inte)
                    NCX% = XVALO(!Npro_Ocom)
                    '
                    !cant_entreg = !cant_asig
                    !CANT_RENDID = !cant_entreg
                    !cant_devu = 0
                    .Update
                    '
                    Call BORRAREGISTROS("RENCONSIG", "nro_ocom like'" & NORCOM$ & "' AND codi_elem=" & CIX%, REGAF&)
                End If
            End If
            .MoveNext
            JJ& = JJ& + 1
        Loop
    End With
    '
    Call COMUNI("Proceso Finalizado")
    '
    Exit Sub
    '
'    CONDI$ = "COD_INTE>0"
'    FIN& = CantRegistros("MACONSIG", CONDI$)
'    SQLX$ = "SELECT * FROM MACONSIG WHERE " & CONDI$
'    Dim RST_CONSIGNA As ADODB.Recordset
'    Set RST_CONSIGNA = New ADODB.Recordset
'    RST_CONSIGNA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'    '
'    JJ& = 1
'    With RST_CONSIGNA
'        Do While Not .EOF
'            Call TRACE(24, JJ&, FIN&)
'            NROOCOM& = XVALO(Mid$(SAFETEXT$(!Nro_Ocom), 4, 6))
'            CONDI$ = "NUME_OCOM = " & NROOCOM&
'            Stat_Ocom = XVALO(VALOBADA("OCOMENCA", "STAT_OCOM", CONDI$))
'            If Stat_Ocom > 0 Then
'                NORCOM$ = Left$(SAFETEXT$(!Nro_Ocom), 9) + "%"
'                CIX% = XVALO(!cod_inte)
'                NCX% = XVALO(!Npro_Ocom)
'                '
'                ' CANTIDAD RENDIDA + CANTIDAD RENDIDA
'                '
'                SQLY$ = "SELECT SUM(CANT_RENDID) AS SUREN FROM RENCONSIG WITH(NOLOCK) "
'                SQLY$ = SQLY$ + " WHERE (NRO_OCOM LIKE '" & NORCOM$ & "'"
'                   If Left$(NORCOM$, 9) = "OC-000000" Then
'                       SQLY$ = SQLY$ + " OR NRO_OCOM = '' OR NRO_OCOM IS NULL)"
'                     Else
'                       SQLY$ = SQLY$ + ")"
'                   End If
'                SQLY$ = SQLY$ + " AND Codi_Elem = " & CIX% & " "
'                SQLY$ = SQLY$ + " AND Nu_Prove = " & NCX%
'                SQLY$ = SQLY$ + " AND (TIPOMOVIM = 'CC' OR TIPOMOVIM = 'QR')"
'                Set QRENDI = READSET(SQLY$)
'                SURENDI = XVALO(QRENDI!SUREN)
'                '
'                ' AJUSTES + DEVOLUCIONES + CANTIDAD DEVUELTA
'                '
'                SQLY$ = "SELECT SUM(CANT_RENDID) AS SUREN FROM RENCONSIG WITH(NOLOCK) "
'                SQLY$ = SQLY$ + " WHERE NRO_OCOM LIKE '" & NORCOM$ & "' "
'                SQLY$ = SQLY$ + " AND Codi_Elem = " & CIX% & " "
'                SQLY$ = SQLY$ + " AND Nu_Prove = " & NCX%
'                SQLY$ = SQLY$ + " AND (TIPOMOVIM = 'QD' OR TIPOMOVIM = 'AJ' OR TIPOMOVIM= 'DV')"
'                Set QRENDI = READSET(SQLY$)
'                SUDEVU = XVALO(QRENDI!SUREN)
'                '
'                CENTREG = XVALO(!cant_entreg)
'                CANTASIG = XVALO(!cant_asig)
'                SALDOENTREG = CENTREG - SURENDI - SUDEVU
'                If SALDOENTREG > 0 Then
'                    SQLX$ = "SELECT * From RENCONSIG where Nu_Prove = " & NCX% & " "
'                    Dim RenConsig As ADODB.Recordset
'                    Set RenConsig = New ADODB.Recordset
'25                  On Error Resume Next
'                    RenConsig.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'                    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'                       On Error GoTo 0
'                       Call CapturaErrorOpen
'                       GoTo 25
'                    End If
'                    On Error GoTo 0
'                    '
'                    CODDX$ = CODEXT$(CIX%)
'                    CAPENDE# = SALDOENTREG
'                    FE% = FERECEP%: If FE% < 1 Then FE% = HOY(HO$)
'                    NOCOM$ = TRIM$(Str$(NROOCOM&))
'                    While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
'                    NORCOM$ = "OC-" + Left$(NOCOM$, 6) + "%"
'                    '
'                    With RenConsig
'                       .AddNew
'                       !Fe_Recep = CVDAT(FE%)
'                       !NUMEBORE = NUBOL&
'                       !Nro_Ocom = NORCOM$
'                       !Nu_Prove = NCX%   ' NCX% ERROR EN NOMBRE DE VARIABLE
'                       !Codi_Conj = 0
'                       !Cant_Conj = 0
'                       !Codi_Elem = CIX%  ' CIX% ERROR EN NOMBRE DE VARIABLE
'                       !cant_rendid = CAPENDE#
'                       !TipoMovim = "AJ"
'                       !MarProce = 1
'                       .Update
'                    End With
'                    RenConsig.Close: Set RenConsig = Nothing
'                Else
'                    !cant_entreg = SURENDI + SUDEVU
'                    !cant_asig = XVALO(!cant_entreg) - SUDEVU
'                    .Update
'                End If
'                '
'            End If
'            JJ& = JJ& + 1
'            .MoveNext
'        Loop
'    End With
'
'    Call COMUNI("Proceso Terminado")
    
'    If InStr(1, EMPREAC$, "GABAL") < 1 Then GoTo 40
'    '
'    Call ABRECONEXION
'    Call SELECHO("PROVED1")
'    NPX% = XVALO(VALACT1$("PROVED1"))
'    '
'    Screen.MousePointer = 11
'    ' ELIMINA LOS MOVIMIENTOS QUE SON SOLO RENDICION Y LOS AJUSTES POR ARMADOR
'    CONDI$ = "NPRO_OCOM=" & NPX% & " AND CANT_ASIG < 0.05 AND CANT_ENTREG < 0.05"
'    Call BORRAREGISTROS("MACONSIG", CONDI$, REGAF&)
'    '
'    '
'    ' BALANCEA ENTREGADO, RENDIDO Y DEVUELTO. O SEA, EL SALDO EN EL ARMADOR
'    SQLX$ = "SELECT * FROM MACONSIG WHERE Npro_Ocom=" & NPX%
'    Dim RST As ADODB.Recordset
'    Set RST = New ADODB.Recordset
'25  On Error Resume Next
'    RST.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 25
'    End If
'    On Error GoTo 0
'
'    With RST
'      Do While Not .EOF
'         E# = XVALO(!CANT_ENTREG)
'         R# = XVALO(!CANT_RENDID)
'         d# = XVALO(!Cant_Devu)
'         '
'         If E# < R# + d# Then
'            E# = R# + d#
'            !CANT_ENTREG = E#
'            .Update
'         End If
'        .MoveNext
'      Loop
'    End With
'    On Error Resume Next
'    RST.Close
'    Set RST = Nothing
'    On Error GoTo 0
'    '
'    '
'    ' BALANCEA ASIGNADO, ENTREGADO Y DEVUELTO. O SEA, EL SALDO CONSIGNADO
'    SQLX$ = "SELECT * FROM MACONSIG WHERE CANT_ASIG > 0.05 AND Npro_Ocom=" & NPX%
'    Dim RST1 As ADODB.Recordset
'    Set RST1 = New ADODB.Recordset
'24  On Error Resume Next
'    RST1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 24
'    End If
'    On Error GoTo 0
'
'    With RST1
'      Do While Not .EOF
'         A# = XVALO(!CANT_ASIG)
'         E# = XVALO(!CANT_ENTREG)
'         d# = XVALO(!Cant_Devu)
'         '
'         If A# < E# - d# Then
'            A# = E# - d#
'            !CANT_ASIG = A#
'            .Update
'         End If
'        .MoveNext
'      Loop
'    End With
'    On Error Resume Next
'    RST1.Close
'    Set RST1 = Nothing
'    On Error GoTo 0
'    '
'    '
'    ' COMPENSA ASIGNADO CON ENTREGADO Y DEVUELTO. AGREGA UN REGISTRO CON CANTIDAD
'    ' ASIGNADA, CUANDO LA CANTIDAD ENTREGADA FUE MAYOR A LA ASIGNADA
'    SQLX$ = "SELECT Cod_Inte, SUM(Cant_Asig) AS SxCant_Asig, SUM(Cant_Entreg) AS SxCant_Entreg , SUM(Cant_Devu) AS SxCant_Devu "
'    SQLX$ = SQLX$ + " FROM MACONSIG WHERE Npro_Ocom = " & NPX% & " "
'    SQLX$ = SQLX$ + " GROUP BY Cod_Inte ORDER BY Cod_Inte ASC"
'
'    Dim MACONSITMP As ADODB.Recordset
'    Set MACONSITMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
'    With MACONSITMP
'      Do While .EOF = False
'        CIX% = XVALO(!cod_inte)
'        CODIGO$ = CODEXT$(CIX%)
'        SURENDI = XVALO(!SxCant_Asig) - (XVALO(!SxCant_Entreg) - XVALO(!SxCant_Devu))
'
'        If SURENDI < 0 Then
'             HOII% = HOY(HO$)
'             FECH = CVDAT(HOII%)
'100          ABUS$ = RANDSTRING$(-10)
'             CONDI$ = "NRO_OCOM = '" & ABUS$ & "'"
'             If CantRegistros("MACONSIG", CONDI$) > 0 Then GoTo 100
'             Call CreaRegistro("MACONSIG", "NRO_OCOM", ABUS$, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "NPRO_OCOM", CONDI$, NPX%, REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "RASO_PROV", CONDI$, RASOCLI(-1 * NCX%), REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "SECC_EJEC", CONDI$, 0, REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "Femi_Ocom", CONDI$, FECH, REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "Deno_Secc", CONDI$, "", REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "Nro_Orfa", CONDI$, "", REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "Cod_Inte", CONDI$, CIX%, REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "Cod_Exte", CONDI$, CODEXT$(CIX%), REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "Descritem", CONDI$, DESCRIT0$(CIX%), REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "Lote_Ocom", CONDI$, "", REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "Unidad", CONDI$, UNIMED(CIX%), REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "Cant_Asig", CONDI$, Abs(SURENDI), REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "Cant_Entreg", CONDI$, 0, REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "Fech_Entreg", CONDI$, FECH, REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "NRemi_Consig", CONDI$, "", REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "Valo_Unit", CONDI$, 0, REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "Moneda", CONDI$, MONEMI%, REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "Cant_Rendid", CONDI$, 0, REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "Saldo_Pend", CONDI$, 0, REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "Observa", CONDI$, "", REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "Valo_Tot", CONDI$, 0, REGAF&, "NOMESS")
'             Call ACTUREGISTRO("MACONSIG", "NRO_OCOM", CONDI$, "OC-999999-99", REGAF&, "NOMESS")
'        End If
'        .MoveNext
'      Loop
'    End With
'    '
'    '
'    Screen.MousePointer = 1
'    Call COMUNI("Proceso Finalizado")
'
'Exit Sub
' deshabilitado a partir del nuevo proceso de armadores de gabal desde 4 noviembre 2011
'90  Call ABRECOMPRAS
'    '
'    SQLX$ = "SELECT * FROM MACONSIG"
'    'SQLX$ = SQLX$ + " WHERE CANT_ASIG - CANT_ENTREG > 0.05 "
'    Dim MACTMP As ADODB.Recordset
'    Set MACTMP = New ADODB.Recordset
'23  On Error Resume Next
'    MACTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 23
'   End If
'   On Error GoTo 0
'
'   With MACTMP
'      Do While Not .EOF
'         OCOMP$ = SAFETEXT$(!Nro_Ocom)
'         nume& = XVALO(Mid$(OCOMP$, 4, 6))
'         ORD% = XVALO(Mid$(OCOMP$, 11, 26))
'         If nume& < 1747 Then
'            !CANT_ASIG = !CANT_ENTREG
'            !Cant_Devu = 0
'            !cant_rendid = !CANT_ENTREG
'            .Update
'         End If
'        .MoveNext
'      Loop
'
'    End With
'    MACTMP.Close
'    Set MACTMP = Nothing
    Call COMUNI("Proceso Terminado")
End Sub

Sub PROVEARTIC()

5        Call SELECHO("MASTER") ' DESPLIEGA LA LISTADO DE ARTICULOS
         '
         COD$ = TRIM$(VALACT1$("MASTER"))
         If COD$ = "" Then Exit Sub
         '
         Ci% = CODINT%(COD$)
         ' BUSCO EN LA TABLA PROVEDCOD TODOS LOS REGISTROS QUE COINCIDAN CON EL CI%
         SQLX$ = "SELECT * FROM PROVECOD WITH (NOLOCK)"
         SQLX$ = SQLX$ + "WHERE COD_INTE = " & Ci% & " AND MARBORRA = 0 "
         Call APERBASDAT("COMPRAS")
         Dim PROVECODTMP As ADODB.Recordset
         Set PROVECODTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         Call SETULTREG("!CODPROVE", 0)
         JJ& = 0
         With PROVECODTMP
            Do While Not .EOF ' RECORRO LA TABLA Y GRABO EN EL ARCHIVO PARA MOSTRAR LA VENTANA
                X$ = REGBLAN$("CODPROVE")
                Call REPLA(X$, MKI(!Num_Prov), 1, 2)
                JJ& = JJ& + 1
                Call GRAREG("!CODPROVE", X$, JJ&)
            .MoveNext
            Loop
         End With
         PROVECODTMP.Close
         Set PROVECODTMP = Nothing
         
10       XYX% = VENTABU%("CODPROVE") ' MUESTRO LA VENTANA
         If XYX% < 0 Then GoTo 5     ' SI CIERRA VUELVE AL SELECHO
         '
         CONDI$ = "COD_INTE=" & Ci%
         Call ACTUREGISTRO("PROVECOD", "MARBORRA", CONDI$, 1, REG&) ' HACE UN MARBORRA
         '
         
         SQLX$ = "SELECT * FROM PROVECOD "
         SQLX$ = SQLX$ + " WHERE COD_INTE = " & Ci%
         Set PROVCODTMP = New ADODB.Recordset
         PROVCODTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
           With PROVCODTMP
             For i& = 1 To ULTREG&("!CODPROVE") ' RECORRO EL ARCHIVO Y GRABO
               XX$ = REGLEIDO$("!CODPROVE", i&) ' LA INFORMACION DE LA VENTANA
               NPRO% = CVI(Mid$(XX$, 1, 2))
               .AddNew   ' YA CONVERTIDO
               On Error Resume Next
               !CODIEMPR = CodiEmp%
               On Error GoTo 0
               !Num_Prov = NPRO%
               !cod_inte = Ci%
               !PRECIO = 0
               !Moneda = 0
               !FECHA = CVDAT(0)
               !MARBORRA = 0
               .Update
             Next i&
           End With
         PROVCODTMP.Close
         Set PROVCODTMP = Nothing
         CONDI$ = "MARBORRA = 1"
         Call BORRAREGISTROS("PROVECOD", CONDI$, REG&) ' BORRO LOS MARCADOS CON 1
'
End Sub
Sub REVERTIR_ITEMCUMPLI()
    'REVERTIR ITEM CUMPLIDOS
    '
    'LISTA DE ORDENES DE COMPRA
    CONDI$ = "Stat_Ocom < 9 AND Nume_Ocom > 0 ORDER BY Nume_Ocom"
    Call CARGALISEL("INDIOCAC", "OCOMENCA", "Nume_Ocom/F6;Refe_OcoM/A44", CONDI$, "DISTINCT")
10  Call SELECHO("INDIOCAC/Indice de Ordenes de Compra")
    '
    NRO& = XVALO(VALACT1("INDIOCAC"))
    If NRO& < 1 Then Exit Sub
    '
    CONDI$ = "CANT_REC < CANT_OCOM AND STAT_OCOM = 8 AND NUME_OCOM = " & NRO&
    Call CARGALISEL("INDICUM", "OCOMDETA", "Nume_Ocom/F6;COD_INTE/C16+MASTER/A48", CONDI$, "DISTINCT")
    '
    'LISTA DE SELECCION DE ITEMS DE LA ORDEN DE COMPRA SELECCIONADA
    If CantRegistros("OCOMDETA", CONDI$, "NOMESS") < 1 Then
      Call COMUNI("No Hay Items Cumplidas En Esta Orden \Que Puedan Restaurarse")
      GoTo 10
    End If
    
20  Call SELECHO("INDICUM/Indice de Items O.C.: " & NRO&)
    NROOC& = XVALO(VALACT1("INDICUM"))
    If NROOC& < 1 Then GoTo 10
    CODEXTER$ = TRIM$(Mid$(VALACT2$("INDICUM"), 1, 16))
    CIXI% = CODINT%(CODEXTER$)
    
    'REVERTIR MARCA DE ITEM CUMPLIDO ' SE AGREGA CONDICION STAT_OCOM = 8 PARA QUE REVIERTA SOLO LOS CUMPLIDOS
    If COMALTER%("¿ Confirma Revertir Item: " & CODEXTER$ & "\Marcado Como Cumplido  \\No, Abandonar\Si, Revertir") = 2 Then
       CONDI$ = " NUME_OCOM = " & NROOC& & " AND STAT_OCOM = 8 and COD_INTE = " & CIXI%
       Call ACTUREGISTRO("OCOMDETA", "STAT_OCOM", CONDI$, 1, REGAF&, "NOMESS")
       Call COMUNI("Se Ha Completado la Operación")
       GoTo 20
    Else
      GoTo 20
    End If
    
End Sub

Function TieneItemsRecibidos&(Noc&)
   'TieneItemsRecibidos& = 0 no recibio nada aun
   'TieneItemsRecibidos& > 0 si ya recibio
   CONDI$ = "Cant_Rec > 0 AND Nume_Ocom = " & Noc&
   TieneItemsRecibidos& = CantRegistros&("OCOMDETA", CONDI$, "NOMESS")
End Function

'
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

Sub Command10_Click()
   Command10.Enabled = False
   If EXISTE%("ORTRAB04.EXE") > 0 Then
       Call MENSERR(24, "No Compatible con Módulo\de Ordenes de Trabajo.")
     Else
        Call DESANULOC
   End If
   Command10.Enabled = True
End Sub


Sub Command12_Click()
    '
    Command12.Enabled = False
    OPX% = COMALTER%("Detalle de Ultimas Compras\\Por Proveedor\Por Artículo\Cancelar")
    If OPX% < 1 Or OPX% > 2 Then
      Command12.Enabled = True
      Exit Sub
    End If
    If OPX% = 1 Then
      Call OPLISCO07.Command12_Click
    Else
      Call OPLISCO07.Command20_Click
    End If
    Command12.Enabled = True
    '
End Sub

Sub Command13_Click()
    '
    Command13.Enabled = False
    Call MODIOCOM
99  Command13.Enabled = True
    '
End Sub

Sub Command14_Click()
   Command14.Enabled = False
   Call OCOMANU
   Command14.Enabled = True
End Sub

Sub Command15_Click()
   Command15.Enabled = False
   '
   '1.-REALIZO ALTERNA
'10 OPX% = ALTERNA%("Consumo de Materiales\Dar Por Cumplida Consignaciones Pendientes\Dar Por Cumplida Consignaciones Pendientes Por Item")
'   If OPX% < 1 Or OPX% > 4 Then GoTo 99
'   '
'   If OPX% = 1 Then
'      Call DEVOLCONSI
'      GoTo 10
'   ElseIf OPX% = 2 Then
'      Call CUMPLECONSIG
'      GoTo 10
'   ElseIf OPX% = 3 Then
'      Call CUMPLECONSIGXART
'      GoTo 10
'   End If


10 OPX% = ALTERNA%("Consumo de Materiales\Ajuste de Consignaciones Por Proveedor\Dar Por Cumplida Consignaciones Pendientes\Dar Por Cumplida Consignaciones Pendientes Por Item")
   If OPX% < 1 Or OPX% > 4 Then GoTo 99
   '
   If OPX% = 1 Then
      Call DEVOLCONSI
      GoTo 10
   ElseIf OPX% = 2 Then
      'AjusteAFechaConsignacionPorProveedorOrdenCompra
      Call COMUNI("Opción No Disponible")
      GoTo 10
   ElseIf OPX% = 3 Then
      Call CUMPLECONSIG
      GoTo 10
   ElseIf OPX% = 4 Then
      Call CUMPLECONSIGXART
      GoTo 10
   End If
99 Command15.Enabled = True
   '
End Sub
Sub CUMPLECONSIG()
    '
    
    If DERACCE%("CUMCONSI", "Dar Por Cumplida Consignación de Materiales") < 2 Then Exit Sub
    
    'PRESENTO ORDENES DE COMPRA PENDIENTES
    '*************************************
    CONDI$ = " OCOMDETA.Stat_Ocom <8 "  ' Excluyo las que estan anuladas O COMPLIDAS
    CONDI$ = CONDI$ + " AND OCOMDETA.Cant_Rec < OCOMDETA.Cant_Ocom "
    CONDI$ = CONDI$ + " AND OCOMENCA.Stat_Ocom = 0 "
    CONDI$ = CONDI$ & " AND OCOMENCA.Nume_Ocom > 0 "
    CONDI$ = CONDI$ + " AND OCOMENCA.CODIEMPR = " & CodiEmp%
    CONDI$ = CONDI$ + " AND OCOMDETA.CODIEMPR = " & CodiEmp%
    CONDI$ = CONDI$ + " ORDER BY OCOMENCA.Nume_Ocom  DESC"
    Call CARGALISEL("ORCOPEN", "OCOMENCA+OCOMDETA/OCOMENCA.Nume_Ocom = OCOMDETA.Nume_Ocom", "OCOMENCA.Nume_Ocom/F7;OCOMENCA.Refe_Ocom/A44", CONDI$, "DISTINCT")
10  Call SELECHO("ORCOPEN/Indice General de Ordenes de Compra (Pendientes)")
    '
    NUMECO& = XVALO(VALACT1$("ORCOPEN"))
    If NUMECO& < 1 Then Exit Sub
    '---------------------------------------------------------------------------------------
    
    'PRESENTO LOS ITEMS DE LA ORDEN DE COMPRA SELECCIONADA
    '*****************************************************
    CONDI$ = "Nume_Ocom = " & NUMECO&
    CONDI$ = CONDI$ + " AND Stat_Ocom < 8"
    CONDI$ = CONDI$ + " ORDER BY Nuord_Item"
    Call CARGALISEL("DETACOMP", "OCOMDETA", "COD_INTE/C16+MASTER/A48;Nuord_Item/F5", CONDI$)
20  Call SELECHO("DETACOMP/Detalle de Orden de Compra: " & SAFETEXT$(NUMECO&) & ")")
    '
    CIXI% = CODINT%(VALACT1$("DETACOMP"))
    If CIXI% < 1 Then GoTo 10
    '
    '---------------------------------------------------------------------------------------
    NUOR% = XVALO(Right(VALACT2$("DETACOMP"), 5))
    codextx$ = VALACT1$("DETACOMP")
    '
    'GRABO EN LA TABLA MACONSIG LOS COMPONENTES ASIGNADOS DEL ITEM SELECCIONADO.
    'CONSIDERNADO EL NUMERO DE ORDEN DE COMPRA Y EL NUMERO DE ORDEN QUE IDENTIFICA
    '---------------------------------------------------------------------------------------
    
    'PRESENTO LOS ITEMS DE LA ORDEN DE COMPRA SELECCIONADA
    '*****************************************************
    CONDI$ = "Nume_Ocom = " & NUMECO&
    CONDI$ = CONDI$ + " AND Stat_Ocom < 8"
    CONDI$ = CONDI$ + " ORDER BY Nuord_Item"
    Call CARGALISEL("DETACOMP", "OCOMDETA", "COD_INTE/C16+MASTER/A48;Nuord_Item/F5", CONDI$)
    'AL ITEM SELECCIONADO.
    '***************************************************************************
    NROCOM$ = "OC-" & Right$("000000" & SAFETEXT$(NUMECO&), 6) & "-" & Right$("00" & SAFETEXT$(NUOR%), 2)
    CONDI$ = "Nro_OCOM = '" & NROCOM$ & "'"
    CONDI$ = CONDI$ + " AND Cant_Asig > Cant_Entreg AND Cant_Asig >= 0.05 "
    If CantRegistros("MACONSIG", CONDI$, "NOMESS") < 1 Then
      Call COMUNI("El Código: '" & codextx$ & "' \ No Posee Saldos de Material Consignado")
      GoTo 20
    End If
    '
    ' PONE LA CANTIDAD ASIGNADA IGUAL A LA ENTREGADA PARA NO DESPACHAR MÁS
    '**********************************************************************
    STRSQL = "update MACONSIG set Cant_Asig = Cant_Entreg where " & CONDI$
    FLEXCONN.Execute (FILTSQL$(STRSQL))
    '---------------------------------------------------------------------------------------
    Call COMUNI("Transacción Finalizada")
    '
End Sub
Private Sub Command17_Click()
    '
    Command17.Enabled = False
    Screen.MousePointer = 11
    J& = 0: k& = 0: OCOA& = 0
    '
    For i& = 1 To ULTREG&("OCOMPRA")
         XX$ = REGLEIDO$("OCOMPRA", i&)
         '
         If CVS(Left(XX$, 4)) >= 0.5 Then    ' depura O/C's
             k& = k& + 1                     ' sin numero
             Call GRAREG("OCOMPRA", XX$, k&)
         End If
         '
         ST% = Asc(Mid$(XX$, 62, 1))
         '
         If ST% < 1 Then     '    excluye cerradas
            If CVS(Left$(XX$, 4)) <> OCOA& Then
               NP% = CVI(Mid$(XX$, 7, 2))
               If NP% < 1 Or NP% = 9999 Then
                  J& = J& + 1
                  Call GRAREG("!OCOMPRA", XX$, J&)
                  OCOA& = CVS(Left$(XX$, 4))
               End If
            End If
         End If
    Next i&
    '
    Call SETULTREG("OCOMPRA", k&)
    Call SETULTREG("!OCOMPRA", J&)
    Call CIERRARCH("!OCOMPRA")
    '
    If J& < 1 Then
          Call MENSERR(24, "No hay Pedidos de Suministro\Pendientes que Puedan Asignarse.")
          Screen.MousePointer = 1
          GoTo 99
    End If
    '
    INDI% = 1
    Screen.MousePointer = 1
    '
    XX$ = REGSEL$("!OCOMPRA")
    If Len(XX$) <= 4 Then GoTo 99
    '
    NUOCO& = CVS(Left$(XX$, 4))
    '
    YY$ = REGSEL$("PROVED1")
    If Len(YY$) <= 4 Then GoTo 99
    NP1% = CVI(Left$(YY$, 2))
    PROVEX$ = Mid$(YY$, 3, 30)
    '
    Screen.MousePointer = 11
    JJ& = 0
    FIN& = ULTREG&("OCOMPRA")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("OCOMPRA", U&)
       '
       If CVS(Left$(XX$, 4)) = NUOCO& Then
          CI0% = CVI(Mid$(XX$, 11, 2))
          CAOR = CVS(Mid$(XX$, 13, 4))
          FENTRE% = CVI(Mid$(XX$, 9, 2))
          YY$ = REGBLAN$("CARDETOC")
          Call REPLA(YY$, MKI$(CI0%), 1, 2)
          Call REPLA(YY$, Unimed$(CI0%), 3, 4)
          Call REPLA(YY$, MKS$(CAOR), 47, 4)
          Call REPLA(YY$, MKI$(FENTRE%), 67, 2)
          Call REPLA(YY$, MKS$(U&), 71, 4)
          JJ& = JJ& + 1
          Call GRAREG("!CARDETOC", YY$, JJ&)
       End If
    Next U&
    '
    Call SETULTREG("!CARDETOC", JJ&)
    Call CIERRARCH("!CARDETOC")
    Screen.MousePointer = 1
    '
    VTB% = VENTABU%("CAROCOM2/TITUPAN=P.S.Nro " + TRIM$(Str$(NUOCO&)) + " - " + TRIM$(rasocli$((-1) * NP1%)))
    If VTB% < 0 Then GoTo 99
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("OCOMPRA")
    For U& = 1 To ULTREG&("!CARDETOC")
       YY$ = REGLEIDO$("!CARDETOC", U&)
       CAOR = CVS(Mid$(YY$, 47, 4))
       REOCO& = CVS(Mid$(YY$, 71, 4))
       If REOCO& > 0 Then
          If REOCO& <= FIN& Then
             XX$ = REGLEIDO$("OCOMPRA", REOCO&)
             If CVS(Left$(XX$, 4)) = NUOCO& Then
                Call REPLA(XX$, MKI$(NP1%), 7, 2)
                Call REPLA(XX$, MKS$(CAOR), 13, 4)
                Call GRAREG("OCOMPRA", XX$, REOCO&)
             End If
          End If
       End If
    Next U&
    '
    FIN& = ULTREG&("MACONSIG")
    For U& = 1 To FIN&
       XX$ = REGLEIDO$("MACONSIG", U&)
       If XVALO(Mid$(XX$, 24, 6)) = NUOCO& Then
          Call REPLA(XX$, MKI$(NP1%), 3, 2)
          Call GRAREG("MACONSIG", XX$, U&)
       End If
    Next U&
    '
    Call CIERRARCH("*.*")
    Call FRESHECHO("!OCOMPRA")
    Screen.MousePointer = 1
    '
99  Command17.Enabled = True
    '
End Sub

Sub Command18_Click()
    '
    Command18.Enabled = False
    If DERACCE%("STPOCOM", "Configuracion de Funcionamiento Compras") = 2 Then
      SETUPCOM.Show 1
      'Call CONECRUN("FSETUP04/SETUPCOM", "Configuración de Funcionamiento")
      Call FINAL("")
    End If
    '
    Command19.Enabled = False
    FORCOMP07.Nuorig.Enabled = False
    FORCOMP07.Nuorig1.Enabled = False
    If Control$("OCOMPRA", "REPLAORI") = "SI" Then
      Command19.Enabled = True
      FORCOMP07.Nuorig.Enabled = True
      FORCOMP07.Nuorig1.Enabled = True
    End If
    '
    Command18.Enabled = True
    '
End Sub

Sub Command19_Click()
'   Command19.Enabled = False
'   Call FILESORT("REPLACOP", "", 1, 2)
'   Call FINAL("*CODORPR")
'   Command19.Enabled = True
   Command19.Enabled = False
   OPX% = COMALTER%("Opciones de Salida\\Actualización de Datos\Listado Impreso")
   If OPX% < 1 Or OPX% > 2 Then GoTo 99
   If OPX% = 2 Then GoTo 30
   
   Call FINAL("*CODORPR")
   GoTo 99
   
30 If EXISTE%(LUBAS$ + "LICORPRP" + ".PRF") < 1 Then
       Call COMUNI("Falta Controlador LICORPRP\Solicite La Instalación del Mismo a su Soporte de Sistemas")
       Exit Sub
    End If
    '
    Call FINAL("*LICORPRP")
    
99  Command19.Enabled = True

End Sub

Sub Command2_Click()
    '
    Command2.Enabled = False
    '
10  OPX% = COMALTER%("Consignación\\Por Proveedor\Por Proveedor y Orden de Compra\Cancelar")
    If OPX% < 1 Or OPX% > 2 Then
      GoTo 99
    End If
    '
    If OPX% = 1 Then
       'Por Proveedor
       Call NUENTRECONSI(0)
       GoTo 10
       
    ElseIf OPX% = 2 Then
        'Por Proveedor y Orden de Compra
20      EBOL$ = OBJPLA$("SELPROVE", "")
        NPRO% = CVI(EBOL$)
        If NPRO% < 1 Then GoTo 10
        '
        'VALIDACION QUE TENGA ORDENES DE COMPRA PENDIENTES DE CONSIGNACION.
        CONDI$ = "Cant_Asig > Cant_Entreg AND Cant_Asig >= 0.05 AND Npro_Ocom = " & NPRO%
        If CantRegistros&("MACONSIG", CONDI$, "NOMESS") < 1 Then
           Call COMUNI("Este Proveedor No Posee Consignaciones Pendientes")
           GoTo 20
        End If
        '
        'PRESENTA LISTA DE ORDENES DE COMPRA CON MATERIALES PENDIENTES DE CONSIGNACION.
        'INNER JOIN CON OCOMENCA PARA QUE NO PRESENTE CONSIGNACIONES DE O.C. ANULADAS.
        CONDI$ = "MACONSIG.Cant_Asig > MACONSIG.Cant_Entreg AND MACONSIG.Cant_Asig >= 0.05 AND MACONSIG.Npro_Ocom = " & NPRO% & " AND OCOMENCA.Stat_Ocom < 8 "
50      Call CARGALISEL("ORCOXPROV", "MACONSIG+OCOMENCA/CAST(SUBSTRING(MACONSIG.NRO_OCOM,4,6)AS INT)=OCOMENCA.Nume_Ocom", "MACONSIG.Nro_Ocom/A12;MACONSIG.Npro_Ocom/F6+proved12/A32;MACONSIG.Femi_Ocom/D8", CONDI$, "NOMESS/DISTINCT/ORDER BY MACONSIG.NPRO_OCOM")
        Call SELECHO("ORCOXPROV/Ordenes de Compras con Consignaciones Pendientes")
        NROCOMX$ = TRIM$(VALACT1$("ORCOXPROV"))
        If NROCOMX$ = "" Then GoTo 20
        Call NUENTRECONSI(NPRO%, NROCOMX$)
        GoTo 50
        '
    End If
99  Command2.Enabled = True
End Sub


Sub Command20_Click()
    '
    Command20.Enabled = False
    OPX% = COMALTER%("Ranking de Compras\\Por Proveedor\Por Artículo\Cancelar")
    If OPX% < 1 Or OPX% > 2 Then
      Command20.Enabled = True
      Exit Sub
    End If
    If OPX% = 1 Then
      Call RANKPRO
    Else
      Call RANKART
    End If
    Command20.Enabled = True
    '
End Sub

Private Sub Command21_Click()
   Command21.Enabled = False
   Screen.MousePointer = 11
   JJ& = 0
   FIN& = ULTREG&("OCOMPRA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("OCOMPRA", U&)
      '
      ST% = Asc(Mid$(XX$, 62, 1))
      If ST% < 1 Then              '    excluye cerradas
         NP% = CVI(Mid$(XX$, 7, 2))
         If NP% < 1 Or NP% = 9999 Then
            CAPEN = CVS(Mid$(XX$, 13, 4)) - CVS(Mid$(XX$, 58, 4))
            If CAPEN > 0 Then
               JJ& = JJ& + 1
               Call GRAREG("!OCOMPRA", XX$, JJ&)
            End If
         End If
      End If
   Next U&
   '
   Call SETULTREG("!OCOMPRA", JJ&)
   Call CIERRARCH("!OCOMPRA")
   '
   Screen.MousePointer = 1
   If JJ& < 1 Then
      Call MENSERR(24, "No Hay Pedidos de Suministro Pendientes.")
      GoTo 99
   End If
   '
   VTB% = VENTABU%("PEDSUPEN")
   '
99 Command21.Enabled = True
   '
End Sub

Sub Command22_Click()
    '
    Command22.Enabled = False
    Call OPLISCO07.Command22_Click
    Command22.Enabled = True
    '
End Sub

Sub Command23_Click()
   Command23.Enabled = False
   Call CONECRUN("ESTOCK02", "Análisis Estadístico de Consumos")
   Command23.Enabled = True
End Sub

Sub Command3_Click()
    Command3.Enabled = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DOSIVA") > 0 Then
        CONTROCOS.Show 1
      Else
        OPX% = COMALTER%("Opciones de Ordenamiento:\\Por Código\Por Proveedor")
        If OPX% = 1 Then
            Call OPLIMAS07.Command4_Click
          ElseIf OPX% = 2 Then
            Call OPLIMAS07.Command5_Click
        End If
    End If
    Command3.Enabled = True
End Sub

Sub Command4_Click()
   '
   Command4.Enabled = False
   If EXISTE%("ORTRAB04.EXE") > 0 Then
       Call MENSERR(24, "No Compatible con Módulo\de Ordenes de Trabajo.")
     Else
10      OPX% = COMALTER%("Opciones de Trabajo\\Marcar Items Cumplidos\Restaurar")
        If OPX% < 1 Or OPX% > 2 Then GoTo 99
        If OPX% = 1 Then
          Call ITEMCUMPLI
          GoTo 10
        ElseIf OPX% = 2 Then
          Call REVERTIR_ITEMCUMPLI
          GoTo 10
        End If
        
   End If
99 Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    TTXX$ = "Materiales Pendientes de Consignación"
    OPX% = COMALTER%(TTXX$ + "\\Todos\Por Orden de Compra")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    If OPX% = 1 Then
      Call FINAL("?MACONSIGNA")
    ElseIf OPX% = 2 Then
      Call MatePendiPorOcompra
    
    End If
99  Command5.Enabled = True
End Sub
'
Sub MatePendiPorOcompra()
    '
    Call APERBASDAT("COMPRAS")
    Call COPYSTRU("INDIOCOM", "INDIOCAC")
    CONDI$ = "Stat_Ocom < 9 AND Nume_Ocom > 0 ORDER BY Nume_Ocom"
    Call CARGALISEL("INDIOCAC", "OCOMENCA", "Nume_Ocom/F6;Refe_OcoM/A44", CONDI$, "DISTINCT")
10  Call SELECHO("INDIOCAC/Indice de Ordenes de Compra")
    '
    NRO& = XVALO(VALACT1("INDIOCAC"))
    If NRO& < 1 Then Exit Sub
    '
    NOCOM$ = TRIM$(Str$(NRO&))
    While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
    NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "%"
    '
    Call FINAL("?MACONSXOC(" & NOCOM$ & ")") '
    GoTo 10
   
End Sub
Private Sub Command6_Click()
    Command6.Enabled = False
    Call PEDISUMI
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
     '
     Command7.Enabled = False
     '
     If DERACCE%("PROVECOD", "Carga Relacion Codigo Proveedor") < 2 Then GoTo 99
     '
5    OPX% = ALTERNA%("Tabla Artículos por Proveedor\Tabla de Proveedor por Artículos\Importación Articulo por Proveedor")
     If OPX% = 3 Then
        Call IMPPROVCOD
     ElseIf OPX% = 2 Then
        Call PROVEARTIC ' TABLA DE PROVEEDOR POR ARTICULO
        GoTo 5
     ElseIf OPX% = 1 Then
         Call SELECHO("PROVED1")
         NPRO% = XVALO(VALACT1("PROVED1"))
         If NPRO% < 1 Then GoTo 5
         '
         SQLX$ = "SELECT * FROM PROVECOD WITH (NOLOCK)"
         SQLX$ = SQLX$ + "WHERE NUM_PROV = " & NPRO% & " AND MARBORRA = 0 "
         Call APERBASDAT("COMPRAS")
         Dim PROVECODTMP As ADODB.Recordset
         Set PROVECODTMP = New ADODB.Recordset
         Set PROVECODTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         Call SETULTREG("!PROVECOD", 0)
         With PROVECODTMP
            On Error Resume Next
            .MoveFirst
            If Err Then
                On Error GoTo 0
                GoTo 10
            End If
            JJ& = 0
             Do While (.EOF = False)
                X$ = REGBLAN$("PROVECOD")
                Call REPLA(X$, MKI(!cod_inte), 1, 2)
                JJ& = JJ& + 1
                Call GRAREG("!PROVECOD", X$, JJ&)
            .MoveNext
            Loop
         End With
         PROVECODTMP.Close
         Set PROVECODTMP = Nothing
         '
10       XYX% = VENTABU%("PROVECOD")
         If XYX% < 0 Then GoTo 5
         '
         CONDI$ = "NUM_PROV =" & NPRO%
         Call ACTUREGISTRO("PROVECOD", "MARBORRA", CONDI$, 1, REG&)
    
         
         SQLX$ = "SELECT * FROM PROVECOD "
         SQLX$ = SQLX$ + " WHERE NUM_PROV = " & NPRO%
         Set PROVCODTMP = New ADODB.Recordset
         PROVCODTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
           With PROVCODTMP
             For i& = 1 To ULTREG&("!PROVECOD")
               XX$ = REGLEIDO$("!PROVECOD", i&)
               Ci% = CVI(Mid$(XX$, 1, 2))
               .AddNew   ' YA CONVERTIDO
               On Error Resume Next
               !CODIEMPR = CodiEmp%
               On Error GoTo 0
               !Num_Prov = NPRO%
               !cod_inte = Ci%
               !PRECIO = 0
               !Moneda = 0
               !FECHA = CVDAT(0)
               !MARBORRA = 0
               .Update
             Next i&
           End With
         PROVCODTMP.Close
         Set PROVCODTMP = Nothing
         CONDI$ = "MARBORRA = 1"
         Call BORRAREGISTROS("PROVECOD", CONDI$, REG&)
     End If
          
99   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    TTOP$ = ""
    OPMAX% = 0
    OPCOS% = 0
    OPREPA% = 0
    OTRAC% = 0
    If InStr(EPAC$, "DOSIVA") > 0 Then
       TTOP$ = "Control Costos Reposición"
       OPCOS% = 1
       OPMAX% = 1
    End If
    '
20  If OPMAX% > 0 Then TTOP$ = TTOP$ + "\"
    OPMAX% = OPMAX% + 1
    OPREPA% = OPMAX%
    TTOP$ = TTOP$ + "Depuración de Consignaciones"
    '
    '
21  If OPMAX% > 0 Then TTOP$ = TTOP$ + "\"
    OPMAX% = OPMAX% + 1
    OCTRLRECE% = OPMAX%
    TTOP$ = TTOP$ + "Control de las Compras Recibidas vs. la Recepción de Materiales"
    '
    If TTOP$ <> "" Then
       TTOP$ = TTOP$ + "\Otros Accesos Directos"
       OPMAX% = OPMAX% + 1
       OTRAC% = OPMAX%
    End If
    '
22  If OPMAX% > 0 Then TTOP$ = TTOP$ + "\"
    OPMAX% = OPMAX% + 1
    ODECOMANU = OPMAX%
    TTOP$ = TTOP$ + "Control de las Compras Anuladas"
    '
    If TTOP$ <> "" Then
       TTOP$ = TTOP$ + "\Otros Accesos Directos"
       OPMAX% = OPMAX% + 1
       OTRAC% = OPMAX%
    End If
    '
30  OPX% = ALTERNA%(TTOP$)
    If OPX% > OPMAX% Then
         GoTo 99
       ElseIf OPX% < 1 Then
         GoTo 99
    End If
    '
    If OPX% = OPCOS% Then
         CONTROCOS.Show 1
         GoTo 30
       ElseIf OPX% = OPREPA% Then
         Call DEPUCONSIG
         GoTo 30
       ElseIf OPX% = OCTRLRECE% Then
         Call CONTROLRECIBIDOCONBR
         GoTo 30
       ElseIf OPX% = ODECOMANU Then
         Call CONTROLCOMPRASANULADAS
         GoTo 30
       ElseIf OPX% = OTRAC% Then
         Call ACCDIR("OCOMPRA")
         GoTo 30
   End If
   '
99 Command8.Enabled = True
   '
End Sub

Private Sub Command9_Click()
  Command9.Enabled = False
  'Call CUMPLECONS_ORD_CERRADAS
  Command9.Enabled = True
End Sub

'''
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Openforms = DoEvents
   Wend
End Sub

Sub PEDISUMI()             ' NUEVO PEDIDO DE SUMINISTRO
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
100 Call CIERRARCH("*.*")
    REGXX& = ULTREG&("OCOMPRA")
    NRO& = 1
    If REGXX& > 0 Then
        NRO& = 1 + CVS(Left$(REGLEIDO$("OCOMPRA", REGXX&), 4))
    End If
    If NRO& <= REGXX& Then
        Screen.MousePointer = 11
        For U& = 1 To REGXX&
            X0$ = REGLEIDO$("OCOMPRA", U&)
            If CVS(Left$(X0$, 4)) >= NRO& Then
                NRO& = 1 + CVS(Left$(X0$, 4))
            End If
        Next U&
        Screen.MousePointer = 1
    End If
    '
    X0$ = REGBLAN$("OCOMPRA")
    Call REPLA(X0$, MKS$(NRO&), 1, 4)
    Call REPLA(X0$, MKI$(HO%), 5, 2)
    Call REPLA(X0$, MKI$(HO%), 9, 2)
    Call REPLA(X0$, "S", 57, 1)
    '
101 OBJE$ = "PEDISUMI"
    '
    X1$ = OBJPLA$(OBJE$, X0$)
    If X1$ = "" Then GoTo 299
    '
    X0$ = X1$
    NRO& = CVS(Left$(X1$, 4))
    If NRO& < 1 Then GoTo 101
    '
    FEMI% = CVI(Mid$(X1$, 5, 2))
    If FEMI% < 1 Then
        Call MENSERR(24, "Falta Fecha de Emisión")
        GoTo 101
    End If
    '
    NPX$ = MKI$(0)
    NPRO% = CVI(NPX$)
    '
    Call BLANARCH("!CARDETOC")
    '
102 TITU$ = TRIM$(Str$(NRO&)) + " - Uso Interno"
    VTB% = VENTABU%("CARPESU1/TITUPAN=Pedido de Suministro Nro. " + TITU$)
    If VTB% < 0 Then GoTo 101
    '
    CAICO% = 0
    For J& = 1 To ULTREG&("!CARDETOC")
       X$ = REGLEIDO$("!CARDETOC", J&)
       CI0% = CVI(Left$(X$, 2))
       If CODEXT$(CI0%) = "" Then
          Call MENSERR(24, "Código de Artículo no Válido\en Item Número " + TRIM$(Str$(J&)) + ".")
          GoTo 102
       End If
       CANTU = CVS(Mid$(X$, 47, 4))
       If CANTU < 0.05 Then
          Call MENSERR(24, "Falta Indicar Cantidad en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          GoTo 102
       End If
       CAICO% = J&
       '
       Call REPLA(X$, Unimed$(CI0%), 3, 4)
       If CVI(Mid$(X$, 67, 2)) < HO% Then
          Call REPLA(X$, MKI$(HO%), 67, 2)
       End If
       Call REPLA(X$, "  ", 69, 2)
       '
       Call LEEEXPLO(CODEXT$(CI0%), 1)
       If CAIT% > 0 Then
          Call REPLA(X$, " *", 69, 2)
       End If
       '
       Call GRAREG("!CARDETOC", X$, J&)
       '
    Next J&
    '
    If CAICO% < 1 Then
       Call MENSERR(24, "Falta Ingresar Detalle Ordenado")
       GoTo 102
    End If
    '
103 VTB% = VENTABU%("CARPESU2/TITUPAN=Pedido de Suministro Nro. " + TITU$)
    If VTB% < 0 Then GoTo 102
    '
    TOTANE# = 0
    For J& = 1 To ULTREG&("!CARDETOC")
       X$ = REGLEIDO$("!CARDETOC", J&)
       CI0% = CVI(Left$(X$, 2))
       CANTU = CVS(Mid$(X$, 47, 4))
       FENTRE% = CVI(Mid$(X$, 67, 2))
       '
       If FENTRE% < FEMI% Then
          Call MENSERR(24, "Fecha de Entrega Incoherente\en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          Call REPLA(X0$, MKI$(FEMI%), 9, 2)
          GoTo 103
       End If
       '
    Next J&
    '
    Call CIERRARCH("!CARDETOC")
    '
    Call BLOQARCH("OCOMPRA")
    Call BLOQARCH("MACONSIG")
    '
    For J& = 1 To ULTREG&("!CARDETOC")
       X$ = REGLEIDO$("!CARDETOC", J&)
       CI0% = CVI(Left$(X$, 2))
       CANTU = CVS(Mid$(X$, 47, 4))
       PreUnid# = CVD(Mid$(X$, 51, 8))
       'PORDE = CVS(Mid$(X$, 55, 4))
       FENTRE% = CVI(Mid$(X$, 67, 2))
       CONSIMAT% = 0: If TRIM$(Mid$(X$, 69, 2)) <> "" Then CONSIMAT% = 1
       '
       X2$ = X1$
       Call REPLA(X2$, MKI$(FENTRE%), 9, 2)
       Call REPLA(X2$, MKI$(CI0%), 11, 2)
       Call REPLA(X2$, MKS$(CANTU), 13, 4)
       'Call REPLA(X2$, MKS$(0), 58, 4)
       'Call REPLA(X2$, Chr$(1 + CONSIMAT%), 57, 1) ' normal
       Call REPLA(X1$, Chr$(0), 62, 1)          ' abierta
       '
       Call REGAPP("OCOMPRA", X2$)
       '
       If CONSIMAT% = 1 Then
          '
          NOCOM$ = TRIM$(Str$(NRO&))
          While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
          ITNU$ = TRIM$(Str$(J&))
          While Len(ITNU$) < 2: ITNU$ = "0" + ITNU$: Wend
          NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-" + ITNU$
          '
          Call LEEEXPLO(CODEXT$(CI0%), CANTU)
          For i% = 1 To CAIT%
             YY$ = REGBLAN$("MACONSIG")
             Call REPLA(YY$, MKI$(FEMI%), 1, 2)
             Call REPLA(YY$, NPX$, 3, 2)
             Call REPLA(YY$, NOCOM$, 21, 12)
             '
             CII% = CIJ%(i%)
             CANTI# = CDbl(CANTU * USOI(i%))
             '
             If TRIM$(CODEXT$(CII%)) <> "" Then
               If CANTI# > 0.05 Then
                 Call REPLA(YY$, MKI$(CII%), 33, 2)
                 Call REPLA(YY$, Unimed$(CII%), 51, 4)
                 Call REPLA(YY$, MKD$(CANTI#), 55, 8)
                 Call REGAPP("MACONSIG", YY$)
               End If
             End If
          Next i%
       End If
       '
    Next J&
    '
    Call CIERRARCH("OCOMPRA")
    Call CIERRARCH("MACONSIG")
    '
' preguntar si imprime si/no
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      FECHDOC$ = FETEXTO$(FEMI%)
      NUMEDOC$ = TRIM$(Str$(NRO&))
      While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Pedido de Suministro"
      '
      CODPARAM$(1) = "COND-PAG"
      CODPARAM$(2) = "OBSERVA"
      CODPARAM$(3) = "IMP-NETO"
      CODPARAM$(4) = "POR-IVA"
      CODPARAM$(5) = "IMP-IVA"
      CODPARAM$(6) = "IMP-TOTA"
      CAPARDOC% = 6
      '
      DOCPARAM$(1) = FORPAG$
      DOCPARAM$(2) = OBSE$
      DOCPARAM$(3) = Str$(TOTANE#)
      DOCPARAM$(4) = Str$(ALIVA)
      DOCPARAM$(5) = Str$(TOTIVA#)
      DOCPARAM$(6) = Str$(TOTATO#)
      
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "UNIMED"
      CODTABU1$(4) = "CANTIDAD"
      CODTABU1$(5) = "FECH-VEN"
      CODTABU1$(6) = "REF-MAT1"
      CODTABU1$(7) = "PRE-LIST"
      CODTABU1$(8) = "PORDESCU"
      CODTABU1$(9) = "PRE-NETO"
      CODTABU1$(10) = "NET-ITEM"
      CACOLTAB1% = 10
      '
      DESTIDOC% = NC1%
      CAITAB1% = 0
      CAITCRU% = 0
      '
      For J& = 1 To ULTREG&("!CARDETOC")
         X$ = REGLEIDO$("!CARDETOC", J&)
         CI0% = CVI(Left$(X$, 2))
         CANTU = CVS(Mid$(X$, 47, 4))
         PreUnid# = CVD(Mid$(X$, 51, 8))
         'PORDE = CVS(Mid$(X$, 55, 4))
         FEX = CVI(Mid$(X$, 67, 2))
         CONSIMAT% = 0: If TRIM$(Mid$(X$, 69, 2)) <> "" Then CONSIMAT% = 1
         '
         CAITAB1% = CAITAB1% + 1
         TETABU1$(1, CAITAB1%) = CODEXT$(CI0%)
         TETABU1$(2, CAITAB1%) = DESCRIT0$(CI0%)
         TETABU1$(3, CAITAB1%) = Mid$(X$, 3, 4)
         TETABU1$(4, CAITAB1%) = TRIM$(Str$(CANTU))
         TETABU1$(5, CAITAB1%) = FECHATEX$(FEX)
         TETABU1$(6, CAITAB1%) = Mid$(X$, 69, 2)
         TETABU1$(7, CAITAB1%) = Str$(PreUnid#)
         TETABU1$(8, CAITAB1%) = Str$(PORDE)
         TETABU1$(9, CAITAB1%) = Str$(PreUnid# * (1 - PORDE / 100))
         TETABU1$(10, CAITAB1%) = Str$(CANTU * PreUnid# * (1 - PORDE / 100))
        '
109   Next J&
      '
      Call PRINTDOC("FORMOCOM")   ' Orden de Compra
      '
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
    End If
    
'consignacion inmediata si/no

    GoTo 100
    '
299 Call CIERRARCH("*.*")
    '
End Sub
'
Sub CoCoPen()
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("OCOMDETA")
   JJ& = 0
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XXX$ = REGLEIDO$("OCOMDETA", U&)
     If Asc(Mid$(XXX$, 124, 1)) < 9 Then ' ordenes abiertas
       If CVS(Mid$(XXX$, 99, 4)) < CVS(Mid$(XXX$, 55, 4)) - 0.1 Then
         XXY$ = XXX$ + String$(32, 0)
         MONEMI% = Asc(Mid$(XXX$, 122, 1))
         If MONEMI% <= 1 Then MONEX$ = "    "
         If MONEMI% = 2 Then MONEX$ = "U$S "
         If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
         TICAMBIO = 1
         If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
         '
         Call REPLA(XXY$, MONEX$, 63, 4)
         CAPEN = CVS(Mid$(XXX$, 55, 4)) - CVS(Mid$(XXX$, 99, 4))
         If CAPEN < 0.05 Then CAPEN = 0
         PreUnid# = CVD(Mid$(XXX$, 79, 8)) * TICAMBIO
         Call REPLA(XXY$, MKS$(CAPEN), 129, 4)
         Call REPLA(XXY$, MKD$(PreUnid#), 133, 8)
         Call REPLA(XXY$, MKD$(PreUnid# * CAPEN), 141, 8)
         JJ& = JJ& + 1
         Call GRAREG("OPENDETA", XXY$, JJ&)
       End If
     End If
   Next U&
   '
   Call SETULTREG("OPENDETA", JJ&)
   Call CIERRARCH("OPENDETA")
   Screen.MousePointer = 1
   '
End Sub

Sub COCOBREC()
   '
   Dim REOCO&(8192), PUBOL$(8192), CAPEN(8192)
   '
   OPX% = COMALTER%("Elija su Opción de Control:\\Cumplimiento\Calidad\Cancelar")
   If OPX% < 1 Or OPX% > 2 Then
      Exit Sub
   End If
   '
   Has% = HOY(HOS$)
   FEX = Has%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   Des% = FECHANUM(DESS$)
   VDEF$ = MKI$(Des%) + MKI$(Has%)
   '
10 OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
      Exit Sub
   End If
   Des% = CVI(Left$(OBX$, 2))
   Has% = CVI(Mid$(OBX$, 3, 2))
   If Has% < Des% Then GoTo 10
   FEX = Des%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
   FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   Screen.MousePointer = 11
   If OPX% = 2 Then GoTo 50
   '
   ' INDICE DE CUMPLIMIENTO DE PROVEEDORES
   '
   J& = 0: PUCON$ = ""
   FIN& = ULTREG&("OCOMDETA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OCOMDETA", U&)
      FF% = CVI(Mid$(X$, 5, 2))        ' fecha de emision
      NP% = CVI(Mid$(X$, 7, 2))        ' proveedor
      STAT% = Asc(Mid$(X$, 124, 1))
      '
      If STAT% <> 9 Then               ' excluye anuladas
        If NP% > 0 Then
          If NP% <> 9999 Then            ' excluye pedidos suministro
            If FF% >= Des% Then
              If FF% <= Has% Then
                J& = J& + 1
                If J& > 8192 Then
                  Call MENSERR(24, "Demasiados Movimientos\en el Rango de Fechas Elegido.")
                  GoTo 10
                End If
                '
                REOCO&(J&) = U&
                CAPEN(J&) = CVS(Mid$(X$, 55, 4))
                CLACO$ = Mid$(X$, 7, 2) + Left$(X$, 4) + Mid$(X$, 9, 2)
                PUCON$ = PUCON$ + CLACO$
              End If
            End If
          End If
        End If
      End If
   Next U&
   CAOCO& = J&
   '
   FIN& = ULTREG&("BOLRECEP")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("BOLRECEP", U&)
      CAREC = CVD(Mid$(X$, 89, 8))
      CLACO$ = Mid$(X$, 105, 2) + Mid$(X$, 151, 4) + Mid$(X$, 83, 2)
      UKA& = 0
      For k& = 1 To CAOCO&
        If Mid$(PUCON$, 8 * k& - 7, 8) = CLACO$ Then
          UKA& = k&
          If CAPEN(k&) > 0 Then
            PUBOL$(k&) = PUBOL$(k&) + MKS$(U&)
            CAPEN(k&) = CAPEN(k&) - CAREC
            GoTo 19
          End If
        End If
      Next k&
      '
      If UKA& > 0 Then
        PUBOL$(UKA&) = PUBOL$(UKA&) + MKS$(U&)
        CAPEN(UKA&) = CAPEN(UKA&) - CAREC
      End If
      '
19 Next U&
   '
   HOII = HOY(HOS$)
   For k& = 1 To CAOCO&
      X$ = REGLEIDO$("OCOMDETA", REOCO&(k&))
      Z0$ = REGBLAN$("COCOBREC")
      Call REPLA(Z0$, Mid$(X$, 7, 2), 1, 2)
      Call REPLA(Z0$, Left$(X$, 4), 3, 4)
      Call REPLA(Z0$, Mid$(X$, 9, 2), 7, 2)
      Call REPLA(Z0$, Mid$(X$, 55, 4), 9, 4)
      Call REPLA(Z0$, Mid$(X$, 95, 2), 13, 2)
      FENTREPRO = CVI(Mid$(X$, 95, 2))
      '
      If Len(PUBOL$(k&)) < 4 Then
           DIATRA% = 0
           If FENTREPRO < HOII Then
             DIATRA% = DIENTRE%(FENTREPRO, HOII)
           End If
           Call REPLA(Z0$, MKI$(DIATRA%), 37, 2)
           NCOCO& = NCOCO& + 1
           Call GRAREG("COCOBREC", Z0$, NCOCO&)
         Else
           For II& = 1 To Len(PUBOL$(k&)) Step 4
             Z$ = Z0$
             REBREC& = CVS(Mid$(PUBOL$(k&), II&, 4))
             Y$ = REGLEIDO$("BOLRECEP", REBREC&)
             Call REPLA(Z$, MKS$(CVD(Mid$(Y$, 89, 8))), 15, 4)
             Call REPLA(Z$, Mid$(Y$, 81, 2), 19, 2)
             Call REPLA(Z$, Mid$(Y$, 119, 16), 21, 16)
             FERECI = CVI(Mid$(Y$, 81, 2))
             DIATRA% = 0
             If FENTREPRO < FERECI Then
               DIATRA% = DIENTRE%(FENTREPRO, FERECI)
             End If
             Call REPLA(Z$, MKI$(DIATRA%), 37, 2)
             NCOCO& = NCOCO& + 1
             Call GRAREG("COCOBREC", Z$, NCOCO&)
           Next II&
      End If
      '
   Next k&
   '
   Call SETULTREG("COCOBREC", NCOCO&)
   Call CIERRARCH("COCOBREC")
   Call HEADERS("COCOBREC", "")
   Call HEADERS("COCOBREC", PERIO$)
   '
   Call FINAL("*SOCUPRO")
   Screen.MousePointer = 1
   Exit Sub
   '
' ************************************************************
   '
50 ' INDICE DE CALIDAD DE PROVEEDORES
   '
   VECLOT$ = ""
   FIN& = ULTREG&("BOLRECEP")
   JJ& = 0
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("BOLRECEP", U&)
      FF% = CVI(Mid$(X$, 81, 2))        ' fecha de emision
      NP% = CVI(Mid$(X$, 105, 2))        ' proveedor
      '
      If FF% >= Des% Then
        If FF% <= Has% Then
          If NP% > 0 Then
            If NP% <> 9999 Then
              CANOM = CVD(Mid$(X$, 89, 8))
              If CANOM > 0 Then
                NRO& = CVS(Left$(X$, 4))
                Ci% = CVI(Mid$(X$, 83, 2))
                REMI$ = Mid$(X$, 119, 16)
                CAPRO = CVD(Mid$(X$, 163, 8))
                LOTE$ = Mid$(X$, 69, 12)
                '
                Z$ = REGBLAN$("COCAREC")
                Call REPLA(Z$, MKI$(NP%), 1, 2)
                Call REPLA(Z$, MKS$(NRO&), 3, 4)
                Call REPLA(Z$, MKI$(FF%), 7, 2)
                Call REPLA(Z$, REMI$, 9, 16)
                Call REPLA(Z$, LOTE$, 25, 12)
                Call REPLA(Z$, MKI$(Ci%), 37, 2)
                Call REPLA(Z$, MKS$(CANOM), 39, 4)
                Call REPLA(Z$, MKI$(1), 55, 2)
                JJ& = JJ& + 1
                Call GRAREG("COCAREC", Z$, JJ&)
                VECLOT$ = VECLOT$ + LOTE$
              End If
            End If
          End If
        End If
      End If
      '
59 Next U&
   '
   FIN& = ULTREG&("MOVISTO")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     Y$ = REGLEIDO$("MOVISTO", U&)
     If Mid$(Y$, 21, 2) = "BR" Then
       CANTI = CVD(Mid$(Y$, 88, 8)) - CVD(Mid$(Y$, 96, 8))
       If CANTI >= 0.05 Then
         LOTE$ = Mid$(Y$, 5, 12)
         If TRIM$(LOTE$) <> "" Then
           For KK& = 1 To Len(VECLOT$) Step 12
             If Mid$(VECLOT$, KK&, 12) = LOTE$ Then
               REG& = (KK& + 11) / 12
               If REG& > 0 Then
                 If REG& <= JJ& Then
                   Z$ = REGLEIDO$("COCAREC", REG&)
                   Call REPLA(Z$, MKS$(CANTI), 43, 4)
                   Call GRAREG("COCAREC", Z$, REG&)
                   GoTo 69
                 End If
               End If
             End If
           Next KK&
         End If
       End If
     End If
     '
     REFE$ = UCase$(Mid$(Y$, 43, 30))
     If Left$(REFE$, 7) = "RECHAZO" Then
       CANTI = CVD(Mid$(Y$, 96, 8)) - CVD(Mid$(Y$, 88, 8))
       If CANTI >= 0.05 Then
         LOTE$ = Mid$(Y$, 5, 12)
         If TRIM$(LOTE$) <> "" Then
           For KK& = 1 To Len(VECLOT$) Step 12
             If Mid$(VECLOT$, KK&, 12) = LOTE$ Then
               REG& = (KK& + 11) / 12
               If REG& > 0 Then
                 If REG& <= JJ& Then
                   Z$ = REGLEIDO$("COCAREC", REG&)
                   Call REPLA(Z$, MKS$(CANTI), 47, 4)
                   Call REPLA(Z$, MKI$(1), 57, 2)
                   Call GRAREG("COCAREC", Z$, REG&)
                   GoTo 69
                 End If
               End If
             End If
           Next KK&
         End If
       End If
     End If
     '
     If Left$(REFE$, 8) = "FALTANTE" Then
       CANTI = CVD(Mid$(Y$, 96, 8)) - CVD(Mid$(Y$, 88, 8))
       If CANTI >= 0.05 Then
         LOTE$ = Mid$(Y$, 5, 12)
         If TRIM$(LOTE$) <> "" Then
           For KK& = 1 To Len(VECLOT$) Step 12
             If Mid$(VECLOT$, KK&, 12) = LOTE$ Then
               REG& = (KK& + 11) / 12
               If REG& > 0 Then
                 If REG& <= JJ& Then
                   Z$ = REGLEIDO$("COCAREC", REG&)
                   Call REPLA(Z$, MKS$(CANTI), 51, 4)
                   'Call REPLA(Z$, MKI$(1), 57, 2)
                   Call GRAREG("COCAREC", Z$, REG&)
                   GoTo 69
                 End If
               End If
             End If
           Next KK&
         End If
       End If
     End If
     '
69 Next U&
   '
   Call SETULTREG("COCAREC", JJ&)
   Call CIERRARCH("COCAREC")
   Call CIERRARCH("*.*")
   Call HEADERS("COCAREC", "")
   Call HEADERS("COCAREC", PERIO$)
   '
   Call FINAL("*SOCAREC")
   Screen.MousePointer = 1
   '
End Sub

Private Sub Form_Load()
    '
    Command19.Enabled = False
    If Control$("OCOMPRA", "REPLAORI") = "SI" Then
      Command19.Enabled = True
    End If
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    'If InStr(EPAC$, "FERVI") > 0 Then Command9.Visible = True
    
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me, Picture5)
    '
    If EXISTE%("ESTOCK02.EXE") > 0 Then
      Command23.Enabled = True
    End If
    '
    Command7.Enabled = False    'Si Val Cod-Prov esta desactivado el botón Co-Pr queda deshabilitado
    If Control$("OCOMPRA", "VALCOPRO") = "SI" Then
        OPCICOM07.Command7.Enabled = True
    End If
    '
    If Control("OCOMPRA", "VALISTPE") <> "SI" Then
        RecibeArchStockMYM.Enabled = False
        RecibirExcepcion.Enabled = False
        RECIBOPEDICOMPRA.Enabled = False
    End If
    
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Sub MODIOCOM()         ' MODIFICAR ORDEN DE COMPRA
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    FORIPRE$ = TRIM$(Control$("", "FORMOCOM"))
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Compra:\  \'" + RMCC$ + "'.")
            Exit Sub
        End If
    End If
    '
    AVALIDA$ = "MASTER"
    '
    'Cargo la lista con las O/C
    Call APERBASDAT("COMPRAS")
    '
    If CODCARPE$ <> "" Then
        CONDI$ = "Stat_Ocom < 9 "  ' EXCLUYE ANULADAS
        CONDI$ = CONDI$ + " AND Nume_Ocom > 0 "
        CONDI$ = CONDI$ + " AND CODCARPIMP = '" & CODCARPE$ & "'"
        CONDI$ = CONDI$ + " ORDER BY Nume_Ocom "
        ' Lista de seleccion asociada a la Carpeta de Importacion
        Call CARGALISEL("!INDIOCOM", "OCOMENCA", "Nume_Ocom/F6;Refe_Ocom/A44", CONDI$, "NOMESS")
        '
        Call SELECHO("!INDIOCOM/Indice General de Ordenes de Compra")
    Else
        CONDI$ = "Stat_Ocom < 9 "  ' EXCLUYE ANULADAS
        CONDI$ = CONDI$ + " AND Nume_Ocom > 0 "
        CONDI$ = CONDI$ + " ORDER BY Nume_Ocom "
        Call CARGALISEL("!INDIOCOM", "OCOMENCA", "Nume_Ocom/F6;Refe_Ocom/A44", CONDI$)
        '
3       Call SELECHO("!INDIOCOM/Indice de Ordenes de Compra (Pendientes-Cumplidas)")
    End If
    '
    NUOCO& = XVALO(VALACT1$("!INDIOCOM")) ' ACA LEO DEL INDIOCOM
    If NUOCO& < 1 Then GoTo 999
    '
    If TIENECOTIZA%(NUOCO&) > 0 Then
       Call COMUNI("Imposible Modificar Esta Orden de Compra\Artículos con Solicitud de Cotización Realizada")
       GoTo 3
    End If
    '
    SQLX$ = "Select * from Ocomenca where Nume_Ocom =" & NUOCO&
    SQLX$ = SQLX$ + " and Stat_Ocom < 9 "
    Call APERBASDAT("COMPRAS")
    Dim OCOMENCA As ADODB.Recordset
    Set OCOMENCA = New ADODB.Recordset
    OCOMENCA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    With OCOMENCA
      If (.EOF And .BOF) Then
        Call MENSERR(24, "Imposible Abrir la O/Compra " & NUOCO& & " - Consulte !")
        GoTo 3
      End If
      '
      If !Stat_Ocom > 8 Then ' esta anulada
        Call MENSERR(24, "O/Compra Anulada - Imposible Modificar")
        GoTo 3
      End If
      '
      NPRO% = !Npro_Ocom
      NC1% = (-1) * NPRO%
      PVCL% = PIVACLI%(NC1%)
      If PVCL% = 2 Or PVCL% = 4 Or PVCL% = 5 Then
        ALIVX = 0
      End If
      '
      FECHEMI% = CVINT%(!Femi_Ocom)
      FENTREGEN% = 32767
      NRO& = NUOCO&
      CPAG% = !Cpag_Ocom
      EMBAL% = 0
      On Error Resume Next
      EMBAL% = !Embala_Ocom
      On Error GoTo 0
      FORPAG$ = TRIM$(ECOARCH$("CONPAG", Chr$(CPAG%)))
      If FORPAG$ = "" Then FORPAG$ = "Contado Contra Entrega"
      ATEN$ = !AtSr_Ocom: If TRIM$(ATEN$) = "" Then ATEN$ = CONTACTCLI$(NC1%)
      '
      MONEMI% = !Mone_Emis
      If MONEMI% < 1 Or ECOARCH$("TAMONED", Chr$(MONEMI%)) = "" Then MONEMI% = 1
      MONECO% = !MONE_COS
      If MONECO% < 1 Or ECOARCH$("TAMONED", Chr$(MONECO%)) = "" Then MONECO% = 1
      '
      FOMODOC07.Label11 = DEMONECO$(MONEMI%)
      FOMODOC07.TICAMBIS = TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4"))
      If MONEMI% > 1 Then
         FOMODOC07.Label11.ForeColor = RGB(255, 0, 0)
      End If
      FOMODOC07.LText1 = NPRO%
      FOMODOC07.Text6 = ATEN$
      FOMODOC07.TEXT21(0) = !Cpag_Ocom
      FOMODOC07.TEXT21(1) = EMBAL%
      FOMODOC07.Text10 = !Lentrega
      FOMODOC07.Text13 = Obse_Ocom
      FOMODOC07.Text7 = ""    ' destino
      FOMODOC07.Text18 = ""   ' sector
      FOMODOC07.Text14 = ""   ' fecha de entrega
      FOMODOC07.Text15 = !MONE_COS
      FOMODOC07.Text11 = NUOCO&
      FOMODOC07.Text9 = FECHATEX$(FECHEMI%)
      FOMODOC07.PORDESCU = !Pdes_Ocom
      FOMODOC07.Total(4) = FORMATNUM(!Iotr1_Ocom, "F10.2")
      '\\\OT-08-0969-OD-24/10/08/// CARGA DE LA CALIFICACION DEL PROVEEDOR
      If EXISTE%(LUDAT$ + "CALINTEG.RFS") > 0 Then
          XX11$ = FILTERGETRECORD$("CALINTEG", 1, MKI$(NPRO%))
          PUNTAPRO = CVS(Mid$(XX11$, 35, 4))
          FOMODOC07.Label18 = FORMATNUM$(PUNTAPRO, "F4.1")
      End If
     '\\\OT-08-0969-OD-FIN///

      '\\\OT-08-0887-OD-24/09/08///
      FOMODOC07.Label118 = SAFETEXT(!PorcIva) ' PORCENTAJE DE IVA PARA ESTA COMPRA
      If FOMODOC07.Label118 = "" Then
         FOMODOC07.Label118 = FORMATNUM$(ALIVA, "F6.2") ' POR SI EL CAMPO ESTA VACIO O NULL
      End If
      '\\\OT-08-0887-OD-FIN///
        PORPER = 0
        On Error Resume Next
        PORPER = 100 * XVALO(FOMODOC07.Total(4)) / (!ToBru_Ocom - !IDes_Ocom)
        On Error GoTo 0
        If PORPER > 99.9 Then PORPER = 99.9
      FOMODOC07.Text17 = TRIM$(FORMATNUM(PORPER, "F5.1"))
      FOMODOC07.Text13 = SAFETEXT$(!Obse_Ocom)
      FOMODOC07.Text22 = SAFETEXT$(!Obse_Ocom)
      On Error Resume Next
      FOMODOC07.Label13 = !Anex_Ocom
      FOMODOC07.Label15 = TRIM$(Str$(!Stat_Ocom))
      '\\\OT-06-0082-WAR-06/03/06///
      NROREV% = 0
      On Error Resume Next
      NROREV% = XVALO(!OCOMREVIS)
        If !Stat_Ocom >= 0 Then 'si no esta aprobada no cambia revisión(0).
          NROREV% = NROREV% + 1
        End If
      On Error GoTo 0
      FOMODOC07.Label16 = TRIM$(Str$(NROREV%)) 'FORMATNUM$(NROREV1%, "F2.0") LO COMENTE XQ DEVUELVE "" SI NUMERO =0
      FOMODOC07.Text1 = SAFETEXT$(!Respon_Ocompra)
      FOMODOC07.TEXT25 = SAFETEXT$(!OComCliente)
      FOMODOC07.TEXT24 = XVALO(!nume_cli)
      If XVALO(FOMODOC07.TEXT24) < 1 Then FOMODOC07.Option1.Value = True
      
      '\\\OT-06-0082-WAR-FIN///
      On Error GoTo 0
      '
    End With
    OCOMENCA.Close
    Set OCOMENCA = Nothing
    '
    Screen.MousePointer = 11
    REDETA$ = ""
    ARCHILIB$ = "!1OCOMDET"
    Call SETULTREG("!OCOMDETA", 0)
    Call SETULTREG(ARCHILIB$, 0)
    '
    REBLA$ = REGBLAN$("OCOMDETA")
    Call REPLA(REBLA$, MKS$(NUOCO&), 1, 4)
    Call REPLA(REBLA$, MKI$(FECHEMI%), 5, 2)
    Call REPLA(REBLA$, MKI$(NPRO%), 7, 2)
    Call REPLA(REBLA$, Chr$(MONECO%), 121, 1)
    Call REPLA(REBLA$, Chr$(MONEMI%), 122, 1)
      TICAMBIO = 1
      If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
    '
    SQLX$ = "SELECT * FROM OCOMDETA WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE Nume_Ocom = " & NUOCO& & " "
    SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
    SQLX$ = SQLX$ + "ORDER BY NuOrd_Item ASC"
    '
    Dim COMDET As ADODB.Recordset
    Set COMDET = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With COMDET
       If .EOF Then
        Call MENSERR(24, "Imposible cargar Detalle de O/Compra\Orden de Compra Anulada")
        Screen.MousePointer = 1
        GoTo 3
      End If
      Do While Not .EOF
        '
        XXY$ = REBLA$
        CAPEN = !Cant_Ocom - !Cant_Rec
        If CAPEN < 0.05 Then CAPEN = 0
        PreUnid# = !prun_neto * TICAMBIO
        PEDINTERNO$ = SAFETEXT$(!Medi_Ocom)
        '
        CIXI% = !cod_inte
        Call REPLA(XXY$, MKI$(CIXI%), 9, 2)
        If CIXI% <= 0 Then HAYMATERIALIB% = 1 'HAY MATERIAL TEXTO LIBRE
        On Error Resume Next
        UNIMA$ = TRIM$(!Uni_Mas)
        On Error GoTo 0
        If UNIMA$ = "" Then UNIMA$ = TRIM$(Unimed$(CIXI%))
        Call REPLA(XXY$, UNIMA$, 11, 4)
        Call REPLA(XXY$, PEDINTERNO$, 31, 20)
        '\\\OT-05-0693-WAR-24/10/05///
          If IsNull(!Sect_Ocom) Then
            SECT1$ = ""
          Else
            SECT1$ = !Sect_Ocom
          End If
          Call REPLA(XXY$, SECT1$, 15, 16)
            CCU% = !IContab
            CCU1$ = CODCUE$(CCU%)
          Call REPLA(XXY$, CCU1$, 105, 12)
        '\\\OT-05-0693-WAR-FIN///
        Call REPLA(XXY$, MKS$(!Cant_Ocom), 55, 4)
        If IsNull(!Uni_Com) Then
          UNICOMOC$ = "  "
        Else
          UNICOMOC$ = !Uni_Com
        End If
        Call REPLA(XXY$, UNICOMOC$, 59, 4)
        If TRIM$(Mid$(XXY$, 11, 4)) = "" Then Call REPLA(XXY$, UNICOMOC$, 11, 4)
        '
        Call REPLA(XXY$, MKS$(!Coef_Unids), 63, 4)
        Call REPLA(XXY$, MKD$(!Pre_Unit), 67, 8)
        Call REPLA(XXY$, MKS$(!Desc_Item), 75, 4)
        Call REPLA(XXY$, MKD$(!prun_neto), 79, 8)
        Call REPLA(XXY$, MKD$(!Cant_Ocom * !prun_neto), 87, 8)
        Call REPLA(XXY$, MKI$(CVINT%(!Fentre_Sol)), 95, 2)
        FENTREPAR% = CVINT(!Fentre_Sol)
        If FENTREPAR% > 0 Then
          If FENTREPAR% < FENTREGEN% Then
            FENTREGEN% = FENTREPAR%
          End If
        End If
        '
        If IsNull(!CONSIMAT) Then
          CONSIMATOC$ = "  "
        Else
          CONSIMATOC$ = !CONSIMAT
        End If
        Call REPLA(XXY$, CONSIMATOC$, 97, 2)
        Call REPLA(XXY$, MKS$(!Cant_Rec), 99, 4)
          STATUI% = !Stat_Ocom: While STATUI% < 0: STATUI% = STATUI% + 256: Wend
        Call REPLA(XXY$, Chr$(STATUI%), 124, 1)
        Call REPLA(XXY$, MKS$(!NuOrd_Item), 125, 4)
        '
          NUECUEIMPO$ = CODCUE$(!IContab)
        Call REPLA(XXY$, NUECUEIMPO$, 105, 12)
          'SECTOROC$ = !Sect_Ocom '***' ESTA LINEA NO ESTABA ATAJADA
          'Call REPLA(XXY$, SECTOROC$, 15, 12)
          '
        Call REGAPP("!OCOMDETA", XXY$)
          '
          DELIBX$ = TRIM$(!DELIBRE)
          If DELIBX$ = "" Then DELIBX$ = DESCRIT0$(CIXI%)
        Call REGAPP(ARCHILIB$, DELIBX$)
        '
        FOMODOC07.Text7 = CODCUE$(!IContab)
        On Error Resume Next
        FOMODOC07.Text18 = !Sect_Ocom
        On Error GoTo 0
        '
      .MoveNext
      Loop
    End With
    COMDET.Close
    Set COMDET = Nothing
    '
    If FENTREGEN% = 32767 Or FENTREGEN% < FECHEMI% Then FENTREGEN% = FECHEMI%
    FOMODOC07.Text14 = FECHATEX$(FENTREGEN%)
    '
    Call FOMODOC07.CARCOANT
    Call FOMODOC07.CARLISCO
    Call FOMODOC07.CALTOCOM
    '
    ' AQUI CARGAR RIVACOM
    TOTANEX = XVALO(FOMODOC07.Total(0)) - XVALO(FOMODOC07.Total(1))
    RIVACOM07.TONECOM = TRIM$(FORMATNUM$(TOTANEX, "F15.2"))
    SQLX$ = "SELECT * FROM DETAIVA "
    SQLX$ = SQLX$ + "WHERE Nume_Ocom = " & NUOCO& & " "
    SQLX$ = SQLX$ + "ORDER BY NORIT ASC"
    Set DETAIVA = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With DETAIVA
      If Not (.BOF And .EOF) Then
        For KKII% = 0 To 5
           RIVACOM07.IMPOGRA(KKII%).TEXT = ""
           RIVACOM07.IMPOIVA(KKII%).Caption = ""
        Next KKII%
        Do While Not .EOF
          KKII% = XVALO(!NORIT)
          If KKII% >= 0 And KKII% <= 5 Then
             RIVACOM07.TASIVA(KKII%) = SAFETEXT$(!Tasa_Iva)
             RIVACOM07.IMPOGRA(KKII%) = TRIM$(FORMATNUM$(XVALO(!Base_Impo), "F15.2"))
             RIVACOM07.IMPOIVA(KKII%) = TRIM$(FORMATNUM$(XVALO(!Impo_Iva), "F15.2"))
          End If
        .MoveNext
        Loop
        FOMODOC07.Label118 = TRIM$(RIVACOM07.Label2)
        FOMODOC07.Total(2) = TRIM$(RIVACOM07.Label3)
      End If
    End With
    DETAIVA.Close
    Set DETAIVA = Nothing
    
    ' EN PROVEEDORES DEL EXTERIOR EL SISTEMA GUARDA EN LA TABLA DETAIVA UN REGISTRO CON EL IVA DEL 21%
    ' AUNQUE SE TRATE DE UNA COMPRA EXENTA DE IVA. LA RUTINA CALTOCOM RECALCULA LOS TOTALES Y QUITA
    ' EL IMPORTE DE IVA CUANDO DETECTA QUE LA POSICION DE IVA ES 5 (COMERCIO EXTERIOR)
    Call FOMODOC07.CALTOCOM
    '
    FOMODOC07.Show 1
    GoTo 3
    '
999 Screen.MousePointer = 1
    '
End Sub
'
Sub ACTCOSTANSER(CI0%, NOPERX%, NPRO%, PREUNINE#, MONEMI%, FECHEMI%)
   RFF$ = RECFILT$("COSOPTER", 1, MKI$(CI0%))
   For UK& = 1 To Len(RFF$) Step 4
     RECOSO& = CVS(Mid$(RFF$, UK&, 4))
     If RECOSO& > 0 Then
       If RECOSO& <= ULTREG&("COSOPTER") Then
         ZXC$ = REGLEIDO$("COSOPTER", RECOSO&)
         If CVI(Mid$(ZXC$, 3, 2)) = NOPERX% Then GoTo 20
       End If
     End If
   Next UK&
   '
   RECOSO& = 1 + ULTREG&("COSOPTER")
   ZXC$ = String$(32, 0)
   Call REPLA(ZXC$, MKI$(CI0%), 1, 2)
   Call REPLA(ZXC$, MKI$(NOPERX%), 3, 2)
   '
20 NPROX% = CVI(Mid$(ZXC$, 5, 2))
   If NPROX% = 0 Or NPROX% = NPRO% Then
     If PREUNINE# >= 0.00005 Then
       MONEI% = MONEMI%: If MONEI% < 1 Then MONEI% = 1
       Call REPLA(ZXC$, MKI$(NPRO%), 5, 2)
       Call REPLA(ZXC$, MKD$(PREUNINE#), 7, 8)
       Call REPLA(ZXC$, MKI$(MONEMI%), 15, 2)
       Call REPLA(ZXC$, MKI$(FECHEMI%), 17, 2)
       Call GRAREG("COSOPTER", ZXC$, RECOSO&)
       Call CIERRARCH("COSOPTER")
     End If
   End If
   '
End Sub

Sub OCOMANU()
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))

    If DERACCE%("ANULOCOM", "Anulación de Ordenes de Compra") < 2 Then Exit Sub

    If COMALTER%("Realmente Desea\ANULAR Ordenes de Compra ?\\No, Abandonar\Si, Anular O/C's") <> 2 Then
      Exit Sub
    End If
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    TITU$ = "PROCESO DE ANULACION DE ORDENES DE COMPRA"
    '
    'Cargo en el list las O/C que puedan ser Anuladas.
5   Screen.MousePointer = 11
    Call APERBASDAT("COMPRAS")
    SQLX$ = " SELECT  DISTINCT OCOMENCA.Nume_Ocom, OCOMENCA.Refe_Ocom "
    SQLX$ = SQLX$ + " FROM OCOMENCA "
    SQLX$ = SQLX$ + " INNER JOIN  OCOMDETA "
    SQLX$ = SQLX$ + " ON (OCOMENCA.Nume_Ocom = OCOMDETA.Nume_Ocom) "
    SQLX$ = SQLX$ + " WHERE OCOMDETA.Stat_Ocom <8 " ' Excluyo las que estan anuladas O COMPLIDAS
    SQLX$ = SQLX$ + " AND OCOMDETA.Cant_Rec < OCOMDETA.Cant_Ocom "
    SQLX$ = SQLX$ + " AND OCOMENCA.Stat_Ocom < 1 "
    SQLX$ = SQLX$ & " AND OCOMENCA.Nume_Ocom > 0 "
    SQLX$ = SQLX$ + " AND OCOMENCA.CODIEMPR = " & CodiEmp%
    SQLX$ = SQLX$ + " AND OCOMDETA.CODIEMPR = " & CodiEmp%
    SQLX$ = SQLX$ + " ORDER BY OCOMENCA.Nume_Ocom  ASC "
    'Set OCANULTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim OCANULTEMP As ADODB.Recordset
    Set OCANULTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With OCANULTEMP
        On Error Resume Next
        .MoveFirst
        If Err Then
          Call MENSERR(24, "No hay Ordenes de Compra\Pendientes que Puedan Anularse.")
          Screen.MousePointer = 1
          On Error GoTo 0
          Exit Sub
        End If
        On Error GoTo 0
        J& = 0
        Do While (.EOF = False)
            XXY$ = REGBLAN$("INDIOCPE")
            Call REPLA(XXY$, MKS(!nume_ocom), 1, 4) 'Nº O/C
            Call REPLA(XXY$, !Refe_OcoM, 5, 44) ' REFERENCIA
            J& = J& + 1
            Call GRAREG("INDIOCPE", XXY$, J&)
        .MoveNext
        Loop
    End With
    OCANULTEMP.Close
    Set OCANULTEMP = Nothing
    '
    Call SETULTREG("INDIOCPE", J&)
    Call CIERRARCH("INDIOCPE")
    Call BAJALDISCO
    Call FRESHECHO("INDIOCPE")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
      Call MENSERR(24, "No hay Ordenes de Compra\Pendientes que Puedan Anularse.")
      Screen.MousePointer = 1
      GoTo 999
    End If
    '
    INDI% = 1
    Screen.MousePointer = 1
    '
10  YY$ = REGSEL$("INDIOCPE/Indice de Ordenes de Compra Pendientes para Anulación")
    If Len(YY$) <= 4 Then GoTo 999
    '
    NUOCO& = CVS(Left$(YY$, 4))
    If TIENECOTIZARECIB%(NUOCO&) > 0 Then
       Call COMUNI("Imposible Anular Esta Orden de Compra\Artículos con Solicitud de Cotización Realizada")
       GoTo 10
    End If
    '
    '
    If InStr(EPAC$, "AHP") > 0 Then
        If TieneItemsRecibidos&(NUOCO&) > 0 Then
           Call COMUNI("Atención !! Imposible Anular O.Compra yá Posee Cantidades Recibidas")
           GoTo 10
        End If
    End If
    
    'SOLICITUD DE CONFIRMACION DE ANULACION DE ORDEN DE COMPRA
    If COMALTER%("¿Está Seguro Que Desea Anular la O/Compra N° " & NUOCO& & "?\\No, Cancelar\Si, Anular ") <> 2 Then
      GoTo 10
    End If
    '
    CONDI$ = "NUME_OCOM = " & NUOCO&
    ORDSER& = XVALO(VALOBADA("OCOMDETA", "N_ORSERV", CONDI$))
    OF$ = SAFETEXT$(VALOBADA("OCOMDETA", "DOC_ORIG", CONDI$))

    MODOGABAL% = 0: If InStr(EPAC$, "GABAL") > 0 And ORDSER& > 0 Then MODOGABAL% = 1
    If MODOGABAL% > 0 Then
        Call COMUNI("Imposible Anular Orden de Compra \ Generada a Partir de O.Fab.: " & OF$)
        GoTo 10
    End If
    '
    Call ANULACOTIZA(NUOCO&, RTA_ANULACOT%) ' ANULA COTIZACION RELACIONADA
    If RTA_ANULACOT% > 0 Then
       GoTo 10
    End If
    '
    'ACTUALIZA OCOMENCA
    STRSQL = "Update OCOMENCA set Stat_Ocom = 9,"
    STRSQL = STRSQL + " Refe_ocom ='(anul) ' + left(Refe_OcoM,57)"
    STRSQL = STRSQL + " where Stat_Ocom < 8"
    STRSQL = STRSQL + " And Nume_Ocom = " & NUOCO&
    FLEXCONN.Execute (FILTSQL$(STRSQL))
    '
    'SE AGREGA ESTO PARA QUE GRABE CON STATUS 10 ANTES DE MARCAR EN 9 LAS QUE ESTAN CON 0
    'POR QUE ANTES MARCABA TODOS LOS REGISTROS CON 9 Y EN EL CASO QUE SI SE HABIA
    'MODIFICADO LA ORDEN Y GENERO NUEVOS REGISTROS Y GRABO LOS ANTERIOES EN 9
    'SI LA ORDEN ANULADA SE REVERTIA LES PONIA A TODAS EN 0 QUEDANADO LOS CODIGOS DUPLIDADOS
    CONDI$ = "Stat_Ocom = 9 AND  Nume_Ocom=" & NUOCO&
    Call ACTUREGISTRO("OCOMDETA", "STAT_OCOM", CONDI$, 10, REGAF&, "NOMESS")
    'FIN CAMBIO
    
    'jandy aca actualizo en ocomdeta
    STRSQL = STRSQL + " Update OCOMDETA set Stat_Ocom = 9"
    STRSQL = STRSQL + " Where Stat_ocom < 8"
    STRSQL = STRSQL + " And Nume_Ocom=" & NUOCO&
    FLEXCONN.Execute (FILTSQL$(STRSQL))
    '
    Call ANULAORSER(NUOCO&) 'Anula O/Serv. Asociada
    '
    ' *** DESVINCULA DE LA COTIZACION, DEJANDO EN CERO EL CAMPO DE NRO DE ORDEN
    STRSQL = STRSQL + " UPDATE COTIRECI SET nume_ocom='0'"
    STRSQL = STRSQL + " WHERE nume_ocom='" & CStr(NUOCO&) & "'"
    FLEXCONN.Execute (FILTSQL$(STRSQL))
    '
    Screen.MousePointer = 11
    '
    J& = 0
    For i& = 1 To ULTREG&("INDIOCPE")
      XX$ = REGLEIDO$("INDIOCPE", i&)
      If CVS(Left$(XX$, 4)) <> NUOCO& Then
        J& = J& + 1
        Call GRAREG("INDIOCPE", XX$, J&)
      End If
    Next i&
    Call SETULTREG("INDIOCPE", J&)
    '
    NOCOM1$ = TRIM$(Str$(NUOCO&))
      While Len(NOCOM1$) < 6: NOCOM1$ = "0" + NOCOM1$: Wend
    ITNU1$ = TRIM$(Str$(NUORIT%))
      While Len(ITNU1$) < 2: ITNU1$ = "0" + ITNU1$: Wend
    NOCOBUS$ = "OC-" + Left$(NOCOM1$, 6)

    ' CANCELAR RESERVAS
    ' BORRA LOS REGISTROS EN LOS QUE NO SE ENTREGO NADA
'    STRSQL = "DELETE MACONSIG where Nro_Ocom = '" & NOCOBUS$ & "' AND Cant_Entreg < 0.05"
'     FLEXCONN.Execute (FILTSQL$(STRSQL))
    CONDI$ = "LEFT(Nro_Ocom,9) = '" & NOCOBUS$ & "' AND Cant_Entreg < 0.05"
    Call BORRAREGISTROS("MACONSIG", CONDI$, REGAF&, "NOMESS")
    '
    ' PONE LA CANTIDAD ASIGNADA IGUAL A LA ENTREGADA PARA NO DESPACHAR MÁS
    STRSQL = "update MACONSIG set Cant_Asig = Cant_Entreg where LEFT(Nro_Ocom,9)='" & NOCOBUS$ & "'"
    FLEXCONN.Execute (FILTSQL$(STRSQL))
    
    'SI QUEDAN REGISTROS CON CANTIDAD ENTREGADA PRESENTAR MENSAJE.
    CONDI$ = "LEFT(Nro_Ocom,9) = '" & NOCOBUS$ & "' AND (Cant_Entreg - Cant_Rendid)>0.05 "
    CANTENTR& = CantRegistros("MACONSIG", CONDI$)
    If CANTENTR& > 0 Then
       Call COMUNI("Hay Materiales Yá Entregados en Consignación\ Al Proveedor para esta O/Compra.\Recuerde Producir el RE-Ingreso")
    End If
    Call COMUNI("Transacción Completada")
    '
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call FRESHECHO("!INDIOCOM")
    Screen.MousePointer = 1
    GoTo 5
    '
999 Call CIERRARCH("*.*")
    '
End Sub
'
Sub CONVERCOM()
   '
   ' conversion de formato de ordenes de compra
   '
   If ULTREG&("OCOMPRA") > 0 Then
     '
     Call BLOQARCH("OCOMPRA")
     Call BLOQARCH("OCOMENCA")
     Call BLOQARCH("OCOMDETA")
     '
     Screen.MousePointer = 11
     OCOA& = 0: NPROA% = 0: ENCAO$ = ""
     FIN& = ULTREG&("OCOMPRA")
     '
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("OCOMPRA", U&)
       NUOCO& = CVS(Left$(XX$, 4))
       FEMI% = CVI(Mid$(XX$, 5, 2))
       NPRO% = CVI(Mid$(XX$, 7, 2))
       FENTRE% = CVI(Mid$(XX$, 9, 2))
       CI0% = CVI(Mid$(XX$, 11, 2))
       CAOR = CVS(Mid$(XX$, 13, 4))
       OBSER$ = Mid$(XX$, 17, 30)
       PreUnid# = CVS(Mid$(XX$, 49, 4))
       DESCUNI = CVS(Mid$(XX$, 53, 4))
       TISU% = Asc(Mid$(XX$, 57, 1))
       CARE = CVS(Mid$(XX$, 58, 4))
       STAT% = Asc(Mid$(XX$, 62, 1))
       '
       If NUOCO& <> OCOA& Or NPRO% <> NPROA% Then
         If OCOA& <> 0 Then
           TOTIVA# = 0
           PVCL% = PIVACLI%((-1) * NPRO%)
           If PVCL% <> 2 Then
             If PVCL% <> 4 Then
               TOTIVA# = (Int(ALIVA * TOTANE# + 0.5)) / 100
             End If
           End If
           Call REPLA(ENCAO$, MKD$(TOTANE#), 89, 8)
           Call REPLA(ENCAO$, MKD$(TOTIVA#), 121, 8)
           Call REPLA(ENCAO$, MKD$(TOTANE# + TOTIVA#), 145, 8)
           Call REGAPP("OCOMENCA", ENCAO$)
         End If
         '
         OCOA& = NUOCO&: NPROA% = NPRO%
         ENCAO$ = REGBLAN$("OCOMENCA")
         '
         Call REPLA(ENCAO$, MKS$(NUOCO&), 1, 4)
         FEX = FEMI%
         REFE$ = FECHATEX$(FEX) + " - " + rasocli$((-1) * NPRO%)
         Call REPLA(ENCAO$, REFE$, 5, 44)
         Call REPLA(ENCAO$, MKI$(FEMI%), 49, 2)
         Call REPLA(ENCAO$, MKI$(NPRO%), 51, 2)
         Call REPLA(ENCAO$, Chr$(TISU%), 53, 1)
         Call REPLA(ENCAO$, Chr$(STAT%), 54, 1)
         '
         HO% = HOY(HOS$)
         HP = 100 * XVALO(Left$(Time$, 2)) + XVALO(Mid$(Time$, 4, 2))
         Call REPLA(ENCAO$, MKI$(HO%), 55, 2)
         Call REPLA(ENCAO$, MKI$(Int(HP)), 57, 2)
         Call REPLA(ENCAO$, MKI$(USNBR%), 59, 2)
         '
         If STAT% = 9 Then
           Call REPLA(ENCAO$, MKI$(HO%), 67, 2)
           Call REPLA(ENCAO$, MKI$(Int(HP)), 69, 2)
           Call REPLA(ENCAO$, MKI$(USNBR%), 71, 2)
           Call REPLA(ENCAO$, "Conversion BD", 73, 16)
         End If
         '
         CPAG% = CPAGCLI%((-1) * NPRO%)
         Call REPLA(ENCAO$, MKI$(CPAG%), 153, 2)
         REDETA& = 1 + ULTREG&("OCOMDETA")
         Call REPLA(ENCAO$, MKS$(REDETA&), 155, 4)
         Call REPLA(ENCAO$, MKS$(REDETA&), 159, 4)
         '
         Call REPLA(ENCAO$, OBSER$, 193, 64)
         '
       End If
       '
       REDET$ = REGBLAN$("OCOMDETA")
       Call REPLA(REDET$, MKS$(NUOCO&), 1, 4)
       Call REPLA(REDET$, MKI$(FEMI%), 5, 2)
       Call REPLA(REDET$, MKI$(NPRO%), 7, 2)
       Call REPLA(REDET$, MKI$(CI0%), 9, 2)
       UNIMAS$ = Unimed$(CI0%)
       Call REPLA(REDET$, UNIMAS$, 11, 4)
       Call REPLA(REDET$, MKS$(CAOR), 55, 4)
       Call REPLA(REDET$, Unimed$(CI0%), 59, 4)
       Call REPLA(REDET$, MKS$(1), 63, 4)
       Call REPLA(REDET$, MKD$(PreUnid#), 67, 8)
       Call REPLA(REDET$, MKD$(PreUnid#), 79, 8)
       Call REPLA(REDET$, MKD$(PreUnid# * CAOR), 87, 8)
       Call REPLA(REDET$, MKI$(FENTRE%), 95, 2)
       If TISU% > 1 Then
         Call REPLA(REDET$, " *", 97, 2)
       End If
       Call REPLA(REDET$, MKS$(CARE), 99, 4)
       Call REPLA(REDET$, Chr$(STAT%), 124, 1)
       RENCA& = 1 + ULTREG&("OCOMENCA")
       Call REPLA(REDET$, MKS$(RENCA&), 125, 4)
       Call REGAPP("OCOMDETA", REDET$)
       '
       TOTANE# = TOTANE# + PreUnid# * CAOR
       REDETA& = ULTREG&("OCOMDETA")
       Call REPLA(ENCAO$, MKS$(REDETA&), 159, 4)
       '
       Call GRAREG("OCOMPRA", String$(128, 0), U&)
       '
     Next U&
     '
     If OCOA& <> 0 Then
       TOTIVA# = 0
       PVCL% = PIVACLI%((-1) * NPRO%)
       If PVCL% <> 2 Then
         If PVCL% <> 4 Then
           TOTIVA# = (Int(ALIVA * TOTANE# + 0.5)) / 100
         End If
       End If
       Call REPLA(ENCAO$, MKD$(TOTANE#), 89, 8)
       Call REPLA(ENCAO$, MKD$(TOTIVA#), 121, 8)
       Call REPLA(ENCAO$, MKD$(TOTANE# + TOTIVA#), 137, 8)
       Call REGAPP("OCOMENCA", ENCAO$)
     End If
     '
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     Screen.MousePointer = 1
     '
   End If
   '
End Sub

Function PRONUOCO&()
   '
   PRONU& = 1
   REGXX& = ULTREG&("OCOMENCA")
   If REGXX& > 0 Then
     PRONU& = 1 + CVS(Left$(REGLEIDO$("OCOMENCA", REGXX&), 4))
   End If
   If PRONU& <= REGXX& Then
     Screen.MousePointer = 11
     For U& = 1 To REGXX&
       X0$ = REGLEIDO$("OCOMENCA", U&)
       If CVS(Left$(X0$, 4)) >= PRONU& Then
         PRONU& = 1 + CVS(Left$(X0$, 4))
       End If
     Next U&
     Screen.MousePointer = 1
   End If
   '
   PRONUOCO& = PRONU&
   '
End Function

Sub SELCOMRAFE()
   '
   Has% = HOY(HOS$)
   FEX = Has%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   Des% = FECHANUM(DESS$)
   VDEF$ = MKI$(Des%) + MKI$(Has%)
   '
10 OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If OBX$ = "" Then
      Exit Sub
   End If
   Des% = CVI(Left$(OBX$, 2))
   Has% = CVI(Mid$(OBX$, 3, 2))
   If Has% < Des% Then GoTo 10
   FEX = Des%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
   FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("OCOMDETA")
   JJ& = 0
   REBLA$ = REGBLAN$("COPRECOM")
   Call HEADERS("COPRECOM", "")
   Call HEADERS("COPRECOM", "Periodo: " + PERIO$)
   '
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XXX$ = REGLEIDO$("OCOMDETA", U&)
     FECHEMI% = CVI(Mid$(XXX$, 5, 2))
     '
     STAT% = Asc(Mid$(XXX$, 124, 1))
     If STAT% <> 9 Then           ' no esta anulada
       If FECHEMI% >= Des% Then
         If FECHEMI% <= Has% Then
           '
           NPRO% = CVI(Mid$(XXX$, 7, 2))
           CIXI% = CVI(Mid$(XXX$, 9, 2))
           DOPE$ = ""
           NOPE% = CVI(Mid$(XXX$, 103, 2))
           If NOPE% > 0 Then
             DOPE$ = DESCRIOP$(CIXI%, NOPE%)
           End If
           NUOC& = CVS(Left$(XXX$, 4))
           UMEDI$ = Mid$(XXX$, 59, 4)
           If TRIM$(UMEDI$) = "" Then
             UMEDI$ = Unimed$(CIXI%)
           End If
           CACOM = CVS(Mid$(XXX$, 55, 4))
           '
           If STAT% = 8 Then                 ' SI SE DIO POR CUMPLIDA,
             CACOM = CVS(Mid$(XXX$, 99, 4))  ' LA CANT COMPRADA = A LA RECIBIDA
           End If
           '
           PREUNI# = CVD(Mid$(XXX$, 79, 8))
           '
           MONEMI% = Asc(Mid$(XXX$, 122, 1))
           If MONEMI% <= 1 Then MONEX$ = "    "
           If MONEMI% = 2 Then MONEX$ = "U$S "
           If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
           TICAMBIO = 1
           If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
           '
           YYY$ = REBLA$
           Call REPLA(YYY$, MKI$(CIXI%), 1, 2)
           Call REPLA(YYY$, MKI$(FECHEMI%), 3, 2)
           Call REPLA(YYY$, MKI$(NPRO%), 5, 2)
           Call REPLA(YYY$, MKI$(CIXI%), 7, 2)
           Call REPLA(YYY$, MKI$(FECHEMI%), 9, 2)
           Call REPLA(YYY$, DOPE$, 11, 22)
           Call REPLA(YYY$, MKS$(NUOC&), 33, 4)
           Call REPLA(YYY$, UMEDI$, 37, 4)
           Call REPLA(YYY$, MKS$(CACOM), 41, 4)
           Call REPLA(YYY$, MONEX$, 49, 4)
           Call REPLA(YYY$, MKD$(PREUNI#), 53, 8)
           ITOCOM# = PREUNI# * CACOM
           Call REPLA(YYY$, MKD$(ITOCOM#), 61, 8)
           Call REPLA(YYY$, MKD$(PREUNI# * TICAMBIO), 69, 8)
           Call REPLA(YYY$, MKD$(ITOCOM# * TICAMBIO), 77, 8)
           JJ& = JJ& + 1
           Call GRAREG("COPRECOM", YYY$, JJ&)
         End If
       End If
     End If
     '
   Next U&
   '
   Call SETULTREG("COPRECOM", JJ&)
   Call CIERRARCH("COPRECOM")
   Screen.MousePointer = 1
   '
End Sub
'
Sub DESANULOC()
    '
    'CARGO LA LISTA DE O/C QUE SE PUEDEN RESTAURAR
3   Call COPYSTRU("INDIOCOM", "IOCANUL")
    Screen.MousePointer = 11
    Call APERBASDAT("COMPRAS")
    SQLX$ = " SELECT Nume_Ocom, Refe_Ocom "
    SQLX$ = SQLX$ + " FROM OCOMENCA "
    SQLX$ = SQLX$ + " WHERE Stat_Ocom = 9 " ' Si esta anulada
    SQLX$ = SQLX$ + " AND Nume_Ocom > 0 "
    SQLX$ = SQLX$ + " ORDER BY Nume_Ocom "
    '
    Dim OCRESTEMP As ADODB.Recordset
    Set OCRESTEMP = New ADODB.Recordset
    OCRESTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With OCRESTEMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No  hay Ordenes de Compra\Anuladas que Puedan Restaurarse.")
        Screen.MousePointer = 1
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      J& = 0
      Do While (.EOF = False)
          XXY$ = REGBLAN$("IOCANUL")
          Call REPLA(XXY$, MKS(!nume_ocom), 1, 4) 'Nº O/C
          Call REPLA(XXY$, !Refe_OcoM, 5, 44) ' REFERENCIA
          J& = J& + 1
          Call GRAREG("IOCANUL", XXY$, J&)
      .MoveNext
      Loop
    End With
    Call SETULTREG("IOCANUL", J&)
    Call CIERRARCH("IOCANUL")
    Call BAJALDISCO
    Call FRESHECHO("IOCANUL")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
     Call MENSERR(24, "No hay Ordenes de Compra\Anuladas que Puedan Restaurarse.")
      Screen.MousePointer = 1
      GoTo 999
    End If
    '
    INDI% = 1
    Screen.MousePointer = 1
    '
10  YY$ = REGSEL$("IOCANUL/Indice de Ordenes de Compra Anuladas para Restauración")
    If Len(YY$) <= 4 Then GoTo 999
    '
    NUOCO& = CVS(Left$(YY$, 4))
    
    'VALIDAR QUE LA ORDEN DE COMPRA A REVERTIR NO TENGA ARTICULOS CON ESTRUCTURA
    SQLX$ = "SELECT * FROM OCOMDETA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE Nume_Ocom = " & NUOCO& & ""
    SQLX$ = SQLX$ + " AND Stat_Ocom = 9 "
    Dim VALIOCOMTMP As ADODB.Recordset
    Set VALIOCOMTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With VALIOCOMTMP
       Do While Not .EOF
          Ci% = XVALO(!cod_inte)
          Call AREXPLO(Ci%, 1, 1)
          If CAIT% > 0 Then
            TTXX$ = "Imposible Revertir Anulación de Esta Orden de Compra.\"
            TTXX$ = TTXX$ + "Posee Materiales con Estructura de Producto.\"
            TTXX$ = TTXX$ + "Debe Generar una Nueva Orden de Compra."
            Call COMUNI(TTXX$)
            GoTo 10
          End If
         .MoveNext
       Loop
    End With
    VALIOCOMTMP.Close
    Set VALIOCOMTMP = Nothing
    '
    'nuevo marcado de anulacion
    SQLX$ = "SELECT * FROM OCOMENCA"
    SQLX$ = SQLX$ + " WHERE Nume_Ocom = " & NUOCO& & ""
    SQLX$ = SQLX$ + " And Stat_Ocom = 9 "
    Dim OCOMMMENCA As ADODB.Recordset
    Set OCOMMMENCA = New ADODB.Recordset
    OCOMMMENCA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
    '
    With OCOMMMENCA
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
        !Stat_Ocom = 0                      'cambio status
        FECHEMII$ = Format(!Femi_Ocom, "dd/mm/yy")
        PRORAI$ = !Raso_Prov
        REFEI$ = FECHEMII$ & " - " & PRORAI$
        !Refe_OcoM = Left$(REFEI$, 64)   'cambio referencia
        .Update
       End If
       On Error GoTo 0
    End With
    Set OCOMMMENCA = Nothing
    '
    YAPRIREC% = 0
    SQLX$ = "SELECT * FROM OCOMDETA "
    SQLX$ = SQLX$ + " WHERE Nume_Ocom = " & NUOCO& & ""
    SQLX$ = SQLX$ + " AND Stat_Ocom = 9 "
    Dim OCOMDDDETA As ADODB.Recordset
    Set OCOMDDDETA = New ADODB.Recordset
    OCOMDDDETA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

    With OCOMDDDETA
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        Do While Not .EOF
          If !Cant_Rec > 0 Then STATUIX% = 2
          If !Cant_Rec >= !Cant_Ocom - 0.05 Then STATUIX% = 3
          !Stat_Ocom = STATUIX%
          .Update
        .MoveNext
        Loop
      End If
      On Error GoTo 0
    End With
    Set OCOMDDDETA = Nothing
    '
    Screen.MousePointer = 11
    '
    J& = 0
    For i& = 1 To ULTREG&("IOCANUL")
      XX$ = REGLEIDO$("IOCANUL", i&)
      If CVS(Left$(XX$, 4)) <> NUOCO& Then
        J& = J& + 1
        Call GRAREG("IOCANUL", XX$, J&)
      End If
    Next i&
    Call SETULTREG("IOCANUL", J&)
    '
    GoTo 3
    '
999 Call CIERRARCH("OCOMENCA")
    Call CIERRARCH("OCOMDETA")
    Call BAJALDISCO
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
End Sub

Sub ITEMCUMPLI()
    '
    TITU$ = "DAR POR CUMPLIDOS ITEMS DE COMPRAS"
    '
2   Screen.MousePointer = 11
    Call COPYSTRU("INDIOCOM", "INDIOCUM")
    Call APERBASDAT("COMPRAS")
    '
    '
10  SQLX$ = " SELECT  DISTINCT OCOMENCA.Nume_Ocom, OCOMENCA.Refe_Ocom "
    SQLX$ = SQLX$ + " FROM OCOMENCA "
    SQLX$ = SQLX$ + " INNER JOIN  OCOMDETA "
    SQLX$ = SQLX$ + " ON (OCOMENCA.Nume_Ocom = OCOMDETA.Nume_Ocom) "
    SQLX$ = SQLX$ + " WHERE OCOMDETA.Stat_Ocom < 8 " ' Excluyo las que estan anuladas O COMPLIDAS
    SQLX$ = SQLX$ + " AND OCOMDETA.Cant_Rec < ( OCOMDETA.Cant_Ocom - 0.05 ) "
    SQLX$ = SQLX$ + " AND OCOMDETA.Cod_Inte > 0 "
    SQLX$ = SQLX$ + " AND OCOMENCA.Stat_Ocom < 9 "
    SQLX$ = SQLX$ + " AND OCOMENCA.Nume_Ocom > 0 "
    SQLX$ = SQLX$ + " AND OCOMENCA.CODIEMPR = " & CodiEmp%
    SQLX$ = SQLX$ + " AND OCOMDETA.CODIEMPR = " & CodiEmp%
    SQLX$ = SQLX$ + " ORDER BY OCOMENCA.Nume_Ocom DESC "
    Dim OCCUMTEMP As ADODB.Recordset
    Set OCCUMTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With OCCUMTEMP
     If .EOF = True Then
        Call MENSERR(24, "No hay Ordenes de Compra\que Puedan Marcarse como cumplidas.")
        Screen.MousePointer = 1
        Exit Sub
     Else
      '
     J& = 0
      Do While (.EOF = False)
         XXY$ = REGBLAN$("INDIOCUM")
         Call REPLA(XXY$, MKS(!nume_ocom), 1, 4) 'Nº O/C
         Call REPLA(XXY$, !Refe_OcoM, 5, 44) ' REFERENCIA
         J& = J& + 1
         Call GRAREG("INDIOCUM", XXY$, J&)
      .MoveNext
      Loop
     End If
    End With
    OCCUMTEMP.Close
    Set OCCUMTEMP = Nothing
    '
    Call SETULTREG("INDIOCUM", J&)
    Call CIERRARCH("INDIOCUM")
    Call FRESHECHO("INDIOCUM")
    Call BAJALDISCO
    '
    If J& < 1 Then
      Call MENSERR(24, "No hay Ordenes de Compra\Pendientes que Puedan Anularse.")
      Screen.MousePointer = 1
      GoTo 999
    End If
    '
    INDI% = 1
    Screen.MousePointer = 1
    '
''''10  YY$ = REGSEL$("INDIOCUM/Indice de Ordenes de Compra Pendientes para Anulación")
''''    If Len(YY$) <= 4 Then GoTo 999
    
    Call SELECHO("INDIOCUM")
    NUOCO& = XVALO(VALACT1$("INDIOCUM"))
    If NUOCO& < 1 Then GoTo 999
    '
    'GABAL
    CONDI$ = "NUME_OCOM = " & NUOCO&
    ORDSER& = XVALO(VALOBADA("OCOMDETA", "N_ORSERV", CONDI$))
    OF$ = SAFETEXT$(VALOBADA("OCOMDETA", "DOC_ORIG", CONDI$))
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    MODOGABAL% = 0: If InStr(EPAC$, "GABAL") > 0 And ORDSER& > 0 Then MODOGABAL% = 1
    If MODOGABAL% > 0 Then
        Call COMUNI("Imposible Marcar Items Cumplidos Para esta O.C. \ Generada a Partir de O.Fab.: " & OF$)
        GoTo 999
    End If
    '
    DAPORCU07.Label14 = NUOCO&
    '
    DAPORCU07.Show 1
    '
    ' A PARTIR DE AQUI MARCAR REGISTROS QUE SE DAN POR CUMPLIDOS
    Call APERBASDAT("COMPRAS")
    If XVALO(DAPORCU07.Label14) = NUOCO& Then
       Screen.MousePointer = 11
       'URE& = ULTREG&("OCOMDETA")
       For U& = 1 To DAPORCU07.LICOPEN.ListCount
         If DAPORCU07.LICOPEN.Selected(U& - 1) = True Then
            TTXX$ = DAPORCU07.LICOPEN.List(U& - 1)
            NUORIT% = XVALO(TRIM$(Mid$(TTXX$, 121, 8)))
            SQLX$ = "SELECT * FROM OCOMDETA "
            SQLX$ = SQLX$ + "WHERE Nume_Ocom = " & NUOCO& & " "
            SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
            SQLX$ = SQLX$ + "AND NuOrd_Item = " & NUORIT% & " "
                         
            Dim MARCTEMP As ADODB.Recordset
            Set MARCTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
            With MARCTEMP
               If Not (.BOF And .EOF) Then
                  Do While Not .EOF
'                     COEFCAN = !Coef_Unids
'                     If COEFCAN < 0.005 Then COEFCAN = 1
                     CAORI = XVALO(!Cant_Ocom) '* COEFCAN se comenta para que tome el valor directo del campo
                     CAREC = XVALO(!Cant_Rec)
                     CAPEN = CAORI - CAREC
                     If CAPEN >= 0.005 Then
                        STRSQL = "Update OCOMDETA set Stat_Ocom = 8"
                        STRSQL = STRSQL + " WHERE Nume_Ocom = " & NUOCO& & " "
                        STRSQL = STRSQL + " AND Stat_Ocom < 9 "
                        STRSQL = STRSQL + " AND NuOrd_Item = " & NUORIT% & " "
                         FLEXCONN.Execute (FILTSQL$(STRSQL))
                     End If
                  .MoveNext
                  Loop
               End If
            End With
            '
            NOCOM1$ = TRIM$(Str$(NUOCO&))
              While Len(NOCOM1$) < 6: NOCOM1$ = "0" + NOCOM1$: Wend
            ITNU1$ = TRIM$(Str$(NUORIT%))
              While Len(ITNU1$) < 2: ITNU1$ = "0" + ITNU1$: Wend
            NOCOBUS$ = "OC-" + Left$(NOCOM1$, 6) + "-" + ITNU1$
            ABUS$ = "Nro_Ocom = '" & NOCOBUS$ & "'"
            '
            ' CANCELAR RESERVAS
            ' BORRA LOS REGISTROS EN LOS QUE NO SE ENTREGO NADA
'            STRSQL = "DELETE MACONSIG where Nro_Ocom = '" & NOCOBUS$ & "' AND Cant_Entreg < 0.05"
'            FLEXCONN.Execute (FILTSQL$(STRSQL))
            CONDI$ = "Nro_Ocom = '" & NOCOBUS$ & "' AND Cant_Entreg < 0.05"
            Call BORRAREGISTROS("MACONSIG", CONDI$, REGAF&, "NOMESS")
            '
            ' PONE LA CANTIDAD ASIGNADA IGUAL A LA ENTREGADA PARA NO DESPACHAR MÁS
            STRSQL = "update MACONSIG set Cant_Asig = Cant_Entreg where Nro_Ocom='" & NOCOBUS$ & "'"
            FLEXCONN.Execute (FILTSQL$(STRSQL))
            '
            'SI QUEDAN REGISTROS CON CANTIDAD ENTREGADA PRESENTAR MENSAJE.
            CONDI$ = "Nro_Ocom = '" & NOCOBUS$ & "' AND (Cant_Entreg - Cant_Rendid)>0.05 "
            CANTENTR& = CantRegistros("MACONSIG", CONDI$)
            If CANTENTR& > 0 Then
               Call COMUNI("Hay Materiales Yá Entregados en Consignación\ Al Proveedor para el Item Seleccionados.\Recuerde Producir el RE-Ingreso")
            End If
            '
         End If
       Next U&
    End If
    '
    Unload DAPORCU07
    Screen.MousePointer = 1
    '
    GoTo 10
    '
999 Call CIERRARCH("*.*")
    '
End Sub

Sub DEVOLCONSI()

        If DERACCE%("DEMATCON", "Devolución de Materiales Consignados") < 2 Then Exit Sub
        
1105    Call CIERRARCH("*.*")
        '
        Screen.MousePointer = 1
        If NPROVE% > 0 Then
             EBOL$ = MKI$(NPROVE%)
          Else
             EBOL$ = OBJPLA$("SELPROVE", "")
             If EBOL$ = "" Then
                GoTo 9999
             End If
        End If
        '
        NC% = CVI(EBOL$)
        If NC% < 1 Or ECOARCH$("PROVED1", MKI$(NC%)) = "" Then
            Call MENSERR(24, "Proveedor Inexistente")
            NPROVE% = 0
            GoTo 1105
        End If
        '
        NC1% = (-1) * NC%
        '
1112    Call CIERRARCH("*.*")
        Call BLANARCH("!BOLREDET")
        '
        Screen.MousePointer = 11
        HAYOC% = 0
        Call APERBASDAT("STOCKS")
        Call PoneEnCeroNull("MACONSIG", "CANT_DEVU", REGAF&, "NOMESS")
        SQLX$ = "SELECT * FROM MACONSIG "
        SQLX$ = SQLX$ + " WHERE ((Cant_Entreg - Cant_Rendid - Cant_Devu)>0 "
        SQLX$ = SQLX$ + " OR CANT_RENDID IS NULL) "
        SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
        'Set MACONSITMP = Stocks.OpenRecordset(SQLX$, 4)
        Dim MACONSITMP As ADODB.Recordset
        Set MACONSITMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With MACONSITMP
         On Error Resume Next
         .MoveFirst
         If Err Then
            On Error GoTo 0
            GoTo 555
         Else
            On Error GoTo 0
            Do While .EOF = False
              CAPENE# = XVALO(!cant_entreg) - XVALO(!CANT_RENDID) - XVALO(!cant_devu)
              NUOC& = XVALO(Mid$(!nro_ocom, 4, 6))
              Ci% = XVALO(!cod_inte)
              '
              Y$ = REGBLAN$("!BOLREDET")
              Call REPLA(Y$, MKI$(Ci%), 83, 2)
              Call REPLA(Y$, Unimed$(Ci%), 85, 4)
              Call REPLA(Y$, MKD$(CAPENE#), 89, 8)
              Call REPLA(Y$, MKS$(NUOC&), 151, 4)
              Call REPLA(Y$, MKS$(i&), 189, 4)
              HAYOC% = 1
              Call REGAPP("!BOLREDET", Y$)
            .MoveNext
            Loop
          End If
        End With
        MACONSITMP.Close
        Set MACONSITMP = Nothing
        
        '
555     If HAYOC% = 0 Then
            Call COMUNI("No hay Consignaciones Pendientes\para este Proveedor.")
        End If
'Jandy si no hay consignaciones no hace el insert
        
        '
1113    VTB% = VENTABU%("ECONSIG/NC/TITUPAN=Rendición Materiales Consignados - " + TRIM$(rasocli$(NC1%)))
        If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
            X$ = "": Y$ = "": YY$ = ""
            EBOL$ = "": PEDY$ = "": DESCRI$ = ""
            GoTo 1105
        End If
        '
11141   CONTA% = ULTREG&("!BOLREDET")
        '
1115    If CONTA% > 0 Then
        '
          Call BLOQARCH("MACONSIG")
          Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
          '
          AINDI% = 0
          NUITEM = 0
          For YX& = 1 To ULTREG&("!BOLREDET")
            YY$ = REGLEIDO$("!BOLREDET", YX&)
            Ci% = CVI(Mid$(YY$, 83, 2))
            CAN# = CVD(Mid$(YY$, 89, 8))
            LOTE$ = Mid$(YY$, 69, 12)
            If Ci% > 0 Then
              If Ci% <= NUMAS% Then
                If CAN# > 0.005 Then
                    '
                    ' NUEVO CODIGO DE RENDICION DE CONSIGNACIONES
                    NUOCO& = CVS(Mid$(YY$, 151, 4))
                    CAREN# = 0
                    CAPEN# = CAN#
                    '
                    ' ABRE LA TABLA DE RENDICION DE CONSIGNACIONES
                    SQLX$ = "SELECT * From RENCONSIG where Nu_Prove = " & NC% & " "
                    Dim RenConsig As ADODB.Recordset
                    Set RenConsig = New ADODB.Recordset
25                  On Error Resume Next
                    RenConsig.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                       On Error GoTo 0
                       Call CapturaErrorOpen
                       GoTo 25
                    End If
                    On Error GoTo 0
                    '
                    CODDX$ = CODEXT$(Ci%)
                    CAPENDE# = CAN#
                    FE% = FERECEP%: If FE% < 1 Then FE% = HOY(HO$)
                    'FECH = FETEXCOR1$(FE%) 'COMENTADO POR Q' SI NO GENERA ERROR EN GRABACION
                    NOCOM$ = TRIM$(Str$(NUOCO&))
                    While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
                    NORCOM$ = "OC-" + Left$(NOCOM$, 6) + "%"
                    '
                    With RenConsig
                       ' AQUÍ DEVOLUCION DIRECTA DEL MATERIAL RECIBIDO
                       .AddNew
                       !Fe_Recep = CVDAT(FE%)
                       !NUMEBORE = NUBOL&
                       !nro_ocom = NORCOM$
                       !Nu_Prove = NC%   ' NCX% ERROR EN NOMBRE DE VARIABLE
                       !Codi_Conj = 0
                       !Cant_Conj = 0
                       !Codi_Elem = Ci%  ' CIX% ERROR EN NOMBRE DE VARIABLE
                       !CANT_RENDID = CAPENDE#
                       !TipoMovim = "DV"
                       !MarProce = 0
                       .Update
                    End With
                    RenConsig.Close: Set RenConsig = Nothing
                    '
                End If
              End If
            End If
            '
          Next YX&
          Call AJUMACON
          '
          Call CIERRARCH("MACONSIG")
          '
          Screen.MousePointer = 1
          '
1199    End If
        '
9999    Call CIERRARCH("*.*")
        '
End Sub

Sub AjusteAFechaConsignacionPorProveedorOrdenCompra()

        'If DERACCE%("DEMATCON", "Devolución de Materiales Consignados") < 2 Then Exit Sub
        
1105    Call CIERRARCH("*.*")
        HAYOC% = 0
        '
        opciones% = COMALTER%("Opciones:\\Consulta Saldos a Fecha (Todos los Proveedores)\Ajuste a Fecha Por Proveedor y O-Compra\Ajuste a Cero a Fecha (Todos los Proveedores)\Correccion de Saldos a Fecha (MACONSIG)\Depurar Consignaciones")
        If opciones% < 1 Or opciones% > 5 Then GoTo 9999
            
        NC% = 0
        If opciones% = 2 Or opciones% = 4 Or opciones% = 5 Then
            EBOL$ = OBJPLA$("SELPROVE", "")
            If EBOL$ = "" Then
                GoTo 9999
            End If
            '
            NC% = CVI(EBOL$)
            NC1% = (-1) * NC%
            If NC% < 1 Or TRIM$(rasocli(NC1%)) = "" Then
                Call MENSERR(24, "Proveedor Inexistente")
                NPROVE% = 0
                GoTo 1105
            End If
            
            If opciones% = 5 Then
                GoTo 10000
            End If
        End If
        '
        Call SELECHO("SELFECHA")
        VDEV$ = TRIM$(VALACT1$("SELFECHA"))
        FechaCorte$ = FETEXCOR1$(HOY(HO$))
        If VDEV$ <> "" Then FechaCorte$ = FETEXCOR1$(FECHANUM(VDEV$))
        '
1112    Call CIERRARCH("*.*")
        Call BLANARCH("!BOLREDET")
        '
        ' *** OBTIENE LAS CONSIGNACIONES REALIZADAS AL PROVEEDOR ELEGIDO Y HASTA LA FECHA DE CORTE
        SQLX$ = "SELECT NPRO_OCOM, COD_INTE, substring(NRO_OCOM,1,9) AS OCOM, SUM(CANT_ENTREG) AS ENTREGADO, SUM(CANT_RENDID) AS RENDIDO FROM MACONSIG "
        CONDI$ = " Npro_Ocom > 0 "
        If NC% > 0 Then CONDI$ = CONDI$ & " AND Npro_Ocom = " & CStr(NC%)
        If FechaCorte$ <> "" Then CONDI$ = CONDI$ & " AND FEMI_OCOM<='" & FechaCorte$ & "'"
        SQLX$ = SQLX$ & " WHERE " & CONDI$
        SQLX$ = SQLX$ & " GROUP BY COD_INTE, substring(NRO_OCOM,1,9), NPRO_OCOM ORDER BY NPRO_OCOM "
        '
        FIN& = CantRegistros("MACONSIG", CONDI$)
        II& = 0
        '
        Dim MACONSITMP As ADODB.Recordset
        Set MACONSITMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With MACONSITMP
         On Error Resume Next
         .MoveFirst
         If Err Then
            On Error GoTo 0
            GoTo 555
         Else
            On Error GoTo 0
            Screen.MousePointer = 11
            Do While .EOF = False
              II& = II& + 1
              Call TRACE(24, II&, FIN&)
                
              Ci% = XVALO(!cod_inte)
              If Ci% = 1608 Then
              A = B
              End If
              NumeroProveedor% = XVALO(!Npro_Ocom)
              ORDENCOMPRA$ = Mid$(!ocom, 1, 9)
              CondicionSql$ = "CODI_ELEM=" & CStr(Ci%) & " AND NRO_OCOM LIKE '" & ORDENCOMPRA$ & "%'" & " And NU_PROVE = " & CStr(NumeroProveedor%)
              CondicionSql$ = CondicionSql$ & " AND (TIPOMOVIM = 'DV' OR TIPOMOVIM = 'CC' OR TIPOMOVIM = 'QR')"
              If FechaCorte$ <> "" Then CondicionSql$ = CondicionSql$ & " AND FE_RECEP<='" & FechaCorte$ & "'"
              
              cantidadRendida# = XVALO(VALOBADA("RENCONSIG", "SUM(CANT_RENDID)", CondicionSql$))
              ' *** SI NO HAY MOVIMIENTOS EN RENCONSIG ASUME QUE NO FUERON MIGRADOS DESDE MACONSIG
              'If cantidadRendida# < 0.05 Then cantidadRendida# = XVALO(!RENDIDO)
              
              If opciones% = 4 Then cantidadRendida# = XVALO(!RENDIDO)
              CAPENE# = XVALO(!ENTREGADO) - cantidadRendida#
              NUOC& = XVALO(Mid$(!ocom, 4, 6))
              '
              If CAPENE# > 0.05 Or (Abs(CAPENE#) > 0 And opciones% = 4) Then
                Y$ = REGBLAN$("!BOLREDET")
                Call REPLA(Y$, MKI$(Ci%), 83, 2)
                Call REPLA(Y$, Unimed$(Ci%), 85, 4)
                Call REPLA(Y$, MKD$(CAPENE#), 89, 8)
                Call REPLA(Y$, MKI$(NumeroProveedor%), 105, 2)
                Call REPLA(Y$, MKS$(NUOC&), 151, 4)
                Call REPLA(Y$, MKD$(CAPENE#), 163, 8)   ' GUARDA SALDO A LA FECHA DE CORTE
                Call REPLA(Y$, MKI$(Ci%), 187, 2)       ' GUARDA EL CODIGO ORIGINAL. SE UTILIZA PARA EVITAR QUE LO CAMBIE
                Call REPLA(Y$, MKS$(i&), 189, 4)
                HAYOC% = 1
                Call REGAPP("!BOLREDET", Y$)
              End If
            .MoveNext
            Loop
          End If
        End With
        MACONSITMP.Close
        Set MACONSITMP = Nothing
        '
555     If HAYOC% = 0 Then
            If opciones% = 1 Then
                Call COMUNI("No hay Consignaciones Pendientes.") ' CONSULTA DE SALDOS
            ElseIf opciones% = 2 Then
                Call COMUNI("No hay Consignaciones Pendientes\para este Proveedor.") ' AJUSTE POR PROVEEDOR
            Else
                Call COMUNI("No hay Consignaciones Pendientes.") ' AJUSTE A CERO
            End If
            Screen.MousePointer = 1
            GoTo 1105
        End If
        '
        Screen.MousePointer = 1
        '
        If opciones% = 3 Then GoTo 11141
        '
        tituloVentana$ = "Saldos a Fecha (Todos los Proveedores)"
        If opciones% = 2 Or opciones% = 4 Then tituloVentana$ = "Ajuste de Materiales Consignados - " & TRIM$(rasocli$(NC1%))
1113    VTB% = VENTABU%("AJUCONSI/NC/TITUPAN=" & tituloVentana$)
        If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
            X$ = "": Y$ = "": YY$ = ""
            EBOL$ = "": PEDY$ = "": DESCRI$ = ""
            GoTo 1105
        End If
        '
11141   CONTA% = ULTREG&("!BOLREDET")
        '
1115    If CONTA% > 0 And (opciones% = 2 Or opciones% = 3 Or opciones% = 4) Then
          '
          Call BLOQARCH("MACONSIG")
          Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
          '
          AINDI% = 0
          NUITEM = 0
          For YX& = 1 To ULTREG&("!BOLREDET")
            Call TRACE(24, YX&, CLng(CONTA%))
            YY$ = REGLEIDO$("!BOLREDET", YX&)
            Ci% = CVI(Mid$(YY$, 83, 2))
            codigoInternoOriginal% = CVI(Mid$(YY$, 187, 2))
            cantidadAjustada# = CVD(Mid$(YY$, 89, 8))
            cantidadConsignada# = CVD(Mid$(YY$, 163, 8))
            NumeroProveedor% = CVI(Mid$(YY$, 105, 2))
            LOTE$ = Mid$(YY$, 69, 12)
            If Ci% > 0 And Ci% <= NUMAS% Then
              If Ci% = codigoInternoOriginal% Then
                If opciones% = 4 Then
                    NUOCO& = CVS(Mid$(YY$, 151, 4))
                    
                    Call AJUMACON(NUOCO&, Ci%, NumeroProveedor%)
                    GoTo 35
                End If
                
                If opciones% = 3 Then cantidadAjustada# = 0
                diferenciaAjustar# = cantidadConsignada# - cantidadAjustada#
                If diferenciaAjustar# > 0.005 Then
                    '
                    ' NUEVO CODIGO DE RENDICION DE CONSIGNACIONES
                    NUOCO& = CVS(Mid$(YY$, 151, 4))
                    CAPENDE# = diferenciaAjustar#
                    '
                    ' ABRE LA TABLA DE RENDICION DE CONSIGNACIONES
                    SQLX$ = "SELECT * From RENCONSIG where Nu_Prove = " & NumeroProveedor% & " "
                    Dim RenConsig As ADODB.Recordset
                    Set RenConsig = New ADODB.Recordset
25                  On Error Resume Next
                    RenConsig.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                       On Error GoTo 0
                       Call CapturaErrorOpen
                       GoTo 25
                    End If
                    On Error GoTo 0
                    '
                    CODDX$ = CODEXT$(Ci%)
                    FE% = FECHANUM(VDEV$): If FE% < 1 Then FE% = HOY(HO$)
                    NOCOM$ = TRIM$(Str$(NUOCO&))
                    While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
                    NORCOM$ = "OC-" + Left$(NOCOM$, 6) + "%"
                    '
                    With RenConsig
                       ' AJUSTA EL SALDO DEL MATERIAL CONSIGNADO. ESTO SIGNIFICA QUE NO LO VUELVE A PEDIR
                       ' FUNCIONA COMO SI ESTUVIERA RENDIDO O SI SE HUBIESE DEVUELTO DESDE EL MODULO DE RECEPCION.
                       .AddNew
                       !Fe_Recep = CVDAT(FE%)
                       !NUMEBORE = NUBOL&
                       !nro_ocom = NORCOM$
                       !Nu_Prove = NumeroProveedor%
                       !Codi_Conj = 0
                       !Cant_Conj = 0
                       !Codi_Elem = Ci%
                       !CANT_RENDID = CAPENDE#
                       !TipoMovim = "DV"
                       !MarProce = 0
                       .Update
                    End With
                    RenConsig.Close: Set RenConsig = Nothing
                    '
                ElseIf diferenciaAjustar# < 0 Then
                    codigoOriginal$ = CODEXT$(Ci%)
                    Call COMUNI("No es Posible Ajustar el Material Consignado en Cantidades Mayores\a la Presentada Para el Artículo '" & codigoOriginal$ & "'." _
                                & "\(Fila número " & CStr(YX&) & ")\ \Debe Generar una Nueva Consignación\Para Reflejar el Aumento del Material en Terceros.")
                    '
                End If
              Else
                    codigoOriginal$ = CODEXT$(codigoInternoOriginal%)
                    codigoCambiado$ = CODEXT$(Ci%)
                    Call COMUNI("No es Posible Cambiar los Materiales Consignados o Agregar Nuevos.\" _
                            & "Artículo '" & codigoOriginal$ & "'" & " Cambiado por '" & codigoCambiado$ & "'." _
                            & "\(Fila número " & CStr(YX&) & ")\ \Debe Generar un Nuevo Remito para Consignar Otros Materiales")
              End If
            End If
            '
35        Next YX&
          Call AJUMACON
          '
          Call CIERRARCH("MACONSIG")
          Screen.MousePointer = 1
          '
          Call COMUNI("Proceso de Ajuste Finalizado")
          GoTo 1105
          '
1199    End If
        '
9999    Call CIERRARCH("*.*")
        Exit Sub
        '
10000
    
    CONDI$ = "COD_INTE>0 AND NPRO_OCOM=" & CStr(NC%)
    FIN& = CantRegistros("MACONSIG", CONDI$)
    SQLX$ = "SELECT * FROM MACONSIG WHERE " & CONDI$
    Dim RST_CONSIGNA As ADODB.Recordset
    Set RST_CONSIGNA = New ADODB.Recordset
    RST_CONSIGNA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    JJ& = 1
    With RST_CONSIGNA
        Do While Not .EOF
            Call TRACE(24, JJ&, FIN&)
            NROOCOM& = XVALO(Mid$(SAFETEXT$(!nro_ocom), 4, 6))
            CONDI$ = "NUME_OCOM = " & NROOCOM&
            Stat_Ocom = XVALO(VALOBADA("OCOMENCA", "STAT_OCOM", CONDI$))
            If Stat_Ocom > 0 Then
                NORCOM$ = Left$(SAFETEXT$(!nro_ocom), 9) + "%"
                CIX% = XVALO(!cod_inte)
                NCX% = XVALO(!Npro_Ocom)
                '
                !cant_entreg = !cant_asig
                !CANT_RENDID = !cant_entreg
                !cant_devu = 0
                .Update
                '
                Call BORRAREGISTROS("RENCONSIG", "nro_ocom like'" & NORCOM$ & "'", REGAF&)
            Else
                ORD% = XVALO(Mid$(SAFETEXT$(!nro_ocom), 11, 2))
                CONDI$ = "NUME_OCOM= " & NROOCOM& & " AND NUORD_ITEM=" & ORD%
                Stat_Ocom = XVALO(VALOBADA("OCOMDETA", "STAT_OCOM", CONDI$))
                CCOMP = XVALO(VALOBADA("OCOMDETA", "cant_ocom", CONDI$))
                CRECI = XVALO(VALOBADA("OCOMDETA", "cant_rec", CONDI$))
                SALDO = CCOMP - CRECI
                If Stat_Ocom > 0 Or SALDO <= 0.005 Then
                    NORCOM$ = Left$(SAFETEXT$(!nro_ocom), 9) + "%"
                    CIX% = XVALO(!cod_inte)
                    NCX% = XVALO(!Npro_Ocom)
                    '
                    !cant_entreg = !cant_asig
                    !CANT_RENDID = !cant_entreg
                    !cant_devu = 0
                    .Update
                    '
                    Call BORRAREGISTROS("RENCONSIG", "nro_ocom like'" & NORCOM$ & "' AND codi_elem=" & CIX%, REGAF&)
                End If
            End If
            .MoveNext
            JJ& = JJ& + 1
        Loop
    End With
    '
    Call COMUNI("Proceso Finalizado")
    '
    GoTo 9999
End Sub

Sub RANKPRO()
'\\\OT-05-0381-WAR-07/06/05///
    Dim TOTCOMPRA#(32767)
    '
    Call DESHASFECHA(Des%, Has%, PERIO$)
    '
    Screen.MousePointer = 11
    J& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTCOMPRA#(U&) = 0
    Next U&
    '
    DESDAT = Int(CDbl(CVDAT(Des%)))
    HASDAT = Int(CDbl(CVDAT(Has%)))
    Desde = FECHATEX(Des%)
    Hasta = FECHATEX(Has%)
    '
    Call APERBASDAT("COMPRAS")
    SQLX$ = "SELECT * FROM OCOMDETA "
    'SQLX$ = SQLX$ + "WHERE INT(CDBL(Femi_Ocom)) >= " & DESDAT & " "
    'SQLX$ = SQLX$ + "AND INT(CDBL(Femi_Ocom)) <= " & HASDAT & " "
    SQLX$ = SQLX$ + "WHERE Femi_Ocom >= '" & Desde & "'"
    SQLX$ = SQLX$ + " AND Femi_Ocom  <= '" & Hasta & "'"
    '
    SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "    ' NO ANULADAS
    '
    SQLX$ = SQLX$ + " AND Nume_Ocom > 0 "
    'Set CABATEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = New ADODB.Recordset
    CABATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    
    PROMAX% = 0
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
      FIN& = .RecordCount
      .MoveFirst
      UK& = 0
      '
      Do While Not .EOF
            UK& = UK& + 1
            Call TRACE(20, UK&, FIN&)
            NPRO% = XVALO(!Npro_Ocom)
            If NPRO% > 0 Then
              If NPRO% > PROMAX% Then PROMAX% = NPRO%
              IMPNETO# = (!prun_neto * !Cant_Ocom)
              TOTCOMPRA#(NPRO%) = TOTCOMPRA#(NPRO%) + (IMPNETO# * TICAMONE(!Mone_Emis))
            End If
        .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    '
    For i% = 1 To PROMAX
      TOTAGEN# = TOTAGEN# + TOTCOMPRA#(i%)
    Next i%
    '
    List1.Clear
    For NCI% = 1 To 32766
       If Abs(TOTCOMPRA#(NCI%)) > 0.005 Then
          RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTCOMPRA#(NCI%)), 4, 12, 2, 2)
          PORPAR = 100 * TOTCOMPRA#(NCI%) / TOTAGEN#
          RELIB$ = RELIB$ + CSTRING$(MKI$(NCI%), 1, 6, 0, 2)
          RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
          RELIB$ = RELIB$ + CSTRING$(MKD$(TOTCOMPRA#(NCI%)), 4, 12, 2, 2)
          List1.AddItem RELIB$
       End If
    Next NCI%
    '
    J& = 0: PORACU = 0
    FIN& = List1.ListCount
    '
    For U& = 1 To FIN&
        List1.ListIndex = U& - 1
        NC% = XVALO(Mid$(List1.TEXT, 14, 6))
        TONEVEN# = XVALO(Mid$(List1.TEXT, 28, 12))
        PORPAR = XVALO(Mid$(List1.TEXT, 20, 8))
        PORACU = PORACU + PORPAR
        '
        Z$ = REGBLAN$("RANKCLI")
        J& = J& + 1: JI% = J&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(NC%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call GRAREG("RANKCLI", Z$, J&)
        '
    Next U&
    '
    Call SETULTREG("RANKCLI", J&)
    Call CIERRARCH("RANKCLI")
    Call HEADERS("RANKCLI", "")
    Call HEADERS("RANKCLI", "Periodo: " + PERIO$)
    Screen.MousePointer = 1
    List1.Clear
    '
    Call FINAL("*LRANPRO")
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
'\\\OT-05-0381-WAR-FIN///
End Sub

Sub RANKART()
'\\\OT-05-0382-WAR-07/06/05
    '
    Dim TOTART#(32767)
    '
    Call DESHASFECHA(Des%, Has%, PERIO$)
    '
    Screen.MousePointer = 11
    J& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTART#(U&) = 0
    Next U&
    '
    DESDAT = Int(CDbl(CVDAT(Des%)))
    HASDAT = Int(CDbl(CVDAT(Has%)))
    Desde = FECHATEX(Des%)
    Hasta = FECHATEX(Has%)
    '
    Call APERBASDAT("COMPRAS")
    SQLX$ = "SELECT * FROM OCOMDETA "
    SQLX$ = SQLX$ + "WHERE Femi_Ocom >= '" & Desde & "'"
    SQLX$ = SQLX$ + " AND Femi_Ocom  <= '" & Hasta & "'"
    SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "  ' NO ANULADAS
    '
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$), RAFE)
    '
    PROMAX% = 0
    '
    With CABATEMP
      On Error Resume Next
      .MoveFirst
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos en el Período Elegido")
        GoTo 99
      End If
      On Error GoTo 0
'      FIN& = .RecordCount
'      .MoveFirst
      UK& = 0
      '
      Do While Not .EOF
         UK& = UK& + 1
         Call TRACE(20, UK&, FIN&)
         NART% = !cod_inte
         IMPNETO# = (!prun_neto * !Cant_Ocom)
         If NART% > 0 Then
           If NART% > PROMAX% Then PROMAX% = NART%
           TOTART#(NART%) = TOTART#(NART%) + (IMPNETO# * TICAMONE(!Mone_Emis))
         Else
           TOTART#(0) = TOTART#(0) + (IMPNETO# * TICAMONE(!Mone_Emis))
         End If
      .MoveNext
      Loop
    End With
    CABATEMP.Close
    Set CABATEMP = Nothing
    '
    For i% = 0 To PROMAX
      TOTAGEN# = TOTAGEN# + TOTART#(i%)
    Next i%
    '
    List1.Clear
    For NCI% = 0 To 32766
       If Abs(TOTART#(NCI%)) > 0.005 Then
          RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTART#(NCI%)), 4, 12, 2, 2)
          PORPAR = 100 * TOTART#(NCI%) / TOTAGEN#
          RELIB$ = RELIB$ + CSTRING$(MKI$(NCI%), 1, 6, 0, 2)
          RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
          RELIB$ = RELIB$ + CSTRING$(MKD$(TOTART#(NCI%)), 4, 12, 2, 2)
          List1.AddItem RELIB$
       End If
    Next NCI%
    '
    J& = 0: PORACU = 0
    FIN& = List1.ListCount
    '
    For U& = 1 To FIN&
        List1.ListIndex = U& - 1
        NC% = XVALO(Mid$(List1.TEXT, 14, 6))
        TONEVEN# = XVALO(Mid$(List1.TEXT, 28, 12))
        PORPAR = XVALO(Mid$(List1.TEXT, 20, 8))
        PORACU = PORACU + PORPAR
        '
        Z$ = REGBLAN$("RANKPIE")
        J& = J& + 1: JI% = J&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(NC%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call GRAREG("RANKPIE", Z$, J&)
        '
    Next U&
    '
    Call SETULTREG("RANKPIE", J&)
    Call CIERRARCH("RANKPIE")
    Call HEADERS("RANKPIE", "")
    Call HEADERS("RANKPIE", "Periodo: " + PERIO$)
    Screen.MousePointer = 1
    List1.Clear
    '
    Call FINAL("*LRANART")
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
'\\\OT-05-0382-WAR-FIN
End Sub

Sub IMPPROVCOD()
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
    Call ABRECOMPRAS
    SQLX$ = "SELECT * FROM PROVECOD "
    Set PROVCODTMP = New ADODB.Recordset
    PROVCODTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   
    With Data1
     .Recordset.MoveFirst
     Do While Not .Recordset.EOF
        Codigo$ = ""
        On Error Resume Next
        NPRO% = .Recordset.Fields(0)
        Codigo$ = .Recordset.Fields(1)
        CIIXX% = CODINT(Codigo$)
        If CIIXX% < 1 Then
          Screen.MousePointer = 1
          Call MENSERR(24, "Código '" & Codigo$ & "' Inexistente o con Baja Logica.\Dar de Alta y Volver a Correr el Proceso de Importación")
          GoTo 20
        End If
        On Error GoTo 0
        '
        CONDI$ = "NUM_PROV = " & NPRO% & " AND Cod_Inte = " & CIIXX%
        If CantRegistros&("PROVECOD", CONDI$) > 0 Then GoTo 20
        With PROVCODTMP
              .AddNew   ' YA CONVERTIDO
              On Error Resume Next
              !CODIEMPR = CodiEmp%
              On Error GoTo 0
              !Num_Prov = NPRO%
              !cod_inte = CIIXX%
              !PRECIO = 0
              !Moneda = 0
              !FECHA = CVDAT(0)
              !MARBORRA = 0
              .Update
        End With
20      Screen.MousePointer = 11
       .Recordset.MoveNext
     Loop
     PROVCODTMP.Close
     Set PROVCODTMP = Nothing
   End With
   Screen.MousePointer = 1
   Call MENSERR(24, "Proceso de Importación Finalizada")
99 '
End Sub
'
Sub ANULAORSER(NUOCO&)
    'Pregunta si la O/Compra tiene una O/Servicio Asociada
    '
    CONDI$ = "NUME_OCOM = " & NUOCO&
    ORDSER& = XVALO(VALOBADA("OCOMDETA", "N_ORSERV", CONDI$))
    
    If ORDSER& <> 0 Then
       '
       Operacion$ = SAFETEXT$(VALOBADA("OCOMDETA", "COD_OPER", CONDI$))
       OF$ = SAFETEXT$(VALOBADA("OCOMDETA", "DOC_ORIG", CONDI$))
       CONDI$ = "IDSUBOR = '" & TRIM$(OF$) & "' AND NUMEOPER = " & Operacion$
       Cantidad# = XVALO(VALOBADA("OPERFAB", "CANTPROG", CONDI$))
       '
       Call ACTUREGISTRO("OPERFAB", "CANTPEND", CONDI$, Cantidad#, REG&)
       Call ACTUREGISTRO("OPERFAB", "CANTREAL", CONDI$, 0, REG&)
       Call FILTERDELETE("OPERTER", 9, MKS(ORDSER&))
       '
    End If
    '
End Sub

Function TIENECOTIZARECIB%(NUOCO&)
    '
    RT% = 0
    SQLX$ = "SELECT * FROM OCOMDETA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE STAT_OCOM < 9 "
    SQLX$ = SQLX$ + " AND NUME_OCOM = " & NUOCO&
    SQLX$ = SQLX$ + " AND CANT_REC > 0"
    Call ABRECOMPRAS
    Dim OCOMTMP As ADODB.Recordset
    Set OCOMTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With OCOMTMP
      Do While Not .EOF
         Ci% = XVALO(!cod_inte)
         CONDI$ = "COD_INTE = " & Ci% & " AND NUME_OCOM = " & NUOCO&
         If CantRegistros("RESERVA", CONDI$) > 0 Then
            RT% = 1
            GoTo 99
         End If
      .MoveNext
      Loop
    End With
    OCOMTMP.Close
    Set OCOMTMP = Nothing
    CONDI$ = "NUME_OCOM = " & NUOCO&
    '
    Call ACTUREGISTRO("RESERVA", "NUME_OCOM", CONDI$, 0, REGAF&, "NOMESS")
    '
99  TIENECOTIZARECIB% = RT%
End Function

Function TIENECOTIZA%(NUOCO&)
    '
    RT% = 0
    SQLX$ = "SELECT * FROM OCOMDETA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE STAT_OCOM < 9 "
    SQLX$ = SQLX$ + " AND NUME_OCOM = " & NUOCO&
    Call ABRECOMPRAS
    Dim OCOMTMP As ADODB.Recordset
    Set OCOMTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With OCOMTMP
      If .BOF = True And .EOF = True Then GoTo 99
      .MoveFirst
      Do While Not .EOF
        Ci% = XVALO(!cod_inte)
        CONDI$ = "COD_INTE = " & Ci% & " AND NUME_OCOM = " & NUOCO&
        If CantRegistros("RESERVA", CONDI$) > 0 Then
           RT% = 1
           GoTo 99
        End If
        .MoveNext
      Loop
    End With
    OCOMTMP.Close
    Set OCOMTMP = Nothing
    CONDI$ = "NUME_OCOM = " & NUOCO&
    '
    Call ACTUREGISTRO("RESERVA", "NUME_OCOM", CONDI$, 0, REGAF&, "NOMESS")
    '
99  TIENECOTIZA% = RT%
End Function

Function TIECONSIGENTR%(NUOCO&) 'VERIFICA SI EL MATERIAL DE LA ORDEN DE COMPRA FUE CONSIGNADO Y ENTREGADO
    '
    TIECONSIGENTR% = 0
    NOCOM$ = TRIM$(Str$(NUOCO&))
    While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
    NOCOM$ = "OC-" + Left$(NOCOM$, 6)
    CONDI$ = "NRO_OCOM LIKE '" & NOCOM$ & "%' AND CANT_ENTREG > 0"
    REGS& = CantRegistros&("MACONSIG", CONDI$)
    If REGS& > 0 Then
        TIECONSIGENTR% = 1
    End If
    '
End Function



Private Sub mnuAjusCons_Click()
   'AjuConsig.Show 1

End Sub

Private Sub RecibeArchStockMYM_Click()
    Archivo$ = "ARCHIVO_TRANSF_MIN_MAX"
    Call RECIBIRARCHIVO(Archivo$)
End Sub
Sub RECIBIRARCHIVO(Archivo$, Optional RETORNOOK%, Optional MODOMESS$)
'CARPETASINCRO$ = "X:\ARCONSQL\FLEXOFT\TRANSFESUCU\COMODORO"
    RETORNOOK% = 0
    'CARPETASINCRO$ = TRIM$(Control("", "CARPEINT"))
    CARPETASINCRO$ = RutaCarpetaLocalCompraInterna$
    If CARPETASINCRO$ = "" Then
        Call COMUNI("Falta Definir Carpeta de Intercambio\\Verifique o Cambie Configuración en Setup de Compras (Valida Stocks y Pedidos)")
        RETORNOOK% = -1
        Exit Sub
    End If
    '
    TRANSF% = 0
    'ASIGNO PROPIEDADES AL FILE1
    On Error Resume Next
    File1.Path = "C:\" 'PARA QUE LIMPIE LO ANTERIOR SI NO CAMBIA SE QUEDA SIEMPRE CON LO MISMO
    On Error GoTo 0
    File1.Path = CARPETASINCRO$
    File1.Pattern = "*.txt"
    TX$ = ""
    List2.Clear
    'RECORRO EL FILE1 Y LE PONGO LA FECHA DEL ARCHIVO POR DELANTE PARA QUE ORDENE CORRECTAMENTE
    For n = 0 To File1.ListCount - 1
      NombreArchivo$ = TRIM$(Left$(File1.List(n), Len(File1.List(n)) - 4))
      FechaArchivo$ = TRIM$(FileDateTime(CARPETASINCRO$ + "\" + NombreArchivo$ + ".txt"))
      Call REPLA(TX$, AJUSTI$(FechaArchivo$ + " | ", 48), 1, 48)
      Call REPLA(TX$, AJUSTI$(NombreArchivo$, 48), 50, 48)
      List2.AddItem TX$
    Next n
    '
    If n < 1 Then
      If UCase$(MODOMESS$) <> "NOMESS" Then
        Call COMUNI("No Hay Archivos de Transferencia Pendientes"): Exit Sub
      End If
    End If
    'SEGUN EL ARCHIVO TOMA LA TABLA CORRESPONDIENTE (VARIABLE ARCHIVO$ ES LA QUE VIENE DE PARAMETRO)
    If Archivo$ = "ARCHIVO_TRANSF_MIN_MAX" Then TABLAX$ = "StockSugXSuc"
    If Archivo$ = "ARCHIVO_TRANSF_EXCEPCION" Then TABLAX$ = "Excepcion"
    If Archivo$ = "ARCHIVO_GENERO_PEDIDO" Then TABLAX$ = "OCOMENCA"
    '
    'GRABA CADA ARCHIVO EN SQL Y SI SI GRABA CORRECTAMENTE SE ELIMINA.
    For i& = 0 To List2.ListCount - 1
      POS1% = InStr(TX$, "|")
      FechaArchivo$ = Mid$(List2.List(i&), 1, POS1% - 2)
      NombreArchivo$ = TRIM$(Mid$(List2.List(i&), POS1% + 2))
      If InStr(NombreArchivo$, Archivo$) > 0 Then
          'VALIDO QUE NO SEA MAYOR A 42 POR QUE SI NO LO TRUNCA Y CUANDO SIGUE DA ERROR POR QUE EL ARCHIVO NO EXISTE
          Call GrabarArchivo(CARPETASINCRO$ + "\" + NombreArchivo$ + ".txt", TABLAX$, "NOMESS") 'GRABA EN LA BASE DE DATOS SQL
          TRANSF% = 1
      End If
    Next i&
    '
    If UCase$(MODOMESS$) <> "NOMESS" Then
        If TRANSF% > 0 Then
          Call COMUNI("Importación Finalizada")
        Else
          Call COMUNI("No Hay Archivos de Transferencia Pendientes")
        End If
    End If

End Sub



Private Sub RecibirExcepcion_Click()
    Archivo$ = "ARCHIVO_TRANSF_EXCEPCION"
    Call RECIBIRARCHIVO(Archivo$, RETOK%, "NOMESS")

End Sub


   Sub DESCARGAR_ARCHIVO_PEDIDO_ACOMPRA(Archivo$)

        'CARPETASINCRO$ = TRIM$(Control("", "CARPEINT"))
        CARPETASINCRO$ = RutaCarpetaLocalCompraInterna$

        Ruta$ = CARPETASINCRO$
        If Ruta$ = "" Then
           Call COMUNI("Imposible Continuar.\Falta Definir Carpeta de Intercambio")
           Exit Sub
        End If
        
        'RUTA$ = TRIM$(CONTROL("", "CARPEINT"))
        If Ruta$ <> "" Then
            'ASIGNO PROPIEDADES AL FILE1
                On Error Resume Next
                File1.Path = "C:\"
                On Error GoTo 0
                File1.Path = Ruta$
                File1.Pattern = "*.txt"
                TX$ = ""
                List2.Clear
                'RECORRO EL FILE1 Y LE PONGO LA FECHA DEL ARCHIVO POR DELANTE PARA QUE ORDENE CORRECTAMENTE
                For n = 0 To File1.ListCount - 1
                  NombreArchivo$ = TRIM$(Left$(File1.List(n), Len(File1.List(n)) - 4))
                  FechaArchivo$ = TRIM$(FileDateTime(Ruta$ + "\" + NombreArchivo$ + ".txt"))
                  Call REPLA(TX$, AJUSTI$(FechaArchivo$ + " | ", 48), 1, 48)
                  Call REPLA(TX$, AJUSTI$(NombreArchivo$, 48), 50, 48)
                  List2.AddItem TX$
                Next n
                '
                'SEGUN EL ARCHIVO TOMA LA TABLA CORRESPONDIENTE
                If InStr(Archivo$, "Ocompra_Sucursal") > 0 Then TABLAX$ = "OCOMENCA"
                
                '
                'GRABA CADA ARCHIVO EN SQL Y SI SI GRABA CORRECTAMENTE SE ELIMINA.
                For i& = 0 To List2.ListCount - 1
                  POS1% = InStr(TX$, "|")
                  FechaArchivo$ = Mid$(List2.List(i&), 1, POS1% - 2)
                  NombreArchivo$ = TRIM$(Mid$(List2.List(i&), POS1% + 2))
                  If UCase("ARCHIVO_GENERO_PEDIDO") = Left$(TRIM$(UCase$(NombreArchivo$)), Len(Archivo$)) Then
                    TABLAX$ = "OCOMENCA"
                    Call GrabarArchivo(Ruta$ + "\" + NombreArchivo$ + ".txt", TABLAX$, "NOMESS") 'GRABA EN LA BASE DE DATOS SQL
                  End If
                Next i&

        Else
            Call COMUNI("Imposible Continuar.\Falta Definir Carpeta de Intercambio")
            Exit Sub
        End If
        '
End Sub


Private Sub RECIBOPEDICOMPRA_Click()
   Call RECIBIRARCHIVO("ARCHIVO_GENERO_PEDIDO", RESOK%, "NOMESS")
End Sub



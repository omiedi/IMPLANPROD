VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FORECEP07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Materiales - Altas de Stock en Depósitos"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   9525
   Icon            =   "FORECEP07.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command20 
      Caption         =   "Prueba QR"
      Height          =   255
      Left            =   7200
      TabIndex        =   62
      Top             =   2880
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   57
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "CONSULTAS Y LISTADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5790
      Left            =   4440
      TabIndex        =   2
      Top             =   120
      Width           =   4110
      Begin VB.Frame Frame3 
         BackColor       =   &H00CDDADA&
         Caption         =   "CONTROL DE RECEPCIONES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2820
         Left            =   210
         TabIndex        =   23
         Top             =   2835
         Width           =   3690
         Begin VB.PictureBox Picture13 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   945
            ScaleHeight     =   615
            ScaleWidth      =   2400
            TabIndex        =   53
            Top             =   1995
            Width           =   2430
            Begin VB.CommandButton Command18 
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
               Left            =   0
               Picture         =   "FORECEP07.frx":030A
               Style           =   1  'Graphical
               TabIndex        =   54
               ToolTipText     =   "Recepciones Rechazadas por Rango de Fechas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Rechazos por Rango de Fechas"
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
               Left            =   720
               TabIndex        =   55
               Top             =   90
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture12 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   735
            ScaleHeight     =   615
            ScaleWidth      =   2400
            TabIndex        =   43
            Top             =   1470
            Width           =   2430
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
               Height          =   645
               Left            =   0
               Picture         =   "FORECEP07.frx":0614
               Style           =   1  'Graphical
               TabIndex        =   45
               ToolTipText     =   "Recepciones sin O/C."
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Recepciones sin O/Compras's"
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
               Left            =   720
               TabIndex        =   44
               Top             =   60
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   525
            ScaleHeight     =   615
            ScaleWidth      =   2400
            TabIndex        =   28
            Top             =   945
            Width           =   2430
            Begin VB.CommandButton Command16 
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
               Left            =   0
               Picture         =   "FORECEP07.frx":075E
               Style           =   1  'Graphical
               TabIndex        =   47
               ToolTipText     =   "Recepciones con Rechazo."
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Recepciones con Rechazos"
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
               TabIndex        =   29
               Top             =   60
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture9 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   300
            ScaleHeight     =   615
            ScaleWidth      =   2400
            TabIndex        =   24
            Top             =   405
            Width           =   2430
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
               Height          =   645
               Left            =   0
               Picture         =   "FORECEP07.frx":0A68
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   0
               Width           =   560
            End
            Begin VB.Label LABEL5 
               BackStyle       =   0  'Transparent
               Caption         =   "Recepciones por Rango de Fechas"
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
               TabIndex        =   26
               Top             =   60
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00CDDADA&
         Caption         =   "CONSULTA DE PENDIENTES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Left            =   210
         TabIndex        =   13
         Top             =   315
         Width           =   3690
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   615
            ScaleWidth      =   2190
            TabIndex        =   20
            Top             =   1545
            Width           =   2220
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
               Height          =   645
               Left            =   -20
               Picture         =   "FORECEP07.frx":0D72
               Style           =   1  'Graphical
               TabIndex        =   21
               ToolTipText     =   "Material en Consignación"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Material en Consignación"
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
               TabIndex        =   22
               Top             =   90
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   735
            ScaleHeight     =   615
            ScaleWidth      =   2190
            TabIndex        =   17
            Top             =   1005
            Width           =   2220
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
               Height          =   645
               Left            =   -20
               Picture         =   "FORECEP07.frx":107C
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "Ordenes de Compra Pendientes de Recepción"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Pedidos y Ordenes de Compra"
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
               TabIndex        =   19
               Top             =   70
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   420
            ScaleHeight     =   615
            ScaleWidth      =   2190
            TabIndex        =   14
            Top             =   510
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
               Height          =   645
               Left            =   0
               Picture         =   "FORECEP07.frx":1386
               Style           =   1  'Graphical
               TabIndex        =   15
               ToolTipText     =   "Detalle por Boleta de Recepción"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Boletas de Recepción"
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
               Left            =   720
               TabIndex        =   16
               Top             =   45
               Width           =   1350
            End
         End
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00CDDADA&
      Caption         =   "SALIDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   120
      TabIndex        =   48
      Top             =   5250
      Width           =   4110
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Pantalla"
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
         Left            =   600
         TabIndex        =   50
         Top             =   280
         Width           =   1380
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00CDDADA&
         Caption         =   "Impresión"
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
         Left            =   2310
         TabIndex        =   49
         Top             =   280
         Value           =   -1  'True
         Width           =   1170
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H80000016&
      Height          =   645
      Left            =   1155
      ScaleHeight     =   585
      ScaleWidth      =   2265
      TabIndex        =   36
      Top             =   6720
      Width           =   2325
      Begin VB.CommandButton Command9 
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
         Picture         =   "FORECEP07.frx":1690
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   0
         Width           =   560
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Devolución       de Materiales"
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
         TabIndex        =   38
         Top             =   105
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture11 
      BackColor       =   &H00808080&
      ForeColor       =   &H00000000&
      Height          =   6210
      Left            =   8640
      ScaleHeight     =   6150
      ScaleWidth      =   1005
      TabIndex        =   30
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   59
         Top             =   5640
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   60
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "FORECEP07.frx":17DA
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   3530
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Tr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FORECEP07.frx":1AE4
         TabIndex        =   51
         ToolTipText     =   "Sub-Sistema de Trazabilidad"
         Top             =   4800
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command12 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FORECEP07.frx":1DEE
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Ver y Re-Imprimir B.Recepción"
         Top             =   1550
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command8 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FORECEP07.frx":20F8
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Acceso Directo a otras Aplicaciones"
         Top             =   2880
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   105
         Picture         =   "FORECEP07.frx":2402
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   805
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FORECEP07.frx":270C
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Terminar - Salir de esta Aplicación"
         Top             =   120
         Width           =   650
      End
      Begin VB.CommandButton Command14 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FORECEP07.frx":2856
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Configuración de Comportamiento y Formularios"
         Top             =   2200
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.CommandButton Command10 
         Caption         =   "QC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FORECEP07.frx":2C98
         TabIndex        =   31
         ToolTipText     =   "Calificación de Proveedores"
         Top             =   4150
         Visible         =   0   'False
         Width           =   650
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "INGRESO A STOCK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4110
      Begin VB.Frame Frame7 
         BackColor       =   &H00CDDADA&
         Caption         =   "GESTIÓN DE COMPRAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   210
         TabIndex        =   39
         Top             =   2200
         Width           =   3690
         Begin VB.PictureBox Picture10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   600
            ScaleHeight     =   615
            ScaleWidth      =   2400
            TabIndex        =   40
            Top             =   420
            Width           =   2430
            Begin VB.CommandButton Command13 
               Caption         =   "OC"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Left            =   0
               TabIndex        =   41
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Consulta de Ordenes de Compra"
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
               TabIndex        =   42
               Top             =   90
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00CDDADA&
         Caption         =   "ENTRADAS DE FÁBRICA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   210
         TabIndex        =   9
         Top             =   3630
         Width           =   3690
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   630
            ScaleHeight     =   615
            ScaleWidth      =   2400
            TabIndex        =   10
            Top             =   430
            Width           =   2430
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
               Height          =   645
               Left            =   0
               Picture         =   "FORECEP07.frx":2FA2
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Por Propia Producción"
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
               TabIndex        =   12
               Top             =   90
               Width           =   1005
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00CDDADA&
         Caption         =   "RECEPCIONES EXTERNAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1730
         Left            =   210
         TabIndex        =   1
         Top             =   315
         Width           =   3690
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   840
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   6
            Top             =   900
            Width           =   2325
            Begin VB.CommandButton Command19 
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
               Left            =   1760
               Picture         =   "FORECEP07.frx":32AC
               Style           =   1  'Graphical
               TabIndex        =   61
               ToolTipText     =   "Clasificación de Material Recibido"
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
               Height          =   645
               Left            =   0
               Picture         =   "FORECEP07.frx":35B6
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Inspección y Aprobación"
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
               Left            =   720
               TabIndex        =   8
               Top             =   90
               Width           =   975
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   525
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   3
            Top             =   420
            Width           =   2325
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
               Height          =   645
               Left            =   0
               Picture         =   "FORECEP07.frx":38C0
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Boleta de Recepción (Provisoria)"
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
               Left            =   630
               TabIndex        =   5
               Top             =   20
               Width           =   1590
            End
         End
         Begin VB.Image Image1 
            Height          =   375
            Left            =   3000
            Top             =   360
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2940
            TabIndex        =   52
            Top             =   420
            Visible         =   0   'False
            Width           =   645
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "FORECEP07.frx":3BCA
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "FORECEP07.frx":3DFE
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label NORVENTA 
      Caption         =   "NORVENTA"
      Height          =   225
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuBoletaRecepciónProvisoria 
         Caption         =   "Boleta de Recepción (Provisoria)"
      End
      Begin VB.Menu mnuInspyAprobacion 
         Caption         =   "Inspección y Aprobación"
      End
      Begin VB.Menu mnuAnulaBolRecProv 
         Caption         =   "Anular Boleta de Recepción Provisoria"
      End
      Begin VB.Menu mnuAprobaciónBR 
         Caption         =   "Aprobación de B-R's"
      End
      Begin VB.Menu s4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdjuntarDocumentosExternos 
         Caption         =   "Adjuntar Documentos Externos"
      End
      Begin VB.Menu mnuModificarRemitoEnRecepcion 
         Caption         =   "Modificar Remito de Proveedor en Recepción"
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClasificacionMaterialesRecibidos 
         Caption         =   "Clasificación de Material Recibido"
      End
      Begin VB.Menu mnutransferirasector 
         Caption         =   "Transferencia de Materia Prima a Sector Productivo"
         Visible         =   0   'False
      End
      Begin VB.Menu mnudevtransf 
         Caption         =   "Devoluciones de Sector Productivo"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuIngresoDatoDespacho 
         Caption         =   "Ingreso de Datos de Despacho a Recepción"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuVerificarDatosSegunSerie 
         Caption         =   "Verificar Datos de Importación Vs N.Serie"
         Visible         =   0   'False
      End
      Begin VB.Menu lnTransacciones1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuCYL 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuBolRec 
         Caption         =   "Boletas de Recepción"
      End
      Begin VB.Menu mnupedidosyordenesdecompra 
         Caption         =   "Pedidos y Ordenes de Compra"
      End
      Begin VB.Menu mnumaterialenconsignacion 
         Caption         =   "Material en Consignación"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecepcionMaterialAClasificar 
         Caption         =   "Recepciones de Material a Clasificar"
      End
      Begin VB.Menu mnuListadoDeClasificaciones 
         Caption         =   "Listado de Recepciones Clasificadas"
      End
      Begin VB.Menu lineConsultas1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecepcionesxrangofecha 
         Caption         =   "Recepciones x Rango de Fechas"
      End
      Begin VB.Menu mnurecepcionesconrechazo 
         Caption         =   "Recepciones con Rechazo"
      End
      Begin VB.Menu mnurecepcionessinocompra 
         Caption         =   "Recepciones Sin O/Compra's"
      End
      Begin VB.Menu mnurechazosxfecha 
         Caption         =   "Rechazos Por Rango de Fechas"
      End
      Begin VB.Menu mnuReporteVarios 
         Caption         =   "Reportes Varios"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
      Begin VB.Menu mnuReimpRecep 
         Caption         =   "Ver y Re-Imprimir Recepciones"
      End
      Begin VB.Menu reimpremidevo 
         Caption         =   "Re-Impresion de Remito Devolución"
      End
      Begin VB.Menu mnuReimpEtiq 
         Caption         =   "Re-Impresión de Etiquetas"
      End
   End
   Begin VB.Menu mnuUtil 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuConfigReportes 
         Caption         =   "Configuración de Acceso a Reportes"
      End
      Begin VB.Menu ControlRecepvscompras 
         Caption         =   "Control Recepciones vs. Compras"
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuración"
   End
   Begin VB.Menu mnuAccDir 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuTrazabilidad 
         Caption         =   "Trazabilidad"
         Visible         =   0   'False
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCalifProveedores 
         Caption         =   "Calificación de Proveedores"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "FORECEP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HAYMATERIALIB%
Dim MODORECHA% '\\\OT-05-0723-WAR-23/11/05///
Dim CANCELX1% '\\\OT-06-0178-WAR-12/05/06///
Dim TXTRAZA_BR$
Public OrdenesSeleccionadas$
Public OrdenesSeleccionadasConsignaciones$
Dim EsReteplast%

Sub AGREGA_A_TABLA_IMPORTADO(ARCHICA$)
      Dim obj As New Art_Importados
      BB = ULTREG&(ARCHICA$)
      For YX& = 1 To ULTREG&(ARCHICA$)
        YY$ = REGLEIDO$(ARCHICA$, YX&)
        CI% = CVI(Mid$(YY$, 83, 2))
        If CI% > 0 Then
           CODI$ = CODEXT$(CI%)
           Call obj.VeriImportado(CODI$) 'LEE LA INFO Y SE LA ASIGNA AL OBJETO OBJ.ESIMPORRTADO
            If obj.EsImportado < 1 Then
              Call obj.GrabarItemImportado(CODI$)
            End If
         End If
      Next YX&
      
End Sub

Private Sub APRORECEP()
   '
   If DERACCE%("APRORECE", "Aprobación de Boleta de Recepción") < 2 Then GoTo 99
   '
   Call COPYSTRU("LILOTSE", "BRECEAP")
   FIN& = ULTREG("BRECEPEN")
   Call SETULTREG("BRECEAP", 0)
   Call TRALOCAL("BRECEAP", "")
   
   'CARGO LA LISTA DE SELECCION DESDE EL ARCHIVO ORFAPEN CON LA OF PENDIENTES DE APROBACION
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     TXTX$ = REGLEIDO$("BRECEPEN", U&)
     CMOV$ = Mid$(TXTX$, 21, 2)
     CIII% = CVI(Mid$(TXTX$, 3, 2))
     If CIII% > 0 Then
           DOPRI$ = TRIM$(Mid$(TXTX$, 5, 16))
           DESCRI$ = AJUSTI$(DESCRIT0$(CIII%), 16)
           CANTINGRE = CVD(Mid$(TXTX$, 88, 8))
           CANTSALID = CVD(Mid$(TXTX$, 96, 8))
           CANTI = CANTINGRE - CANTSALID
5           If CANTI > 0.05 Then
             Call REPLA(TTXX$, TRIM$(DOPRI$), 1, 16)
             Call REPLA(TTXX$, CODEXT$(CIII%), 18, 16)
             Call REPLA(TTXX$, AJUSTI$(DESCRIT0$(CIII%), 24), 36, 24)
             Call REPLA(TTXX$, FORMATNUM$(CANTI, "F10.1"), 62, 10)
             Call REGAPP("!BRECEAP", TTXX$)
           End If
     End If
   Next U&
   Call CIERRARCH("!BRECEAP")
   Call TRACENTRAL("BRECEAP", "")
   If U& < 2 Then Call COMUNI("No Hay Recepciones Para Autorizar"): Exit Sub
   Call FRESHECHO("BRECEAP")
   '
   'PRESENTO LISTA DE SELECCION
50 Call SELECHO("BRECEAP/Recepciones Pendientes de Aprobación")
   '
   DOPRI$ = TRIM$(VALACT1$("BRECEAP"))
   REFESEL$ = TRIM$(VALACT2$("BRECEAP"))
   If DOPRI$ = "" Then Exit Sub
   '
   'CONSULTA DE CONFIRMACION DE O-FABRICACION
   NUORIT% = 0
   POSI% = 1
   ' **** GRABA LOS MOVIMIENTOS DE STOCK
   CANTINOM = 0
   RFF$ = RECFILT$("BRECEPEN", 4, DOPRI$)
   'LEO LA INFORMACION DEL REGISTRO SELECCIONADO
55 REG& = CVS(Mid$(RFF$, POSI%, 4))
   TX1$ = REGLEIDO$("BRECEPEN", REG&)
   CIII% = CVI(Mid$(TX1$, 3, 2))
   CIXI% = CODINT%(Mid$(REFESEL$, 1, 16))
   If CIII% <> CIXI% Then
        POSI% = POSI% + 4
        REG& = CVS(Mid$(RFF$, POSI%, 4))
        GoTo 55
   End If
   REFE$ = TRIM$(Mid$(TX1$, 43, 30))
   FF% = CVI(Mid$(TX1$, 1, 2))  'TOMA LA FECHA EN QUE SE GENERO, EL CIERRE, NO LA DE LA APROBACION
   LTX$ = TRIM$(Mid$(TX1$, 5, 16))
   CMOV$ = TRIM$(Mid$(TX1$, 21, 2))
   DOPRICO$ = TRIM$(Mid$(TX1$, 23, 10))
   DOSECO$ = TRIM$(Mid$(TX1$, 33, 10))
   DEPO% = Asc(Mid$(TX1$, 83, 1))
   NC% = CVI(Mid$(TX1$, 86, 2))
   CANTIENT1 = CVD(Mid$(TX1$, 88, 8))
   CANTISAL1 = CVD(Mid$(TX1$, 96, 8))
   CANTI = CANTIENT1 - CANTISAL1
   '
   'OBTIENE CANTIDAD NOMINAL-APROBADA-FALTANTE-RECHAZADA
   CANTIAPRO = CANTI
   CANMODIF = XVALO(InputBox("Cantidad Aprobada", "Aprobación de Recepción", Str$(CANTIAPRO)))
   '
   If CANMODIF > 0 Then
       OPXX% = COMALTER%("Confirma Cantidades Aprobadas Para '" & CODEXT$(CIII%) & "' Por " _
       & TRIM$(FORMATNUM$(CANMODIF, "F10.1")) & "\\No, Cancelar\Si, Continuar")
       If OPXX% = 2 Then
            'GENERA EL MOVIMIENTO DE STOCK TANTO DE LA OF, COMO EL CONSUMO DE OF O EL SCRAP
            CANTI = CANMODIF
            Call MOVISTOK(FF%, CIII%, LTX$, CMOV$, DOPRICO$, DOPRICO$, DOSECO$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NC% * (-1), DEPO%, CANTI)
            'ACTUALIZA CANTIDAD APROBADA
            CONDI$ = "IdenLote='" & DOPRI$ & "' AND Cod_Inte=" & CIII%
            Call ACTUREGISTRO("BOLRECEP", "CanToApro", CONDI$, CANMODIF, REGAF&)
            '
            CANTNOMI = XVALO(VALOBADA("BOLRECEP", "CANTNOM", CONDI$))
            NUMOITOC% = XVALO(VALOBADA("BOLRECEP", "NroItemOC", CONDI$))
            
            NUOC& = XVALO(VALOBADA("BOLRECEP", "NumOC", CONDI$))
            If NUOC& > 0 Then
                CAPRO = CANMODIF: NUOR% = 0
                'ACTUALIZA ORDEN DE COMPRA
                CONDI$ = "Nume_Ocom = " & NUOC&
                CONDI$ = CONDI$ + " AND Cod_Inte = " & CIII%
                CONDI$ = CONDI$ + " AND Stat_Ocom < 8 "
                If NUMOITOC% = 0 Then
                    CONDI$ = CONDI$ + " And Cant_Rec >= " & TRIM$(FORMATNUM$(Abs(CANTINOM - 0.005), "F15.3"))
                Else
                    CONDI$ = CONDI$ + " AND NUORD_ITEM=" & CStr(NUMOITOC%) ' REGISTRA EL RECHAZO EN EL ITEM QUE CORRESPONDE
                End If
                
                SQLX$ = "SELECT * FROM OCOMDETA WHERE " & CONDI$
                Dim Rst As ADODB.Recordset
                Set Rst = New ADODB.Recordset
                Rst.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                With Rst
                    If Not (.EOF And .BOF) Then
                        If CAPRO > 0.05 Then
                            COEFUNI = XVALO(!Coef_Unids)
                               If COEFUNI < 0.0001 Then COEFUNI = 1
                            CANCOM = XVALO(!cant_ocom)
                            
                            'QUITA LA CANTIDAD NOMINAL APROBADA ORIGINALMENTE Y SUMA LA MODIFICADA
                            CANTORECI = XVALO(!Cant_Rec) - (CANTNOMI / COEFUNI) + (CAPRO / COEFUNI)
                            
                            'GRABA CANTIDAD RECIBIDA
                            !Cant_Rec = CANTORECI
                            
                            .Update
                        End If
80                      .MoveNext
                    End If
                    Rst.Close
                    Set Rst = Nothing
                End With
            End If
            'BORRO EL REGISTRO DEL ARCHIVO DE PENDIENTES DE APROBACION
            Call FILTERDELETE("BRECEPEN", 4, AJUSTI(DOPRI$, 16))
            '
            Call COMUNI("Aprobación Procesada")
        End If
   End If
   '
99 Screen.MousePointer = 1
   '
End Sub

Function canticampos(TABLA$)
    SQLX$ = " SELECT * FROM " & TABLA$
    SQLX$ = SQLX$ + " WITH(NOLOCK)"
    Set RS = READSET(SQLX$)
    With RS
      canticampos = RS.Fields.Count
      
    End With
End Function

Sub GENERA_OF_NSERIE_PROD_IMPORTADOS(NROBOL&)
        GENEROOF% = 0
        FIN& = ULTREG&("!BOLREDET")
        Dim CANTIAFAB()
        ReDim Preserve CANTIAFAB(1)
        '
        Dim obj As New Art_Importados
        'LEO EL BOLREDET Y CARGO UN VECTOR CON LA CANTIDAD A FABRICAR DE LOS PRODUCTOS TRAZABLES
        For U& = 1 To FIN&
          Call TRACE(20, U&, FIN&)
          X$ = REGLEIDO$("!BOLREDET", U&)
          CI% = CVI(Mid$(X$, 83, 2))
          If ESTRAZABLE%(CI%, 1) = 2 Then
            Call obj.VeriImportado(CODEXT$(CI%)) 'LEE LA INFO Y SE LA ASIGNA AL OBJETO OBJ.ESIMPORRTADO
            If obj.EsImportado > 0 Then
               If UBound(CANTIAFAB) < CI% Then ReDim Preserve CANTIAFAB(CI%)
               CANOM# = CVD(Mid$(X$, 89, 8))
               APRO# = CVD(Mid$(X$, 163, 8))
               CANTIAFAB(CI%) = CANTIAFAB(CI%) + APRO#
            End If
          End If
        Next U&
        '
        'ACA GENERAR LAS OF RECORRIENDO EL VECTOR
        
        FIN& = UBound(CANTIAFAB)
        Label12.Visible = True
        For i% = 1 To UBound(CANTIAFAB)
            k& = i%
            Call TRACE(20, k&, UBound(CANTIAFAB))
            If CANTIAFAB(i%) > 0 Then
                 Label12.Caption = k&
                 GENEROOF% = 1
                 Call GENEOF(i%, 0, APRO#, HOY(HO$), 0, "", RET_NORFA$, 1, "Eq.Imp - BR-", TRIM$(SAFETEXT$(NROBOL&)))
                 Call ACTUECO(i%, APRO#, RET_NORFA$, "Eq.Imp - BR-" & TRIM$(SAFETEXT$(NROBOL&)))
            End If
        Next i%
        '
        Label12.Visible = False
        Erase CANTIAFAB
        Set obj = Nothing
        'LOS NUMEROS DE SERIE SE GENERAN DE MANERA POSTERIOR
'        If EXISTE%("TRANSE07.EXE") > 0 And ESTRAZABLE(CI%) = 2 Then
'          Call CONECRUN("TRANSE07/GENSEOF" + RET_NORFA$, "Generador de Numeros de Serie")
'          Screen.MousePointer = 1
'          Exit Sub
'        End If

'        GENUMSE07.Label13 = RET_NORFA$
'        If GENEROOF% > 0 Then GENUMSE07.Show 1
End Sub


Sub GRABAR_CONTEXTO_PROPUESTO(NC, NROBOL&, FF%)
'CI/F0
';CODIGO/A16
';DESCRIPCION/A32
';U.M./A4
';Q.INGR/F10.1
';Q.IDEAL/F10.1
';O.COMPRA/F8
';ENTR.SOL./D8
';PENDIENTE/F10.1
';Q.IDEAL/F10.1
';LOTECO/A0
';ORDEN/F0".
        RAPRO$ = Left$(rasocli$(-1 * NC), 64)
        Call CREATABLA_CONTRECE
        With VENTABNEW.MSF
          For i& = 1 To .Rows - 1
           CI1% = XVALO(.TextMatrix(i&, 1))
           CODI$ = .TextMatrix(i&, 2)
           DESCRIX$ = Left$(.TextMatrix(i&, 3), 64)
           UME$ = .TextMatrix(i&, 4)
           CantIngUser = XVALO(.TextMatrix(i&, 5))
           CantPropuEdit = XVALO(.TextMatrix(i&, 6)) 'CANTIDAD PROPUESTA EDITABLE
           NCompra& = XVALO(.TextMatrix(i&, 7))
           EntSolic = .TextMatrix(i&, 8)
           CAPENCOMP = XVALO(.TextMatrix(i&, 9))
           CantOrigProp = XVALO(.TextMatrix(i&, 10))
           LOTE$ = Left$(.TextMatrix(i&, 11), 12)
           ORDENIT% = XVALO(.TextMatrix(i&, 12))
           
           'AGREGA UN NUEVO REGISTRO
           SQLX2$ = "INSERT INTO CONTRECE"
           SQLX2$ = SQLX2$ + " (NumeBoRe,IdenLote,FechRecep,Cod_Inte,Cod_Exte,DescripItem,Unidad,CantNom,"
           SQLX2$ = SQLX2$ + " CanPropEdit,CanProOrig,NumProv,RaSo_Prov,NumOC,NroItemOC,codiempr,CAPENCOMP)"
           SQLX2$ = SQLX2$ + " VALUES(" & NROBOL&
           SQLX2$ = SQLX2$ + " ,'" & LOTE$ & "'"
           SQLX2$ = SQLX2$ + " ,'" & CVDAT(FF%) & "'"
           SQLX2$ = SQLX2$ + " ," & CI1%
           SQLX2$ = SQLX2$ + " ,'" & CODI$ & "'"
           SQLX2$ = SQLX2$ + " ,'" & DESCRIX$ & "'"
           SQLX2$ = SQLX2$ + " ,'" & UME$ & "'"
           SQLX2$ = SQLX2$ + " ," & CantIngUser
           SQLX2$ = SQLX2$ + " ," & CantPropuEdit
           SQLX2$ = SQLX2$ + " ," & CantOrigProp
           SQLX2$ = SQLX2$ + " ," & NC
           SQLX2$ = SQLX2$ + " ,'" & RAPRO$ & "'"
           SQLX2$ = SQLX2$ + " ," & NCompra&
           SQLX2$ = SQLX2$ + " ," & ORDENIT%
           SQLX2$ = SQLX2$ + " ," & CodiEmp%
           SQLX2$ = SQLX2$ + " ," & CAPENCOMP
           SQLX2$ = SQLX2$ + ")"
           FLEXCONN.Execute (SQLX2$)
          Next i&
        End With

End Sub

Private Sub GuardaFotoMaterialRecibido(LoteRecepcion As String)
   
    If LoteRecepcion <> "" Then
        RutaGuardaCaptura$ = TRIM$(Control("", "RUTARGRA"))
        If RutaGuardaCaptura$ = "" Then RutaGuardaCaptura$ = LUDAT$
        RutaGuardaCaptura$ = RutaGuardaCaptura$ & "\" & LoteRecepcion & ".JPG"
        
        If EXISTE%(RutaGuardaCaptura$) > 0 Then
            CondicionSql$ = "DOCUMENTOORIGEN='" & LoteRecepcion & "'"
            If CantRegistros("INFOCALIDAD", CondicionSql$) = 0 Then
                Call CreaRegistro("INFOCALIDAD", "DOCUMENTOORIGEN", LoteRecepcion)
            End If
            Call ACTUREGISTRO("INFOCALIDAD", "PATHFOTOCALIDAD", CondicionSql$, RutaGuardaCaptura$, REGAF&)
        End If
        
    End If

End Sub

Function ComprasPendientesDeLoConsignado%(Proveedor&, ArchivoRecepcion$)
    
    ' RECORRO EL ARCHIVO BOLREDET EN BUSCA DE LAS ORDENES DE COMPRAS A RECIBIR
    
    Screen.MousePointer = 11
    Call ABRESTOCKS
    RecepcionesEncontradas% = 0
    Campos$ = "Código Consignado/A15;Descripción/A24;O-Compra/I10;Proveedor/A24;Saldo O-Compra/F14.2;Uso según fórmula/F18.2"
    Call CARGA_ENCABEZADO(MUESTRADATOS9.GRID, Campos$, MUESTRADATOS9, 1, 0, ANTOT, , 1)
    MUESTRADATOS9.Caption = "Consignaciones Pendientes de Recepción"
    MUESTRADATOS9.GRID.Redraw = False
    
    For II& = 1 To ULTREG&("!" & ArchivoRecepcion)
            
        XX$ = REGLEIDO$("!" & ArchivoRecepcion, II&)
        CodigoARecibir& = CVI(Mid$(XX$, 83, 2))
        CantidadARecibir# = CVD(Mid$(XX$, 89, 8))
        OrdenDeCompra& = CVS(Mid$(XX$, 151, 4))
        NumeroDeOrden% = XVALO(Asc(Mid$(XX$, 98, 1)))
        FechaOrdenDeCompra% = CVINT(VALOBADA("OCOMENCA", "Femi_Ocom", "Nume_Ocom=" & CStr(OrdenDeCompra&)))
        
        NUMECOMPRA$ = CStr(OrdenDeCompra&)
        ORDEN$ = CStr(NumeroDeOrden%)
        Do While Len(NUMECOMPRA$) < 6: NUMECOMPRA$ = "0" & NUMECOMPRA$: Loop
        Do While Len(ORDEN$) < 2: ORDEN$ = "0" & ORDEN$: Loop
        NUMECOMPRA$ = NUMECOMPRA$ & "-" & ORDEN$
               
        ' SI SE TRATA DE UNA ORDEN DE COMPRA, INGRESA A LA VERIFICACION
        If CantidadARecibir# > 0.05 Then
            ' CONSIGNACIONES DE LA ORDEN DE COMPRA EN BASE A LA ESTRUCTURA DEL PRODUCTO
            SQLX$ = "SELECT CodIConj, CodiElem, UsoBruto FROM FORMULAS WITH(NOLOCK) "
            SQLX$ = SQLX$ + " WHERE CodIConj='" & CStr(CodigoARecibir&) & "'"

            Dim Consignaciones As ADODB.Recordset
            Set Consignaciones = READSET(SQLX$)
            With Consignaciones
                Do While .EOF = False
                    CodigoConsignado& = XVALO(!codielem)
                    USO# = XVALO(!usobruto)
                    Consignado# = CantidadARecibir# * USO#
                    '
                    ' BUSCA LAS ORDENES DE COMPRAS PENDIENTES DE LOS ARTICULOS CONSIGNADOS
                    ConsultaSql$ = "SELECT Nume_Ocom, Npro_Ocom, Cant_Ocom, Cant_Rec FROM OCOMDETA WITH(NOLOCK) WHERE Cod_Inte=" & CStr(CodigoConsignado&) & " AND Stat_Ocom < 3 AND Cant_Ocom - Cant_Rec > 0.05 AND Femi_Ocom <= '" & FETEXCOR1$(FechaOrdenDeCompra%) & "'"
                    Dim ComprasPendientes As ADODB.Recordset
                    Set ComprasPendientes = READSET(ConsultaSql$)
                    With ComprasPendientes
                        Do While Not .EOF
                            '
                            NumeroCompra& = XVALO(!nume_ocom)
                            Proveedor& = XVALO(!Npro_Ocom)
                            CantidadCompra# = XVALO(!cant_ocom)
                            cantidadRecibida# = XVALO(!Cant_Rec)
                            SaldoPendienteCompra# = CantidadCompra# - cantidadRecibida#
                            ARendir# = USO# * CantidadARecibir#
                            '
                            NumeroFila& = NumeroFila& + 1
                            MUESTRADATOS9.GRID.Rows = NumeroFila& + 1
                            MUESTRADATOS9.GRID.TextMatrix(NumeroFila&, 1) = CODEXT$(CInt(CodigoConsignado&))
                            MUESTRADATOS9.GRID.TextMatrix(NumeroFila&, 2) = DESCRIT0$(CInt(CodigoConsignado&))
                            MUESTRADATOS9.GRID.TextMatrix(NumeroFila&, 3) = CStr(NumeroCompra&)
                            MUESTRADATOS9.GRID.TextMatrix(NumeroFila&, 4) = rasocli$((-1) * Proveedor&)
                            MUESTRADATOS9.GRID.TextMatrix(NumeroFila&, 5) = FORMATNUM$(SaldoPendienteCompra#, "F12.2")
                            MUESTRADATOS9.GRID.TextMatrix(NumeroFila&, 6) = FORMATNUM$(Consignado#, "F12.2")
                            '
                            RecepcionesEncontradas% = 1
                            .MoveNext
                            '
                        Loop
                    End With
                    On Error Resume Next
                    ComprasPendiente.Close
                    Set ComprasPendiente = Nothing
                    On Error GoTo 0
                    .MoveNext
                Loop
            End With

            On Error Resume Next
            Consignaciones.Close
            Set Consignaciones = Nothing
            On Error GoTo 0
        
'            ' CONSIGNACIONES DE LA ORDEN DE COMPRA
'            SQLX$ = "SELECT Cod_Inte, Nro_Ocom, cant_entreg, cant_rendid, cant_devu FROM MACONSIG WITH(NOLOCK) "
'            SQLX$ = SQLX$ + " WHERE (Cant_Rendid + Cant_Devu < Cant_Entreg - 0.001 "
'            SQLX$ = SQLX$ + " AND Nro_Ocom NOT LIKE 'AR-%' "
'            SQLX$ = SQLX$ + " AND Nro_Ocom LIKE '%" & NUMECOMPRA$ & "%' "
'            SQLX$ = SQLX$ + " AND Npro_Ocom = " & CStr(Proveedor&) & ") "
'
'            Dim Consignaciones As ADODB.Recordset
'            Set Consignaciones = READSET(SQLX$)
'            With Consignaciones
'                Do While .EOF = False
'                    CodigoConsignado& = XVALO(!COD_INTE)
'                    CantidadConsignada# = XVALO(!cant_entreg)
'                    CantidadRendida# = XVALO(!cant_rendid)
'                    SaldoConsignado# = CantidadConsignada# - CantidadRendida#
'                    '
'                    ConsultaSQL$ = "SELECT Nume_Ocom, Npro_Ocom, Cant_Ocom, Cant_Rec FROM OCOMDETA WITH(NOLOCK) WHERE Cod_Inte=" & CStr(CodigoConsignado&) & " AND Stat_Ocom < 3 AND Cant_Ocom - Cant_Rec > 0.05"
'                    Dim ComprasPendientes As ADODB.Recordset
'                    Set ComprasPendientes = READSET(ConsultaSQL$)
'                    With ComprasPendientes
'                        Do While Not .EOF
'                            '
'                            NumeroCompra& = XVALO(!nume_ocom)
'                            Proveedor& = XVALO(!Npro_Ocom)
'                            CantidadCompra# = XVALO(!cant_ocom)
'                            CantidadRecibida# = XVALO(!Cant_Rec)
'                            SaldoPendienteCompra# = CantidadCompra# - CantidadRecibida#
'                            Uso# = XVALO(VALOBADA("FORMULAS", "UsoBruto", "CODICONJ=" & CStr(CodigoARecibir&) & "AND CODIELEM=" & CStr(CodigoConsignado&), ""))
'                            ARendir# = Uso# * CantidadARecibir#
'                            '
'                            NumeroFila& = NumeroFila& + 1
'                            MUESTRADATOS9.GRID.Rows = NumeroFila& + 1
'                            MUESTRADATOS9.GRID.TextMatrix(NumeroFila&, 1) = CODEXT$(CInt(CodigoConsignado&))
'                            MUESTRADATOS9.GRID.TextMatrix(NumeroFila&, 2) = DESCRIT0$(CInt(CodigoConsignado&))
'                            MUESTRADATOS9.GRID.TextMatrix(NumeroFila&, 3) = CStr(NumeroCompra&)
'                            MUESTRADATOS9.GRID.TextMatrix(NumeroFila&, 4) = RASOCLI$((-1) * Proveedor&)
'                            MUESTRADATOS9.GRID.TextMatrix(NumeroFila&, 5) = FORMATNUM$(SaldoPendienteCompra#, "F12.2")
'                            MUESTRADATOS9.GRID.TextMatrix(NumeroFila&, 6) = FORMATNUM$(SaldoConsignado#, "F12.2")
'                            MUESTRADATOS9.GRID.TextMatrix(NumeroFila&, 7) = FORMATNUM$(ARendir#, "F12.2")
'                            '
'                            RecepcionesEncontradas% = 1
'                            .MoveNext
'                            '
'                        Loop
'                    End With
'                    On Error Resume Next
'                    ComprasPendiente.Close
'                    Set ComprasPendiente = Nothing
'                    On Error GoTo 0
'                    .MoveNext
'                Loop
'            End With
'
'            On Error Resume Next
'            Consignaciones.Close
'            Set Consignaciones = Nothing
'            On Error GoTo 0
        End If
    Next II&
    '
    Call COLOREAR_GRILLASQL(MUESTRADATOS9.GRID, &H8000000F, 0)
    MUESTRADATOS9.GRID.Redraw = True
    '
    Screen.MousePointer = 1
    '
    If RecepcionesEncontradas% = 1 Then
        Respuesta% = COMALTER%("Se Han Encontrado Compras Pendientes de Materiales Consignados.\ \" _
                    & "¿Desea Ver el Detalle de las Compras?\\Ver Detalle\Continuar")
        If Respuesta% = 1 Then
            MUESTRADATOS9.GRID.Visible = True
            MUESTRADATOS9.GRID.ZOrder 0
            MUESTRADATOS9.Show 1
        End If
    End If
    
    ComprasPendientesDeLoConsignado% = RecepcionesEncontradas%
    
End Function

Private Sub FiltroRecepcionPorCompras(NumeroProveedor%)

20  Campos$ = "O-Compra/I10;Fecha/D8;Contacto/A32;Imagen/A5"
    Call CARGA_ENCABEZADO(MUESTRADATOS6.GRID, Campos$, MUESTRADATOS6, 1, 0, ANTOT, , 1)
        
    MUESTRADATOS6.Caption = "Ordenes de Compra Pendientes de Recepción"
   
    'PRESENTACION DE LOS COMPROBANTES QUE NO TIENEN CUIT
    '
    SQLX$ = " SELECT OC.NUME_OCOM,OC.FGEN_OCOM, OC.ATSR_OCOM FROM OCOMENCA OC WITH(NOLOCK)"
    SQLX$ = SQLX$ + "INNER JOIN OCOMDETA OD WITH(NOLOCK) ON OC.NUME_OCOM = OD.NUME_OCOM "
    SQLX$ = SQLX$ + " Where OD.Npro_Ocom = " & CStr(NumeroProveedor%) & " AND OD.Stat_Ocom < 8 "
    SQLX$ = SQLX$ + " AND  OD.Stat_Ocom <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
    SQLX$ = SQLX$ + " AND OD.Cant_Rec < OD.Cant_Ocom - 0.001 "
    SQLX$ = SQLX$ + " GROUP BY OC.NUME_OCOM, OC.FGEN_OCOM, OC.ATSR_OCOM "
    SQLX$ = SQLX$ + " ORDER BY OC.NUME_OCOM "
    Set rs1 = READSET(SQLX$)
    Fila& = 0
    MUESTRADATOS6.GRID.Redraw = False
    With rs1
      Do While Not .EOF
        NUMCO& = XVALO(!nume_ocom)
        If NUMCO& = NUMCO_ANT& Then GoTo 50
        Comprobante$ = SAFETEXT$(!nume_ocom)
        FECHA% = CVINT(!Fgen_Ocom)
        CONTACTO$ = SAFETEXT$(!AtSr_Ocom)
        '
        Fila& = Fila& + 1
        MUESTRADATOS6.GRID.Rows = Fila& + 1
        MUESTRADATOS6.GRID.TextMatrix(Fila&, 1) = Comprobante$
        MUESTRADATOS6.GRID.TextMatrix(Fila&, 2) = FECHATEX$(FECHA%)
        MUESTRADATOS6.GRID.TextMatrix(Fila&, 3) = CONTACTO$
        NUMCO_ANT& = NUMCO&
50      .MoveNext
      Loop
    End With
    rs1.Close
    Set rs1 = Nothing
    Call COLOREAR_GRILLASQL(MUESTRADATOS6.GRID, &H8000000F, 0)
    MUESTRADATOS6.GRID.Redraw = True
    
    If Fila& > 0 Then
       Screen.MousePointer = 1
       MUESTRADATOS6.GRID.Visible = True
       MUESTRADATOS6.GRID.ZOrder 0
       MUESTRADATOS6.Show 1
    Else
       Call COMUNI("No Se Han Encontrado Ordenes de Compra Pendientes")
       Screen.MousePointer = 1
    End If

    
    
    
    
    
'    SQLX$ = " SELECT OCOMENCA.NUME_OCOM FROM OCOMENCA WITH(NOLOCK)"
'    SQLX$ = SQLX$ + "INNER JOIN OCOMDETA WITH(NOLOCK) ON OCOMENCA.NUME_OCOM = OCOMDETA.NUME_OCOM "
'    SQLX$ = SQLX$ + " Where OCOMDETA.Stat_Ocom < 8 "
'    SQLX$ = SQLX$ + " AND  OCOMDETA.Stat_Ocom <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
'    SQLX$ = SQLX$ + " AND OCOMDETA.Cant_Rec < OCOMDETA.Cant_Ocom - 0.001 "
'    SQLX$ = SQLX$ + " AND OCOMDETA.Npro_Ocom = " & NumeroProveedor% & " "
'    SQLX$ = SQLX$ + " AND OCOMDETA.Npro_Ocom = " & NumeroProveedor% & " "
'    SQLX$ = SQLX$ + " GROUP BY OCOMENCA.NUME_OCOM "
'    SQLX$ = SQLX$ + " ORDER BY OCOMENCA.NUME_OCOM "
'    Dim RS1 As ADODB.Recordset
'    Set RS1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
'    FILA& = 0
'    ColorGris = &H8000000F
'    ColorBlanco = vbWhite
'    UltimoColor = vbWhite
'    With RS1
'      Do While Not .EOF
'        COMPROBANTE$ = SAFETEXT$(!Nume_Ocom)
'            CONDI$ = "NUME_OCOM=" & COMPROBANTE$
'        FECHA% = CVINT(VALOBADA("OCOMENCA", "FGEN_OCOM", CONDI$))
'        CONTACTO$ = SAFETEXT$(VALOBADA("OCOMENCA", "ATSR_OCOM", CONDI$))
'        '
'        FILA& = FILA& + 1
'        MUESTRADATOS6.GRID.Rows = FILA& + 1
'        MUESTRADATOS6.GRID.TextMatrix(FILA&, 1) = COMPROBANTE$
'        MUESTRADATOS6.GRID.TextMatrix(FILA&, 2) = FECHATEX$(FECHA%)
'        MUESTRADATOS6.GRID.TextMatrix(FILA&, 3) = CONTACTO$
'        If UltimoColor = ColorGris Then
'            UltimoColor = vbWhite
'        ElseIf FILA& > 1 Then
'            UltimoColor = ColorGris
'        End If
'        Call COLOREAR_GRILLA_SOLAFILA(MUESTRADATOS6.GRID, UltimoColor, FILA&, 0, 1, 0)
'        .MoveNext
'      Loop
'    End With
'    RS1.Close
'    Set RS1 = Nothing
'    If FILA& > 0 Then
'       Screen.MousePointer = 1
'       MUESTRADATOS6.GRID.Visible = True
'       MUESTRADATOS6.GRID.ZOrder 0
'       MUESTRADATOS6.Show 1
'    Else
'       Call COMUNI("No Se Han Encontrado Ordenes de Compra Pendientes")
'       Screen.MousePointer = 1
'    End If

End Sub

Sub GENERABOLREPEN(NROBOL&)
    '
    SQLX$ = " SELECT * FROM BOLRECEP WITH(NOLOCK) "
    SQLX$ = SQLX$ & " WHERE Status <= 1 "
    SQLX$ = SQLX$ & " AND NUMEBORE= " & NROBOL&
    SQLX$ = SQLX$ & " ORDER BY NumeBoRe "
    Dim COBRECEPTMP As ADODB.Recordset
    Set COBRECEPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With COBRECEPTMP
      JJ& = 0
      REBLA$ = REGBLAN$("BOLREPEN")
      '
      Do While (.EOF = False)
         '
         NUORSE& = XVALO(!NUMEORDSER)
         If NUORSE& > 0 Then GoTo 80
         '
         If XVALO(!CanToApro) > 0 Then GoTo 85
         If XVALO(!CanToRech) > 0 Then GoTo 85
         If XVALO(!CanToFalta) > 0 Then GoTo 85
         '
80       Y$ = REBLA$
         Call REPLA(Y$, MKS$(XVALO(!NUMEBORE)), 1, 4)
         Call REPLA(Y$, SAFETEXT$(!idenlote), 69, 12)
         Call REPLA(Y$, MKI$(CVINT(!FechRecep)), 81, 2)
            
         Call REPLA(Y$, MKI$(XVALO(!cod_inte)), 83, 2)
         Call REPLA(Y$, SAFETEXT$(!Unidad), 85, 4)
            
         Call REPLA(Y$, MKD$(XVALO(!CantNom)), 89, 8)
            
         Call REPLA(Y$, Chr$(XVALO(!DEPOENT)), 97, 1)
            
         Call REPLA(Y$, MKI$(XVALO(!NumProv)), 105, 2)
         Call REPLA(Y$, SAFETEXT$(!LOTECOL_PROV), 107, 12)
         Call REPLA(Y$, SAFETEXT$(!NumRem_Prov), 119, 16)
         Call REPLA(Y$, SAFETEXT$(!NumFac_Prov), 135, 16)
            
         Call REPLA(Y$, MKS$(XVALO(!NUMOC)), 151, 4)
         Call REPLA(Y$, SAFETEXT$(!MoneVaUni), 155, 4)
            
         Call REPLA(Y$, MKS$(XVALO(!Valo_Unit)), 159, 4)
            
         Call REPLA(Y$, MKD$(XVALO(!CanToApro)), 163, 8)
            
         Call REPLA(Y$, MKD$(XVALO(!CanToRech)), 171, 8)
         
         Call REPLA(Y$, MKD$(XVALO(!CanToFalta)), 179, 8)
         Call REPLA(Y$, MKI$(CVINT(!FechLibe)), 187, 2)
         
           NAENT& = 0
           On Error Resume Next
           NAENT& = !Nume_Aent
           On Error GoTo 0
         Call REPLA(Y$, MKS$(XVALO(NAENT&)), 189, 4) '151, 4)'\\\OT-06-0241-WAR-14/06/06///
         Call REPLA(Y$, SAFETEXT$(!Observa), 193, 64)
         JJ& = JJ& + 1
         Call GRAREG("BOLREPEN", Y$, JJ&)
         '
85    .MoveNext
      Loop
    End With
    '
90  COBRECEPTMP.Close
    Set COBRECEPTMP = Nothing
    '
95  Call SETULTREG("BOLREPEN", JJ&)
    Call CIERRARCH("BOLREPEN")
    '
    Screen.MousePointer = 1
    '
End Sub

Sub IMPREETIQUETA(NROBOL&)
    '
    FORIPRE$ = TRIM$(Control("", "STIKAPRO"))
    If FORIPRE$ <> "" Then
        PRINTERPORT$ = ""
        PFAC$ = Control$(IDENTER$, "PORTSTIK") ' IMPRESORA CONFIGURADA
        If PFAC$ = "" Then PFAC$ = Control$("*", "PORTSTIK")
        Call SETPRINPORT(PFAC$)
        NOCONFIRM% = 1 ' NO PRESENTA MENSAJE DE REIMPRESION
        CANCELPRINT% = 0
        '
        FIN& = ULTREG&("!BOLREDET")
        For JJ& = 1 To FIN&
            XX$ = REGLEIDO$("!BOLREDET", JJ&)
            
            ' DETERMINA SI SE TRATA DE UNA M.PRIMA A PARTIR DE TENER O NO UBICACION.
            ' CASO CONTRARIO OBVIA LA IMPRESION DE LA ETIQUETA
            UB$ = TRIM$(Mid$(XX$, 193, 16))
            If UB$ = "" Then GoTo 50
            
            CIXI% = CVI(Mid$(XX$, 83, 2))
            CAN# = CVD(Mid$(XX$, 89, 8))
            LINDIV# = ROUND(LINDIVIS(CIXI%), 0)
            LOTE$ = SAFETEXT$(VALOBADA("BOLRECEP", "IDENLOTE", "NUMEBORE=" & NROBOL& & " AND COD_INTE=" & CIXI%))
            FF% = CVINT(VALOBADA("BOLRECEP", "FechRecep", "NUMEBORE=" & NROBOL& & " AND COD_INTE=" & CIXI%))
            CANTI_ETIQ_A_GENERAR% = CANBULTOS(CIXI%, CAN#, "COMPRA")
            '
            TIPODOC$ = "Etiqueta Lote Recepcion " & LOTE$
            NUMEDOC$ = LOTE$
            '
                COXX$ = CODEXT$(CIXI%)
            CODPARAM$(1) = "COD-MAT": DOCPARAM$(1) = COXX$
            CODPARAM$(2) = "DES-MPRI": DOCPARAM$(2) = DESCRIT0$(CIXI%)
            CODPARAM$(3) = "CAN-MPRI"
            CODPARAM$(4) = "UNI-MPRI": DOCPARAM$(4) = Unimed(CIXI%)
            CODPARAM$(5) = "FECH-LAN": DOCPARAM$(5) = FECHATEX$(FF%)
            CODPARAM$(6) = "DESTINO": DOCPARAM$(6) = CODIBARRA_EAN13$(COLOR$(CIOF0%))
            CODPARAM$(7) = "REF-MAT2": DOCPARAM$(7) = LOTE$
            CODPARAM$(8) = "REF-CONJ":
            '
            CAPARDOC% = 8
            CAITAB1% = 0
            For i% = 1 To CANTI_ETIQ_A_GENERAR%
                If CAN# > LINDIV# Then
                    DOCPARAM$(3) = LINDIV#
                    CAN# = CAN# - LINDIV#
                    TX$ = TRIM$(COXX$) & Chr$(124) & TRIM$(LOTE$) & Chr$(124) & FORMATNUM$(CAN#, "F10.1")
                Else
                    DOCPARAM$(3) = CAN#
                    TX$ = TRIM$(COXX$) & Chr$(124) & TRIM$(LOTE$) & Chr$(124) & TRIM$(FORMATNUM$(CAN#, "F10.1"))
                End If
                DOCPARAM$(8) = CODIBARRA$(TX$, 2)
                Call PRINTDOC("@" + FORIPRE$)
            Next i%
50      Next JJ&
        NOCONFIRM% = 0
        Call SETPRINPORT("RESTORE")
    End If
    '
End Sub


Sub ImprimeEtiqRecepción(NROBOL&)
    '
    '02/08/18
    PRINTETI.mnuMenu3.Caption = "Revertir Selección"
    Call CREATABLA_ARCHIVOS_QR
    PRINTETI.mnuMenu3.Visible = True
    PRINTETI.Text5 = SAFETEXT$(NROBOL&)
    PRINTETI.Show 1
    Exit Sub
    
    '============================================================================================================
    '============================================================================================================
    '============================================================================================================
    
    FORIPRE$ = TRIM$(Control("BOLRECEP", "ETIQRECE"))
'    FORIPRE$ = TRIM$(CONTROL("", "STIKAPRO"))
    If FORIPRE$ = "" Then
       Call COMUNI("Falta Configurar Formulario de Impresión de Etiquetas")
       Exit Sub
    
    ElseIf FORIPRE$ <> "" Then
        PRINTERPORT$ = ""
        PFAC$ = Control$(IDENTER$, "PORTSTIK") ' IMPRESORA CONFIGURADA
        If PFAC$ = "" Then PFAC$ = Control$("*", "PORTSTIK")
        Call SETPRINPORT(PFAC$)
        NOCONFIRM% = 1 ' NO PRESENTA MENSAJE DE REIMPRESION
        CANCELPRINT% = 0
        '
        FIN& = ULTREG&("!BOLREDET")
        Call LIMPIATETAB

        For JJ& = 1 To FIN&
            XX$ = REGLEIDO$("!BOLREDET", JJ&)
            CIXI% = CVI(Mid$(XX$, 83, 2))
            CAN# = CVD(Mid$(XX$, 89, 8))
            LOTE$ = SAFETEXT$(VALOBADA("BOLRECEP", "IDENLOTE", "NUMEBORE=" & NROBOL& & " AND COD_INTE=" & CIXI%))
            FF% = CVINT(VALOBADA("BOLRECEP", "FechRecep", "NUMEBORE=" & NROBOL& & " AND COD_INTE=" & CIXI%))
            CANTI_ETIQ_A_GENERAR% = CANBULTOSXKG(CIXI%, CAN#)
            If CANTI_ETIQ_A_GENERAR% < 1 Then CANTI_ETIQ_A_GENERAR% = 1
            '
            '
            COXX$ = CODEXT$(CIXI%)
            OPX% = COMALTER%("Se van a Generar " & CANTI_ETIQ_A_GENERAR% & " Etiquetas Para el Artículo: " & DOCPARAM$(2) & "\\Cancelar\Imprimir Etiquetas")
            If OPX% = 2 Then
                PESOX = PESUNI(CI%)
                For i% = 1 To CANTI_ETIQ_A_GENERAR%
                    '
                    CODI$ = COXX$
                    TXORDEN$ = i% & " Etiquetas de " & CANTI_ETIQ_A_GENERAR%
                    TXTX$ = CODI$ & "|" & LOTE$ & "|" & PESOX
                    FILTX$ = LOTE$ & ";" & TRIM$(Str$(ORDEN%) & ";" & TXTX$ & ";" & TXORDEN$)
                    Call STDFORM(FORIPRE$, FILTX$, "PRINT")

                Next i%
            End If
50      Next JJ&
    End If
    '
End Sub


Sub PRESENTAVENTANAVALORIZADAXXXXXX(NPROV)
VENTANA.Inicializar = 0
VENTANA.Formula(5) = "3;*;4"
VENTANA.CampEditables = "1;3;4;5"
VENTANA.ColEcoCampo(1) = "CODIGO;DESCRIPCION;MASTER"
VENTANA.Campox = "CODIGO/A12;DESCRIPCION/A32;CANTIDAD/f12.2;PRECIO/F12.2;IMPORTE/F12.2"

VENTANA.AgregaRegistro = 1
VENTANA.UtilizaFormula = 1
VENTANA.HabilitaBorrar = 1
VENTANA.HabilitaInsertar = 1
VENTANA.NoMostrarBotonTilde = 0
VENTANA.TITULO = ""
VENTANA.Inicializar = 1
        VENTABNEW.Width = 9000
        'CARGAR ENCABEZADO
'        Call CARGA_ENCABEZADO(VENTABNEW.MSF, Campos$, VENTABNEW)
'        Call CARGA_ENCABEZADO(VENTABNEW.MSFFORMULA, Campos$, VENTABNEW)
        'GENERAR FORMATOS
'        FORMATOS$ = GENERA_FORMATOS$(CAMPOS$) ' DEVUELVE ESTO = " /A16; /A20; /A3; /f10.1; /D8; /F10.2; /F12.2; /F10.2; /F7; /F10.2; /F0, /F0; /D8"
'        FORMATOSCONBARRA$ = GENERA_FORMATOS_CON_BARRA$(CAMPOS$)
'        'CAMPOS$ = " /A16; /A20; /A3; /f10.1; /D8; /F10.2; /F12.2; /F10.2; /F7; /F10.2; /F0, /F0; /D8"
'        Call CARGA_ENCABEZADO(VENTABNEW.MSFTOTALES, FORMATOSCONBARRA$, VENTABNEW)
'        '
'        VENTABNEW.MSFFORMULA.Rows = 4
        
        'CARGAR LAS FORMULAS EN LAS FILA 1 SEGUN CELDA QUE CORRESPONDA
'        VENTABNEW.MSFFORMULA.TextMatrix(1, 7) = "4;*;6" 'DEBEN IR EN FILA 1
'        VENTABNEW.MSFFORMULA.TextMatrix(1, 9) = "=PORCENT_ENTRE_2_VALORES;10;6" 'DEBE LLEVAR SIMBOLO =
        '
        'ACA SE LE PASA LOS FORMATOS DE CELDA PARA QUE LA GRILLA INFERIOR CON LOS TOTALES QUEDE IGUAL QUE LA DE ARRIBA.
'        Call AGREGAFILA(VENTABNEW.MSFFORMULA, FORMATOS$, 2)  'DEBEN IR EN LA FILA 2
        '
        'EDITABLES (INDICAR CUALES SON LAS COLUMNAS EDITABLES)
'        VENTABNEW.EDITABLES.Caption = "1;4;10"
        
        'INDICAR SI CAMPO EDITABLE LLEVA LISTA DE SELECCION Y ARCHIVO ECO
'        VENTABNEW.MSFFORMULA.TextMatrix(3, 1) = "CODIGO;Descripcion;Master"
        '
        'INDICAR SI SI VA A GREGAR FILAS NUEVAS COMO VENTABU
'        VENTABNEW.AGREGAFILASNUEVAS.Caption = "1"
        '
        
'        VENTABNEW.MSFTOTALES.Rows = 2
'        VENTABNEW.MSFTOTALES.TextMatrix(1, 7) = "TOTAL"
        

        '
        J& = 0
        FIN& = ULTREG&("!BOLREDET")
        For YX& = 1 To FIN&
           YY$ = REGLEIDO$("!BOLREDET", YX&)
           CI1% = CVI(Mid$(YY$, 83, 2))
           CAN# = CVD(Mid$(YY$, 89, 8))
           If CI1% > 0 Then
               If CAN# > 0.005 Then
                 '
                 UNID$ = TRIM$(Mid$(YY$, 85, 4))
                 If UNID$ = "" Then UNID$ = Unimed$(CI1%)
                 FEX = CVI(Mid$(YY$, 155, 2))
                 FESOLEN$ = FECHATEX$(FEX)
                 NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
                 '
                 If NULOTE$ = "" Then
                    NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
                 End If
                 NUOC& = CVS(Mid$(YY$, 151, 4))
                 If NUOC& <> NUOC_ANT& Then 'SI CAMBIA LA ORDEN DE COMPRA VUELVE A BUSCAR EL COEFICIENTE
                   coefi = COEFI_COMPRA_SIEMPRE_DOLAR(NUOC&)
                 End If
                 If coefi < 1 Then coefi = 1
                 
                 NORD% = Asc(Mid$(YY$, 98, 1))
                 ESPRIN$ = Mid$(YY$, 5, 64)
                 NUOC& = CVS(Mid$(YY$, 151, 4))
                 FENTRESO% = CVI(Mid$(YY$, 155, 2))
                 
                 With VENTABNEW.MSF
                    J& = J& + 1
                    VENTABNEW.MSF.Rows = J& + 1
                    .TextMatrix(J&, 1) = CODEXT$(CI1%)
                    .TextMatrix(J&, 2) = DESCRIT0$(CI1%)
                    .TextMatrix(J&, 3) = UNID$
                    .TextMatrix(J&, 4) = FORMATNUM$(CAN#, "F10.1")
                    .TextMatrix(J&, 5) = FESOLEN$
                        COSTOFOB1# = COSTO_SEGUN_COMPRA(CI1%, NUOC&, NORD%)
                        COSTOFOB1# = COSTOFOB1# / coefi
                    .TextMatrix(J&, 6) = TRIM$(FORMATNUM$(COSTOFOB1#, "F12.2"))
                    .TextMatrix(J&, 7) = TRIM$(FORMATNUM$(COSTOFOB1# * CAN#, "F12.2"))
                        VIAENTR% = VIAENTRE_COMPRA(NUOC&)
                        PVTA# = PRECIOVENTA_X_FORMULA(COSTOFOB1#, VIAENTR%, CI1%, 0, NPROV, (-1))
                        PVTAEDITADO# = PRECIODEVENTA_AHP(CI1%, NPROV, VIAENTR%)
                        'If PVTAEDITADO# < 0.005 Then PVTAEDITADO# = PVTA# 'PRESENTO VACIO SI NO TIENE IMPORTE.
                    .TextMatrix(J&, 8) = TRIM$(FORMATNUM$(PVTA#, "F12.2"))
                    .TextMatrix(J&, 9) = TRIM$(FORMATNUM$(PORCENT_ENTRE_2_VALORES(PVTAEDITADO#, COSTOFOB1#), "F9.2"))
                    .TextMatrix(J&, 10) = TRIM$(FORMATNUM$(PVTAEDITADO#, "F9.2"))
                    .TextMatrix(J&, 11) = SAFETEXT$(NORD%)
                    .TextMatrix(J&, 12) = SAFETEXT$(NUOC&)
                    .TextMatrix(J&, 13) = FECHATEX$(FENTRESO%)
                    NUOC_ANT& = NUOC&
                    
                 End With
               End If
           End If
         Next YX&
         If VENTABNEW.MSF.Rows > 1 Then
            VENTABNEW.Show 1
            If VENTABNEW.APROBADO% > 0 Then
             JJ& = 0: CONDI$ = ""
             Dim viaentregax()
             ReDim Preserve viaentregax(1)
             With VENTABNEW
               For i& = 1 To .MSF.Rows - 1
                 Y$ = REGBLAN$("!BOLREDET")
                   CI1% = XVALO(CODINT(.MSF.TextMatrix(i&, 1)))
                   prexiventa# = XVALO(.MSF.TextMatrix(i&, 10))
                   If prexiventa# > 0 Then ' cargo en un  VECTOR PARA LEER SOLO UNA VEZ  Y HACER MAS RAPIDA LA ACTUALIZACION
                     If NUOC& > UBound(viaentregax) Then ReDim Preserve viaentregax(NUOC&)
                     If XVALO(viaentregax(NUOC&)) < 1 Then viaentregax(NUOC&) = VIAENTRE_COMPRA(NUOC&)
                     Call RUTINA_ACTU_COSTOPROVE(NPROV, CodiEmp%, CI1%, XVALO(viaentregax(NUOC&)), prexiventa#)
                   End If
                 Call REPLA(Y$, MKI$(CI1%), 83, 2)
                 Call REPLA(Y$, .MSF.TextMatrix(i&, 3), 85, 4)
                   CAPEN = XVALO(.MSF.TextMatrix(i&, 4))
                 Call REPLA(Y$, MKD$(CDbl(CAPEN)), 89, 8)
'                If ANUCA% = 1 Then Call REPLA(Y$, MKD$(0), 89, 8)
                 Call REPLA(Y$, MKD$(CDbl(CAPEN)), 163, 8)
                 '
                 Call REPLA(Y$, Chr$(XVALO(.MSF.TextMatrix(i&, 11))), 98, 1)
                 Call REPLA(Y$, MKS$(0), 100, 4) 'no tiene AE
                 '
                 NUOC& = XVALO(.MSF.TextMatrix(i&, 12))
                 Call REPLA(Y$, MKS$(NUOC&), 151, 4)
                   FENTRESO% = FECHANUM(.MSF.TextMatrix(i&, 11))
                 Call REPLA(Y$, MKI$(FENTRESO%), 155, 2)
                 Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
                 '
                 JJ& = JJ& + 1
                 Call GRAREG("!BOLREDET", Y$, JJ&)
               Next i&
             End With
            End If
         End If
         Call SETULTREG("!BOLREDET", JJ&)
         '
'         'RECORRO SOLO LOS CODIGOS QUE SE LES AGREGO PRECIO DE VENTA EN AL GRILLA Y LE GRABO EL PRECIO DEL VECTOR QUE CARGUE CUANDO LEI LA GRILLA.
'         CONDI$ = Left$(CONDI$, Len(CONDI$) - 2) ' LE QUITO LA COMA FINAL
'         CONDI$ = "(" & CONDI$ & ")"
'         SQLX$ = "SELECT PRECIOVENTA,CODINT FROM COSTOXPROVE "
'         SQLX$ = SQLX$ + " WHERE " & CONDI$
'         SQLX$ = SQLX$ + " and Nume_Cli = " & NPROV
'         'ver si esta bien el campo de proveedor y falta la via de entregea
'         SQLX$ = SQLX$ + " ORDER BY CODINT"
'         Dim RS As ADODB.Recordset
'         Set RS = New ADODB.Recordset
'         RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'         With RS
'            Do While Not .EOF
'                CI1% = XVALO(!CODINT)
'                !PRECIOVENTA = XVALO(prexiventa#(CI1%))
'                .Update
'              .MoveNext
'            Loop
'         End With
         Erase viaentregax
         
'     Set RS = READSET(SQLX$)
'       Y$ = REGBLAN$("!BOLREDET")
'             Call REPLA(Y$, MKI$(CI%), 83, 2)
'             Call REPLA(Y$, UNIMED$(CI%), 85, 4)
'             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 89, 8)
'                If ANUCA% = 1 Then Call REPLA(Y$, MKD$(0), 89, 8)
'             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 163, 8)
'             '
'             Call REPLA(Y$, Chr$(XVALO(!NuOrd_Item)), 98, 1)
'             Call REPLA(Y$, MKS$(0), 100, 4) 'no tiene AE
'             '
'             Call REPLA(Y$, LOTECO$, 107, 12)
'             Call REPLA(Y$, MKS$(NUOC&), 151, 4)
'             Call REPLA(Y$, MKI$(FENTRESO%), 155, 2)
'             Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
'             Call REPLA(Y$, MKS$(!N_ORSERV), 189, 4)
    

End Sub

Sub PRESENTAVENTANAVALORIZADA(NPROV)
        Campos$ = "Código/A16;Descripción/A20;U-M/A3;Cantidad/f10.1;F.Solic./D8;Costo/F10.2;Importe/F12.2;Formula/F10.2;%/F7.2;Venta/F10.2;NUMORDEN/F0;O.C./F7;Entregar/D8"
        Call CARGA_ENCABEZADO(BRECPAHP.MSF, Campos$, BRECPAHP)
        Call CARGA_ENCABEZADO(BRECPAHP.MSFFORMULA, Campos$, BRECPAHP)
        FORMATOS$ = GENERA_FORMATOS$(Campos$) ' DEVUELVE ESTO = " /A16; /A20; /A3; /f10.1; /D8; /F10.2; /F12.2; /F10.2; /F7; /F10.2; /F0, /F0; /D8"
        FORMATOSCONBARRA$ = GENERA_FORMATOS_CON_BARRA$(Campos$)
        'CAMPOS$ = " /A16; /A20; /A3; /f10.1; /D8; /F10.2; /F12.2; /F10.2; /F7; /F10.2; /F0, /F0; /D8"
        Call CARGA_ENCABEZADO(BRECPAHP.MSFTOTALES, FORMATOSCONBARRA$, BRECPAHP)
        
        BRECPAHP.MSFFORMULA.Rows = 2  'DEBEN IR EN FILA 1
        BRECPAHP.MSFFORMULA.TextMatrix(1, 7) = "4;*;6"
        BRECPAHP.MSFFORMULA.TextMatrix(1, 9) = "=PORCENT_ENTRE_2_VALORES;10;6" 'DEBE LLEVAR SIMBOLO =
        Call AGREGAFILA(BRECPAHP.MSFFORMULA, FORMATOS$, 2) 'DEBEN IR EN LA FILA 2
        '
        'EDITABLES
        BRECPAHP.EDITABLES.Caption = "4;10"
        BRECPAHP.MSFTOTALES.Rows = 2
        BRECPAHP.MSFTOTALES.TextMatrix(1, 7) = "TOTAL"
        

        '
        J& = 0
        FIN& = ULTREG&("!BOLREDET")
        For YX& = 1 To FIN&
           YY$ = REGLEIDO$("!BOLREDET", YX&)
           CI1% = CVI(Mid$(YY$, 83, 2))
           CAN# = CVD(Mid$(YY$, 89, 8))
           If CI1% > 0 Then
               If CAN# > 0.005 Then
                 '
                 UNID$ = TRIM$(Mid$(YY$, 85, 4))
                 If UNID$ = "" Then UNID$ = Unimed$(CI1%)
                 FEX = CVI(Mid$(YY$, 155, 2))
                 FESOLEN$ = FECHATEX$(FEX)
                 NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
                 '
                 If NULOTE$ = "" Then
                    NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
                 End If
                 NUOC& = CVS(Mid$(YY$, 151, 4))
                 If NUOC& <> NUOC_ANT& Then 'SI CAMBIA LA ORDEN DE COMPRA VUELVE A BUSCAR EL COEFICIENTE
                   coefi = COEFI_COMPRA_SIEMPRE_DOLAR(NUOC&)
                 End If
                 If coefi < 1 Then coefi = 1
                 
                 NORD% = Asc(Mid$(YY$, 98, 1))
                 ESPRIN$ = Mid$(YY$, 5, 64)
                 NUOC& = CVS(Mid$(YY$, 151, 4))
                 FENTRESO% = CVI(Mid$(YY$, 155, 2))
                 
                 With BRECPAHP.MSF
                    J& = J& + 1
                    BRECPAHP.MSF.Rows = J& + 1
                    .TextMatrix(J&, 1) = CODEXT$(CI1%)
                    .TextMatrix(J&, 2) = DESCRIT0$(CI1%)
                    .TextMatrix(J&, 3) = UNID$
                    .TextMatrix(J&, 4) = FORMATNUM$(CAN#, "F10.1")
                    .TextMatrix(J&, 5) = FESOLEN$
                        COSTOFOB1# = COSTO_SEGUN_COMPRA(CI1%, NUOC&, NORD%)
                        COSTOFOB1# = COSTOFOB1# / coefi
                    .TextMatrix(J&, 6) = TRIM$(FORMATNUM$(COSTOFOB1#, "F12.2"))
                    .TextMatrix(J&, 7) = TRIM$(FORMATNUM$(COSTOFOB1# * CAN#, "F12.2"))
                        VIAENTR% = VIAENTRE_COMPRA(NUOC&)
                        PVTA# = PRECIOVENTA_X_FORMULA(COSTOFOB1#, VIAENTR%, CI1%, 0, NPROV, (-1))
                        PVTAEDITADO# = PRECIODEVENTA_AHP(CI1%, NPROV, VIAENTR%)
                        'If PVTAEDITADO# < 0.005 Then PVTAEDITADO# = PVTA# 'PRESENTO VACIO SI NO TIENE IMPORTE.
                    .TextMatrix(J&, 8) = TRIM$(FORMATNUM$(PVTA#, "F12.2"))
                    .TextMatrix(J&, 9) = TRIM$(FORMATNUM$(PORCENT_ENTRE_2_VALORES(PVTAEDITADO#, COSTOFOB1#), "F9.2"))
                    .TextMatrix(J&, 10) = TRIM$(FORMATNUM$(PVTAEDITADO#, "F9.2"))
                    .TextMatrix(J&, 11) = SAFETEXT$(NORD%)
                    .TextMatrix(J&, 12) = SAFETEXT$(NUOC&)
                    .TextMatrix(J&, 13) = FECHATEX$(FENTRESO%)
                    NUOC_ANT& = NUOC&
                    
                 End With
               End If
           End If
         Next YX&
         If BRECPAHP.MSF.Rows > 1 Then
            BRECPAHP.Show 1
            If BRECPAHP.APROBADO% > 0 Then
             JJ& = 0: CONDI$ = ""
             Dim viaentregax()
             ReDim Preserve viaentregax(1)
             With BRECPAHP
               For i& = 1 To .MSF.Rows - 1
                 Y$ = REGBLAN$("!BOLREDET")
                   CI1% = XVALO(CODINT(.MSF.TextMatrix(i&, 1)))
                   prexiventa# = XVALO(.MSF.TextMatrix(i&, 10))
                   If prexiventa# > 0 Then ' cargo en un  VECTOR PARA LEER SOLO UNA VEZ  Y HACER MAS RAPIDA LA ACTUALIZACION
                     If NUOC& > UBound(viaentregax) Then ReDim Preserve viaentregax(NUOC&)
                     If XVALO(viaentregax(NUOC&)) < 1 Then viaentregax(NUOC&) = VIAENTRE_COMPRA(NUOC&)
                     Call RUTINA_ACTU_COSTOPROVE(NPROV, CodiEmp%, CI1%, XVALO(viaentregax(NUOC&)), prexiventa#)
                   End If
                 Call REPLA(Y$, MKI$(CI1%), 83, 2)
                 Call REPLA(Y$, .MSF.TextMatrix(i&, 3), 85, 4)
                   CAPEN = XVALO(.MSF.TextMatrix(i&, 4))
                 Call REPLA(Y$, MKD$(CDbl(CAPEN)), 89, 8)
'                If ANUCA% = 1 Then Call REPLA(Y$, MKD$(0), 89, 8)
                 Call REPLA(Y$, MKD$(CDbl(CAPEN)), 163, 8)
                 '
                 Call REPLA(Y$, Chr$(XVALO(.MSF.TextMatrix(i&, 11))), 98, 1)
                 Call REPLA(Y$, MKS$(0), 100, 4) 'no tiene AE
                 '
                 NUOC& = XVALO(.MSF.TextMatrix(i&, 12))
                 Call REPLA(Y$, MKS$(NUOC&), 151, 4)
                   FENTRESO% = FECHANUM(.MSF.TextMatrix(i&, 11))
                 Call REPLA(Y$, MKI$(FENTRESO%), 155, 2)
                 Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
                 '
                 JJ& = JJ& + 1
                 Call GRAREG("!BOLREDET", Y$, JJ&)
               Next i&
             End With
            End If
         End If
         Call SETULTREG("!BOLREDET", JJ&)
         '
'         'RECORRO SOLO LOS CODIGOS QUE SE LES AGREGO PRECIO DE VENTA EN AL GRILLA Y LE GRABO EL PRECIO DEL VECTOR QUE CARGUE CUANDO LEI LA GRILLA.
'         CONDI$ = Left$(CONDI$, Len(CONDI$) - 2) ' LE QUITO LA COMA FINAL
'         CONDI$ = "(" & CONDI$ & ")"
'         SQLX$ = "SELECT PRECIOVENTA,CODINT FROM COSTOXPROVE "
'         SQLX$ = SQLX$ + " WHERE " & CONDI$
'         SQLX$ = SQLX$ + " and Nume_Cli = " & NPROV
'         'ver si esta bien el campo de proveedor y falta la via de entregea
'         SQLX$ = SQLX$ + " ORDER BY CODINT"
'         Dim RS As ADODB.Recordset
'         Set RS = New ADODB.Recordset
'         RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'         With RS
'            Do While Not .EOF
'                CI1% = XVALO(!CODINT)
'                !PRECIOVENTA = XVALO(prexiventa#(CI1%))
'                .Update
'              .MoveNext
'            Loop
'         End With
         Erase viaentregax
         
'     Set RS = READSET(SQLX$)
'       Y$ = REGBLAN$("!BOLREDET")
'             Call REPLA(Y$, MKI$(CI%), 83, 2)
'             Call REPLA(Y$, UNIMED$(CI%), 85, 4)
'             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 89, 8)
'                If ANUCA% = 1 Then Call REPLA(Y$, MKD$(0), 89, 8)
'             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 163, 8)
'             '
'             Call REPLA(Y$, Chr$(XVALO(!NuOrd_Item)), 98, 1)
'             Call REPLA(Y$, MKS$(0), 100, 4) 'no tiene AE
'             '
'             Call REPLA(Y$, LOTECO$, 107, 12)
'             Call REPLA(Y$, MKS$(NUOC&), 151, 4)
'             Call REPLA(Y$, MKI$(FENTRESO%), 155, 2)
'             Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
'             Call REPLA(Y$, MKS$(!N_ORSERV), 189, 4)
    

End Sub


Function PRESENTAVENTANASCRAP(NPROV) As Boolean
    'RECEPCION DE BOBINAS EN ENVAPLAST DE CAMBACERES CON EL CODIGO QR QUE IMPRIMEN EN CAMBACERES CUANDO HACEN EL REMITO
    PRESENTAVENTANASCRAP = False
'    Call COPYSTRU("BOLREPAP", "BOLRETMP")
'    FORIPRE$ = "ETR2-CTP"
'    If EXISTE%(LUDAT$ + FORIPRE$ + ".SQR") = 0 Then Exit Function
    
'    IMAG$ = "IMAGRECE_QR.JPG"
'    IDENTIFICADOR$ = "QR_RECEPCION"
    ANCHOX% = 0
    ALTOX% = 0

    On Error Resume Next
    Unload VENTABNEW
    On Error GoTo 0
    
    
    VENTANA.Inicializar = 0
    VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 0
    VENTANA.HabilitaInsertar = 1
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "RECEPCION DE SCRAP"
    VENTANA.Campox = "Codigo/A16;Descripcion/A32;Kgs./F8"
    VENTANA.CampEditables = "1;3"
    VENTANA.ColEcoCampo(1) = "codigo;descripcion;master;status>0"

    VENTANA.Inicializar = 1
  
10  VENTABNEW.Show 1
    If VENTABNEW.APROBADO% < 1 Then PRESENTAVENTANASCRAP = False
    If VENTABNEW.APROBADO% > 0 Then
        'valido que todo sea scrap
        For J& = 1 To VENTABNEW.MSF.Rows - 1
            If TRIM$(VENTABNEW.MSF.TextMatrix(J&, 3)) <> "" Then ' VERIFICA QUE TENGA METROS
                CI1% = CODINT%(VENTABNEW.MSF.TextMatrix(J&, 1))
                familia1$ = TRIM$(UCase$(CODFAMIL$(CI1%)))
                If familia1$ <> "SC" Then
                   Call COMUNI("Se detecto  Material Ingresado que no Corresponde a Scrap (" & CODEXT$(CI1%) & " )")
                   PRESENTAVENTANASCRAP = False: Exit Function
                End If
            End If
        Next J&
        
        Call SETULTREG("!BOLREDET", 0)
        JJ& = ULTREG&("!BOLREDET")
        For i = 1 To VENTABNEW.MSF.Rows - 1
            If TRIM$(VENTABNEW.MSF.TextMatrix(i, 3)) <> "" Then ' VERIFICA QUE TENGA METROS
                CI% = CODINT%(VENTABNEW.MSF.TextMatrix(i, 1))
        
                ' *** GRABA DATOS DE LA BOBINA PROCESADA CORRECTAMENTE
                CAPEN = 0
                ORDOC% = 0
                NUOC& = 0
                FENTRESO% = HOY(HO$)
                AnchoBobina# = 0
                LargoBobina# = XVALO(VENTABNEW.MSF.TextMatrix(i, 3))
                Peso# = XVALO(VENTABNEW.MSF.TextMatrix(i, 3)) 'es la cantidad por que son kg
'                VENTABNEW.MSF.TextMatrix(i, 6) = UCase$(REPLACAR(VENTABNEW.MSF.TextMatrix(i, 6), "'", "-"))
'    VENTANA.Campox = "Codigo/A16;Descripcion/A32;Entrega/F8;Metros/F10.1;Peso/F10.1;REMITO/A18;Lote Proveedor"

                Y$ = REGBLAN("BOLREPAP")
                Call REPLA(Y$, MKI$(CI%), 83, 2)
                Call REPLA(Y$, Unimed$(CI%), 85, 4)
                Call REPLA(Y$, MKD$(CDbl(LargoBobina#)), 89, 8)
                   If ANUCA% = 1 Then Call REPLA(Y$, MKD$(0), 89, 8)
                Call REPLA(Y$, MKD$(CDbl(CAPEN)), 163, 8)
                '
                Call REPLA(Y$, Chr$(XVALO(ORDOC%)), 98, 1)
                Call REPLA(Y$, MKS$(0), 100, 4) 'no tiene AE
                '
'                LOTECO$ = UCase$(REPLACAR(VENTABNEW.MSF.TextMatrix(i, 7), "'", "-"))
                Call REPLA(Y$, "", 107, 12)
                Call REPLA(Y$, MKS$(NUOC&), 151, 4)
                Call REPLA(Y$, MKI$(FENTRESO%), 155, 2)
                Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
                Call REPLA(Y$, MKS$(0), 189, 4)

                Call REPLA(Y$, MKD$(AnchoBobina#), 193, 8)
                Call REPLA(Y$, MKD$(LargoBobina#), 201, 8)
                Call REPLA(Y$, MKS$(1), 209, 4)
                Call REPLA(Y$, MKS$(Peso#), 253, 4)
                
                JJ& = JJ& + 1
'                Call GRAREG("BOLRETMP", Y$, JJ&)
                Call GRAREG("!BOLREDET", Y$, JJ&)

                
                ' *** MARCA REGISTRO / BOBINA COMO PROCESADO OK
                
                ' *** IMPRIME ETIQUETA
'                PFAC$ = Control$(IDENTER$, "PORTSTIK")
'                If PFAC$ = "" Then PFAC$ = Control$("*", "PORTSTIK")
'                Call SETPRINPORT(PFAC$)
'
'                Codigo$ = VENTABNEW.MSF.TextMatrix(i, 1)
'                Descripcion$ = VENTABNEW.MSF.TextMatrix(i, 2)
'                LOTE$ = "BR-" & CStr(NumeroRecepcion&) & "-" & CStr(JJ&)
'                Largo# = LargoBobina#
'                Proveedor$ = rasocli$(NumeroProveedor% * -1)
'                LoteProveedor$ = LOTECO$ 'el lote de cambaceres lo lee mas arriba
'                If TRIM$(LoteProveedor$) = "" Then LoteProveedor$ = "-"
'                '
'                REMITOYLOTEPROVEEDOR$ = TRIM$(Left$(Remito$, 18)) & " " & LoteProveedor$
'                TEXTOQR$ = Codigo$ & "|" & LOTE$
'                IMAG$ = "IMAGRECE_QR.JPG"
'                IDENTIFICADOR$ = "QR_RECEPCION"
'                ANCHOX% = 0
'                ALTOX% = 0
'
'                On Error Resume Next
'                Kill LUCOM$ & IMAG$
'                On Error GoTo 0
'                'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
'                QRX$ = TRIM$(IMAGENCODBARRA$(LUCOM$, IMAG$, TEXTOQR$, "QR_CODE", ANCHOX%, ALTOX%))
'                Call T_Espera(3)
'                Image1.Picture = LoadPicture(LUCOM$ & IMAG$)
'                Call GRABAR_QR(LUCOM$ & IMAG$, IDENTIFICADOR$)
'
'                FILTX$ = LOTE$ & ";" & Codigo$ & ";" & Descripcion$ & ";" & Proveedor$ & ";" & TRIM$(FORMATNUM$(Largo#, "F12.1")) & ";" & TRIM$(FORMATNUM$(Peso#, "F12.1")) & ";" & REMITOYLOTEPROVEEDOR$
'                FILTX1$ = LOTE$ & ";" & Codigo$ & ";" & Descripcion$ & ";" & Proveedor$ & ";" & TRIM$(FORMATNUM$(Largo#, "F12.1")) & ";" & TRIM$(FORMATNUM$(Peso#, "F12.1")) & ";" & REMITOYLOTEPROVEEDOR$
'                Call STDFORM(FORIPRE$, FILTX$, "PRINT")
'                Call STDFORM(FORIPRE$, FILTX1$, "PRINT")
'
'                Call SETPRINPORT("RESTORE")
'                PRINTERPORT$ = ""
            End If
15      Next i
        
        If ULTREG&("!BOLREDET") > 0 Then
            Call SETULTREG("BOLREPAP", 0)
            'Call TRALOCAL("BOLRETMP", "BOLREPAP")
'            Call TRALOCAL("!BOLRETMP", "!BOLREDET")
'            Call SETULTREG("BOLRETMP", 0)
            
'            ' *** VERIFICA SI TODOS LOS ARTICULOS SON TRAZABLES. CASO CONTRARIO LOS CONFIGURA.
'            For IX& = 1 To ULTREG&("!BOLREPAP")
'                XX$ = REGLEIDO$("!BOLREPAP", IX&)
'                CI% = CVI(Mid$(XX$, 83, 2))
'                If ESTRAZABLE%(CI%) <> 1 Then
'                    Call ACTUREGISTRO("MASTER", "TRAZABLE", "CODINT=" & CStr(CI%), "1", REGAF&)
'                End If
'            Next IX&
            
            PRESENTAVENTANASCRAP = True
        End If
        
        On Error Resume Next
        Unload VENTABNEW
        On Error GoTo 0
    End If

End Function



Sub PrintEtiquetaRecepcionBobinaCambaceres(LoteRecepcion$)
    
    If PFAC$ = "" Then PFAC$ = Control$("*", "PORTSTIK")
    Call SETPRINPORT(PFAC$)
    
    ConsultaSql$ = "select fechrecep, cod_inte, cod_exte, descripitem, idenlote, numrem_prov, m_largo, cantnom, peso, raso_prov, NumRem_Prov, CantidadBultos from bolrecep with(nolock) where idenlote='" & LoteRecepcion$ & "'"
    Set Rst = READSET(ConsultaSql$)

  
70  With Rst
       If Not .EOF Then
            CodInte% = !cod_inte
            Codigo$ = CODEXT$(CodInte%)
            Descripcion$ = DESCRIT0$(CodInte%)
            LOTE$ = !idenlote
            Largo$ = "0": If XVALO(!M_Largo) > 0.0005 Then Largo$ = TRIM$(FORMATNUM$(!M_Largo, "F12.1"))
            Peso$ = "0": If XVALO(!Peso) > 0.005 Then Peso$ = TRIM$(FORMATNUM$(!Peso, "F12.1"))
            Proveedor$ = SAFETEXT$(!Raso_Prov)
            Remito$ = SAFETEXT$(!NumRem_Prov)
            cantidadNominal# = XVALO(!CantNom)
            FechaRecepcion% = CVINT(!FechRecep)
            .MoveNext
       Else
          Exit Sub
       End If
    End With
         
             familia1$ = TRIM$(UCase$(CODFAMIL$(CodInte%)))
             TEXTOQR$ = Codigo$ & "|" & LOTE$
             IMAG$ = "IMAGRECE_QR.JPG"
             IDENTIFICADOR$ = "QR_RECEPCION"
             ANCHOX% = 0
             ALTOX% = 0
             USUARIO% = 0
            '
            On Error Resume Next
            Kill LUCOM$ & IMAG$
            On Error GoTo 0
            'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
            QRX$ = TRIM$(IMAGENCODBARRA$(LUCOM$, IMAG$, TEXTOQR$, "QR_CODE", ANCHOX%, ALTOX%))
            Call T_Espera(3)
            Image1.Picture = LoadPicture(LUCOM$ & IMAG$)
            Call GRABAR_QR(LUCOM$ & IMAG$, IDENTIFICADOR$)
                'ACA CAMBACERES USA LA MISMA ETIQUETA DE RECEPCION DE CONTEPLAST
                FILTX$ = LOTE$ & ";" & Codigo$ & ";" & Descripcion$ & ";" & Proveedor$ & ";" & TRIM$(Largo$) & ";" & TRIM$(Peso$) & ";" & IDENTIFICADOR$ & ";" & CStr(USUARIO%) & ";" & Remito$
                Call STDFORM("ETRE-CMB", FILTX$, "PRINT")
                FILTX$ = LOTE$ & ";" & Codigo$ & ";" & Descripcion$ & ";" & Proveedor$ & ";" & TRIM$(Largo$) & ";" & TRIM$(Peso$) & ";" & IDENTIFICADOR$ & ";" & CStr(USUARIO%) & ";" & Remito$
                Call STDFORM("ETRE-CMB", FILTX$, "PRINT")
            Call SETPRINPORT("RESTORE")
            PRINTERPORT$ = ""
        

         
  
End Sub

Sub PROCESO_REC_VS_COMPRAS()
   'RECORRO EL BOLRECEP AGRUPADO POR CODIGO Y ORDEN DE COMPRA
   '
   OPX% = COMALTER%("Atención Este Proceso Recupera Recepciones de Ordenes de Compra No Vinculadas\\Cancelar\Continuar")
   If OPX% <> 2 Then Exit Sub
   '
   'ESTE PROCESO REVISA LAS RECEPCIONES REALIZADAS Y VERIFICA SI LA CANTIDAD RECEPCIONADA SEGUN ORDEN DE COMPRA
   'COINCIDE CON LA CANTIDAD REGISTRADA EN EL OCOMDETA COMO CANTIDAD RECIBIDA.
   
   'CRITERIO:
   '*********
   'RECORRE EL BOLRECEP
   'TRAE LA CANTIDAD AGRUPADA POR CODIGO Y ORDEN DE COMPRA
   'GENERA CONSULTA EN OCOMDETA FILTRANDO POR ORDEN DE COMPRA, CODIGO,STATUS < 3 Y CANTIOCOM > CANTIRECIBIDA
   'TENIENDO EN CUENTA EL CASO QUE HAYA UN ITEM MAS DE UNA VEZ EN LA ORDEN DE COMPRA
   'ORDENA POR FECHA SOLICITADA Y VA ASIGNANDO LA DIFERENCIA EN EL CASO QUE HAYA DIFERENCIA ENTRE LA CANTIDAD
   'RECIBIDA Y LA COMPRADA,
   'SI LA CANTIDAD TOTAL RECEPCIONADA PARA ESA ORDEN DE COMPRA CON LA CANTIDAD TOTAL RECIBIDA DE DCICHA ORDEN
   'ES MAYOR , SOLO EN ESTE CASO AJUSTA.
   '
    Has% = HOY(HOS$)
    FEX = Has%
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
10  OKKX% = 0
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
       Exit Sub
    End If
    Des% = CVI(Left$(OBX$, 2))
    Has% = CVI(Mid$(OBX$, 3, 2))
    If Has% < Des% Then GoTo 10
    
    Desde = FETEXCOR1$(Des%)
    Hasta = FETEXCOR1$(Has%)
 

   Screen.MousePointer = 11
   
   CONDI$ = "NUMOC > 0 "
   CONDI$ = CONDI$ + " AND COD_INTE > 0 "
   CONDI$ = CONDI$ + " AND FechRecep >= '" & Desde & "'"
   CONDI$ = CONDI$ + " AND FechRecep <= '" & Hasta & "'"
   '
   'TRAE LA CANTIDAD AGRUPADA POR CODIGO Y ORDEN DE COMPRA
   SQLX$ = "SELECT SUM(CANTOAPRO) AS CANTI,COD_INTE,NUMOC FROM BOLRECEP WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE " & CONDI$
   'SQLX$ = SQLX$ + " and numprov=716"
   SQLX$ = SQLX$ + " GROUP BY NUMOC, COD_INTE "
   FIN& = CantRegistros("BOLRECEP", CONDI$, "NOMESS")
   '
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF
        i& = i& + 1
        Call TRACE(24, i&, FIN&)
        VUELTA% = 0
        CI% = XVALO(!cod_inte)
        OC& = XVALO(!NUMOC)
        CANTI_RECEPCIONADA = XVALO(!CANTI)
        'BUSCO EN OCOMDETA CALCULANDO Y RECORRIENDO POR SI EL ITEM ESTA REPTIDO EN LA O.C.
        CONDI$ = " NUME_OCOM = " & OC&
        CONDI$ = CONDI$ + " AND COD_INTE = " & CI%
        CONDI$ = CONDI$ + " AND STAT_OCOM < 10 "
        'CONDI$ = CONDI$ + " AND CANT_OCOM > CANT_REC"
        '
        CANTIDAD_TOTAL_RECIBIDA_COMPRA = XVALO(VALOBADA("OCOMDETA", "SUM(CANT_REC)", CONDI$, "NOMESS"))
        '
        'GENERA CONSULTA EN OCOMDETA FILTRANDO POR ORDEN DE COMPRA, CODIGO,STATUS < 3 Y CANTIOCOM > CANTIRECIBIDA
'        If (CANTI_RECEPCIONADA) > CANTIDAD_TOTAL_RECIBIDA_COMPRA Then
            SQLX$ = "SELECT * FROM OCOMDETA WITH(NOLOCK) WHERE " & CONDI$
            SQLX$ = SQLX$ + " ORDER BY FENTRE_SOL, CANT_REC "
            '
            Set rs1 = READSET(SQLX$)
            With rs1
              Do While Not .EOF
                CANTICOMPRARECI = XVALO(!Cant_Rec)
                CANTIDAD_COMPRADA = XVALO(!cant_ocom)
                COEFUNI = XVALO(!Coef_Unids)
                If COEFUNI < 0.0001 Then COEFUNI = 1

                SALDOPENDIENTE_COMPRA = CANTIDAD_COMPRADA - CANTICOMPRARECI
                If VUELTA% = 0 Then
                   VUELTA% = 1
                   SALDOPENDIENTE_RECPCION = (CANTI_RECEPCIONADA / COEFUNI) - CANTIDAD_TOTAL_RECIBIDA_COMPRA
                End If
                '
                If SALDOPENDIENTE_RECPCION > 0 Then
                   If SALDOPENDIENTE_COMPRA > SALDOPENDIENTE_RECPCION Then SALDOPENDIENTE_COMPRA = SALDOPENDIENTE_RECPCION
                       CANTIAGRABAR = SALDOPENDIENTE_COMPRA + CANTICOMPRARECI
                       '
                       NUOR% = XVALO(!NuOrd_Item)
                       '
                       TX$ = "Atención!! Se va Actualizar La Cantidad Recibida: (" & CANTICOMPRARECI & ") Del Artículo: " & CODEXT$(CI%)
                       TX$ = TX$ + " \De la Orden de Compra: " & OC& & " Cantidad Comprada: " & CANTIDAD_COMPRADA
                       TX$ = TX$ + " \Ajustando La Cantidad Recibida a " & CANTIAGRABAR
                       TX$ = TX$ + " \\Ajustar\Cancelar\Salir de Proceso"
                       OPX1% = COMALTER%(TX$)
                       If OPX1% = 1 Then
                            CONDI$ = " NUORD_ITEM = " & NUOR%
                            CONDI$ = CONDI$ + " AND COD_INTE = " & CI%
                            CONDI$ = CONDI$ + " AND  NUME_OCOM = " & OC&

                            Call ACTUREGISTRO("OCOMDETA", "CANT_REC", CONDI$, CANTIAGRABAR, REGAF&, "NOMESS")
                            If CANTIAGRABAR >= CANTIDAD_COMPRADA Then
                               Call ACTUREGISTRO("OCOMDETA", "STAT_OCOM", CONDI$, 3, REGAF&, "NOMESS")
                            End If
                            '
                            SALDOPENDIENTE_RECPCION = SALDOPENDIENTE_RECPCION - SALDOPENDIENTE_COMPRA
                        ElseIf OPX1% = 3 Then Screen.MousePointer = 1: rs1.Close: Set rs1 = Nothing: GoTo 99
                       End If
                End If
                If SALDOPENDIENTE_RECPCION <= o Then GoTo 80
                .MoveNext
              Loop
            End With
80          rs1.Close
            Set rs1 = Nothing
'        End If
       .MoveNext
     Loop
   End With
99 On Error Resume Next
   RS.Close
   Set RS = Nothing
   On Error GoTo 0
   Screen.MousePointer = 1
End Sub

Sub Recepcion_AdjuntarDocumentosExternos(NumeroRecepcion As Long)

    If NumeroRecepcion > 0 Then
        RECEDOCEX.txtRecepcion = CStr(NumeroRecepcion)
        RECEDOCEX.Show 1
    Else
        ' LISTA DE RECEPCIONES
        TituloLista$ = "Recepciones de materiales"
        
        Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Recepcion/A12;Fecha/D8;Remito/A20;Proveedor/A32", FRMZELECHO)
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Recepcion/A12;Fecha/D8;Remito/A20;Proveedor/A32", FRMZELECHO)
        FRMZELECHO.Caption = TituloLista$
        
        Dim obj As New RECORXET
        Set rs1 = obj.RS_BOLRECEPEN()
        Call Carga_MSF_Recordset(rs1, "Recepcion/A12;Fecha/D8;Remito/A20;Proveedor/A32", FRMZELECHO.msf2, 1)
        On Error Resume Next
        Set obj = Nothing
        rs1.Close
        Set rs1 = Nothing
        On Error GoTo 0
            
        FRMZELECHO.Width = 10000
50      FRMZELECHO.Show 1
        
        NroRecepcion$ = TRIM(RESP_ZELE_VECT(1))
        If NroRecepcion$ = "" Then Exit Sub
        
        RECEDOCEX.txtRecepcion = NroRecepcion$
        RECEDOCEX.Show 1
    End If
    
End Sub

Function RecepcionDeBobinas(NumeroRecepcion&, NumeroProveedor%, Remito$, TipoRecepcion%) As Boolean

    RecepcionDeBobinas = False
    
    EPAC$ = UCase$(TRIM$(EMPREAC$))
    SOYCAMBACERES% = 0
    If InStr(1, EMPREAC$, "CAMBACERES") > 0 Then SOYCAMBACERES% = 1
    
    Call COPYSTRU("BOLREPAP", "BOLRETMP")
    Call SETULTREG("BOLRETMP", 9)
    If ULTREG&("BOLRETMP") > 0 Then
        OpcionRecuperacion% = COMALTER%("Se Ha Detectado Una Recepción Pendiente de Registración.\¿Desea Continuar Ingresando?\\Si, Continuar\No, Cancelar")
        If OpcionRecuperacion% = 2 Then Call SETULTREG("BOLRETMP", 0)
    End If
    
    IDLIST& = ObtenerIdTablaAuxiliar("TARATIRECE")
    IDTARA% = XVALO(VALOBADA("datasql", "campo3", "id_lista=" & CStr(IDLIST&) & " and campo1=" & CStr(TipoRecepcion%)))
    IDLIST& = ObtenerIdTablaAuxiliar("TABLATARAS")
    Tara# = XVALO(VALOBADA("datasql", "campo3", "id_lista=" & CStr(IDLIST&) & " and campo1=" & CStr(IDTARA%)))
    
    FORIPRE$ = TRIM$(Control$("", "STIKAPRO"))
    If EXISTE%(LUDAT$ + FORIPRE$ + ".SQR") = 0 Then Exit Function
    
    On Error Resume Next
    Unload BALANZA
    Unload VENTABNEW
    On Error GoTo 0
    
    VENTANA.Inicializar = 0
    VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 0
    VENTANA.HabilitaInsertar = 1
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "RECEPCION DE BOBINAS"
      VENTANA.CampEditables = "2;9"
    VENTANA.ColEcoCampo(2) = "codigo;descripcion;master;status>0"

    VENTANA.Campox = "Orden/I5;Codigo/A16;Descripcion/A32;UM/A3;O-C/A8;Pendiente/F10.1;Ord OC/I5;Entrega/D8;Metros/F10.1;Peso/F0;Proce/A0"
    VENTANA.Inicializar = 1
        
    ' *** CARGA LAS ORDENES DE COMPRAS PENDIENTE
    J& = 0
    
    FIN& = ULTREG&("!BOLREPAP")
    For YX& = 1 To FIN&
        YY$ = REGLEIDO$("!BOLREPAP", YX&)
        CI% = CVI(Mid$(YY$, 83, 2))
        
        NUOC& = CVS(Mid$(YY$, 151, 4))
        CAN# = CVD(Mid$(YY$, 163, 8))
        ORDOC% = Asc(Mid$(YY$, 98, 1))
        FE% = CVI(Mid$(YY$, 155, 2))
        
        With VENTABNEW.MSF
            J& = J& + 1
            .Rows = J& + 1
            .TextMatrix(J&, 1) = J&
            .TextMatrix(J&, 2) = CODEXT$(CI%)
            .TextMatrix(J&, 3) = DESCRIT0$(CI%)
            .TextMatrix(J&, 4) = Unimed$(CI%)
            .TextMatrix(J&, 5) = CStr(NUOC&)
            .TextMatrix(J&, 6) = TRIM$(FORMATNUM$(CAN#, "F10.1"))
            .TextMatrix(J&, 7) = CStr(ORDOC%)
            .TextMatrix(J&, 8) = FECHATEX$(FE%)
            .TextMatrix(J&, 9) = ""
        End With
    Next YX&
    
10  VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
        If ValidayoAplicaOrdenCompra% < 1 Then GoTo 10
12      JJ& = ULTREG&("BOLRETMP")
        For i = 1 To VENTABNEW.MSF.Rows - 1
            If TRIM$(VENTABNEW.MSF.TextMatrix(i, 9)) <> "" Then ' VERIFICA QUE TENGA METROS
                CI% = CODINT%(VENTABNEW.MSF.TextMatrix(i, 2))
                
                ' *** BALANZA
                BALANZA.CodigoProducto = CODINT%(VENTABNEW.MSF.TextMatrix(i, 2))
                BALANZA.txtCodigo = VENTABNEW.MSF.TextMatrix(i, 2)
                BALANZA.txtDescripcion = VENTABNEW.MSF.TextMatrix(i, 3)
                BALANZA.Show 1
        
                Peso# = 0
                If BALANZA.Respuesta = "APROBO" Then Peso# = XVALO(BALANZA.txtPeso)
                Unload BALANZA
                If Peso# < 0.005 Then GoTo 15
                Peso# = Peso# - Tara#
                
                ' *** CONTRASTA LOS METROS CALCULADOS A PARTIR DEL PESO EN BALANZA Y EL PESO POR METRO DEL ARTICULO,
                ' *** CON EL PESO TEORICO INGRESADO POR EL USUARIO A PARTIR DEL REMITO DEL PROVEEDOR
                LargoBobina# = XVALO(VENTABNEW.MSF.TextMatrix(i, 9))
                If Control("BOLRECEP", "CTRLPESO") = "SI" Then
                     PesoPorMetro# = PESMETRO(CI%)
                     SobreEmbarque# = XVALO(Control("BOLRECEP", "PORSOBRE"))
                     If PesoPorMetro# > 0.0005 And SobreEmbarque# > 0.05 Then
                          PesoBalanza# = Peso#
                          MetrosCalculados# = PesoBalanza# / PesoPorMetro#
                          
                          If Abs(LargoBobina# - MetrosCalculados#) > LargoBobina# * (SobreEmbarque# / 100) Then
                             If DERACCE%("RECSOBRE/NOMESS", "Recepción con Sobre-Embarque") >= 2 Then
                                 Mensaje$ = "Se Ha Detectado una Diferencia\Entre los Metros Configurados Para la Bobina\y los Calculados a partir del Peso\Para el Código '" & CODEXT$(CI%) & "'\ \Configurado: " & TRIM$(FORMATNUM$(LargoBobina#, "F10.2")) & "\Calculado: " & TRIM$(FORMATNUM$(MetrosCalculados#, "F10.2"))
                                 If COMALTER%(Mensaje$ + ".\\Corregir\Continuar") = 1 Then GoTo 15
                             Else
                                 Call MENSERR(24, "Configuración y Derechos de Acceso\Impiden Sobre-Embarque\en Código " + DECODE$ + ".")
                                 GoTo 15
                             End If
                          End If
                     End If
                End If

                ' *** GRABA DATOS DE LA BOBINA PROCESADA CORRECTAMENTE
                CAPEN = XVALO(VENTABNEW.MSF.TextMatrix(i, 6))
                ORDOC% = XVALO(VENTABNEW.MSF.TextMatrix(i, 7))
                NUOC& = XVALO(VENTABNEW.MSF.TextMatrix(i, 5))
                FENTRESO% = FECHANUM(XVALO(VENTABNEW.MSF.TextMatrix(i, 8)))
                AnchoBobina# = 1000
                
                Y$ = REGBLAN("BOLREPAP")
                Call REPLA(Y$, MKI$(CI%), 83, 2)
                Call REPLA(Y$, Unimed$(CI%), 85, 4)
                  CANTIX# = LargoBobina#
                If SOYCAMBACERES% > 0 And Unimed$(CI%) = "U." Then
                   CANTIX# = 1
                End If
                Call REPLA(Y$, MKD$(CDbl(CANTIX#)), 89, 8) '>>>>>>>>> aca antes tenia LargoBobina# si no es camba va a mantener el mismo valor
                   If ANUCA% = 1 Then Call REPLA(Y$, MKD$(0), 89, 8)  'si es camba y unimed = u. pone 1 y mantiene el mt en LargoBobina# que lo
                Call REPLA(Y$, MKD$(CDbl(CAPEN)), 163, 8)             'termina grabando en la aprobacion en el movisto con los metros.
                '                                                     'clave seguimiento: RecepcionCambaBobina
                Call REPLA(Y$, Chr$(XVALO(ORDOC%)), 98, 1)
                Call REPLA(Y$, MKS$(0), 100, 4) 'no tiene AE
                '
                Call REPLA(Y$, LOTECO$, 107, 12)
                Call REPLA(Y$, MKS$(NUOC&), 151, 4)
                Call REPLA(Y$, MKI$(FENTRESO%), 155, 2)
                Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
                Call REPLA(Y$, MKS$(0), 189, 4)

                Call REPLA(Y$, MKD$(AnchoBobina#), 193, 8)
                Call REPLA(Y$, MKD$(LargoBobina#), 201, 8)
                Call REPLA(Y$, MKS$(1), 209, 4)
                Call REPLA(Y$, MKS$(Peso#), 253, 4)
                
                JJ& = JJ& + 1
                Call GRAREG("BOLRETMP", Y$, JJ&)
                
                ' *** MARCA REGISTRO / BOBINA COMO PROCESADO OK
                VENTABNEW.MSF.TextMatrix(i, 11) = "PROCESADO"
                
                ' *** IMPRIME ETIQUETA
                PFAC$ = Control$(IDENTER$, "PORTSTIK")
                If PFAC$ = "" Then PFAC$ = Control$("*", "PORTSTIK")
                Call SETPRINPORT(PFAC$)
                Codigo$ = VENTABNEW.MSF.TextMatrix(i, 2)
                CodInte% = CODINT%(Codigo$)
                
                Descripcion$ = VENTABNEW.MSF.TextMatrix(i, 3)
                LOTE$ = "BR-" & CStr(NumeroRecepcion&) & "-" & CStr(JJ&)
                Largo# = LargoBobina#
                Proveedor$ = rasocli$(NumeroProveedor% * -1)
                familia1$ = TRIM$(UCase$(CODFAMIL$(CodInte%)))
                If TIMATE%(CodInte%) = 41 And familia1$ = "TE" Then  ' SI ES BOBINA -- TELA 08_02_24 (HABIA IMPRESO ASI PARA ESLINGAS)
                  TEXTOQR$ = Codigo$ & "| |" & LOTE$
                Else
                  TEXTOQR$ = Codigo$ & "|" & LOTE$
                End If
                IMAG$ = "IMAGRECE_QR.JPG"
                IDENTIFICADOR$ = "QR_RECEPCION"
                ANCHOX% = 0
                ALTOX% = 0
                '
                On Error Resume Next
                Kill LUCOM$ & IMAG$
                On Error GoTo 0
                'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
                QRX$ = TRIM$(IMAGENCODBARRA$(LUCOM$, IMAG$, TEXTOQR$, "QR_CODE", ANCHOX%, ALTOX%))
                Call T_Espera(3)
                Image1.Picture = LoadPicture(LUCOM$ & IMAG$)
                Call GRABAR_QR(LUCOM$ & IMAG$, IDENTIFICADOR$)
                If SOYCAMBACERES% < 1 Then
                    FILTX$ = LOTE$ & ";" & Codigo$ & ";" & Descripcion$ & ";" & Proveedor$ & ";" & TRIM$(FORMATNUM$(Largo#, "F12.1")) & ";" & TRIM$(FORMATNUM$(Peso#, "F12.1")) & ";" & IDENTIFICADOR$ & ";" & CStr(USUARIO%) & ";" & Remito$ & ";" & FECHATEX$(HOY(HO$))
                    Call STDFORM(FORIPRE$, FILTX$, "PRINT")
                    FILTX$ = LOTE$ & ";" & Codigo$ & ";" & Descripcion$ & ";" & Proveedor$ & ";" & TRIM$(FORMATNUM$(Largo#, "F12.1")) & ";" & TRIM$(FORMATNUM$(Peso#, "F12.1")) & ";" & IDENTIFICADOR$ & ";" & CStr(USUARIO%) & ";" & Remito$ & ";" & FECHATEX$(HOY(HO$))
                    Call STDFORM(FORIPRE$, FILTX$, "PRINT")
                Else
                    'ACA CAMBACERES USA LA MISMA ETIQUETA DE RECEPCION DE CONTEPLAST
                    FILTX$ = LOTE$ & ";" & Codigo$ & ";" & Descripcion$ & ";" & Proveedor$ & ";" & TRIM$(FORMATNUM$(Largo#, "F12.1")) & ";" & TRIM$(FORMATNUM$(Peso#, "F12.1")) & ";" & IDENTIFICADOR$ & ";" & CStr(USUARIO%) & ";" & Remito$
                    Call STDFORM("ETRE-CMB", FILTX$, "PRINT")
                    FILTX$ = LOTE$ & ";" & Codigo$ & ";" & Descripcion$ & ";" & Proveedor$ & ";" & TRIM$(FORMATNUM$(Largo#, "F12.1")) & ";" & TRIM$(FORMATNUM$(Peso#, "F12.1")) & ";" & IDENTIFICADOR$ & ";" & CStr(USUARIO%) & ";" & Remito$
                    Call STDFORM("ETRE-CMB", FILTX$, "PRINT")
                End If
                  Call SETPRINPORT("RESTORE")
                PRINTERPORT$ = ""
            End If
15      Next i
        
        ' *** ELIMINA LOS PROCESADOS CORRECTAMENTE
        For i = VENTABNEW.MSF.Rows - 1 To 1 Step -1
            If VENTABNEW.MSF.TextMatrix(i, 11) = "PROCESADO" Then
                If i > 1 Or (i = 1 And VENTABNEW.MSF.Rows > 2) Then
                    VENTABNEW.MSF.RemoveItem (i)
                Else
                    VENTABNEW.MSF.Rows = 1
                End If
            End If
        Next i
        
        If VENTABNEW.MSF.Rows = 1 Then VENTABNEW.MSF.Rows = 2
        
        GoTo 10
    Else
        If ULTREG&("BOLRETMP") > 0 Then
            Call SETULTREG("BOLREPAP", 0)
            Call TRALOCAL("BOLRETMP", "BOLREPAP")
            Call SETULTREG("BOLRETMP", 0)
            
            ' *** VERIFICA SI TODOS LOS ARTICULOS SON TRAZABLES. CASO CONTRARIO LOS CONFIGURA.
            For IX& = 1 To ULTREG&("!BOLREPAP")
                XX$ = REGLEIDO$("!BOLREPAP", IX&)
                CI% = CVI(Mid$(XX$, 83, 2))
                If ESTRAZABLE%(CI%) <> 1 Then
                    Call ACTUREGISTRO("MASTER", "TRAZABLE", "CODINT=" & CStr(CI%), "1", REGAF&)
                End If
            Next IX&
            
            RecepcionDeBobinas = True
        End If
        
        On Error Resume Next
        Unload BALANZA
        Unload VENTABNEW
        On Error GoTo 0
    End If

End Function

Function RecepcionDeBobinasCambaQR(NumeroRecepcion&, NumeroProveedor%, Remito$, TipoRecepcion%) As Boolean
    'RECEPCION DE BOBINAS EN ENVAPLAST DE CAMBACERES CON EL CODIGO QR QUE IMPRIMEN EN CAMBACERES CUANDO HACEN EL REMITO
    RecepcionDeBobinasCambaQR = False
    Call COPYSTRU("BOLREPAP", "BOLRETMP")
    FORIPRE$ = "ETR2-CTP"
    If EXISTE%(LUDAT$ + FORIPRE$ + ".SQR") = 0 Then Exit Function
    
    IMAG$ = "IMAGRECE_QR.JPG"
    IDENTIFICADOR$ = "QR_RECEPCION"
    ANCHOX% = 0
    ALTOX% = 0

    On Error Resume Next
    Unload VENTABNEW
    On Error GoTo 0
    
    
    VENTANA.Inicializar = 0
    VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 0
    VENTANA.HabilitaInsertar = 1
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "RECEPCION DE BOBINAS"
    VENTANA.Campox = "Codigo/A16;Descripcion/A32;Entrega/F8;Metros/F10.1;Peso/F10.1;REMITO/A18;Lote Proveedor"
    VENTANA.CampEditables = "1;3;4;5;6;7"
    VENTANA.ColEcoCampo(1) = "codigo;descripcion;master;status>0"

    VENTANA.Inicializar = 1
  
10  VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
        Call SETULTREG("BOLRETMP", 0)
        JJ& = ULTREG&("BOLRETMP")
        For i = 1 To VENTABNEW.MSF.Rows - 1
            If TRIM$(VENTABNEW.MSF.TextMatrix(i, 4)) <> "" Then ' VERIFICA QUE TENGA METROS
                CI% = CODINT%(VENTABNEW.MSF.TextMatrix(i, 1))
        
                Peso# = XVALO(VENTABNEW.MSF.TextMatrix(i, 5))
                ' *** GRABA DATOS DE LA BOBINA PROCESADA CORRECTAMENTE
                CAPEN = 0
                ORDOC% = 0
                NUOC& = 0
                FENTRESO% = HOY(HO$)
                AnchoBobina# = 1000
                LargoBobina# = XVALO(VENTABNEW.MSF.TextMatrix(i, 4))
                VENTABNEW.MSF.TextMatrix(i, 6) = UCase$(REPLACAR(VENTABNEW.MSF.TextMatrix(i, 6), "'", "-"))
                Y$ = REGBLAN("BOLREPAP")
                Call REPLA(Y$, MKI$(CI%), 83, 2)
                Call REPLA(Y$, Unimed$(CI%), 85, 4)
                Call REPLA(Y$, MKD$(CDbl(LargoBobina#)), 89, 8)
                   If ANUCA% = 1 Then Call REPLA(Y$, MKD$(0), 89, 8)
                Call REPLA(Y$, MKD$(CDbl(CAPEN)), 163, 8)
                '
                Call REPLA(Y$, Chr$(XVALO(ORDOC%)), 98, 1)
                Call REPLA(Y$, MKS$(0), 100, 4) 'no tiene AE
                '
                LOTECO$ = UCase$(REPLACAR(VENTABNEW.MSF.TextMatrix(i, 7), "'", "-"))
                Call REPLA(Y$, LOTECO$, 107, 12)
                Call REPLA(Y$, MKS$(NUOC&), 151, 4)
                Call REPLA(Y$, MKI$(FENTRESO%), 155, 2)
                Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
                Call REPLA(Y$, MKS$(0), 189, 4)

                Call REPLA(Y$, MKD$(AnchoBobina#), 193, 8)
                Call REPLA(Y$, MKD$(LargoBobina#), 201, 8)
                Call REPLA(Y$, MKS$(1), 209, 4)
                Call REPLA(Y$, MKS$(Peso#), 253, 4)
                
                JJ& = JJ& + 1
                Call GRAREG("BOLRETMP", Y$, JJ&)
                
                ' *** MARCA REGISTRO / BOBINA COMO PROCESADO OK
                
                ' *** IMPRIME ETIQUETA
                PFAC$ = Control$(IDENTER$, "PORTSTIK")
                If PFAC$ = "" Then PFAC$ = Control$("*", "PORTSTIK")
                Call SETPRINPORT(PFAC$)
    
                Codigo$ = VENTABNEW.MSF.TextMatrix(i, 1)
                Descripcion$ = VENTABNEW.MSF.TextMatrix(i, 2)
                LOTE$ = "BR-" & CStr(NumeroRecepcion&) & "-" & CStr(JJ&)
                Largo# = LargoBobina#
                Proveedor$ = rasocli$(NumeroProveedor% * -1)
                LoteProveedor$ = LOTECO$ 'el lote de cambaceres lo lee mas arriba
                If TRIM$(LoteProveedor$) = "" Then LoteProveedor$ = "-"
                '
                REMITOYLOTEPROVEEDOR$ = TRIM$(Left$(Remito$, 18)) & " " & LoteProveedor$
                TEXTOQR$ = Codigo$ & "|" & LOTE$
                IMAG$ = "IMAGRECE_QR.JPG"
                IDENTIFICADOR$ = "QR_RECEPCION"
                ANCHOX% = 0
                ALTOX% = 0

                On Error Resume Next
                Kill LUCOM$ & IMAG$
                On Error GoTo 0
                'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
                QRX$ = TRIM$(IMAGENCODBARRA$(LUCOM$, IMAG$, TEXTOQR$, "QR_CODE", ANCHOX%, ALTOX%))
                Call T_Espera(3)
                Image1.Picture = LoadPicture(LUCOM$ & IMAG$)
                Call GRABAR_QR(LUCOM$ & IMAG$, IDENTIFICADOR$)
        
                FILTX$ = LOTE$ & ";" & Codigo$ & ";" & Descripcion$ & ";" & Proveedor$ & ";" & TRIM$(FORMATNUM$(Largo#, "F12.1")) & ";" & TRIM$(FORMATNUM$(Peso#, "F12.1")) & ";" & REMITOYLOTEPROVEEDOR$
                FILTX1$ = LOTE$ & ";" & Codigo$ & ";" & Descripcion$ & ";" & Proveedor$ & ";" & TRIM$(FORMATNUM$(Largo#, "F12.1")) & ";" & TRIM$(FORMATNUM$(Peso#, "F12.1")) & ";" & REMITOYLOTEPROVEEDOR$
                Call STDFORM(FORIPRE$, FILTX$, "PRINT")
                Call STDFORM(FORIPRE$, FILTX1$, "PRINT")

                Call SETPRINPORT("RESTORE")
                PRINTERPORT$ = ""
            End If
15      Next i
        
        If ULTREG&("BOLRETMP") > 0 Then
            Call SETULTREG("BOLREPAP", 0)
            Call TRALOCAL("BOLRETMP", "BOLREPAP")
            Call SETULTREG("BOLRETMP", 0)
            
'            ' *** VERIFICA SI TODOS LOS ARTICULOS SON TRAZABLES. CASO CONTRARIO LOS CONFIGURA.
'            For IX& = 1 To ULTREG&("!BOLREPAP")
'                XX$ = REGLEIDO$("!BOLREPAP", IX&)
'                CI% = CVI(Mid$(XX$, 83, 2))
'                If ESTRAZABLE%(CI%) <> 1 Then
'                    Call ACTUREGISTRO("MASTER", "TRAZABLE", "CODINT=" & CStr(CI%), "1", REGAF&)
'                End If
'            Next IX&
            
            RecepcionDeBobinasCambaQR = True
        End If
        
        On Error Resume Next
        Unload VENTABNEW
        On Error GoTo 0
    End If

End Function

Sub PRINT_ETIQRECEPCAMBA(Codigo$, LOTE$, Cantidad#, Proveedor$, Remito$, Peso#)
                FORIPRE$ = TRIM$(Control$("", "STIKAPRO"))
                If EXISTE%(LUDAT$ + FORIPRE$ + ".SQR") = 0 Then Exit Sub

                IMAG$ = "IMAGRECE_QR.JPG"
                IDENTIFICADOR$ = "QR_RECEPCION"
                ANCHOX% = 0
                ALTOX% = 0
'                TEXTOQR$ = Codigo$ & "|" & LOTE$ & "|" & Cantidad & "|" & Proveedor$ & "|" & Remito$ & "|" & TRIM$(FORMATNUM$(Peso#, "F12.1"))
                
                TEXTOQR$ = Codigo$ & "|" & LOTE$ & "|" & Proveedor$ & "|" & TRIM$(FORMATNUM$(Cantidad#, "F12.1")) & "|" & TRIM$(FORMATNUM$(Peso#, "F12.1")) & "|" & Remito$ & "|" & FECHATEX$(HOY(HO$))

                On Error Resume Next
                Kill LUCOM$ & IMAG$
                On Error GoTo 0
                
                'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
                QRX$ = TRIM$(IMAGENCODBARRA$(LUCOM$, IMAG$, TEXTOQR$, "QR_CODE", ANCHOX%, ALTOX%))
                Call T_Espera(3)
                Image1.Picture = LoadPicture(LUCOM$ & IMAG$)
                Call GRABAR_QR(LUCOM$ & IMAG$, IDENTIFICADOR$)
        
                FILTX$ = LOTE$ & ";" & Codigo$ & ";" & Descripcion$ & ";" & Proveedor$ & ";" & TRIM$(FORMATNUM$(Cantidad#, "F12.1")) & ";" & TRIM$(FORMATNUM$(Peso#, "F12.1")) & ";" & IDENTIFICADOR$ & ";" & CStr(USUARIO%) & ";" & Remito$
                Call STDFORM(FORIPRE$, FILTX$, "PRINT")
                FILTX$ = LOTE$ & ";" & Codigo$ & ";" & Descripcion$ & ";" & Proveedor$ & ";" & TRIM$(FORMATNUM$(Cantidad#, "F12.1")) & ";" & TRIM$(FORMATNUM$(Peso#, "F12.1")) & ";" & IDENTIFICADOR$ & ";" & CStr(USUARIO%) & ";" & Remito$
                Call STDFORM(FORIPRE$, FILTX$, "PRINT")

                Call SETPRINPORT("RESTORE")
                PRINTERPORT$ = ""

End Sub
Function RecepcionDeMateriales(idRecepcion As Long) As Boolean

    RecepcionDeMateriales = False
    
    idRecepcion = 0
    
    On Error Resume Next
    Unload BALANZA
    Unload VENTABNEW
    On Error GoTo 0
   
    VENTANA.Inicializar = 0
    VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 0
    VENTANA.HabilitaInsertar = 1
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "RECEPCION DE MATERIALES"
    VENTANA.ColEcoCampo(2) = "codigo;descripcion;master;status>0"

    VENTANA.Campox = "Orden/I5;Codigo/A16;Descripcion/A32;UM/A3;O-C/A0;Ord OC/I0;Entrega/D0;Pendiente/F0;Lote/A24;Recibido/F10.1"
    VENTANA.CampEditables = "2;9;10"
    If InStr(1, EMPREAC$, "RETEPLAST") > 0 Then
        VENTANA.Campox = "Orden/I5;Codigo/A16;Descripcion/A50;UM/A3;O-C/A0;Ord OC/I0;Entrega/D0;Pendiente/F0;Lote/A12;Control/F10.1;Remito/F10.1;Balanza/F10.1;Bultos/I6"
        VENTANA.CampEditables = "2;9;10;11;12;13"
    End If
    VENTANA.Inicializar = 1
        
    ' *** CARGA LAS ORDENES DE COMPRAS PENDIENTE
    J& = 0
    FIN& = ULTREG&("!BOLREDET")
    For YX& = 1 To FIN&
        YY$ = REGLEIDO$("!BOLREDET", YX&)
        CI% = CVI(Mid$(YY$, 83, 2))
        
        NUOC& = CVS(Mid$(YY$, 151, 4))
        CAN# = CVD(Mid$(YY$, 163, 8))
        ORDOC% = Asc(Mid$(YY$, 98, 1))
        FE% = CVI(Mid$(YY$, 155, 2))
        
        With VENTABNEW.MSF
            J& = J& + 1
            .Rows = J& + 1
            .TextMatrix(J&, 1) = J&
            .TextMatrix(J&, 2) = CODEXT$(CI%)
            .TextMatrix(J&, 3) = DESCRIT0$(CI%)
            .TextMatrix(J&, 4) = Unimed$(CI%)
            .TextMatrix(J&, 5) = CStr(NUOC&)
            .TextMatrix(J&, 6) = CStr(ORDOC%)
            .TextMatrix(J&, 7) = FECHATEX$(FE%)
            .TextMatrix(J&, 8) = TRIM$(FORMATNUM$(CAN#, "F10.1"))
            .TextMatrix(J&, 9) = ""
            .TextMatrix(J&, 10) = ""
        End With
    Next YX&
    
10  VENTABNEW.Show 1
    If VENTABNEW.APROBADO% > 0 Then
        Call COPYSTRU("BOLREPAP", "BOLRETMP")
        JJ& = ULTREG&("BOLRETMP")
        If VENTABNEW.MSF.Rows < 2 Then
            RecepcionDeMateriales = False
            GoTo 99
        End If
        
        ordenId% = ordenId% + 1
15      idRecepcion = XVALO(Format(Now, "yyyymmdd")) * 100 + ordenId%
        If CantRegistros("BOLRETMP", "idRecepcion=" & idRecepcion) > 0 Then
            ordenId% = ordenId% + 1
            If ordenId% > 99 Then GoTo 99
            GoTo 15
        End If
        On Error GoTo ControlError
        Dim rstRecepcion As New ADODB.Recordset
        rstRecepcion.Open FILTSQL$("select * from bolretmp"), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With rstRecepcion
            For i = 1 To VENTABNEW.MSF.Rows - 1
                CodigoInterno% = CODINT%(VENTABNEW.MSF.TextMatrix(i, 2))
                NroOrdenCompra& = XVALO(VENTABNEW.MSF.TextMatrix(i, 5))
                OrdenItemOCompra% = XVALO(VENTABNEW.MSF.TextMatrix(i, 6))
                FechaEntrega% = FECHANUM(XVALO(VENTABNEW.MSF.TextMatrix(i, 7)))
                CantPendiente# = XVALO(VENTABNEW.MSF.TextMatrix(i, 8))
                LoteMaterial$ = VENTABNEW.MSF.TextMatrix(i, 9)
                CantRecibida# = XVALO(VENTABNEW.MSF.TextMatrix(i, 10))
                Cantidad_1# = XVALO(VENTABNEW.MSF.TextMatrix(i, 11))
                Cantidad_2# = XVALO(VENTABNEW.MSF.TextMatrix(i, 12))
                CantBultos% = XVALO(VENTABNEW.MSF.TextMatrix(i, 13))
                
                .AddNew
                !idRecepcion = idRecepcion
                !ORDENITEM = i
                !cod_inte = CodigoInterno%
                !ORDENCOMPRA = NroOrdenCompra&
                !OrdenItemCompra = OrdenItemOCompra%
                !FechaEntrega = CVDAT(FechaEntrega%)
                !CantidadPendiente = CantPendiente#
                !LOTE = LoteMaterial$
                !Recibido = CantRecibida#
                !Cantidad_1 = Cantidad_1#
                !Cantidad_2 = Cantidad_2#
                !cantidadBultos = CantBultos%
                !Status = 0
                .Update
                
                Y$ = REGBLAN("BOLREDET")
                Call REPLA(Y$, MKI$(CodigoInterno%), 83, 2)
                Call REPLA(Y$, Unimed$(CodigoInterno%), 85, 4)
                Call REPLA(Y$, MKD$(CDbl(CantRecibida#)), 89, 8)
                Call REPLA(Y$, MKD$(CDbl(CantPendiente#)), 163, 8)
                '
                Call REPLA(Y$, Chr$(OrdenItemOCompra%), 98, 1)
                Call REPLA(Y$, MKS$(0), 100, 4)
                '
                Call REPLA(Y$, LoteMaterial$, 107, 12)
                Call REPLA(Y$, MKS$(NroOrdenCompra&), 151, 4)
                Call REPLA(Y$, MKI$(FechaEntrega%), 155, 2)
                Call REPLA(Y$, MKI$(FechaEntrega%), 187, 2)
                Call REPLA(Y$, MKS$(0), 189, 4)

                Call REPLA(Y$, MKD$(0), 193, 8)
                Call REPLA(Y$, MKD$(0), 201, 8)
                Call REPLA(Y$, MKS$(1), 209, 4)
                Call REPLA(Y$, MKS$(Cantidad_1#), 213, 4)
                Call REPLA(Y$, MKS$(Cantidad_2#), 217, 4)
                Call REPLA(Y$, MKI$(CantBultos%), 221, 2)
                Call REPLA(Y$, MKS$(0), 253, 4)
                
                JJ& = JJ& + 1
                Call GRAREG("BOLRETMP", Y$, JJ&)
            Next i
        
            On Error Resume Next
            rstRecepcion.Close
            Set rstRecepcion = Nothing
            On Error GoTo 0
        
            If ULTREG&("BOLRETMP") > 0 Then
                Call SETULTREG("BOLREDET", 0)
                Call TRALOCAL("BOLRETMP", "BOLREDET")
                Call SETULTREG("BOLRETMP", 0)
                Call ACTUREGISTRO("BOLRETMP", "status", "idrecepcion=" & idRecepcion, 3, REGAF&)
                
                RecepcionDeMateriales = True
            End If
            On Error GoTo 0
            
            On Error Resume Next
            Unload VENTABNEW
            On Error GoTo 0
        End With
    End If

99  Exit Function

ControlError:
    idRecepcion = 0
    GoTo 99

End Function
Sub Reimpre_Etiquetas()

    PRINTETI.Show 1
    EPAC$ = UCase$(TRIM$(EMPREAC$))
    If InStr(1, EPAC$, "PERGAMINO") > 0 Then Exit Sub

'=======================================================================================================
'=======================================================================================================
'=======================================================================================================
'=======================================================================================================

   FORIPRE$ = TRIM$(Control("BOLRECEP", "ETIQRECE"))
   
   If FORIPRE$ = "" Then
       Call COMUNI("Falta Configurar Formulario de Impresión de Etiquetas")
       Exit Sub
   ElseIf FORIPRE$ <> "" Then
        PRINTERPORT$ = ""
        PFAC$ = Control$(IDENTER$, "PORTSTIK") ' IMPRESORA CONFIGURADA
        If PFAC$ = "" Then PFAC$ = Control$("*", "PORTSTIK")
        Call SETPRINPORT(PFAC$)
   End If
10 CONDI$ = "NUMEBORE > 0 ORDER BY NUMEBORE"
    i& = 0
    
   Campos$ = "N° Rec./F8;Cta./F6;Proveedor/A32;F.Rec./D8"
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO, , , ANTOT)
   FRMZELECHO.Width = 9000
   FRMZELECHO.Caption = "Re-Imprimir Etiquetas"
   SQLX$ = "SELECT DISTINCT NUMEBORE,NUMPROV,RASO_PROV,FECHRECEP  FROM BOLRECEP WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " & CONDI$
   Set rs1 = READSET(SQLX$)
   With rs1
      Do While Not .EOF
        
        i& = i& + 1
        FRMZELECHO.msf2.Rows = i& + 1
        FRMZELECHO.msf2.TextMatrix(i&, 1) = XVALO(!NUMEBORE)
        FRMZELECHO.msf2.TextMatrix(i&, 2) = XVALO(!NumProv)
        FRMZELECHO.msf2.TextMatrix(i&, 3) = SAFETEXT$(!Raso_Prov)
        FRMZELECHO.msf2.TextMatrix(i&, 4) = FECHATEX$(CVINT(!FechRecep))
        .MoveNext
      Loop
   End With
   rs1.Close
   Set rs1 = Nothing
   FRMZELECHO.Show 1
      
      
   NUMEBOREX = XVALO(RESP_ZELE_VECT(1))

   i& = 0
   
30 If NUMEBOREX > 0 Then
        Campos$ = "Lote/A12;Código/A16;Descripción/A32;Cantidad/F0;N° Etiqueta/A30"
        Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
        Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO, , , ANTOT)
        FRMZELECHO.Width = 12000
        FRMZELECHO.Caption = "Re-Imprimir Etiquetas"

        CONDI$ = "NUMEBORE = " & NUMEBOREX
        SQLX$ = "SELECT * FROM BOLRECEP WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE " & CONDI$
        SQLX$ = SQLX$ + " ORDER BY IDENLOTE,COD_INTE"
        Set RS = READSET(SQLX$)
        With RS
          Do While Not .EOF
            CAN# = XVALO(!CantNom)
            CIXI% = XVALO(!cod_inte)
            PESOX = PESUNI(CIXI%)
            CANTIETIQ = CANBULTOSXKG(CIXI%, CAN#)
            For i& = 1 To CANTIETIQ
              FRMZELECHO.msf2.Rows = i& + 1
              FRMZELECHO.msf2.TextMatrix(i&, 1) = SAFETEXT$(!idenlote)
              FRMZELECHO.msf2.TextMatrix(i&, 2) = SAFETEXT$(!Cod_Exte)
              FRMZELECHO.msf2.TextMatrix(i&, 3) = SAFETEXT$(!descripitem)
              FRMZELECHO.msf2.TextMatrix(i&, 4) = i&
              FRMZELECHO.msf2.TextMatrix(i&, 5) = "Etiqueta " & i& & " de " & CANTIETIQ
            Next i&
            .MoveNext
          Loop
        End With
        RS.Close
        Set rs1 = Nothing
        FRMZELECHO.Show 1
        LOTE$ = TRIM$(RESP_ZELE_VECT(1))
50      If LOTE$ = "" Then Exit Sub
        CODI$ = TRIM$(RESP_ZELE_VECT(2))
        ORDEN% = XVALO(RESP_ZELE_VECT(4))
        TXORDEN$ = RESP_ZELE_VECT(5)
        TXTX$ = CODI$ & "|" & LOTE$ & "|" & PESOX
        FILTX$ = LOTE$ & ";" & TRIM$(Str$(ORDEN%) & ";" & TXTX$ & ";" & TXORDEN$)
        Call STDFORM(FORIPRE$, FILTX$, "PRINT")
        Call SETPRINPORT("RESTORE")
        '
        FRMZELECHO.Show 1
        LOTE$ = TRIM$(RESP_ZELE_VECT(1))
        If LOTE$ = "" Then GoTo 10
   
        If NUMEBOREX <> "" Then GoTo 50

      End If
      
      
   
End Sub

Sub ReimpresionEtiquetasParaBobinas(NroRecepcion As Long, Optional IncluyeQR As Boolean)

    FORIPRE$ = TRIM$(Control$("", "STIKAPRO"))
    EPAC$ = UCase$(TRIM$(EMPREAC$))
    
    If InStr(1, EPAC$, "CONTEPLAST") > 0 Then
       FORIPRE$ = "ETR2-CTP" 'este tiene la fecha que la lee de la b.d.
    End If
    
    If EXISTE%(LUDAT$ + FORIPRE$ + ".SQR") = 0 Then
        Call COMUNI("Falta Formulario de Impresión de Etiqueta")
        Exit Sub
    End If
    
    Call CREATABLA_ARCHIVOS_QR
    USUARIO% = USNBR% Mod 100
    PFAC$ = Control$(IDENTER$, "PORTSTIK")
    If PFAC$ = "" Then PFAC$ = Control$("*", "PORTSTIK")
    If PFAC$ = "" Then Exit Sub ' POR SI NOHAY NINGUNA IMPRESORA CONFIGURADA
    Call SETPRINPORT(PFAC$)
    
    Dim Rst As ADODB.Recordset
   
    ' *** IMPRIME TODAS LAS ETIQUETAS DE LA RECEPCION
    ' ***********************************************
    If NroRecepcion > 0 Then
        NumeroRecepcion$ = CStr(NroRecepcion)
        GoTo 65
    End If
    
    
    ' *** REIMPRIME UNA ETIQUETA
    ' **************************
    ' LISTA DE RECEPCIONES
40  TituloLista$ = "Recepciones de materiales"
    If TRIM$(FRMZELECHO.Caption) = TituloLista$ Then GoTo 50
    
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Recepcion/A12;Fecha/D8;Remito/A20;Proveedor/A32", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Recepcion/A12;Fecha/D8;Remito/A20;Proveedor/A32", FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim obj As New RECORXET
    Set rs1 = obj.RS_BOLRECEPEN()
    Call Carga_MSF_Recordset(rs1, "Recepcion/A12;Fecha/D8;Remito/A20;Proveedor/A32", FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 10000
50  FRMZELECHO.Show 1
    
    NumeroRecepcion$ = TRIM(RESP_ZELE_VECT(1))
    If NumeroRecepcion$ = "" Then Exit Sub
    
    ' LISTA DE ITEMS DE LA RECEPCION ELEGIDA
    TituloLista$ = "Items de la Recepción '" & NumeroRecepcion$ & "'"
    If TRIM$(FRMZELECHO.Caption) = TituloLista$ Then GoTo 50
    
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Lote/A12;Código/A16;Descripción/A32;Cantidad/F10.1;Peso/F10.2", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Lote/A12;Código/A16;Descripción/A32;Cantidad/F10.1;Peso/F10.2", FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim rst_items As ADODB.Recordset
    ConsultaSql$ = "select idenlote, cod_exte, descripitem, m_largo, peso, cantnom from bolrecep with(nolock) where numebore=" & NumeroRecepcion$
    Set rst_items = READSET(ConsultaSql$)
    
    Fila& = 0
    With rst_items
        Do While Not .EOF
            Fila& = Fila& + 1
            FRMZELECHO.msf2.Rows = Fila& + 1
            
            FRMZELECHO.msf2.TextMatrix(Fila&, 1) = !idenlote
            FRMZELECHO.msf2.TextMatrix(Fila&, 2) = !Cod_Exte
            FRMZELECHO.msf2.TextMatrix(Fila&, 3) = !descripitem
            cantidadRecibida# = XVALO(!M_Largo): If cantidadRecibida# < 0.05 Then cantidadRecibida# = XVALO(!CantNom)
            FRMZELECHO.msf2.TextMatrix(Fila&, 4) = TRIM$(FORMATNUM$(cantidadRecibida#, "F12.1"))
            FRMZELECHO.msf2.TextMatrix(Fila&, 5) = TRIM$(FORMATNUM$(!Peso, "F12.2"))
        
            .MoveNext
        Loop
    End With
            
    On Error Resume Next
    rst_items.Close
    Set rst_items = Nothing
    On Error GoTo 0
    
    If Fila& = 0 Then Exit Sub
    
    FRMZELECHO.Width = 10500
60  FRMZELECHO.Show 1
    
    LoteRecepcion$ = TRIM(RESP_ZELE_VECT(1))
    If LoteRecepcion$ = "" Then GoTo 40

    ' OBTIENE DATOS DEL LOTE
    ConsultaSql$ = "select fechrecep, cod_inte, cod_exte, descripitem, idenlote, numrem_prov, m_largo, cantnom, peso, raso_prov, NumRem_Prov, CantidadBultos,LoteCol_Prov"
    ConsultaSql$ = ConsultaSql$ + " from bolrecep with(nolock) where idenlote='" & LoteRecepcion$ & "'"
    Set Rst = READSET(ConsultaSql$)
    GoTo 70
        
    ' OBTIENE DATOS DE LA RECEPCION
65  ConsultaSql$ = "select fechrecep, cod_inte, cod_exte, descripitem, idenlote, numrem_prov, m_largo, cantnom, peso, raso_prov, NumRem_Prov, CantidadBultos,LOTECOL_PROV from bolrecep with(nolock) where numebore='" & CStr(NroRecepcion) & "'"
    Set Rst = READSET(ConsultaSql$)
    
70  With Rst
        Do While Not .EOF
            CodInte% = !cod_inte
            If InStr(1, EPAC$, "CAMBACERES") > 0 And TIMATE%(CodInte%) = 41 Then GoTo 80 'ACA VIENE AL FINAL de la rece.provisoria
            Codigo$ = CODEXT$(CodInte%)                                                  'si es bobina no imprime por que ya imprimio etiqueta
            Descripcion$ = DESCRIT0$(CodInte%)
            LOTE$ = SAFETEXT$(!idenlote)
            LoteProveedor$ = ""
            If InStr(1, EPAC$, "CONTEPLAST") > 0 Then
              LoteProveedor$ = SAFETEXT$(!LOTECOL_PROV) 'ES EL DE CAMBACERES / SI ES DE OTRO LO VA A DEVOLVER VACIO
            End If
            Largo$ = "0": If XVALO(!M_Largo) > 0.0005 Then Largo$ = TRIM$(FORMATNUM$(!M_Largo, "F12.1"))
            Peso$ = "0": If XVALO(!Peso) > 0.005 Then Peso$ = TRIM$(FORMATNUM$(!Peso, "F12.1"))
            Proveedor$ = !Raso_Prov
            Remito$ = SAFETEXT$(!NumRem_Prov)
            Remito$ = TRIM$(Left$(Remito$, 18)) & " " & LoteProveedor$ ' SI LO TIENE LO CONCATENA SI NO VA CONCATENA VACIO
            cantidadNominal# = !CantNom
            FechaRecepcion% = CVINT(!FechRecep)
            cantidadBultos% = XVALO(!cantidadBultos)
               
            TEXTOQR$ = Codigo$ & "|" & LOTE$
            IMAG$ = "IMAGRECE_QR.JPG"
            IDENTIFICADOR$ = "QR_RECEPCION"
            ANCHOX% = 0
            ALTOX% = 0
            '
            If IncluyeQR = True Then
                On Error Resume Next
                Kill LUCOM$ & IMAG$
                On Error GoTo 0
                'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
                QRX$ = TRIM$(IMAGENCODBARRA$(LUCOM$, IMAG$, TEXTOQR$, "QR_CODE", ANCHOX%, ALTOX%))
                Call T_Espera(3)
                Image1.Picture = LoadPicture(LUCOM$ & IMAG$)
                Call GRABAR_QR(LUCOM$ & IMAG$, IDENTIFICADOR$)
            End If
            If InStr(1, EPAC$, "CONTEPLAST") < 1 Then 'si NO es conteplast
                If cantidadBultos% > 0 Then
                    For i = 1 To cantidadBultos%
                        numeroBulto$ = CStr(i) & " / " & CStr(cantidadBultos%)
        
                        FILTX$ = LOTE$ & ";" & Codigo$ & ";" & Descripcion$ & ";" & Proveedor$ & ";" & Largo$ & ";" & Peso$ & ";" & IDENTIFICADOR$ & ";" & CStr(USUARIO%) & ";" & Remito$ & ";" & TRIM$(FORMATNUM$(cantidadNominal#, "F12.1")) & ";" & FECHATEX$(FechaRecepcion%) & ";Recepcion de material" & ";" & numeroBulto$
                        Call STDFORM(FORIPRE$, FILTX$, "PRINT")
                    Next i
                Else
                    FILTX$ = LOTE$ & ";" & Codigo$ & ";" & Descripcion$ & ";" & Proveedor$ & ";" & Largo$ & ";" & Peso$ & ";" & IDENTIFICADOR$ & ";" & CStr(USUARIO%) & ";" & Remito$ & ";" & TRIM$(FORMATNUM$(cantidadNominal#, "F12.1")) & ";" & FECHATEX$(FechaRecepcion%) & ";Recepcion de material" & ";" & numeroBulto$
                    Call STDFORM(FORIPRE$, FILTX$)
                End If
             Else  'si es CONTEPLAST
                    FILTX$ = LOTE$ & ";" & Codigo$ & ";" & Descripcion$ & ";" & Proveedor$ & ";" & Largo$ & ";" & Peso$ & ";" & Remito$
                    Call STDFORM(FORIPRE$, FILTX$)
             
             End If
             
80          .MoveNext
        Loop
    End With
        
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
    If PFAC$ <> "" Then Call SETPRINPORT("RESTORE")

End Sub

Sub ReImprimirRD()
     
     Screen.MousePointer = 11
     CMO$ = TRIM$(Control$("", "CODRRECH"))
     If Len(CMO$) <> 2 Then
        CMO$ = "RD"
     End If
     '
     Call ABRESTOCKS
     '
     CONDI$ = "substring(REFERCOR,1,4) <> 'ANUL' " 'PARA EVITAR COMPROBANTES ARRUINADOS
     CONDI$ = CONDI$ + " AND CodiMovi = '" & CMO$ & "' AND CODIEMPR = " & CodiEmp% & " ORDER BY FECHMOV DESC, DOPRILAR DESC"
     '
     Call COPYSTRU("SELDOCAP", "SELREMIT")
     Call CARGALISEL("!SELREMIT", "MOVISTO", "DOPRILAR/A11;FechMov/D8;REFERCOR/A24", CONDI$, "DISTINCT")
     '
15   Call SELECHO("!SELREMIT/Remitos de Devolución")
     '
     If ULTREG&("!SELREMIT") < 1 Then
        Call MENSERR(24, "No se Encontraron Remitos en el Rango Seleccionado.")
        GoTo 99
     End If
     '
     Screen.MousePointer = 1
     '
     NUDOCU$ = TRIM$(VALACT1$("!SELREMIT"))
     If VALACT1$("!SELREMIT") = "" Then GoTo 99
     TIDOCUM$ = "Remito de Devolución"
     '
     Screen.MousePointer = 11
      '
     For UI& = ULTREG&("PDOCLST") To 1 Step -1
        XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
        '\\\OT-08-0646-OD-20/08/08///
         If InStr(XX$, TIDOCUM$) > 4 Then
            AA = InStr(XX$, NUDOCU$)
            If InStr(XX$, NUDOCU$) > 4 Then
                PPXX% = InStr(XX$, "Nro.")
                If PPXX% > 0 Then
                 XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
                End If
                If InStr(XX$, NUDOCU$) > 4 Then
                 DOCUNU& = CVS(Left$(XX$, 4))
                 GoTo 50
                End If
            End If
         End If
     Next UI&
     Screen.MousePointer = 1
40   Call COMUNI("No se Encontró el Documento Impreso")
     GoTo 15
     '
50   Call MUESTRADOC(DOCUNU&)
     GoTo 15
     '
99   Command17.Enabled = True
     '
End Sub

Function ValidayoAplicaOrdenCompra%()
        'VALIDA SI TIENE O.C. SI TIENE UNA AL RESTO LE AGREGA LA MISMA O.C
        ValidayoAplicaOrdenCompra% = 1
        For J& = 1 To VENTABNEW.MSF.Rows - 1
           ORDOC1& = XVALO(VENTABNEW.MSF.TextMatrix(J&, 5))
           If ORDENITEM% < 1 Then
             ORDENITEM% = XVALO(VENTABNEW.MSF.TextMatrix(J&, 7))
             UNIM$ = VENTABNEW.MSF.TextMatrix(J&, 4)
           End If
           CI% = CODINT%(VENTABNEW.MSF.TextMatrix(J&, 2))
           If CI_ANT% > 0 And CI% <> CI_ANT% Then ' SI LOS CODIGOS SON DIFERERNTES NO VALIDA EL TEMA DE LA O.C.
              ValidayoAplicaOrdenCompra% = 1 'CONTINUA SIN VAIDACION
           End If
           If ORDOC1& > 0 Then
              OrdenEncontrada& = ORDOC1&
           Else
              FaltaOrdenCompra% = 1
           End If
           CI_ANT% = CI%
        Next J&
        If OrdenEncontrada& > 0 And FaltaOrdenCompra% > 0 Then 'tiene que tener o.c. alguna y alguna no si no no valida
               OPX% = COMALTER%("Se Detectaron Filas sin O.Compra\\Cancelar\Aplicar O.C." & OrdenEncontrada& & " A Todas")
               If OPX% < 2 Then 'VUELVE A LA VENTANA
                 ValidayoAplicaOrdenCompra% = -1
                 Exit Function 'se va
               End If
               
               For J& = 1 To VENTABNEW.MSF.Rows - 1
                  VENTABNEW.MSF.TextMatrix(J&, 5) = OrdenEncontrada&
                  VENTABNEW.MSF.TextMatrix(J&, 7) = ORDENITEM% 'TODOS CON EL MISMO ORDEN DE ITEM
                  VENTABNEW.MSF.TextMatrix(J&, 4) = UNIM$
               Next J&
               ValidayoAplicaOrdenCompra% = 1
        End If
End Function

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command10_Click()
    Command10.Enabled = False
    If EXISTE%("CALPRO07.EXE") > 0 Then
        Call CONECRUN("CALPRO07", "Sistema Calificación de Proveedores")
    ElseIf EXISTE%("CALPRO09.EXE") > 0 Then
        Call CONECRUN("CALPRO09", "Sistema Calificación de Proveedores")
    End If
    Command10.Enabled = True
End Sub

Private Sub Command11_Click()
    '
    If EsReteplast% = 1 Then
        Call DESHASFECHA(Desde%, Hasta%, PERIODO$)
        If PERIODO$ <> "" Then
            Campos$ = "Fecha/D8;Recepción/I10;Cta/I10;Proveedor/A24;Remito/A12;Código/A24;Descripción/A32;Lote/A20;Kg Control/F10.1;Kg Remito/F10.1;Kg Balanza/F10.1"
            Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
            Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
            GRILLAGRAL.Caption = "RECEPCIONES POR RANGO DE FECHAS | " & PERIODO$
            
            DesdeSql$ = FETEXCOR1$(Desde%)
            HastaSql$ = FETEXCOR1$(Hasta%)
            
            Dim Rst As ADODB.Recordset
            ConsultaSql$ = "select fechrecep, numebore, numprov, raso_prov, numrem_prov, cod_exte, descripitem, lotecol_prov, cantnom, peso, kgbalanzapublica from bolrecep with(nolock)" _
            & " where fechrecep>='" & DesdeSql$ & "' and fechrecep<='" & HastaSql$ & "'" _
            & " order by fechrecep, numebore"
            Set Rst = READSET(ConsultaSql$)
            
            Call Carga_MSF_Recordset(Rst, Campos$, GRILLAGRAL.GRID)
                        
            GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
            GRILLAGRAL.mnuMenu1.Visible = True
            GRILLAGRAL.MnuEdición.Visible = False
            GRILLAGRAL.GRID.ZOrder 0
            GRILLAGRAL.TXTBUSCAR = ""
            
            GRILLAGRAL.Show 1
            
            Dim MSF As msFlexGrid
            If GRILLAGRAL.MENU1 > 0 Then
                Set MSF = GRILLAGRAL.GRID
                If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
                Call CARGA_TABLA_IMPRE(MSF, Campos$, "RECEFECH", GRILLAGRAL.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1)
                Call FINAL("?" & RET_NAMECONS$)
                GRILLAGRAL.MENU1 = 0
                Screen.MousePointer = 1
            End If
        End If
        
        Exit Sub
    End If
    
    If Option1.Value = True Then
       Call MENSERR(24, "Opcion no Disponible por Pantalla.\  \Emita Listado Impreso y Exporte a Archivo de Texto.")
       GoTo 20
    End If
    '
20  HOI% = HOY(HO$)
    FE$ = "01" + Mid$(HO$, 3)
    FF1% = FECHANUM(FE$): FF2% = HOI%
    '
    LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
    If LONCODI% < 16 Or LONCODI% > 24 Then LONCODI% = 16

    XX$ = Left$(REGLEIDO$("ECOMAST", 1), LONCODI%): CI1% = Abs(CODINT%(XX$))
    XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), LONCODI%): CI2% = Abs(CODINT%(XX$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0) + MKI$(FF1%) + MKI$(FF2%)
    '
22  RACOD$ = OBJPLA$("DESHASCOD2", VDEF$)
    If RACOD$ = "" Then
        GoTo 99
    End If
    CI1% = CVI(Left$(RACOD$, 2)): COD1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): COD2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TMT% = 0 Or TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(RACOD$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación"
    End If
    FF1% = CVI(Mid$(RACOD$, 6, 2))
    FEX = FF1%: FLI1$ = FECHATEX$(FEX)
    FF2% = CVI(Mid$(RACOD$, 8, 2))
    FEX = FF2%: FLI2$ = FECHATEX$(FEX)
    '
    If COD1$ = "" Or COD2$ = "" Or COD2$ < COD1$ Or FF2% < FF1% Then
        VDEF$ = RACOD$
        GoTo 22
    End If
    '
    Rango$ = TRIM$(COD1$) + " - " + TRIM$(COD2$)
    Rango$ = Rango$ + " / " + TRIM$(TIMAT$)
    Rango$ = Rango$ + " / " + FLI1$ + " - " + FLI2$
    '
    Screen.MousePointer = 11
    '
    Call COPYSTRU("BOLRECEP", "BOLSELEC")
    J& = 0
    '
    Call ABRESTOCKS
    'FF1D = FECHATEX$(FF1%)
    'FF2H = FECHATEX$(FF2%)
    FF1D = FETEXCOR1$(FF1%)
    FF2H = FETEXCOR1$(FF2%)
    '
    SQLX$ = " SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + "INNER JOIN MASTER ON Master.Codint = BolRecep.Cod_Inte "
    SQLX$ = SQLX$ + " WHERE FechRecep >= '" & FF1D & "' "
    SQLX$ = SQLX$ + " AND FechRecep <= '" & FF2H & "' "
    SQLX$ = SQLX$ + " AND Master.Codigo >= '" & COD1$ & "' "
    SQLX$ = SQLX$ + " AND Master.Codigo <= '" & COD2$ & "' "
    SQLX$ = SQLX$ + " ORDER BY NumeBoRe "
    '
    Dim BRECEPRATMP As ADODB.Recordset
    Set BRECEPRATMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With BRECEPRATMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Boletas de Recepcion\que Cumplan Condición de Filtro.")
        On Error GoTo 0
        Screen.MousePointer = 1
        GoTo 99
      End If
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
         CIXI% = !cod_inte
         If TMT% = 0 Or TIMATE%(CIXI%) = TMT% Then
            Y$ = REGBLAN("BOLSELEC")
            Call REPLA(Y$, MKS$(!NUMEBORE), 1, 4)
            Call REPLA(Y$, !Descrip, 5, 64)
            Call REPLA(Y$, !idenlote, 69, 12)
            Call REPLA(Y$, MKI$(CVINT(!FechRecep)), 81, 2)
            Call REPLA(Y$, MKI$(!cod_inte), 83, 2)
            Call REPLA(Y$, !Unidad, 85, 4)
            Call REPLA(Y$, MKD$(!CantNom), 89, 8)
            Call REPLA(Y$, Chr$(!DEPOENT), 97, 1)
            Call REPLA(Y$, MKI$(!NumProv), 105, 2)
            Call REPLA(Y$, !LOTECOL_PROV, 107, 12)
            Call REPLA(Y$, !NumRem_Prov, 119, 16)
              NCX% = !NumProv
            Call REPLA(Y$, CODICLI$((-1) * NCX%), 135, 16)
            Call REPLA(Y$, MKS$(!NUMOC), 151, 4)
            Call REPLA(Y$, !MoneVaUni, 155, 4)
            Call REPLA(Y$, MKS$(!Valo_Unit), 159, 4)
            Call REPLA(Y$, MKD$(!CanToApro), 163, 8)
            Call REPLA(Y$, MKD$(!CanToRech), 171, 8)
            Call REPLA(Y$, MKD$(!CanToFalta), 179, 8)
            Call REPLA(Y$, MKI$(CVINT(!FechLibe)), 187, 2)
            If IsNull(!Observa) Then
                USERAPRO$ = " "
            Else
                USERAPRO$ = !Observa
            End If
            Call REPLA(Y$, USERAPRO$, 193, 48) 'antes estaba el campo observaciones
            '
            JJ& = JJ& + 1
            Call GRAREG("BOLSELEC", Y$, JJ&)
         End If
      .MoveNext
      Loop
    End With
    '
98  BRECEPRATMP.Close
    Set BRECEPRATMP = Nothing
    '
    Call SETULTREG("BOLSELEC", JJ&)
    Call CIERRARCH("BOLSELEC")
    Call CIERRARCH("*.*")
    '
    Call HEADERS("BOLSELEC", "")
    Call HEADERS("BOLSELEC", "Rango: " + Rango$)
    Call CONECRUN("*BOLSELEC", "Recepciones por Rango de Fechas")
    '
99  Screen.MousePointer = 1
    '
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   '
   Call GENINDIBOL(0)
   '
3  Call SELECHO("INDIBOL")
   NPX& = Val(VALACT1$("INDIBOL"))
   If NPX& > 0 Then
     Call SHOWBORE(NPX&)
     GoTo 3
   End If
   '
99 Command12.Enabled = True
End Sub

Private Sub Command13_Click()
      Command13.Enabled = False
      Screen.MousePointer = 11
      'agrego todas las O/C de ocomenca
      Call ABRECOMPRAS
      JJ& = 0
      SQLX$ = " SELECT Nume_Ocom, Refe_Ocom FROM OCOMENCA WITH(NOLOCK)"
      SQLX$ = SQLX$ + " ORDER BY Nume_Ocom  DESC "
      '
      Dim VEROCTEMP As ADODB.Recordset
'      Set VEROCTEMP = New ADODB.Recordset 'MODIFICADO PARA MEJORAR VELOCIDAD OD-30/11/11
'      VEROCTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'      Set VEROCTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      Set VEROCTEMP = READSET(SQLX$) 'OD-02/12/11 ACTUALIZADO A LA FUNCION READSET
      With VEROCTEMP
         Do While Not .EOF
             XXY$ = REGBLAN$("INDIOCOM")
             Call REPLA(XXY$, MKS(XVALO(!nume_ocom)), 1, 4)
             Call REPLA(XXY$, SAFETEXT$(!Refe_Ocom), 5, 44)
             JJ& = JJ& + 1
             Call GRAREG("INDIOCOM", XXY$, JJ&)
             .MoveNext
         Loop
      End With
      VEROCTEMP.Close
      '
      ' LO QUE SIGUE PARA ORDENES DE COMPRA ABIERTAS
10    SQLX$ = " SELECT Nume_Ocom, Refe_Ocom FROM ENCOCABI WITH(NOLOCK) "
      SQLX$ = SQLX$ + " ORDER BY Nume_Ocom  DESC "
      '
'      Set VEROCTEMP = New ADODB.Recordset 'MODIFICADO PARA MEJORAR VELOCIDAD OD-30/11/11
'      VEROCTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'      Set VEROCTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      Set VEROCTEMP = READSET(SQLX$) 'OD-02/12/11 ACTUALIZADO A LA FUNCION READSET
      
      With VEROCTEMP
         Do While Not .EOF
             XXY$ = REGBLAN$("INDIOCOM")
             Call REPLA(XXY$, MKS(XVALO(!nume_ocom)), 1, 4)
             Call REPLA(XXY$, "(abierta) " + SAFETEXT$(!Refe_Ocom), 5, 44)
             JJ& = JJ& + 1
             Call GRAREG("INDIOCOM", XXY$, JJ&)
             .MoveNext
         Loop
      End With
      VEROCTEMP.Close
      Set VEROCTEMP = Nothing
      '
20    Call SETULTREG("INDIOCOM", JJ&)
      Call CIERRARCH("INDIOCOM")
      If JJ& < 1 Then
         Screen.MousePointer = 1
         Call MENSERR(24, "No hay Ordenes de Compra Emitidas")
         GoTo 99
      End If
      '
      Screen.MousePointer = 1
3     Call SELECHO("INDIOCOM/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
      
      NPX& = Val(VALACT1$("INDIOCOM"))
      If Left$(VALACT2$("INDIOCOM"), 9) = "(abierta)" Then
         NPX& = (-1) * Abs(NPX&)
      End If
      '
      If Abs(NPX&) > 0 Then
        Call SHOWOCOM(NPX&)
        GoTo 3
      End If
      '
99  Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   '
   If DERACCE%("STPRECEP", "Configuración Procesos de Recepción") >= 2 Then
     OPRECEP07.Show 1
   End If
   '
End Sub

Private Sub Command15_Click()
    '
    Has% = HOY(HOS$)
    FEX = Has%
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
10  OKKX% = 0
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
       Exit Sub
    End If
    OKKX% = 1
    Des% = CVI(Left$(OBX$, 2)): DES1# = CDbl(CVDAT(Des%))
    Has% = CVI(Mid$(OBX$, 3, 2)): HAS1# = CDbl(CVDAT(Has%))
    If Has% < Des% Then GoTo 10
    FEX = Des%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    
    FF1D = FECHATEX$(Des%)
    FF2H = FECHATEX$(Has%)
 
    
    '
    Screen.MousePointer = 11
    '
    'SQLX$ = "SELECT * FROM BOLRECEP "
    'SQLX$ = SQLX$ + " WHERE (isnull(NumOC)=true or NumOC < 1) "
    'SQLX$ = SQLX$ + " AND INT(CDBL(FechRecep)) >= " & DES1# & " "
    'SQLX$ = SQLX$ + " AND INT(CDBL(FechRecep)) <= " & HAS1# & " "
    'SQLX$ = SQLX$ + " ORDER BY NumeBoRe ASC "
    
    SQLX$ = "SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + " WHERE NumOC is null or NumOC < 1 "
    SQLX$ = SQLX$ + " AND FechRecep >= '" & FF1D & "'"
    SQLX$ = SQLX$ + " AND FechRecep <= '" & FF2H & "'"
    SQLX$ = SQLX$ + " ORDER BY NumeBoRe ASC "
    Call ABRESTOCKS
    Dim RECEPTMP As ADODB.Recordset
    Set RECEPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    
    'Set RECEPTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    With RECEPTMP
    '
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Call COMUNI("No Existen Materiales Recibidos Sin O/C\en Rango de Fechas Elegido")
        Screen.MousePointer = 1
        Exit Sub
      End If
      JJ& = 0
      Call COPYSTRU("BOLSELEC", "BOLSELE1")
      Do While Not .EOF
        Y$ = REGBLAN("BOLSELE1")
        Call REPLA(Y$, MKS$(!NUMEBORE), 1, 4)
        Call REPLA(Y$, !Descrip, 5, 64)
        Call REPLA(Y$, !idenlote, 69, 12)
        Call REPLA(Y$, MKI$(CVINT(!FechRecep)), 81, 2)
        Call REPLA(Y$, MKI$(!cod_inte), 83, 2)
        Call REPLA(Y$, !Unidad, 85, 4)
        Call REPLA(Y$, MKD$(!CantNom), 89, 8)
        Call REPLA(Y$, Chr$(!DEPOENT), 97, 1)
        Call REPLA(Y$, MKI$(!NumProv), 105, 2)
        Call REPLA(Y$, !LOTECOL_PROV, 107, 12)
        Call REPLA(Y$, !NumRem_Prov, 119, 16)
        Call REPLA(Y$, !NumFac_Prov, 135, 16)
        Call REPLA(Y$, MKS$(!NUMOC), 151, 4)
        Call REPLA(Y$, !MoneVaUni, 155, 4)
        Call REPLA(Y$, MKS$(!Valo_Unit), 159, 4)
        Call REPLA(Y$, MKD$(!CanToApro), 163, 8)
        Call REPLA(Y$, MKD$(!CanToRech), 171, 8)
        Call REPLA(Y$, MKD$(!CanToFalta), 179, 8)
        Call REPLA(Y$, MKI$(CVINT(!FechLibe)), 187, 2)
        Call REPLA(Y$, !Observa, 193, 64)
        '
        JJ& = JJ& + 1
        Call GRAREG("BOLSELE1", Y$, JJ&)
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("BOLSELE1", JJ&)
    Call CIERRARCH("BOLSELE1")
    Call CIERRARCH("*.*")
    '
    Call HEADERS("BOLSELE1", "")
    Call HEADERS("BOLSELE1", PERIO$)
    Call FILESORT("BOLSELE1", "", 1, 3, "ASC")
    Call CONECRUN("*BOLSELE1", "Recepciones sin O/C's por Rango de Fechas")
    '
    Screen.MousePointer = 1
    '
    RECEPTMP.Close
    Set RECEPTMP = Nothing
    
99 End Sub

Private Sub Command16_Click()
    '
    Has% = HOY(HOS$)
    FEX = Has%
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
10  OKKX% = 0
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
       Exit Sub
    End If
    OKKX% = 1
    Des% = CVI(Left$(OBX$, 2)): DES1# = CDbl(CVDAT(Des%))
    Has% = CVI(Mid$(OBX$, 3, 2)): HAS1# = CDbl(CVDAT(Has%))
    If Has% < Des% Then GoTo 10
    FEX = Des%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
   
    FF1D = FECHATEX$(Des%)
    FF2H = FECHATEX$(Has%)
   
   
    '
    Screen.MousePointer = 11
    '
    SQLX$ = "SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + " WHERE CanToRech > 0 "
    SQLX$ = SQLX$ + " AND FechRecep >= '" & FF1D & "' "
    SQLX$ = SQLX$ + " AND FechRecep <= '" & FF2H & "' "
    SQLX$ = SQLX$ + " ORDER BY NumeBore ASC "
    Call ABRESTOCKS
    'Set RECEPTMP = Stocks.OpenRecordset(sqlx$, dbOpenSnapshot)
    Dim RECEPTMP As ADODB.Recordset
    Set RECEPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With RECEPTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Call COMUNI("No Existen Recepciones con Rechazos\en Rango de Fechas Elegido.")
        Screen.MousePointer = 1
        Exit Sub
      End If
      JJ& = 0
      Call COPYSTRU("BOLSELEC", "BOLSELER")
      Do While Not .EOF
        Y$ = REGBLAN("BOLSELER")
        Call REPLA(Y$, MKS$(!NUMEBORE), 1, 4)
        Call REPLA(Y$, !Descrip, 5, 64)
        Call REPLA(Y$, !idenlote, 69, 12)
        Call REPLA(Y$, MKI$(CVINT(!FechRecep)), 81, 2)
        Call REPLA(Y$, MKI$(!cod_inte), 83, 2)
        Call REPLA(Y$, !Unidad, 85, 4)
        Call REPLA(Y$, MKD$(!CantNom), 89, 8)
        Call REPLA(Y$, Chr$(!DEPOENT), 97, 1)
        Call REPLA(Y$, MKI$(!NumProv), 105, 2)
        Call REPLA(Y$, !LOTECOL_PROV, 107, 12)
        Call REPLA(Y$, !NumRem_Prov, 119, 16)
        Call REPLA(Y$, !NumFac_Prov, 135, 16)
        Call REPLA(Y$, MKS$(!NUMOC), 151, 4)
        Call REPLA(Y$, !MoneVaUni, 155, 4)
        Call REPLA(Y$, MKS$(!Valo_Unit), 159, 4)
        Call REPLA(Y$, MKD$(!CanToApro), 163, 8)
        Call REPLA(Y$, MKD$(!CanToRech), 171, 8)
        Call REPLA(Y$, MKD$(!CanToFalta), 179, 8)
        Call REPLA(Y$, MKI$(CVINT(!FechLibe)), 187, 2)
        Call REPLA(Y$, !Observa, 193, 64)
        '
        JJ& = JJ& + 1
        Call GRAREG("BOLSELER", Y$, JJ&)
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("BOLSELER", JJ&)
    Call CIERRARCH("BOLSELER")
    Call CIERRARCH("*.*")
    '
    Call HEADERS("BOLSELER", "")
    Call HEADERS("BOLSELER", PERIO$)
    Call FILESORT("BOLSELER", "", 1, 3, "ASC")
    Call CONECRUN("*BOLSELER", "Recepciones con Rechazos por Rango de Fechas")
    '
    Screen.MousePointer = 1
    '
    RECEPTMP.Close
    Set RECEPTMP = Nothing
    
End Sub

Private Sub Command17_Click()
    Command17.Enabled = False
    If DERACCE%("TRAZABR", "Registro de Trazabilidad de B/R's") = 2 Then
      ACONEC$ = "TRAZAB07"
      If EXISTE%("TRAZAB07.EXE") < 1 Then
         ACONEC$ = "TRAZAB08"
         If EXISTE%("TRAZAB08.EXE") < 1 Then
            ACONEC$ = "TRANSA07"
         End If
      End If
      Call CONECRUN(ACONEC$, "Sub-Sistema de Trazabilidad")
    End If
    Command17.Enabled = True
End Sub

Private Sub Command18_Click()
    '
    Command18.Enabled = False
    Has% = HOY(HOS$)
    FEX = Has%
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
10  OKKX% = 0
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
       GoTo 99 'Exit Sub X EL SCREENMOUSEPOINTER
    End If
    OKKX% = 1
    Des% = CVI(Left$(OBX$, 2)): DES1# = CDbl(CVDAT(Des%))
    Has% = CVI(Mid$(OBX$, 3, 2)): HAS1# = CDbl(CVDAT(Has%))
    If Has% < Des% Then GoTo 10
    FEX = Des%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
    FEX = Has%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    '
    FF1D = FECHATEX$(Des%)
    FF2H = FECHATEX$(Has%)
    
    SQLX$ = "SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + " WHERE CanToRech > 0 "
    SQLX$ = SQLX$ + " AND FechRecep >= '" & FF1D & "' "
    SQLX$ = SQLX$ + " AND FechRecep <= '" & FF2H & "' "
    SQLX$ = SQLX$ + " ORDER BY NumeBore ASC "
    Call ABRESTOCKS
    '
    Dim RECEPTMP As ADODB.Recordset
    Set RECEPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With RECEPTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Call COMUNI("No Existen Recepciones con Rechazos\en Rango de Fechas Elegido.")
        GoTo 99
      End If
      JJ& = 0
      Call COPYSTRU("BOLSELEC", "BOLSELER")
      Do While Not .EOF
        Y$ = REGBLAN("BOLSELER")
        Call REPLA(Y$, MKS$(!NUMEBORE), 1, 4)
        Call REPLA(Y$, !Descrip, 5, 64)
        Call REPLA(Y$, !idenlote, 69, 12)
        Call REPLA(Y$, MKI$(CVINT(!FechRecep)), 81, 2)
        Call REPLA(Y$, MKI$(!cod_inte), 83, 2)
        Call REPLA(Y$, !Unidad, 85, 4)
        Call REPLA(Y$, MKD$(!CantNom), 89, 8)
        Call REPLA(Y$, Chr$(!DEPOENT), 97, 1)
        Call REPLA(Y$, MKI$(!NumProv), 105, 2)
        Call REPLA(Y$, !LOTECOL_PROV, 107, 12)
        Call REPLA(Y$, !NumRem_Prov, 119, 16)
        Call REPLA(Y$, !NumFac_Prov, 135, 16)
        Call REPLA(Y$, MKS$(!NUMOC), 151, 4)
        Call REPLA(Y$, MKD$(!CanToRech), 171, 8)
        '
        JJ& = JJ& + 1
        Call GRAREG("BOLSELER", Y$, JJ&)
      .MoveNext
      Loop
    End With
    RECEPTMP.Close
    Set RECEPTMP = Nothing
    '
    Call SETULTREG("BOLSELER", JJ&)
    Call CIERRARCH("BOLSELER")
    Call CIERRARCH("*.*")
    '
    Call HEADERS("BOLSELER", "")
    Call HEADERS("BOLSELER", PERIO$)
    Call FILESORT("BOLSELER", "", 5, 5, "ASC")
    Call CONECRUN("*BOLSELEX", "Rechazos por Rango de Fechas")
    '
99  Screen.MousePointer = 1
    Command18.Enabled = True
End Sub




Private Sub Command19_Click()
    
    Call mnuClasificacionMaterialesRecibidos_Click
    
End Sub

Private Sub Command2_Click()
    RTA% = 1   ' COMALTER%("Modalidad de Procesamiento:\\Aprobación/Rechazo\Discrepancia/Faltantes")
    Select Case RTA%
      Case 1
        Call APROBOL
      Case 2
        Call DISCREBOL
    End Select
End Sub
Sub PROCESO_ASIGNACION_DE_LOTES_A_CONSUMOS(CI%, LOTE_RECIBIDO$, CANTIDAD_INGRESADA)
   If InStr(UCase$(EMPREAC$), "AHP") > 0 Then
      FIN& = canticampos("Movisto")
      Dim VALRS(256)
      'REALIZAR CONSULTA SOBRE MOVISTO VERIFICANDO SI HAY MOVIMIENTOS DE SALIDA CON CODIMOVI = 'RT' SIN LOTES
      'EN CASO AFIRMATIVO ASIGNAR LOTES A LOS REGISTROS
      'SI ES POR EL TOTAL ASIGNA EL LOTE Y VA DESCONTANDO LA CANTIDAD PARA UNO  O MAS REGISTROS
      'SI NO ALCANZA VA APLICANDO DE  MANERA PARCIAL
      'SE DEBERIA CONFIGURAR LA FECHA DE INICIO DE SALIDAS.
      SQLX$ = "SELECT * FROM MOVISTO "
      SQLX$ = SQLX$ + " WHERE CODIMOVI = 'RT' "
      SQLX$ = SQLX$ + " AND (IDENLOTE = '' OR IDENLOTE IS NULL)"
      'SQLX$ = SQLX$ + " AND CANTSALID > 0 "
      SQLX$ = SQLX$ + " AND COD_INTE = " & CI%
      SQLX$ = SQLX$ + " ORDER BY FECHMOV"
      Dim RS As ADODB.Recordset
      Set RS = New ADODB.Recordset
25    On Error Resume Next
      RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
      End If
      On Error GoTo 0
      With RS
        If Not .EOF Then
           
        End If
          
        Do While Not .EOF
           CANTISAL = XVALO(!CANTSALID)
           'SI LA CANTIDAD INGRESADA ES MAYOR A LA CANTIDAD DEL MOVIMIENTO DIRECTAMENTE LE PONE EL LOTE EN EL CAMPO IDENLOTE
           If XVALO(CANTIDAD_INGRESADA) >= CANTISAL Then
              !idenlote = LOTE_RECIBIDO$
              .Update
           Else
             'SI LA CANTIDAD ES MENOR DEBERA APLICAR EL LOTE A UN REGISTRO CON DICHA CANTIDAD Y EL LOTE Y GENERAR UNO POR EL SALDO
             !CANTSALID = CANTISAL - CANTIDAD_INGRESADA 'SOBRE ESCRIBE LA CANTIDAD ORIGINAL CON LA CANT. QUE VA A QUEDAR SIN LOTE
             .Update
              For KKI% = 1 To FIN& ' RECORRE TODOS LOS REGISTROS Y LOS GUARDA EN UN VECTOR
               VALRS(KKI%) = RS.Fields(KKI% - 1).Value
              Next KKI%
              .AddNew   'CREA UN NUEVO REGISTRO
              On Error Resume Next
              !CODIEMPR = CodiEmp%
              On Error GoTo 0
              'GRABA TODA LA INFORMACION COPIADA EN EL NUEVO REGISTRO
              For KKI% = 1 To FIN&
               .Fields(KKI% - 1).Value = VALRS(KKI%)
              Next KKI%
              '
             !CANTSALID = CANTIDAD_INGRESADA 'SOBRE-ESCRIBE LA CANTIDAD SALIDA.
             !idenlote = LOTE_RECIBIDO$      'LE AGREGA EL LOTE  CORRESPONDIENTE.
             .Update
             End If
             CANTIDAD_INGRESADA = CANTIDAD_INGRESADA - CANTISAL
             If CANTIDAD_INGRESADA <= 0 Then Exit Do
             .MoveNext
        Loop
      End With
   End If
End Sub

Private Sub Command20_Click()
    Unload VENTABNEW
    On Error GoTo 0
    
    VENTANA.Inicializar = 0
    VENTANA.AgregaRegistro = 1
    VENTANA.UtilizaFormula = 1
    VENTANA.HabilitaBorrar = 0
    VENTANA.HabilitaInsertar = 1
    VENTANA.NoMostrarBotonTilde = 0
    VENTANA.TITULO = "RECEPCION DE BOBINAS"
    VENTANA.Campox = "Codigo/A16;Descripcion/A32;Entrega/F8;Metros/F10.1;Peso/F10.1;REMITO/A18"
    VENTANA.CampEditables = "1;3;4;5;6"
    VENTANA.ColEcoCampo(1) = "codigo;descripcion;master;status>0"

    VENTANA.Inicializar = 1
    VENTABNEW.Show
    
End Sub

Private Sub Command3_Click()
    If Option1.Value = True Then
        Call COBORECEP
        If ULTREG("BOLREPEN") < 1 Then GoTo 99
        Call FINAL("*COBOLPE")
      ElseIf Option2.Value = True Then
        Call COBORECEP
        If ULTREG("BOLREPEN") < 1 Then GoTo 99
        If HAYMATERIALIB% = 0 Then
            Call FINAL("*LIBOLPE")
        Else
            Call FINAL("*LIBOPEL")
        End If
    End If
    '
99  Screen.MousePointer = 1 '\\\OT-06-0179-WAR-12/05/06///
    '
End Sub

Private Sub Command4_Click()
    '
    Call Selcopen
3   Call SELECHO("INDIOCPE/Indice de Ordenes de Compra Pendientes")
    NPX& = Val(VALACT1$("INDIOCPE"))
    If NPX& > 0 Then
      Call SHOWOCOM(NPX&)
      GoTo 3
    End If
    '
End Sub

Private Sub Command5_Click()
    '
    Screen.MousePointer = 11
    Call ABRESTOCKS
    SQLX$ = "Select * from MACONSIG"
    SQLX$ = SQLX$ + " WHERE Cant_Rendid < Cant_Entreg - 0.001 "
    '
    Dim CONSIGTMP As ADODB.Recordset
    Set CONSIGTMP = New ADODB.Recordset
    CONSIGTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With CONSIGTMP
      '
      JJ& = 0
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
          zz$ = REGBLAN$("MACONSIG")
          Call REPLA(zz$, MKI$(CVINT(!Femi_Ocom)), 1, 2)
            NC% = !Npro_Ocom
          Call REPLA(zz$, MKI$(NC%), 3, 2)
            NORFA$ = !NRO_ORFA
          Call REPLA(zz$, NOCOM$, 9, 12)
            NOCOM$ = !nro_ocom
          Call REPLA(zz$, NOCOM$, 21, 12)
            CI% = !cod_inte
          Call REPLA(zz$, MKI$(CI%), 33, 2)
          Call REPLA(zz$, Unimed$(CI%), 51, 4)
            CAN# = !cant_asig
          Call REPLA(zz$, MKD$(CAN#), 55, 8)
            CAN# = !cant_entreg
          Call REPLA(zz$, MKD$(CAN#), 63, 8)
            FF% = CVINT%(!FECHA_ENTREG)
          Call REPLA(zz$, MKI$(FF%), 71, 2)
            CAN# = !CANT_RENDID
          Call REPLA(zz$, MKD$(CAN#), 111, 8)
            CAN# = !Saldo_Pend
          Call REPLA(zz$, MKD$(CAN#), 119, 8)
            '
          JJ& = JJ& + 1
          Call GRAREG("MACONSIG", zz$, JJ&)
        .MoveNext
        Loop
      End If
    End With
    CONSIGTMP.Close
    Set CONSIGTMP = Nothing
    Call SETULTREG("MACONSIG", JJ&)
    Screen.MousePointer = 1
    '
    If Option1.Value = True Then
        Call FINAL("*COMATCO")
      ElseIf Option2.Value = True Then
        Call FINAL("*LIMATCO")
    End If
    '
End Sub

Private Sub Command6_Click()
    Call RECPROP
End Sub

Private Sub Command7_Click()
     Call BOLERECEP
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    OPX% = ALTERNA%("Anulación de B.R.'s Pendientes\Re-Impresion de Remito Devolucion\Aprobación de B-R's\Control Recepciones vs. Compras \Re-Impresión de Etiquetas\Otros Accesos Directos")
    If OPX% = 1 Then
        Call ANUBORE
      ElseIf OPX% = 2 Then
        Call ReImprimirRD
      ElseIf OPX% = 3 Then
        If InStr(1, EMPREAC$, "ARCON") < 1 Then
            Call COMUNI("Opción No Disponible en la Presente Versión")
            Exit Sub
        End If
        Call APRORECEP
      ElseIf OPX% = 4 Then
        Call PROCESO_REC_VS_COMPRAS
      ElseIf OPX% = 5 Then
        Call Reimpre_Etiquetas
      ElseIf OPX% = 6 Then
        Call ACCDIR("RECEPMA")
    End If
    Command8.Enabled = True
End Sub

Private Sub ControlRecepvscompras_Click()
    Call PROCESO_REC_VS_COMPRAS
End Sub

Private Sub Form_Load()
    '
    Visible = False
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    Call CONAPLI(APLINVO$, TERMINA%) ' Busca la aplicacion a mostrar
    If TERMINA% > 0 Then Call FINAL("")
    '
    'PONER EN 0 VALORES NULL DE TABLA MACONSIG
    Call ABRESTOCKS
    Call PoneEnCeroNull("MACONSIG", "Cant_Rendid", REGAF&, "NOMESS")
    Call PoneEnCeroNull("MACONSIG", "CANT_ENTREG", REGAF&, "NOMESS")
    Call PoneEnCeroNull("MACONSIG", "CANT_ASIG", REGAF&, "NOMESS")
    Call PoneEnCeroNull("MACONSIG", "CANT_DEVU", REGAF&, "NOMESS")
    '
    EPAC$ = UCase$(TRIM$(EMPREAC$))

   

    'If InStr(UCase$(EMPREAC$), "VOITH") > 0 Or InStr(UCase$(EMPRELI$), "VOITH") > 0 Then
    If InStr(UCase$(EMPREAC$), "VOITH") > 0 Then
      Call REACOMODA
    End If
    '
    Visible = True
    If APLINVO$ = "CONRECEP" Then
      Frame2.Left = 210
      Picture11.Left = Picture11.Left - 4305
      Width = Width - 4305
      Call CENTRAFORM(Me)
      DoEvents
    End If
    '
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    If InStr(1, EPAC$, "CAMBACERES") > 0 Then
       Command20.Visible = True
    End If
    If InStr(EPAC$, "ALEJPA") > 0 Then mnutransferirasector.Visible = True
    
    If EXISTE%("CALPRO07.EXE") > 0 Or EXISTE%("CALPRO09.EXE") > 0 Then
       Command10.Visible = True
       mnuCalifProveedores.Visible = True
    End If
    '
    If EXISTE%("TRAZAB07.EXE") > 0 Or EXISTE%("TRAZAB08.EXE") > 0 Or EXISTE%("TRANSA07.EXE") > 0 Then
       Command17.Visible = True
       mnuTrazabilidad.Visible = True
    End If
    '
    If XVALO(Control("BOLRECEP", "RECEIMPO")) > 0 Then
       mnuIngresoDatoDespacho.Visible = True
       mnuVerificarDatosSegunSerie.Visible = True
    End If
    '
    EsReteplast% = 0
    If InStr(1, EMPREAC$, "RETEPLAST") > 0 Then EsReteplast% = 1

    If EsReteplast% < 1 Then
        mnuClasificacionMaterialesRecibidos.Visible = False
        lnTransacciones1.Visible = False
        Command19.Visible = False
    End If
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
    If XVALO(Control("BOLRECEP", "RECECRON")) > 0 Then
        Call CREATABLA_CONTRECE
    End If
    Call CREAR_STORED_PROCEDURE
    Call GENERAR_ESTRUCTURAS_SQL
    Call CREATABLA_STOCKS
    Call CREATABLA_CALIDAD
    
    If InStr(1, EMPREAC$, "RETEPLAST") < 1 Then
        'mnuRecepcionMaterialAControlar.Visible = False
        mnuListadoDeClasificaciones.Visible = False
        lineConsultas1.Visible = False
        mnuClasificacionMaterialesRecibidos.Visible = False
        lnTransacciones1.Visible = False
    End If
    
End Sub

Sub CREATABLA_CONTRECE()
        Call CREACAMPO("STOCKS", "CONTRECE", "NumeBoRe", 6, 4, 1) ' CONTEXTO DE RECEPCION
        Call CREACAMPO("STOCKS", "CONTRECE", "IdenLote", 10, 12, 1)
        Call CREACAMPO("STOCKS", "CONTRECE", "FechRecep", 8, 0)
        Call CREACAMPO("STOCKS", "CONTRECE", "FEYHORecep", 8, 0)
        Call CREACAMPO("STOCKS", "CONTRECE", "Cod_Inte", 3, 0, 1)
        Call CREACAMPO("STOCKS", "CONTRECE", "Cod_Exte", 10, 24)
        Call CREACAMPO("STOCKS", "CONTRECE", "DescripItem", 10, 64)
        Call CREACAMPO("STOCKS", "CONTRECE", "Unidad", 10, 4)
        Call CREACAMPO("STOCKS", "CONTRECE", "CantNom", 7, 8)
        Call CREACAMPO("STOCKS", "CONTRECE", "CanPropEdit", 7, 8)
        Call CREACAMPO("STOCKS", "CONTRECE", "CanProOrig", 7, 8)
        Call CREACAMPO("STOCKS", "CONTRECE", "CAPENCOMP", 7, 8)
        Call CREACAMPO("STOCKS", "CONTRECE", "NumProv", 3, 2, 1)
        Call CREACAMPO("STOCKS", "CONTRECE", "RaSo_Prov", 10, 64)
        Call CREACAMPO("STOCKS", "CONTRECE", "NumOC", 4, 4, 1)
        Call CREACAMPO("STOCKS", "CONTRECE", "NroItemOC", 3, 0)
        Call CREACAMPO("STOCKS", "CONTRECE", "CodiEmpr", 3, 2, 1)
        
End Sub

Sub REACOMODA()
    '
    Frame3.Visible = False
    Frame2.Height = 2910
    Frame6.Left = 4515
    Frame6.Top = 3180
    '
    Frame5.Visible = False
    Frame1.Height = 3615
    '
    Command10.Visible = False
    Command17.Visible = False
    '
    MARCOBARRA.Top = 3090
    Image2.Top = 3405
    Height = 4320
    '
    Command29.Visible = False
    Command12.Top = 1050
    Command14.Top = 1695
    Command8.Top = 2340
    '
    DoEvents
    '
End Sub
'
Sub APROCLI(NC5%, FF5%, NROBOL5&, REMI5$, DP15%)
    'APROBAR RECEPCIONES DE ORDEN DE COMPRA TEXTO LIBRE
    Call ABRESTOCKS
    With Bolrecep
         SQLX$ = " NumeBoRe = " & NROBOL5& & " AND NumRem_Prov = '" & REMI5$ & "' "
         ' CON ESTE CRITERIO DE BUSQUEDA ALCANZA Y SOBRE
         ' YA QUE NO HAY RECEPCIONES PARCIALES.
         .FindFirst SQLX$
         If .NoMatch = False Then
               CANREN5# = !CantNom
               NOC3& = !NUMOC
               DP2% = !DEPOENT
           Else
               Call COMUNI("Imposible Continuar. Número de B/R no Válido")
               Screen.MousePointer = 1
               Exit Sub
         End If
    End With
    '
    '
    Call ABRECOMPRAS
    With OCOMDETA
            SQLX$ = " NUME_OCOM = " & NOC3& & ""
            'ALCANZA TAMBIEN POR Q LA O/C SOLO PUEDE
            'TENER UN ITEM.
            .FindFirst SQLX$
            If .NoMatch = False Then
                On Error Resume Next
                IMPCONCO% = !IContab
                'ACA TOMO 'EV' y 'Grupo'
                NEV$ = !Nro_Otrab
                NGR$ = !Nro_Grup
                DESCRICOR$ = !Descritem
                DESCRILAR$ = !Delibre
                On Error GoTo 0
                '
            Else
               Call COMUNI("Imposible Continuar. Número de O/C no Válido")
               Screen.MousePointer = 1
               Exit Sub
            End If
        End With
    'MUESTRO LOS DATOS DE LA RECEPCION PROVISORIA
    With APROCLIB04
       .Label4 = NOC3&
       .Text1 = DESCRICOR$
       .Text2 = DESCRILAR$
       .Label5 = NC5%
       .Label11 = rasocli(NC5% * (-1))
       .Label7 = FORMATNUM$(CANREN5#, "F8.1")
       .Label2 = FECHATEX(FF5%)
       .Label9 = NROBOL5&
       .Label22 = REMI5$
       .Label17 = NEV$
       .Label19 = NGR$
       .Label13 = DP2%
       .Label14 = IMPCONCO%
       .Label16 = ECOARCH("TADEPOSI", Chr(DP2%))
       .Show 1
    End With
    '
End Sub

Sub INITATIREC()
     '
     If ULTREG&("TATIREC") < 1 Then
         X$ = REGBLAN$("TATIREC")
         Call REPLA(X$, Chr$(1), 1, 1)
         Call REPLA(X$, "Material General", 2, 31)
         Call REPLA(X$, "BOLREGEN", 49, 16)
         Call REPLA(X$, "BOLREGEN", 65, 16)
         Call REGAPP("TATIREC", X$)
         Call CIERRARCH("TATIREC")
     End If
     '
End Sub

Function PRONUBOLRE&()
     '
     Call ABRESTOCKS
     PRONUBOLRE& = 1 + XVALO(VALOBADA("BOLRECEP", "max(NumeBoRe)", "NUMEBORE > 0", "NOMESS"))
     '
End Function
'
Sub BOLERECEP()
     '
     '
     Dim IMPCONOC%

     Call CREATABLA_STOCKS
     Call INITATIREC                  ' INICIALIZA TABLA DE TIPOS DE RECEPCION
     '
     ' ****** CAMBIE ACA
1100 MODOGABAL% = 0
     If InStr(EMPREAC$, "GABAL") > 0 Then MODOGABAL% = 1
     '
     Call CIERRARCH("*.*")
     Call SETULTREG("!BOLREDET", 0)
     JJ& = 0
     '
     NROBOL& = PRONUBOLRE&
     NORDE% = 0
     NUOC& = 0 'INICIALIZO POR QUE  EN CONTEPLAST AL HACER SIGUIENTE RECEP SIN O.C. TOMABA LA ANTERIOR
     FORECEP07.OrdenesSeleccionadas$ = ""
     FORECEP07.OrdenesSeleccionadasConsignaciones = ""

     '
     TXTRAZA_BR$ = ""
     If TIMATEI% = 0 Then X$ = REGLEIDO$("TATIREC", 1): TIMATEI% = Asc(Left$(X$, 1))
     If TRIM$(Mid$(EBOLAN$, 4, 12)) <> "" Then NC% = 0
     '
     VDEF$ = MKI$(NC%) + Chr$(TIMATEI%) + CODICLI$((-1) * NC%)
     EBOL$ = OBJPLA$("SELPROVRECACC", VDEF$)
     If EBOL$ = "" Then
        Screen.MousePointer = 1 '\\\OT-06-0179-WAR-12/05/06///
        Exit Sub
     End If
     EBOLAN$ = EBOL$
     '
     NC3% = CVI(Left$(EBOL$, 2))
     If NC3% > 0 Then
       NC% = NC3%
       GoTo 15
     End If
     '
     NC% = NUMEPRO%(TRIM$(Mid$(EBOL$, 4, 12)))
15   If NC% < 1 Or STATCLI%((-1) * NC%) < 1 Then
       Call MENSERR(24, "Proveedor Inexistente")
       GoTo 1100
     End If
     '
     TIMATEI% = Asc(Mid$(EBOL$, 3, 1))
     If TIMATEI% < 1 Or ECOARCH("TATIREC", Chr$(TIMATEI%)) = "" Then
         Call MENSERR(24, "Falta Tipo de Material")
         GoTo 1100
     End If
     BobinaMtsLineales$ = Control("", "REPAMTLI")
     '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 1 = " & Now & vbCrLf

     REGT& = TIMATEI%
     X$ = REGLEIDO$("TATIREC", REGT&)
     FORMPLA$ = TRIM$(UCase$(Mid$(X$, 33, 16)))
     FORMTAB$ = TRIM$(UCase$(Mid$(X$, 49, 16)))
     FORIPRE$ = TRIM$(UCase$(Mid$(X$, 65, 16)))
     '
     NC1% = (-1) * NC%
     '
1105 Y$ = PEDY$
     If Y$ = "" Then
       Y$ = Space$(128)
       Call REPLA(Y$, Chr$(0), 121, 1)
     End If
     '
TX1$ = "BR-: " & SAFETEXT$(NROBOL&) & " Prov." & SAFETEXT$(NC1%)
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 1.1 = " & TX1$ & "  -  " & Now & vbCrLf
     
     Call REPLA(Y$, MKS$(NROBOL&), 1, 4)
     Call REPLA(Y$, EBOL$, 7, 2)
     Call REPLA(Y$, MKI$(HOY(HO$)), 5, 2)
     Call REPLA(Y$, rasocli$(NC1%), 9, 32)
     Call REPLA(Y$, DOMICLI$(NC1%), 41, 32)
     CLOCA$ = TRIM$(CPOSCLI$(NC1%)): If CLOCA$ <> "" Then CLOCA$ = CLOCA$ + " "
     Call REPLA(Y$, CLOCA$ + LOCACLI$(NC1%), 73, 32)
     Call REPLA(Y$, Chr$(0), 121, 1)
    '
     OBJE$ = "BOLREC-ENCB"
     If Control$("", "NUMEBREC") = "MANUAL" Then
         OBJE$ = "BOLREC-ENCB-M"
     End If
     '
     ' CONTROL DE VIGILANCIA
     If Control$("BOLRECEP", "NUCONVIG") = "SI" Then
         OBJE$ = "BOLREC-ENCC"
     End If
     '
     TICOMPRO$ = TRIM$(Control$("", "COMPBREC"))
     '
     PEDY$ = OBJPLA$(OBJE$, Y$)
     If PEDY$ = "" Then
        Screen.MousePointer = 1 '\\\OT-06-0179-WAR-12/05/06///
        Exit Sub
     End If
     '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 2 = " & Now & vbCrLf

     NUCONVI$ = TRIM$(Mid$(PEDY$, 145, 24))
     FF% = CVI(Mid$(PEDY$, 5, 2))
     FEX = FF%: FECHAREC$ = FECHATEX$(FEX)
     If TRIM$(FECHAREC$) = "" Then
         Call MENSERR(24, "Fecha Incorrecta")
         GoTo 1105
     End If
     '
     REMI$ = TRIM$(Mid$(PEDY$, 105, 16))
     If REMI$ = "" Then
        Call MENSERR(24, "Falta Número de Remito del Proveedor")
        GoTo 1105
     End If
     '
     patentetransporte$ = TRIM$(Mid$(PEDY$, 145, 16))
     '
     DP1% = Asc(Mid$(PEDY$, 121, 1))
     DEPOSAL$ = TRIM$(ECOARCH$("ECODEP", Chr$(DP1%)))
     If DP1% < 1 Or DEPOSAL$ = "" Then
         Call MENSERR(24, "Falta Depósito de Entrada")
         GoTo 1105
     End If
     DEPOSAL$ = DEPOSAL$ + " (" + TRIM$(Str$(DP1%)) + ")"
     '
     NROPRO& = NROBOL&
     NROBOL& = CVS(Left$(PEDY$, 4))
     If NROBOL& < 1 Then
         Call MENSERR(24, "Número Incorrecto de B.R.")
         NROBOL& = NROPRO&
         Call REPLA(PEDY$, MKS$(NROBOL&), 1, 4)
         GoTo 1105
     End If
     ' COMPRUEBA QUE NO EXISTA EL REMI
     NC2% = NC%
     NR1$ = REMI$
     If Control$("BOLRECEP", "SUPREMI") <> "SI" Then  '\\\OT-06-0515-RG-28/11/06///
        R1% = EXIREMI(NR1$, NC2%)
        If R1% = 1 Then
           Call COMUNI("El Número de Remito ya fue Procesado\en Recepción Anterior.")
           GoTo 15
        End If
     End If
     '
1112 Call CIERRARCH("*.*")
     '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 3 = " & Now & vbCrLf
     
     ' ****** CAMBIE ACA
     If MODOGABAL% = 1 Then ' SOLO GABAL
113     OPX% = COMALTER%("Opciones de Carga\\Carga Manual\OF's de Armadores")
        If OPX% < 1 Or OPX% > 2 Then GoTo 1105
        '
        If OPX% = 1 Then
            If DERACCE%("CARMANU", "Carga Manual de Recepcion") < 2 Then GoTo 1100
        Else
           ' RECEPCION DE ARMADORES
           Call CARGALISEL("OFARPERE", "MOVISTO", "DOPRICOR/A12;COD_EXTE/A16;DESCRIP/A32;CANTINGRE/F10.1", "TIPOTRAN=1 AND CODIMOVI='OF' AND CANTINGRE > 0.05")
           Call SELECHO("OFARPERE/O/Fabricación Pendientes de Cierre")
           OFRECIBI$ = TRIM$(VALACT1$("OFARPERE"))
           OFRECIB2$ = TRIM$(VALACT2$("OFARPERE"))
           If OFRECIBI$ = "" Then GoTo 113
           '
           GB_CIXI% = XVALO(VALOBADA("ORDEFABR", "COD_INTE", "IDSUBOR='" & OFRECIBI$ & "'"))
           GB_NUME_OCOM& = XVALO(VALOBADA("ORDEFABR", "NUME_OCOM", "IDSUBOR='" & OFRECIBI$ & "'"))
           GB_FENTRESO% = CVINT(VALOBADA("OCOMDETA", "Fentre_Sol", "NUME_OCOM=" & Str$(GB_NUME_OCOM&) & " AND COD_INTE=" & Str$(GB_CIXI%)))
           GB_OFRECI$ = TRIM$(VALACT2$("OFARPERE"))
           GB_CANTIDAD# = XVALO(Mid$(OFRECIB2$, 53))
           '
           If GB_CIXI% = 0 Or GB_NUME_OCOM& = 0 Or GB_CANTIDAD# < 0.1 Then
                Call COMUNI("Imposible Continuar.\\No se Ha Encontrado Código o O/Compra o Cantidad")
                GoTo 113
           End If
           '
           MODOGABAL% = 2
           Y$ = REGBLAN$("BOLREDET")
           Call REPLA(Y$, MKI$(GB_CIXI%), 83, 2)
           Call REPLA(Y$, Unimed$(GB_CIXI%), 85, 4)
           Call REPLA(Y$, MKD$(GB_CANTIDAD#), 89, 8)
           Call REPLA(Y$, Chr$(1), 98, 1)
           Call REPLA(Y$, MKD$(GB_CANTIDAD#), 163, 8)
           Call REPLA(Y$, MKS$(GB_NUME_OCOM&), 151, 4)
           Call REPLA(Y$, MKI$(GB_FENTRESO%), 187, 2)
           Call REPLA(Y$, MKS$(GB_OFRECIBI$), 189, 4)
           '
           Call SETULTREG("!BOLREDET", 0)
           Call GRAREG("!BOLREDET", Y$, 1)
           Call CIERRARCH("!BOLREDET")
           '
           GoTo 1113
           '
        End If
     End If
     '
     Screen.MousePointer = 11
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 4 = " & Now & vbCrLf
     'si es Empres Conteplast y Proveedor Cambaceres
     If TIMATEI% > 1 Then
        If InStr(1, EMPREAC$, "CONTEPLAST") > 0 Or InStr(1, EMPREAC$, "QUILMES") > 0 Or InStr(1, EMPREAC$, "CAMBACERES") > 0 Then   'en el caso que proveedor cambaceres
            If (InStr(1, UCase$(rasocli$((-1) * NC2%)), "CAMBACERES") > 0 Or InStr(1, UCase$(rasocli$((-1) * NC2%)), "QUILMES") > 0) Then
                OPX% = COMALTER%("Opciones de Carga\\Carga Manual\Por Documento Adjunto Código QR (Sin Orden de Compra)")
                If OPX% > 1 Then
                   If RecepcionDeBobinasCambaQR(NROBOL&, NC%, REMI$, TIMATEI%) = True Then
                     AVENTA$ = "BOLREPAP": ARCHICA$ = "BOLREPAP": TIMATEFIL% = 1
                     VTB% = 1
                     GoTo 321
                   End If
                End If
            End If
        End If
     End If
     
     Call ABRECOMPRAS
     SQLX$ = "Select COD_INTE from OCOMDETA WITH(NOLOCK) "
     SQLX$ = SQLX$ + " Where Npro_Ocom = " & CStr(NC%) & " AND Stat_Ocom < 8 "
     SQLX$ = SQLX$ + " AND Stat_Ocom <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
     SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
     '
     Dim COMPTEMP As ADODB.Recordset
     Set COMPTEMP = READSET(SQLX$)
     
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 5 = " & Now & vbCrLf
     
     HAYCODI% = 0: HAYLIBRE% = 0
     With COMPTEMP
       On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         GoTo 45       ' no hay pendiente de recepcion --> va a codificado
       End If
       On Error GoTo 0
       Do While Not .EOF
         If XVALO(!cod_inte) > 0 Then
             HAYCODI% = 1
           Else
             HAYLIBRE% = 1
         End If
       .MoveNext
       Loop
     End With
     COMPTEMP.Close
     Set COMPTEMP = Nothing
     '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 7 = " & Now & vbCrLf
     
     'CONSIDERA CONSIGANCIONES PENDIENTES
     CONDI$ = "Npro_Ocom = " & CStr(NC%) & " AND ((Cant_Rendid < Cant_Entreg - 0.001) Or Cant_Entreg Is Null)"
     If CantRegistros("MACONSIG", CONDI$) > 0 Then HAYCODI% = 1
     '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 8 = " & Now & vbCrLf
     
     Screen.MousePointer = 1
     If HAYLIBRE% < 1 Then GoTo 45
     If HAYCODI% < 1 And HAYLIBRE% > 0 Then GoTo 42
     '
     JJ& = 0
     '
     'ACA NUEVO
     RESP4% = COMALTER("Elija el Material Recibido:\\Codificado\Texto Libre")
     If RESP4% < 1 Or RESP4% > 2 Then
         Screen.MousePointer = 1
         GoTo 1100
     End If
     '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 10 = " & Now & vbCrLf
     
     If RESP4% = 1 Then
         GoTo 45
       Else
42       Call RECEOCLI(NC%, FF%, NROBOL&, REMI$, DP1%, NUCONVI$)
         Screen.MousePointer = 1
         GoTo 1100
     End If
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 11 = " & Now & vbCrLf
     
     '
45   DARX$ = TRIM$(Control$("BOLRECEP", "DIANREC"))
     If DARX$ = "" Then
          DIANRE& = 10000
        Else
          DIANRE& = Val(DARX$)
     End If
     HOIX1 = HOY(HOS$)
     '
     ANUCA% = 0
     If Control$("BOLRECEP", "ANUCANTI") = "SI" Then
       ANUCA% = 1
     End If
     '
     FiltroModoOrdenCompra$ = ""
     FiltroModoOrdenCompraConsignaciones$ = ""
     MODOFILTRO% = COMALTER%("Opciones de Recepción:\\Todos Los Pendientes\Selección Por O-Compra")
     If MODOFILTRO% < 1 Or MODOFILTRO% > 2 Then GoTo 1100
     If MODOFILTRO% = 2 Then
        Call FiltroRecepcionPorCompras(NC%)
        FiltroModoOrdenCompra$ = OrdenesSeleccionadas$
        FiltroModoOrdenCompraConsignaciones$ = OrdenesSeleccionadasConsignaciones$
     End If
     '
     Screen.MousePointer = 11
     Call ABRECOMPRAS
     '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 12 = " & Now & vbCrLf
     
     ArchivoRecepcion$ = "BOLREDET"
     If TIMATEI% > 1 Then
        ArchivoRecepcion$ = "BOLREPAP"
     End If
     CONDI$ = "Npro_Ocom = " & CStr(NC%) & " AND Cod_Inte > 0 AND Stat_Ocom < 8 "
     CONDI$ = CONDI$ + " AND Stat_Ocom <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
     CONDI$ = CONDI$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
     '
     SQLX$ = "Select cant_ocom, Cant_Rec,Coef_Unids , Nume_Ocom, Fentre_Sol, COD_INTE, Doc_Orig, NuOrd_Item, N_ORSERV "
     SQLX$ = SQLX$ + " from OCOMDETA WITH(NOLOCK) "
     SQLX$ = SQLX$ + " Where " + CONDI$
     ' CONTEMPLA LA MODALIDAD DE FILTRO POR ORDEN DE COMPRA
     If TRIM$(FiltroModoOrdenCompra$) <> "" Then SQLX$ = SQLX$ + " AND " & FiltroModoOrdenCompra$
     '
     If InStr(1, EMPREAC$, "FERVI") > 0 Then SQLX$ = SQLX$ + " ORDER BY NUME_OCOM, FEMI_OCOM"
     '
     Dim COMPTEMP_1 As ADODB.Recordset
     Set COMPTEMP_1 = READSET(SQLX$)
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 13 = " & Now & vbCrLf
     With COMPTEMP_1
       '
       On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Screen.MousePointer = 1
         HAYOC% = 0
         GoTo 1777
       End If
       On Error GoTo 0
       '
       JJ& = 0
        
       Do While (.EOF = False)
         CAPEN = XVALO(!cant_ocom) - XVALO(!Cant_Rec)
         COEFUNI = XVALO(!Coef_Unids)
         If COEFUNI < 0.0001 Then COEFUNI = 1
         If CAPEN > 0.005 Then
           NUOC& = XVALO(!nume_ocom)
           FENTRESO% = CVINT(!Fentre_Sol)
           FESO2 = FENTRESO%
           If DIENTRE%(HOIX1, FESO2) <= DIANRE& Then
             CI% = XVALO(!cod_inte)
             
             If TIMATEI% > 1 And (InStr(1, EMPREAC$, "CONTEPLAST") > 0 Or InStr(1, EMPREAC$, "CAMBACERES")) > 0 Then
                If TIMATE%(CI%) <> 41 Then GoTo 49 '
             End If                                                  '
             
             LOTECO$ = SAFETEXT$(!Doc_Orig)
             Y$ = REGBLAN$("!BOLREDET")
             Call REPLA(Y$, MKI$(CI%), 83, 2)
             Call REPLA(Y$, Unimed$(CI%), 85, 4)
             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 89, 8)
                If ANUCA% = 1 Then Call REPLA(Y$, MKD$(0), 89, 8)
             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 163, 8)
             '
             Call REPLA(Y$, Chr$(XVALO(!NuOrd_Item)), 98, 1)
             Call REPLA(Y$, MKS$(0), 100, 4) 'no tiene AE
             '
             Call REPLA(Y$, LOTECO$, 107, 12)
             Call REPLA(Y$, MKS$(NUOC&), 151, 4)
             Call REPLA(Y$, MKI$(FENTRESO%), 155, 2)
             Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
             Call REPLA(Y$, MKS$(!N_ORSERV), 189, 4)
             '
             If ESTRAZABLE(CI%) = 2 Then     ' BUSCA POR NUMEROS DE SERIE
                NOFAX$ = SAFETEXT$(!Doc_Orig)
                If NOFAX$ Like "OF*" Then
                   SQLX$ = "SELECT UBICACION, NUMEROSERIE, IDSUBOR FROM TANUMSE WITH(NOLOCK) "
                   SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CI%
                   SQLX$ = SQLX$ + " AND IdSubOr = '" & NOFAX$ & "'"
                   SQLX$ = SQLX$ + "ORDER BY NUMEROSERIE "
                   Dim SELNUSE As ADODB.Recordset
                   ' Set SELNUSE = New ADODB.Recordset
                   Set SELNUSE = READSET(SQLX$)
                   With SELNUSE
                     Do While Not .EOF
                       If InStr(UCase$(SAFETEXT$(!UBICACION)), "PLANTA MEDITEA") < 1 Then
                         Y1$ = Y$
                         Call REPLA(Y1$, MKD$(1), 89, 8)
                           If ANUCA% = 1 Then Call REPLA(Y1$, MKD$(0), 89, 8)
                         Call REPLA(Y1$, MKD$(1), 163, 8)
                         Call REPLA(Y1$, SAFETEXT$(!NUMEROSERIE), 107, 12)
                         Call REPLA(Y1$, SAFETEXT$(!idsubor), 135, 16)
                         HAYOC% = 1
                         JJ& = JJ& + 1
                         Call GRAREG("!" & ArchivoRecepcion$, Y1$, JJ&)
                       End If
                     .MoveNext
                     Loop
                   End With
                   GoTo 49
                End If
             End If
             
             If TIMATEI% > 1 And (InStr(1, EMPREAC$, "CONTEPLAST") > 0 Or InStr(1, EMPREAC$, "CAMBACERES") > 0) Then
                    AnchoBobina# = 0: If BobinaMtsLineales$ = "SI" Then AnchoBobina# = 1000
                Call REPLA(Y$, MKD$(AnchoBobina#), 193, 8)
                    LargoBobina# = 0: If BobinaMtsLineales$ = "SI" Then LargoBobina# = XVALO(PESUNI(CI%))
                Call REPLA(Y$, MKD$(LargoBobina#), 201, 8)
                Call REPLA(Y$, MKS$(1), 209, 4)
             End If
             
             HAYOC% = 1
             JJ& = JJ& + 1
             Call GRAREG("!" & ArchivoRecepcion$, Y$, JJ&)
             
             If TIMATEI% > 1 And (InStr(1, EMPREAC$, "CONTEPLAST") > 0 Or InStr(1, EMPREAC$, "CAMBACERES") > 0) Then
                MetrosPorUnidad# = XVALO(PESUNI(CI%))
                If MetrosPorUnidad# < 0.005 Then
                    CodigoProducto$ = CODEXT$(CI%)
                    Call COMUNI("Falta Definir Metros Por Unidad para el Artículo '" & CodigoProducto$ & "'")
                    GoTo 49
                    MetrosPorUnidad# = XVALO(InputBox("Metros por Bobina", "Metros Por Unidad para " & CodigoProducto$, ""))
                    If MetrosPorUnidad# < 0.005 Then
                        GoTo 49
                    Else
                        Call ACTUREGISTRO("master", "color", "codint=" & CStr(CI%), TRIM$(FORMATNUM$(MetrosPorUnidad#, "F14.5")), REGAF&)
                    End If
                End If
                If (CAPEN / MetrosPorUnidad#) > 32000 Then GoTo 49
                CantidadBobinasPendientes% = Int(CAPEN / MetrosPorUnidad#) - 1
                Resto# = CAPEN Mod MetrosPorUnidad#
                If Resto# > 0.005 Then CantidadBobinasPendientes% = CantidadBobinasPendientes% + 1
                For BOBINA% = 1 To CantidadBobinasPendientes%
                    JJ& = JJ& + 1
                    Call GRAREG("!" & ArchivoRecepcion$, Y$, JJ&)
                Next BOBINA%
             End If
           End If
         End If
49      .MoveNext
       Loop
     End With
     COMPTEMP_1.Close
     Set COMPTEMP_1 = Nothing
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 14 = " & Now & vbCrLf
   
   '
     'TENGO QUE FIJARME SI TIENE ALGUN MATERIAL EL PROV
     'CONSIGNADO EN MACONSIG CON EL MISMO Nº DE PROV
1777 Call ABRESTOCKS
     SQLX$ = "Select COD_INTE, Nro_Ocom, cant_entreg, cant_rendid, cant_devu  from MACONSIG WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE Npro_Ocom = " & CStr(NC%)
     SQLX$ = SQLX$ + " AND NRO_OCOM NOT LIKE 'AR-%' " ' OBVIA AJUSTES CON CANT_DEVU < 0
     SQLX$ = SQLX$ + " AND (Cant_Rendid + Cant_Devu < Cant_Entreg - 0.001)"
     'SQLX$ = SQLX$ + " OR CANT_RENDID IS NULL) "
     ' CONTEMPLA LA MODALIDAD DE FILTRO POR ORDEN DE COMPRA
     If TRIM$(FiltroModoOrdenCompraConsignaciones$) <> "" Then SQLX$ = SQLX$ + " AND " + FiltroModoOrdenCompraConsignaciones$
     '
     Dim CONSIGTMP As ADODB.Recordset
     Set CONSIGTMP = READSET(SQLX$)
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 15 = " & Now & vbCrLf
     With CONSIGTMP
       '
       On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Screen.MousePointer = 1
         GoTo 1129
       End If
       On Error GoTo 0
       '
       Y$ = REGBLAN$("!BOLREDET")
       Do While .EOF = False
         CI% = XVALO(!cod_inte)
         NUOC& = Val(Mid$(!nro_ocom, 4, 6))
         CAPEN = XVALO(!cant_entreg) - (XVALO(!CANT_RENDID) + XVALO(!cant_devu))
         If CAPEN >= 0.005 Then
            Call REPLA(Y$, MKI$(CI%), 83, 2)
            Call REPLA(Y$, Unimed$(CI%), 85, 4)
            Call REPLA(Y$, MKD$(0), 89, 8)
            Call REPLA(Y$, "MP " + TRIM$(CSTRING$(MKS$(CAPEN), 3, 12, 1, 2)) + " " + Unimed$(CI%), 107, 12)
            Call REPLA(Y$, MKS$(NUOC&), 151, 4)
            Call REPLA(Y$, MKS$((-1) * i&), 189, 4)
            JJ& = JJ& + 1
            Call GRAREG("!BOLREDET", Y$, JJ&)
         End If
       .MoveNext
       Loop
     End With
     CONSIGTMP.Close
     Set CONSIGTMP = Nothing
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 16 = " & Now & vbCrLf

     '\\\OT-06-0149-WAR-25/04/06///
1129 Call ABRECOMPRAS
     SQLX$ = "Select * from AUTOENT WITH(NOLOCK) "
     
     SQLX$ = SQLX$ + " Where Stat_Aten < 8 "
     SQLX$ = SQLX$ + " AND  Stat_Aten <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
     SQLX$ = SQLX$ + " AND Acuen_Aten < Cant_Aten - 0.001 "
     SQLX$ = SQLX$ + " AND Npro_Aten = " & NC% & " "
     SQLX$ = SQLX$ + " AND Cod_Inte > 0 " 'EXCLUYE LAS DE 0/C's TEXTO LIBRE
     'Set COMPTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     Dim COMPTEMP_2 As ADODB.Recordset
     Set COMPTEMP_2 = READSET(SQLX$)
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 17 = " & Now & vbCrLf
     
     With COMPTEMP_2
       On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Screen.MousePointer = 1
         GoTo 1130
       End If
       On Error GoTo 0
       '
       Do While (.EOF = False)
         CAPEN = XVALO(!Cant_Aten) - XVALO(!Acuen_Aten)
         COEFUNI = 1 '!Coef_Unids
         If COEFUNI < 0.0001 Then COEFUNI = 1
         If CAPEN > 0.005 Then
           NUOC& = XVALO(!nume_ocom)
           FENTRESO% = CVINT(!Fentre_Sol)
           FESO2 = FENTRESO%
           If DIENTRE%(HOIX1, FESO2) <= DIANRE& Then
             CI% = XVALO(!cod_inte)
             Y$ = REGBLAN$("!BOLREDET")
             Call REPLA(Y$, MKI$(CI%), 83, 2)
             Call REPLA(Y$, Unimed$(CI%), 85, 4)
             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 89, 8)
             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 163, 8)
             If ANUCA% = 1 Then Call REPLA(Y$, MKD$(0), 89, 8)
             '
             Call REPLA(Y$, Chr$(XVALO(!NuOrd_Item)), 98, 1)
               NANTEN& = 0
               On Error Resume Next
               NANTEN& = XVALO(!Nume_Aten)
               On Error GoTo 0
             Call REPLA(Y$, MKS$(NANTEN&), 100, 4) 'Nº Aut. Ent.
             '
             Call REPLA(Y$, LOTECO$, 107, 12)
             Call REPLA(Y$, MKS$(NUOC&), 151, 4)
             Call REPLA(Y$, MKI$(FENTRESO%), 155, 2)
             Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
             'Call REPLA(Y$, MKS$(!N_OrServ), 189, 4)
             HAYOC% = 1
             JJ& = JJ& + 1
             Call GRAREG("!BOLREDET", Y$, JJ&)
           End If
         End If
        .MoveNext
       Loop
     End With
     COMPTEMP_2.Close
     Set COMPTEMP_2 = Nothing
     '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 18 = " & Now & vbCrLf
     
1130 If HAYOC% = 0 Then  '1129 antes de OT-06-0149
       Call COMUNI("No hay Ordenes de Compra Pendientes\para este Proveedor.")
     End If
     '
     Call SETULTREG("!BOLREDET", JJ&)
     Call CIERRARCH("!BOLREDET")
     If TIMATEI% > 1 Then
        Call SETULTREG("!BOLREPAP", JJ&)
        Call CIERRARCH("!BOLREPAP")
     End If
     '
1113 TIMATEFIL% = 0
     AVENTA$ = "BOLREDET": ARCHICA$ = "BOLREDET"
     If ANUCA% = 1 Then AVENTA$ = "BOLRENUE"
     If TIMATEI% > 1 Then AVENTA$ = "BOLREPAP": ARCHICA$ = "BOLREPAP": TIMATEFIL% = 1
     '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 19 = " & Now & vbCrLf
     
     ' ****** CAMBIE ACA
     ATRI$ = ""
     If MODOGABAL% = 2 Then
        ATRI$ = ATRI$ + "/POSEDIT(1)=2"
        ATRI$ = ATRI$ + "/POSEDIT(2)=2"
        ATRI$ = ATRI$ + "/POSEDIT(3)=2"
        ATRI$ = ATRI$ + "/POSEDIT(4)=2"
        ATRI$ = ATRI$ + "/POSEDIT(5)=2"
        ATRI$ = ATRI$ + "/POSEDIT(6)=2"
        ATRI$ = ATRI$ + "/POSEDIT(9)=2"
        ATRI$ = ATRI$ + "/POINDEL=NO" ' DESHABILITA EL INSERTAR Y EL DEL
     End If
     '
     If TIMATEI% > 1 Then
        ATRI$ = ATRI$ + "/POSEDIT(7)=2"
        'ATRI$ = ATRI$ + "/POSEDIT(8)=2"
        ATRI$ = ATRI$ + "/POSEDIT(9)=2"
     End If
     ' ****** CAMBIE ACA
     If Control("BOLRECEP", "CONPRECI") = "SI" And ULTREG("!BOLREDET") > 0 And MODOFILTRO% = 2 Then
        Call PRESENTAVENTANAVALORIZADA(Abs(NC1%))
        If BRECPAHP.APROBADO% > 0 Then
          Unload BRECPAHP
          VTB% = 0
        Else
          VTB% = -1
        End If
          GoTo 321
     End If
     '
     ES_SCRAP% = 0
     If InStr(EMPREAC$, "QUILMES") > 0 And TIMATEI% = 1 Then
        OPX% = COMALTER%("Opciones de Carga\\Carga Manual\Carga - SOLO SCRAP x QR")
        If OPX% = 2 Then
            RESUL = PRESENTAVENTANASCRAP(Abs(NC1%))
            If RESUL = True Then
              Unload BREQUILMFRM
              VTB% = 0: ES_SCRAP% = 1
            Else
              VTB% = -1
            End If
              GoTo 321
        End If
     End If


     
     If TIMATEI% > 1 And (InStr(1, EMPREAC$, "CONTEPLAST") > 0 Or InStr(1, EMPREAC$, "CAMBACERES") > 0) Then
       SQLX$ = " SELECT max(NumeBoRe) as maximo from BOLRECEP"
       Dim BolreTemp_0 As ADODB.Recordset
       Set BolreTemp_0 = FLEXCONN.Execute(FILTSQL$(SQLX$))
       If Not IsNull(BolreTemp_0!Maximo) Then
           NROBOL& = 1 + (BolreTemp_0!Maximo)
       End If
       BolreTemp_0.Close
       Set BolreTemp_0 = Nothing
       If HAYOC% < 1 Then Call SETULTREG("!BOLREPAP", 0) ' para liberar por que estaba quedando info y en recepcion de bobinas mostraba datos asi no haya compreas pendientes
       If RecepcionDeBobinas(NROBOL&, NC%, REMI$, TIMATEI%) = True Then
            VTB% = 1
            GoTo 321
       Else
            GoTo 1100
       End If
     End If
     
     If InStr(1, EMPREAC$, "RETEPLAST") > 0 Then
        idRecepcion& = 0
        If RecepcionDeMateriales(idRecepcion&) = True Then
            VTB% = 1
            GoTo 321
        Else
            GoTo 1100
        End If
     End If
          
320  VTB% = VENTABU%(AVENTA$ + "/NC/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(rasocli$(NC1%)) + " - " + REPLACAR$(FECHAREC$, "/", "-") & ATRI$)
321  TIMATEFIL% = 0
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 20 = " & Now & vbCrLf
     If VTB% < 0 Or ULTREG&("!" + ARCHICA$) < 1 Then
       X$ = "": Y$ = "": YY$ = ""
       EBOL$ = "": PEDY$ = "": DESCRI$ = ""
       GoTo 1100
     End If
     '
     If ARCHICA$ = "BOLREPAP" Then
        ' CALCULA LOS METROS CUADRADOS
        For YX& = 1 To ULTREG&("!" + ARCHICA$)
           YY$ = REGLEIDO$("!" + ARCHICA$, YX&)
           CI% = CVI(Mid$(YY$, 83, 2))
           If TIMATE%(CI%) <> 1 And InStr(1, EMPREAC$, "CONTEPLAST") < 1 And InStr(1, EMPREAC$, "CAMBACERES") < 1 Then
               DECODE$ = TRIM$(CODEXT$(CI%)) + " - " + DESCRIT0$(CI%)
               Call MENSERR(24, "Código " + DECODE$ + "\No Corresponde a Tipo de Material 'PAPELES'")
               GoTo 1113
           End If
           ANCHOX# = CVD(Mid$(YY$, 193, 8))
           If ANCHOX# < 0.005 Then ANCHOX# = 1
           LARGOX# = CVD(Mid$(YY$, 201, 8))
           UNIDS& = CVS(Mid$(YY$, 209, 4))
           '
           If UNIDS& < 1 Then
              UNIDS& = 1
              Call REPLA(YY$, MKS$(1), 209, 4)
              Call GRAREG("!" + ARCHICA$, YY$, YX&)
           End If
           '
           If UNIDS > 99 Then
               Call MENSERR(24, "Cantidad no válida - Supera 99 Unidades.")
               GoTo 1113
           End If
           '
           If Control("BOLRECEP", "CTRLPESO") = "SI" Then
                PesoBalanza# = CVS(Mid$(YY$, 253, 4))
                If PesoBalanza# < 0.005 Then
                    Call MENSERR(24, "Imposible Continuar. Debe Informar un Peso para cada Artículo Recibido.")
                    GoTo 1113
                End If
           End If
           
           If InStr(1, EMPREAC$, "CAMBACERES") > 0 And Unimed$(CI%) = "U." Then 'clave seguimiento: RecepcionCambaBobina
              If TIMATE%(CI%) = 41 Then 'SI ES BOBINAS
                  CAN# = 1
              End If
           Else
              CAN# = LARGOX# * ANCHOX# / 1000 'ESTO COMO ESTABA
           End If
              
           Call REPLA(YY$, MKD$(CAN#), 89, 8)
           Call GRAREG("!" + ARCHICA$, YY$, YX&)
        Next YX&
     End If
     '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 21 = " & Now & vbCrLf
     
     For YX& = 1 To ULTREG&("!" + ARCHICA$)
       YY$ = REGLEIDO$("!" + ARCHICA$, YX&)
       CI% = CVI(Mid$(YY$, 83, 2))
       CAN# = CVD(Mid$(YY$, 89, 8))
       
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           If CAN# > 0.005 Then
             GoTo 1141
           End If
         End If
       End If
     Next YX&
     Call MENSERR(24, "Recepción Cancelada !!!\Faltan Cantidades o Códigos.")
     GoTo 1113
     '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 22 = " & Now & vbCrLf
     
1141 CONTA% = ULTREG&("!" + ARCHICA$)
     '
     BLOSOBRE% = 0
     If Control$("BOLRECEP", "BLOSOBRE") = "SI" Then
       BLOSOBRE% = 1
       PORSOBRE = Val(Control("BOLRECEP", "PORSOBRE"))
     End If
     '
1115 If CONTA% > 0 Then
       '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 23 = " & Now & vbCrLf
       
       STMOVI09.List1.Clear
       '
       '************ACA INGRESAR VALIDACION CRONOLOGICA*********************
       If XVALO(Control("BOLRECEP", "RECECRON")) > 0 Then
          Call CONTROL_CRONOLOGICO_COMPRAS(NC, ARCHICA$, APROBO1%)
          If APROBO1% < 1 Then Unload VENTABNEW: GoTo 320
       End If
       '************FIN INGRESAR VALIDACION CRONOLOGICA *********************
       For YX& = 1 To ULTREG&("!" + ARCHICA$)
         YY$ = REGLEIDO$("!" + ARCHICA$, YX&)
         CI% = CVI(Mid$(YY$, 83, 2))
         ' validacion que no tenga marca en proceso de articulos con salidas menores a 5000 e
         ' importes por cantidad salida menores a 5000
         If STATMOVI%(CI%) < 0 Then
              ' ARMA UN LIST INFORMATIVO CON LOS ITEMS CON STATMOV - 1
              TX$ = ""
              Call REPLA(TX$, CODEXT$(CI%), 1, 16)
              Call REPLA(TX$, DESCRIT$(CI%), 17, 36)
              STMOVI09.List1.AddItem TX$
              ' LUEGO DE RECORRER LO MUESTRA
         End If

         CAN# = CVD(Mid$(YY$, 89, 8))
         UNIDS& = CVS(Mid$(YY$, 209, 4))
             If UNIDS& < 1 Then UNIDS& = 1
         If CI% < 1 Then
           Call MENSERR(24, "Código no Válido.")
           GoTo 1113
         End If
         If ANUCA% < 1 Then
           If CAN# < 0.005 Then
             Call MENSERR(24, "Falta Cantidad.")
             GoTo 1113
           End If
         End If
         '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 24 = " & Now & vbCrLf
         ' VERIFICA TRENSFERENCIA DE COMPONENTES
         If CAN# >= 1 Then
            If ESTRAZABLE(CI%) = 2 Then
               LOTECO$ = TRIM$(Mid$(YY$, 107, 12))
               If LOTECO$ <> "" Then
                  DEPARMAD% = XVALO(Control("", "DEPOARMA")) ' DEPOSITO DE ARMADOR
                  CONDII$ = "Cod_Inte = " & CI% & " and NUMEROSERIE = '" & LOTECO$ & "'"
                  NUMEORF$ = SAFETEXT$(VALOBADA("TANUMSE", "IdSubOr", CONDII$, "NOMESS"))
                  If TRIM$(NUMEORF$) <> "" Then
                     CONDI1$ = "CANTINGRE > 0 AND DepoMovi = " & DEPARMAD% & " And DoSecCor = '" & NUMEORF$ & "' AND CODIMOVI = 'TR'"
                     If CantRegistros("MOVISTO", CONDI1$, "NOMESS") < 1 Then
                        Call MENSERR(24, "Imposible Procesar Recepción de Armador Externo.\  \Aun Falta Transferencia Inicial de Materiales Componentes")
                        GoTo 1113
                     End If
                  End If
               End If
            End If
         End If
         '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 25 = " & Now & vbCrLf

         'VALIDACION SEGUN CONTROL DE RECEPCION SOLO POR ORDEN DE COMPRA
         If Control$("BOLRECEP", "RESOLOOC") = "SI" Then
           If Mid$(YY$, 107, 2) <> "MP" Then ' no es devolucion de MP
             NUMEOCOMPRA& = Abs(CVS(Mid$(YY$, 151, 4)))
             '
             If NUMEOCOMPRA& < 1 Then
                Call COMUNI("Imposible Continuar con la Recepción.\El Código: " & CODEXT(CI%) & " en Fila: " & YX& & "\No Posee una O-Compra\Para Este Proveedor")
                GoTo 1113
             End If
             '
             CONDI$ = "Npro_Ocom = " & NC% & " AND COD_INTE = " & CStr(CI%) & " AND NUME_OCOM = " & CStr(NUMEOCOMPRA&)
             CONDI$ = CONDI$ + " AND Stat_Ocom < 3 AND Stat_Ocom <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
             '
             NCOM& = XVALO(VALOBADA("OCOMDETA", "NUME_OCOM", CONDI$, "NOMESS"))
             If NCOM& < 1 Then
               Call COMUNI("Imposible Continuar con la Recepción.\El Código: " & CODEXT(CI%) & " en Fila: " & YX& & "\No Corresponde a una O-Compra\Para Este Proveedor")
               GoTo 1113
             End If
             '
           End If
         End If
         '
         
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 26 = " & Now & vbCrLf
         
         If BLOSOBRE% > 0 Then
           If Mid$(YY$, 107, 2) <> "MP" Then ' no es devolucion de MP
             '\\\OT-05-0694-WAR-25/10/05///
             'SI ES MATERIAL AUTORIZADO A RECIBIR SIN O/C
             X2$ = FILTERGETRECORD$("TAMATCON", 1, MKI$(CI%))
             If TRIM$(X2$) <> "" Then
               FELI% = CVI(Mid$(X2$, 3, 2))
               If FF% <= FELI% Then 'SI VALIDA FECHA LIMITE
                 GoTo 1114
               End If
             End If
             '\\\OT-05-0694-WAR-FIN///
             SS = CVD(Mid$(YY$, 163, 8)) * (1 + PORSOBRE / 100) + 0.05
             If CAN# * UNIDS& > CVD(Mid$(YY$, 163, 8)) * (1 + PORSOBRE / 100) + 0.05 Then
               DECODE$ = TRIM$(CODEXT$(CI%)) + " - " + DESCRIT0$(CI%)
               If DERACCE%("RECSOBRE/NOMESS", "Recepción con Sobre-Embarque") >= 2 Then
                   If COMALTER%("ATENCIóN: Sobre-Embarque en Código\" + DECODE$ + ".\\Corregir\Continuar") <> 2 Then GoTo 1113
                 Else
                   Call MENSERR(24, "Configuración y Derechos de Acceso\Impiden Sobre-Embarque\en Código " + DECODE$ + ".")
                   GoTo 1113
               End If
             End If
           End If
         End If
1114   Next YX&
       '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 27 = " & Now & vbCrLf
       If STMOVI09.List1.ListCount > 0 Then
         If COMALTER%("Se Han Detectado Items\ Con Estado de Movimiento de Baja \\Ver Detalle\Continuar") = 1 Then
            STMOVI09.Label2 = ""
            STMOVI09.Show 1 ' MUESTRA DETALLE
            If STMOVI09.Label2 = "MODIFICAR" Then ' SI ELIGIO MODIFICAR VUELVE A VENTANA
              GoTo 1113
            End If
         End If
       End If
       Call SETULTREG("!OBSERVOF", 0)
       Call CIERRARCH("!OBSERVOF")
       '
       ' VERIFICA SI EXISTEN RECEPCIONES DE LOS COMPONENTES CONSIGNADOS DE LOS ARTICULOS CON O/C QUE SE ESTAN RECIBIENDO
       If Control("BOLRECEP", "VEBRECON") = "SI" Then
            If ECOARCH$("VEBRECON", MKI$(NC%)) <> "" Then
                If ComprasPendientesDeLoConsignado%(CLng(NC%), ARCHICA$) = 1 Then
                     Respuesta% = COMALTER("¿Desea Continuar con la Recepción de los Materiales?\\Si, Continuar\No, Cancelar")
                     If Respuesta% = 2 Then
                         X$ = "": Y$ = "": YY$ = ""
                         EBOL$ = "": PEDY$ = "": DESCRI$ = ""
                         Screen.MousePointer = 1
                         GoTo 1100
                     End If
                     Screen.MousePointer = 11
                End If
            End If
       End If
       '
       Call SETULTREG("!OBSERVOF", 0)
       Call CIERRARCH("!OBSERVOF")
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 28 = " & Now & vbCrLf
       VTB% = VENTABU%("OBSERVOF/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(rasocli$(NC1%)) + " - " + FECHAREC$)
       If VTB% < 0 Then
         If TIMATEI% > 1 And InStr(1, EMPREAC$, "CONTEPLAST") > 0 Or InStr(1, EMPREAC$, "CAMBACERES") > 0 Then GoTo 1125
         If COMALTER%("Realmente Desea Cancelar ?\\Continuar\Cancelar") = 2 Then
           X$ = "": Y$ = "": YY$ = ""
           EBOL$ = "": PEDY$ = "": DESCRI$ = ""
           GoTo 1100
         End If
       End If
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 29 = " & Now & vbCrLf
       
       '
1125   Call ABRESTOCKS
       Call BLOQARCH("BOLRECEP")
       TINI = Timer
51     If Timer < TINI + 2 Then
         GoTo 51
       End If
       Call BLOQARCH("BOLRECEP")
       '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 30 = " & Now & vbCrLf
       SQLX$ = " SELECT max(NumeBoRe) as maximo from BOLRECEP"
       Dim BolreTemp_1 As ADODB.Recordset
       Set BolreTemp_1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
       If Not IsNull(BolreTemp_1!Maximo) Then
           NROBOL& = 1 + (BolreTemp_1!Maximo)
       End If
       BolreTemp_1.Close
       Set BolreTemp_1 = Nothing
       '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 31 = " & Now & vbCrLf
       Call CreaRegistro("BOLRECEP", "NumeBoRe", NROBOL&, "NOMESS")
       Call ACTUREGISTRO("BOLRECEP", "STATUS", "NumeBoRe=" & NROBOL&, 999, REG&, "NOMESS")
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 32 = " & Now & vbCrLf
       TINI = Timer
52     If Timer < TINI + 1 Then
         GoTo 52
       End If
       Call LIBARCH("BOLRECEP")
       '
       Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
       Call ABRESTOCKS
       Call ABRECOMPRAS
       '
1126   AINDI% = 0
       NORVENTA = ""
       NUOR% = 0
       ENVIAMAILCALIDAD% = 0
       ' PARA GABAL
       Call COPYSTRU("BOLREDET", "BOLREDEG")
       Call SETULTREG("BOLREDEG", 0)
       GABALMODOOF% = 0
       ExigeCalidadCompleto = False
       If InStr(1, EMPREAC$, "ASEMBLI") > 0 Then
            ExigeCalidadCompleto = True
            RutaGuardaCaptura$ = TRIM$(Control("", "RUTARGRA"))
            If RutaGuardaCaptura$ = "" Then RutaGuardaCaptura$ = LUDAT$
       End If
       '
       'VECTORES PARA ACUMULAR LA CANTIDAD SEGUN ORDEN DE FAB. P/LA TRANSFERENCIA DE COMPONENTES DE MEDITEA.
       Dim NUMEORFA$(), CANTIX1%()
       ReDim NUMEORFA$(0), CANTIX1%(0)
       ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 32 = " & Now & vbCrLf
       '
       '''PERGAMINO
       'BLOQUE PARA QUE PUEDA CARGAR PERGAMINO DE UNA CANTIDAD COMPRADA VARIAS FILAS Y QUE LAS DESCUENTE DE LA MISMA O.C.
       'EJ: SI DE UNA BOBINA COMPRO 3000 MTS DEBERA CARGAR EN LA PRIMERA QUE FIGURA CON LA O.C 1000 Y AGREGAR 2 FILAS MAS CON 1000 CADA UNO CON EL MISMO CODIGO
       If TRIM$(OrdenesSeleccionadas$) <> "" Then
         If InStr(1, TRIM$(UCase$(EMPREAC$)), "PERGAMINO") > 0 Then
            Dim ORDENXCODIGO(32767)
            NumeroCompra1& = 0
            For YX1& = 1 To ULTREG&("!" + ARCHICA$)
                 YY$ = REGLEIDO$("!" + ARCHICA$, YX1&)
                 CI% = CVI(Mid$(YY$, 83, 2))
                 CAN# = CVD(Mid$(YY$, 89, 8))
                 NUOC& = Abs(CVS(Mid$(YY$, 151, 4)))
                 NUMOIT% = Asc(Mid$(YY$, 98, 1))
                 If NUOC& > 0 And ORDENXCODIGO(CI%) < 1 Then
                    ORDENXCODIGO(CI%) = NUMOIT% 'GUARDO EL NRO DE ORDEN DEL ITEM PARA CADA CODIGO (SOLO LO HARA CON LOS PRIMEROS QUE DEBEN TENER ORDEN DE COMPRA
                 End If
                 If NumeroCompra1& < 1 Then
                    NumeroCompra1& = NUOC&
                 End If
                 If NUOC& < 1 Then
                    Call REPLA(YY$, Chr$(XVALO(ORDENXCODIGO(CI%))), 98, 1)
                    Call REPLA(YY$, MKS$(NumeroCompra1&), 151, 4)
                    Call GRAREG("!" + ARCHICA$, YY$, YX1&)
                 End If
            Next YX1&
         End If
       End If
       ''PERGAMINO
       
       
       For YX& = 1 To ULTREG&("!" + ARCHICA$)
            YY$ = REGLEIDO$("!" + ARCHICA$, YX&)
            CI% = CVI(Mid$(YY$, 83, 2))
            CAN# = CVD(Mid$(YY$, 89, 8))
            UNIDS& = CVS(Mid$(YY$, 209, 4))
                If UNIDS& < 1 Then UNIDS& = 1
            ANCHOX# = CVD(Mid$(YY$, 193, 8))
            LARGOX# = CVD(Mid$(YY$, 201, 8))
            Cantidad_1# = CVS(Mid$(YY$, 213, 4))
            Cantidad_2# = CVS(Mid$(YY$, 217, 4))
            cantidadBultos% = CVI(Mid$(YY$, 221, 2))
            If cantidadBultos% = 8224 Then cantidadBultos% = 1
            NUOC& = Abs(CVS(Mid$(YY$, 151, 4)))
            NUMOIT% = Asc(Mid$(YY$, 98, 1))   'EL Nº DE ORDEN DEL ITEM
            Peso# = 0: If (TIMATEI% > 1 Or (ES_SCRAP% > 0 And TIMATEI% = 1)) Then Peso# = CVS(Mid$(YY$, 253, 4))
            If CI% > 0 Then
              If CAN# > 0.005 Then
                CARECI = CAN# * UNIDS&
                LOTECO$ = TRIM$(Mid$(YY$, 107, 12))
                NUORSE& = 0
                '
                If CVD(Mid$(YY$, 163, 8)) >= 0.05 Then   ' solo si es comprado - no M.P. devuelto
                   NATEN& = CVS(Mid$(YY$, 100, 4))       ' NUMERO DE AUTORIZACION DE ENTREGAS
                   ' ACTUALIZA ORDENES DE COMPRA PENDIENTES
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 33 = " & Now & vbCrLf

                   Call ACTUCOMPEN(NUOC&, NUMOIT%, CARECI, NANTEN&)
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 34 = " & Now & vbCrLf
                   
                   'TRAIGO EL NUMERO DE ORDEN DE SERVICIO
                   CONDI2$ = "Nume_Ocom = " & NUOC&
                   CONDI2$ = CONDI2$ + " AND NuOrd_Item = " & NUMOIT%
                   CONDI2$ = CONDI2$ + " AND Stat_Ocom < 8 "
                   NUORSE& = XVALO(VALOBADA("OCOMDETA", "N_OrServ", CONDI2$, "NOMESS"))
                Else
                   If (InStr(1, EMPREAC$, "CONTEPLAST") > 0 Or InStr(1, EMPREAC$, "PERGAMINO") > 0) And NUOC& > 0 Then
                      Call ACTUCOMPEN(NUOC&, NUMOIT%, CARECI, NANTEN&)
                   Else 'SI NO ES CONTEPLAST O ES PERO NO TIENE NUM O.C. VIENE ACA COMO LO HACE SIEMPRE
                      Call COMUNI("Material '" & CODEXT$(CI%) & "' Sin Orden de Compra")
                   End If
                End If
                '
                ' SE AGREGA POR ERROR 6 QUE SE PRODUCE EN AHP
                ' *******************************************
                On Error GoTo ControlError
                '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 34_1 = " & Now & vbCrLf
                ' aca arma un string con las ordenes de compra recibidas
                NORVE& = Abs(CVS(Mid$(YY$, 151, 4)))
                If NORVE& > 0 Then
                  If NORVE& <> NORVEA& Then
                    NORVEA& = NORVE&
                    If NORVENTA <> "" Then NORVENTA = NORVENTA + " - "
                    NORVENTA = NORVENTA + TRIM$(Str$(NORVE&))
                  End If
                End If
                '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 34_2 = " & Now & vbCrLf
                NUMSE$ = "" ' ESCRIBO NUMERO DE SERIE COMO REFERENCIA
                If ESTRAZABLE(CI%) > 1 Then
                   NUMSE$ = "S-Nº: " + TRIM$(Mid$(YY$, 107, 12)) + " - "
                End If
                
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 34_3 = " & Now & vbCrLf
                DESCRI$ = TRIM$(CODEXT$(CI%)) + " - " + NUMSE$ + TRIM$(DESCRIT$(CI%))
                If AINDI% = 0 Then
                   X$ = REGBLAN$("INDIBOL")
                   Call REPLA(X$, MKS$(NROBOL&), 1, 4)
                   FEX = FF%
                   Call REPLA(X$, FECHATEX$(FEX), 5, 8)
                   Call REPLA(X$, REMI$, 15, 16)
                   Call REPLA(X$, ECOARCH$("PROVED1", MKI$(NC%)), 32, 29)
                   Call REGAPP("INDIBOL", X$)
                   If NUORSE& < 1 Then Call REGAPP("IBOLPEN", X$)
                   AINDI% = 1
                End If
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 34_4 = " & Now & vbCrLf
                '
                Call CIERRARCH("INDIBOL")
                Call CIERRARCH("IBOLPEN")
                '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 34_5 = " & Now & vbCrLf
                Call REPLA(YY$, MKS$(NROBOL&), 1, 4)
                If TIMATEI% < 2 Then Call REPLA(YY$, DESCRI$, 5, 64)
                '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 34_6 = " & Now & vbCrLf
                If TRIM$(Mid$(YY$, 69, 12)) = "" Then
                  NBR$ = TRIM$(Str$(NROBOL&))
                  NUOR% = NUOR% + 1 ' AGREGADO PARA QUE GUARDE EL ORDEN DE LAS QUE SE RECIBEN
                  IDENLO$ = "BR-" + NBR$ + "-" + TRIM$(Str$(NUOR%))
                  Call REPLA(YY$, IDENLO$, 69, 12)
                End If
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 34_7 = " & Now & vbCrLf
                Call REPLA(YY$, MKI$(FF%), 81, 2)
                DEPX% = DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(CI%)
                Call REPLA(YY$, Chr$(DEPX%), 97, 1)
                Call REPLA(YY$, MKI$(NC%), 105, 2)
                Call REPLA(YY$, REMI$, 119, 16)
                If CVD(Mid$(YY$, 163, 8)) >= 0.05 Then   ' solo si es comprado - no M.P. devuelto
                  If InStr(1, EMPREAC$, "ASEMBLI") < 1 Then
                    Call REPLA(YY$, MKS$(NUORSE&), 189, 4)
                  End If
                End If
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 35 = " & Now & vbCrLf
                '
                Call GRAREG("!" + ARCHICA$, YY$, YX&)
                '
                On Error GoTo 0
                ' FIN CONTROL DE ERROR 6 EN AHP
                ' *****************************
                '
                strSQL$ = " Select * from Bolrecep"
                strSQL$ = strSQL$ + " where NumeBoRe=" & NROBOL&
                Dim Bolrecep As ADODB.Recordset
                Set Bolrecep = New ADODB.Recordset
25              On Error Resume Next
                Bolrecep.Open FILTSQL$(strSQL$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                   On Error GoTo 0
                   Call CapturaErrorOpen
                   GoTo 25
                End If
                On Error GoTo 0
                '
                
                '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 36 = " & Now & vbCrLf

                With Bolrecep
                 For IUNI% = 1 To UNIDS&
                   NBR$ = TRIM$(Str$(NROBOL&))
                   NORDE% = NORDE% + 1
                   NORDES$ = TRIM$(Str$(NORDE%))
                   IDENLO$ = "BR-" + NBR$ + "-" + NORDES$
                   '
                   .AddNew   ' YA CONVERTIDO
                   On Error Resume Next
                   !CODIEMPR = CodiEmp%
                   On Error GoTo 0
                    !NUMEBORE = NROBOL&
                    !Descrip = Left$(DESCRI$, 64)
                    !idenlote = Left$(IDENLO$, 12)
                    !FechRecep = CVDAT(FF%)
                    !FEYHORecep = Now
                    !cod_inte = CI%
                    !Cod_Exte = Left$(CODEXT$(CI%), 24)
                    !descripitem = Left$(DESCRIT0$(CI%), 64)
                    !Unidad = Left$(Unimed$(CI%), 4)
                    !CantNom = CAN#
                    !M_Ancho = ANCHOX#
                    !M_Largo = LARGOX#
                    !DEPOENT = DEPX%
                    !NumProv = NC%
                    !Raso_Prov = Left$(rasocli$((-1) * NC%), 64)
                       LOTECO$ = TRIM$(LOTECO$)
                       While Len(LOTECO$) > 12
                         LOTECO$ = Mid$(LOTECO$, 2)
                       Wend
                    !LOTECOL_PROV = Left$(LOTECO$, 12)
                    !NumConVig = Left$(NUCONVI$, 16)
                    !NumRem_Prov = Left$(REMI$, 16)
                    !NumFac_Prov = ""
                    !patentetransporte = patentetransporte$
                    !NUMOC = NUOC&
                    !NroItemOC = NUMOIT% ' GUARDA EL NUMERO DE ITEM EN EL DETALLE DE LA O/COMPRA
                    !IContab = IMPCONOC% 'Imputación contable de la OC.
                    !MoneVaUni = ""
                    !Valo_Unit = 0 'ACA AGREGAR EL VALOR DE LA ORDEN DE COMPRA PARA AHP
                    !CanToApro = 0
                    !CanToRech = 0
                    !CanToFalta = 0
                    !FechLibe = CVDAT(0)
                    !NUMEORDSER = NUORSE&
                    !Observa = " " 'OBSE$
                    !Status = 0
                    !USUAPRO = " "
                    !FeHoApro = CVDAT(0)
                    If NUORSE& > 0 Then
                       !CanToApro = CAN#
                       !Status = 1
                       !USUAPRO = USERNAME$
                       !FeHoApro = CVDAT(HOY(HOS$))
                    End If
                    '\\\OT-06-0149-WAR-26/04/06///
                    !Nume_Aten = NATEN&
                    '\\\OT-06-0149-WAR-FIN///
                    !Via_Entrega = VIAENTRE_COMPRA(NUOC&)
                    !Peso = Peso# + Cantidad_1#
                    !kgbalanzapublica = Cantidad_2#
                    !cantidadBultos = cantidadBultos%
                   .Update
                 Next IUNI%
                End With
                Bolrecep.Close
                Set Bolrecep = Nothing
                '

TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 37 = " & Now & vbCrLf

               CANTUI& = CAN#
               If ESTRAZABLE(CI%) = 2 Then
                  LOTECO$ = TRIM$(Mid$(YY$, 107, 12))
                  If LOTECO$ <> "" Then
                     '** OJO ESTA CONDI SE UTILIZA TAMBIEN PARA LA OBTENCION DE LA OF MAS ABAJO**
                     CONDI$ = "Cod_Inte = " & CI% & " and NUMEROSERIE = '" & LOTECO$ & "'"
                     'QUITA NOMESS PARA QUE PRESENTE EL MENSAJE QUE CAUSA QUE NO SE ACTUALICE
10                   Call ACTUREGISTRO("TANUMSE", "UBICACION", CONDI$, "Planta Meditea", RAFE&)

                     'SE INCORPORA VERIFICACION PARA GARANTIZAR QUE SE ACTUALIZO LA TABLA Y SI REINTENTA
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 38 = " & Now & vbCrLf
                     
                     UBICA$ = TRIM$(SAFETEXT$(VALOBADA("TANUMSE", "UBICACION", CONDI$)))
                     '

                     ENVIAMAILCALIDAD% = 1
                     CUERPO$ = "Recepcion de " + rasocli$((-1) * NC%)
                     '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 39 = " & Now & vbCrLf
                     
                     Call ANOTANUSE(LOTECO$, CUERPO$)
                     '**
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 40 = " & Now & vbCrLf
                     
                     'ACUMULO EN UN VECTOR LA CANTIDAD QUE SE ESTA RECEPCIONANDO PARA CADA ORDEN DE FABRICACION.
                     NORFA1$ = VALOBADA("TANUMSE", "IdSubOr", CONDI$, "NOMESS")
                     '
                     ' VALIDANDO SI EL VECTOR YA TIENE LA OF. ACUMULO CANTIDAD
                     For J% = 1 To UBound(NUMEORFA$)
                       If NUMEORFA$(J%) = NORFA1$ Then
                          CANTIX1%(J%) = CANTIX1%(J%) + 1
                          GoTo 12 'SI ES MISMA ORDEN ACUMULA Y SALTA A 12
                       End If
                     Next J%
                     '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 41 = " & Now & vbCrLf
                     'SI ES UNA ORDEN QUE NO ENCONTRO EN EL VECTOR INCREMENTA EL INDICE CON UNA NUEVA OF
                     IND% = UBound(NUMEORFA$) + 1
                     ReDim Preserve NUMEORFA$(IND%)
                     ReDim Preserve CANTIX1%(IND%)
                     NUMEORFA$(IND%) = NORFA1$
                     CANTIX1%(IND%) = CANTIX1%(IND%) + 1
'                     Call TRANSFIERE_COMPONENTES(CI%, LOTECO$, "BR-" + TRIM$(Str$(NROBOL&)))
12                End If
               End If
               '
               If NUORSE& < 1 Then
                  MODO$ = "CC": If Left$(UCase$(LOTECO$), 2) = "MP" Then MODO$ = "DV"
                  Call DESCACONSI(NC%, CI%, CAN#, NUOC&, FF%, NROBOL&, IDENLO$, MODO$)
               End If
               '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 42 = " & Now & vbCrLf
               
               If NUORSE& > 0 Then
                  If CAN# > 0.05 Then
                    GoSub BLANTEROP
'                    ' GRABA LOTE DE RECEPCION A LA ETIQUETA LEIDA
'                    If InStr(1, EMPREAC$, "GABAL") > 0 Then
'                        GABALMODOOF% = 1
'                        NORFA$ = Mid$(YY$, 241, 2) & "-" & Mid$(YY$, 243, 6) & "-" & Mid$(YY$, 249, 1)
'                        NORDEN% = XVALO(Mid$(YY$, 251, 4))
'                        CONDI$ = "IDSUBOR='" & NORFA$ & "' AND NUORDEN=" & NORDEN%
'                        Call ACTUREGISTRO("ORFAETI", "LOTERECE", CONDI$, Left$(IDENLO$, 12), REGAF&)
'                        '
'                        'AGRUPA PARA QUE NO REPITA EL MISMO CODIGO POR CADA ETIQUETA. PARA ESO SE UTILIZA OTRO ARCHIVO
'                        '(BOLREDEG) QUE SE GENERA CON LOS DATO DE BOLREDET PERO ACUMULANDO CUANDO SE REPITEN LOS
'                        'CODIGOS. ESTO FUNCIONA CUANDO SE LEE ETIQUETAS DE OF COMO CODIGOS DE BARRA UPC
'                        finbr& = ULTREG&("BOLREDEG")
'                        ENCONTRO% = 0: JJ& = 0
'                        For II& = 1 To finbr&
'                            YYX$ = REGLEIDO$("BOLREDEG", II&)
'                            CIXI% = CVI(Mid$(YYX$, 83, 2))
'                            If CIXI% = CI% Then
'                                CANBR# = CVD(Mid$(YYX$, 89, 8))
'                                CANBR# = CANBR# + CAN#
'                                Call REPLA(YYX$, MKD$(CANBR#), 89, 8)
'                                Call GRAREG("BOLREDEG", YYX$, II&)
'                                ENCONTRO% = 1
'                            End If
'                        Next II&
'                        If ENCONTRO% = 0 Then
'                            XXY$ = YY$: Call REPLA(XXY$, Space(12), 107, 12)
'                            Call REGAPP("BOLREDEG", XXY$)
'                        End If
'                    End If
                  End If
               End If
               
               If MODOGABAL% = 2 Then
                    Dim RST_CIERREARMADOR As ADODB.Recordset
                    Set RST_CIERREARMADOR = New ADODB.Recordset
                    SQLX$ = "SELECT * FROM MOVISTO WHERE COD_INTE=" & Str$(GB_CIXI%) & " AND DOPRICOR='" & OFRECIBI$ & "' AND (CANTINGRE-" & TRIM$(FORMATNUM$(GB_CANTIDAD#, "F10.1")) & ")< 0.05"
                    RST_CIERREARMADOR.Open SQLX$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                      With RST_CIERREARMADOR
                        Do While Not .EOF
                            !TIPOTRAN = 3
                            .Update
                            .MoveNext
                        Loop
                      End With
               End If
               
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 43 = " & Now & vbCrLf
               '
               ' DETERMINA UBICACION PARA GABAL SOLO PARA MATERIAS PRIMAS
               ' OBVIA LOS TRABAJOS DE ARMADORES NUORSE& > 0
'               If InStr(1, EMPREAC$, "GABAL") > 0 And NUORSE& < 1 Then
'                  UB$ = UBISTONUE$(CI%, CAN#)
'                  If UB$ <> "" Then Call GRABAUBICACION(CI%, CAN#, UB$, IDENLO$)
'                  UB$ = AJUSTI$(UB$, 16)
'                  Call REPLA(YY$, UB$, 193, 16)
'                  Call GRAREG("!" + ARCHICA$, YY$, YX&)
'               End If
               '
               
                ' *** AGREGA IMAGEN DEL MATERIAL RECIBIDO CAPTURANDO LA MISMA VIA WEBCAM
'                If ExigeCalidadCompleto Then
'                    ArchivoOriginal$ = RutaGuardaCaptura$ & "\" & TRIM$(Mid$(YY$, 189, 4)) & ".JPG"
'                    ArchivoRenombrado$ = RutaGuardaCaptura$ & "\" & IDENLO$ & ".JPG"
'                    If FileExists(ArchivoOriginal$) = True Then
'                        Name ArchivoOriginal$ As ArchivoRenombrado$
'                        Call GuardaFotoMaterialRecibido(IDENLO$)
'                    End If
'                End If
                '
              End If
            End If
       Next YX&
       '
       Call BORRAREGISTROS("BOLRECEP", "STATUS=999", REG&, "NOMESS")
       
       If idRecepcion& > 0 Then
            Call ACTUREGISTRO("BOLRETMP", "status", "idRecepcion=" & CStr(idRecepcion), 9, REGAF&)
            
            Call Recepcion_AdjuntarDocumentosExternos(NROBOL&)
       End If
        
       '***********ACA GRABAR EL CONTEXTO DE LA RECEPCION *********************
       If XVALO(Control("BOLRECEP", "RECECRON")) > 0 Then 'CONTROL CRONOLOGICO SEGUN FECHA DE ENTREGA DE ORDEN DE COMPRA
           Call GRABAR_CONTEXTO_PROPUESTO(NC, NROBOL&, FF%)
           Unload VENTABNEW
       End If
       '************FIN GRABAR EL CONTEXTO DE LA RECEPCION *********************
       
       '************ PRODUCTOS IMPORTADOS  *********************
       If PIVACLI%(NC% * (-1)) = 5 Then
           If XVALO(Control("BOLRECEP", "RECEIMPO")) > 0 Then
             Call AGREGA_A_TABLA_IMPORTADO("!" + ARCHICA$) 'GRABACION AUTOMATICA A TABLA DE IMPORTADOS
             ReceDespa.Text5 = NROBOL&
             ReceDespa.Text1 = FECHAREC$
             ReceDespa.Text3 = Abs(NC1%)
             ReceDespa.Text4 = rasocli$(NC1%)
             ReceDespa.Show 1 ' OPCION DE INFORMAR DATOS DE DESPACHO DE IMPORTACION.
           End If
         End If
        '************ FIN PRODUCTOS IMPORTADOS  *********************
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 44 = " & Now & vbCrLf
       
       '
       'TRANSFERENCIA DE COMPONENTES PARA CADA ORDEN DE FABRICACION, CON LA CANTIDAD ACUMULADA DE EQUIPOS POR OF
       
       For J% = 1 To UBound(NUMEORFA$)
          NORFA1$ = NUMEORFA$(J%)
          CANTIX% = XVALO(CANTIX1%(J%))
          'Call TRANSFIERE_COMPONENTES(CI%, LOTECO$, "BR-" + TRIM$(Str$(NROBOL&)))
          Call TRANSFIERE_COMPONENTES(CODIFAB%(NORFA1$), NORFA1$, "BR-" + TRIM$(Str$(NROBOL&)), CANTIX%)
       Next J%
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 45 = " & Now & vbCrLf
       'FIN TRANSFERENCIA DE COMPONENTES
       '
       Call AJUMACON
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 46 = " & Now & vbCrLf
       '
       If InStr(EMPREAC$, "MEDITEA") > 0 Then
            Call ALERTATRANSFE
       End If
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 47 = " & Now & vbCrLf
       '
       '
       If TRIM$(FORIPRE$) <> "" Then
         '
         CANCELPRINT% = 0
         If Control("BOLRECEP", "SUPRINT") = "SI" Then
            CANCELPRINT% = 1
         Else
            'VALIDACION DE MODO DE IMPRESION
            MODOIMPRE07.Check7.Caption = "Imprime Recepción"
            MODOIMPRE07.Label1.Caption = "" 'LIMPIA EL LABEL DONDE PONE OK, SI ACEPTA
            MODOIMPRE07.Check8.Value = 0
            Screen.MousePointer = 1
            MODOIMPRE07.Show 1
            If MODOIMPRE07.Label1.Caption = "OK" Then
             If MODOIMPRE07.Check8.Value > 0 Then CANCELPRINT% = 1
            End If
            
            'AGREGADO PARA EYM POR QUE LUEGO DE IMPRIMIR UNAS ETIQUETAS CUANDO VOLVIA A RECEPCIONAR LE SALIAN POR LA MAQUINA DE ETIQUETAS
            'SIENDO QUE EN DICHA IMPRESION REALIZABA UN RESTORE-
            If InStr(1, EMPREAC$, "EYM") > 0 Then Call SETPRINPORT("RESTORE")
            
            Screen.MousePointer = 11
         End If
         '
         TIPODOC$ = "Boleta de Recepción"
         DESTIDOC% = 0 'SE AGREGA POR QUE LUEGO QUE SE HACIA UN RECHAZO y DESPUES UNA RECEPCION MANTENIA EL NRO DE PROVEEDOR.

         DESTIDOCNUE = (-1) * Abs(NC1%)
         NUMEDOC$ = Str$(NROBOL&)
         FEX = FF%: FECHDOC$ = FECHATEX$(FEX)
         '
         FACTU$ = TRIM$(Mid$(PEDY$, 129, 16)) ' ASIGNA NUMERO DE FACTURA QUE TOMA DEL OBJETO
         '
         CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = DEPOSAL$
         CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = REMI$
         CODPARAM$(3) = "N-AUTENT": DOCPARAM$(3) = NORVENTA
         CODPARAM$(4) = "NRO-PASE": DOCPARAM$(4) = NUCONVI$
         CODPARAM$(5) = "DOCUORIG": DOCPARAM$(5) = FACTU$
         CODPARAM$(6) = "PATE-TRA": DOCPARAM$(6) = patentetransporte$
         '
         CAPARDOC% = 6
         '
         CODTABU1$(1) = "COD-MAT"
         CODTABU1$(2) = "DES-MAT"
         CODTABU1$(3) = "COLOR"
         CODTABU1$(4) = "UNIMED"
         CODTABU1$(5) = "NRO-LOTE"
         CODTABU1$(6) = "CANTIDAD"
         CODTABU1$(7) = "N-OCOMPR"
         CODTABU1$(8) = "CAKILOS"
         CODTABU1$(9) = "F-PLANO"
         CODTABU1$(10) = "FECH-REC"
         CODTABU1$(11) = "ORD-ITEM"
         CODTABU1$(12) = "UBI-MPRI"
         CODTABU1$(13) = "CANT-DIS"
         CODTABU1$(14) = "DESTINO"
         CODTABU1$(15) = "F-PLANO"
         CODTABU1$(16) = "CAN-PLCO"
         CODTABU1$(17) = "NUBULTOS"
         
         CACOLTAB1% = 17
         '
         Call LIMPIATETAB
         '
         CAITAB1% = 0
         ORDITEM% = 0
         '
         'CONTROLES DE IMPRESION DE DESCRIPCION
         DESCPRI% = 0: DESCOMP% = 0: Especif% = 0
         If Control("BOLRECEP", "DESCPRIM") = "SI" Then DESCPRI% = 1
         If Control("BOLRECEP", "DESCOMPL") = "SI" Then DESCOMP% = 1
         If Control("BOLRECEP", "ESPECIFI") = "SI" Then Especif% = 1
         If DESCOMP% + Especif% < 1 Then DESCPRI% = 1 'POR SI NO SE SELECCIONO NADA
         '
'         ' UTILIZA EL ARCHIVO BOLREDEG GENERADO PARA EL CASO DE GABAL. VER MAS ARRIBA
'         If GABALMODOOF% = 1 Then
'            Call TRALOCAL("BOLREDEG", "")
'            ARCHICA$ = "BOLREDEG"
'         End If
         '
         For YX& = 1 To ULTREG&("!" + ARCHICA$)
           YY$ = REGLEIDO$("!" + ARCHICA$, YX&)
           UNIDS& = CVS(Mid$(YY$, 209, 4))          'Lee del archivo la cantidad de unidades
           CABOBIN& = UNIDS&: If CABOBIN& < 1 Then CABOBIN& = 1  'Asigna a cabobin el total de unidades, si es "" asigna 1
           CI% = CVI(Mid$(YY$, 83, 2))
           CAN# = CVD(Mid$(YY$, 89, 8)) * CABOBIN&
           If CI% > 0 Then
             If CI% <= NUMAS% Then
               If CAN# > 0.005 Then
                 '
                 UNID$ = TRIM$(Mid$(YY$, 85, 4))
                 If UNID$ = "" Then UNID$ = Unimed$(CI%)
                 CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
                 KILAJE$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 10, 3, 2)
                    If TIMATEI% > 1 Then KILAJE$ = TRIM$(FORMATNUM$(CVS(Mid$(YY$, 253, 4)), "F10.2")) ' PESO BOBINA
                    If TIMATEI% = 1 Then KILAJE$ = TRIM$(FORMATNUM$(CVS(Mid$(YY$, 213, 4)), "F10.2")) ' CANTIDAD_1
                 FEX = CVI(Mid$(YY$, 155, 2))
                 FESOLEN$ = FECHATEX$(FEX)
                 FEPLA$ = FEPLANO$(CI%)
                 NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
                 ' SON EQUIPOS CON NUMEROS DE SERIE Y SE INGRESA COMO CANTIDAD <> 1 LE ASIGNA 1
                 If ESTRAZABLE%(CI%) > 1 And XVALO(CANS$) <> 1 And NULOTE$ <> "" Then
                    CANS$ = "1"
                 End If
                 '
                 If NULOTE$ = "" Then
                    NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
                 End If
                 NUOCO$ = FORMATNUM$(Abs(CVS(Mid$(YY$, 151, 4))), "F7,0")
                 '
                 cantidadBultos% = CVI(Mid$(YY$, 221, 2))
                 '
                 CAITAB1% = CAITAB1% + 1
                 TETABU1$(1, CAITAB1%) = TRIM$(CODEXT$(CI%))
                 '
                 DESPRIN$ = DESCRIT$(CI%)
                 DESCOMPLEM$ = DESCRIT2$(CI%)
                 ESPECIFICAC$ = ESPECIFIT$(CI%)
                 '
                 If FORMPLA$ = "RECEP-ESPUMA" Then
                       DESPRIN$ = Mid$(YY$, 5, 64)
                 End If

                 DIMPRE$ = ""
                 If DESCPRI% > 0 Then
                   If DESPRIN$ <> "" Then
                      DIMPRE$ = DESPRIN$
                   End If
                 End If
                 '
                 'DESCRIPCION COMPLEMENTARIA
                 If DESCOMP% > 0 Then
                   If DESCOMPLEM$ <> "" Then
                     If DIMPRE$ <> "" Then
                       DIMPRE$ = DIMPRE$ + vbCrLf
                     End If
                     DIMPRE$ = DIMPRE$ + DESCOMPLEM$
                   End If
                 End If
                 '
                 'ESPECIFICACION
                 If Especif% > 0 Then
                   If ESPECIFICAC$ <> "" Then
                     If DIMPRE$ <> "" Then
                       DIMPRE$ = DIMPRE$ + vbCrLf
                     End If
                     DIMPRE$ = DIMPRE$ + ESPECIFICAC$
                   End If
                 End If

                 ' KG BALANZA PUBLICA O CANTIDAD_2
                 Cantidad_2# = CVS(Mid$(YY$, 217, 4))
                 '
                 ORDITEM% = ORDITEM% + 1
'                 TETABU1$(2, CAITAB1%) = DIMPRE$
                 TETABU1$(3, CAITAB1%) = COLOR$(CI%)
                 TETABU1$(4, CAITAB1%) = UNID$
                 TETABU1$(5, CAITAB1%) = NULOTE$
                 TETABU1$(6, CAITAB1%) = CANS$
                 TETABU1$(7, CAITAB1%) = NUOCO$
                 TETABU1$(8, CAITAB1%) = KILAJE$
                 TETABU1$(9, CAITAB1%) = FEPLA$
                 TETABU1$(10, CAITAB1%) = FESOLEN$
                 TETABU1$(11, CAITAB1%) = ORDITEM%
                 TETABU1$(12, CAITAB1%) = UBISTO$(CI%)
                    DEPX% = DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(CI%)
                 TETABU1$(14, CAITAB1%) = CStr(DEPX%)
                 TETABU1$(15, CAITAB1%) = REVICOD$(CI%) ' REVISION
                 TETABU1$(16, CAITAB1%) = TRIM$(FORMATNUM$(Cantidad_2#, "F10.1"))
                 TETABU1$(17, CAITAB1%) = TRIM$(FORMATNUM$(cantidadBultos%, "I5"))
                 'TETABU1$(13, CAITAB1%) = STOCKDIS(CI%) ' *********** SE COMENTA POR VELOCIDAD
                 ' IMPRIME UBICACION
'                 If InStr(1, EMPREAC$, "GABAL") > 0 Then
'                    UB$ = TRIM$(Mid$(YY$, 193, 16))
'                    TETABU1$(12, CAITAB1%) = UB$
'                 End If
                 '
                 'IMPRESION DE LA DESCRIPCION CON TODAS SUS VARIANTES CONFIGURADAS
                 If Control("BOLRECEP", "DESCUNRE") = "SI" Then
                    TETABU1$(2, CAITAB1%) = DIMPRE$
                    GoTo 48
                 End If
                 FORTELIB$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")))
                 ATELIB% = XVALO(Mid$(FORTELIB$, 2))
                 Call FRATEXTO(DIMPRE$, ATELIB%)
                 For KKU1% = 1 To CARETEX%
                      TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
                      CAITAB1% = CAITAB1% + 1
                 Next KKU1%
                 '
48               If InStr(1, EMPREAC$, "DOSIVAC") > 0 Then
                   CONDI$ = "Nume_Ocom = " & NUOCO$
                   CONDI$ = CONDI$ + " AND cod_inte = " & CI%
                   CONDI$ = CONDI$ + " AND Stat_Ocom < 8 "
                   NUORSE& = XVALO(VALOBADA("OCOMDETA", "N_OrServ", CONDI$, "NOMESS"))
                   '
                   If NUORSE& > 0 Then
                        SQLX$ = "SELECT NUMEOPER, DESCOPER,  FROM SECOPER WITH(NOLOCK) WHERE CODICONJ=" & CI%
                        Dim Rst As ADODB.Recordset
                        Set Rst = READSET(SQLX$)
                        With Rst
                            Do While Not .EOF
                                NOP% = XVALO(!NUMEOPER)
                                DESCOP$ = SAFETEXT$(!descoper)
                                TXTX$ = Str$(NOP%) & " ... " & DESCOP$
                                CAITAB1% = CAITAB1% + 1
                                TETABU1$(2, CAITAB1%) = TXTX$
                                .MoveNext
                            Loop
                        End With
                        On Error Resume Next
                        Rst.Close
                        Set Rst = Nothing
                        On Error GoTo 0
                   End If
                 End If
                 '
                 'LINEA DE SEPARACION
                 Call CARLINSEP
                 '
                 If TICOMPRO$ = "FICHA" Then
                   Call PRINTDOC("FORMBREC")
                 End If
                 '
               End If
             End If
           End If
         Next YX&
         '
 TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 48 = " & Now & vbCrLf
         If TICOMPRO$ <> "FICHA" Then
           Call PRINTDOC("FORMBREC")
         End If
         '
TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 49 = " & Now & vbCrLf
         If ENVIAMAILCALIDAD% > 0 Then
            'ENVIO DE MAIL DE AVISO DE RECEPCION DE EQUIPOS CON NUMERO DE SERIE
            EMAI$ = Control("OPVALEMA", "CONTMAIL")
            If EMAI$ = "" Or InStr(EMAI$, "@") < 1 Then
               GoTo 11999
            End If
 TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 50 = " & Now & vbCrLf
           '
            SENDBYMAIL$ = EMAI$
            For Each PRX1 In Printers
             If InStr(UCase$(PRX1.DeviceName), "PDFWRITER") > 0 Then GoTo 50
            Next
            Call MENSERR(24, "Imposible Enviar por Mail.\   \Falta Impresora Virtual PDFWRITER.")
            SENDBYMAIL = ""
            GoTo 11999
50          Call PRINTDOC("FORMBREC")
         End If
         
 TXTRAZA_BR$ = TXTRAZA_BR$ & "PUNTO 51 = " & Now & vbCrLf
        '
        
         If Control("BOLRECEP", "IMETIREC") = "SI" Then
            Call ImprimeEtiqRecepción(NROBOL&)
         End If
         
         If TRIM$(Control$("", "STIKAPRO")) <> "" Then
            IncluyeQR = False
            If InStr(1, EMPREAC$, "CONTEPLAST") > 0 Or InStr(1, EMPREAC$, "RETEPLAST") > 0 Or InStr(1, EMPREAC$, "CAMBACERES") > 0 Then
                IncluyeQR = True
            End If
            Call ReimpresionEtiquetasParaBobinas(NROBOL&, True)
         End If
         'If InStr(1, EMPREAC$, "GABAL") > 0 Then Call IMPREETIQUETA(NROBOL&)
         '
       End If

     End If
     '
11999 Call CIERRARCH("*.*")
'
'GRABO EL ARCHIVO DE TRAZABILIDAD
'For I1% = 1 To 32767
'     ARCH$ = LUDAT$ & "TRAZABR" & TRIM$(Str$(I1%)) & ".TXT"
'     If EXISTE%(ARCH$) < 1 Then
'       NARCHI2% = FILEOPEN%(ARCH$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
'       Print #NARCHI2%, TXTRAZA_BR$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
'       Close #NARCHI2%
'       Exit For
'     End If
'Next I1%

     '
     X$ = "": Y$ = "": YY$ = ""
     EBOL$ = "": PEDY$ = "": DESCRI$ = ""
     '
     CANCELPRINT% = 0 'QUE NO CANCELE LA IMPRESIÓN DEL REMITO DE DEVOLUCIÓN \\\OT-05-0826-WAR-21/11/05///
     '
     GoTo 1100
     '
BLANTEROP:
   '
   FINTER& = ULTREG&("OPERTER")
   For U& = FINTER& To 1 Step -1
     X0$ = REGLEIDO$("OPERTER", U&)
     '///OT-8-667-04/09/08
     If CVS(Mid$(X0$, 75, 4)) = NUORSE& Then
       REGITER& = U&
       GoTo 21
     End If
   Next U&
   GoTo 29
   '
21 ' Marca como "Cumplida" Operación en OPERTER.DAT
   SALRE# = CVD(Mid$(X0$, 79, 8)) - CAN#
   If SALRE# < 0 Then SALRE# = 0
   '
   FEBA% = HOY(HOS$)
   HOBA% = Int(HORANUM(Time$))
   USUA% = USNBR% Mod 100
   DBAJA$ = MKI$(FEBA%) + MKI$(HOBA%) + Chr$(USUA%) + Chr$(1) + "B.Rec." + TRIM$(Str$(NROBOL&))
   '
   NUOROP% = Asc(Mid$(X0$, 67, 1))
   Call REPLA(X0$, MKD$(CAN#), 87, 8)
   Call REPLA(X0$, MKD$(SALRE#), 95, 8)
   Call REPLA(X0$, "3ROS  ", 151, 6)
   Call REPLA(X0$, MKI$(FEBA%), 157, 2)
   Call REPLA(X0$, MKI$(Int(HORANUM(Time$))), 159, 2)
   Call REPLA(X0$, MKS$(0), 161, 4)
   Call REPLA(X0$, String$(8, 0), 165, 8)
   Call REPLA(X0$, MKS$(0), 173, 4)
   Call REPLA(X0$, DBAJA$, 177, 22)
   Call GRAREG("OPERTER", X0$, REGITER&)
   '
' Si se recibio un parcial, genera registro en OPERTER.DAT
' por el saldo pendiente
   If SALRE# > 0.01 Then
      NUOROP% = 1 + Asc(Mid$(X0$, 67, 1))
      CIIK% = CVI(Mid$(X0$, 13, 2))
      NOPE% = CVI(Mid$(X0$, 15, 2))
      DALTA$ = Chr$(NUOROP%) + Mid$(DBAJA$, 2, 5)
      INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
      PROGBLA$ = Space$(6) + String$(26, 0)
      MINVARI = TIESTAN(CIIK%, NOPE%, SALRE#, NUOROP%)
      '
      Call REPLA(X0$, DALTA$, 67, 6)
      Call REPLA(X0$, MKD$(SALRE#), 79, 8)
      Call REPLA(X0$, MKD$(0), 87, 8)
      Call REPLA(X0$, MKD$(SALRE#), 95, 8)
      Call REPLA(X0$, String$(32, 0), 103, 32)
      Call REPLA(X0$, MKS$(MINVARI), 111, 4)
      Call REPLA(X0$, MKS$(MINVARI), 131, 4)
      Call REPLA(X0$, String$(16, 0), 135, 16)
      Call REPLA(X0$, INFOBLA$, 151, 48)
      Call REPLA(X0$, PROGBLA$, 225, 32)
      Call REGAPP("OPERTER", X0$)
      AGREGA% = 1
   End If
   '
29 Return
   '
   
ControlError:

   TXTRAZA_BR$ = TXTRAZA_BR$ & "Se Produjo el Error de AHP - " & Err.Description
   Return
   
End Sub


Sub CONTROL_CRONOLOGICO_COMPRAS(NPROV, ARCHICA$, APROBO%)
        APROBO% = 0
        BLOQUEA% = XVALO(Control("BOLRECEP", "CRONBLOQ")) 'SI BLOQUEA NO PERMITE EDITAR SOLO SE PUEDE RECIBIR COMO SE PROPONE
        Unload VENTABNEW
        VENTANA.Inicializar = 0

        If BLOQUEA% < 1 Then VENTANA.CampEditables = "6" 'SI NO BLOQUEA LO PONE EDITABLE
        If DERACCE%("MODIRESI/NOMESS", "MODIFICACIÓN DE RECEPCIÓN CRONOLÓGICA") >= 2 Then
           VENTANA.CampEditables = "6" 'SI TIENE DERECHOS LO PONE EDITABLE
        End If
        VENTANA.AgregaRegistro = 0
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        VENTANA.HabilitaInsertar = 0
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = "Propuesta Cronológica Según Fecha Solicitada de Orden de Compra "
        VENTANA.TOTALES(5) = "S"
        VENTANA.TOTALES(6) = "S"
        VENTANA.TOTALES(9) = "S"
        VENTANA.Inicializar = 1
        VENTANA.Campox = "CI/F0;CODIGO/A16;DESCRIPCION/A32;U.M./A4;Q.INGR/F10.1;Q.IDEAL/F10.1;O.COMPRA/F8;ENTR.SOL./D8;PENDIENTE/F10.1;Q.IDEAL/F10.1;LOTECO/A0;ORDEN/F0"
        
        '
        'CONTROLA SI TIENE SOBRE EMBARQUE
        BLOSOBRE% = 0
        PORCENTAJE_SOBRE = 0
        If Control$("BOLRECEP", "BLOSOBRE") = "SI" Then
          BLOSOBRE% = 1
          PORCENTAJE_SOBRE = XVALO(Control("BOLRECEP", "PORSOBRE"))
        End If

     
       i& = 0
       Dim CANTIXCODIGO(32767), CODILEIDO(32767)
       Dim CANTIXOCOM()
      '***************************************************************************
      ' - ACUMULO LA CANTIDAD DE CADA ARTICULO QUE ESTA CON CANTIDAD EN EL VENTABU
      '***************************************************************************
       For YX& = 1 To ULTREG&("!" + ARCHICA$)
         YY$ = REGLEIDO$("!" + ARCHICA$, YX&)
         CI1% = CVI(Mid$(YY$, 83, 2))
         
         CAN# = CVD(Mid$(YY$, 89, 8))
         If CAN# > 0.005 Then
           If ESTRAZABLE%(CI1%) < 2 Then
             UNIDS& = CVS(Mid$(YY$, 209, 4))
             If UNIDS& < 1 Then UNIDS& = 1
             NUMEOCOMPRA& = Abs(CVS(Mid$(YY$, 151, 4)))
             CANTIXCODIGO(CI1%) = CANTIXCODIGO(CI1%) + CAN#
           End If
         End If
       Next YX&
       '*********************************************************************************
       ' - FIN - ACUMULO LA CANTIDAD DE CADA ARTICULO QUE ESTA CON CANTIDAD EN EL VENTABU
       '*********************************************************************************
       '
       '
       For YX& = 1 To ULTREG&("!" + ARCHICA$)
         YY$ = REGLEIDO$("!" + ARCHICA$, YX&)
         CI1% = CVI(Mid$(YY$, 83, 2))
         
         CAN# = CVD(Mid$(YY$, 89, 8))
         If CAN# > 0.005 Then
           If ESTRAZABLE%(CI1%) < 2 Then
             ACUMULACAPEN = 0
             UNIDS& = CVS(Mid$(YY$, 209, 4))
             If UNIDS& < 1 Then UNIDS& = 1
             NUMEOCOMPRA& = Abs(CVS(Mid$(YY$, 151, 4)))
             On Error Resume Next
             If UBound(CANTIXOCOM) < NUMEOCOMPRA& Then ReDim Preserve CANTIXOCOM(NUMEOCOMPRA&)
             On Error GoTo 0
             CANTIXOCOM(NUMEOCOMPRA&) = CANTIXOCOM(NUMEOCOMPRA&) + CAN#

             'ESTE IF ES PARA AGREGAR EN LA FILA DE LA ORDEN DE COMPRA LA CANTIDAD QUE INFORMO EL USUARIO _
             'EJEMPLO SI SE INGRESO PARA UN MISMO ITEM CANTIDADES Y LA PROPUESTA PUEDE ACOMODAR TODO EN EN LA PRIMERA_
             'SOLO ESTABA MOSTRANDO COMO CANTIDAD INGRESADA DEL USUARIO LA PRIMERA P Q CUANDO PEGABA LA VUELTA
             'YA ESTABA CUMPLIDA Y NO LO PONIA
             If CANTIXOCOM(NUMEOCOMPRA&) > 0 And CANTIXCODIGO(CI1%) <= 0 Then
                For i& = 1 To VENTABNEW.MSF.Rows - 1
                   If VENTABNEW.MSF.TextMatrix(i&, 7) = NUMEOCOMPRA& And XVALO(VENTABNEW.MSF.TextMatrix(i&, 5)) <= 0 Then
                     VENTABNEW.MSF.TextMatrix(i&, 5) = FORMATNUM$(CANTIXOCOM(NUMEOCOMPRA&), "F12.1")
                     CANTIXOCOM(NUMEOCOMPRA&) = 0
                   End If
                Next i&
             End If
             '
             If CODILEIDO(CI1%) < 1 Then 'SI NO SE LEYO EL CODIGO AUN
                'CANTIDAD TOTAL PENDIENTE DE ESTE CODIGO
                CONDI$ = " Cod_Inte = " & CI1% & " AND Stat_Ocom < 8 "
                CONDI$ = CONDI$ + " AND Stat_Ocom <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
                CONDI$ = CONDI$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
                CONDI$ = CONDI$ + " AND NPRO_OCOM = " & NPROV

                CANCOMPORA = XVALO(VALOBADA("OCOMDETA", "Sum(cant_ocom)", CONDI$, "NOMESS"))
                CANREC = XVALO(VALOBADA("OCOMDETA", "Sum(Cant_Rec)", CONDI$, "NOMESS"))
                PEND_TOTAL_OCS = CANCOMPORA - CANREC
                '
                'If CANTIXCODIGO(CI1%) <= PEND_TOTAL_OCS Then
                If CANTIXCODIGO(CI1%) > 0 Then
                    ACUMULACAPEN = 0
                    CAN_ACUMULA_X_CODIGO = CANTIXCODIGO(CI1%)
                    APLICA_SOBEMB% = 0
                    SQLX$ = "Select cant_ocom, Cant_Rec,COD_EXTE , Nume_Ocom,DESCRITEM, Fentre_Sol, COD_INTE, "
                    SQLX$ = SQLX$ + " UNI_MAS,Doc_Orig,NuOrd_Item  "
                    SQLX$ = SQLX$ + " from OCOMDETA WITH(NOLOCK) "
                    SQLX$ = SQLX$ + " Where " + CONDI$
                    SQLX$ = SQLX$ + " ORDER BY Fentre_Sol"
                    Set RS = READSET(SQLX$)
                    With RS
                      Do While Not .EOF
                         i& = i& + 1
                         VENTABNEW.MSF.Rows = i& + 1
                         CAPEN = XVALO(!cant_ocom) - XVALO(!Cant_Rec)
                         ACUMULACAPEN = ACUMULACAPEN + CAPEN
                         NROCOM = XVALO(!nume_ocom)
                         LOTECO$ = SAFETEXT$(!Doc_Orig)
                         VENTABNEW.MSF.TextMatrix(i&, 1) = CI1%
                         VENTABNEW.MSF.TextMatrix(i&, 2) = SAFETEXT$(!Cod_Exte)
                         VENTABNEW.MSF.TextMatrix(i&, 3) = SAFETEXT$(!Descritem)
                         VENTABNEW.MSF.TextMatrix(i&, 4) = SAFETEXT$(!Uni_Mas)
                         VENTABNEW.MSF.TextMatrix(i&, 12) = XVALO(!NuOrd_Item)
                         
                         If NROCOM = NUMEOCOMPRA& Then
                           VENTABNEW.MSF.TextMatrix(i&, 5) = FORMATNUM$(CANTIXOCOM(NROCOM), "F12.1") 'CANTIDAD INGRESADA POR EL USUARIO
                           CANTIXOCOM(NROCOM) = 0
                         Else
                           VENTABNEW.MSF.TextMatrix(i&, 5) = 0
                         End If
                         
                         PROPONGO = CAPEN
                         If CAPEN > CANTIXCODIGO(CI1%) Then
                           PROPONGO = CANTIXCODIGO(CI1%)
                         End If
                         VENTABNEW.MSF.TextMatrix(i&, 6) = FORMATNUM$(PROPONGO, "F10.1")
                         CANTIXCODIGO(CI1%) = CANTIXCODIGO(CI1%) - PROPONGO
                         If CANTIXCODIGO(CI1%) < 0.005 Then CANTIXCODIGO(CI1%) = 0
                         
                         VENTABNEW.MSF.TextMatrix(i&, 7) = NROCOM
                         VENTABNEW.MSF.TextMatrix(i&, 8) = FECHATEX$(CVINT(!Fentre_Sol))
                         VENTABNEW.MSF.TextMatrix(i&, 9) = FORMATNUM$(CAPEN, "F10.1")
                         VENTABNEW.MSF.TextMatrix(i&, 10) = FORMATNUM$(PROPONGO, "F10.1")
                         VENTABNEW.MSF.TextMatrix(i&, 11) = LOTECO$
                         .MoveNext
                      Loop
                      CODILEIDO(CI1%) = 1
                      EXCEDENTE = CANTIXCODIGO(CI1%)
                      ACUMULO_PROPUESTO = 0
                      If EXCEDENTE > 0.005 Then
                         If CAN_ACUMULA_X_CODIGO <= (ACUMULACAPEN * (1 + PORCENTAJE_SOBRE / 100)) Then
                           'ACA AGREGO EN LA GRILLA DE ABAJO HACIA ARRIBA VALIDANDO QUE COINCIDA EL CODIGO HASTA LLEGAR A 0
                            ACUMULO_SOBRE_EMBARQUE_APLICADO = 0
                            VENTABNEW.Caption = VENTANA.TITULO + " (Considerando Sobre-Emarque " & PORCENTAJE_SOBRE & "%)"
                            For i& = VENTABNEW.MSF.Rows - 1 To 1 Step -1
                              If VENTABNEW.MSF.TextMatrix(i&, 1) = CI1% Then
                                CANTIPROPUESTA = XVALO(VENTABNEW.MSF.TextMatrix(i&, 6))
                                CANTIPROPUESTA_MAS_SOBREMBARQUE = Int(CANTIPROPUESTA * (1 + PORCENTAJE_SOBRE / 100))
                                SOBRE_EMBARQUE_APLICADO = CANTIPROPUESTA_MAS_SOBREMBARQUE - CANTIPROPUESTA
                                If SOBRE_EMBARQUE_APLICADO > EXCEDENTE Then SOBRE_EMBARQUE_APLICADO = EXCEDENTE
                                EXCEDENTE = EXCEDENTE - SOBRE_EMBARQUE_APLICADO
                                ACUMULO_SOBRE_EMBARQUE_APLICADO = ACUMULO_SOBRE_EMBARQUE_APLICADO + SOBRE_EMBARQUE_APLICADO
                                ACUMULO_PROPUESTO = ACUMULO_PROPUESTO + CANTIPROPUESTA_MAS_SOBREMBARQUE
                                VENTABNEW.MSF.TextMatrix(i&, 6) = FORMATNUM$(XVALO(CANTIPROPUESTA + SOBRE_EMBARQUE_APLICADO), "F10.1")
                                JJJ& = JJJ& + 1
                                If EXCEDENTE <= 0 Then
                                   Exit For
                                End If
                                
                              End If
                           Next i&
                           'ESTO LO PONGO POR SI POR CUESTION DE LOS DECIMALES CUANDO CALCULA EL SOBRE EMBARQUE _
                           NO LLEGA A LA CANTIDAD INGRESADA (-POR EL TEMA DEL REDONDEO HACIA ABAJO-) AJUSTA EN LA ULTIMA FILA
                           If XVALO(ACUMULO_PROPUESTO) < (ACUMULO_PROPUESTO + EXCEDENTE) Then
                                  VENTABNEW.MSF.TextMatrix(JJJ&, 6) = FORMATNUM$(XVALO(ACUMULO_PROPUESTO - CANTIPROPUESTA_MAS_SOBREMBARQUE) + XVALO(EXCEDENTE), "F10.1")
                                  Call COLOREAR_TEXTO_SOLO_UNA_AFILA(VENTABNEW.MSF, vbRed, i&, 6)
                           End If

                         Else 'SI VIENE AQUI ES POR QUE EL EXCEDENTE SUPERA LAS COMPRAS PENDIENTES + EL SOBRE EMBARQUE
                           'APLICO EL EXCEDENTE EN LA ULTIMA FILA Y YA, EL RESTO LO DEJO COMO ESTA.
                           'DE TODOS MODOS VA A PROTESTAR POR EL SOBRE EMBARQUE, PASARA EL QUE POSEA DERECHOS.
                           For i& = VENTABNEW.MSF.Rows - 1 To 1 Step -1
                              If VENTABNEW.MSF.TextMatrix(i&, 1) = CI1% Then
                                CANTIPROPUESTA = XVALO(VENTABNEW.MSF.TextMatrix(i&, 6))
                                VENTABNEW.MSF.TextMatrix(i&, 6) = FORMATNUM$(XVALO(CANTIPROPUESTA + EXCEDENTE), "F10.1")
                                Call COLOREAR_TEXTO_SOLO_UNA_AFILA(VENTABNEW.MSF, vbRed, i&, 6)
                                Exit For
                              End If
                           Next i&
                         End If
                      End If
                    End With
                    RS.Close
                    Set RS = Nothing
                End If
             End If
            'ESTE IF ES PARA AGREGAR EN LA FILA DE LA ORDEN DE COMPRA LA CANTIDAD QUE INFORMO EL USUARIO _
            'EJEMPLO SI SE INGRESO PARA UN MISMO ITEM CANTIDADES Y LA PROPUESTA PUEDE ACOMODAR TODO EN EN LA PRIMERA_
            'SOLO ESTABA MOSTRANDO COMO CANTIDAD INGRESADA DEL USUARIO LA PRIMERA P Q CUANDO PEGABA LA VUELTA
            'YA ESTABA CUMPLIDA Y NO LO PONIA
            If CANTIXOCOM(NUMEOCOMPRA&) > 0 Then
              For i& = 1 To VENTABNEW.MSF.Rows - 1
                 If VENTABNEW.MSF.TextMatrix(i&, 7) = NUMEOCOMPRA& And XVALO(VENTABNEW.MSF.TextMatrix(i&, 5)) <= 0 Then
                   VENTABNEW.MSF.TextMatrix(i&, 5) = FORMATNUM$(CANTIXOCOM(NUMEOCOMPRA&), "F12.1")
                   CANTIXOCOM(NUMEOCOMPRA&) = 0
                 End If
              Next i&
            End If
           End If
         End If
       Next YX&
       VENTABNEW.Width = 14000
       VENTABNEW.Show 1
       Set VENTANA = Nothing
       '
       '
       'SI SE APROBO LA VENTANA VUELVE A CARGAR EL BOLREDET PARA QUE CONTINUE SU CAMINO.
       If VENTABNEW.APROBADO% > 0 Then
           APROBO% = 1
           For i& = 1 To VENTABNEW.MSF.Rows - 1
             CI1% = XVALO(VENTABNEW.MSF.TextMatrix(i&, 1))
             Y$ = REGBLAN$("!" + ARCHICA$)
             Call REPLA(Y$, MKI$(CI1%), 83, 2)
             Call REPLA(Y$, Unimed$(CI1%), 85, 4)
              CAPEN = XVALO(VENTABNEW.MSF.TextMatrix(i&, 6))
             Call REPLA(Y$, MKD$(CDbl(CAPEN)), 89, 8)
              CANT_ORIG_PEND = XVALO(VENTABNEW.MSF.TextMatrix(i&, 9))
             Call REPLA(Y$, MKD$(CDbl(CANT_ORIG_PEND)), 163, 8)
             '
'             Call REPLA(y$, Chr$(XVALO(!NuOrd_Item)), 98, 1)
             Call REPLA(Y$, MKS$(0), 100, 4) 'no tiene AE
             '
             LOTECO$ = VENTABNEW.MSF.TextMatrix(i&, 11)
             Call REPLA(Y$, LOTECO$, 107, 12)
              NUOR% = XVALO(VENTABNEW.MSF.TextMatrix(i&, 12))
             Call REPLA(Y$, Chr$(NUOR%), 98, 1)
              NUOC& = XVALO(VENTABNEW.MSF.TextMatrix(i&, 7))
             Call REPLA(Y$, MKS$(NUOC&), 151, 4)
             FENTRESO% = FECHANUM(VENTABNEW.MSF.TextMatrix(i&, 8))
             Call REPLA(Y$, MKI$(FENTRESO%), 155, 2)
             Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
             Call GRAREG("!" + ARCHICA$, Y$, i&)
           Next i&
           Call SETULTREG&("!" + ARCHICA$, (i& - 1))
       Else
            APROBO% = 0
       End If
        
End Sub

Sub TRANSFIERE_COMPONENTES(CI%, NUMEORF$, NUMEBORE$, CANTIEQUI%)
        '
        'SINTESIS
        'RECORRER EL RESERVA SEGUN ORDEN DE FABRICACION
        'OBTENER EL CODIGO , CANTIDAD UNITARIA, Y EL LOTE
        'OBTENER EL STOCK SEGUN EL DEPOSITO DEL ARMADOR, TR, CI%,O.F
        'SI LA CANTIDAD RESERVADA ES <= AL STOCK EN EL DEPOSITO DEL ARMADOR
        'TRANSFERIR DEL ARMADOR AL DEPOSITO DESTINO
        '******************************************
        '
        'LO VA A REALIZAR EN RECEPCION
        
        'PARAMETRO PARA PASAR A RUTINA MOVISTOK
12      FE% = HOY(HOS$)
        CMOV$ = "TR"
        NVM$ = TRIM$(ProNroComprobante("TR"))
        NRO& = XVALO(NVM$)
        DOPRI$ = "TR-" + NVM$
        DEPARMAD% = XVALO(Control("", "DEPOARMA")) ' DEPOSITO DE ARMADOR
        DEPFABRI% = Control("", "DEPODEST")  ' DEPOSITO DE DESTINO
'        CONDI$ = "Cod_Inte = " & CI% & " and NUMEROSERIE = '" & NUMSERIE$ & "'"
'        NUMEORF$ = VALOBADA("TANUMSE", "IdSubOr", CONDI$, "NOMESS")
        DOSE$ = TRIM$(NUMEORF$) ' NUMERO DE ORDEN DE FABRICACION
        NUARMA = XVALO(VALOBADA("ORDEFABR", "ARMADOR", "IdSubOr = '" & NUMEORF$ & "'", "NOMESS"))
        DARMA$ = ECOARCH$("ARMADORS", MKI$(Int(NUARMA)))
        REFE32$ = "Recep.Equipo Arm." + DARMA$
        '
        SQLX$ = "SELECT UsoUnit, COD_INTE, IDENLOTE  FROM RESERVA "
        SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NUMEORF$ & "'"
        Dim RESTMP As New ADODB.Recordset
        Set RESTMP = FLEXCONN.Execute(SQLX$)
        FIN& = CantRegistros("RESERVA", "IdSubOr = '" & NUMEORF$ & "'", "NOMESS")
        i& = 0
        With RESTMP
          Do While Not .EOF
             i& = i& + 1
             Call TRACE(24, i&, FIN&)
             'RECORRO EL RESERVA
             CANTIUNIT = XVALO(!UsoUnit)
             CANTIUNIT = CANTIUNIT * CANTIEQUI%
             LOTEASIG$ = SAFETEXT$(!idenlote)
             CiComp% = XVALO(!cod_inte)
             '
             'OBTENGO EL STOCKS SEGUN O.F. , CODIGO, DEPOSITO Y CODIMOVI = TR
             CONDI1$ = "Cod_Inte = " & CiComp% & " AND DepoMovi = " & DEPARMAD% & " And DoSecCor = '" & DOSE$ & "' AND CODIMOVI = 'TR'"
             STOCK32 = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE - CANTSALID)", CONDI1$, "NOMESS"))
             '
             'TRANSFIERE DEL 32 AL 54
             If CANTIUNIT <= STOCK32 Then
                ' PRIMERO INTENTA AGREGAR A LOS MOVIMIENTOS EXISTENTES
                CONDI5$ = "CODIMOVI = '" + CMOV$ + "' "
                CONDI5$ = CONDI5$ & " AND COD_INTE = " & CiComp%
                If TRIM$(LOTEASIG$) <> "" Then CONDI5$ = CONDI5$ & " AND IDENLOTE = '" & LOTEASIG$ & "'"
                CONDI5$ = CONDI5$ + " AND DepoMovi = " & DEPARMAD%
                CONDI5$ = CONDI5$ & " AND DOSECCOR = '" & DOSE$ & "'"
                CONDI5$ = CONDI5$ & " AND CANTSALID > 0"
                QTRX& = CantRegistros("MOVISTO", CONDI5$, "NOMESS")
                If QTRX& > 0 Then
                   CANYACO = CANTIUNIT + XVALO(VALOBADA("MOVISTO", "Sum(CANTSALID)", CONDI5$, "NOMESS"))
                   Call ACTUREGISTRO("MOVISTO", "CANTSALID", CONDI5$, CANYACO, RAFE&, "NOMESS")
                   GoTo 15      ' LO AGREGO AL MOVIMIENTO PREVIOS
                End If
                '
                NUORIT% = NUORIT% + 1
                Call MOVISTOK(FE%, CiComp%, LOTEASIG$, CMOV$, DOPRI$, DOPRI$, DOSE$, DOSE$, REFE32$, NUORIT%, 0, 1, 0, DEPARMAD%, CANTIUNIT * (-1))
                '
15              ' PRIMERO INTENTA AGREGAR A LOS MOVIMIENTOS EXISTENTES
                CONDI5$ = "CODIMOVI = '" + CMOV$ + "' "
                CONDI5$ = CONDI5$ & " AND COD_INTE = " & CiComp%
                If TRIM$(LOTEASIG$) <> "" Then CONDI5$ = CONDI5$ & " AND IDENLOTE = '" & LOTEASIG$ & "'"
                CONDI5$ = CONDI5$ + " AND DepoMovi = " & DEPFABRI%
                CONDI5$ = CONDI5$ & " AND DOSECCOR = '" & DOSE$ & "'"
                CONDI5$ = CONDI5$ & " AND CANTINGRE > 0"
                QTRX& = CantRegistros("MOVISTO", CONDI5$, "NOMESS")
                If QTRX& > 0 Then
                   CANYACO = CANTIUNIT + XVALO(VALOBADA("MOVISTO", "Sum(CANTINGRE)", CONDI5$, "NOMESS"))
                   REFEXX$ = SAFETEXT$(VALOBADA("MOVISTO", "ReferCor", CONDI5$, "NOMESS")) + "," + NUMEBORE$
                   Call ACTUREGISTRO("MOVISTO", "CANTINGRE", CONDI5$, CANYACO, RAFE&, "NOMESS")
                   Call ACTUREGISTRO("MOVISTO", "ReferCor", CONDI5$, REFEXX$, RAFE&, "NOMESS")
                   GoTo 25      ' LO AGREGO AL MOVIMIENTO PREVIOS
                End If
                '
                Call MOVISTOK(FE%, CiComp%, LOTEASIG$, CMOV$, DOPRI$, DOPRI$, DOSE$, DOSE$, NUMEBORE$, NUORIT%, 0, 1, 0, DEPFABRI%, CANTIUNIT)
             End If
             '
25           .MoveNext
          Loop
        End With
        RESTMP.Close
        Set RESTMP = Nothing
        
        
        
End Sub
Sub TRANSF_MASIVA_COMPONENTES(NUMEORF$)
     'PARAMETRO PARA PASAR A RUTINA MOVISTOK
12      FE% = HOY(HOS$)
        CMOV$ = "TR"
        NVM$ = TRIM$(ProNroComprobante("TR"))
        NRO& = XVALO(NVM$)
        DOPRI$ = "TR-" + NVM$
        DEPARMAD% = XVALO(Control("", "DEPOARMA")) ' DEPOSITO DE ARMADOR
        DEPFABRI% = Control("", "DEPODEST")  ' DEPOSITO DE DESTINO
        CONDI$ = "Cod_Inte = " & CI% & " and NUMEROSERIE = '" & NUMSERIE$ & "'"
        NUMEORF$ = VALOBADA("TANUMSE", "IdSubOr", CONDI$, "NOMESS")
        DOSE$ = TRIM$(NUMEORF$) ' NUMERO DE ORDEN DE FABRICACION
        NUARMA = XVALO(VALOBADA("ORDEFABR", "ARMADOR", "IdSubOr = '" & NUMEORF$ & "'", "NOMESS"))
        DARMA$ = ECOARCH$("ARMADORS", MKI$(Int(NUARMA)))
        REFE32$ = "Recep.Equipo Arm." + DARMA$
        '
        SQLX$ = "SELECT * FROM RESERVA "
        SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NUMEORF$ & "'"
        Dim RESTMP As New ADODB.Recordset
        Set RESTMP = FLEXCONN.Execute(SQLX$)
        FIN& = CantRegistros("RESERVA", "IdSubOr = '" & NUMEORF$ & "'", "NOMESS")
        i& = 0
        With RESTMP
          Do While Not .EOF
             i& = i& + 1
             Call TRACE(24, i&, FIN&)
             'RECORRO EL RESERVA
             CANTIUNIT = XVALO(!UsoUnit)
             LOTEASIG$ = SAFETEXT$(!idenlote)
             CiComp% = XVALO(!cod_inte)
             '
             'OBTENGO EL STOCKS SEGUN O.F. , CODIGO, DEPOSITO Y CODIMOVI = TR
             CONDI1$ = "Cod_Inte = " & CiComp% & " AND DepoMovi = " & DEPARMAD% & " And DoSecCor = '" & DOSE$ & "' AND CODIMOVI = 'TR'"
             STOCK32 = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE - CANTSALID)", CONDI1$, "NOMESS"))
             '
             'TRANSFIERE DEL 32 AL 54
             If CANTIUNIT <= STOCK32 Then
                ' PRIMERO INTENTA AGREGAR A LOS MOVIMIENTOS EXISTENTES
                CONDI5$ = "CODIMOVI = '" + CMOV$ + "' "
                CONDI5$ = CONDI5$ & " AND COD_INTE = " & CiComp%
                If TRIM$(LOTEASIG$) <> "" Then CONDI5$ = CONDI5$ & " AND IDENLOTE = '" & LOTEASIG$ & "'"
                CONDI5$ = CONDI5$ + " AND DepoMovi = " & DEPARMAD%
                CONDI5$ = CONDI5$ & " AND DOSECCOR = '" & DOSE$ & "'"
                CONDI5$ = CONDI5$ & " AND CANTSALID > 0"
                QTRX& = CantRegistros("MOVISTO", CONDI5$, "NOMESS")
                If QTRX& > 0 Then
                   CANYACO = CANTIUNIT + XVALO(VALOBADA("MOVISTO", "Sum(CANTSALID)", CONDI5$, "NOMESS"))
                   Call ACTUREGISTRO("MOVISTO", "CANTSALID", CONDI5$, CANYACO, RAFE&, "NOMESS")
                   GoTo 15      ' LO AGREGO AL MOVIMIENTO PREVIOS
                End If
                '
                NUORIT% = NUORIT% + 1
                Call MOVISTOK(FE%, CiComp%, LOTEASIG$, CMOV$, DOPRI$, DOPRI$, DOSE$, DOSE$, REFE32$, NUORIT%, 0, 1, 0, DEPARMAD%, CANTIUNIT * (-1))
                '
15              ' PRIMERO INTENTA AGREGAR A LOS MOVIMIENTOS EXISTENTES
                CONDI5$ = "CODIMOVI = '" + CMOV$ + "' "
                CONDI5$ = CONDI5$ & " AND COD_INTE = " & CiComp%
                If TRIM$(LOTEASIG$) <> "" Then CONDI5$ = CONDI5$ & " AND IDENLOTE = '" & LOTEASIG$ & "'"
                CONDI5$ = CONDI5$ + " AND DepoMovi = " & DEPFABRI%
                CONDI5$ = CONDI5$ & " AND DOSECCOR = '" & DOSE$ & "'"
                CONDI5$ = CONDI5$ & " AND CANTINGRE > 0"
                QTRX& = CantRegistros("MOVISTO", CONDI5$, "NOMESS")
                If QTRX& > 0 Then
                   CANYACO = CANTIUNIT + XVALO(VALOBADA("MOVISTO", "Sum(CANTINGRE)", CONDI5$, "NOMESS"))
                   REFEXX$ = SAFETEXT$(VALOBADA("MOVISTO", "ReferCor", CONDI5$, "NOMESS")) + "," + NUMEBORE$
                   Call ACTUREGISTRO("MOVISTO", "CANTINGRE", CONDI5$, CANYACO, RAFE&, "NOMESS")
                   Call ACTUREGISTRO("MOVISTO", "ReferCor", CONDI5$, REFEXX$, RAFE&, "NOMESS")
                   GoTo 25      ' LO AGREGO AL MOVIMIENTO PREVIOS
                End If
                '
                Call MOVISTOK(FE%, CiComp%, LOTEASIG$, CMOV$, DOPRI$, DOPRI$, DOSE$, DOSE$, NUMEBORE$, NUORIT%, 0, 1, 0, DEPFABRI%, CANTIUNIT)
             End If
             '
25           .MoveNext
          Loop
        End With
        RESTMP.Close
        Set RESTMP = Nothing
        
End Sub
Sub ACTUCOMPEN(NUOC&, NUMOIT%, CARECI, NANTEN&)
    '
    TXXTRAZA$ = ""
    
    If NATEN& = 0 Then 'O/C CERRADA'\\\OT-06-0149-WAR-26/04/06///
          CONDI$ = "Nume_Ocom = " & NUOC&
          CONDI$ = CONDI$ + " AND NuOrd_Item = " & NUMOIT%
          CONDI$ = CONDI$ + " AND Stat_Ocom < 8 "
          
TXXTRAZA$ = TXXTRAZA$ + "O\Compra= " & TRIM$(Str$(NUOC&)) & " - N\Orden= " & TRIM$(Str$(NUMOIT%)) & " - Cantidad= " & TRIM$(Str$(CARECI)) & vbCrLf

          SQLX$ = "SELECT * FROM OCOMDETA WHERE " & CONDI$
          Dim RSTBRECEP As ADODB.Recordset
          Set RSTBRECEP = New ADODB.Recordset
          RSTBRECEP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          
TXXTRAZA$ = TXXTRAZA$ + "Generó consulta SQL" & vbCrLf

          With RSTBRECEP
              If (.BOF And .EOF) Then
                Call MENSERR(24, "Imposible Netear OC='" & TRIM$(Str$(NUOC&)) & " e Item='" & TRIM$(Str$(NUMOIT%)))
              End If
              On Error Resume Next
              .MoveFirst
              On Error GoTo 0
              CI1% = XVALO(!cod_inte)
              IMPCONOC% = XVALO(!IContab)
              COEFUNI = XVALO(!Coef_Unids)
                If COEFUNI < 0.0001 Then COEFUNI = 1
              CANTIORIG = XVALO(!Cant_Rec)
              CANTORECI = XVALO(!Cant_Rec) + XVALO(CARECI / COEFUNI)
              !Cant_Rec = XVALO(CANTORECI)
              CAPENDIEN = XVALO(!cant_ocom) - CANTORECI
              NUORSE& = XVALO(!N_ORSERV)
              DOCUORIG$ = SAFETEXT$(!Doc_Orig)
              
TXXTRAZA$ = TXXTRAZA$ + "Cantidad Original " & FORMATNUM$(CANTIORIG, "F12.2") & vbCrLf
TXXTRAZA$ = TXXTRAZA$ + "Cantidad Recibida " & FORMATNUM$(CARECI / COEFUNI, "F12.2") & vbCrLf
TXXTRAZA$ = TXXTRAZA$ + "Actualizó Codigo '" & CODEXT$(CI1%) & "' - Cantidad: " & FORMATNUM$(CANTORECI, "F12.2") & vbCrLf

    '          CI1% = XVALO(VALOBADA("OCOMDETA", "Cod_Inte", CONDI$))
    '          IMPCONOC% = XVALO(VALOBADA("OCOMDETA", "IContab", CONDI$)) 'ASIGNO LA IMPUTACION CONTABLE
    '          COEFUNI = XVALO(VALOBADA("OCOMDETA", "Coef_Unids", CONDI$))
    '             If COEFUNI < 0.0001 Then COEFUNI = 1
    '          CANTORECI = XVALO(VALOBADA("OCOMDETA", "Cant_Rec", CONDI$)) + CARECI / COEFUNI
    '             Call ACTUREGISTRO("OCOMDETA", "Cant_Rec", CONDI$, CANTORECI, RAFE&)
    '          CAPENDIEN = XVALO(VALOBADA("OCOMDETA", "Cant_Ocom", CONDI$)) - CANTORECI
    '             NUORSE& = XVALO(VALOBADA("OCOMDETA", "N_OrServ", CONDI$))
    '          DOCUORIG$ = VALOBADA("OCOMDETA", "Doc_Orig", CONDI$)
              '
              If CAPENDIEN >= 0.05 Then
                    If InStr(1, EMPREAC$, "GABAL") > 0 Then
                        STATOF% = XVALO(VALOBADA("ORDEFABR", "STATUORF", "IDSUBOR='" & DOCUORIG$ & "'"))
                        If STATOF% = 3 Then
                            QOFAPRO# = XVALO(VALOBADA("ORDEFABR", "CANAPRO", "IDSUBOR='" & DOCUORIG$ & "'"))
                            'QOCRECI# = XVALO(VALOBADA("OCOMDETA", "Cant_Rec", CONDI$))
                            QOCRECI# = CANTORECI
                            If Abs(QOCRECI# - QOFAPRO#) < 0.05 Then
                                STATOC% = 3
                                If (QOFORIG# - QOFAPRO#) > 0.05 Then STATOC% = 8
                                'Call ACTUREGISTRO("OCOMDETA", "Stat_Ocom", CONDI$, STATOC%, RAFE&)
                                !Stat_Ocom = STATOC%
                            End If
                        End If
                    ElseIf Control("BOLRECEP", "CIERREOC") = "SI" Then
                        If COMALTER%("Esta O/C por Codigo " + TRIM$(CODEXT$(CI1%)) + "\Queda con un Saldo Pendiente de " + TRIM$(FORMATNUM$(CAPENDIEN, "F10.1")) + " " + TRIM$(Unimed$(CI1%)) + ".\\Mantener Abierta\Cancelar Pendiente") = 2 Then
                            'Call ACTUREGISTRO("OCOMDETA", "Stat_Ocom", CONDI$, 8, RAFE&)
                            STATOC% = 8
                            !Stat_Ocom = STATOC%
                        End If
                    End If
              Else
                 'Call ACTUREGISTRO("OCOMDETA", "Stat_Ocom", CONDI$, 3, RAFE&)
                 STATOC% = 3
                 !Stat_Ocom = STATOC%
              End If
              .Update
              
TXXTRAZA$ = TXXTRAZA$ + "Grabó registro" & vbCrLf & vbCrLf
              '
          End With
          
       ElseIf NATEN& > 0 Then '0/C ABIERTA (AE)
          CONDI$ = "Nume_Ocom = " & NUOC&
          CONDI$ = CONDI$ + " + " And cod_inte = " & CI1%"
          CONDI$ = CONDI$ + " AND Stat_Ocom < 8 "
          '
          IMPCONOC% = XVALO(VALOBADA("DETOCABI", "IContab", CONDI$)) 'ASIGNO LA IMPUTACION CONTABLE
          COEFUNI = XVALO(VALOBADA("DETOCABI", "Coef_Unids", CONDI$))
             If COEFUNI < 0.0001 Then COEFUNI = 1
          CANTORECI = XVALO(VALOBADA("DETOCABI", "Cant_Rec", CONDI$)) + CARECI / COEFUNI
             Call ACTUREGISTRO("DETOCABI", "Cant_Rec", CONDI$, CANTORECI, RAFE&)
          CAPENDIEN = XVALO(VALOBADA("DETOCABI", "Cant_Ocom", CONDI$)) - CANTORECI
             NUORSE& = XVALO(VALOBADA("DETOCABI", "N_OrServ", CONDI$))
          If CAPENDIEN < 0.05 Then
             Call ACTUREGISTRO("DETOCABI", "Stat_Ocom", CONDI$, 3, RAFE&)
          End If
                 '
                 ' ESTA PARTE FALTA ARREGLARLA - EPB 28/FEB/2012
                 strSQL = "Select * from AutoEnt where Nume_Ocom = " & NUOC&
                 strSQL = strSQL + " AND NuOrd_Item = " & NUMOIT%
                 strSQL = strSQL + " AND Stat_Ocom < 8 "
                 strSQL = strSQL + " AND Nume_Aten= " & NATEN&
                
                 Dim AutoEnt As ADODB.Recordset
                 Set AutoEnt = New ADODB.Recordset
                 AutoEnt.CursorType = adOpenDynamic
                 AutoEnt.LockType = adLockPessimistic
                 AutoEnt.Open FILTSQL$(strSQL), FLEXCONN, , , adCmdText
                 If Not (AutoEnt.EOF And AutoEnt.BOF) Then 'si no esta vacio actualiza
                   With AutoEnt
                       COEFUNI = 1 ' !Coef_Unids
                     If COEFUNI < 0.0001 Then COEFUNI = 1
                       !Acuen_Aten = XVALO(!Acuen_Aten) + CARECI / COEFUNI
                       CAPENDIEN = XVALO(!Cant_Aten) - XVALO(!Acuen_Aten)
                       !Sald_Aten = CAPENDIEN
                     If CAPENDIEN < 0.05 Then
                       !Stat_Aten = 3
                     End If
                     .Update
                   End With
                 End If
                 AutoEnt.Close
                 Set AutoEnt = Nothing
    End If
    '
''GRABO EL ARCHIVO DE TRAZABILIDAD
'For I1% = 1 To 32767
'     NOMVARIA$ = REPLACAR(Mid$(TRIM$(Time), 1, 8), ":", "")
'     ARCH$ = LUDAT$ & "TRAZABRR_" & NOMVARIA$ & ".TXT"
'     If EXISTE%(ARCH$) < 1 Then
'       NARCHI2% = FILEOPEN%(ARCH$, 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
'       Print #NARCHI2%, TXXTRAZA$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
'       Close #NARCHI2%
'       Exit For
'     End If
'Next I1%
    '
End Sub

Private Sub RECEOCLI(NC2%, FR%, NBR&, NREMI$, DP2%, NUCONVI$)
   'MATERIAL TEXTO LIBRE
   DARX$ = TRIM$(Control$("BOLRECEP", "DIANREC"))
     If DARX$ = "" Then
          DIANRE& = 10000
        Else
          DIANRE& = Val(DARX$)
     End If
     HOIX1 = HOY(HOS$)
     '
     ANUCA% = 0
     If Control$("BOLRECEP", "ANUCANTI") = "SI" Then
       ANUCA% = 1
     End If
     '
     Call COPYSTRU("INDIOCOM", "INDIRETE")
     '
10   Screen.MousePointer = 11
     Call ABRECOMPRAS
     'jandy_1
     SQLX$ = "Select * from OCOMDETA"
     SQLX$ = SQLX$ + " Where Npro_Ocom = " & CStr(NC2%) & " AND Cod_Inte = 0 AND Stat_Ocom < 8 "
     SQLX$ = SQLX$ + " AND Stat_Ocom <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
     SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
     'Set COMPTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     
     Dim COMPTEMP_3 As ADODB.Recordset
     Set COMPTEMP_3 = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With COMPTEMP_3
       '
       JJ& = 0
       On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Screen.MousePointer = 1
         GoTo 1117
       End If
       On Error GoTo 0
       '
       Y$ = REGBLAN("INDIRETE")
       Do While (.EOF = False)
         DESIT$ = TRIM$(!Descritem)
         If DESIT$ = "" Or UCase$(DESIT$) = "MATERIAL NO CODIFICADO" Then
           DESIT$ = TRIM$(!Delibre)
         End If
         Call REPLA(Y$, MKS(!nume_ocom), 1, 4)
         Call REPLA(Y$, TRIM$(Str$(!NuOrd_Item)) + ") " + DESIT$ + " (S=" + TRIM$(FORMATNUM$(!cant_ocom - !Cant_Rec, "F8.1")) + ")", 5, 44)
         JJ& = JJ& + 1
         Call GRAREG("!INDIRETE", Y$, JJ&)
        .MoveNext
       Loop
     End With
     '
     Call SETULTREG("!INDIRETE", JJ&)
     Call FRESHECHO("!INDIRETE")
1117 If JJ& < 1 Then
       Call COMUNI("No hay Compras Libres Pendientes\para este Proveedor.")
       Screen.MousePointer = 1
       Exit Sub
     End If
     '
     CANCELX1% = 0 '\\\OT-06-0178-WAR-12/05/06///
     Call SELECHO("!INDIRETE")
     VDEVU$ = VALACT1$("!INDIRETE")
     If VDEVU$ = "" Then
       CANCELX1% = 1 '\\\OT-06-0178-WAR-12/05/06///
       Screen.MousePointer = 1
       Exit Sub
     End If
     '
     'ver Ernesto
     'With OcomDeta
       NOC3& = Val(VDEVU$)
       NORIT& = Val(VALACT2$("!INDIRETE"))
       SQLX$ = "SELECT * FROM OCOMDETA"
       SQLX$ = SQLX$ + " WHERE NUME_OCOM = " & NOC3&
       SQLX$ = SQLX$ & " AND NuOrd_Item = " & NORIT&
       SQLX$ = SQLX$ + " AND Stat_Ocom < 8 "         ' NO ANULADAS
       SQLX$ = SQLX$ + " AND  Stat_Ocom <> - 1 "     ' EXCLUYE NO APROB.
       '
       '.FindFirst SQLX$
       Dim OCOMMMDETA As ADODB.Recordset
       Set OCOMMMDETA = New ADODB.Recordset
       OCOMMMDETA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
'
'       If .NoMatch = False Then
     With OCOMMMDETA
       If Not (.EOF And .BOF) Then
           On Error Resume Next
           NROC& = !nume_ocom
           DECRICOR$ = TRIM$(!Descritem)
           DECRILAR$ = !Delibre
           NC% = !Npro_Ocom
           CANREN# = !cant_ocom - !Cant_Rec
           IMPCONCO% = !IContab
           'ACA TOMO 'EV' y 'Grupo'
           NEV$ = !Nro_Otrab
           NGR$ = !Nro_Grup
           NORIT& = !NuOrd_Item
           On Error GoTo 0
           '
         Else
           Call COMUNI("Imposible Continuar. Número de O/C no Válido")
           Screen.MousePointer = 1
           Exit Sub
       End If
     End With
     '
1129 'LEVANTA LA VENTANA CON O/C
     ANUCA% = 0
     If Control$("BOLRECEP", "ANUCANTI") = "SI" Then
       ANUCA% = 1
     End If
     '
     With PRESOCLIB07
       .Label4 = NROC&
       .Label12 = NORIT&
       .Text1 = DECRICOR$
       .Text2 = DECRILAR$
       .Label5 = NC%
       .Label11 = rasocli(NC% * (-1))
       .Label7 = FORMATNUM$(CANREN#, "F8.1")
       .Text3 = .Label7: If ANUCA% > 0 Then Text3 = ""
       .Label2 = FECHATEX(FR%)
       .Label9 = NBR&
       .Label22 = NREMI$
       .Label17 = NEV$
       .Label19 = NGR$
       .Label13 = DP2%
       .Label14 = IMPCONCO%
       .Label16 = ECOARCH("TADEPOSI", Chr(DP2%))
       .Label24 = NUCONVI$
       .Show 1
     End With
     GoTo 10    ' PRESENTA NUEVAMENTE LISTA DE COMPRAS PENDIENTES
   '
   'HASTA ACA NUEVO MAT TEXTO LIBRE
   '
End Sub
'RUTINA ANTIGUA DE RECEPCION
'Sub BOLERECEP()
'Dim IMPCONOC%
''
''falta generar ficha de recepcion fervi-air
'     '
'     If ULTREG&("TATIREC") < 1 Then
'         X$ = REGBLAN$("TATIREC")
'         Call REPLA(X$, Chr$(1), 1, 1)
'         Call REPLA(X$, "Material General", 2, 31)
'         Call REPLA(X$, "BOLREGEN", 49, 16)
'         Call REPLA(X$, "BOLREGEN", 65, 16)
'         Call REGAPP("TATIREC", X$)
'         Call CIERRARCH("TATIREC")
'     End If
'     '
'1100 Call CIERRARCH("*.*")
'     '
'     NROBOL& = 1
'     Call ABRESTOCKS
'     Set BolreTemp = Stocks.OpenRecordset("SELECT DISTINCT NumeBoRe from BOLRECEP ORDER BY NumeBoRe DESC")
'     On Error Resume Next
'     BolreTemp.MoveFirst
'     If Err < 1 Then
'       NROBOL& = 1 + BolreTemp!NumeBore
'     End If
'     On Error GoTo 0
'     '
'     If TIMATEI% = 0 Then
'         X$ = REGLEIDO$("TATIREC", 1)
'         TIMATEI% = Asc(Left$(X$, 1))
'     End If
'     '
'     VDEF$ = MKI$(NC%) + Chr$(TIMATEI%) + CODICLI$((-1) * NC%)
'     EBOL$ = OBJPLA$("SELPROVRECACC", VDEF$)
'     If EBOL$ = "" Then
'        Exit Sub
'     End If
'     '
'     NC% = NUMEPRO%(TRIM$(Mid$(EBOL$, 4, 12)))
'     If NC% < 1 Or STATCLI%((-1) * NC%) < 1 Then
'       Call MENSERR(24, "Proveedor Inexistente")
'       GoTo 1100
'     End If
'     '
'     TIMATEI% = Asc(Mid$(EBOL$, 3, 1))
'     If TIMATEI% < 1 Or ECOARCH("TATIREC", Chr$(TIMATEI%)) = "" Then
'         Call MENSERR(24, "Falta Tipo de Material")
'         GoTo 1100
'     End If
'     '
'     REGT& = TIMATEI%
'     X$ = REGLEIDO$("TATIREC", REGT&)
'     FORMPLA$ = TRIM$(UCase$(Mid$(X$, 33, 16)))
'     FORMTAB$ = TRIM$(UCase$(Mid$(X$, 49, 16)))
'     FORIPRE$ = TRIM$(UCase$(Mid$(X$, 65, 16)))
'     '
'     NC1% = (-1) * NC%
'     '
'1105 Y$ = PEDY$
'     If Y$ = "" Then
'       Y$ = Space$(128)
'       Call REPLA(Y$, Chr$(0), 121, 1)
'     End If
'     '
'     Call REPLA(Y$, MKS$(NROBOL&), 1, 4)
'     Call REPLA(Y$, EBOL$, 7, 2)
'     Call REPLA(Y$, MKI$(HOY(HO$)), 5, 2)
'     Call REPLA(Y$, RASOCLI$(NC1%), 9, 32)
'     Call REPLA(Y$, DOMICLI$(NC1%), 41, 32)
'     CLOCA$ = TRIM$(CPOSCLI$(NC1%)): If CLOCA$ <> "" Then CLOCA$ = CLOCA$ + " "
'     Call REPLA(Y$, CLOCA$ + LOCACLI$(NC1%), 73, 32)
'     Call REPLA(Y$, Chr$(0), 121, 1)
'    '
'     OBJE$ = "BOLREC-ENCB"
'     If CONTROL$("", "NUMEBREC") = "MANUAL" Then
'         OBJE$ = "BOLREC-ENCB-M"
'     End If
'     '
'     TICOMPRO$ = TRIM$(CONTROL$("", "COMPBREC"))
'     '
'     PEDY$ = OBJPLA$(OBJE$, Y$)
'     If PEDY$ = "" Then Exit Sub
'     '
'     FF% = CVI(Mid$(PEDY$, 5, 2))
'     FEX = FF%: FECHAREC$ = FECHATEX$(FEX)
'     If TRIM$(FECHAREC$) = "" Then
'         Call MENSERR(24, "Fecha Incorrecta")
'         GoTo 1105
'     End If
'     '
'     REMI$ = TRIM$(Mid$(PEDY$, 105, 16))
'     If REMI$ = "" Then
'         Call MENSERR(24, "Falta Número de Remito del Proveedor")
'         GoTo 1105
'     End If
'     '
'     DP1% = Asc(Mid$(PEDY$, 121, 1))
'     DEPOSAL$ = TRIM$(ECOARCH$("ECODEP", Chr$(DP1%)))
'     If DP1% < 1 Or DEPOSAL$ = "" Then
'         Call MENSERR(24, "Falta Depósito de Entrada")
'         GoTo 1105
'     End If
'     '
'     NROPRO& = NROBOL&
'     NROBOL& = CVS(Left$(PEDY$, 4))
'     If NROBOL& < 1 Then
'         Call MENSERR(24, "Número Incorrecto de B.R.")
'         NROBOL& = NROPRO&
'         Call REPLA(PEDY$, MKS$(NROBOL&), 1, 4)
'         GoTo 1105
'     End If
'     '
'1112 Call CIERRARCH("*.*")
'     JJ& = 0
'     '
'     DARX$ = TRIM$(CONTROL$("BOLRECEP", "DIANREC"))
'     If DARX$ = "" Then
'          DIANRE& = 10000
'        Else
'          DIANRE& = Val(DARX$)
'     End If
'     HOIX1 = HOY(HOS$)
'     '
'     ANUCA% = 0
'     If CONTROL$("BOLRECEP", "ANUCANTI") = "SI" Then
'       ANUCA% = 1
'     End If
'     '
'     Screen.MousePointer = 11
'     Call ABRECOMPRAS
'     SQLX$ = "Select * from OCOMDETA"
'     SQLX$ = SQLX$ + " Where Stat_Ocom < 8 "
'     SQLX$ = SQLX$ + " AND  Stat_Ocom <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
'     SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
'     SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
'     Set COMPTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
'     '
'     With COMPTEMP
'       '
'       On Error Resume Next
'       .MoveFirst
'       If Err Then
'         On Error GoTo 0
'         Screen.MousePointer = 1
'         HAYOC% = 0
'         GoTo 11129
'       End If
'       On Error GoTo 0
'       '
'       JJ& = 0
'
'       Do While (.EOF = False)
'         CAPEN = !Cant_Ocom - !Cant_Rec
'         COEFUNI = !Coef_Unids
'         If COEFUNI < 0.0001 Then COEFUNI = 1
'         If CAPEN > 0.005 Then
'           NUOC& = !Nume_Ocom
'           FENTRESO% = CVINT(!Fentre_Sol)
'           FESO2 = FENTRESO%
'           If DIENTRE%(HOIX1, FESO2) <= DIANRE& Then
'             CI% = !Cod_Inte
'             LOTECO$ = !Doc_Orig
'             Y$ = REGBLAN$("!BOLREDET")
'             Call REPLA(Y$, MKI$(CI%), 83, 2)
'             Call REPLA(Y$, UNIMED$(CI%), 85, 4)
'             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 89, 8)
'             If ANUCA% = 1 Then
'               Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 163, 8)
'                Call REPLA(Y$, MKD$(0), 89, 8)
'             End If
'             '
'             Call REPLA(Y$, CHR$(!NuOrd_Item), 98, 1)
'             '
'             Call REPLA(Y$, LOTECO$, 107, 12)
'             Call REPLA(Y$, MKS$(NUOC&), 151, 4)
'             Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
'             Call REPLA(Y$, MKS$(REDET&), 189, 4)
'             HAYOC% = 1
'             JJ& = JJ& + 1
'             Call GRAREG("!BOLREDET", Y$, JJ&)
'           End If
'         End If
'        .MoveNext
'       Loop
'     End With
'     '/////****NUEVO****/////
'     'TENGO QUE FIJARME SI TIENE ALGUN MATERIAL EL PROV
'     'CONSIGNADO EN MACONSIG CON EL MISMO Nº DE PROV
'      Call ABRESTOCKS
'      SQLX$ = "Select * from MACONSIG"
'      SQLX$ = SQLX$ + " WHERE Cant_Rendid < Cant_Entreg - 0.001 "
'      SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
'      Set CONSIGTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
'      With CONSIGTMP
'        '
'        On Error Resume Next
'        .MoveFirst
'        If Err Then
'          On Error GoTo 0
'          Screen.MousePointer = 1
'          GoTo 11129
'        End If
'        On Error GoTo 0
'        '
'        Y$ = REGBLAN$("!BOLREDET")
'        Do While .EOF = False
'          CI% = !Cod_Inte
'          NUOC& = Val(Mid$(!Nro_Ocom, 4, 6))
'          CAPEN = !Cant_Entreg - !Cant_Rendid
'          Call REPLA(Y$, MKI$(CI%), 83, 2)
'          Call REPLA(Y$, UNIMED$(CI%), 85, 4)
'          Call REPLA(Y$, MKD$(0), 89, 8)
'          Call REPLA(Y$, "MP " + TRIM$(CSTRING$(MKS$(CAPEN), 3, 12, 1, 2)) + " " + UNIMED$(CI%), 107, 12)
'          Call REPLA(Y$, MKS$(NUOC&), 151, 4)
'          Call REPLA(Y$, MKS$((-1) * I&), 189, 4)
'          JJ& = JJ& + 1
'          Call GRAREG("!BOLREDET", Y$, JJ&)
'        .MoveNext
'        Loop
'      End With
'         'HASTA ACA LLEGA OK
'
''        UREOCO& = ULTREG&("OCOMDETA")
''        HAYOC% = 0
''        RFF$ = RECFILT$("OCOMDETA", 3, MKI$(NC%))
''        For UI& = 1 To Len(RFF$) Step 4
''          REDET& = CVS(Mid$(RFF$, UI&, 4))
''          If REDET& > 0 Then
''            If REDET& <= UREOCO& Then
''              X$ = REGLEIDO$("OCOMDETA", REDET&)
''              If CVI(Mid$(X$, 7, 2)) = NC% Then
''                If Asc(Mid$(X$, 124, 1)) < 8 Then ' abierta Y NO CUMPLIDA
''                  CAPEN = CVS(Mid$(X$, 55, 4)) - CVS(Mid$(X$, 99, 4))
''                  COEFUNI = CVS(Mid$(X$, 63, 4))
''                  If COEFUNI < 0.0001 Then COEFUNI = 1
''                  If CAPEN > 0 Then
''                    NUOC& = CVS(Left$(X$, 4))
''                    FENTRESO% = CVI(Mid$(X$, 95, 2))
''                    FESO2 = FENTRESO%
''                    If DIENTRE%(HOIX1, FESO2) <= DIANRE& Then
''                      CI% = CVI(Mid$(X$, 9, 2))
''                      LOTECO$ = Mid$(X$, 105, 12)
''                      Y$ = REGBLAN$("!BOLREDET")
''                      Call REPLA(Y$, MKI$(CI%), 83, 2)
''                      Call REPLA(Y$, UNIMED$(CI%), 85, 4)
''                      Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 89, 8)
''                      If ANUCA% = 1 Then
''                        Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 163, 8)
''                        Call REPLA(Y$, MKD$(0), 89, 8)
''                      End If
''                      Call REPLA(Y$, LOTECO$, 107, 12)
''                      Call REPLA(Y$, MKS$(NUOC&), 151, 4)
''                      Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
''                      Call REPLA(Y$, MKS$(REDET&), 189, 4)
''                      HAYOC% = 1
''                      JJ& = JJ& + 1
''                      Call GRAREG("!BOLREDET", Y$, JJ&)
''                    End If
''                  End If
''                End If
''              End If
''            End If
''          End If
''        Next UI&
'        '
'11129     If HAYOC% = 0 Then
'       Call COMUNI("No hay Ordenes de Compra Pendientes\para este Proveedor.")
'     End If
'     '
'' QUEDA PENDIENTE LA DEVOLUCION DE CONSIGNACIONES
''        For I& = 1 To ULTREG&("MACONSIG")
''          X$ = REGLEIDO$("MACONSIG", I&)
''          If CVI(Mid$(X$, 3, 2)) = NC% Then
''            CAPEN = CVD(Mid$(X$, 63, 8)) - CVD(Mid$(X$, 111, 8))
''            If CAPEN > 0 Then
''              NUOC& = Val(Mid$(X$, 24, 6))
''              CI% = CVI(Mid$(X$, 33, 2))
''              Y$ = REGBLAN$("!BOLREDET")
''              Call REPLA(Y$, MKI$(CI%), 83, 2)
''              Call REPLA(Y$, UNIMED$(CI%), 85, 4)
''              Call REPLA(Y$, MKD$(0), 89, 8)
''              Call REPLA(Y$, "MP " + TRIM$(CSTRING$(MKS$(CAPEN), 3, 12, 1, 2)) + " " + UNIMED$(CI%), 107, 12)
''              Call REPLA(Y$, MKS$(NUOC&), 151, 4)
''              Call REPLA(Y$, MKS$((-1) * I&), 189, 4)
''              For UI& = 1 To JJ&
''                If Left$(Y$, 154) = Left$(REGLEIDO$("!BOLREDET", UI&), 154) Then GoTo 11129 ' SUPRIME SI COINCIDE
''              Next UI&
''              JJ& = JJ& + 1
''              Call GRAREG("!BOLREDET", Y$, JJ&)
''            End If
''          End If
''11129   Next I&
'
'     Call SETULTREG("!BOLREDET", JJ&)
'     Call CIERRARCH("!BOLREDET")
'     '
'1113 AVENTA$ = "BOLREDET"
'     If ANUCA% = 1 Then AVENTA$ = "BOLRENUE"
'     '
'     VTB% = VENTABU%(AVENTA$ + "/NC/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(RASOCLI$(NC1%)) + " - " + FECHAREC$)
'     If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
'       X$ = "": Y$ = "": YY$ = ""
'       EBOL$ = "": PEDY$ = "": DESCRI$ = ""
'       GoTo 1100
'     End If
'     '
'     For YX& = 1 To ULTREG&("!BOLREDET")
'       YY$ = REGLEIDO$("!BOLREDET", YX&)
'       CI% = CVI(Mid$(YY$, 83, 2))
'       CAN# = CVD(Mid$(YY$, 89, 8))
'       If CI% > 0 Then
'         If CI% <= NUMAS% Then
'           If CAN# > 0.005 Then
'             GoTo 1114
'           End If
'         End If
'       End If
'     Next YX&
'     Call MENSERR(24, "Recepción Cancelada !!!\Faltan Cantidades o Códigos.")
'     GoTo 1113
'     '
'1114 Call SETULTREG("!OBSERVOF", 0)
'     Call CIERRARCH("!OBSERVOF")
'     '
'     VTB% = VENTABU%("OBSERVOF/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(RASOCLI$(NC1%)) + " - " + FECHAREC$)
'     If VTB% < 0 Then
'       If COMALTER%("Realmente Desea Cancelar ?\\Continuar\Cancelar") = 2 Then
'         X$ = "": Y$ = "": YY$ = ""
'         EBOL$ = "": PEDY$ = "": DESCRI$ = ""
'         GoTo 1100
'       End If
'     End If
'     '
'11141 CONTA% = ULTREG&("!BOLREDET")
'     '
'1115 If CONTA% > 0 Then
'       '
'       For YX& = 1 To ULTREG&("!BOLREDET")
'         YY$ = REGLEIDO$("!BOLREDET", YX&)
'         CI% = CVI(Mid$(YY$, 83, 2))
'         CAN# = CVD(Mid$(YY$, 89, 8))
'         If CI% < 1 Then
'           Call MENSERR(24, "Código no Válido.")
'           GoTo 1113
'         End If
'         If ANUCA% < 1 Then
'           If CAN# < 0.005 Then
'             Call MENSERR(24, "Falta Cantidad.")
'             GoTo 1113
'           End If
'         End If
'       Next YX&
'       '
'1125   Call ABRESTOCKS
'       Set BolreTemp = Stocks.OpenRecordset("SELECT DISTINCT NumeBoRe from BOLRECEP ORDER BY NumeBoRe DESC")
'       On Error Resume Next
'       BolreTemp.MoveFirst
'       If Err < 1 Then
'         If NROBOL& <= BolreTemp!NumeBore Then
'           NROBOL& = 1 + BolreTemp!NumeBore
'         End If
'       End If
'       On Error GoTo 0
'       BolreTemp.Close
'       '
'       Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
'       Call ABRESTOCKS
'       Call ABRECOMPRAS
'       '
'       AINDI% = 0
'       NORVENTA = ""
'       For YX& = 1 To ULTREG&("!BOLREDET")
'         YY$ = REGLEIDO$("!BOLREDET", YX&)
'         CI% = CVI(Mid$(YY$, 83, 2))
'         CAN# = CVD(Mid$(YY$, 89, 8))
'         NUOC& = CVS(Mid$(YY$, 151, 4))
'         NUMOIT% = ASC(Mid$(YY$, 98, 1))   'EL Nº DE ORDEN DEL ITEM
'         If CI% > 0 Then
'           If CAN# > 0.005 Then
'             CARECI = CAN#
'                '
'
'                  'If REGOC& > 0 Then
'                   ' If REGOC& <= ULTREG&("OCOMDETA") Then
'                     ' ZZ$ = REGLEIDO$("OCOMDETA", REGOC&)
'             COEFUNI = CVS(Mid$(ZZ$, 63, 4))
'             If COEFUNI < 0.0001 Then COEFUNI = 1
'             LOTECO$ = TRIM$(Mid$(YY$, 107, 12))
'             '
'             With OcomDeta
'               .FindFirst ("Nume_Ocom = " & NUOC& & " AND NuOrd_Item = " & NUMOIT%)
'               If .NoMatch = False Then
'                 .Edit
'                 If IsNull(!ICONTAB) Then
'                   IMPCONOC% = 0
'                 Else
'                    IMPCONOC% = !ICONTAB 'ASIGNO LA IMPUTACION CONTABLE
'                 End If
'                 !Cant_Rec = !Cant_Rec + CARECI / COEFUNI
'                 CAPENDIEN = !Cant_Ocom - !Cant_Rec
'                   NUORSE& = !N_OrServ
'                 If CAPENDIEN >= 0.005 Then
'                     If COMALTER%("Esta O/C por Código " + TRIM$(CODEXT$(CI%)) + "\Queda con un Saldo Pendiente de " + TRIM$(FORMATNUM$(CAPENDIEN, "F10.1")) + " " + TRIM$(UNIMED$(CI%)) + ".\\Mantener Abierta\Cancelar Pendiente") = 2 Then
'                       !Stat_Ocom = 8
'                     End If
'                   Else
'                     !Stat_Ocom = 3
'                 End If
'                 .Update
'               End If
'             End With
'             '///**NUEVO ACA EDITO EL REGISTRO EN MACONSIG***///
'             Call ABRESTOCKS
'             With MaConsig
'               .FindFirst (" VAL(MID$(NRO_Ocom,4,6)) = " & NUOC& & " AND Cod_Inte = " & CI% & " ")
'               If .NoMatch = False Then
'                 .Edit
'                   !Cant_Rendid = Cant_Rendid + CARECI / COEFUNI
'                     CAPENDIEN = !Cant_Entreg - !Cant_Rendid
'                   !Saldo_Pend = CAPENDIEN
'                 .Update
'               End If
'             End With
'            '////***HASTA ACA NUEVO****////
'
'             '
'             ' aca arma un string con las ordenes de compra recibidas
'             NORVE& = CVS(Mid$(YY$, 151, 4))
'             If NORVE& > 0 Then
'               If NORVE& <> NORVEA& Then
'                 NORVEA& = NORVE&
'                 If NORVENTA <> "" Then NORVENTA = NORVENTA + " - "
'                 NORVENTA = NORVENTA + TRIM$(Str$(NORVE&))
'               End If
'             End If
'             '
'             DESCRI$ = TRIM$(CODEXT$(CI%)) + " - " + TRIM$(DESCRIT$(CI%))
'                  If AINDI% = 0 Then
'                     X$ = REGBLAN$("INDIBOL")
'                     Call REPLA(X$, MKS$(NROBOL&), 1, 4)
'                     FEX = FF%
'                     Call REPLA(X$, FECHATEX$(FEX), 5, 8)
'                     Call REPLA(X$, REMI$, 15, 16)
'                     Call REPLA(X$, ECOARCH$("PROVED1", MKI$(NC%)), 32, 29)
'                     Call REGAPP("INDIBOL", X$)
'                     If NUORSE& < 1 Then Call REGAPP("IBOLPEN", X$)
'                     AINDI% = 1
'                  End If
'                  '
'             NBR$ = TRIM$(Str$(NROBOL&))
'             NORDE% = NORDE% + 1
'             NORDES$ = TRIM$(Str$(NORDE%))
'             IDENLO$ = "BR-" + NBR$ + "-" + NORDES$
'             '
'             Call REPLA(YY$, MKS$(NROBOL&), 1, 4)
'             If TIMATEI% <> 2 Then Call REPLA(YY$, DESCRI$, 5, 64)
'             '
'             If TRIM$(Mid$(YY$, 69, 12)) = "" Then
'               Call REPLA(YY$, IDENLO$, 69, 12)
'             End If
'             Call REPLA(YY$, MKI$(FF%), 81, 2)
'             DEPX% = DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(CI%)
'             Call REPLA(YY$, Chr$(DEPX%), 97, 1)
'             Call REPLA(YY$, MKI$(NC%), 105, 2)
'             Call REPLA(YY$, REMI$, 119, 16)
'             Call REPLA(YY$, MKS$(NUORSE&), 189, 4)
'             Call GRAREG("!BOLREDET", YY$, YX&)
'             '
'             With Bolrecep
'              .AddNew
'               !NumeBore = NROBOL&
'               !Descrip = Left$(DESCRI$, 64)
'               !IdenLote = Left$(IDENLO$, 12)
'               !FechRecep = CVDAT(FF%)
'               !Cod_Inte = CI%
'               !Cod_Exte = Left$(CODEXT$(CI%), 16)
'               !DescripItem = Left$(DESCRIT0$(CI%), 64)
'               !Unidad = Left$(UNIMED$(CI%), 4)
'               !CantNom = CAN#
'               !DepoEnt = DEPX%
'               !NUMPROV = NC%
'               !RaSo_Prov = Left$(RASOCLI$((-1) * NC%), 64)
'                  LOTECO$ = TRIM$(LOTECO$)
'                  While Len(LOTECO$) > 12
'                    LOTECO$ = Mid$(LOTECO$, 2)
'                  Wend
'               !LoteCol_Prov = Left$(LOTECO$, 12)
'               !NumRem_Prov = Left$(REMI$, 16)
'               !NumFac_Prov = ""
'               !NumOC = NUOC&
'               !ICONTAB = IMPCONOC% 'Imputación contable de la OC.
'               !MoneVaUni = ""
'               !Valo_Unit = 0
'               !CanToApro = 0
'               !CanToRech = 0
'               !CanToFalta = 0
'               !FechLibe = CVDAT(0)
'               !NumeOrdSer = NUORSE&
'               !Observa = " " 'OBSE$
'               !STATUS = 0
'               !USUAPRO = " "
'               !FeHoApro = CVDAT(0)
'              .Update
'             End With
'             '
'             If NUORSE& < 1 Then
'               If REGOC& > 0 Then
'                   Call DESCACONSI(NC%, CI%, CAN#)
'                 ElseIf REGOC& < 0 Then
'                   Call DEVOLCONSI(Abs(REGOC&), NC%, CI%, CAN#)
'               End If
'             End If
'             '
'             If NUORSE& > 0 Then
'               If CAN# > 0.05 Then
'                 GoSub BLANTEROP
'               End If
'             End If
'             '
'           End If
'         End If
'       Next YX&
'       '
'       Bolrecep.Close
'       OcomEnca.Close
'       OcomDeta.Close
''          For YX& = 1 To ULTREG&("!OBSERVOF")
''             YY$ = REGLEIDO$("!OBSERVOF", YX&)
''             YY$ = MKS$(NROBOL&) + YY$
''             Call REGAPP("OBSEBREC", YY$)
''          Next YX&
'          '
''          Call CIERRARCH("BOLRECEP")
''          Call CIERRARCH("BOLREPEN")
''          Call CIERRARCH("INDIBOL")
''          Call CIERRARCH("IBOLPEN")
''          Call CIERRARCH("OBSEBREC")
''          '
''          Call FRESHECHO("IBOLPEN")
''          Screen.MousePointer = 1
''          '
'       If TRIM$(FORIPRE$) <> "" Then
'         '
'         CANCELPRINT% = 0
'         If CONTROL("BOLRECEP", "SUPRINT") = "SI" Then CANCELPRINT% = 1
'         '
'         TIPODOC$ = "Boleta de Recepción"
'         DESTIDOC% = NC1%
'         NUMEDOC$ = Str$(NROBOL&)
'         FEX = FF%: FECHDOC$ = FECHATEX$(FEX)
'         '
'         CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = DEPOSAL$
'         CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = REMI$
'         CODPARAM$(3) = "N-AUTENT": DOCPARAM$(3) = NORVENTA
'         CAPARDOC% = 3
'         '
'         CODTABU1$(1) = "COD-MAT"
'         CODTABU1$(2) = "DES-MAT"
'         CODTABU1$(3) = "COLOR"
'         CODTABU1$(4) = "UNIMED"
'         CODTABU1$(5) = "NRO-LOTE"
'         CODTABU1$(6) = "CANTIDAD"
'         CODTABU1$(7) = "N-OCOMPR"
'         CODTABU1$(8) = "CAKILOS"
'         CODTABU1$(9) = "F-PLANO"
'         CODTABU1$(10) = "FECH-REC"
'         CACOLTAB1% = 10
'         '
'         CAITAB1% = 0
'         For YX& = 1 To ULTREG&("!BOLREDET")
'           YY$ = REGLEIDO$("!BOLREDET", YX&)
'           CI% = CVI(Mid$(YY$, 83, 2))
'           CAN# = CVD(Mid$(YY$, 89, 8))
'           If CI% > 0 Then
'             If CI% <= NUMAS% Then
'               If CAN# > 0.005 Then
'                 '
'                 UNID$ = TRIM$(Mid$(YY$, 85, 4))
'                 If UNID$ = "" Then UNID$ = UNIMED$(CI%)
'                 CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
'                 KILAJE$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 10, 3, 2)
'                 FEX = CVI(Mid$(YY$, 155, 2))
'                 FESOLEN$ = FECHATEX$(FEX)
'                 FEPLA$ = FEPLANO$(CI%)
'                 NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
'                 If NULOTE$ = "" Then
'                    NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
'                 End If
'                 NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
'                 '
'                 CAITAB1% = CAITAB1% + 1
'                 TETABU1$(1, CAITAB1%) = TRIM$(CODEXT$(CI%))
'                 '
'                 If FORMPLA$ = "RECEP-ESPUMA" Then
'                      DESMAT$ = Mid$(YY$, 5, 64)
'                    Else
'                      DESMAT$ = DESCRIT$(CI%)
'                 End If
'                 TETABU1$(2, CAITAB1%) = DESMAT$
'                 TETABU1$(3, CAITAB1%) = COLOR$(CI%)
'                 TETABU1$(4, CAITAB1%) = UNID$
'                 TETABU1$(5, CAITAB1%) = NULOTE$
'                 TETABU1$(6, CAITAB1%) = CANS$
'                 TETABU1$(7, CAITAB1%) = NUOCO$
'                 TETABU1$(8, CAITAB1%) = KILAJE$
'                 TETABU1$(9, CAITAB1%) = FEPLA$
'                 TETABU1$(10, CAITAB1%) = FESOLEN$
'                 '
'                 If TICOMPRO$ = "FICHA" Then
'                   Call PRINTDOC("FORMBREC")
'                 End If
'                 '
'               End If
'             End If
'           End If
'         Next YX&
'         '
'         If TICOMPRO$ <> "FICHA" Then
'           Call PRINTDOC("FORMBREC")
'         End If
'         '
'       End If
'          '
'1199 End If
'     '
'11999 Call CIERRARCH("*.*")
'     '
'     X$ = "": Y$ = "": YY$ = ""
'     EBOL$ = "": PEDY$ = "": DESCRI$ = ""
'     GoTo 1100
'     '
'BLANTEROP:
'   '
'   FINTER& = ULTREG&("OPERTER")
'   For U& = FINTER& To 1 Step -1
'     X0$ = REGLEIDO$("OPERTER", U&)
'     If CVS(Mid$(X0$, 75, 4)) >= NUORSE& Then
'       REGITER& = U&
'       GoTo 21
'     End If
'   Next U&
'   GoTo 29
'   '
'21 ' Marca como "Cumplida" Operación en OPERTER.DAT
'   SALRE# = CVD(Mid$(X0$, 79, 8)) - CAN#
'   If SALRE# < 0 Then SALRE# = 0
'   '
'   FEBA% = HOY(HOS$)
'   HOBA% = HORANUM%(Time$)
'   USUA% = USNBR% Mod 100
'   DBAJA$ = MKI$(FEBA%) + MKI$(HOBA%) + Chr$(USUA%) + Chr$(1) + "B.Rec." + TRIM$(Str$(NROBOL&))
'   '
'   NUOROP% = Asc(Mid$(X0$, 67, 1))
'   Call REPLA(X0$, MKD$(CAN#), 87, 8)
'   Call REPLA(X0$, MKD$(SALRE#), 95, 8)
'   Call REPLA(X0$, "3ROS  ", 151, 6)
'   Call REPLA(X0$, MKI$(FEBA%), 157, 2)
'   Call REPLA(X0$, MKI$(HORANUM%(Time$)), 159, 2)
'   Call REPLA(X0$, MKS$(0), 161, 4)
'   Call REPLA(X0$, String$(8, 0), 165, 8)
'   Call REPLA(X0$, MKS$(0), 173, 4)
'   Call REPLA(X0$, DBAJA$, 177, 22)
'   Call GRAREG("OPERTER", X0$, REGITER&)
'   '
'' Si se recibio un parcial, genera registro en OPERTER.DAT
'' por el saldo pendiente
'   If SALRE# > 0.01 Then
'      NUOROP% = 1 + Asc(Mid$(X0$, 67, 1))
'      CIIK% = CVI(Mid$(X0$, 13, 2))
'      NOPE% = CVI(Mid$(X0$, 15, 2))
'      DALTA$ = Chr$(NUOROP%) + Mid$(DBAJA$, 2, 5)
'      INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
'      PROGBLA$ = Space$(6) + String$(26, 0)
'      MINVARI = TIESTAN(CIIK%, NOPE%, SALRE#, NUOROP%)
'      '
'      Call REPLA(X0$, DALTA$, 67, 6)
'      Call REPLA(X0$, MKD$(SALRE#), 79, 8)
'      Call REPLA(X0$, MKD$(0), 87, 8)
'      Call REPLA(X0$, MKD$(SALRE#), 95, 8)
'      Call REPLA(X0$, String$(32, 0), 103, 32)
'      Call REPLA(X0$, MKS$(MINVARI), 111, 4)
'      Call REPLA(X0$, MKS$(MINVARI), 131, 4)
'      Call REPLA(X0$, String$(16, 0), 135, 16)
'      Call REPLA(X0$, INFOBLA$, 151, 48)
'      Call REPLA(X0$, PROGBLA$, 225, 32)
'      Call REGAPP("OPERTER", X0$)
'      AGREGA% = 1
'   End If
'   '
'29 Return
'   '
'   '
'End Sub
''

Sub GENIBOLPE(Optional MODO%)
    '
    Screen.MousePointer = 11
    Call ABRESTOCKS
    '
'    Dim MIRS As ADODB.Recordset
''10  STRSQL = "Select count(NumeBoRe) as CdadReg from BOLRECEP WITH(NOLOCK) "
'10  STRSQL = "Select count(*) as CdadReg from BOLRECEP WITH(NOLOCK) "
'    CONDI1$ = " WHERE Status <1 and Cod_Inte > 0 "
'    'If MODO% = 1 Then CONDI1$ = " WHERE Status <=1 and Cod_Inte > 0  "
'    STRSQL = STRSQL + CONDI1$
'    Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL))
'    If Not IsNull(MIRS!CdadReg) Then
'      FIN& = MIRS!CdadReg 'cantidad de registros
'    End If
'    MIRS.Close
'    Set MIRS = Nothing

'    ' recuperacion de información Dosivac
'    FELIMI = FETEXCOR1$(FECHANUM("22/08/08"))
'    If FIN& < 1 Or FIN& > 512 Then
'      If HOY(HOS$) < FECHANUM("23/08/08") Then
'        If VUELTA% < 1 Then
'           SQLZ$ = "UPDATE BOLRECEP SET STATUS = 1 WHERE STATUS < 2 AND FechRecep < '" & FELIMI & "' "
'           FLEXCONN.Execute SQLZ$
'           For U& = 1 To ULTREG&("BOLREPEN")
'             x$ = REGLEIDO$("BOLREPEN", U&)
'             APRO# = CVD(Mid$(x$, 163, 8))
'             RECHA# = CVD(Mid$(x$, 171, 8))
'             FALTA# = CVD(Mid$(x$, 179, 8))
'             If Abs(APRO#) < 0.005 Then
'               If Abs(RECHA#) < 0.005 Then
'                 If Abs(FALTA#) < 0.005 Then
'                    NOOB& = CVS(Left$(x$, 4))
'                    SQLZ$ = "UPDATE BOLRECEP SET STATUS = 0 WHERE NumeBoRe = " & NOOB&
'                    FLEXCONN.Execute SQLZ$
'                 End If
'               End If
'             End If
'           Next U&
'           VUELTA% = 1
'           GoTo 10
'        End If
'      End If
'    End If
    '
'    If FIN& < 1 Then
'        Call MENSERR(24, "No Existen Boletas de Recepcion Pendientes.")
'        On Error GoTo 0
'        JJ& = 0: K& = 0
'        GoTo 95
'    End If
    '
    SQLX$ = " SELECT * FROM BOLRECEP WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE Status < 1 "
    SQLX$ = SQLX$ + " AND Cod_Inte > 0 "   ' SOLO MATERIAL CODIFICADO
    SQLX$ = SQLX$ + " ORDER BY NumeBoRe "
    Dim COBRECEPTMP As ADODB.Recordset
    Set COBRECEPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With COBRECEPTMP
      II& = 0
      JJ& = 0
      k& = 0
      REBLA$ = REGBLAN$("BOLREPEN")
      REBLB$ = REGBLAN$("IBOLPEN")
      '
      Do While (.EOF = False)
         II& = II& + 1
         'Call TRACE(20, II&, FIN&)
         '
         NUORSE& = XVALO(!NUMEORDSER)
         If MODO% = 1 And NUORSE& > 0 Then GoTo 80
         '
         If XVALO(!CanToApro) > 0 Then GoTo 85
         If XVALO(!CanToRech) > 0 Then GoTo 85
         If XVALO(!CanToFalta) > 0 Then GoTo 85
         '
80       Y$ = REBLA$
         Call REPLA(Y$, MKS$(XVALO(!NUMEBORE)), 1, 4)
         Call REPLA(Y$, SAFETEXT$(!idenlote), 69, 12)
         Call REPLA(Y$, MKI$(CVINT(!FechRecep)), 81, 2)
            
         Call REPLA(Y$, MKI$(XVALO(!cod_inte)), 83, 2)
         Call REPLA(Y$, SAFETEXT$(!Unidad), 85, 4)
            
         Call REPLA(Y$, MKD$(XVALO(!CantNom)), 89, 8)
            
         Call REPLA(Y$, Chr$(XVALO(!DEPOENT)), 97, 1)
            
         Call REPLA(Y$, MKI$(XVALO(!NumProv)), 105, 2)
         Call REPLA(Y$, SAFETEXT$(!LOTECOL_PROV), 107, 12)
         Call REPLA(Y$, SAFETEXT$(!NumRem_Prov), 119, 16)
         Call REPLA(Y$, SAFETEXT$(!NumFac_Prov), 135, 16)
            
         Call REPLA(Y$, MKS$(XVALO(!NUMOC)), 151, 4)
         Call REPLA(Y$, SAFETEXT$(!MoneVaUni), 155, 4)
            
         Call REPLA(Y$, MKS$(XVALO(!Valo_Unit)), 159, 4)
            
         Call REPLA(Y$, MKD$(XVALO(!CanToApro)), 163, 8)
            
         Call REPLA(Y$, MKD$(XVALO(!CanToRech)), 171, 8)
         
         Call REPLA(Y$, MKD$(XVALO(!CanToFalta)), 179, 8)
         Call REPLA(Y$, MKI$(CVINT(!FechLibe)), 187, 2)
         
           NAENT& = XVALO(!Nume_Aten)
         Call REPLA(Y$, MKS$((NAENT&)), 189, 4) '151, 4)'\\\OT-06-0241-WAR-14/06/06///
         Call REPLA(Y$, SAFETEXT$(!Observa), 193, 64)
         JJ& = JJ& + 1
         Call GRAREG("BOLREPEN", Y$, JJ&)
         '
         If XVALO(!NUMEBORE) <> ANTNUM Then
           XX$ = REBLB$
           Call REPLA(XX$, MKS$(XVALO(!NUMEBORE)), 1, 4)
           FEX$ = Format(!FechRecep, "DD/MM/YY")
           Call REPLA(XX$, FEX$, 5, 8)
           Call REPLA(XX$, SAFETEXT$(!NumRem_Prov), 15, 16)
           Call REPLA(XX$, ECOARCH$("PROVED1", MKI$(XVALO(!NumProv))), 32, 29)
           '///OT-8-667-04/09/08 - SE INCLUYE CODIGO PARA QUE OBVIE LAS RECEPCIONES
           ' PROVENIENTES DE ORDENES DE SERVICIO
           'NUORSE& = XVALO(!NUMEORDSER)
           k& = k& + 1
           Call GRAREG("IBOLPEN", XX$, k&)
'           If NUORSE& < 1 Or (MODO% = 1 And NUORSE& > 0) Then
'             K& = K& + 1
'             Call GRAREG("IBOLPEN", XX$, K&)
'           End If
           '///OT-8-667-FIN
           ANTNUM = XVALO(!NUMEBORE)
         End If
         '
85    .MoveNext
      Loop
    End With
    '
90  COBRECEPTMP.Close
    Set COBRECEPTMP = Nothing
    '
95  Call SETULTREG("BOLREPEN", JJ&)
    Call CIERRARCH("BOLREPEN")
    Call SETULTREG("IBOLPEN", k&)
    Call CIERRARCH("IBOLPEN")
    Call FRESHECHO("IBOLPEN")
    '
    Screen.MousePointer = 1
    '
    
End Sub

Sub REPIBOL()
        '
        Dim REGBO&(2048)
        '
        TICOMPRO$ = TRIM$(Control$("", "COMPBREC"))
        
1120    'Call CLEARVEN
        X$ = "": Y$ = "": Z$ = "": YY$ = ""
        EBOL$ = "": PEDY$ = "": DESCRI$ = ""
        NBR$ = OBJPLA$("SELBOLET", "")
        If NBR$ = "" Then
            Exit Sub
        End If
        '
        NROBOL& = CVS(Left$(NBR$, 4))
        If NROBOL& < 1 Then
            Exit Sub
        End If
        '
        Screen.MousePointer = 11
        FIN& = ULTREG&("BOLRECEP")
        Call SETULTREG("!BOLREDET", 0)
        Call CIERRARCH("!BOLREDET")
        Call SETULTREG("!OBSERVOF", 0)
        Call CIERRARCH("!OBSERVOF")
        '
        For U& = 1 To ULTREG&("BOLRECEP")
            X$ = REGLEIDO$("BOLRECEP", U&)
            If CVS(Left$(X$, 4)) = NROBOL& Then
                NC% = CVI(Mid$(X$, 7, 2))
                Call REGAPP("!BOLREDET", X$)
            End If
        Next U&
        '
        For U& = 1 To ULTREG&("OBSEBREC")
            X$ = REGLEIDO$("OBSEBREC", U&)
            If CVS(Left$(X$, 4)) = NROBOL& Then
                Call REGAPP("!OBSERVOF", Mid$(X$, 5, 72))
            End If
        Next U&
        '
        Screen.MousePointer = 1
        '
        Call CIERRARCH("!BOLREDET")
        Call CIERRARCH("!OBSERVOF")
        '
        If ULTREG&("!BOLREDET") < 1 Then
            Call COMUNI("Boleta de Recepcion Nro." + Str$(NROBOL&) + "\no Figura en la Base de Datos.")
            GoTo 1120
        End If
        '
        X$ = REGLEIDO$("TATIREC", 1)
        FORMPLA$ = TRIM$(UCase$(Mid$(X$, 33, 16)))
        FORMTAB$ = TRIM$(UCase$(Mid$(X$, 49, 16)))
        FORIPRE$ = TRIM$(UCase$(Mid$(X$, 65, 16)))
        '
        If TRIM$(FORIPRE$) = "" Then
            Call COMUNI("No se Definió Formulario de Impresión de B.R.")
            GoTo 1120
        End If
        '
        YY$ = REGLEIDO$("!BOLREDET", 1)
        NC% = CVI(Mid$(YY$, 105, 2))
        NC1% = (-1) * NC%
        FF% = CVI(Mid$(YY$, 81, 2))
        DP1% = Asc(Mid$(YY$, 97, 1))
        DEPOSAL$ = TRIM$(ECOARCH$("ECODEP", Chr$(DP1%)))
        REMI$ = Mid$(YY$, 119, 16)
        '
        If TRIM$(FORIPRE$) <> "" Then
               '
               TIPODOC$ = "Boleta de Recepción"
               DESTIDOC% = NC1%
               NUMEDOC$ = Str$(NROBOL&)
               FEX = FF%: FECHDOC$ = FECHATEX$(FEX)
               '
               CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = DEPOSAL$
               CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = REMI$
               CAPARDOC% = 2
               '
               CODTABU1$(1) = "COD-MAT"
               CODTABU1$(2) = "DES-MAT"
               CODTABU1$(3) = "COLOR"
               CODTABU1$(4) = "UNIMED"
               CODTABU1$(5) = "NRO-LOTE"
               CODTABU1$(6) = "CANTIDAD"
               CODTABU1$(7) = "N-OCOMPR"
               CODTABU1$(8) = "CAKILOS"
               CODTABU1$(9) = "F-PLANO"
               CACOLTAB1% = 9
               '
               CAITAB1% = 0
               For YX& = 1 To ULTREG&("!BOLREDET")
                  YY$ = REGLEIDO$("!BOLREDET", YX&)
                  CI% = CVI(Mid$(YY$, 83, 2))
                  CAN# = CVD(Mid$(YY$, 89, 8))
                  If CI% > 0 Then
                     If CI% <= NUMAS% Then
                        If CAN# > 0.005 Then
                           '
                           UNID$ = TRIM$(Mid$(YY$, 85, 4))
                           If UNID$ = "" Then UNID$ = Unimed$(CI%)
                           CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
                           KILAJE$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 10, 3, 2)
                           FEPLA$ = FEPLANO$(CI%)
                           NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
                           If NULOTE$ = "" Then
                              NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
                           End If
                           NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
                           '
                           CAITAB1% = CAITAB1% + 1
                           TETABU1$(1, CAITAB1%) = TRIM$(CODEXT$(CI%))
                           '
                           If FORMPLA$ = "RECEP-ESPUMA" Then
                                DESMAT$ = Mid$(YY$, 5, 64)
                              Else
                                DESMAT$ = DESCRIT$(CI%)
                           End If
                           TETABU1$(2, CAITAB1%) = DESMAT$
                           TETABU1$(3, CAITAB1%) = COLOR$(CI%)
                           TETABU1$(4, CAITAB1%) = UNID$
                           TETABU1$(5, CAITAB1%) = NULOTE$
                           TETABU1$(6, CAITAB1%) = CANS$
                           TETABU1$(7, CAITAB1%) = NUOCO$
                           TETABU1$(8, CAITAB1%) = KILAJE$
                           TETABU1$(9, CAITAB1%) = FEPLA$
                           '
                           If TICOMPRO$ = "FICHA" Then
                               Call PRINTDOC("FORMBREC")
                           End If
                           '
                        End If
                     End If
                  End If
               Next YX&
               '
               If TICOMPRO$ <> "FICHA" Then
                   Call PRINTDOC("FORMBREC")
               End If
               '
1199    End If
        '
        Call CIERRARCH("*.*")
        '
        X$ = "": Y$ = "": YY$ = ""
        EBOL$ = "": PEDY$ = "": DESCRI$ = ""
        GoTo 1120
        '
End Sub
'
Function EXIREMI(NR1$, NC1%)
    '
    Call ABRESTOCKS
    SQLX$ = "SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + " WHERE NumRem_Prov = '" & TRIM(NR1$) & "' "
    SQLX$ = SQLX$ + " AND NumProv = " & NC1% & ""
    'Set EXITMP = Stocks.OpenRecordset(sqlx$, 4)
    
    EXIREMI = 0
    '
    Dim EXITMP As ADODB.Recordset
    Set EXITMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    If Not (EXITMP.EOF And EXITMP.BOF) Then
      EXIREMI = 1
    End If
    If Err Then
        On Error GoTo 0
        Exit Function
    End If
    
    '
    'With EXITMP
    '  On Error Resume Next
    '  .MoveFirst
    '  If Err Then
    '    On Error GoTo 0
    '    Exit Function
    '  End If
    '  On Error GoTo 0
    '  EXIREMI = 1
    'End With
    '
End Function
'
Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub
'
Sub SHOWOCOM(NPX&)
    '
    TIDOCUM$ = "Nota de Pedido"
    If NPX& < 0 Then
       TIDOCUM$ = "Nota de Pedido - Abierta"
    End If
    NUDOCU$ = TRIM$(Str$(Abs(NPX&)))
    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
    NUDOCU1$ = " " + NUDOCU$ + " "
    NUDOCU2$ = " " + NUDOCU$ + "-"
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
          DOCUNU& = CVS(Left$(XX$, 4))
          GoTo 5
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    Exit Sub
    '
5   Call MUESTRADOC(DOCUNU&)

End Sub
'
Sub SHOWBORE(NPX&)
    '
    TIDOCUM$ = "Boleta de Recepción"
    NUDOCU$ = TRIM$(Str$(NPX&))
    Screen.MousePointer = 11
    Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$, NUDOCU$)
    Exit Sub
    '
''    For UI& = ULTREG&("PDOCLST") To 1 Step -1
''      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
''      If InStr(XX$, TIDOCUM$) > 4 Then
''        If InStr(XX$, NUDOCU$) > 4 Then
''          DOCUNU& = CVS(Left$(XX$, 4))
''          GoTo 5
''        End If
''      End If
''    Next UI&
''    Screen.MousePointer = 1
''    Call COMUNI("Documento no Encontrado")
''''    Exit Sub
''    '
''5   Call MUESTRADOC(DOCUNU&)
    '
End Sub
'
Sub COBORECEP()
    '
    'MARCA PARA SABER SI HAY MATERIALES DE TEXTO LIBRE
    HAYMATERIALIB% = 0
    '
    Call ABRESTOCKS
    '
    SQLX$ = " SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + " WHERE Status = 0 "
    SQLX$ = SQLX$ + " ORDER BY NumeBoRe "
    '
    'Set COBRECEPTMP = Stocks.OpenRecordset(sqlx$, dbOpenSnapshot)
    '
    Dim COBRECEPTMP As ADODB.Recordset
    Set COBRECEPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    
    With COBRECEPTMP
      JJ& = 0
      k& = 0
      TC& = 0
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Boletas de Recepcion Pendientes.")
        On Error GoTo 0
        GoTo 90
      End If
      On Error GoTo 0
      Do While (.EOF = False)
            Y$ = REGBLAN("BOLREPEN")
            '
            Call REPLA(Y$, MKS$(!NUMEBORE), 1, 4)
            Call REPLA(Y$, !idenlote, 69, 12)
            Call REPLA(Y$, MKI$(CVINT(!FechRecep)), 81, 2)
            Call REPLA(Y$, MKI$(!cod_inte), 83, 2)
            Call REPLA(Y$, !Unidad, 85, 4)
            Call REPLA(Y$, MKD$(!CantNom), 89, 8)
            Call REPLA(Y$, Chr$(!DEPOENT), 97, 1)
            Call REPLA(Y$, MKI$(!NumProv), 105, 2)
            Call REPLA(Y$, !LOTECOL_PROV, 107, 12)
            Call REPLA(Y$, !NumRem_Prov, 119, 16)
            Call REPLA(Y$, !NumFac_Prov, 135, 16)
            Call REPLA(Y$, MKS$(!NUMOC), 151, 4)
            Call REPLA(Y$, !MoneVaUni, 155, 4)
            Call REPLA(Y$, MKS$(!Valo_Unit), 159, 4)
            Call REPLA(Y$, MKD$(!CanToApro), 163, 8)
            Call REPLA(Y$, MKD$(!CanToRech), 171, 8)
            Call REPLA(Y$, MKD$(!CanToFalta), 179, 8)
            Call REPLA(Y$, MKI$(CVINT(!FechLibe)), 187, 2)
            Call REPLA(Y$, !Observa, 193, 64)
            '
            If !cod_inte <= 0 Then HAYMATERIALIB% = 1 'HAY MATERIAL TEXTO LIBRE
            '
            JJ& = JJ& + 1
            If !cod_inte > 0 Then
                 TC& = TC& + 1
                Call GRAREG("BOLREPEN", Y$, TC&)
            End If
            '
            XXZ$ = REGBLAN("BOREPELI")
            XXZ$ = Left$(Y$, 82) + Space$(174) + Mid$(Y$, 85)
            If !cod_inte <= 0 Then
                'Call REPLA(XXZ$, MKI$(!COD_INTE), 9, 2)
                Call REPLA(XXZ$, !Cod_Exte, 83, 16)
                Call REPLA(XXZ$, !descripitem, 99, 64)
              Else
               CIXI% = !cod_inte
               Call REPLA(XXZ$, CODEXT$(CIXI%), 83, 16)
               Call REPLA(XXZ$, DESCRIT0$(CIXI%), 99, 64)
            End If
            '
            NUOCO1& = !NUMOC
            'Call ABRECOMPRAS
            strSQL = "Select Nume_Ocom,Nro_Otrab,Nro_grup from Ocomdeta"
            strSQL = strSQL + " where Nume_Ocom=" & NUOCO1&
            Dim OCOMDETA As ADODB.Recordset
            Set OCOMDETA = FLEXCONN.Execute(FILTSQL$(strSQL))
            With OCOMDETA 'PARA OBTENER EL DESTINO Y GRUPO
            '    sqlx$ = " NUME_OCOM = " & NUOCO1& & ""
            '    .FindFirst sqlx$
            '    If .NoMatch = False Then
                If Not (.EOF And .BOF) Then      '
                  If IsNull(!Nro_Otrab) Then
                     NOTR$ = ""
                  Else
                     NOTR$ = !Nro_Otrab
                  End If
                  If IsNull(!Nro_Grup) Then
                      NGR$ = ""
                  Else
                      NGR$ = !Nro_Grup
                  End If
                  DEST1$ = NOTR$
                  If DEST1$ <> "" Then
                    If NGR$ <> "" Then
                      DEST1$ = DEST1$ + " / "
                    End If
                  End If
                  DEST1$ = DEST1$ + NGR$
                  Call REPLA(XXZ$, DEST1$, 163, 24)
                  '
                End If
                '
            End With
            '
            Call GRAREG("BOREPELI", XXZ$, JJ&)
            '
            '
            If !NUMEBORE <> ANTNUM Then
              XX$ = REGBLAN$("IBOLPEN")
            '
              Call REPLA(XX$, MKS$(!NUMEBORE), 1, 4)
              FEX$ = Format(!FechRecep, "DD/MM/YY")
              Call REPLA(XX$, FEX$, 5, 8)
              Call REPLA(XX$, !NumRem_Prov, 15, 16)
              Call REPLA(XX$, ECOARCH$("PROVED1", MKI$(!NumProv)), 32, 29)
              k& = k& + 1
              Call GRAREG("IBOLPEN", XX$, k&)
              ANTNUM = !NUMEBORE
           End If
           '
      .MoveNext
      Loop
    End With
    COBRECEPTMP.Close
    Set COBRECEPTMP = Nothing
    '
90  Call SETULTREG("BOLREPEN", JJ&)
    Call CIERRARCH("BOLREPEN")
    Call SETULTREG("IBOLPEN", k&)
    Call CIERRARCH("IBOLPEN")
    
    '
End Sub

Sub Selcopen()
    '
    Screen.MousePointer = 11
    Call COPYSTRU("INDIOCOM", "INDIOCPE")
    Call ABRECOMPRAS
    SQLX$ = " SELECT  DISTINCT OCOMENCA.Nume_Ocom, OCOMENCA.Refe_Ocom "
    SQLX$ = SQLX$ + " FROM OCOMENCA "
    SQLX$ = SQLX$ + " INNER JOIN  OCOMDETA "
    SQLX$ = SQLX$ + " ON (OCOMENCA.Nume_Ocom = OCOMDETA.Nume_Ocom) "
    SQLX$ = SQLX$ + " WHERE OCOMDETA.Stat_Ocom < 8 "   ' para que no este anulada ni cumplida
    SQLX$ = SQLX$ + " AND OCOMDETA.Cant_Rec < (OCOMDETA.Cant_Ocom - 0.005) "
    SQLX$ = SQLX$ + " AND OCOMENCA.Stat_Ocom < 1 "
    SQLX$ = SQLX$ + " AND OCOMDETA.Stat_Ocom >= 0 " 'YA APROBADAS
    SQLX$ = SQLX$ + " AND OCOMENCA.Stat_Ocom >= 0 " 'YA APROBADAS
    SQLX$ = SQLX$ + " ORDER BY OCOMENCA.Nume_Ocom  ASC "
    'Set OCMODIFTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim OCMODIFTEMP As ADODB.Recordset
    Set OCMODIFTEMP = New ADODB.Recordset
    OCMODIFTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    With OCMODIFTEMP
       On Error Resume Next
        .MoveFirst
        If Err Then
          Call MENSERR(24, "No Hay Ordenes de Compra\Pendientes.")
          Screen.MousePointer = 1
          On Error GoTo 0
          Exit Sub
        End If
        On Error GoTo 0
        '
        JJ& = 0
        Do While (.EOF = False)
            XXY$ = REGBLAN$("INDIOCPE")
            Call REPLA(XXY$, MKS(!nume_ocom), 1, 4) 'Nº O/C
            Call REPLA(XXY$, !Refe_Ocom, 5, 44) ' REFERENCIA
            JJ& = JJ& + 1
            Call GRAREG("INDIOCPE", XXY$, JJ&)
            .MoveNext
        Loop
    End With
    Call SETULTREG("INDIOCPE", JJ&)
    Call CIERRARCH("INDIOCPE")
    '
End Sub

Sub APROBOL()
        '
        ' MODIFICADO EL 30/01/06 PARA QUE LA COINCIDENCIA SEA POR NUMERO DE br + NUMERO REMITO PROVEEDOR
        '
        Dim REGBO&(2048)
        '
1100    Call GENINDIBOL(2) 'Call GENIBOLPE(0)
        '
        EPAC$ = UCase$(TRIM$(EMPREAC$))
        SOYCAMBACERES% = 0
        If InStr(1, EMPREAC$, "CAMBACERES") > 0 Then SOYCAMBACERES% = 1

1120    Call FRESHECHO("IBOPEAN")
        Screen.MousePointer = 1
        'Call CLEARVEN
        If ULTREG&("IBOPEAN") < 1 Then
          Exit Sub
        End If
        '
        X$ = "": Y$ = "": Z$ = "": YY$ = ""
        EBOL$ = "": PEDY$ = "": DESCRI$ = ""
        Call SELECHO("IBOPEAN")
        NBR1$ = TRIM$(VALACT1$("IBOPEAN"))
        If NBR1$ = "" Then
          Exit Sub
        End If
        NBR2$ = VALACT2$("IBOPEAN")
        NUREPRO$ = TRIM$(Mid$(NBR2$, 11, 16))
        '
        NROBOL& = Val(NBR1$)
        If NROBOL& < 1 Then
            Exit Sub
        End If
        '
        ESIMPO% = VALIDAR_DESPA_IMPO%(NROBOL&)
        If ESIMPO% < 0 Then
            'DEVUELVE -1 SI ES DE EXTERIOR,ESTA CONFIGURADO Y LE FALTAN LOS DATOS DE DESPACHO
            'DEVUELVE 0 SI NO ES DEL EXTERIOR O NO ESTA CONFIGURADO
            'DEVUELVE 1 SI ESTA CONFIGURADO, ES DEL EXTERIOR Y TIENE TODOS LOS DATOS
            Exit Sub
        ElseIf ESIMPO% = 1 Then
           DEPOSITO_EQ_IMP% = XVALO(Control("BOLRECEP", "DEPEQIMP"))
           If DEPOSITO_EQ_IMP% <= 0 Then
              Call COMUNI("Falta Definir Depósito de Ingreso de Equipos Importados")
              Exit Sub
           End If
        End If
        Call GENERABOLREPEN(NROBOL&)
        '
        FIN& = ULTREG&("BOLREPEN")
        Call COPYSTRU("BOLREDET", "BOLRESEL")
        '
        For KU& = 1 To NUMAS%
          CIMOVI%(KU&) = 0
        Next KU&
        '
        CABOPE% = 0: JJ& = 0
        For i& = 1 To FIN&
          X$ = REGLEIDO$("BOLREPEN", i&)
          If CVS(Left$(X$, 4)) = NROBOL& Then
            If TRIM$(Mid$(X$, 119, 16)) = NUREPRO$ Then
              CANOM# = CVD(Mid$(X$, 89, 8))
              APRO# = CVD(Mid$(X$, 163, 8))
              RECHA# = CVD(Mid$(X$, 171, 8))
              FALTA# = CVD(Mid$(X$, 179, 8))
              '
              If CANOM# < 0.05 Or APRO# > 0.05 Or RECHA# > 0.05 Or Abs(FALTA#) > 0.05 Then GoTo 1121
              Call REPLA(X$, MKD$(CANOM#), 163, 8)
              '
              IDLO$ = TRIM$(Mid$(X$, 69, 12))
825           If Len(IDLO$) > 0 Then
                If InStr(IDLO$, "-") > 0 Then
                  IDLO$ = Mid$(IDLO$, 2)
                  GoTo 825
                End If
              End If
              Call REPLA(X$, MKS$(Val(IDLO$)), 1, 4)
              '
              If CABOPE% = 0 Then
                NC% = CVI(Mid$(X$, 105, 2)): NC1% = (-1) * NC%
                FF% = CVI(Mid$(X$, 81, 2))
                FEX = FF%: FECHAREC$ = FECHATEX$(FEX)
                'REMI$ = TRIM$(Mid$(X$, 119, 16))
                DP1% = Asc(Mid$(X$, 97, 1))
                If ECOARCH$("ECODEP", Chr$(DP1%)) = "" Then
                    DP1% = 0
                End If
                '
                Y$ = String$(128, 0)
                Call REPLA(Y$, MKS$(NROBOL&), 1, 4)
                Call REPLA(Y$, MKI$(FF%), 5, 2)
                Call REPLA(Y$, MKI$(NC%), 7, 2)
                Call REPLA(Y$, rasocli$(NC1%), 9, 32)
                Call REPLA(Y$, DOMICLI$(NC1%), 41, 32)
                    CLOCA$ = TRIM$(CPOSCLI$(NC1%)): If CLOCA$ <> "" Then CLOCA$ = CLOCA$ + " "
                Call REPLA(Y$, CLOCA$ + LOCACLI$(NC1%), 73, 32)
                Call REPLA(Y$, NUREPRO$, 105, 16)
                Call REPLA(Y$, Chr$(DP1%), 121, 1)
                AINDI% = AINDI% + 1
              End If
              '
              CABOPE% = CABOPE% + 1: REGBO&(CABOPE%) = i&
              JJ& = JJ& + 1: Call GRAREG("BOLRESEL", X$, JJ&)
              '
            End If
           End If
          '
1121    Next i&
        Call SETULTREG("BOLRESEL", JJ&)
        '
        If CABOPE% < 1 Then
            Call COMUNI("No hay Recepciones Pendientes\para la Boleta Elegida")
            GoTo 1120
        End If
        '
        Call FILESORT("BOLRESEL", "", 1, 1, "ASC")
        Call TRALOCAL("BOLRESEL", "BOLREDET")
        '
11215   Z$ = OBJPLA$("BOLREC-MUESTRW", Y$)
        If Z$ = "" Then
            GoTo 1120
        End If
        DP1% = Asc(Mid$(Z$, 121, 1))
        If DP1% < 1 Or ECOARCH$("ECODEP", Chr$(DP1%)) = "" Then
          Call MENSERR(24, "Falta Indicar Depósito de Entrada")
          GoTo 11215
        End If
        '
        Call ABRESTOCKS
        'ABUS$ = "NumeBoRe = " & NROBOL & " AND NumRem_Prov = '" & NUREPRO$ & "' "
        'With BolRecep
        '  .FindFirst ABUS$
        '  If .NoMatch = True Then
        '    Call MENSERR(24, "Boleta de Recepción no Encontrada.")
        '    GoTo 1100
        '  End If
          '
        '  CIXBR% = !Cod_Inte
        '  RESP5% = 1  ' Material Codificado
        '  If CIXBR% = 0 Then RESP5% = 2    ' mATERIAL lIBRE
        'End With
        
        
        strSQL = "Select Cod_Inte from BolRecep where NumeBoRe=" & NROBOL
        strSQL = strSQL + " And NumRem_Prov = '" & NUREPRO$ & "'"
        Dim Bolrecep As ADODB.Recordset
        Set Bolrecep = New ADODB.Recordset
        Set Bolrecep = FLEXCONN.Execute(FILTSQL$(strSQL))
        If (Bolrecep.EOF And Bolrecep.BOF) Then
            Call MENSERR(24, "Boleta de Recepción no Encontrada.")
            GoTo 1100
        Else
          CIXBR% = Bolrecep!cod_inte
          RESP5% = 1  ' Material Codificado
          If CIXBR% = 0 Then RESP5% = 2    ' mATERIAL lIBRE
        End If
        Bolrecep.Close
        Set Bolrecep = Nothing
        '
        '
        If RESP5% = 2 Then
           Call APROCLI(NC%, FF%, NROBOL&, NUREPRO$, DP1%)
           Screen.MousePointer = 1
           GoTo 1100
        End If
        'HASTA ACA NUEVO
        '
1122    AVENTA$ = "APROBOL"
        If Control$("BOLRECEP", "CERCAPRO") = "SI" Then AVENTA$ = "APROBOLG"
        VTB% = VENTABU%(AVENTA$ + "/NC/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(rasocli$(NC1%)) + " - " + FECHAREC$)
        If VTB% < 0 Then GoTo 1120
        '
        ATX% = 0: HAYRECHA% = 0
        For U& = 1 To ULTREG&("!BOLREDET")
          X$ = REGLEIDO$("!BOLREDET", U&)
          CANOM# = CVD(Mid$(X$, 89, 8))
          APRO# = CVD(Mid$(X$, 163, 8))
          RECHA# = CVD(Mid$(X$, 171, 8))
            If RECHA# >= 0.05 Then HAYRECHA% = 1
          FALTA# = CVD(Mid$(X$, 179, 8))
          CERCAPRO$ = TRIM$(Mid$(X$, 241, 16))
          If Abs(APRO#) < 0.005 Then
            If Abs(RECHA#) < 0.005 Then
              If Abs(FALTA#) < 0.005 Then
                 GoTo 1123                        ' NO SE PROCESO MOVIMIENTO
              End If
            End If
          End If
          '
          FALTB# = CANOM# - APRO# - RECHA#
          If Abs(FALTB# - FALTA#) > 0.005 Then
            ATX% = ATX% + 1
            Call REPLA(X$, MKD$(FALTB#), 179, 8)
            Call GRAREG("!BOLREDET", X$, U&)
          End If
1123    Next U&
        '
        If ATX% > 0 Then
          Call COMUNI("Cantidades Incoherentes.\Revise y Corrija !")
          GoTo 1122
        End If
        '
'        ' ASIGNACION DE UBICACIONES AL MATERIAL APROBADO
'        If InStr(1, EMPREAC$, "GABAL") > 0 Then
'            For UU& = 1 To ULTREG&("!BOLREDET")
'               x$ = REGLEIDO$("!BOLREDET", U&)
'               CIXI% = CVI(Mid$(x$, 83, 2))
'               APRO# = CVD(Mid$(x$, 163, 8))
'               IDENLO$ = Mid$(x$, 69, 12)
'               UB$ = UBISTONUE$(CIXI%, CAN#)
'               If UB$ <> "" Then Call GRABAUBICACION(CIXI%, APRO#, UB$, IDENLO$)
'               UB$ = AJUSTI$(UB$, 16)
'               Call REPLA(x$, UB$, 193, 16)
'               Call GRAREG("!BOLREDET", x$, UU&)
'            Next UU&
'        End If
        '
1124    If VTB% < 0 Then GoTo 1120
        '
        PORECHLIN% = 0
        CODRECHA$ = "RD"
        REMIRECH$ = "RD." + TRIM$(Str$(NROBOL&))
        If HAYRECHA% > 0 Then
          '\\\OT-05-0723-WAR-23/11/05///
          MODORECHA% = 0
          
          If InStr(1, EMPREAC$, "FERVI") > 0 Then GoTo 1125
          
          If EXISTE%("RECHLI07.EXE") > 0 Then
             If HAYRECHA% > 1 Then
               Call MENSERR(24, "Imposible Procesar más de un Rechazo.")
               GoTo 1120
             End If
             PORECHLIN% = 1
             FERECHA% = CVI(Mid$(X$, 81, 2))
             CIRECHA% = CVI(Mid$(X$, 83, 2))
             CARECHA = RECHA#
             GoTo 1125
          Else
             If ULTREG&("TAMOTIV") > 0 Then 'VERIFICA SI HAY RECHAZOS
                FIN& = ULTREG&("!BOLREDET") 'Y SI HAY MUESTRA EL FRM DE MOTIVOS
                For i& = 1 To FIN&          ' DE RECHAZOS
                   XX$ = REGLEIDO$("!BOLREDET", i&)
                   CI% = CVI(Mid$(XX$, 83, 2))
                   If CI% > 0 Then
                     RECH# = CVD(Mid$(XX$, 171, 8))
                     If RECH# >= 0.05 Then
                      IDLO$ = Mid$(XX$, 69, 12)
                      Call Mot_Rechazo(NROBOL&, CI%, IDLO$, RECH#, MOTIRECHA%, DESTRECHA%, RECH_OBSERVA$)
                      Call REPLA(XX$, Chr(MOTIRECHA%), 99, 1)
                      Call REPLA(XX$, Chr(DESTRECHA%), 104, 1)
                      Call REPLA(XX$, RECH_OBSERVA$, 193, 48)
                      Call GRAREG("!BOLREDET", XX$, i&)
                     End If
                   End If
                Next i&
            End If
          End If
          '
          If ULTREG&("TADESTI") > 0 Then
            If ULTREG&("TAMOTIV") > 0 Then
              MODORECHA% = 1
            End If
          End If
          '\\\OT-05-0723-WAR-FIN///
          If TRIM$(Control$("", "REMIRECH")) <> "" Then
            CODRECHA$ = TRIM$(Control$("", "CODRRECH"))
            If CODRECHA$ = "" Then CODRECHA$ = "RD"
          End If
        End If
        '
1125    QUEDAPEN% = 0
        HAYRECHA% = 0
        Call ABRESTOCKS
        NOCLOSE% = 1: NOARRAS% = 1
        FIN& = ULTREG&("!BOLREDET")
        '
        NOCLOSE% = 1: NOARRAS% = 1
        Label12.Visible = True
        '
        CREALOTE% = (-1)
        For U& = 1 To FIN&
          Call TRACE(20, U&, FIN&)
          Label12.Caption = TRIM$(Str$(U&))
          DoEvents
          X$ = REGLEIDO$("!BOLREDET", U&)
          CI% = CVI(Mid$(X$, 83, 2))
            CIMOVI%(CI%) = 1
            ESTRAX% = ESTRAZABLE%(CI%, 1)
            If ESTRAX% = 1 Then CREALOTE% = 1
          REFX$ = Mid$(X$, 5, 64)
          LOTE$ = Mid$(X$, 69, 12)
          CANOM# = CVD(Mid$(X$, 89, 8))
          APRO# = CVD(Mid$(X$, 163, 8))
          RECHA# = CVD(Mid$(X$, 171, 8))
          FALTA# = CVD(Mid$(X$, 179, 8))
          LOTPROV$ = Mid$(X$, 107, 12)      'LOTE DEL PROVEEDOR
          MOTIRECHA% = Asc(Mid$(X$, 99, 1))
          DESTRECHA% = Asc(Mid$(X$, 104, 1))
          OBSERVARECHAZ$ = RTrim(Mid$(X$, 193, 48))
          UB$ = TRIM$(Mid$(X$, 135, 16)) ' GABAL - UBICACION ASIGNADA
          ' pos% = InStr(OBSRVRE$, ";")
          ' NUMERECHA% = XVALO(Mid$(OBSRVRE$, 1, pos% - 1))
          ' OBSERVARECHAZ$ = Mid$(OBSRVRE$, pos% + 1)

          If Abs(APRO#) < 0.005 Then
            If Abs(RECHA#) < 0.005 Then
              If Abs(FALTA#) < 0.005 Then
                QUEDAPEN% = 1
                GoTo 1128                        ' NO SE PROCESO MOVIMIENTO
              End If
            End If
          End If
          '
          SQLX$ = " SELECT * FROM BOLRECEP "
          SQLX$ = SQLX$ + " WHERE Status = 0 "
          SQLX$ = SQLX$ + " AND NumeBoRe = " & NROBOL& & " "
          SQLX$ = SQLX$ + " AND NumRem_Prov = '" & NUREPRO$ & "' "
          SQLX$ = SQLX$ + " AND IdenLote = '" & LOTE$ & "' "
          SQLX$ = SQLX$ + " AND Cod_Inte = " & CI%
          
          Dim COBRECEPTMP As ADODB.Recordset
          Set COBRECEPTMP = New ADODB.Recordset
          COBRECEPTMP.CursorType = adOpenKeyset
          COBRECEPTMP.LockType = adLockPessimistic
          COBRECEPTMP.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
          '
          With COBRECEPTMP
            If Not (.EOF And .EOF) Then
              !CanToApro = APRO#
              !CanToRech = RECHA#
              !CanToFalta = FALTA#
              !Status = 1
              !FechLibe = Date
              !USUAPRO = USERNAME   'NUEVO
              !FeHoApro = Now       'NUEVO
              !Observa = USERNAME   'nuevo
              !CerCal_Prov = CERCAPRO$
              '
              NANTEN& = 0
              On Error Resume Next
              NANTEN& = !Nume_Aten
              On Error GoTo 0
              '
              !DESTIRECHA = DESTRECHA%
              !MOTIRECHA = MOTIRECHA%
              !OBSERECHA = OBSERVARECHAZ$
              !UBICFIS = UB$
              NUMOITOC% = XVALO(!NroItemOC)
              .Update
              '
1127          DOPRI$ = "BR." + TRIM$(Str$(NROBOL&))
              DOPRILA$ = DOPRI$
              DOSELA$ = TRIM$(NUREPRO$)
              DOSEC$ = TRIM$(NUREPRO$)
              DEPX% = DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(CI%)
              If CANOM# >= 0.05 Then
                CANTI = CANOM#
                REFE$ = "Rto." + TRIM$(NUREPRO$) + " - " + rasocli$(NC1%)
                VUNI# = PRECIOUNITCOMPRA#(LOTE$) 'VER ACA DE GRABAR EL VALOR DE COMPRA VER TEMA DE MONEDA.
                MONEII% = MONELOTE%(LOTE$)
                PPX% = InStr(REFX$, "- BC D.")
                If PPX% > 0 Then
                    REFE$ = Mid$(REFX$, PPX% + 5)
                End If
                While Len(DOSEC$) > 10: DOSEC$ = Mid$(DOSEC$, 2): Wend
                If Len(DOSEC$) < 8 Then DOSEC$ = "Rt." + DOSEC$
                APROBOLE% = 1   '*****
                If InStr(1, EMPREAC$, "ARCON") > 0 Then
                    ' GRABA EN TABLA DE APROBACION SOLO LAS CANTIDADES APROBADAS
                    RECHFAL = Abs(RECHA# + FALTA#)
                    CANTIAPRO = CANTI - RECHFAL
                    Call MOVISTOCK_APRBR(FF%, CI%, LOTE$, "BR", DOPRILA$, DOSELA$, REFE$, VUNI#, "PESOS", NC1%, DEPX%, CANTIAPRO)
                    ' GENERA UN MOVIMIENTO DE ENTRADA PARA COMPENSAR EL RECHAZO Y EL FALTANTE
                    If RECHFAL > 0.005 Then
                        Call MOVISTOK(FF%, CI%, LOTE$, "BR", DOPRILA$, DOPRI$, DOSELA$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC1%, DEPX%, RECHFAL)
                    End If
                Else
                    GUARDO_DEPX% = DEPX%
                    'valido si es
                    If ESTRAX% = 2 And ESIMPO% = 1 Then DEPX% = DEPOSITO_EQ_IMP%: If DEPX% <= 0 Then DEPX% = GUARDO_DEPX%
                     METROX# = 0 'clave seguimiento: RecepcionCambaBobina
                     KILOX# = 0  ' EN LA RECPECION AGREGA AL MOVISTO EL PESO Y  LOS METROS
                    If SOYCAMBACERES% > 0 Then
                       CONDI$ = "IDENLOTE='" & LOTE$ & "'"
                       METROX# = XVALO(VALOBADA("BOLRECEP", "M_Largo", CONDI$, "NOMESS"))
                       KILOX# = XVALO(VALOBADA("BOLRECEP", "PESO", CONDI$, "NOMESS"))
                    End If
                    Call MOVISTOK(FF%, CI%, LOTE$, "BR", DOPRILA$, DOPRI$, DOSELA$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC1%, DEPX%, CANTI, , , , , , , , , , , METROX#, KILOX#)
                    DEPX% = GUARDO_DEPX%
                End If
                APROBOLE% = 0
              End If
              '
              If RECHA# >= 0.05 Then
                HAYRECHA% = 1
                If PORECHLIN% = 0 Then
                     CANTI = RECHA# * (-1)
                     REFE$ = "RECHAZO - Rto." + TRIM$(NUREPRO$)
                     APROBOLE% = 1 '******
                     Call MOVISTOK(FF%, CI%, LOTE$, CODRECHA$, REMIRECH$, REMIRECH$, DOSELA$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC1%, DEPX%, CANTI)
                     '
                     If DESTRECHA% > 0 Then ' DEPOSITO RELACIONADO CON DESTINO
                       CANTI = -1 * CANTI
                       REG$ = FILTERGETRECORD("TADESTI", 1, Chr$(DESTRECHA%))
                       DTINO$ = TRIM$(Mid$(REG$, 33, 20))
                       DPSITO% = Asc(Mid$(REG$, 32, 1))
                       REFE$ = "Para : " & DTINO$
                       If DPSITO% > 0 Then
                          Call MOVISTOK(FF%, CI%, LOTE$, CODRECHA$, REMIRECH$, REMIRECH$, DOSELA$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC1%, DPSITO%, CANTI)
                       End If
                     End If
                     APROBOLE% = 0
                  ElseIf PORECHLIN% = 1 Then
                     DOPRI12$ = "BR." + TRIM$(Str$(NROBOL&))
                     CIRECHA% = CI%
                     FERECHA% = FF%
                     CARECHA = RECHA#
                End If
              End If
              '
              If Abs(FALTA#) >= 0.05 Then
                CANTI = FALTA# * (-1)
                REFE$ = "FALTANTE - Rto." + TRIM$(NUREPRO$)
                If CANTI > 0 Then
                  REFE$ = "SOBRANTE - Rto." + TRIM$(NUREPRO$)
                End If
                APROBOLE% = 1
                Call MOVISTOK(FF%, CI%, LOTE$, "BR", DOPRILA$, DOPRI$, DOSELA$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC1%, DEPX%, CANTI)
                APROBOLE% = 0
              End If
              '
              ' aqui re-activa la Orden de Compra por el rechazado mas el faltante
              RECHFAL = RECHA# + FALTA#
              If RECHFAL > 0.005 Then
                NOCORI& = CVS(Mid$(X$, 151, 4))
                If NOCORI& > 0 Then
                   If NANTEN& <= 0 Then 'SI ES O/C cerrada'\\\OT-06-0149-WAR-27/04/06///
                     Call ABRECOMPRAS
                     SQLX$ = "Select Coef_Unids, Cant_Rec,  Stat_Ocom  from OCOMDETA "
                     SQLX$ = SQLX$ + " Where Nume_Ocom= " & NOCORI&
                     SQLX$ = SQLX$ + " And Cod_Inte=" & CI%
                     SQLX$ = SQLX$ + " AND Stat_Ocom < 8" ' SE AGREGO PARA QUE NO APLIQUE SOBRE UNA ANULADA O CUMPLIDA PROBLEMA EN  17/02/16 FERVI
                     If NUMOITOC% = 0 Then
                        SQLX$ = SQLX$ + " And Cant_Rec >= " & TRIM$(FORMATNUM$(Abs(CANOM# - 0.005), "F15.3"))
                     Else
                        SQLX$ = SQLX$ + " AND NUORD_ITEM=" + CStr(NUMOITOC%)
                     End If
                     
                     Dim OCOMDETA As ADODB.Recordset
                     Set OCOMDETA = New ADODB.Recordset
                     OCOMDETA.CursorType = adOpenKeyset
                     OCOMDETA.LockType = adLockPessimistic
                     OCOMDETA.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
                     With OCOMDETA
1135                   If Not (.EOF And .BOF) Then
                            CFXUNI# = !Coef_Unids
                            If CFXUNI# < 0.005 Then CFXUNI# = 1
                            RECHAX# = RECHFAL / CFXUNI#
                            If !Cant_Rec >= RECHAX# - 0.0005 Then
                                !Cant_Rec = !Cant_Rec - RECHAX#
                                !Stat_Ocom = 0
                                .Update
                                GoTo 1145
                            End If
                         Else
                            Call MENSERR(24, "No fue Posible Actualizar la Cantidad Pendiente\en la Orden de Compra Número " + TRIM$(Str$(NOCORI&)))
                        End If
1145                 End With
                     OCOMDETA.Close
                     Set OCOMDETA = Nothing
                   End If
                End If
                '\\\OT-06-0149-WAR-27/04/06///
                
                If NANTEN& > 0 Then
                  NOCORI& = CVS(Mid$(X$, 151, 4))
                  If NOCORI& > 0 Then
                   'RE ACTIVA O/C ABIERTA
                   Call ABRECOMPRAS
1136               SQLX$ = "Select * from DETOCABI"
                   SQLX$ = SQLX$ + " Where Nume_Ocom= " & NOCORI&
                   SQLX$ = SQLX$ + " And Cod_Inte=" & CI%
                   SQLX$ = SQLX$ + " And Cant_Rec >= " & TRIM$(FORMATNUM$(Abs(CANOM# - 0.005), "F15.3"))
                   
                   Dim DetoCabi As ADODB.Recordset
                   Set DetoCabi = New ADODB.Recordset
                   DetoCabi.CursorType = adOpenKeyset
                   DetoCabi.LockType = adLockPessimistic
                   DetoCabi.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
                   '
                   With DetoCabi
                    If Not (.EOF And .BOF) Then
                        CFXUNI# = !Coef_Unids
                        If CFXUNI# < 0.005 Then CFXUNI# = 1
                        RECHAX# = RECHFAL / CFXUNI#
                        If !Cant_Rec >= RECHAX# - 0.0005 Then
                            !Cant_Rec = !Cant_Rec - RECHAX#
                            !Stat_Ocom = 0
                            .Update
                            GoTo 1137
                        End If
                      Else
                        Call MENSERR(24, "No fue Posible Actualizar la Cantidad Pendiente\en la Orden de Compra Abierta Número " + TRIM$(Str$(NOCORI&)))
                    End If '
1137              End With
                  'RE ACTIVA AUTORIZACIÓN DE ENTREGA
                   SQLX$ = "Select * from DETOCABI"
                   SQLX$ = SQLX$ + " Where Nume_Ocom= " & NOCORI&
                   SQLX$ = SQLX$ + " And Cod_Inte=" & CI%
                   SQLX$ = SQLX$ + " And Nume_Aten=" & NANTEN&
                   SQLX$ = SQLX$ + " And Cant_Rec >= " & TRIM$(FORMATNUM$(Abs(CANOM# - 0.005), "F15.3"))
                   
                   Dim AutoEnt As ADODB.Recordset
                   Set AutoEnt = New ADODB.Recordset
                   AutoEnt.CursorType = adOpenDynamic
                   AutoEnt.LockType = adLockPessimistic
                   AutoEnt.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
                   
                   With AutoEnt
1155                 If Not (.EOF And .BOF) Then
                       CFXUNI# = 1 ' !Coef_Unids
                       If CFXUNI# < 0.005 Then CFXUNI# = 1
                       RECHAX# = RECHFAL / CFXUNI#
                       If !Acuen_Aten >= RECHAX# - 0.0005 Then
                           !Acuen_Aten = !Acuen_Aten - RECHAX#
                           !Sald_Aten = !Cant_Aten - !Acuen_Aten
                           !Stat_Aten = 0
                           .Update
                           GoTo 1165
                       End If
                     Else
                        Call MENSERR(24, "No fue Posible Actualizar la Cantidad Pendiente\en la Autorización de Entrega Número " + TRIM$(Str$(NANTEN&)))
                     End If
                     
1165               End With
                  End If
                End If
              End If
            End If
          End With
          COBRECEPTMP.Close
          Set COBRECEPTMP = Nothing
          '
          ' AL TRABAJAR ARCON EN MODALIDAD DE APROBACION, LA CREACION DEL LOTE SE HACE AL APROBAR
          If InStr(1, EMPREAC$, "ARCON") > 0 Then GoTo 1128
          '
          If ESTRAZABLE%(CI%) = 1 Then
            NPROV$ = rasocli$(NC1%)
            Call CREALOTIN(HOY(HOS$), HOY(HOS$), HOY(HOS$), CI%, LOTE$, "", LOTPROV$, CANOM#, 0, 0, DEPX%, "", NPROV$)
            Call ACSALOTE(CI%, LOTE$, "NOMESS")
          End If
          '
1128    Next U&
        '
        If ESIMPO% > 0 Then
           Call GENERA_OF_NSERIE_PROD_IMPORTADOS(NROBOL&)
        End If
        
199     NOCLOSE% = 0: NOARRAS% = 0
        '
        Label12.Visible = False
        '
        Call CIERRARCH("*.*")
        Call BAJALDISCO
        '
        '
        LU$ = LUDAT$
        If EXISTE%(LU$ + "BRAP-ARC.FRM") > 0 Then
           HOII% = HOY(HOS$)
           FECHDOC$ = HOS$
           NUMEDOC$ = ""
           TIPODOC$ = "Vale Aprobación Recepción"
           DOCUORIG$ = "BR." + TRIM$(Str$(NROBOL&))
           '
           'CODPARAM$(1) ="TRANSPOR": DOCPARAM$(1) = TRIM$(FOREMIFAC.Text8.TEXT)
           CAPARDOC% = 0
           '
           CODTABU1$(1) = "COD-MAT"
           CODTABU1$(2) = "DES-MAT"
           CODTABU1$(3) = "CANTIDAD"
           CACOLTAB1% = 3
           '
           DESTIDOC% = 0
           CAITAB1% = 0
           '
           For YX& = 1 To ULTREG&("!BOLREDET")
              YY$ = REGLEIDO$("!BOLREDET", YX&)
              CI% = CVI(Mid$(YY$, 83, 2))
              CAN# = CVD(Mid$(YY$, 163, 8))
              If CI% > 0 Then
                If CI% <= NUMAS% Then
                  If CAN# > 0.005 Then
                    '
                    'UNID$ = UNIMED$(CI%)
                    'If UNID$ = "" Then UNID$ = UNIMED$(CI%)
                    CANS$ = CSTRING$(MKD$(CAN#), 3, 12, 1, 2)
                    '
                    CAITAB1% = CAITAB1% + 1
                    CODDX$ = CODEXT$(CI%)
                    TETABU1$(1, CAITAB1%) = CODDX$
                    TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
                    TETABU1$(3, CAITAB1%) = CANS$
                  End If
                End If
              End If
           Next YX&
           '
           Call PRINTDOC("@BRAP-ARC")   ' Remito de Facturacion
           '
        End If
        '
        
        Call FILTERDELETE("IBOPEAN", 1, MKS$(NROBOL&))
        '
        If HAYRECHA% = 1 Then
          DEST% = XVALO(Control("BOLRECEP", "DETIRECH"))
          ' CONFIGURACION DE DESTINO EN SETUP SI ESTA CONFIGURADA 1 DEVOLUCION IMPRIME EL REMITO
          If (MODORECHA% = 0 And PORECHLIN% = 0) Or DEST% = 1 Then 'SALTEO IMPRESIÓN DE REMITO DE RECHAZO
            Call REMIRECHA(NC1%, DEPX%, REMIRECH$)
          End If
        End If
        '
'        ' GABAL - IMPRIME FORMULARIO CON LAS UBICACIONES
'        If EXISTE%(LU$ + "UBIMARE.SQR") > 0 Then
'            FILTX$ = TRIM$(Str$(NROBOL&))
'            Call STDFORM("UBIMARE", FILTX$)
'        End If
        '
200     Call CIERRARCH("*.*")
        Call BAJALDISCO
        Screen.MousePointer = 1
        '
        If CREALOTE% = 1 Then
           If EXISTE%("TRAZAB07.EXE") > 0 Or EXISTE%("TRAZAB08.EXE") > 0 Then
              ACONEC$ = "TRAZAB07"
              If EXISTE%("TRAZAB07.EXE") < 1 Then ' verifica con que modulo de trazabilidad debe conectarse
                 ACONEC$ = "TRAZAB08"
                 If EXISTE%("TRAZAB08.EXE") < 1 Then
                    ACONEC$ = "TRANSA07"
                 End If
              End If
              Call CONECRUN(ACONEC$ + "/BORECEP" + TRIM$(NBR1$), "Registro de Trazabilidad de B/R's")
              GoTo 1120
           End If
        End If
        '
        If EXISTE%("CALPRO07.EXE") > 0 Then
          ICALREC07.Label5 = ""
          ICALREC07.Label5 = TRIM$(Str$(NROBOL&))
          ICALREC07.Show 1
          'Call GENIBOLPE
        End If
        '
        If HAYRECHA% > 0 Then
           If PORECHLIN% = 1 Then
              Call PROCERECHA(CIRECHA%, FERECHA%, CARECHA, DOPRI12$) '\\\OT-06-0191-WAR-09/06/06///
              Call RECONECTA
              Call FINAL("")
              Exit Sub
           End If
           HAYRECHA% = 0
        End If
        GoTo 1120
        '
'PROCERECHA:
'     'Muestro el Objeto
'     VDEF$ = Space$(256)
'     Call REPLA(VDEF$, MKI$(CI%), 1, 2)
'     Call REPLA(VDEF$, MKI$(FF%), 3, 2)
'     Call REPLA(VDEF$, MKI$(NC%), 5, 2)
'     Call REPLA(VDEF$, Chr$(0), 7, 1)
'     Call REPLA(VDEF$, Chr$(0), 8, 1)
'1130 OBJ$ = OBJPLA$("DETARECHA2", VDEF$)
'     'Valido el OBJ$
'     OBJANT$ = OBJ$
'     If TRIM(OBJ$) <> "" Then
'       DEST% = Asc(Mid$(OBJANT$, 7, 1))
'       MOTI% = Asc(Mid$(OBJANT$, 8, 1))
'       REFX$ = TRIM$(Mid$(OBJANT$, 9, 48))
'       '
'       If TRIM$(ECOARCH$("TADESTI", Chr(DEST%))) = "" Then
'         Call MENSERR(24, "Destino de Rechazo no Válido.")
'         VDEF$ = OBJANT$
'         GoTo 1130
'       End If
'       If TRIM$(ECOARCH$("TAMOTIV", Chr(MOTI%))) = "" Then
'         Call MENSERR(24, "Motivo de Rechazo no Válido.")
'         VDEF$ = OBJANT$
'         GoTo 1130
'       End If
'       If TRIM$(REFX$) = "" Then
'         Call MENSERR(24, "Referencia de Rechazo no Válida.")
'         VDEF$ = OBJANT$
'         GoTo 1130
'       End If
'       '
'       GoTo 1131
'     End If
'     Call MENSERR(24, "Es Mandatorio Procesar el Rechazo")
'     GoTo 1130
'     '
'1131 TDEST$ = TRIM$(ECOARCH$("TADESTI", Chr(DEST%)))
'     TMOTI$ = TRIM$(ECOARCH$("TAMOTIV", Chr(MOTI%)))
'Return
'
End Sub


Function VALIDAR_DESPA_IMPO%(NROLOTE&, Optional MODO$)
    '
    'DEVUELVE 0 (SI NO ESTA CONFIGURADO O SI EL PROVEEDOR NO ES DEL EXTERIOR
    'DEVUELVE -1 (SI FALTA ALGUN DATO)
    'DEVUELVE 1 (SI ES PROVEEDOR DEL EXTERIOR, ESTA CONFIGURADO Y TIENE LOS DATOS DE IMPORTACION
    
    VALIDAR_DESPA_IMPO% = 0
    PROVELOTEX% = PROVELOTENUM(NROLOTE&)
    If PIVACLI%(PROVELOTEX% * (-1)) <> 5 Then Exit Function
    
    If XVALO(Control("BOLRECEP", "RECEIMPO")) > 0 Then
       Dim obj As New Art_Importados
       Call obj.READ_DATOS_IMPO(NROLOTE&)
       If IsDate(obj.FECHA) = False Or obj.Despacho = "" Or obj.Aduana = "" Then
          If MODO$ <> "NOMESS" Then
             Call COMUNI("Atención Proveedor Comercio Exterior\Faltan Datos de Importación\Deberá Informar los Mismos Para Poder Aprobar la Recepción Seleccionada")
          End If
          VALIDAR_DESPA_IMPO% = -1
       Else ' SI TIENE TODOS LOS DATOS
          VALIDAR_DESPA_IMPO% = 1
     End If
       Set obj = Nothing
    End If
End Function

Sub MOVISTOCK_APRBR(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, CANTI)
   '
   On Error GoTo ControlDeErrores
   
   Static CTX%, REBLAN$, RELIBR$, UREMOVI&, DECONSUM$
   Call COPYSTRU("MOVISTO", "BRECEPEN")
   '
   Screen.MousePointer = 11
   UREMAS& = ULTREG&("MASTER")
   '
   User% = USNBR% Mod 100
   If User% < 1 Then User% = 1
   Call CARDEPOS
   ' ********** ARMAR REGISTRO DE MOVIMIENTO **********
9  DPRX$ = DOPRI$
10 PPXX% = InStr(DPRX$, "-0")
   If PPXX% > 0 Then
      DPRX$ = Left$(DPRX$, PPXX%) + Mid$(DPRX$, PPXX% + 2)
      GoTo 10
   End If
   DOPRIX$ = AJUSTI$(DPRX$, 10)
   '
   DPRX$ = DOSEC$
20 PPXX% = InStr(DPRX$, "-0")
   If PPXX% > 0 Then
      DPRX$ = Left$(DPRX$, PPXX%) + Mid$(DPRX$, PPXX% + 2)
      GoTo 20
   End If
   DOSECX$ = AJUSTI$(DPRX$, 10)
   '
   NCLIE% = NC%: NPROV% = 0
   If NC% < 0 Then
      NCLIE% = 0: NPROV% = Abs(NC%)
   End If
   '
   Entrada# = CDbl(CANTI): SALIDA# = 0
   If CANTI < 0 Then
      Entrada# = 0: SALIDA# = CDbl(Abs(CANTI))
   End If
   CANACU = SALDETOT + CANTI
   '
   DEPOMOVI% = DEPO%
   RECORD$ = REBLAN$
   Call REPLA(RECORD$, MKI$(FECHA%), 1, 2)
   Call REPLA(RECORD$, MKI$(CI%), 3, 2)
   Call REPLA(RECORD$, LOTE$, 5, 16)
   Call REPLA(RECORD$, CMOV$, 21, 2)
   If CMOV$ = "*R" Then
        Call REPLA(RECORD$, "RP", 21, 2)      ' reparación
      ElseIf CMOV$ = "*D" Then
        Call REPLA(RECORD$, "DR", 21, 2)      ' DESPACHO reparación
   End If
   Call REPLA(RECORD$, DOPRIX$, 23, 10)
   Call REPLA(RECORD$, DOSECX$, 33, 10)
   Call REPLA(RECORD$, REFE$, 43, 30)
   Call REPLA(RECORD$, MKD$(VUNI#), 73, 8)
   Call REPLA(RECORD$, MONE$, 81, 2)
   Call REPLA(RECORD$, Chr$(DEPOMOVI%), 83, 1)
   Call REPLA(RECORD$, MKI$(NCLIE%), 84, 2)
   Call REPLA(RECORD$, MKI$(NPROV%), 86, 2)
   Call REPLA(RECORD$, MKD$(Entrada#), 88, 8)
   Call REPLA(RECORD$, MKD$(SALIDA#), 96, 8)
   Call REPLA(RECORD$, String$(8, 0), 104, 8)
   Call REPLA(RECORD$, String$(6, 0), 112, 6)
   Call REPLA(RECORD$, MKS$(CANACU), 118, 4)
   '
   HO% = HOY(HOS$)
   HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
   '
   Call REPLA(RECORD$, MKI$(HO%), 122, 2)
   Call REPLA(RECORD$, MKI$(Int(HP)), 124, 2)
   Call REPLA(RECORD$, Chr$(User%), 126, 1)
   Call REPLA(RECORD$, String$(2, 0), 127, 2)
   '
   Dim BUGRA As String * 140
   Call REGAPP("BRECEPEN", RECORD$)
   REMOVIST& = ULTREG&("BRECEPEN")
   If REMOVIST& > UREMOVI& Then UREMOVI& = REMOVIST&
   '
   NKX% = FreeFile
   Open LUCOM$ + "\BRECEPEN.LOC" For Random Access Read Write As #NKX% Len = 140
   RECONO& = 1 + LOF(NKX%) / 140
   BUGRA$ = "BRECEPEN " + MKS$(REMOVIST&) + RECORD$
   Put #NKX%, RECONO&, BUGRA$
   Close #NKX%
   '
   If CI% > 0 Then FLAGMOVI%(CI%) = 1
   '
   Screen.MousePointer = 1
   '
   Exit Sub
   '
ControlDeErrores:
   Screen.MousePointer = 1
   Call COMUNI("Error Al Registrar Como Pendiente a la Recepción '" + DOPRIX$)
End Sub

Sub Mot_Rechazo(NROBOLRECE&, CI%, IDLOTE$, CANTRECHA#, MOTIRECHA%, DESTRECHA%, RECH_OBSERVA$)
   '
   RECHARE07.Label3 = NROBOLRECE&
   RECHARE07.Label5 = CODEXT$(CI%)
   RECHARE07.Label8 = FORMATNUM$(CANTRECHA#, "F9." & CNTDC$ & "")
   RECHARE07.Label9 = IDLOTE$
   RECHARE07.Show 1
   On Error Resume Next
   MOTIRECHA% = XVALO(RECHARE07.Text4)
   DESTRECHA% = XVALO(RECHARE07.Text3)
   RECH_OBSERVA$ = Left$(RECHARE07.text9, 48)       'OBSERVACION
   On Error GoTo 0
   '
   RECHARE07.Label3 = ""
   RECHARE07.Label5 = ""
   RECHARE07.Label8 = ""
   RECHARE07.Label9 = ""
   RECHARE07.Label11 = ""
   RECHARE07.Text4 = ""
   RECHARE07.text9 = ""
   RECHARE07.Text3 = ""

End Sub

'\\\OT-05-0608-WAR-06/09/05///
Sub DISCREBOL()
        '
        Dim REGBO&(2048)
        '
1100    Call GENIBOLPE
        '
1120    Call FRESHECHO("IBOLPEN")
        Screen.MousePointer = 1
        'Call CLEARVEN
        If ULTREG&("IBOLPEN") < 1 Then
          Exit Sub
        End If
        '
        X$ = "": Y$ = "": Z$ = "": YY$ = ""
        EBOL$ = "": PEDY$ = "": DESCRI$ = ""
        NBR$ = OBJPLA$("SELBOLPEN", "")
        If NBR$ = "" Then
            On Error Resume Next
            Movisto.Close
            Bolrecep.Close
            SalStock.Close
            On Error GoTo 0
            Exit Sub
        End If
        '
        NROBOL& = CVS(Left$(NBR$, 4))
        If NROBOL& < 1 Then
            Exit Sub
        End If
        '
        'Call BLIMESS("Seleccionando ...")
        FIN& = ULTREG&("BOLREPEN")
        Call COPYSTRU("BOLREDET", "BOLRESEL")
        'Call BLANARCH("BOLRESEL")
        
        For KU& = 1 To NUMAS%
          CIMOVI%(KU&) = 0
        Next KU&
        '
        CABOPE% = 0: JJ& = 0
        For i& = 1 To FIN&
          X$ = REGLEIDO$("BOLREPEN", i&)
          If CVS(Left$(X$, 4)) = NROBOL& Then
            CANOM# = CVD(Mid$(X$, 89, 8))
            APRO# = CVD(Mid$(X$, 163, 8))
            RECHA# = CVD(Mid$(X$, 171, 8))
            FALTA# = CVD(Mid$(X$, 179, 8))
            '
            If CANOM# < 0.05 Or APRO# > 0.05 Or RECHA# > 0.05 Or Abs(FALTA#) > 0.05 Then GoTo 1121
            Call REPLA(X$, MKD$(CANOM#), 163, 8)
            '
            IDLO$ = TRIM$(Mid$(X$, 69, 12))
825         If Len(IDLO$) > 0 Then
              If InStr(IDLO$, "-") > 0 Then
                IDLO$ = Mid$(IDLO$, 2)
                GoTo 825
              End If
            End If
            Call REPLA(X$, MKS$(Val(IDLO$)), 1, 4)
            '
            If CABOPE% = 0 Then
              NC% = CVI(Mid$(X$, 105, 2)): NC1% = (-1) * NC%
              FF% = CVI(Mid$(X$, 81, 2))
              FEX = FF%: FECHAREC$ = FECHATEX$(FEX)
              REMI$ = Mid$(X$, 119, 16)
              DP1% = Asc(Mid$(X$, 97, 1))
              If ECOARCH$("ECODEP", Chr$(DP1%)) = "" Then
                  DP1% = 0
              End If
              '
              Y$ = String$(128, 0)
              Call REPLA(Y$, MKS$(NROBOL&), 1, 4)
              Call REPLA(Y$, MKI$(FF%), 5, 2)
              Call REPLA(Y$, MKI$(NC%), 7, 2)
              Call REPLA(Y$, rasocli$(NC1%), 7, 32)
              Call REPLA(Y$, DOMICLI$(NC1%), 39, 32)
                  CLOCA$ = TRIM$(CPOSCLI$(NC1%)): If CLOCA$ <> "" Then CLOCA$ = CLOCA$ + " "
              Call REPLA(Y$, CLOCA$ + LOCACLI$(NC1%), 71, 32)
              '\\\OT-05-0662-WAR-06/10/05///
              Call REPLA(Y$, CONTACTCLI$(NC1%), 103, 48)
              '\\\OT-05-0662-WAR-FIN///
              Call REPLA(Y$, REMI$, 153, 16)
              Call REPLA(Y$, MKI$(0), 187, 2)
              AINDI% = AINDI% + 1
            End If
            '
            CABOPE% = CABOPE% + 1: REGBO&(CABOPE%) = i&
            JJ& = JJ& + 1: Call GRAREG("BOLRESEL", X$, JJ&)
            '
          End If
          '
1121    Next i&
        Call SETULTREG("BOLRESEL", JJ&)
        '
        'Call BLIMESS("-")
        If CABOPE% < 1 Then
            Call COMUNI("No hay Recepciones Pendientes\para la Boleta Elegida")
            GoTo 1120
        End If
        '
        Call FILESORT("BOLRESEL", "", 1, 1, "ASC")
        Call TRALOCAL("BOLRESEL", "BOLREDET")
        '
11215   Z$ = OBJPLA$("DECLADISCRE", Y$)
        If Z$ = "" Then
          GoTo 1120
        End If
        '
        'VALIDACION DE DATOS
        PEDI1$ = Mid$(Z$, 169, 18)
        LOTE1$ = Mid$(Z$, 189, 18)
        TRACL1$ = Mid$(Z$, 207, 48)
        PROCL1% = CVI(Mid$(Z$, 187, 2))
        CONTA1$ = Mid$(Z$, 103, 48)
        '
        If TRIM$(PEDI1$) = "" Then
          Call MENSERR(24, "Falta Indicar Número de Pedido")
          Y$ = Z$
          GoTo 11215
        End If
        If TRIM$(LOTE1$) = "" Then
          Call MENSERR(24, "Falta Indicar Número de Lote")
          Y$ = Z$
          GoTo 11215
        End If
        If TRIM$(TRACL1$) = "" Then
          Call MENSERR(24, "Falta Indicar Trazabilidad del Cliente")
          Y$ = Z$
          GoTo 11215
        End If
        If PROCL1% <= 0 Then
          Call MENSERR(24, "Falta Indicar Producto")
          Y$ = Z$
          GoTo 11215
        End If
        '
        'ACA GRABO DISCREPA
        Call BLANARCH("!DISCREPA")
        For U& = 1 To ULTREG&("!BOLREDET")
          XXX$ = REGLEIDO$("!BOLREDET", U&)
          CODI1% = CVI(Mid$(XXX$, 83, 2))
          CANOM1# = CVD(Mid$(XXX$, 89, 8))
          zz$ = REGBLAN("!DISCREPA")
          Call REPLA(zz$, MKI$(CODI1%), 1, 2)
          Call REPLA(zz$, MKS$(CANOM1#), 27, 4)
          Call REPLA(zz$, MKS$(CANOM1#), 31, 4)
          Call GRAREG("!DISCREPA", zz$, U&)
        Next U&
        Call SETULTREG("!DISCREPA", U&)
        Call CIERRARCH("!DISCREPA")
        '
1122    VTB% = VENTABU%("DISCREPA/NC/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(rasocli$(NC1%)) + " - " + FECHAREC$)
        If VTB% < 0 Then GoTo 1120
        'IMPRESION DEL COMPROBANTE
        LU$ = LUDAT$
        If EXISTE%(LU$ + "DREC-ASB.FRM") > 0 Then
           HOII% = HOY(HOS$)
           FECHDOC$ = HOS$
           NUMEDOC$ = ""
           TIPODOC$ = "Vale Aprobación Recepción"
           DOCUORIG$ = "BR." + TRIM$(Str$(NROBOL&))
           '
           CODPARAM$(1) = "FECH-EMI": DOCPARAM$(1) = FECHDOC$
           CODPARAM$(2) = "RASO-CLI": DOCPARAM$(2) = rasocli$(NC1%)
           CODPARAM$(3) = "NUME-COM": DOCPARAM$(3) = PEDI1$
           CODPARAM$(4) = "CTAC-CLI": DOCPARAM$(4) = CONTA1$
           CODPARAM$(5) = "NRO-REMI": DOCPARAM$(5) = REMI$
           CODPARAM$(6) = "NRO-LOTE": DOCPARAM$(6) = LOTE1$
           CODPARAM$(7) = "REF-MAT1": DOCPARAM$(7) = TRACL1$
           CODPARAM$(8) = "COD-CONJ": DOCPARAM$(8) = CODEXT$(PROCL1%)
           CODPARAM$(9) = "DES-CONJ": DOCPARAM$(9) = DESCRIT0$(PROCL1%)
           CAPARDOC% = 9
           '
           CODTABU1$(1) = "ORD-ITEM"
           CODTABU1$(2) = "COD-MAT"
           CODTABU1$(3) = "DES-MAT"
           CODTABU1$(4) = "CANTIDAD"
           CODTABU1$(5) = "CAN-MPRI"
           CODTABU1$(6) = "OBS-OPER"
           CACOLTAB1% = 6
           '
           DESTIDOC% = 0
           CAITAB1% = 0
           '
           For YX& = 1 To ULTREG&("!DISCREPA")
              YY$ = REGLEIDO$("!DISCREPA", YX&)
              CI% = CVI(Mid$(YY$, 1, 2))
              CAN# = CVS(Mid$(YY$, 27, 4))
              CARE1# = CVS(Mid$(YY$, 31, 4))
              CARE2$ = CSTRING$(MKS$(CARE1#), 3, 12, 1, 2)
              OBSER1$ = Mid$(YY$, 35, 48)
              If CI% > 0 Then
                If CI% <= NUMAS% Then
                  If CAN# > 0.005 Then
                    '
                    'UNID$ = UNIMED$(CI%)
                    'If UNID$ = "" Then UNID$ = UNIMED$(CI%)
                    CANS$ = CSTRING$(MKS$(CAN#), 3, 12, 1, 2)
                    '
                    CAITAB1% = CAITAB1% + 1
                    CODDX$ = CODEXT$(CI%)
                    TETABU1$(1, CAITAB1%) = FORMATNUM$(CAITAB1%, "F2.0")
                    TETABU1$(2, CAITAB1%) = CODEXT$(CI%)
                    TETABU1$(3, CAITAB1%) = DESCRIT0$(CI%)
                    TETABU1$(4, CAITAB1%) = CANS$
                    TETABU1$(5, CAITAB1%) = CARE2$
                    TETABU1$(6, CAITAB1%) = OBSER1$
                  End If
                End If
              End If
           Next YX&
           '
           
           Call PRINTDOC("@DREC-ASB")
           '
        End If
          '
End Sub
'\\\OT-05-0608-WAR-FIN///

Sub ANUBORE()
        '
        If DERACCE%("ANUBORE", "Anulación Boletas de Recepción") < 2 Then
          Exit Sub
        End If
        '
        Dim REGBO&(2048)
        '
        ' SE QUITA LA ACTUALIZACION DEL IBOLPEN PORQUE YA SE HACE EN LA GRABACION DE LA
        ' RECEPCION Y EN LA APROBACION
        Call GENINDIBOL(2)
1100    If ULTREG&("IBOPEAN") < 1 Then
          Exit Sub
        End If
'1100    Call GENIBOLPE(1)
'        If ULTREG&("IBOLPEN") < 1 Then
'          Exit Sub
'        End If
        '
1120    'Call FRESHECHO("IBOPEAN")
        Screen.MousePointer = 1
        Call SELECHO("IBOPEAN")
        NROBOL& = Val(VALACT1$("IBOPEAN"))
        If NROBOL& < 1 Then Exit Sub
        '
        If COMALTER%("Confirme la Anulación de la B.R. " + Str$(NROBOL&) + "\\Cancelar Transacción\Confirmar Anulación") <> 2 Then
          Exit Sub
        End If
        '
        Call GENERABOLREPEN(NROBOL&)
        '
'        SQLX$ = " SELECT * FROM BOLRECEP "
'        SQLX$ = SQLX$ + "WHERE Status = 0 "
'        SQLX$ = SQLX$ + "AND NumeBoRe = " & NROBOL& & " "
'        '
'        Dim COBRECEPTMP As ADODB.Recordset
'        Set COBRECEPTMP = New ADODB.Recordset
'        COBRECEPTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

        ACTU% = 0
        FIN& = ULTREG&("BOLREPEN")
        For i& = 1 To FIN&
          X$ = REGLEIDO$("BOLREPEN", i&)
          If CVS(Left$(X$, 4)) = NROBOL& Then
            CANOM# = CVD(Mid$(X$, 89, 8))
            APRO# = CVD(Mid$(X$, 163, 8))
            RECHA# = CVD(Mid$(X$, 171, 8))
            FALTA# = CVD(Mid$(X$, 179, 8))
            '
            'If CANOM# < 0.05 Or APRO# > 0.05 Or RECHA# > 0.05 Or Abs(FALTA#) > 0.05 Then GoTo 1121
            '
            CI% = CVI(Mid$(X$, 83, 2))
            LOTE$ = TRIM$(Mid$(X$, 69, 12))
            NOCORI& = CVS(Mid$(X$, 151, 4))
            NAENT& = CVS(Mid$(X$, 189, 4))
            '
            If NAENT& > 0 Then GoTo 1146 ' ORDEN DE COMPRA ABIERTA CON A/ENTREGA
            '
            If NOCORI& > 0 Then
              Call ABRECOMPRAS
              ABUS$ = "SELECT Coef_Unids, Cant_Rec, Stat_Ocom FROM OCOMDETA"
              ABUS$ = ABUS$ + " WHERE Nume_Ocom = " & NOCORI& & " "
              ABUS$ = ABUS$ + "AND Cod_Inte = " & CI% & " "
              Dim OCCCOMDETA As ADODB.Recordset
              Set OCCCOMDETA = New ADODB.Recordset
              OCCCOMDETA.Open FILTSQL$(ABUS$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
              
              With OCCCOMDETA
                On Error Resume Next
                .MoveFirst
                If Err Then
                   On Error GoTo 0
                   GoTo 1145
                End If
                On Error GoTo 0
                Do While Not .EOF
                  CFXUNI# = XVALO(!Coef_Unids)
                  If CFXUNI# < 0.005 Then CFXUNI# = 1
                  RECHAX# = CANOM# / CFXUNI#
                  If XVALO(!Cant_Rec) >= RECHAX# - 0.0005 Then
                      !Cant_Rec = XVALO(!Cant_Rec) - RECHAX#
                      !Stat_Ocom = 0
                      .Update
                      GoTo 1145
                  End If
                .MoveNext
                Loop
                Call MENSERR(24, "No fue Posible Actualizar la Cantidad Pendiente\en la Orden de Compra Número " + TRIM$(Str$(NOCORI&)))
                '
1145          End With
              OCCCOMDETA.Close
              Set OCCCOMDETA = Nothing
            End If
            GoTo 11471
            '
            '
1146        ' CASO DE O/COMPRA ABIERTA CON A/ENTREGA
            Call ABRECOMPRAS
            ABUS$ = "SELECT * FROM AUTOENT"
            ABUS$ = ABUS$ + " WHERE Nume_Aten = " & NAENT& & " "
            ABUS$ = ABUS$ + "AND Cod_Inte = " & CI% & " "
            Dim AUTTTOENT As ADODB.Recordset
            Set AUTTTOENT = New ADODB.Recordset
            AUTTTOENT.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
            '
            With AUTTTOENT
                On Error Resume Next
                .MoveFirst
                If Err Then
                   On Error GoTo 0
                   GoTo 11469
                End If
                On Error GoTo 0
                Do While Not .EOF
                  RECHAX# = CANOM#
                  If XVALO(!Acuen_Aten) >= RECHAX# - 0.0005 Then
                    !Acuen_Aten = XVALO(!Acuen_Aten) - RECHAX#
                    !Stat_Aten = 0
                    .Update
                    '
                    NOCORI& = XVALO(!nume_ocom)
                    ABUS2$ = "SELECT * FROM DetoCabi "
                    ABUS2$ = ABUS2$ + " WHERE Nume_Ocom = " & NOCORI& & " "
                    ABUS2$ = ABUS2$ + " AND Cod_Inte = " & CI% & " "
                    Dim DETTTOCABI As ADODB.Recordset
                    Set DETTTOCABI = New ADODB.Recordset
                    DETTTOCABI.Open FILTSQL$(ABUS2$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
                    With DETTTOCABI
                       On Error Resume Next
                       .MoveFirst
                       If Err Then
                          On Error GoTo 0
                          GoTo 11463
                       End If
                       On Error GoTo 0
                       Do While Not .EOF
                         If XVALO(!Cant_Rec) >= RECHAX# - 0.0005 Then
                             !Cant_Rec = XVALO(!Cant_Rec) - RECHAX#
                             !Stat_Ocom = 0
                             .Update
                             GoTo 11463
                         End If
                       .MoveNext
                       Loop
11463               End With
                    DETTTOCABI.Close
                    Set DETTTOCABI = Nothing
                    GoTo 11469
                  End If
                .MoveNext
                Loop
                '
                Call MENSERR(24, "No fue Posible Actualizar la Cantidad Pendiente\en la Orden de Compra Número " + TRIM$(Str$(NOCORI&)))
                '
11469       End With
            AUTTTOENT.Close
            Set AUTTTOENT = Nothing
            GoTo 11471
            '
11471       SQLX$ = "DELETE FROM BOLRECEP "
            SQLX$ = SQLX$ + "WHERE NumeBoRe = " & NROBOL & " "
            SQLX$ = SQLX$ + "AND IdenLote = '" & LOTE$ & "' "
            SQLX$ = SQLX$ + "AND Cod_Inte = " & CI%
            FLEXCONN.Execute SQLX$
            '
            Call GRAREG("BOLREPEN", String$(512, 0), i&)
            ACTU% = 1
            '
            If InStr(1, EMPREAC$, "GABAL") > 0 Then
                ' BLANQUEA LA TABLA DE ETIQUETAS QUITANDO EL LOTE DE RECEPCION A LAS QUE CORRESPONDE
                Call ACTUREGISTRO("ORFAETI", "LOTERECE", "LOTERECE='" & LOTE$ & "'", "", REGAF&)
                ' BLANQUEA LA TABLA DE UBICACIONES
                Call BORRAREGISTROS("UBICALOT", "LOTE='" & LOTE$ & "'", REGAF&)
            End If
            '
          End If
1121    Next i&


'        FIN& = ULTREG&("BOLREPEN")
'        Call COPYSTRU("BOLREDET", "BOLRESEL")
'        '
'        ACTU% = 0
'        For I& = 1 To FIN&
'          x$ = REGLEIDO$("BOLREPEN", I&)
'          If CVS(Left$(x$, 4)) = NROBOL& Then
'            CANOM# = CVD(Mid$(x$, 89, 8))
'            APRO# = CVD(Mid$(x$, 163, 8))
'            RECHA# = CVD(Mid$(x$, 171, 8))
'            FALTA# = CVD(Mid$(x$, 179, 8))
'            '
'            'If CANOM# < 0.05 Or APRO# > 0.05 Or RECHA# > 0.05 Or Abs(FALTA#) > 0.05 Then GoTo 1121
'            '
'            CI% = CVI(Mid$(x$, 83, 2))
'            LOTE$ = TRIM$(Mid$(x$, 69, 12))
'            NOCORI& = CVS(Mid$(x$, 151, 4))
'            NAENT& = CVS(Mid$(x$, 189, 4))
'            '
'            If NAENT& > 0 Then GoTo 1146 ' ORDEN DE COMPRA ABIERTA CON A/ENTREGA
'            '
'            If NOCORI& > 0 Then
'              Call ABRECOMPRAS
'              ABUS$ = "SELECT * FROM OCOMDETA"
'              ABUS$ = ABUS$ + " WHERE Nume_Ocom = " & NOCORI& & " "
'              ABUS$ = ABUS$ + "AND Cod_Inte = " & CI% & " "
'              Dim OCCCOMDETA As ADODB.Recordset
'              Set OCCCOMDETA = New ADODB.Recordset
'              OCCCOMDETA.Open FILTSQL$(ABUS$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
'
'              With OCCCOMDETA
'                On Error Resume Next
'                .MoveFirst
'                If Err Then
'                   On Error GoTo 0
'                   GoTo 1145
'                End If
'                On Error GoTo 0
'                Do While Not .EOF
'                  CFXUNI# = XVALO(!Coef_Unids)
'                  If CFXUNI# < 0.005 Then CFXUNI# = 1
'                  RECHAX# = CANOM# / CFXUNI#
'                  If XVALO(!Cant_Rec) >= RECHAX# - 0.0005 Then
'                      !Cant_Rec = XVALO(!Cant_Rec) - RECHAX#
'                      !Stat_Ocom = 0
'                      .Update
'                      GoTo 1145
'                  End If
'                .MoveNext
'                Loop
'                Call MENSERR(24, "No fue Posible Actualizar la Cantidad Pendiente\en la Orden de Compra Número " + TRIM$(Str$(NOCORI&)))
'                '
'1145          End With
'              OCCCOMDETA.Close
'              Set OCCCOMDETA = Nothing
'            End If
'            GoTo 11471
'            '
'            '
'1146        ' CASO DE O/COMPRA ABIERTA CON A/ENTREGA
'            Call ABRECOMPRAS
'            ABUS$ = "SELECT * FROM AUTOENT"
'            ABUS$ = ABUS$ + " WHERE Nume_Aten = " & NAENT& & " "
'            ABUS$ = ABUS$ + "AND Cod_Inte = " & CI% & " "
'            Dim AUTTTOENT As ADODB.Recordset
'            Set AUTTTOENT = New ADODB.Recordset
'            AUTTTOENT.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
'            '
'            With AUTTTOENT
'                On Error Resume Next
'                .MoveFirst
'                If Err Then
'                   On Error GoTo 0
'                   GoTo 11469
'                End If
'                On Error GoTo 0
'                Do While Not .EOF
'                  RECHAX# = CANOM#
'                  If XVALO(!Acuen_Aten) >= RECHAX# - 0.0005 Then
'                    !Acuen_Aten = XVALO(!Acuen_Aten) - RECHAX#
'                    !Stat_Aten = 0
'                    .Update
'                    '
'                    NOCORI& = XVALO(!Nume_Ocom)
'                    ABUS2$ = "SELECT * FROM DetoCabi "
'                    ABUS2$ = ABUS2$ + " WHERE Nume_Ocom = " & NOCORI& & " "
'                    ABUS2$ = ABUS2$ + " AND Cod_Inte = " & CI% & " "
'                    Dim DETTTOCABI As ADODB.Recordset
'                    Set DETTTOCABI = New ADODB.Recordset
'                    DETTTOCABI.Open FILTSQL$(ABUS2$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
'                    With DETTTOCABI
'                       On Error Resume Next
'                       .MoveFirst
'                       If Err Then
'                          On Error GoTo 0
'                          GoTo 11463
'                       End If
'                       On Error GoTo 0
'                       Do While Not .EOF
'                         If XVALO(!Cant_Rec) >= RECHAX# - 0.0005 Then
'                             !Cant_Rec = XVALO(!Cant_Rec) - RECHAX#
'                             !Stat_Ocom = 0
'                             .Update
'                             GoTo 11463
'                         End If
'                       .MoveNext
'                       Loop
'11463               End With
'                    DETTTOCABI.Close
'                    Set DETTTOCABI = Nothing
'                    GoTo 11469
'                  End If
'                .MoveNext
'                Loop
'                '
'                Call MENSERR(24, "No fue Posible Actualizar la Cantidad Pendiente\en la Orden de Compra Número " + TRIM$(Str$(NOCORI&)))
'                '
'11469       End With
'            AUTTTOENT.Close
'            Set AUTTTOENT = Nothing
'            GoTo 11471
'            '
'11471       SQLX$ = "DELETE FROM BOLRECEP "
'            SQLX$ = SQLX$ + "WHERE NumeBoRe = " & NROBOL & " "
'            SQLX$ = SQLX$ + "AND IdenLote = '" & LOTE$ & "' "
'            SQLX$ = SQLX$ + "AND Cod_Inte = " & CI%
'            FLEXCONN.Execute SQLX$
'            '
'            Call GRAREG("BOLREPEN", String$(512, 0), I&)
'            ACTU% = 1
'            '
'            If InStr(1, EMPREAC$, "GABAL") > 0 Then
'                ' BLANQUEA LA TABLA DE ETIQUETAS QUITANDO EL LOTE DE RECEPCION A LAS QUE CORRESPONDE
'                Call ACTUREGISTRO("ORFAETI", "LOTERECE", "LOTERECE='" & LOTE$ & "'", "", REGAF&)
'                ' BLANQUEA LA TABLA DE UBICACIONES
'                Call BORRAREGISTROS("UBICALOT", "LOTE='" & LOTE$ & "'", REGAF&)
'            End If
'            '
'          End If
'1121    Next I&
        '
        CONDI$ = "NUMEBORE = " & NROBOL& & " AND NUMEBORE > 0 "
        Call ACTUREGISTRO("RENCONSIG", "Cant_Conj", CONDI$, 0, RAFE&)
        Call ACTUREGISTRO("RENCONSIG", "Cant_Rendid", CONDI$, 0, RAFE&)
        Call ACTUREGISTRO("RENCONSIG", "MarProce", CONDI$, 0, RAFE&)
        '
        Call AJUMACON
        Call BORRAREGISTROS("RENCONSIG", CONDI$, RAFE&)
        '
        Call GENINDIBOL(0)
        '
        If ACTU% > 0 Then
           Call COMUNI("Boleta de Recepción Anulada")
        End If
        GoTo 1100
        '
End Sub

Sub GENINDIBOL(MODO%)
        '\\\OT-06-0543-RG-12/12/06///
        Call COPYSTRU("IBOLPEN", "IBOPEAN")
        
        If MODO% = 0 Then
            Call APERBASDAT("STOCKS")
            SQLY$ = " SELECT NUMEBORE, FechRecep, NumRem_Prov, NUMEORDSER, Raso_Prov, NumProv, Status FROM BOLRECEP WITH(NOLOCK) "
            SQLY$ = SQLY$ + " WHERE STATUS <= 1 "
            SQLY$ = SQLY$ + " ORDER BY NumeBoRe "
            '
            Screen.MousePointer = 11
            Dim BRECEPTMP As ADODB.Recordset
            Set BRECEPTMP = READSET(SQLY$)
            '
            ANTNUM = 0
            i& = 0
            J& = 0
            k& = 0
            With BRECEPTMP
              If Not (.EOF And .BOF) Then
                .MoveFirst
                Do While (BRECEPTMP.EOF = False)
                  k& = k& + 1
                  Call TRACE(20, k&, FIN&)
                  NROBOL& = !NUMEBORE
                  If NROBOL& <> ANTNUM Then
                     X$ = REGBLAN$("INDIBOL")
                     FEX = CVINT(!FechRecep)
                     REMI$ = SAFETEXT$(!NumRem_Prov)
                     ST& = XVALO(!Status)
                     NC1% = (-1) * XVALO(!NumProv)
                     NUORSE& = XVALO(!NUMEORDSER)
                     '
                     Call REPLA(X$, MKS$(NROBOL&), 1, 4)
                     Call REPLA(X$, FECHATEX$(FEX), 5, 8)
                     Call REPLA(X$, REMI$, 15, 16)
                     Call REPLA(X$, SAFETEXT$(!Raso_Prov), 32, 29)
                     i& = i& + 1
                     Call GRAREG("INDIBOL", X$, i&)
                     'If (ST& < 1) Or (ST& = 1 And NUORSE& > 0) Then
                     If (ST& < 1) Then
                        J& = J& + 1
                        Call GRAREG("IBOPEAN", X$, J&)
                     End If
                     ANTNUM = NROBOL&
                  End If
                .MoveNext
                Loop
                Call SETULTREG("INDIBOL", i&)
                Call SETULTREG("IBOPEAN", J&)
                Call FRESHECHO("IBOPEAN")
                Call FRESHECHO("IBOLPEN")
              End If
            End With
            BRECEPTMP.Close
            Set BRECEPTMP = Nothing
            '
            Screen.MousePointer = 1
            '\\\OT-06-0543-RG-FIN///
        ElseIf MODO% = 1 Then
            Call APERBASDAT("STOCKS")
            SQLY$ = " SELECT NUMEBORE, FechRecep, NumRem_Prov, NUMEORDSER, Raso_Prov, NumProv, Status FROM BOLRECEP WITH(NOLOCK) "
            SQLY$ = SQLY$ + " WHERE STATUS <= 1 "
            SQLY$ = SQLY$ + " ORDER BY NumeBoRe "
            '
            Screen.MousePointer = 11
            Dim BRECEPTMP1 As ADODB.Recordset
            Set BRECEPTMP1 = READSET(SQLY$)
            '
            ANTNUM = 0
            i& = 0
            J& = 0
            k& = 0
            With BRECEPTMP1
              If Not (.EOF And .BOF) Then
                .MoveFirst
                Do While (.EOF = False)
                  k& = k& + 1
                  Call TRACE(20, k&, FIN&)
                  NROBOL& = !NUMEBORE
                  If NROBOL& <> ANTNUM Then
                     X$ = REGBLAN$("INDIBOL")
                     FEX = CVINT(!FechRecep)
                     REMI$ = SAFETEXT$(!NumRem_Prov)
                     ST& = XVALO(!Status)
                     NC1% = (-1) * XVALO(!NumProv)
                     NUORSE& = XVALO(!NUMEORDSER)
                     '
                     Call REPLA(X$, MKS$(NROBOL&), 1, 4)
                     Call REPLA(X$, FECHATEX$(FEX), 5, 8)
                     Call REPLA(X$, REMI$, 15, 16)
                     Call REPLA(X$, SAFETEXT$(!Raso_Prov), 32, 29)
                     i& = i& + 1
                     Call GRAREG("INDIBOL", X$, i&)
                     If (ST& < 1) Or NUORSE& > 0 Then
                        J& = J& + 1
                        Call GRAREG("IBOPEAN", X$, J&)
                     End If
                     ANTNUM = NROBOL&
                  End If
                .MoveNext
                Loop
                Call SETULTREG("INDIBOL", i&)
                Call SETULTREG("IBOPEAN", J&)
                Call FRESHECHO("IBOPEAN")
                Call FRESHECHO("IBOLPEN")
              End If
            End With
            BRECEPTMP1.Close
            Set BRECEPTMP1 = Nothing
            '
            Screen.MousePointer = 1
            '\\\OT-06-0543-RG-FIN///
        Else
            Call APERBASDAT("STOCKS")
            SQLY$ = " SELECT NUMEBORE, FechRecep, NumRem_Prov, NUMEORDSER, Raso_Prov, NumProv, Status FROM BOLRECEP WITH(NOLOCK) "
            SQLY$ = SQLY$ + " WHERE STATUS < 1 "
            SQLY$ = SQLY$ + " ORDER BY NumeBoRe "
            '
            Screen.MousePointer = 11
            Dim BRECEPTMP2 As ADODB.Recordset
            Set BRECEPTMP2 = READSET(SQLY$)
            '
            ANTNUM = 0
            i& = 0
            J& = 0
            k& = 0
            With BRECEPTMP2
              If Not (.EOF And .BOF) Then
                .MoveFirst
                Do While (.EOF = False)
                  k& = k& + 1
                  Call TRACE(20, k&, FIN&)
                  NROBOL& = !NUMEBORE
                  If NROBOL& <> ANTNUM Then
                     X$ = REGBLAN$("INDIBOL")
                     FEX = CVINT(!FechRecep)
                     REMI$ = SAFETEXT$(!NumRem_Prov)
                     ST& = XVALO(!Status)
                     NC1% = (-1) * XVALO(!NumProv)
                     NUORSE& = XVALO(!NUMEORDSER)
                     '
                     Call REPLA(X$, MKS$(NROBOL&), 1, 4)
                     Call REPLA(X$, FECHATEX$(FEX), 5, 8)
                     Call REPLA(X$, REMI$, 15, 16)
                     Call REPLA(X$, SAFETEXT$(!Raso_Prov), 32, 29)
                     J& = J& + 1
                     Call GRAREG("IBOPEAN", X$, J&)
                     ANTNUM = NROBOL&
                  End If
                .MoveNext
                Loop
                Call SETULTREG("INDIBOL", i&)
                Call SETULTREG("IBOPEAN", J&)
                Call FRESHECHO("IBOPEAN")
                Call FRESHECHO("IBOLPEN")
              End If
            End With
            BRECEPTMP2.Close
            Set BRECEPTMP2 = Nothing
            '
            Screen.MousePointer = 1
            '\\\OT-06-0543-RG-FIN///
        End If
End Sub

Sub RECPROP()
    '
    If PRECOCIE% = 0 Then
       PRECOCIE% = (-1)
       If Control$("", "PRECOCIE") = "SI" Then
          PRECOCIE% = 1
       End If
    End If
    '
10  HOOI% = HOY(HO$)
    If ULTREG&("ORDEFABR") < 1 Then   ' por numero de O/F
        GoTo 100
    End If
    '
    MODO% = COMALTER%("Puede Declarar Entradas a Stock:\  \a) por Cierre Parcial o Total de Orden de Fabricación\o\b) en Forma Directa por Armado de Conjunto\\Cierre de O./Fabricación\Armado de Conjunto\Cancelar")
    On MODO% GoTo 20, 100
    Exit Sub
    '
20  Call ORFCIE
    Exit Sub
    '
100 Screen.MousePointer = 11
    For U& = ULTREG&("RECPROP") To 1 Step -1
        XX$ = REGLEIDO$("RECPROP", U&)
        If CVS(Mid$(XX$, 59, 4)) > 0 Then
            UPARTE& = CVS(Mid$(XX$, 59, 4))
            GoTo 105
        End If
    Next U&
    UPARTE& = 0
    '
105 Screen.MousePointer = 1
    Call BLANARCH("!CAREPRO")
    '
    VDEF$ = String$(88, 0)
    Call REPLA(VDEF$, MKI$(HOOI%), 1, 2)
    Call REPLA(VDEF$, MKS$(UPARTE& + 1), 85, 4)
    Call REPLA(VDEF$, Chr$(0), 37, 1)
    Call REPLA(VDEF$, Chr$(0), 38, 1)
    '
110 XX$ = OBJPLA$("RECPROP", VDEF$)
    If XX$ = "" Then Exit Sub
    '
160 VDEF$ = XX$
    FECHA% = CVI(Left$(XX$, 2))
    If FECHA% > HOOI% Or FECHA% < 1 Then
        Call MENSERR(24, "Fecha Incorrecta")
        GoTo 110
    End If
    '
'    CI% = CVI(Mid$(XX$, 3, 2))
'    If TRIM$(CODEXT$(CI%)) = "" Then
'        Call MENSERR(24, "Codigo Inexistente.")
'        GoTo 110
'    End If
'    CODCON$ = CODEXT$(CI%)
'    '
    DP1% = Asc(Mid$(XX$, 37, 1))
    DP2% = Asc(Mid$(XX$, 38, 1))
    If TRIM$(ECOARCH$("ECODEP", Chr$(DP1%))) = "" Then
        Call MENSERR(24, "Falta Depósito de Entrada")
        GoTo 110
    End If
    '
    If TRIM$(ECOARCH$("ECODEP", Chr$(DP2%))) = "" Then
        Call MENSERR(24, "Falta Depósito de Consumo")
        GoTo 110
    End If
    '
120 U3& = 0
    For U1& = 1 To ULTREG&("!CAREPRO")
      XY$ = REGLEIDO$("!CAREPRO", U1&)
      CANTI = CVD(Mid$(XY$, 23, 8))
      If Abs(CANTI) >= 0.05 Then
        U3& = U3& + 1
        Call GRAREG("!CAREPRO", XY$, U3&)
      End If
    Next U1&
    Call SETULTREG("!CAREPRO", U3&)
    '
    VTB% = VENTABU%("CAREPRO")
    If VTB% < 0 Then GoTo 110
    '
    CARE& = ULTREG&("!CAREPRO")
    If CARE& < 1 Then GoTo 120
    '
    NUORIT% = 0
    For U& = 1 To CARE&
       XY$ = REGLEIDO$("!CAREPRO", U&)
       CANTI = CVD(Mid$(XY$, 23, 8))
       Scrap = 0
       CI% = CVI(Left$(XY$, 2))
       CODICONJ$ = CODEXT$(CI%)
       CICONJ% = CI%
       '
       '\\\OT-06-0165-WAR-08/05/06///
       'If CANTI < 0 Or SCRAP < 0 Or CANTI + SCRAP < 0.05 Then
       If Scrap < 0 Then
       '\\\OT-06-0165-WAR-FIN///
          Call MENSERR(24, "Cantidades Incorrectas en Código " + TRIM$(CODICONJ$))
          GoTo 120
       End If
       '
       NORFABRI$ = ""
       CACONJU# = CDbl(CANTI + Scrap)
       VACONTROL% = 1
       '
       Call SETULTREG("!MATEROF", 0)
       Call CIERRARCH("!MATEROF")
       Call CARCOMPO(CODICONJ$, NORFABRI$, CACONJU#, VACONTROL%, 2)
       '
       If VACONTROL% < 0 Then GoTo 120           ' cancelacion
       '
       '\\\OT-05-0730-WAR-10/11/05///
       CAITMO% = 0
         For YX& = 1 To ULTREG&("!MATEROF")
            YY$ = REGLEIDO$("!MATEROF", YX&)
            CI% = CVI(Mid$(YY$, 1, 2))
            CAN# = CVS(Mid$(YY$, 5, 4))
             If CI% > 0 Then
               If CI% <= NUMAS% Then
                 If CAN# > 0.005 Then
                   CAITMO% = CAITMO% + 1
                   CIMOVIM%(CAITMO%) = CI%
                   CAMOVIM(CAITMO) = (-1) * CAN#
                 End If
               End If
             End If
         Next YX&
       If PUEDEMOVI%(DP2%) < 0 Then GoTo 110 ' CANCELA SALIDA
       '\\\OT-05-0730-WAR-FIN///
       '
       NRO& = CVS(Mid$(XX$, 85, 4))
       NROOF& = CVS(Mid$(XY$, 19, 4))
       If NROOF& < 1 Then NROOF& = NRO&
       DOPRI$ = "OF-" + TRIM$(Str$(NROOF&))
       DOSEC$ = "PF-" + TRIM$(Str$(NRO&))
       LOTE$ = Mid$(XY$, 3, 16)
       If LOTE$ = "" Then
          LOTE$ = DOPRI$
       End If
       REFE$ = Mid$(XX$, 55, 30)
       If TRIM$(REFE$) = "" Then
          REFE$ = "Informe de Produccion"
       End If
       If CANTI < (-0.005) Then
          REFE$ = "Desarmado de Conjunto"
       End If
       '
       NC% = 0
       VUNI# = 0
       MONE$ = ""
       MONEII% = 0
       CMOV$ = "PF"
       '
       NUORIT% = 0
       DEPX% = DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(CICONJ%)
       Call MOVISTOK(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CANTI)
       '
       YY$ = REGBLAN$("RECPROP")
       Call REPLA(YY$, MKI$(FECHA%), 1, 2)
       Call REPLA(YY$, Chr$(DP1%), 3, 1)
       Call REPLA(YY$, Chr$(DP2%), 4, 1)
       Call REPLA(YY$, REFE$, 5, 32)
       Call REPLA(YY$, MKI$(CI%), 37, 2)
       Call REPLA(YY$, MKS$(CANTI), 39, 4)
       Call REPLA(YY$, LOTE$, 43, 16)
       Call REPLA(YY$, MKS$(NRO&), 59, 4)
       Call REPLA(YY$, MKI$(-1), 63, 2)
       Call REGAPP("RECPROP", YY$)
    
       COD$ = CODEXT$(CICONJ%)
       CODCON$ = COD$
       '
       For U2& = 1 To ULTREG&("!MATEROF")
          zz$ = REGLEIDO$("!MATEROF", U2&)
          CI2% = CVI(Left$(zz$, 2))
          CAN = -1 * CVS(Mid$(zz$, 5, 4))
          '
          If TRIM$(UCase$(Unimed$(CI2%))) = "GK" Then
             CAN = CAN / 1000
          End If
          '
          If CI2% > 0 Then
             If CI2% <= NUMAS% Then
                If Abs(CAN) >= 0.05 Then
                   LOTE$ = ""
                   CMOV$ = "CF"
                   DOPRI$ = "OF-" + TRIM$(Str$(NROOF&))
                   DOSEC$ = "PF-" + TRIM$(Str$(NRO&))
                   REFE$ = "CONSUMO S/DESP." + CODCON$
                   '\\\OT-06-0165-WAR-08/05/06///
                   If CAN > 0 Then
                    REFE$ = "DESARMADO CONJ." + CODCON$
                   End If
                   '\\\OT-06-0165-WAR-FIN///
                   DEPX% = DP2%: If DP2% = 99 Then DEPX% = DEPOPRE%(CI2%)
                   '
                   If STODEPOS(CI2%, DEPX%) < (-1) * CAN Then
                     If InStr(EPAC$, "YALTRE") > 0 Then
                       Call CARDEPOS
                       For KKII% = 1 To NUSU%
                         If STODEPOS(CI2%, COSU%(KKII%)) >= (-1) * CAN Then
                           DEPX% = COSU%(KKII%)
                           GoTo 124
                         End If
                       Next KKII%
                     End If
                   End If
                   '
124                NUORIT% = NUORIT% + 1
                   Call MOVISTOK(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CAN)
                End If
             End If
          End If
       Next U2&
       Call CIERRARCH("!MATEROF")
       '
       XY$ = REGLEIDO$("!CAREPRO", U&)
       Call REPLA(XY$, MKD$(0), 23, 8)
       Call GRAREG("!CAREPRO", XY$, U&)
       '
129 Next U&
    Call CIERRARCH("MOVISTO")
    Call CIERRARCH("RECPROP")
    '
180 Call CIERRARCH("*.*")
    Call BLANARCH("!CAREPRO")
    Screen.MousePointer = 1
    GoTo 100
    '
End Sub

'
Sub CONAPLI(APLINVO$, TERMINA%)
   TERMINA% = 1
   Select Case APLINVO$
       '
       'TRANSACCIONES
       '
       Case "BOLREPRO"
           Call Command7_Click                 ' BOLETA DE RECEPCION
       Case "INSPAPRO"
           Call Command2_Click                 ' INSPECCION Y APROBACION DE BOLETA DE RECEPCION
       Case "ANUBOLREC"
           Call Command8_Click                 ' ANULACION DE BOLETA DE RECEPCION
       Case "ENTPROPROD"
           Call Command6_Click                 ' ENTRADA POR PROPIA PRODUCCION
       '
       'CONSULTAS Y LISTADOS
       '
       Case "CONBOLREC"
           Call Command3_Click                 ' BOLETAS DE RECEPCION
       Case "CONPEDORCOM"
           Call Command4_Click                 ' PEDIDOS Y ORDENES DE COMPRAS
       Case "CONMATCON"
           Call Command5_Click                 ' MATERIAL EN CONSIGNACION
       Case "RECRANFEC"
           Call Command11_Click                ' RECEPCIONES POR RANGO DE FECHAS
       Case "RECONRECH"
           Call Command16_Click                ' RECEPCIONES CON RECHAZOS
       Case "RECSINORD"
           Call Command15_Click                ' RECEPCIONES SIN ORDEN DE COMPRA
       Case "RECHRANFEC"
           Call Command18_Click                ' RECHAZOS POR RANGO DE FECHAS
       Case "CONORDCOMP"
           Call Command13_Click                ' CONSULTA DE ORDENES DE COMPRAS
       Case "VERBOLRECE"
           Call Command12_Click                ' VER Y REIMPRIMIR BOLETA DE RECEPCION
       '
       'CONFIGURACION BOLETA DE RECEPCION
       '
       Case "CONGENBRE"
           Call Command14_Click                ' CONFIGURACION GENERAL DE BOLETA DE RECEPCION
           
       Case "TRANFSEC"
            Call mnutransferirasector_Click
       Case Else
           TERMINA% = 0
   End Select
End Sub

Sub ALERTATRANSFE()
    '
    'SE COMENTO POR QUE NO LO ESTAN USANDO Y TENIA SU DEMORA
    Exit Sub
    
    TXTX$ = ""
    TXTXMSG$ = ""
    '
    DEPPAÑOL% = XVALO(Control("", "DEPOCEN"))
    '
    For YX& = 1 To ULTREG&("!BOLREDET")
        YY$ = REGLEIDO$("!BOLREDET", YX&)
        CIIXI% = CVI(Mid$(YY$, 83, 2))
        CANTUI& = CVD(Mid$(YY$, 89, 8))
        If CANTUI& >= 1 Then
            SQLX$ = "SELECT RESERVA.Cod_Inte, RESERVA.CantRes, ORDEFABR.IdSubOr, ORDEFABR.StatuOrf "
            SQLX$ = SQLX$ + " FROM RESERVA "
            SQLX$ = SQLX$ + " INNER JOIN ORDEFABR "
            SQLX$ = SQLX$ + " ON RESERVA.IdSubOr = ORDEFABR.IdSubOr "
            SQLX$ = SQLX$ + " WHERE ORDEFABR.StatuOrf = 1 AND RESERVA.COD_INTE=" & CIIXI%
            
            'OBTENGO TOTAL RESERVADO Y ARMO LISTA DE OF QUE TIENEN RESERVADO AL ARTICULO
            TXTX1$ = "Codigo: '" & CODEXT$(CIIXI%) & "' Para: "
            QRES# = 0: JJ& = 0
            Set RESETMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
            With RESETMP
                Do While .EOF = False
                    QRES# = QRES# + !CANTRES
                    NROOF$ = SAFETEXT$(!idsubor)
                    SEPARADOR$ = "; "
                    JJ& = JJ& + 1
                    If JJ& = 4 Then SEPARADOR$ = "\": JJ& = 0
                    TXTX1$ = TXTX1$ + NROOF$ + SEPARADOR$
                .MoveNext
                Loop
            End With
            'STOCK DEL ARTICULO
            'QSTOCK# = STOCKACT(CIIXI%)
            QSTOCK# = STODEPOS(CIIXI%, DEPPAÑOL%)
            
            'VERIFICO SI EL STOCK CUBRE LA RESERVA
            If QRES# > QSTOCK# And QRES# > 0.05 Then
                TXTXMSG$ = TXTXMSG$ + TXTX1$ & "\"
            End If
        End If
     Next YX&
     '
     If TRIM$(TXTXMSG$) <> "" Then
        TXTX$ = "Los Siguientes Materiales Recibidos\Son Requeridos para O-F's en Curso:\ \"
        TXTX$ = TXTX$ + TXTXMSG$
        Call COMUNI(TXTX$)
        '
        'AGREGA MENSAJE A LAS OBSERVACIONES
        ALERTA$ = REPLACAR$(TXTX$, "\", " ")
        Call FRATEXTO(ALERTA$, 72)
        JJX& = ULTREG&("!OBSERVOF")
        For KKU1% = 1 To CARETEX%
            XX$ = RETEX$(KKU1%)
            JJX& = JJX& + 1
            Call GRAREG("!OBSERVOF", XX$, JJX&)
        Next KKU1%
        '
        'ENVIO POR MAIL
        MAIPA$ = Control("BOLRECEP", "MAIPAÑOL")
        If MAIPA$ <> "" And InStr(MAIPA$, "@") > 0 Then
            TXTX$ = REPLACAR(TXTX$, "\", vbCrLf)
            MAI_DESTI$ = "Pañol"
            MAI_DIREL$ = MAIPA$
            MAI_TEXTO$ = TXTX$
            MAI_ASUNT$ = "Disponibilidad a Transferir"
            Call GENERAMAIL
        End If
        
     End If
     '
     
End Sub

Sub BOLERECEP_VIE()
Dim IMPCONOC%
'
'falta generar ficha de recepcion fervi-air
     '
     If ULTREG&("TATIREC") < 1 Then
         X$ = REGBLAN$("TATIREC")
         Call REPLA(X$, Chr$(1), 1, 1)
         Call REPLA(X$, "Material General", 2, 31)
         Call REPLA(X$, "BOLREGEN", 49, 16)
         Call REPLA(X$, "BOLREGEN", 65, 16)
         Call REGAPP("TATIREC", X$)
         Call CIERRARCH("TATIREC")
     End If
     '
1100 Call CIERRARCH("*.*")
     Call SETULTREG("!BOLREDET", 0)
     JJ& = 0
     '
     NROBOL& = 1
     Call ABRESTOCKS
     'Set BolreTemp = Stocks.OpenRecordset("SELECT DISTINCT NumeBoRe from BOLRECEP ORDER BY NumeBoRe DESC")
     'On Error Resume Next
     'BolreTemp.MoveFirst
     'If Err < 1 Then
     '  NROBOL& = 1 + BolreTemp!NumeBoRe
     'End If
     'On Error GoTo 0
     
         
     SQLX$ = " SELECT max(NumeBoRe) as maximo from BOLRECEP"
     Dim BolreTemp As ADODB.Recordset
     Set BolreTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
     If Not IsNull(BolreTemp!Maximo) Then
         NROBOL& = 1 + (BolreTemp!Maximo)
     End If
     BolreTemp.Close
     Set BolreTemp = Nothing

     NORDE% = 0
     '
     If TIMATEI% = 0 Then
         X$ = REGLEIDO$("TATIREC", 1)
         TIMATEI% = Asc(Left$(X$, 1))
     End If
     '
     If TRIM$(Mid$(EBOLAN$, 4, 12)) <> "" Then
       NC% = 0
     End If
     '
     VDEF$ = MKI$(NC%) + Chr$(TIMATEI%) + CODICLI$((-1) * NC%)
     EBOL$ = OBJPLA$("SELPROVRECACC", VDEF$)
     If EBOL$ = "" Then
        Screen.MousePointer = 1 '\\\OT-06-0179-WAR-12/05/06///
        Exit Sub
     End If
     EBOLAN$ = EBOL$
     '
     NC3% = CVI(Left$(EBOL$, 2))
     If NC3% > 0 Then
       NC% = NC3%
       GoTo 15
     End If
     '
     NC% = NUMEPRO%(TRIM$(Mid$(EBOL$, 4, 12)))
15   If NC% < 1 Or STATCLI%((-1) * NC%) < 1 Then
       Call MENSERR(24, "Proveedor Inexistente")
       GoTo 1100
     End If
     '
     TIMATEI% = Asc(Mid$(EBOL$, 3, 1))
     If TIMATEI% < 1 Or ECOARCH("TATIREC", Chr$(TIMATEI%)) = "" Then
         Call MENSERR(24, "Falta Tipo de Material")
         GoTo 1100
     End If
     '
     REGT& = TIMATEI%
     X$ = REGLEIDO$("TATIREC", REGT&)
     FORMPLA$ = TRIM$(UCase$(Mid$(X$, 33, 16)))
     FORMTAB$ = TRIM$(UCase$(Mid$(X$, 49, 16)))
     FORIPRE$ = TRIM$(UCase$(Mid$(X$, 65, 16)))
     '
     NC1% = (-1) * NC%
     '
1105 Y$ = PEDY$
     If Y$ = "" Then
       Y$ = Space$(128)
       Call REPLA(Y$, Chr$(0), 121, 1)
     End If
     '
     Call REPLA(Y$, MKS$(NROBOL&), 1, 4)
     Call REPLA(Y$, EBOL$, 7, 2)
     Call REPLA(Y$, MKI$(HOY(HO$)), 5, 2)
     Call REPLA(Y$, rasocli$(NC1%), 9, 32)
     Call REPLA(Y$, DOMICLI$(NC1%), 41, 32)
     CLOCA$ = TRIM$(CPOSCLI$(NC1%)): If CLOCA$ <> "" Then CLOCA$ = CLOCA$ + " "
     Call REPLA(Y$, CLOCA$ + LOCACLI$(NC1%), 73, 32)
     Call REPLA(Y$, Chr$(0), 121, 1)
    '
     OBJE$ = "BOLREC-ENCB"
     If Control$("", "NUMEBREC") = "MANUAL" Then
         OBJE$ = "BOLREC-ENCB-M"
     End If
     '
     ' CONTROL DE VIGILANCIA
     If Control$("BOLRECEP", "NUCONVIG") = "SI" Then
         OBJE$ = "BOLREC-ENCC"
     End If
     '
     TICOMPRO$ = TRIM$(Control$("", "COMPBREC"))
     '
     PEDY$ = OBJPLA$(OBJE$, Y$)
     If PEDY$ = "" Then
        Screen.MousePointer = 1 '\\\OT-06-0179-WAR-12/05/06///
        Exit Sub
     End If
     '
     NUCONVI$ = TRIM$(Mid$(PEDY$, 145, 24))
     FF% = CVI(Mid$(PEDY$, 5, 2))
     FEX = FF%: FECHAREC$ = FECHATEX$(FEX)
     If TRIM$(FECHAREC$) = "" Then
         Call MENSERR(24, "Fecha Incorrecta")
         GoTo 1105
     End If
     '
     REMI$ = TRIM$(Mid$(PEDY$, 105, 16))
     If REMI$ = "" Then
        Call MENSERR(24, "Falta Número de Remito del Proveedor")
        GoTo 1105
     End If
     '
     DP1% = Asc(Mid$(PEDY$, 121, 1))
     DEPOSAL$ = TRIM$(ECOARCH$("ECODEP", Chr$(DP1%)))
     If DP1% < 1 Or DEPOSAL$ = "" Then
         Call MENSERR(24, "Falta Depósito de Entrada")
         GoTo 1105
     End If
     DEPOSAL$ = DEPOSAL$ + " (" + TRIM$(Str$(DP1%)) + ")"
     '
     NROPRO& = NROBOL&
     NROBOL& = CVS(Left$(PEDY$, 4))
     If NROBOL& < 1 Then
         Call MENSERR(24, "Número Incorrecto de B.R.")
         NROBOL& = NROPRO&
         Call REPLA(PEDY$, MKS$(NROBOL&), 1, 4)
         GoTo 1105
     End If
     'ACA NUEVO PARA GUIDI(COMPRUEBA QUE NO EXISTA EL REMI)
     NC2% = NC%
     NR1$ = REMI$
     If Control$("BOLRECEP", "SUPREMI") <> "SI" Then  '\\\OT-06-0515-RG-28/11/06///
        R1% = EXIREMI(NR1$, NC2%)
        If R1% = 1 Then
           Call COMUNI("El Número de Remito ya fue Procesado\en Recepción Anterior.")
           GoTo 15
        End If
     End If
    'HASTA ACA NUEVO PARA GUIDI
1112 Call CIERRARCH("*.*")
     '
     If InStr(EMPREAC$, "GABAL") > 0 Then ' SOLO GABAL
113     OPX% = COMALTER%("Opciones de Carga\\Carga Manual\Código de Barras")
        If OPX% < 1 Or OPX% > 2 Then GoTo 1105
        '
        If OPX% = 2 Then
           CARCOBAR07.Label2 = TRIM$(Str$(NC%))
           CARCOBAR07.Show 1
           '
           If XVALO(CARCOBAR07.List1.ListCount) < 1 Then GoTo 113
           CARCOBAR07.List1.Clear
           CARCOBAR07.List2.Clear
           JJ& = ULTREG&("!BOLREDET")
           HAYOC% = 1
           GoTo 1130
        End If
     End If
     
     '
     ' verificar que hay pendiente
     'jandy
     
     Screen.MousePointer = 11
     Call ABRECOMPRAS
     SQLX$ = "Select * from OCOMDETA"
     SQLX$ = SQLX$ + " Where Stat_Ocom < 8 "
     SQLX$ = SQLX$ + " AND  Stat_Ocom <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
     SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
     SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
     'Set COMPTEMP = Compras.OpenRecordset(sqlx$, dbOpenSnapshot)
     '
     Dim COMPTEMP As ADODB.Recordset
     Set COMPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     
     HAYCODI% = 0: HAYLIBRE% = 0
     With COMPTEMP
       On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         GoTo 45       ' no hay pendiente de recepcion --> va a codificado
       End If
       On Error GoTo 0
       Do While Not .EOF
         If XVALO(!cod_inte) > 0 Then
             HAYCODI% = 1
           Else
             HAYLIBRE% = 1
         End If
       .MoveNext
       Loop
     End With
     COMPTEMP.Close
     Set COMPTEMP = Nothing
     '
     'CONSIDERA CONSIGANCIONES PENDIENTES
     CONDI$ = "((Cant_Rendid < Cant_Entreg - 0.001) Or Cant_Entreg Is Null) And Npro_Ocom = " & NC%
     If CantRegistros("MACONSIG", CONDI$) > 0 Then HAYCODI% = 1
     '
     Screen.MousePointer = 1
     If HAYLIBRE% < 1 Then GoTo 45
     If HAYCODI% < 1 And HAYLIBRE% > 0 Then GoTo 42
     '
     JJ& = 0
     '
     'ACA NUEVO
     RESP4% = COMALTER("Elija el Material Recibido:\\Codificado\Texto Libre")
     If RESP4% < 1 Or RESP4% > 2 Then
         Screen.MousePointer = 1
         GoTo 1100
     End If
     '
     If RESP4% = 1 Then
         GoTo 45
       Else
'42       Call RECEOCLI(NC%, FF%, NROBOL&, REMI$, DP1%)
42       Call RECEOCLI(NC%, FF%, NROBOL&, REMI$, DP1%, NUCONVI$)
         Screen.MousePointer = 1
         'If CANCELX1% = 1 Then GoTo 45 '\\\OT-06-00178-WAR-12/05/06///
         GoTo 1100
     End If
     'HASTA ACA NUEVO
     '
45  DARX$ = TRIM$(Control$("BOLRECEP", "DIANREC"))
     If DARX$ = "" Then
          DIANRE& = 10000
        Else
          DIANRE& = Val(DARX$)
     End If
     HOIX1 = HOY(HOS$)
     '
     ANUCA% = 0
     If Control$("BOLRECEP", "ANUCANTI") = "SI" Then
       ANUCA% = 1
     End If
     '
     Screen.MousePointer = 11
     Call ABRECOMPRAS
     '
     CONDI$ = "Stat_Ocom < 8 "
     CONDI$ = CONDI$ + " AND  Stat_Ocom <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
     CONDI$ = CONDI$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
     CONDI$ = CONDI$ + " AND Npro_Ocom = " & NC% & " "
     CONDI$ = CONDI$ + " AND Cod_Inte > 0 " 'EXCLUYE LAS DE 0/C's TEXTO LIBRE
     '
'     CONDI1$ = CONDI$ + " AND DOC_ORIG LIKE 'OF%"
'     If CantRegistros("OCOMDETA", CONDI1$) > 0 Then
'        TTXX$ = "Hay Ordenes de Servicio Pendientes de Recepción"
'        IF COMALTER%(TTXX$+"\\Re
'     End If
'
     
     SQLX$ = "Select * from OCOMDETA"
     SQLX$ = SQLX$ + " Where " + CONDI$
     '
     Dim COMPTEMP_1 As ADODB.Recordset
     Set COMPTEMP_1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With COMPTEMP_1
       '
       On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Screen.MousePointer = 1
         HAYOC% = 0
         GoTo 1777
       End If
       On Error GoTo 0
       '
       JJ& = 0
        
       Do While (.EOF = False)
         CAPEN = XVALO(!cant_ocom) - XVALO(!Cant_Rec)
         COEFUNI = XVALO(!Coef_Unids)
         If COEFUNI < 0.0001 Then COEFUNI = 1
         If CAPEN > 0.005 Then
           NUOC& = XVALO(!nume_ocom)
           FENTRESO% = CVINT(!Fentre_Sol)
           FESO2 = FENTRESO%
           If DIENTRE%(HOIX1, FESO2) <= DIANRE& Then
             CI% = XVALO(!cod_inte)
             LOTECO$ = SAFETEXT$(!Doc_Orig)
             Y$ = REGBLAN$("!BOLREDET")
             Call REPLA(Y$, MKI$(CI%), 83, 2)
             Call REPLA(Y$, Unimed$(CI%), 85, 4)
             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 89, 8)
                If ANUCA% = 1 Then Call REPLA(Y$, MKD$(0), 89, 8)
             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 163, 8)
             '
             Call REPLA(Y$, Chr$(XVALO(!NuOrd_Item)), 98, 1)
             Call REPLA(Y$, MKS$(0), 100, 4) 'no tiene AE
             '
             Call REPLA(Y$, LOTECO$, 107, 12)
             Call REPLA(Y$, MKS$(NUOC&), 151, 4)
             Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
             Call REPLA(Y$, MKS$(!N_ORSERV), 189, 4)
             '
             If ESTRAZABLE(CI%) = 2 Then     ' BUSCA POR NUMEROS DE SERIE
                NOFAX$ = SAFETEXT$(!Doc_Orig)
                If NOFAX$ Like "OF*" Then
                   SQLX$ = "SELECT * FROM TANUMSE"
                   SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CI%
                   SQLX$ = SQLX$ + " AND IdSubOr = '" & NOFAX$ & "'"
                   SQLX$ = SQLX$ + "ORDER BY NUMEROSERIE "
                   Dim SELNUSE As ADODB.Recordset
                   Set SELNUSE = New ADODB.Recordset
                   Set SELNUSE = FLEXCONN.Execute(FILTSQL$(SQLX$))
                   With SELNUSE
                     Do While Not .EOF
                       If InStr(UCase$(SAFETEXT$(!UBICACION)), "PLANTA MEDITEA") < 1 Then
                         Y1$ = Y$
                         Call REPLA(Y1$, MKD$(1), 89, 8)
                           If ANUCA% = 1 Then Call REPLA(Y1$, MKD$(0), 89, 8)
                         Call REPLA(Y1$, MKD$(1), 163, 8)
                         Call REPLA(Y1$, SAFETEXT$(!NUMEROSERIE), 107, 12)
                         HAYOC% = 1
                         JJ& = JJ& + 1
                         Call GRAREG("!BOLREDET", Y1$, JJ&)
                       End If
                     .MoveNext
                     Loop
                   End With
                   GoTo 49
                End If
             End If
             HAYOC% = 1
             JJ& = JJ& + 1
             Call GRAREG("!BOLREDET", Y$, JJ&)
           End If
         End If
49      .MoveNext
       Loop
     End With
     COMPTEMP_1.Close
     Set COMPTEMP_1 = Nothing
     '
     'TENGO QUE FIJARME SI TIENE ALGUN MATERIAL EL PROV
     'CONSIGNADO EN MACONSIG CON EL MISMO Nº DE PROV
1777 Call ABRESTOCKS
     SQLX$ = "Select * from MACONSIG"
     SQLX$ = SQLX$ + " WHERE (Cant_Rendid < Cant_Entreg - 0.001 "
     SQLX$ = SQLX$ + "OR CANT_RENDID IS NULL) "
     SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
     'Set consigtmp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     Dim CONSIGTMP As ADODB.Recordset
     Set CONSIGTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With CONSIGTMP
       '
       On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Screen.MousePointer = 1
         GoTo 1129
       End If
       On Error GoTo 0
       '
       Y$ = REGBLAN$("!BOLREDET")
       Do While .EOF = False
         CI% = XVALO(!cod_inte)
         NUOC& = Val(Mid$(!nro_ocom, 4, 6))
         CAPEN = XVALO(!cant_entreg) - (XVALO(!CANT_RENDID) + XVALO(!cant_devu))
         If CAPEN >= 0.005 Then
            Call REPLA(Y$, MKI$(CI%), 83, 2)
            Call REPLA(Y$, Unimed$(CI%), 85, 4)
            Call REPLA(Y$, MKD$(0), 89, 8)
            Call REPLA(Y$, "MP " + TRIM$(CSTRING$(MKS$(CAPEN), 3, 12, 1, 2)) + " " + Unimed$(CI%), 107, 12)
            Call REPLA(Y$, MKS$(NUOC&), 151, 4)
            Call REPLA(Y$, MKS$((-1) * i&), 189, 4)
            JJ& = JJ& + 1
            Call GRAREG("!BOLREDET", Y$, JJ&)
         End If
       .MoveNext
       Loop
     End With
     CONSIGTMP.Close
     Set CONSIGTMP = Nothing
     
     '\\\OT-06-0149-WAR-25/04/06///
1129 Call ABRECOMPRAS
     SQLX$ = "Select * from AUTOENT"
     
     SQLX$ = SQLX$ + " Where Stat_Aten < 8 "
     SQLX$ = SQLX$ + " AND  Stat_Aten <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
     SQLX$ = SQLX$ + " AND Acuen_Aten < Cant_Aten - 0.001 "
     SQLX$ = SQLX$ + " AND Npro_Aten = " & NC% & " "
     SQLX$ = SQLX$ + " AND Cod_Inte > 0 " 'EXCLUYE LAS DE 0/C's TEXTO LIBRE
     'Set COMPTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     Dim COMPTEMP_2 As ADODB.Recordset
     Set COMPTEMP_2 = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With COMPTEMP_2
       On Error Resume Next
       .MoveFirst
       If Err Then
         On Error GoTo 0
         Screen.MousePointer = 1
         'HAYOC% = 0
         GoTo 1130
       End If
       On Error GoTo 0
       '
       Do While (.EOF = False)
         CAPEN = XVALO(!Cant_Aten) - XVALO(!Acuen_Aten)
         COEFUNI = 1 '!Coef_Unids
         If COEFUNI < 0.0001 Then COEFUNI = 1
         If CAPEN > 0.005 Then
           NUOC& = XVALO(!nume_ocom)
           FENTRESO% = CVINT(!Fentre_Sol)
           FESO2 = FENTRESO%
           If DIENTRE%(HOIX1, FESO2) <= DIANRE& Then
             CI% = XVALO(!cod_inte)
             Y$ = REGBLAN$("!BOLREDET")
             Call REPLA(Y$, MKI$(CI%), 83, 2)
             Call REPLA(Y$, Unimed$(CI%), 85, 4)
             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 89, 8)
             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 163, 8)
             If ANUCA% = 1 Then Call REPLA(Y$, MKD$(0), 89, 8)
             '
             Call REPLA(Y$, Chr$(XVALO(!NuOrd_Item)), 98, 1)
               NANTEN& = 0
               On Error Resume Next
               NANTEN& = XVALO(!Nume_Aten)
               On Error GoTo 0
             Call REPLA(Y$, MKS$(NANTEN&), 100, 4) 'Nº Aut. Ent.
             '
             Call REPLA(Y$, LOTECO$, 107, 12)
             Call REPLA(Y$, MKS$(NUOC&), 151, 4)
             Call REPLA(Y$, MKI$(FENTRESO%), 155, 2)
             Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
             'Call REPLA(Y$, MKS$(!N_OrServ), 189, 4)
             HAYOC% = 1
             JJ& = JJ& + 1
             Call GRAREG("!BOLREDET", Y$, JJ&)
           End If
         End If
        .MoveNext
       Loop
     End With
     COMPTEMP_2.Close
     Set COMPTEMP_2 = Nothing
     
     '
     '\\\OT-06-0149-WAR-FIN///
1130 If HAYOC% = 0 Then  '1129 antes de OT-06-0149
       Call COMUNI("No hay Ordenes de Compra Pendientes\para este Proveedor.")
     End If
     '
     Call SETULTREG("!BOLREDET", JJ&)
     Call CIERRARCH("!BOLREDET")
     If TIMATEI% = 2 Then
        Call SETULTREG("!BOLREPAP", JJ&)
        Call CIERRARCH("!BOLREPAP")
     End If
     '
1113 TIMATEFIL% = 0
     AVENTA$ = "BOLREDET": ARCHICA$ = "BOLREDET"
     If ANUCA% = 1 Then AVENTA$ = "BOLRENUE"
     If TIMATEI% = 2 Then AVENTA$ = "BOLREPAP": ARCHICA$ = "BOLREPAP": TIMATEFIL% = 1
     '
     VTB% = VENTABU%(AVENTA$ + "/NC/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(rasocli$(NC1%)) + " - " + REPLACAR$(FECHAREC$, "/", "-"))
     TIMATEFIL% = 0
     If VTB% < 0 Or ULTREG&("!" + ARCHICA$) < 1 Then
       X$ = "": Y$ = "": YY$ = ""
       EBOL$ = "": PEDY$ = "": DESCRI$ = ""
       GoTo 1100
     End If
     '
     If ARCHICA$ = "BOLREPAP" Then
        ' CALCULA LOS METROS CUADRADOS
        For YX& = 1 To ULTREG&("!" + ARCHICA$)
           YY$ = REGLEIDO$("!" + ARCHICA$, YX&)
           CI% = CVI(Mid$(YY$, 83, 2))
           If TIMATE%(CI%) <> 1 Then
               DECODE$ = TRIM$(CODEXT$(CI%)) + " - " + DESCRIT0$(CI%)
               Call MENSERR(24, "Código " + DECODE$ + "\No Corresponde a Tipo de Material 'PAPELES'")
               GoTo 1113
           End If
           ANCHOX# = CVD(Mid$(YY$, 193, 8))
           LARGOX# = CVD(Mid$(YY$, 201, 8))
           UNIDS& = CVS(Mid$(YY$, 209, 4))
           '
           If UNIDS& < 1 Then
              UNIDS& = 1
              Call REPLA(YY$, MKS$(1), 209, 4)
              Call GRAREG("!" + ARCHICA$, YY$, YX&)
           End If
           '
           If UNIDS > 99 Then
               Call MENSERR(24, "Cantidad no válida - Supera 99 Unidades.")
               GoTo 1113
           End If
           '
           CAN# = LARGOX# * ANCHOX# / 1000
           Call REPLA(YY$, MKD$(CAN#), 89, 8)
           Call GRAREG("!" + ARCHICA$, YY$, YX&)
        Next YX&
     End If
     '
     For YX& = 1 To ULTREG&("!" + ARCHICA$)
       YY$ = REGLEIDO$("!" + ARCHICA$, YX&)
       CI% = CVI(Mid$(YY$, 83, 2))
       CAN# = CVD(Mid$(YY$, 89, 8))
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           If CAN# > 0.005 Then
             GoTo 1141
           End If
         End If
       End If
     Next YX&
     Call MENSERR(24, "Recepción Cancelada !!!\Faltan Cantidades o Códigos.")
     GoTo 1113
     '
1141 CONTA% = ULTREG&("!" + ARCHICA$)
     '
     BLOSOBRE% = 0
     If Control$("BOLRECEP", "BLOSOBRE") = "SI" Then
       BLOSOBRE% = 1
       PORSOBRE = Val(Control("BOLRECEP", "PORSOBRE"))
     End If
     '
1115 If CONTA% > 0 Then
       '
       STMOVI09.List1.Clear
       For YX& = 1 To ULTREG&("!" + ARCHICA$)
         YY$ = REGLEIDO$("!" + ARCHICA$, YX&)
         CI% = CVI(Mid$(YY$, 83, 2))
         ' validacion que no tenga marca en proceso de articulos con salidas menores a 5000 e
         ' importes por cantidad salida menores a 5000
         If STATMOVI%(CI%) < 0 Then
              ' ARMA UN LIST INFORMATIVO CON LOS ITEMS CON STATMOV - 1
              TX$ = ""
              Call REPLA(TX$, CODEXT$(CI%), 1, 16)
              Call REPLA(TX$, DESCRIT$(CI%), 17, 36)
              STMOVI09.List1.AddItem TX$
              ' LUEGO DE RECORRER LO MUESTRA
         End If

         CAN# = CVD(Mid$(YY$, 89, 8))
         UNIDS& = CVS(Mid$(YY$, 209, 4))
             If UNIDS& < 1 Then UNIDS& = 1
         If CI% < 1 Then
           Call MENSERR(24, "Código no Válido.")
           GoTo 1113
         End If
         If ANUCA% < 1 Then
           If CAN# < 0.005 Then
             Call MENSERR(24, "Falta Cantidad.")
             GoTo 1113
           End If
         End If
         '
         ' VERIFICA TRENSFERENCIA DE COMPONENTES
         If CAN# >= 1 Then
            If ESTRAZABLE(CI%) = 2 Then
               LOTECO$ = TRIM$(Mid$(YY$, 107, 12))
               If LOTECO$ <> "" Then
                  DEPARMAD% = XVALO(Control("", "DEPOARMA")) ' DEPOSITO DE ARMADOR
                  CONDII$ = "Cod_Inte = " & CI% & " and NUMEROSERIE = '" & LOTECO$ & "'"
                  NUMEORF$ = SAFETEXT$(VALOBADA("TANUMSE", "IdSubOr", CONDII$, "NOMESS"))
                  If TRIM$(NUMEORF$) <> "" Then
                     CONDI1$ = "CANTINGRE > 0 AND DepoMovi = " & DEPARMAD% & " And DoSecCor = '" & NUMEORF$ & "' AND CODIMOVI = 'TR'"
                     If CantRegistros("MOVISTO", CONDI1$, "NOMESS") < 1 Then
                        Call MENSERR(24, "Imposible Procesar Recepción de Armador Externo.\  \Aun Falta Transferencia Inicial de Materiales Componentes")
                        GoTo 1113
                     End If
                  End If
               End If
            End If
         End If
         '
         If BLOSOBRE% > 0 Then
           If Mid$(YY$, 107, 2) <> "MP" Then ' no es devolucion de MP
             '\\\OT-05-0694-WAR-25/10/05///
             'SI ES MATERIAL AUTORIZADO A RECIBIR SIN O/C
             X2$ = FILTERGETRECORD$("TAMATCON", 1, MKI$(CI%))
             If TRIM$(X2$) <> "" Then
               FELI% = CVI(Mid$(X2$, 3, 2))
               If FF% <= FELI% Then 'SI VALIDA FECHA LIMITE
                 GoTo 1114
               End If
             End If
             '\\\OT-05-0694-WAR-FIN///
             If CAN# * UNIDS& > CVD(Mid$(YY$, 163, 8)) * (1 + PORSOBRE / 100) + 0.05 Then
               DECODE$ = TRIM$(CODEXT$(CI%)) + " - " + DESCRIT0$(CI%)
               If DERACCE%("RECSOBRE/NOMESS", "Recepción con Sobre-Embarque") >= 2 Then
                   If COMALTER%("ATENCIóN: Sobre-Embarque en Código\" + DECODE$ + ".\\Corregir\Continuar") <> 2 Then GoTo 1113
                 Else
                   Call MENSERR(24, "Configuración y Derechos de Acceso\Impiden Sobre-Embarque\en Código " + DECODE$ + ".")
                   GoTo 1113
               End If
             End If
           End If
         End If
1114   Next YX&
       '
       If STMOVI09.List1.ListCount > 0 Then
         If COMALTER%("Se Han Detectado Items\ Con Estado de Movimiento de Baja \\Ver Detalle\Continuar") = 1 Then
            STMOVI09.Label2 = ""
            STMOVI09.Show 1 ' MUESTRA DETALLE
            If STMOVI09.Label2 = "MODIFICAR" Then ' SI ELIGIO MODIFICAR VUELVE A VENTANA
              GoTo 1113
            End If
         End If
       End If
       Call SETULTREG("!OBSERVOF", 0)
       Call CIERRARCH("!OBSERVOF")
       '
       VTB% = VENTABU%("OBSERVOF/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(rasocli$(NC1%)) + " - " + FECHAREC$)
       If VTB% < 0 Then
         If COMALTER%("Realmente Desea Cancelar ?\\Continuar\Cancelar") = 2 Then
           X$ = "": Y$ = "": YY$ = ""
           EBOL$ = "": PEDY$ = "": DESCRI$ = ""
           GoTo 1100
         End If
       End If
       '
1125   Call ABRESTOCKS
       Call BLOQARCH("BOLRECEP")
       TINI = Timer
51     If Timer < TINI + 3 Then
         GoTo 51
       End If
       Call BLOQARCH("BOLRECEP")
       '
       SQLX$ = " SELECT max(NumeBoRe) as maximo from BOLRECEP"
       Dim BolreTemp_1 As ADODB.Recordset
       Set BolreTemp_1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
       If Not IsNull(BolreTemp_1!Maximo) Then
           NROBOL& = 1 + (BolreTemp_1!Maximo)
       End If
       BolreTemp_1.Close
       Set BolreTemp_1 = Nothing
       '
       Call CreaRegistro("BOLRECEP", "NumeBoRe", NROBOL&, "NOMESS")
       Call ACTUREGISTRO("BOLRECEP", "STATUS", "NumeBoRe=" & NROBOL&, 999, REG&, "NOMESS")
       TINI = Timer
52     If Timer < TINI + 3 Then
         GoTo 52
       End If
       Call LIBARCH("BOLRECEP")
       '
       Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
       Call ABRESTOCKS
       Call ABRECOMPRAS
       '
1126   AINDI% = 0
       NORVENTA = ""
       NUOR% = 0
       For YX& = 1 To ULTREG&("!" + ARCHICA$)
         YY$ = REGLEIDO$("!" + ARCHICA$, YX&)
         CI% = CVI(Mid$(YY$, 83, 2))
         CAN# = CVD(Mid$(YY$, 89, 8))
         UNIDS& = CVS(Mid$(YY$, 209, 4))
             If UNIDS& < 1 Then UNIDS& = 1
         ANCHOX# = CVD(Mid$(YY$, 193, 8))
         LARGOX# = CVD(Mid$(YY$, 201, 8))
         NUOC& = Abs(CVS(Mid$(YY$, 151, 4)))
         NUMOIT% = Asc(Mid$(YY$, 98, 1))   'EL Nº DE ORDEN DEL ITEM
         If CI% > 0 Then
           If CAN# > 0.005 Then
             CARECI = CAN# * UNIDS&
             LOTECO$ = TRIM$(Mid$(YY$, 107, 12))
             NUORSE& = 0
             '
             If CVD(Mid$(YY$, 163, 8)) >= 0.05 Then   ' solo si es comprado - no M.P. devuelto
               NATEN& = CVS(Mid$(YY$, 100, 4))
               
               If NATEN& = 0 Then 'O/C CERRADA'\\\OT-06-0149-WAR-26/04/06///
                 strSQL = "Select * from Ocomdeta where Nume_Ocom = " & NUOC&
                 strSQL = strSQL + " AND NuOrd_Item = " & NUMOIT%
                 strSQL = strSQL + " AND Stat_Ocom < 8 "
                 
                 Dim OCOMDETA As ADODB.Recordset
                 Set OCOMDETA = New ADODB.Recordset
                 OCOMDETA.CursorType = adOpenDynamic
                 OCOMDETA.LockType = adLockPessimistic
                 OCOMDETA.Open FILTSQL$(strSQL), FLEXCONN, , , adCmdText
                 If Not (OCOMDETA.EOF And OCOMDETA.BOF) Then 'si no esta vacio actualiza
                  With OCOMDETA
                     If IsNull(!IContab) Then
                          IMPCONOC% = 0
                        Else
                          IMPCONOC% = XVALO(!IContab) 'ASIGNO LA IMPUTACION CONTABLE
                     End If
                       COEFUNI = XVALO(!Coef_Unids)
                       If COEFUNI < 0.0001 Then COEFUNI = 1
                     !Cant_Rec = XVALO(!Cant_Rec) + CARECI / COEFUNI
                       CAPENDIEN = XVALO(!cant_ocom) - XVALO(!Cant_Rec)
                       NUORSE& = XVALO(!N_ORSERV)
                       If CAPENDIEN >= 0.05 Then
                           If COMALTER%("Esta O/C por Codigo " + TRIM$(CODEXT$(CI%)) + "\Queda con un Saldo Pendiente de " + TRIM$(FORMATNUM$(CAPENDIEN, "F10.1")) + " " + TRIM$(Unimed$(CI%)) + ".\\Mantener Abierta\Cancelar Pendiente") = 2 Then
                              !Stat_Ocom = 8
                           End If
                        Else
                           !Stat_Ocom = 3
                       End If
                     .Update
                  End With
                  OCOMDETA.Close
                  Set OCOMDETA = Nothing
                 End If
              ElseIf NATEN& > 0 Then '0/C ABIERTA (AE)
                 strSQL = "Select * from DetoCabi where Nume_Ocom = " & NUOC&
                 strSQL = strSQL + " AND Cod_Inte = " & CI%
                 strSQL = strSQL + " AND Stat_Ocom < 8 "
                 
                 Dim DetoCabi As ADODB.Recordset
                 Set DetoCabi = New ADODB.Recordset
                 DetoCabi.CursorType = adOpenDynamic
                 DetoCabi.LockType = adLockPessimistic
                 DetoCabi.Open FILTSQL$(strSQL), FLEXCONN, , , adCmdText
                 If Not (DetoCabi.EOF And DetoCabi.BOF) Then 'si no esta vacio actualiza
                    With DetoCabi
                        If IsNull(!IContab) Then
                             IMPCONOC% = 0
                           Else
                             IMPCONOC% = XVALO(!IContab) 'ASIGNO LA IMPUTACION CONTABLE
                        End If
                          COEFUNI = XVALO(!Coef_Unids)
                          If COEFUNI < 0.0001 Then COEFUNI = 1
                          !Cant_Rec = XVALO(!Cant_Rec) + CARECI / COEFUNI
                          CAPENDIEN = XVALO(!cant_ocom) - XVALO(!Cant_Rec)
                          NUORSE& = XVALO(!N_ORSERV)
                        If CAPENDIEN < 0.05 Then
                          !Stat_Ocom = 3
                        End If
                        .Update
                    End With
                 End If
                 DetoCabi.Close
                 Set DetoCabi = Nothing
                 '
                 strSQL = "Select * from AutoEnt where Nume_Ocom = " & NUOC&
                 strSQL = strSQL + " AND NuOrd_Item = " & NUMOIT%
                 strSQL = strSQL + " AND Stat_Ocom < 8 "
                 strSQL = strSQL + " AND Nume_Aten= " & NATEN&
                
                 Dim AutoEnt As ADODB.Recordset
                 Set AutoEnt = New ADODB.Recordset
                 AutoEnt.CursorType = adOpenDynamic
                 AutoEnt.LockType = adLockPessimistic
                 AutoEnt.Open FILTSQL$(strSQL), FLEXCONN, , , adCmdText
                 If Not (AutoEnt.EOF And AutoEnt.BOF) Then 'si no esta vacio actualiza
                   With AutoEnt
                       COEFUNI = 1 ' !Coef_Unids
                     If COEFUNI < 0.0001 Then COEFUNI = 1
                       !Acuen_Aten = XVALO(!Acuen_Aten) + CARECI / COEFUNI
                       CAPENDIEN = XVALO(!Cant_Aten) - XVALO(!Acuen_Aten)
                       !Sald_Aten = CAPENDIEN
                     If CAPENDIEN < 0.05 Then
                       !Stat_Aten = 3
                     End If
                     .Update
                   End With
                 End If
                 AutoEnt.Close
                 Set AutoEnt = Nothing
               End If
             End If
             '
             ' aca arma un string con las ordenes de compra recibidas
             NORVE& = Abs(CVS(Mid$(YY$, 151, 4)))
             If NORVE& > 0 Then
               If NORVE& <> NORVEA& Then
                 NORVEA& = NORVE&
                 If NORVENTA <> "" Then NORVENTA = NORVENTA + " - "
                 NORVENTA = NORVENTA + TRIM$(Str$(NORVE&))
               End If
             End If
             '
             NUMSE$ = "" ' ESCRIBO NUMERO DE SERIE COMO REFERENCIA
             If ESTRAZABLE(CI%) > 1 Then
                NUMSE$ = "S-Nº: " + TRIM$(Mid$(YY$, 107, 12)) + " - "
             End If
             
             DESCRI$ = TRIM$(CODEXT$(CI%)) + " - " + NUMSE$ + TRIM$(DESCRIT$(CI%))
             If AINDI% = 0 Then
                X$ = REGBLAN$("INDIBOL")
                Call REPLA(X$, MKS$(NROBOL&), 1, 4)
                FEX = FF%
                Call REPLA(X$, FECHATEX$(FEX), 5, 8)
                Call REPLA(X$, REMI$, 15, 16)
                Call REPLA(X$, ECOARCH$("PROVED1", MKI$(NC%)), 32, 29)
                Call REGAPP("INDIBOL", X$)
                If NUORSE& < 1 Then Call REGAPP("IBOLPEN", X$)
                AINDI% = 1
             End If
             '
             Call CIERRARCH("INDIBOL")
             Call CIERRARCH("IBOLPEN")
             '
             Call REPLA(YY$, MKS$(NROBOL&), 1, 4)
             If TIMATEI% <> 2 Then Call REPLA(YY$, DESCRI$, 5, 64)
             '
             If TRIM$(Mid$(YY$, 69, 12)) = "" Then
               NBR$ = TRIM$(Str$(NROBOL&))
               NUOR% = NUOR% + 1 ' AGREGADO PARA QUE GUARDE EL ORDEN DE LAS QUE SE RECIBEN
               IDENLO$ = "BR-" + NBR$ + "-" + TRIM$(Str$(NUOR%))
               Call REPLA(YY$, IDENLO$, 69, 12)
             End If
             Call REPLA(YY$, MKI$(FF%), 81, 2)
             DEPX% = DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(CI%)
             Call REPLA(YY$, Chr$(DEPX%), 97, 1)
             Call REPLA(YY$, MKI$(NC%), 105, 2)
             Call REPLA(YY$, REMI$, 119, 16)
             If CVD(Mid$(YY$, 163, 8)) >= 0.05 Then   ' solo si es comprado - no M.P. devuelto
               Call REPLA(YY$, MKS$(NUORSE&), 189, 4)
             End If
             Call GRAREG("!" + ARCHICA$, YY$, YX&)
             '
            strSQL = " Select * from Bolrecep"
            strSQL = strSQL + " where NumeBoRe=" & NROBOL&
            Dim Bolrecep As ADODB.Recordset
            Set Bolrecep = New ADODB.Recordset
            Bolrecep.CursorType = adOpenDynamic
            Bolrecep.LockType = adLockPessimistic
            Bolrecep.Open FILTSQL$(strSQL), FLEXCONN, , , adCmdText
            '
            With Bolrecep
              For IUNI% = 1 To UNIDS&
                NBR$ = TRIM$(Str$(NROBOL&))
                NORDE% = NORDE% + 1
                NORDES$ = TRIM$(Str$(NORDE%))
                IDENLO$ = "BR-" + NBR$ + "-" + NORDES$
                '
                .AddNew   ' YA CONVERTIDO
                On Error Resume Next
                !CODIEMPR = CodiEmp%
                On Error GoTo 0
                 !NUMEBORE = NROBOL&
                 !Descrip = Left$(DESCRI$, 64)
                 !idenlote = Left$(IDENLO$, 12)
                 !FechRecep = CVDAT(FF%)
                 !cod_inte = CI%
                 !Cod_Exte = Left$(CODEXT$(CI%), 24)
                 !descripitem = Left$(DESCRIT0$(CI%), 64)
                 !Unidad = Left$(Unimed$(CI%), 4)
                 !CantNom = CAN#
                 !M_Ancho = ANCHOX#
                 !M_Largo = LARGOX#
                 !DEPOENT = DEPX%
                 !NumProv = NC%
                 !Raso_Prov = Left$(rasocli$((-1) * NC%), 64)
                    LOTECO$ = TRIM$(LOTECO$)
                    While Len(LOTECO$) > 12
                      LOTECO$ = Mid$(LOTECO$, 2)
                    Wend
                 !LOTECOL_PROV = Left$(LOTECO$, 12)
                 !NumConVig = Left$(NUCONVI$, 16)
                 !NumRem_Prov = Left$(REMI$, 16)
                 !NumFac_Prov = ""
                 !NUMOC = NUOC&
                 !IContab = IMPCONOC% 'Imputación contable de la OC.
                 !MoneVaUni = ""
                 !Valo_Unit = 0
                 !CanToApro = 0
                 !CanToRech = 0
                 !CanToFalta = 0
                 !FechLibe = CVDAT(0)
                 !NUMEORDSER = NUORSE&
                 !Observa = " " 'OBSE$
                 !Status = 0
                 !USUAPRO = " "
                 !FeHoApro = CVDAT(0)
                 If NUORSE& > 0 Then
                    !CanToApro = CAN#
                    !Status = 1
                    !USUAPRO = USERNAME$
                    !FeHoApro = CVDAT(HOY(HOS$))
                 End If
                 '\\\OT-06-0149-WAR-26/04/06///
                 !Nume_Aten = NATEN&
                 '\\\OT-06-0149-WAR-FIN///
                .Update
              Next IUNI%
            End With
            Bolrecep.Close
            Set Bolrecep = Nothing
            '
            If NUORSE& < 1 Then
               MODO$ = "CC": If Left$(UCase$(LOTECO$), 2) = "MP" Then MODO$ = "DV"
               Call DESCACONSI(NC%, CI%, CAN#, NUOC&, FF%, NROBOL&, IDENLO$, MODO$)
            End If
            '
            If NUORSE& > 0 Then
               If CAN# > 0.05 Then
                 GoSub BLANTEROP
               End If
            End If
            '
           End If
         End If
       Next YX&
       '
       Call BORRAREGISTROS("BOLRECEP", "STATUS=999", REG&, "NOMESS")
       Call AJUMACON
       '
       ' AQUI RESOLVER TEMA DE NUMEROS DE SERIE
       ENVIAMAILCALIDAD% = 0
       '
       For YX& = 1 To ULTREG&("!" + ARCHICA$)
         YY$ = REGLEIDO$("!" + ARCHICA$, YX&)
         CI% = CVI(Mid$(YY$, 83, 2))
         CANTUI& = CVD(Mid$(YY$, 89, 8))
         If CANTUI& >= 1 Then
            If ESTRAZABLE(CI%) = 2 Then
               LOTECO$ = TRIM$(Mid$(YY$, 107, 12))
               If LOTECO$ <> "" Then
                  CONDI$ = "Cod_Inte = " & CI% & " and NUMEROSERIE = '" & LOTECO$ & "'"
                  '**** ESTO ES UNA SOLUCION PROVISORIA !!!!
                  'QUITA NOMESS PARA QUE PRESENTE EL MENSAJE QUE CAUSA QUE NO SE ACTUALICE
10                Call ACTUREGISTRO("TANUMSE", "UBICACION", CONDI$, "Planta Meditea", RAFE&)
                  'SE INCORPORA VERIFICACION PARA GARANTIZAR QUE SE ACTUALIZO LA TABLA Y SI REINTENTA
                  UBICA$ = TRIM$(SAFETEXT$(VALOBADA("TANUMSE", "UBICACION", CONDI$)))
                  If UCase(UBICA$) <> "PLANTA MEDITEA" Then
                    OP% = COMALTER%("No se Ha Podido Ejecutar La Consulta SQL.\¿Desea Intentarlo Nuevamente?\\No, Continuar\Si, Reintentar")
                    If OP% = 2 Then GoTo 10
                  End If
                  '**** FIN SOLUCION PROVISORIA !!!!
                  '
                  ENVIAMAILCALIDAD% = 1
                  CUERPO$ = "Recepcion de " + rasocli$((-1) * NC%)
                  '
                  Call ANOTANUSE(LOTECO$, CUERPO$)
                  '
'                  Call TRANSFIERE_COMPONENTES(CI%, LOTECO$, "BR-" + TRIM$(Str$(NROBOL&)))
                  '
               End If
            End If
         End If
       Next YX&
       '
       If InStr(EMPREAC$, "MEDITEA") > 0 Then
            'Call ALERTATRANSFE
            'comentado por que dijo que no lo usan y para acelerar los tiempos
       End If
       '
       If TRIM$(FORIPRE$) <> "" Then
         '
         CANCELPRINT% = 0
         If Control("BOLRECEP", "SUPRINT") = "SI" Then CANCELPRINT% = 1
         '
         TIPODOC$ = "Boleta de Recepción"
         DESTIDOC% = (-1) * Abs(NC1%)
         NUMEDOC$ = Str$(NROBOL&)
         FEX = FF%: FECHDOC$ = FECHATEX$(FEX)
         '
         FACTU$ = TRIM$(Mid$(PEDY$, 129, 16)) ' ASIGNA NUMERO DE FACTURA QUE TOMA DEL OBJETO
         '
         CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = DEPOSAL$
         CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = REMI$
         CODPARAM$(3) = "N-AUTENT": DOCPARAM$(3) = NORVENTA
         CODPARAM$(4) = "NRO-PASE": DOCPARAM$(4) = NUCONVI$
         CODPARAM$(5) = "DOCUORIG": DOCPARAM$(5) = FACTU$
         '
         CAPARDOC% = 5
         '
         CODTABU1$(1) = "COD-MAT"
         CODTABU1$(2) = "DES-MAT"
         CODTABU1$(3) = "COLOR"
         CODTABU1$(4) = "UNIMED"
         CODTABU1$(5) = "NRO-LOTE"
         CODTABU1$(6) = "CANTIDAD"
         CODTABU1$(7) = "N-OCOMPR"
         CODTABU1$(8) = "CAKILOS"
         CODTABU1$(9) = "F-PLANO"
         CODTABU1$(10) = "FECH-REC"
         CODTABU1$(11) = "ORD-ITEM"
         CODTABU1$(12) = "UBI-MPRI"
         CACOLTAB1% = 12
         '
         Call LIMPIATETAB
         '
         CAITAB1% = 0
         ORDITEM% = 0
         '
         'CONTROLES DE IMPRESION DE DESCRIPCION
         DESCPRI% = 0: DESCOMP% = 0: Especif% = 0
         If Control("BOLRECEP", "DESCPRIM") = "SI" Then DESCPRI% = 1
         If Control("BOLRECEP", "DESCOMPL") = "SI" Then DESCOMP% = 1
         If Control("BOLRECEP", "ESPECIFI") = "SI" Then Especif% = 1
         If DESCOMP% + Especif% < 1 Then DESCPRI% = 1 'POR SI NO SE SELECCIONO NADA
         '
         For YX& = 1 To ULTREG&("!" + ARCHICA$)
           YY$ = REGLEIDO$("!" + ARCHICA$, YX&)
           UNIDS& = CVS(Mid$(YY$, 209, 4))          'Lee del archivo la cantidad de unidades
           CABOBIN& = UNIDS&: If CABOBIN& < 1 Then CABOBIN& = 1  'Asigna a cabobin el total de unidades, si es "" asigna 1
           CI% = CVI(Mid$(YY$, 83, 2))
           CAN# = CVD(Mid$(YY$, 89, 8)) * CABOBIN&
           If CI% > 0 Then
             If CI% <= NUMAS% Then
               If CAN# > 0.005 Then
                 '
                 UNID$ = TRIM$(Mid$(YY$, 85, 4))
                 If UNID$ = "" Then UNID$ = Unimed$(CI%)
                 CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
                 KILAJE$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 10, 3, 2)
                 FEX = CVI(Mid$(YY$, 155, 2))
                 FESOLEN$ = FECHATEX$(FEX)
                 FEPLA$ = FEPLANO$(CI%)
                 NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
                 ' SON EQUIPOS CON NUMEROS DE SERIE Y SE INGRESA COMO CANTIDAD <> 1 LE ASIGNA 1
                 If ESTRAZABLE%(CI%) > 1 And XVALO(CANS$) <> 1 And NULOTE$ <> "" Then
                    CANS$ = "1"
                 End If
                 '
                 If NULOTE$ = "" Then
                    NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
                 End If
                 NUOCO$ = FORMATNUM$(Abs(CVS(Mid$(YY$, 151, 4))), "F7,0")
                 '
                 CAITAB1% = CAITAB1% + 1
                 TETABU1$(1, CAITAB1%) = TRIM$(CODEXT$(CI%))
                 '
                 DESPRIN$ = DESCRIT$(CI%)
                 DESCOMPLEM$ = DESCRIT2$(CI%)
                 ESPECIFICAC$ = ESPECIFIT$(CI%)
                 '
                 If FORMPLA$ = "RECEP-ESPUMA" Then
                       DESPRIN$ = Mid$(YY$, 5, 64)
                 End If

                 DIMPRE$ = ""
                 If DESCPRI% > 0 Then
                   If DESPRIN$ <> "" Then
                      DIMPRE$ = DESPRIN$
                   End If
                 End If
                 '
                 'DESCRIPCION COMPLEMENTARIA
                 If DESCOMP% > 0 Then
                   If DESCOMPLEM$ <> "" Then
                     If DIMPRE$ <> "" Then
                       DIMPRE$ = DIMPRE$ + vbCrLf
                     End If
                     DIMPRE$ = DIMPRE$ + DESCOMPLEM$
                   End If
                 End If
                 '
                 'ESPECIFICACION
                 If Especif% > 0 Then
                   If ESPECIFICAC$ <> "" Then
                     If DIMPRE$ <> "" Then
                       DIMPRE$ = DIMPRE$ + vbCrLf
                     End If
                     DIMPRE$ = DIMPRE$ + ESPECIFICAC$
                   End If
                 End If


''''                 'DESCRIPCION COMPLEMENTARIA
''''                 If DESCOMP% > 0 Then
''''                    If TRIM$(DESCRIT2$(CI%)) <> "" Then
''''                       If DIMPRE$ <> "" Then DIMPRE$ = DIMPRE$ + vbCrLf
''''                       DIMPRE$ = DIMPRE$ & DESCRIT2$(CI%)
''''                    End If
''''                 End If
''''                 '
''''                 'ESPECIFICACION
''''                 If ESPECIF% > 0 Then
''''                    If TRIM$(ESPECIFIT$(CI%)) <> "" Then
''''                       If DIMPRE$ <> "" Then DIMPRE$ = DIMPRE$ + vbCrLf
''''                       DIMPRE$ = DIMPRE$ & ESPECIFIT$(CI%)
''''                    End If
''''                 End If
                 '
                 ORDITEM% = ORDITEM% + 1
'                 TETABU1$(2, CAITAB1%) = DIMPRE$
                 TETABU1$(3, CAITAB1%) = COLOR$(CI%)
                 TETABU1$(4, CAITAB1%) = UNID$
                 TETABU1$(5, CAITAB1%) = NULOTE$
                 TETABU1$(6, CAITAB1%) = CANS$
                 TETABU1$(7, CAITAB1%) = NUOCO$
                 TETABU1$(8, CAITAB1%) = KILAJE$
                 TETABU1$(9, CAITAB1%) = FEPLA$
                 TETABU1$(10, CAITAB1%) = FESOLEN$
                 TETABU1$(11, CAITAB1%) = ORDITEM%
                 TETABU1$(12, CAITAB1%) = UBISTO$(CI%)
                 '
                 'IMPRESION DE LA DESCRIPCION CON TODAS SUS VARIANTES CONFIGURADAS
                 FORTELIB$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")))
                 ATELIB% = XVALO(Mid$(FORTELIB$, 2))
                 Call FRATEXTO(DIMPRE$, ATELIB%)
                 For KKU1% = 1 To CARETEX%
                      TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
                      CAITAB1% = CAITAB1% + 1
                 Next KKU1%
                 'r
                 'LINEA DE SEPARACION
                 Call CARLINSEP
                 '
                 If TICOMPRO$ = "FICHA" Then
                   Call PRINTDOC("FORMBREC")
                 End If
                 '
               End If
             End If
           End If
         Next YX&
         '
         If TICOMPRO$ <> "FICHA" Then
           Call PRINTDOC("FORMBREC")
         End If
         '
         If ENVIAMAILCALIDAD% > 0 Then
            'ENVIO DE MAIL DE AVISO DE RECEPCION DE EQUIPOS CON NUMERO DE SERIE
            EMAI$ = Control("OPVALEMA", "CONTMAIL")
            If EMAI$ = "" Or InStr(EMAI$, "@") < 1 Then
               GoTo 11999
            End If
            '
            SENDBYMAIL$ = EMAI$
            For Each PRX1 In Printers
             If InStr(UCase$(PRX1.DeviceName), "PDFWRITER") > 0 Then GoTo 50
            Next
            Call MENSERR(24, "Imposible Enviar por Mail.\   \Falta Impresora Virtual PDFWRITER.")
            SENDBYMAIL = ""
            GoTo 11999
50          Call PRINTDOC("FORMBREC")
         End If
       End If

     End If
          '
11999 Call CIERRARCH("*.*")
     '
     X$ = "": Y$ = "": YY$ = ""
     EBOL$ = "": PEDY$ = "": DESCRI$ = ""
     '
     CANCELPRINT% = 0 'QUE NO CANCELE LA IMPRESIÓN DEL REMITO DE DEVOLUCIÓN \\\OT-05-0826-WAR-21/11/05///
     '
     GoTo 1100
     '
BLANTEROP:
   '
   FINTER& = ULTREG&("OPERTER")
   For U& = FINTER& To 1 Step -1
     X0$ = REGLEIDO$("OPERTER", U&)
     '///OT-8-667-04/09/08
     If CVS(Mid$(X0$, 75, 4)) = NUORSE& Then
       REGITER& = U&
       GoTo 21
     End If
   Next U&
   GoTo 29
   '
21 ' Marca como "Cumplida" Operación en OPERTER.DAT
   SALRE# = CVD(Mid$(X0$, 79, 8)) - CAN#
   If SALRE# < 0 Then SALRE# = 0
   '
   FEBA% = HOY(HOS$)
   HOBA% = Int(HORANUM(Time$))
   USUA% = USNBR% Mod 100
   DBAJA$ = MKI$(FEBA%) + MKI$(HOBA%) + Chr$(USUA%) + Chr$(1) + "B.Rec." + TRIM$(Str$(NROBOL&))
   '
   NUOROP% = Asc(Mid$(X0$, 67, 1))
   Call REPLA(X0$, MKD$(CAN#), 87, 8)
   Call REPLA(X0$, MKD$(SALRE#), 95, 8)
   Call REPLA(X0$, "3ROS  ", 151, 6)
   Call REPLA(X0$, MKI$(FEBA%), 157, 2)
   Call REPLA(X0$, MKI$(Int(HORANUM(Time$))), 159, 2)
   Call REPLA(X0$, MKS$(0), 161, 4)
   Call REPLA(X0$, String$(8, 0), 165, 8)
   Call REPLA(X0$, MKS$(0), 173, 4)
   Call REPLA(X0$, DBAJA$, 177, 22)
   Call GRAREG("OPERTER", X0$, REGITER&)
   '
' Si se recibio un parcial, genera registro en OPERTER.DAT
' por el saldo pendiente
   If SALRE# > 0.01 Then
      NUOROP% = 1 + Asc(Mid$(X0$, 67, 1))
      CIIK% = CVI(Mid$(X0$, 13, 2))
      NOPE% = CVI(Mid$(X0$, 15, 2))
      DALTA$ = Chr$(NUOROP%) + Mid$(DBAJA$, 2, 5)
      INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
      PROGBLA$ = Space$(6) + String$(26, 0)
      MINVARI = TIESTAN(CIIK%, NOPE%, SALRE#, NUOROP%)
      '
      Call REPLA(X0$, DALTA$, 67, 6)
      Call REPLA(X0$, MKD$(SALRE#), 79, 8)
      Call REPLA(X0$, MKD$(0), 87, 8)
      Call REPLA(X0$, MKD$(SALRE#), 95, 8)
      Call REPLA(X0$, String$(32, 0), 103, 32)
      Call REPLA(X0$, MKS$(MINVARI), 111, 4)
      Call REPLA(X0$, MKS$(MINVARI), 131, 4)
      Call REPLA(X0$, String$(16, 0), 135, 16)
      Call REPLA(X0$, INFOBLA$, 151, 48)
      Call REPLA(X0$, PROGBLA$, 225, 32)
      Call REGAPP("OPERTER", X0$)
      AGREGA% = 1
   End If
   '
29 Return
   '
End Sub

Private Sub mnuAdjuntarDocumentosExternos_Click()

    Call Recepcion_AdjuntarDocumentosExternos(0)

End Sub

Private Sub mnuAnulaBolRecProv_Click()
   Call ANUBORE
End Sub

Private Sub mnuAprobaciónBR_Click()
        If InStr(1, EMPREAC$, "ARCON") < 1 Then
            Call COMUNI("Opción No Disponible en la Presente Versión")
            Exit Sub
        End If
        Call APRORECEP
End Sub

Private Sub mnuBoletaRecepciónProvisoria_Click()
  Call Command7_Click
End Sub

Private Sub mnuBolRec_Click()
    Call Command3_Click
End Sub

Private Sub mnuClasificacionMaterialesRecibidos_Click()

    CLASIMATRE.Show 1
    
End Sub

Private Sub mnuConfig_Click()
    Call Command14_Click
End Sub

Private Sub mnuConfigReportes_Click()
   Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
End Sub

Private Sub mnudevtransf_Click()
   'SI ESTA EN MEMORIA LO BORRA.
   If IsFormLoaded(BCODEGR09) = True Then
      Unload BCODEGR09
   End If
   BCODEGR09.DEVOLUCION% = 1
   BCODEGR09.Caption = "Devoluciones de Sector Productivo"
10 BCODEGR09.Show 1
   If BCODEGR09.VALI_BCODE09 = "APROB" Then
      Call TRANSFALEJPA(OK%)
      If OK% < 0 Then GoTo 10
   End If
   Screen.MousePointer = 1

End Sub

Private Sub mnuIngresoDatoDespacho_Click()
    ReceDespa.Show 1
End Sub

Private Sub mnuInspyAprobacion_Click()
   Call Command2_Click
End Sub

Private Sub mnuListadoDeClasificaciones_Click()

    Call CLASIMATRE.ListadoDeClasificaciones

End Sub

Private Sub mnumaterialenconsignacion_Click()
   Call Command5_Click
End Sub

Private Sub mnuModificarRemitoEnRecepcion_Click()
    
    Screen.MousePointer = 11
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Recepcion/A12;Fecha/D8;Remito/A20;Proveedor/A32", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Recepcion/A12;Fecha/D8;Remito/A20;Proveedor/A32", FRMZELECHO)
    FRMZELECHO.Caption = "Recepciones de materiales"

    Dim obj As New RECORXET
    Set rs1 = obj.RS_BOLRECE()
    Call Carga_MSF_Recordset(rs1, "Recepcion/A12;Fecha/D8;Remito/A20;Proveedor/A32", FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0

    FRMZELECHO.Width = 10000
    Screen.MousePointer = 1
50  FRMZELECHO.Show 1

    NumeroRecepcion$ = TRIM(RESP_ZELE_VECT(1))
    If NumeroRecepcion$ = "" Then Exit Sub

    Screen.MousePointer = 11
    Remito$ = TRIM(RESP_ZELE_VECT(3))
    Proveedor$ = TRIM(RESP_ZELE_VECT(4))
    Screen.MousePointer = 1

    RemitoModificado$ = TRIM$(InputBox("Remito Proveedor", "BR: " & NumeroRecepcion$ & " | Modificar Remito Proveedor", Remito$))

    Screen.MousePointer = 11
    If RemitoModificado$ <> "" And RemitoModificado$ <> Remito$ Then
        Call ACTUREGISTRO("bolrecep", "numrem_prov", "numebore=" & NumeroRecepcion$, RemitoModificado$, REGAF&)
        REFE$ = "Rto." + RemitoModificado$ + " - " + Proveedor$
        DOPRILA$ = "BR." & NumeroRecepcion$
        Call ACTUREGISTRO("movisto", "refercor", "dopricor='" & DOPRILA$ & "'", REFE$, REGAF&)
        Call ACTUREGISTRO("movisto", "doseccor", "dopricor='" & DOPRILA$ & "'", RemitoModificado$, REGAF&)
        Call ACTUREGISTRO("movisto", "doseclar", "dopricor='" & DOPRILA$ & "'", RemitoModificado$, REGAF&)
    End If
    Screen.MousePointer = 1

    Call COMUNI("Cambio de Remito Registrado")
    
End Sub

Private Sub mnupedidosyordenesdecompra_Click()
   Call Command4_Click
End Sub

Private Sub mnurecepcionesconrechazo_Click()
   Call Command16_Click
End Sub

Private Sub mnurecepcionessinocompra_Click()
   Call Command15_Click
End Sub

Private Sub mnuRecepcionesxrangofecha_Click()

    Call Command11_Click
     
End Sub

Private Sub mnuRecepcionMaterialAClasificar_Click()

    Dim MSF As msFlexGrid
    
    Call DESHASFECHA(Des%, Has%, PERIO$)
    If PERIO$ = "" Then GoTo 99
    Desde$ = FETEXCOR1$(Des%)
    Hasta$ = FETEXCOR1$(Has%)

    Campos$ = "Recepcion/A18;Fecha/D8;Remito/A20;Patente/A16;Cta/I6;Proveedor/A24;Código/A16;Descripción/A32;Kg Remito/F10.1;Kg Balanza/F10.1;Kg Control/F10.1;NumeroRecep/I0;Documentos/A10"
'    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
'    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)

    ConsultaSql$ = "select * from bolrecep with(nolock)" _
    & " where fechrecep>='" & Desde$ & "' and fechrecep<='" & Hasta$ & "'" _
    & " order by numebore"
    Set Rst = READSET(ConsultaSql$)
    
    With Rst
        Do While Not .EOF
            
            GRILLAGRAL.GRID.Rows = GRILLAGRAL.GRID.Rows + 1
                LoteRecepcion$ = SAFETEXT$(!idenlote)
            GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 1) = LoteRecepcion$
            GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 2) = FECHATEX$(CVINT(!FechRecep))
            GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 3) = SAFETEXT$(!NumRem_Prov)
            GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 4) = SAFETEXT$(!patentetransporte)
            GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 5) = SAFETEXT$(!NumProv)
            GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 6) = SAFETEXT$(!Raso_Prov)
            GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 7) = CODEXT$(!cod_inte)
            GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 8) = DESCRIT0$(!cod_inte)
            GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 9) = FORMATNUM$(!Peso, "F10.1")
            GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 10) = FORMATNUM$(!kgbalanzapublica, "F10.1")
            GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 11) = FORMATNUM$(SAFETEXT$(!CantNom), "F10.1")
            GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 12) = SAFETEXT$(!NUMEBORE)
            If CantRegistros("infocalidad", "documentoorigen='" & LoteRecepcion$ & "'") > 0 Then
                GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, 13) = "SI"
            End If
        
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
85  GRILLAGRAL.Caption = "Recepciones de Material a Controlar | Período: " & PERIO$
    
    GRILLAGRAL.mnuMenu1.Caption = "Exportar a Excel"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.mnuMenu2.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu2.Visible = True
    GRILLAGRAL.MnuEdición.Visible = False
    GRILLAGRAL.GRID.ZOrder 0
    
    Screen.MousePointer = 1
    GRILLAGRAL.Show 1
    If GRILLAGRAL.SeleccionaFilaDobleClick% = 1 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        NumeroRecepcion& = XVALO(RESP_ZELE_VECT(12))
        Recepcion_AdjuntarDocumentosExternos (NumeroRecepcion&)
        GoTo 85
    End If
    
    If GRILLAGRAL.MENU1 > 0 Then
        CommonDialog1.CancelError = True
        CommonDialog1.DialogTitle = "Exportar a Excel"
        CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|(*.XLSX)|*.XLSX|"
        CommonDialog1.FilterIndex = 1
        On Error GoTo 99
        CommonDialog1.ShowOpen
        On Error GoTo 0
        '
        DoEvents
        Screen.MousePointer = 11
        AREXPO$ = UCase$(CommonDialog1.FileName)
        If TRIM$(AREXPO$) <> "" Then
           Call GeneraExcelDeGrilla(AREXPO$, GRILLAGRAL.GRID, GRILLAGRAL.Caption, False)
        End If
        
    ElseIf GRILLAGRAL.MENU2 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "REMATCON", GRILLAGRAL.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If

99 Exit Sub
    
    
End Sub

Private Sub mnurechazosxfecha_Click()
   Call Command18_Click
End Sub

Private Sub mnuReimpEtiq_Click()
    Dim IncluyeQR As Boolean
    IncluyeQR = False
    If InStr(1, EMPREAC$, "CAMBACERES") > 0 Then
        IncluyeQR = True
        Call ReimpresionEtiquetasRecepcionCambaceres(0, IncluyeQR)
       Exit Sub
    End If
    If InStr(1, EMPREAC$, "PERGAMINO") > 0 Then
      Call Reimpre_Etiquetas
      Exit Sub
    End If

    If TRIM$(Control$("", "RECPAPEL")) = "SI" Or TRIM$(Control$("", "STIKAPRO")) <> "" Then
        If InStr(1, EMPREAC$, "CONTEPLAST") > 0 Or InStr(1, EMPREAC$, "RETEPLAST") > 0 Or InStr(EPAC$, "PERGAMINO") > 0 Then IncluyeQR = True
        Call ReimpresionEtiquetasParaBobinas(0, IncluyeQR)
    Else
        Call Reimpre_Etiquetas
    End If
End Sub

Sub ReimpresionEtiquetasRecepcionCambaceres(NroRecepcion As Long, Optional IncluyeQR As Boolean)

    OPX% = COMALTER%("Seleccione Tipo de Etiqueta\\Cancelar\Etiqueta Pequeña\Etiqueta Grande ")
    If OPX% <= 1 Then
       Exit Sub
    ElseIf OPX% = 2 Then
       FORIPRE$ = "ETR2-CTP"
    Else
       FORIPRE$ = TRIM$(Control$("", "STIKAPRO"))
    End If
    
    If EXISTE%(LUDAT$ + FORIPRE$ + ".SQR") = 0 Then
        Call COMUNI("Falta Formulario de Impresión de Etiqueta")
        Exit Sub
    End If

    EPAC$ = UCase$(TRIM$(EMPREAC$))
    SOYCAMBACERES% = 0
    If InStr(1, EPAC$, "CAMBACERES") > 0 Then SOYCAMBACERES% = 1
    SOYQUILMES% = 0
    If InStr(1, EPAC$, "QUILMES") > 0 Then SOYQUILMES% = 1: SOYCAMBACERES% = 0
    
    Call CREATABLA_ARCHIVOS_QR
    USUARIO% = USNBR% Mod 100
    
    PFAC$ = Control$(IDENTER$, "PORTSTIK")
    If PFAC$ = "" Then PFAC$ = Control$("*", "PORTSTIK")
    If PFAC$ = "" Then Exit Sub ' POR SI NOHAY NINGUNA IMPRESORA CONFIGURADA
    Call SETPRINPORT(PFAC$)
    
    Dim Rst As ADODB.Recordset
   
    ' *** IMPRIME TODAS LAS ETIQUETAS DE LA RECEPCION
    ' ***********************************************
    If NroRecepcion > 0 Then
        NumeroRecepcion$ = CStr(NroRecepcion)
        GoTo 65
    End If
    
    
    ' *** REIMPRIME UNA ETIQUETA
    ' **************************
    ' LISTA DE RECEPCIONES
40  TituloLista$ = "Recepciones de materiales"
    If TRIM$(FRMZELECHO.Caption) = TituloLista$ Then GoTo 50
    
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Recepcion/A12;Fecha/D8;Remito/A20;Proveedor/A32", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Recepcion/A12;Fecha/D8;Remito/A20;Proveedor/A32", FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim obj As New RECORXET
    Set rs1 = obj.RS_BOLRECEPEN()
    Call Carga_MSF_Recordset(rs1, "Recepcion/A12;Fecha/D8;Remito/A20;Proveedor/A32", FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 10000
50  FRMZELECHO.Show 1
    
    NumeroRecepcion$ = TRIM(RESP_ZELE_VECT(1))
    If NumeroRecepcion$ = "" Then Exit Sub
    
    ' LISTA DE ITEMS DE LA RECEPCION ELEGIDA
    TituloLista$ = "Items de la Recepción '" & NumeroRecepcion$ & "'"
    If TRIM$(FRMZELECHO.Caption) = TituloLista$ Then GoTo 50
    
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Lote/A12;Código/A16;Descripción/A32;Cantidad/F10.1;Peso/F10.2", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Lote/A12;Código/A16;Descripción/A32;Cantidad/F10.1;Peso/F10.2", FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim rst_items As ADODB.Recordset
    ConsultaSql$ = "select idenlote, cod_exte, descripitem, m_largo, peso, cantnom from bolrecep with(nolock) where numebore=" & NumeroRecepcion$
    Set rst_items = READSET(ConsultaSql$)
    
    Fila& = 0
    With rst_items
        Do While Not .EOF
            Fila& = Fila& + 1
            FRMZELECHO.msf2.Rows = Fila& + 1
            
            FRMZELECHO.msf2.TextMatrix(Fila&, 1) = !idenlote
            FRMZELECHO.msf2.TextMatrix(Fila&, 2) = !Cod_Exte
            FRMZELECHO.msf2.TextMatrix(Fila&, 3) = !descripitem
            cantidadRecibida# = XVALO(!M_Largo): If cantidadRecibida# < 0.05 Then cantidadRecibida# = XVALO(!CantNom)
            FRMZELECHO.msf2.TextMatrix(Fila&, 4) = TRIM$(FORMATNUM$(cantidadRecibida#, "F12.1"))
            FRMZELECHO.msf2.TextMatrix(Fila&, 5) = TRIM$(FORMATNUM$(!Peso, "F12.2"))
        
            .MoveNext
        Loop
    End With
            
    On Error Resume Next
    rst_items.Close
    Set rst_items = Nothing
    On Error GoTo 0
    
    If Fila& = 0 Then Exit Sub
    
    FRMZELECHO.Width = 10500
60  FRMZELECHO.Show 1
    
    LoteRecepcion$ = TRIM(RESP_ZELE_VECT(1))
    CodigoExt$ = TRIM(RESP_ZELE_VECT(2))
    CI1% = CODINT(CodigoExt$)
    
    If LoteRecepcion$ = "" Then GoTo 40
    If SOYQUILMES% > 0 Then
       OPX% = COMALTER%("Opciones de Impresión: \\Cancelar\Etiqueta con QR(Datos Varios)\Etiqueta con QR(Solo Código + Lote)")
       If OPX% < 1 Then Exit Sub
       'If OPX% = 2 Then 'Continua como siempre
       If OPX% = 3 Then
          Call PrintEtiquetaRecepcionBobinaCambaceres(LoteRecepcion$) 'QR(Solo Código + Lote)
       End If
    End If
    If SOYCAMBACERES% > 0 And TIMATE%(CI1%) = 41 Then
      Call PrintEtiquetaRecepcionBobinaCambaceres(LoteRecepcion$)
      Exit Sub
    End If

    ' OBTIENE DATOS DEL LOTE
    ConsultaSql$ = "select fechrecep, cod_inte, cod_exte, descripitem, idenlote, numrem_prov, m_largo, cantnom, peso, raso_prov, NumRem_Prov, CantidadBultos,LOTECOL_PROV from bolrecep with(nolock) where idenlote='" & LoteRecepcion$ & "'"
    Set Rst = READSET(ConsultaSql$)
    GoTo 70
        
    ' OBTIENE DATOS DE LA RECEPCION
65  ConsultaSql$ = "select fechrecep, cod_inte, cod_exte, descripitem, idenlote, numrem_prov, m_largo, cantnom, peso, raso_prov, NumRem_Prov, CantidadBultos,LOTECOL_PROV from bolrecep with(nolock) where numebore='" & CStr(NroRecepcion) & "'"
    Set Rst = READSET(ConsultaSql$)
    
70  With Rst
        Do While Not .EOF
            CodInte% = !cod_inte
            Codigo$ = CODEXT$(CodInte%)
            Descripcion$ = DESCRIT0$(CodInte%)
            LOTE$ = !idenlote
            Largo$ = "0": If XVALO(!M_Largo) > 0.0005 Then Largo$ = TRIM$(FORMATNUM$(!M_Largo, "F12.1"))
            Peso$ = "0": If XVALO(!Peso) > 0.005 Then Peso$ = TRIM$(FORMATNUM$(!Peso, "F12.1"))
            Proveedor$ = !Raso_Prov
            Remito$ = SAFETEXT$(!NumRem_Prov)
           
            cantidadNominal# = XVALO(!CantNom)
            FechaRecepcion% = CVINT(!FechRecep)
            cantidadBultos% = XVALO(!cantidadBultos)
10          CantidadModificada# = XVALO(InputBox("Cantidad", "Puede Modificar Las Cantidades", TRIM$(FORMATNUM$(cantidadNominal#, "F12.1"))))
            
            If CantidadModificada# < 0.005 Then
                   Call COMUNI("Atención Debe Ingresar Una Cantidad Mayor a 0")
                   GoTo 10
            End If
        
            PesoModificado# = XVALO(InputBox("KiloGramos", "Corrección de Peso", Peso$))

            CantidadAImprimir& = XVALO(InputBox("Cantidad de Etiquetas", "INGRESE CANTIDAD DE ETIQUETAS", "1"))
            If CantidadAImprimir& < 1 Then CantidadAImprimir& = 1
            
            TEXTOQR$ = Codigo$ & "|" & LOTE$ & "|" & Proveedor$ & "|" & TRIM$(FORMATNUM$(CantidadModificada#, "F12.1")) & "|" & TRIM$(FORMATNUM$(PesoModificado#, "F12.1")) & "|" & Remito$ & "|" & FECHATEX$(FechaRecepcion%)
            IMAG$ = "IMAGRECE_QR.JPG"
            IDENTIFICADOR$ = "QR_RECEPCION"
            ANCHOX% = 0
            ALTOX% = 0
            '
            'ETIQUETA PARA REALIZAR LA TRANSFERERNCIA EN MODULO DE STOCK POR ITEM UNICO A OBJETO 'TRANSFESTOCK'
            DEPOSITOSALIDA$ = Control$("", "DEPOCEN"): If TRIM$(DEPOSITOSALIDA$) = "" Then DEPOSITOSALIDA$ = "1"
            familia1$ = TRIM$(UCase$(CODFAMIL$(CodInte%)))
            REFERENCIA$ = "Lote Transf.: " & LOTE$
            If SOYCAMBACERES% > 0 Then
               CODIGOHASTA$ = CodigoTransformacionSegunFamilia_Camb$(familia1$)
            ElseIf SOYQUILMES% > 0 Then 'QUILMES ES CODIGOHASTA EL MISMO
               CODIGOHASTA$ = CODEXT$(CodInte%)
            End If
            DEPODESTINO$ = SAFETEXT$(DepositoDestinoTransformacionSegunFamilia_Camb%(familia1$))
            TEXTOQR2$ = Codigo$ & Chr$(13) & DEPOSITOSALIDA$ & Chr$(13) & Chr$(13) & CODIGOHASTA$ & Chr$(13) & DEPODESTINO$
            TEXTOQR2$ = TEXTOQR2$ + Chr$(13) & Chr$(13) & TRIM$(FORMATNUM$(PesoModificado#, "F12.1")) & Chr$(13) & REFERENCIA$
            IMAG2$ = "IMAGTRAN_QR.JPG"
            IDENTIFICADOR2$ = "QR_RECTRANF"
            ANCHOX2% = 0
            ALTOX2% = 0

            If IncluyeQR = True Then
                On Error Resume Next
                Kill LUCOM$ & IMAG$
                On Error GoTo 0
                'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
                QRX$ = TRIM$(IMAGENCODBARRA$(LUCOM$, IMAG$, TEXTOQR$, "QR_CODE", ANCHOX%, ALTOX%))
                Call T_Espera(3)
                Image1.Picture = LoadPicture(LUCOM$ & IMAG$)
                Call GRABAR_QR(LUCOM$ & IMAG$, IDENTIFICADOR$)
                
                'GRABA QR DE TRANSFERENCIA DE STOCKS
                On Error Resume Next
                Kill LUCOM$ & IMAG2$
                On Error GoTo 0
                'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
                QRX$ = TRIM$(IMAGENCODBARRA$(LUCOM$, IMAG2$, TEXTOQR2$, "QR_CODE", ANCHOX%, ALTOX%))
                Call T_Espera(3)
                Image1.Picture = LoadPicture(LUCOM$ & IMAG2$)
                Call GRABAR_QR(LUCOM$ & IMAG2$, IDENTIFICADOR2$)
            End If
            If SOYQUILMES% < 1 Then
               FILTX$ = LOTE$ & ";" & Codigo$ & ";" & Descripcion$ & ";" & Proveedor$ & ";" & "= " & TRIM$(FORMATNUM$(CantidadModificada#, "F12.1")) & "; = " & TRIM$(FORMATNUM$(PesoModificado#, "F12.1")) & "; = " & Remito$ & ";: " & FECHATEX$(FechaRecepcion%)
            ElseIf SOYQUILMES% > 0 Then 'EJEMPLO EN QUILMES BR-336-1
               FILTX$ = LOTE$ & ";" & Codigo$ & ";" & Descripcion$ & ";" & Proveedor$ & ";" & "= " & TRIM$(FORMATNUM$(Largo$, "F12.1")) & "; = " & TRIM$(FORMATNUM$(PesoModificado#, "F12.1")) & "; = " & Remito$ & ";: " & FECHATEX$(FechaRecepcion%)
            End If
            Call STDFORM(FORIPRE$, FILTX$, , , , CantidadAImprimir&)
            .MoveNext
        Loop
    End With
        
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
    If PFAC$ <> "" Then Call SETPRINPORT("RESTORE")

End Sub


Private Sub mnuReimpRecep_Click()
   Call Command12_Click
End Sub

Private Sub mnuReporteVarios_Click()
   Call AGenRep_Click
End Sub

Private Sub mnuSalir_Click()
   Call Command1_Click
End Sub


Private Sub mnutransferirasector_Click()
   'SI ESTA EN MEMORIA LO BORRA.
   If IsFormLoaded(BCODEGR09) = True Then
      Unload BCODEGR09
   End If
   BCODEGR09.DEVOLUCION% = 0
   BCODEGR09.Caption = "Transferencia de Materia Prima a Sector Productivo"
10 BCODEGR09.Show 1
   If BCODEGR09.VALI_BCODE09 = "APROB" Then
      Call TRANSFALEJPA(OK%)
   End If
   Screen.MousePointer = 1
End Sub

Sub TRANSFALEJPA(OK%)

      With BCODEGR09
        If .DEVOLUCION < 1 Then
          DEPOSAL% = XVALO(.Text2(.DEVOLUCION))
          DEPOING% = XVALO(.Label5(.DEVOLUCION))
        Else
          DEPOING% = XVALO(.Text2(.DEVOLUCION))
          DEPOSAL% = XVALO(.Label5(.DEVOLUCION))
        End If
        
        If DEPOSAL% < 1 Then Call COMUNI("Falta Seleccionar Depósito"): OK% = -1: Exit Sub
        If DEPOING% < 1 Then Call COMUNI("Falta Sector Depósito"): OK% = -1: Exit Sub

        CMOV$ = TRIM$(Control$("", "CODTRAST"))
        If CMOV$ = "" Then CMOV$ = "TR"
        NRO& = ProNroComprobante(CMOV$)
        FECHA% = HOY(HO$)
        NC% = 0
        VUNI# = 0
        MONE$ = ""
        CMOV$ = "TR"
        DOPRI$ = "TR." + TRIM$(Str$(NRO&))
        DOSEC$ = "TR." + Mid$(HO$, 7, 2) + Mid$(HO$, 4, 2) + Left$(HO$, 2) '& "-" & Right(SAFETEXT$(NRO&), 2)
               '

        For i& = 1 To .GRID.Rows - 1
            CI% = XVALO(.GRID.TextMatrix(i&, 1))
            CANTI = XVALO(.GRID.TextMatrix(i&, 5))
            LOTE$ = .GRID.TextMatrix(i&, 4)
            '
            'MOVIMIENTO DE SALIDA
            CANTI = -1 * CANTI
            REFE$ = "Transf. de Dep: " + TRIM$(Str$(DEPOSAL%))
            Call MOVISTOK(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, VUNI#, 1, NC%, DEPOSAL%, CANTI)
            '
            'MOVIMIENTO DE ENTRADA
            CANTI = -1 * CANTI ' LO VUELVO A PONER EN POSITIVO
            REFE$ = "Transf. a Dep: " + TRIM$(Str$(DEPOING%))
            Call MOVISTOK(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, VUNI#, 1, NC%, DEPOING%, CANTI)
         Next i&
         
         If .DEVOLUCION < 1 Then
            Call COMUNI("Transferencia Completada Exitosamente")
         Else
            Call COMUNI("Devolución Completada Exitosamente")
         End If
      End With
      
      

End Sub

Private Sub mnuVerificarDatosSegunSerie_Click()

10          Campos$ = "Número de Serie/A24;ULT/F0"
            VENTANA.Inicializar = 0
            VENTANA.AgregaRegistro = 0
            VENTANA.UtilizaFormula = 1
            VENTANA.NoMostrarBotonTilde = 0
            VENTANA.TITULO = "Ingrese N° de Serie"
            VENTANA.CampEditables = "1"
            VENTANA.Campox = Campos$
            VENTANA.Inicializar = 1
            VENTABNEW.MSF.TextMatrix(1, 1) = NUSER1$
            
20          VENTABNEW.Show 1
            If VENTABNEW.APROBADO% > 0 Then
               NUSER1$ = TRIM$(VENTABNEW.MSF.TextMatrix(1, 1))
               CI1% = CODINT_SEGUN_SERIE(NUSER1$)
               If CI1% > 0 Then
                  NORFA$ = OrdenFabricacionNumSerieTanumse$(CI1%, NUSER1$)
                  NUBORE& = XVALO(OrigenSegunOF$(NORFA$))
               Else
                  Call COMUNI("N.Serie Incorrecto (No Devuelve Código Interno)")
                  GoTo 20
                  Exit Sub
               End If
               Unload VENTABNEW
               Dim obj As New Art_Importados
               Call obj.READ_DATOS_IMPO(NUBORE&)
               Campos$ = "Número de Serie/A24;Fecha/D8;Despacho/A24;Aduana/A24;AULT/F0"
               VENTANA.Campox = Campos$
               VENTANA.Inicializar = 0
               VENTANA.CampEditables = "0;0;0;0"
               VENTANA.AgregaRegistro = 0
               VENTANA.UtilizaFormula = 1
               VENTANA.NoMostrarBotonTilde = 0
               VENTANA.TITULO = "Datos de N° de Serie: " & NUSER1$
               VENTANA.Inicializar = 1
               VENTABNEW.MSF.TextMatrix(1, 1) = NUSER1$
               VENTABNEW.MSF.TextMatrix(1, 2) = obj.FECHA
               VENTABNEW.MSF.TextMatrix(1, 3) = obj.Despacho
               VENTABNEW.MSF.TextMatrix(1, 4) = obj.Aduana
               VENTABNEW.Show 1
               Set obj = Nothing
               Unload VENTABNEW
               GoTo 10
           End If
End Sub

Private Sub reimpremidevo_Click()
   Call ReImprimirRD
End Sub

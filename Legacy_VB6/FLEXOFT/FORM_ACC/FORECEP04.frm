VERSION 5.00
Begin VB.Form FORECEP04 
   BackColor       =   &H00E0E0F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Materiales - Altas de Stock en Depósitos"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   9720
   Icon            =   "FORECEP04.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture15 
      Height          =   6200
      Left            =   8820
      ScaleHeight     =   6135
      ScaleWidth      =   1035
      TabIndex        =   59
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   6255
         Left            =   0
         Picture         =   "FORECEP04.frx":030A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture14 
      Height          =   6150
      Left            =   360
      ScaleHeight     =   6090
      ScaleWidth      =   8940
      TabIndex        =   58
      Top             =   5280
      Visible         =   0   'False
      Width           =   9000
      Begin VB.Image Image1 
         Height          =   6135
         Left            =   120
         Picture         =   "FORECEP04.frx":26A2
         Stretch         =   -1  'True
         Top             =   360
         Width           =   8775
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0F0&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5790
      Left            =   4450
      TabIndex        =   2
      Top             =   210
      Width           =   4110
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0F0&
         Caption         =   "Control de Recepciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2820
         Left            =   210
         TabIndex        =   23
         Top             =   2835
         Width           =   3690
         Begin VB.PictureBox Picture13 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2370
            TabIndex        =   55
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
               Height          =   600
               Left            =   0
               Picture         =   "FORECEP04.frx":109554
               Style           =   1  'Graphical
               TabIndex        =   56
               ToolTipText     =   "Recepciones Rechazadas por Rango de Fechas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Rechazos por Rango de Fechas"
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
               Left            =   720
               TabIndex        =   57
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture12 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2370
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
               Height          =   600
               Left            =   0
               Picture         =   "FORECEP04.frx":10985E
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
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   720
               TabIndex        =   44
               Top             =   20
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2370
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
               Height          =   600
               Left            =   0
               Picture         =   "FORECEP04.frx":1099A8
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
               TabIndex        =   29
               Top             =   20
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture9 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   300
            ScaleHeight     =   585
            ScaleWidth      =   2370
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
               Height          =   600
               Left            =   -20
               Picture         =   "FORECEP04.frx":109CB2
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   0
               Width           =   560
            End
            Begin VB.Label LABEL5 
               BackStyle       =   0  'Transparent
               Caption         =   "Recepciones por Rango de Fechas"
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
               TabIndex        =   26
               Top             =   20
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0F0&
         Caption         =   "Consulta de Pendientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Left            =   210
         TabIndex        =   13
         Top             =   315
         Width           =   3690
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   1050
            ScaleHeight     =   585
            ScaleWidth      =   2160
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
               Height          =   600
               Left            =   -20
               Picture         =   "FORECEP04.frx":109FBC
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
               TabIndex        =   22
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   735
            ScaleHeight     =   585
            ScaleWidth      =   2160
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
               Height          =   600
               Left            =   -20
               Picture         =   "FORECEP04.frx":10A2C6
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "Ordenes de Compra Pendientes de Recepción"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Pedidos y Orde- nes de Compra"
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
               TabIndex        =   19
               Top             =   50
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   420
            ScaleHeight     =   585
            ScaleWidth      =   2160
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
               Height          =   600
               Left            =   -20
               Picture         =   "FORECEP04.frx":10A5D0
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
               TabIndex        =   16
               Top             =   50
               Width           =   1590
            End
         End
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0F0&
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   650
      Left            =   210
      TabIndex        =   48
      Top             =   5355
      Width           =   4110
      Begin VB.OptionButton Option1 
         BackColor       =   &H00E0E0F0&
         Caption         =   "Pantalla"
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
         Left            =   630
         TabIndex        =   50
         Top             =   320
         Width           =   1380
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00E0E0F0&
         Caption         =   "Impresión"
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
         Left            =   2310
         TabIndex        =   49
         Top             =   320
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
         Picture         =   "FORECEP04.frx":10A8DA
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
      BackColor       =   &H00000080&
      ForeColor       =   &H00000000&
      Height          =   6210
      Left            =   8600
      ScaleHeight     =   6150
      ScaleWidth      =   1005
      TabIndex        =   30
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   52
         Top             =   5250
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   53
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Tr"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FORECEP04.frx":10AA24
         TabIndex        =   51
         ToolTipText     =   "Sub-Sistema de Trazabilidad"
         Top             =   4450
         Visible         =   0   'False
         Width           =   650
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "FORECEP04.frx":10AD2E
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Ver y Re-Imprimir B.Recepción"
         Top             =   1890
         Width           =   650
      End
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
         Height          =   640
         Left            =   105
         Picture         =   "FORECEP04.frx":10B038
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Acceso Directo a otras Aplicaciones"
         Top             =   3180
         Width           =   650
      End
      Begin VB.CommandButton Command29 
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
         Left            =   105
         Picture         =   "FORECEP04.frx":10B342
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   955
         Width           =   650
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
         Height          =   640
         Left            =   105
         Picture         =   "FORECEP04.frx":10B64C
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Terminar - Salir de esta Aplicación"
         Top             =   315
         Width           =   650
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
         Height          =   640
         Left            =   105
         Picture         =   "FORECEP04.frx":10B796
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Configuración de Comportamiento y Formularios"
         Top             =   2535
         Width           =   650
      End
      Begin VB.CommandButton Command10 
         Caption         =   "QC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   640
         Left            =   105
         Picture         =   "FORECEP04.frx":10BBD8
         TabIndex        =   31
         ToolTipText     =   "Calificación de Proveedores"
         Top             =   3810
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -240
         Picture         =   "FORECEP04.frx":10BEE2
         Stretch         =   -1  'True
         Top             =   5565
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0F0&
      Caption         =   "Ingresos a Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4110
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0F0&
         Caption         =   "Gestión de Compras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   210
         TabIndex        =   39
         Top             =   2200
         Width           =   3690
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2370
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
               Height          =   600
               Left            =   0
               TabIndex        =   41
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Consulta de Or- denes de Compra"
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
               TabIndex        =   42
               Top             =   105
               Width           =   1695
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0F0&
         Caption         =   "Entradas de Fábrica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   210
         TabIndex        =   9
         Top             =   3630
         Width           =   3690
         Begin VB.PictureBox Picture6 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2370
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
               Height          =   600
               Left            =   0
               Picture         =   "FORECEP04.frx":10DE04
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label7 
               BackStyle       =   0  'Transparent
               Caption         =   "Por Propia Produccion"
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
               TabIndex        =   12
               Top             =   50
               Width           =   1005
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0F0&
         Caption         =   "Recepciones Externas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1730
         Left            =   210
         TabIndex        =   1
         Top             =   315
         Width           =   3690
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   840
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   6
            Top             =   900
            Width           =   2325
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
               Height          =   600
               Left            =   0
               Picture         =   "FORECEP04.frx":10E10E
               Style           =   1  'Graphical
               TabIndex        =   7
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Inspección y Aprobación"
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
               TabIndex        =   8
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2265
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
               Height          =   600
               Left            =   0
               Picture         =   "FORECEP04.frx":10E418
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Boleta de Recep- ción (Provisoria)"
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
               Left            =   630
               TabIndex        =   5
               Top             =   0
               Width           =   1590
            End
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
            TabIndex        =   54
            Top             =   420
            Visible         =   0   'False
            Width           =   645
         End
      End
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
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu recepcion 
      Caption         =   "Recepción"
      Begin VB.Menu recepcionprovisoria 
         Caption         =   "Boleta de Recepción Provisoria"
      End
      Begin VB.Menu inspeccionyaprobacion 
         Caption         =   "Inspección y Aprobación"
      End
      Begin VB.Menu entradaxfabricacion 
         Caption         =   "Entradas x Fabricación"
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu anularecepcionprovisoria 
         Caption         =   "Anulación de Boleta de Recepción Provisoria"
      End
      Begin VB.Menu s4 
         Caption         =   "-"
      End
      Begin VB.Menu seleccionpararechazo 
         Caption         =   "Selección de Lote Para Rechazo"
      End
      Begin VB.Menu s6 
         Caption         =   "-"
      End
      Begin VB.Menu reimprimirremitodevolucion 
         Caption         =   "Re-Impresión de Remito de Devolución"
      End
   End
   Begin VB.Menu ver 
      Caption         =   "Ver"
      Begin VB.Menu verboletarecepcion 
         Caption         =   "Visualizar Boletas de Rececpcion"
      End
   End
   Begin VB.Menu consultasylistados 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu detallexboletarecepcion 
         Caption         =   "Detalle por Boleta de Recepción"
         Begin VB.Menu pantalla 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu impresora 
            Caption         =   "Impresora"
         End
      End
      Begin VB.Menu pendientesderecepcion 
         Caption         =   "Ordenes de Compra Pendientes de Recepción"
      End
      Begin VB.Menu listadodematerialenconsignacion 
         Caption         =   "Material en Consignación"
         Begin VB.Menu pantalla2 
            Caption         =   "Pantalla"
         End
         Begin VB.Menu impresora2 
            Caption         =   "Impresora"
         End
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu recepcionesporrangdefecha 
         Caption         =   "Recepciones por Rango de Fecha"
      End
      Begin VB.Menu RecepcionesconRechazos 
         Caption         =   "Recepciones con Rechazos"
      End
      Begin VB.Menu recepcionessinordendecompra 
         Caption         =   "Recepciones sin Orden de Compras"
      End
      Begin VB.Menu rechazoporrangodefechas 
         Caption         =   "Rechazos por Rango de Fechas"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu consultadeordenesdecompra 
         Caption         =   "Consulta de Ordenes de Compra"
      End
   End
   Begin VB.Menu utiles 
      Caption         =   "Útiles"
      Begin VB.Menu calificacionproveedores 
         Caption         =   "Calificación de Proveedores"
      End
      Begin VB.Menu Trazabilidad 
         Caption         =   "Trazabilidad"
      End
   End
   Begin VB.Menu configuracion2 
      Caption         =   "Configuración"
      Begin VB.Menu configuracion 
         Caption         =   "Configuración"
      End
   End
End
Attribute VB_Name = "FORECEP04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HAYMATERIALIB%
Dim MODORECHA% '\\\OT-05-0723-WAR-23/11/05///
Dim CANCELX1% '\\\OT-06-0178-WAR-12/05/06///

Sub REIMPREDEV()
    '
    If TRIM$(CONTROL$("", "REMIRECH")) <> "" Then ' TIPO DE MOVIMIENTO
      CODRECHA$ = TRIM$(CONTROL$("", "CODRRECH"))
      If CODRECHA$ = "" Then CODRECHA$ = "RD"
    End If
    '
    SQLX$ = " SELECT DISTINCT DoPriCor, Nume_Prov, FechMov FROM MOVISTO "
    'PARA EVITAR COMPROBANTES ARRUINADOS
    SQLX$ = SQLX$ + " WHERE LEFT$(UCASE$(REFERCOR),4) <> 'ANUL' "
    SQLX$ = SQLX$ + " AND CodiMovi =  '" & CODRECHA$ & "'  "
    SQLX$ = SQLX$ + " ORDER BY FECHMOV"
    '
    Call APERBASDAT("STOCKS")
    Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
         Call MENSERR(24, " No Existen Registros que Cumplan\con la Condición SQL.")
         On Error GoTo 0
         Exit Sub
      End If
      On Error GoTo 0
       '
      J& = 0
      Call COPYSTRU("SELREMI", "REMIDEVO")
      Do While (.EOF = False)
         PRO% = !Nume_Prov
         REMI$ = !DoPriCor
         FF% = CVINT(!fechMov)
         F1$ = FECHATEX$(FF%)
         REFE1$ = F1$ + " - " + RASOCLI$(-1 * PRO%)
         z$ = Space$(40) + String$(8, 0)
         Call REPLA(z$, REMI$, 1, 10)
         Call REPLA(z$, REFE1$, 11, 30)
         Call REPLA(z$, MKS$(U&), 41, 4)
         Call REPLA(z$, MKI$(FF%), 45, 2)
         Call REPLA(z$, MKI$(PRO%), 47, 2)
         J& = J& + 1
         Call GRAREG("!REMIDEVO", z$, J&)
       .MoveNext
      Loop
    End With
     '
    Call SETULTREG("!REMIDEVO", J&)
    Call CIERRARCH("!REMIDEVO")
    Screen.MousePointer = 1
    '
    '
15  RERE$ = REGSEL$("!REMIDEVO")
    If Len(RERE$) <= 4 Then Exit Sub
    
     NPX& = XVALO(TRIM$(Mid$(RERE$, 4, 5)))
     TIDOCUM$ = "Remito de Devolución"
     FECHA1$ = Mid$(RERE$, 11, 8)
     NUDOCU$ = TRIM$(Str$(NPX&))
     While Len(NUDOCU$) < 4: NUDOCU$ = "0" + NUDOCU$: Wend
     NUDOCU$ = NUDOCU$ + " "
     Screen.MousePointer = 11
     '
     For UI& = ULTREG&("PDOCLST") To 1 Step -1
       XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
        If InStr(XX$, TIDOCUM$) > 4 Then
           If InStr(XX$, NUDOCU$) > 4 Then
                PPXX% = InStr(XX$, "RD.")
                If PPXX% > 0 Then
                 XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
                End If
                If InStr(RERE$, NUDOCU$) > 3 Then
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
    

End Sub

Private Sub anularecepcionprovisoria_Click()
   Call ANUBORE
End Sub

Private Sub calificacionproveedores_Click()
   Call Command10_Click
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command10_Click()
Command10.Enabled = False
Call CONECRUN("CALPRO02", "Sistema Calificación de Proveedores")
Command10.Enabled = True
End Sub

Private Sub Command11_Click()
    '
    If Option1.Value = True Then
       Call MENSERR(24, "Opcion no Disponible por Pantalla.\  \Emita Listado Impreso y Exporte a Archivo de Texto.")
       GoTo 20
    End If
    '
    RTA% = COMALTER("Opciones de Listado: \\Mat.Codificado\Texto Libre\Unificado")
    If RTA% = 1 Then
        GoTo 20
      ElseIf RTA% = 2 Then
        Call PorTextoLib
      ElseIf RTA% = 3 Then
        Call UNIFICARECEP
    End If
    Exit Sub
    '
20  HOI% = HOY(HO$)
    FE$ = "01" + Mid$(HO$, 3)
    FF1% = FECHANUM(FE$): FF2% = HOI%
    XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
    XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
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
    RANGO$ = TRIM$(COD1$) + " - " + TRIM$(COD2$)
    RANGO$ = RANGO$ + " / " + TRIM$(TIMAT$)
    RANGO$ = RANGO$ + " / " + FLI1$ + " - " + FLI2$
    '
    Screen.MousePointer = 11
    '
    Call COPYSTRU("BOLRECEP", "BOLSELEC")
    J& = 0
    '
    Call ABRESTOCKS
    FF1D = CDbl(CVDAT(FF1%))
    FF2H = CDbl(CVDAT(FF2%))
    '
    SQLX$ = " SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + " WHERE cdbl(FechRecep) >= " & FF1D & " "
    SQLX$ = SQLX$ + " AND cdbl(FechRecep) <= " & FF2H & " "
    'SQLX$ = SQLX$ + " AND Cod_Exte >= '" & cod1$ & "' "
    'SQLX$ = SQLX$ + " AND Cod_Exte <= '" & cod2$ & "' "
    'SQLX$ = SQLX$ + " AND CODEXT$(Cod_Inte) >= '" & cod1$ & "' "
    'SQLX$ = SQLX$ + " AND CODEXT$(Cod_Inte) <= '" & cod2$ & "' "
    SQLX$ = SQLX$ + " ORDER BY NumeBoRe "
    '
    Set BRECEPRATMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With BRECEPRATMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Boletas de Recepcion\que Cumplan Condición de Filtro.")
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
        CIXI% = !Cod_Inte
        CODD$ = CODEXT$(CIXI%)
        If CODD$ >= COD1$ Then
          If CODD$ <= COD2$ Then
            If TMT% = 0 Or TIMATE%(CIXI%) = TMT% Then
                Y$ = REGBLAN("BOLSELEC")
                Call REPLA(Y$, MKS$(!NUMEBORE), 1, 4)
                Call REPLA(Y$, !Descrip, 5, 64)
                Call REPLA(Y$, !IdenLote, 69, 12)
                Call REPLA(Y$, MKI$(CVINT(!FechRecep)), 81, 2)
                Call REPLA(Y$, MKI$(!Cod_Inte), 83, 2)
                Call REPLA(Y$, !Unidad, 85, 4)
                Call REPLA(Y$, MKD$(!CantNom), 89, 8)
                Call REPLA(Y$, Chr$(!DepoENT), 97, 1)
                Call REPLA(Y$, MKI$(!NUMPROV), 105, 2)
                Call REPLA(Y$, !LoteCol_Prov, 107, 12)
                Call REPLA(Y$, !NUMREM_PROV, 119, 16)
                  NCX% = !NUMPROV
                Call REPLA(Y$, CODICLI$((-1) * NCX%), 135, 16)
                Call REPLA(Y$, MKS$(!NumOC), 151, 4)
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
          End If
        End If
      .MoveNext
      Loop
    End With
    Call SETULTREG("BOLSELEC", JJ&)
    Call CIERRARCH("BOLSELEC")
    Call CIERRARCH("*.*")
    
    
    
    
'    FIN& = ULTREG&("BOLRECEP")
'    For U& = 1 To FIN&
'       Call TRACE(20, U&, FIN&)
'       X$ = REGLEIDO$("BOLRECEP", U&)
'       CIIX% = CVI(Mid$(X$, 83, 2))
'       FERE% = CVI(Mid$(X$, 81, 2))
'       '
'       If FERE% >= FF1% Then
'         If FERE% <= FF2% Then
'           If CVD(Mid$(X$, 89, 8)) > 0.0005 Then ' CANTIDAD NO NULA
'             If TMT% = 0 Or TIMATE%(CIIX%) = TMT% Then
'               CODE$ = CODEXT$(CIIX%)
'               If CODE$ >= cod1$ Then
'                 If CODE$ <= cod2$ Then
'                   J& = J& + 1
'                   Call GRAREG("BOLSELEC", X$, J&)
'                 End If
'               End If
'             End If
'           End If
'         End If
'       End If
'       '
'    Next U&
'    '
'    Call SETULTREG("BOLSELEC", J&)
'    Call CIERRARCH("BOLSELEC")
'    Call CIERRARCH("*.*")
'    '
    Call HEADERS("BOLSELEC", "")
    Call HEADERS("BOLSELEC", "Rango: " + RANGO$)
    Call CONECRUN("*BOLSELEC", "Recepciones por Rango de Fechas")
    '
99 Screen.MousePointer = 1
   '
End Sub

Sub PorTextoLib()
      
      '//OT-07-0466 //
      
10    Call DESHASFECHA(DES%, HAS%, PERIO$)
      '\\\OT-08-0287-FL-22/05/08///
      Call SELECHO("PROVED1")
      NCP% = XVALO(VALACT1$("PROVED1"))
      '\\\OT-08-0287-FL-FIN///
      '
      DES1# = CDbl(CVDAT(DES%))
      HAS1# = CDbl(CVDAT(HAS%))
      SQLX$ = "SELECT * FROM BOLRECEP "
      SQLX$ = SQLX$ + " WHERE Cod_Inte =" & 0 & " "
      '
      '\\\OT-08-0287-FL-22/05/08///
      If NCP% <> 0 Then
        SQLX$ = SQLX$ + "AND NumProv =" & NCP% & " "
        PROVEED$ = RASOCLI$(-1 * NCP%)
      End If
      '\\\OT-08-0287-FL-FIN///
      '
      SQLX$ = SQLX$ + " AND INT(CDBL(FechRecep)) >= " & DES1# & " "
      SQLX$ = SQLX$ + " AND INT(CDBL(FechRecep)) <= " & HAS1# & " "
      SQLX$ = SQLX$ + " ORDER BY NumeBoRe ASC "
      
      Call ABRESTOCKS
      Set RECEPTXL = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
      With RECEPTXL
        On Error Resume Next
        .MoveFirst
        If Err Then
          On Error GoTo 0
          Call COMUNI("No Existen Materiales Recibidos de Texto Libre \En Rango de Fechas Elegido")
          Screen.MousePointer = 1
          Exit Sub
        Else
          JJ& = 0
          Do While Not .EOF
            Y$ = REGBLAN("BOLSELEC")
            Call REPLA(Y$, MKS$(!NUMEBORE), 1, 4)
            Call REPLA(Y$, !Descrip, 5, 64)
            'Call REPLA(Y$, !IdenLote, 69, 12)
            Call REPLA(Y$, MKI$(CVINT(!FechRecep)), 81, 2)
            'Call REPLA(Y$, MKI$(!Cod_Inte), 83, 2)
            Call REPLA(Y$, !Unidad, 85, 4)
            Call REPLA(Y$, MKD$(!CantNom), 89, 8)
            'Call REPLA(Y$, Chr$(!DepoEnt), 97, 1)
            Call REPLA(Y$, MKI$(!NUMPROV), 105, 2)
            Call REPLA(Y$, !LoteCol_Prov, 107, 12)
            Call REPLA(Y$, !NUMREM_PROV, 119, 16)
            Call REPLA(Y$, !NumFac_Prov, 135, 16)
            Call REPLA(Y$, MKS$(!NumOC), 151, 4)
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
          .MoveNext
          Loop
        End If
      End With
      '
      Call SETULTREG("BOLSELEC", JJ&)
      '
      Call HEADERS("BOLSELEC", "")
      Call HEADERS("BOLSELEC", "Rango: " + PERIO$)
      If NCP% <> 0 Then
        Call HEADERS("BOLSELEC", "Proveedor: " + PROVEED$)
      End If
      Call CONECRUN("*BOLSELIB", "Recepciones Libres por Rango de Fechas ")
      '
      Screen.MousePointer = 1
      '//OT-07-0466 //
End Sub

Sub UNIFICARECEP()
    '
    Screen.MousePointer = 11
    '
    Call DESHASFECHA(FF1%, FF2%, PERIO$)
    '
    'Call COPYSTRU("BOLRECEP", "BOLSELE2")
    
    Call ABRESTOCKS
    DES = CDbl(CVDAT(FF1%))
    HAS = CDbl(CVDAT(FF2%))
    '
    SQLX$ = " SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + " WHERE FechRecep >= " & DES & " "
    SQLX$ = SQLX$ + " AND FechRecep <= " & HAS & " "
    SQLX$ = SQLX$ + " ORDER BY NumeBoRe "
    '
    Set BRECEPRATMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With BRECEPRATMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Boletas de Recepcion\dentro de Rango de Fechas Elegido.")
        On Error GoTo 0
        GoTo 99
      End If
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
        CIXI% = !Cod_Inte
        CODD$ = CODEXT$(CIXI%)
            If TMT% = 0 Or TIMATE%(CIXI%) = TMT% Then
                Y$ = REGBLAN("BOLSELEC")
                Call REPLA(Y$, MKS$(XVALO(!NUMEBORE)), 1, 4)
                Call REPLA(Y$, SAFETEXT(!Descrip), 5, 64)
                Call REPLA(Y$, SAFETEXT(!IdenLote), 69, 12)
                Call REPLA(Y$, MKI$(CVINT(!FechRecep)), 81, 2)
                Call REPLA(Y$, MKI$(XVALO(!Cod_Inte)), 83, 2)
                Call REPLA(Y$, SAFETEXT(!Unidad), 85, 4)
                Call REPLA(Y$, MKD$(XVALO(!CantNom)), 89, 8)
                Call REPLA(Y$, Chr$(!DepoENT), 97, 1)
                Call REPLA(Y$, MKI$(XVALO(!NUMPROV)), 105, 2)
                Call REPLA(Y$, SAFETEXT(!LoteCol_Prov), 107, 12)
                Call REPLA(Y$, SAFETEXT(!NUMREM_PROV), 119, 16)
                  NCX% = !NUMPROV
                Call REPLA(Y$, CODICLI$((-1) * NCX%), 135, 16)
                Call REPLA(Y$, MKS$(XVALO(!NumOC)), 151, 4)
                Call REPLA(Y$, SAFETEXT(!MoneVaUni), 155, 4)
                Call REPLA(Y$, MKS$(XVALO(!Valo_Unit)), 159, 4)
                Call REPLA(Y$, MKD$(XVALO(!CanToApro)), 163, 8)
                Call REPLA(Y$, MKD$(XVALO(!CanToRech)), 171, 8)
                Call REPLA(Y$, MKD$(XVALO(!CanToFalta)), 179, 8)
                Call REPLA(Y$, MKI$(CVINT(!FechLibe)), 187, 2)
                If IsNull(!Observa) Then
                    USERAPRO$ = " "
                Else
                    USERAPRO$ = SAFETEXT(!Observa)
                End If
                Call REPLA(Y$, USERAPRO$, 193, 48) 'antes estaba el campo observaciones
                '
                JJ& = JJ& + 1
                Call GRAREG("BOLSELEC", Y$, JJ&)
            End If

      .MoveNext
      Loop
    End With
    Call SETULTREG("BOLSELEC", JJ&)
    Call CIERRARCH("BOLSELEC")
    Call CIERRARCH("*.*")
    Call HEADERS("BOLSELEC", "")
    Call HEADERS("BOLSELEC", "Rango : " & PERIO$)
    Call CONECRUN("*BOLSELE2", "Recepciones por Rango de Fechas")
    '
99  Screen.MousePointer = 1

End Sub


Private Sub Command12_Click()
   Command12.Enabled = False
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
     'agrego todas las O/C de ocomenca
      Call ABRECOMPRAS
      JJ& = 0
      SQLX$ = " SELECT Nume_Ocom, Refe_Ocom FROM OCOMENCA"
      SQLX$ = SQLX$ + " ORDER BY Nume_Ocom  DESC "
      Set VEROCTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
      With VEROCTEMP
         On Error Resume Next ' si hay algun error msgbox err
         .MoveFirst
         If Err Then
           On Error GoTo 0
           GoTo 10
         End If
         On Error GoTo 0
         '
         Do While (.EOF = False)
             XXY$ = REGBLAN$("INDIOCOM")
             Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4)
             Call REPLA(XXY$, !Refe_OcoM, 5, 44)
             JJ& = JJ& + 1
             Call GRAREG("INDIOCOM", XXY$, JJ&)
             .MoveNext
         Loop
      End With
      Set VEROCTEMP = Nothing
      '
      ' LO QUE SIGUE PARA ORDENES DE COMPRA ABIERTAS
10    SQLX$ = " SELECT Nume_Ocom, Refe_Ocom FROM ENCOCABI"
      SQLX$ = SQLX$ + " ORDER BY Nume_Ocom  DESC "
      Set VEROCTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
      With VEROCTEMP
         On Error Resume Next ' si hay algun error msgbox err
         .MoveFirst
         If Err Then
           On Error GoTo 0
           GoTo 20
         End If
         On Error GoTo 0
         '
         Do While (.EOF = False)
             XXY$ = REGBLAN$("INDIOCOM")
             Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4)
             Call REPLA(XXY$, "(abierta) " + !Refe_OcoM, 5, 44)
             JJ& = JJ& + 1
             Call GRAREG("INDIOCOM", XXY$, JJ&)
             .MoveNext
         Loop
      End With
      Set VEROCTEMP = Nothing
      '
20    Call SETULTREG("INDIOCOM", JJ&)
      Call CIERRARCH("INDIOCOM")
      If JJ& < 1 Then
         Call MENSERR(24, "No hay Ordenes de Compra Emitidas")
         GoTo 99
      End If
      '
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
     OPRECEP04.Show 1
   End If
   '
End Sub

Private Sub Command15_Click()
    '
    HAS% = HOY(HOS$)
    FEX = HAS%
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
10  OKKX% = 0
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
       Exit Sub
    End If
    OKKX% = 1
    DES% = CVI(Left$(OBX$, 2)): DES1# = CDbl(CVDAT(DES%))
    HAS% = CVI(Mid$(OBX$, 3, 2)): HAS1# = CDbl(CVDAT(HAS%))
    If HAS% < DES% Then GoTo 10
    FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    '
    SQLX$ = "SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + " WHERE (isnull(NumOC)=true or NumOC < 1) "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechRecep)) >= " & DES1# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechRecep)) <= " & HAS1# & " "
    SQLX$ = SQLX$ + " ORDER BY NumeBoRe ASC "
    
    Call ABRESTOCKS
    Set RECEPTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    With RECEPTMP
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
        Call REPLA(Y$, !IdenLote, 69, 12)
        Call REPLA(Y$, MKI$(CVINT(!FechRecep)), 81, 2)
        Call REPLA(Y$, MKI$(!Cod_Inte), 83, 2)
        Call REPLA(Y$, !Unidad, 85, 4)
        Call REPLA(Y$, MKD$(!CantNom), 89, 8)
        Call REPLA(Y$, Chr$(!DepoENT), 97, 1)
        Call REPLA(Y$, MKI$(!NUMPROV), 105, 2)
        Call REPLA(Y$, !LoteCol_Prov, 107, 12)
        Call REPLA(Y$, !NUMREM_PROV, 119, 16)
        Call REPLA(Y$, !NumFac_Prov, 135, 16)
        Call REPLA(Y$, MKS$(!NumOC), 151, 4)
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
99 End Sub

Private Sub Command16_Click()
    '
    HAS% = HOY(HOS$)
    FEX = HAS%
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
10  OKKX% = 0
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
       Exit Sub
    End If
    OKKX% = 1
    DES% = CVI(Left$(OBX$, 2)): DES1# = CDbl(CVDAT(DES%))
    HAS% = CVI(Mid$(OBX$, 3, 2)): HAS1# = CDbl(CVDAT(HAS%))
    If HAS% < DES% Then GoTo 10
    FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    '
    SQLX$ = "SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + " WHERE CanToRech > 0 "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechRecep)) >= " & DES1# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechRecep)) <= " & HAS1# & " "
    SQLX$ = SQLX$ + " ORDER BY NumeBore ASC "
    Call ABRESTOCKS
    Set RECEPTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    With RECEPTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Call COMUNI("No Existen Recepciones con Rechazos\en Rango de Fechas Elegido.")
        Exit Sub
      End If
      JJ& = 0
      Call COPYSTRU("BOLSELEC", "BOLSELER")
      Do While Not .EOF
        Y$ = REGBLAN("BOLSELER")
        Call REPLA(Y$, MKS$(!NUMEBORE), 1, 4)
        Call REPLA(Y$, !Descrip, 5, 64)
        Call REPLA(Y$, !IdenLote, 69, 12)
        Call REPLA(Y$, MKI$(CVINT(!FechRecep)), 81, 2)
        Call REPLA(Y$, MKI$(!Cod_Inte), 83, 2)
        Call REPLA(Y$, !Unidad, 85, 4)
        Call REPLA(Y$, MKD$(!CantNom), 89, 8)
        Call REPLA(Y$, Chr$(!DepoENT), 97, 1)
        Call REPLA(Y$, MKI$(!NUMPROV), 105, 2)
        Call REPLA(Y$, !LoteCol_Prov, 107, 12)
        Call REPLA(Y$, !NUMREM_PROV, 119, 16)
        Call REPLA(Y$, !NumFac_Prov, 135, 16)
        Call REPLA(Y$, MKS$(!NumOC), 151, 4)
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
End Sub

Private Sub Command17_Click()
    Command17.Enabled = False
    If DERACCE%("TRAZABR", "Registro de Trazabilidad de B/R's") = 2 Then
      Call CONECRUN("TRAZAB04", "Sub-Sistema de Trazabilidad")
    End If
    Command17.Enabled = True
End Sub

Private Sub Command18_Click()
    '
    Command18.Enabled = False
    HAS% = HOY(HOS$)
    FEX = HAS%
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
10  OKKX% = 0
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
       GoTo 99 'Exit Sub X EL SCREENMOUSEPOINTER
    End If
    OKKX% = 1
    DES% = CVI(Left$(OBX$, 2)): DES1# = CDbl(CVDAT(DES%))
    HAS% = CVI(Mid$(OBX$, 3, 2)): HAS1# = CDbl(CVDAT(HAS%))
    If HAS% < DES% Then GoTo 10
    FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    '
    SQLX$ = "SELECT * FROM BOLRECEP "
    SQLX$ = SQLX$ + " WHERE CanToRech > 0 "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechRecep)) >= " & DES1# & " "
    SQLX$ = SQLX$ + " AND INT(CDBL(FechRecep)) <= " & HAS1# & " "
    SQLX$ = SQLX$ + " ORDER BY NumeBore ASC "
    Call ABRESTOCKS
    Set RECEPTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
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
        Call REPLA(Y$, !IdenLote, 69, 12)
        Call REPLA(Y$, MKI$(CVINT(!FechRecep)), 81, 2)
        Call REPLA(Y$, MKI$(!Cod_Inte), 83, 2)
        Call REPLA(Y$, !Unidad, 85, 4)
        Call REPLA(Y$, MKD$(!CantNom), 89, 8)
        Call REPLA(Y$, Chr$(!DepoENT), 97, 1)
        Call REPLA(Y$, MKI$(!NUMPROV), 105, 2)
        Call REPLA(Y$, !LoteCol_Prov, 107, 12)
        Call REPLA(Y$, !NUMREM_PROV, 119, 16)
        Call REPLA(Y$, !NumFac_Prov, 135, 16)
        Call REPLA(Y$, MKS$(!NumOC), 151, 4)
        Call REPLA(Y$, MKD$(!CanToRech), 171, 8)
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
    Call FILESORT("BOLSELER", "", 5, 5, "ASC")
    Call CONECRUN("*BOLSELEX", "Rechazos por Rango de Fechas")
    '
99  Screen.MousePointer = 1
    Command18.Enabled = True
End Sub

Private Sub Command2_Click()
'   Call APROBOL
'   Exit Sub
'\\\OT-05-0608-WAR-06/09/05///
RTA% = 1   ' COMALTER%("Modalidad de Procesamiento:\\Aprobación/Rechazo\Discrepancia/Faltantes")
Select Case RTA%
  Case 1
    Call APROBOL
  Case 2
    Call DISCREBOL
End Select
'\\\OT-05-0608-WAR-FIN///
End Sub

Private Sub Command3_Click()
    If Option1.Value = True Then
        Call COBORECEP
        Call FINAL("*COBOLPE")
      ElseIf Option2.Value = True Then
        Call COBORECEP
        If HAYMATERIALIB% = 0 Then
            Call FINAL("*LIBOLPE")
        Else
            Call FINAL("*LIBOPEL")
        End If
    End If
    '
    Screen.MousePointer = 1 '\\\OT-06-0179-WAR-12/05/06///
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
    Set consigtmp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With consigtmp
      '
      JJ& = 0
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        Do While Not .EOF
          ZZ$ = REGBLAN$("MACONSIG")
          Call REPLA(ZZ$, MKI$(CVINT(!Femi_Ocom)), 1, 2)
            NC% = !Npro_Ocom
          Call REPLA(ZZ$, MKI$(NC%), 3, 2)
            NORFA$ = !Nro_Orfa
          Call REPLA(ZZ$, NOCOM$, 9, 12)
            NOCOM$ = !Nro_Ocom
          Call REPLA(ZZ$, NOCOM$, 21, 12)
            CI% = !Cod_Inte
          Call REPLA(ZZ$, MKI$(CI%), 33, 2)
          Call REPLA(ZZ$, UNIMED$(CI%), 51, 4)
            CAN# = !Cant_Asig
          Call REPLA(ZZ$, MKD$(CAN#), 55, 8)
            CAN# = !Cant_Entreg
          Call REPLA(ZZ$, MKD$(CAN#), 63, 8)
            FF% = CVINT%(!Fecha_Entreg)
          Call REPLA(ZZ$, MKI$(FF%), 71, 2)
            CAN# = !Cant_Rendid
          Call REPLA(ZZ$, MKD$(CAN#), 111, 8)
            CAN# = !Saldo_Pend
          Call REPLA(ZZ$, MKD$(CAN#), 119, 8)
            '
          JJ& = JJ& + 1
          Call GRAREG("MACONSIG", ZZ$, JJ&)
        .MoveNext
        Loop
      End If
    End With
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
10  OPX% = ALTERNA%("Anulación de B.R.'s Pendientes\Devolución de Materiales\Re-Impresion de Remitos de Devolución\Otros Accesos Directos")
    If OPX% = 1 Then
        EPAC$ = TRIM$(UCase$(EMPREAC$))
        If InStr(EPAC$, "ASEMBLI") > 0 Then
          If HOY(HOS$) <= FECHANUM("01/08/08") Then
            If COMALTER%("Desea Eliminar BR's con Guiones ?\\No, Continuar\Si, Eliminarlos") = 2 Then
              Call BORRAGUION
              GoTo 99
            End If
          End If
        End If
        Call ANUBORE
      ElseIf OPX% = 2 Then
        Call DEVOLMATE ' SELECCION DE LOTE PARA RECHAZO
      ElseIf OPX% = 3 Then
        Call REIMPREDEV ' REIMPRESION DE REMITOS DE DEVOLUCION
        GoTo 10
      ElseIf OPX% = 4 Then
        Call ACCDIR("RECEPMA")
    End If
99  Command8.Enabled = True
End Sub
Sub DEVOLMATE()
   ' llamada a formulario de devolucion de materiales
   If EXISTE%("TRAZAB04.EXE") > 0 Then
     DEVOLMAT04.Show 1
   Else
     Call MENSERR(24, "No Disponible en la Presente Versión.")
   End If
   
End Sub
Sub BORRAGUION()
    Call APERBASDAT("STOCKS")
    SQLX$ = "DELETE * FROM MOVISTO "
    SQLX$ = SQLX$ + " WHERE CODIMOVI = 'BR' AND DOPRICOR = '-' AND DOSECCOR = '-'"
    Stocks.Execute SQLX$
    '
    Screen.MousePointer = 1
    Call COMUNI("Movimientos Eliminados")
End Sub

Private Sub configuracion_Click()
   Call Command14_Click
End Sub

Private Sub consultadeordenesdecompra_Click()
     Call Command13_Click
End Sub

Private Sub entradaxfabricacion_Click()
    Call Command6_Click
End Sub

Private Sub Form_Load()
    '
    Visible = False
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture14.Visible = True
      Picture15.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Frame7.BackColor = &HFCD7B6
      Me.Frame8.BackColor = &HFCD7B6
      Me.Option1.BackColor = &HFCD7B6
      Me.Option2.BackColor = &HFCD7B6
      MARCOBARRA.Top = 5950
      Set MARCOBARRA.Container = Me.Picture15
      
    End If
    If InStr(UCase$(EMPREAC$), "VOITH") > 0 Or InStr(UCase$(EMPRELI$), "VOITH") > 0 Then
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
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    If EXISTE%("CALPRO02.EXE") > 0 Then
      Command10.Visible = True
    End If
    '
    If EXISTE%("TRAZAB04.EXE") > 0 Then
      Command17.Visible = True
    End If
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
               NOC3& = !NumOC
               DP2% = !DepoENT
           Else
               Call COMUNI("Imposible Continuar. Número de B/R no Válido")
               Screen.MousePointer = 1
               Exit Sub
         End If
    End With
    '
    '
    Call ABRECOMPRAS
    With OcomDeta
            SQLX$ = " NUME_OCOM = " & NOC3& & ""
            'ALCANZA TAMBIEN POR Q LA O/C SOLO PUEDE
            'TENER UN ITEM.
            .FindFirst SQLX$
            If .NoMatch = False Then
                On Error Resume Next
                IMPCONCO% = !ICONTAB
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
       .LABEL5 = NC5%
       .Label11 = RASOCLI(NC5% * (-1))
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

Sub BOLERECEP()
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
     Set BolreTemp = Stocks.OpenRecordset("SELECT DISTINCT NumeBoRe from BOLRECEP ORDER BY NumeBoRe DESC")
     On Error Resume Next
     BolreTemp.MoveFirst
     If Err < 1 Then
       NROBOL& = 1 + BolreTemp!NUMEBORE
     End If
     On Error GoTo 0
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
     Call REPLA(Y$, RASOCLI$(NC1%), 9, 32)
     Call REPLA(Y$, DOMICLI$(NC1%), 41, 32)
     CLOCA$ = TRIM$(CPOSCLI$(NC1%)): If CLOCA$ <> "" Then CLOCA$ = CLOCA$ + " "
     Call REPLA(Y$, CLOCA$ + LOCACLI$(NC1%), 73, 32)
     Call REPLA(Y$, Chr$(0), 121, 1)
    '
     OBJE$ = "BOLREC-ENCB"
     If CONTROL$("", "NUMEBREC") = "MANUAL" Then
         OBJE$ = "BOLREC-ENCB-M"
     End If
     '
     ' CONTROL DE VIGILANCIA
     If CONTROL$("BOLRECEP", "NUCONVIG") = "SI" Then
         OBJE$ = "BOLREC-ENCC"
     End If
     '
     TICOMPRO$ = TRIM$(CONTROL$("", "COMPBREC"))
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
     '\\\OT-08-0684-OD-09/09/08///
     FACTU$ = TRIM$(Mid$(PEDY$, 129, 16)) ' NUMERO DE FACTURA
     '\\\OT-08-0684-OD-FIN///
     '
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
     If CONTROL$("BOLRECEP", "SUPREMI") <> "SI" Then  '\\\OT-06-0515-RG-28/11/06///
        R1% = EXIREMI(NR1$, NC2%)
        If R1% = 1 Then
           Call COMUNI("El Número de Remito ya fue Procesado\en Recepción Anterior.")
           GoTo 15
        End If
     End If
    'HASTA ACA NUEVO PARA GUIDI
1112 Call CIERRARCH("*.*")
     '
     ' verificar que hay pendiente
     Screen.MousePointer = 11
     Call ABRECOMPRAS
     SQLX$ = "Select * from OCOMDETA"
     SQLX$ = SQLX$ + " Where Stat_Ocom < 8 "
     SQLX$ = SQLX$ + " AND  Stat_Ocom <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
     SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
     SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
     Set COMPTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
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
         If !Cod_Inte > 0 Then
             HAYCODI% = 1
           Else
             HAYLIBRE% = 1
         End If
       .MoveNext
       Loop
     End With
     ' INCLUYE DENTRO DE LOS CODIFICADOS A LAS CONSIGNACIONES NO RENDIDAS
     Call ABRESTOCKS
     SQLX$ = "Select * from MACONSIG"
     SQLX$ = SQLX$ + " WHERE ((Cant_Rendid < Cant_Entreg - 0.001) "
     SQLX$ = SQLX$ + " OR Cant_Entreg IS NULL) "
     SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
     Set consigna = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     With consigna
       On Error Resume Next
       .MoveFirst
       If Not Err Then HAYCODI% = 1
     End With
     consigna.Close
     Set consigna = Nothing
     On Error GoTo 0
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
42       Call RECEOCLI(NC%, FF%, NROBOL&, REMI$, DP1%, NUCONVI$)
         Screen.MousePointer = 1
         'If CANCELX1% = 1 Then GoTo 45 '\\\OT-06-00178-WAR-12/05/06///
         GoTo 1100
     End If
     'HASTA ACA NUEVO
     '
45  DARX$ = TRIM$(CONTROL$("BOLRECEP", "DIANREC"))
     If DARX$ = "" Then
          DIANRE& = 10000
        Else
          DIANRE& = Val(DARX$)
     End If
     HOIX1 = HOY(HOS$)
     '
     ANUCA% = 0
     If CONTROL$("BOLRECEP", "ANUCANTI") = "SI" Then
       ANUCA% = 1
     End If
     '
     Screen.MousePointer = 11
     Call ABRECOMPRAS
     SQLX$ = "Select * from OCOMDETA"
     SQLX$ = SQLX$ + " Where Stat_Ocom < 8 "
     SQLX$ = SQLX$ + " AND  Stat_Ocom <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
     SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
     SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
     SQLX$ = SQLX$ + " AND Cod_Inte > 0 " 'EXCLUYE LAS DE 0/C's TEXTO LIBRE
     Set COMPTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With COMPTEMP
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
         CAPEN = !Cant_Ocom - !Cant_rec
         COEFUNI = !Coef_Unids
         If COEFUNI < 0.0001 Then COEFUNI = 1
         If CAPEN > 0.005 Then
           NUOC& = !Nume_Ocom
           FENTRESO% = CVINT(!Fentre_Sol)
           FESO2 = FENTRESO%
           If DIENTRE%(HOIX1, FESO2) <= DIANRE& Then
             CI% = !Cod_Inte
             LOTECO$ = !Doc_Orig
             Y$ = REGBLAN$("!BOLREDET")
             Call REPLA(Y$, MKI$(CI%), 83, 2)
             Call REPLA(Y$, UNIMED$(CI%), 85, 4)
             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 89, 8)
             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 163, 8)
             If ANUCA% = 1 Then
               Call REPLA(Y$, MKD$(0), 89, 8)
             End If
             '
             Call REPLA(Y$, MKI$(!NuOrd_Item), 98, 2)
             Call REPLA(Y$, MKS$(0), 100, 4) 'no tiene AE
             '
             Call REPLA(Y$, LOTECO$, 107, 12)
             Call REPLA(Y$, MKS$(NUOC&), 151, 4)
             Call REPLA(Y$, MKI$(FENTRESO%), 187, 2)
             Call REPLA(Y$, MKS$(!N_OrServ), 189, 4)
             HAYOC% = 1
             JJ& = JJ& + 1
             Call GRAREG("!BOLREDET", Y$, JJ&)
           End If
         End If
        .MoveNext
       Loop
     End With
     '
     '/////****NUEVO****/////
     'TENGO QUE FIJARME SI TIENE ALGUN MATERIAL EL PROV
     'CONSIGNADO EN MACONSIG CON EL MISMO Nº DE PROV
1777 Call ABRESTOCKS
     SQLX$ = "Select * from MACONSIG"
     SQLX$ = SQLX$ + " WHERE ((Cant_Rendid < Cant_Entreg - 0.001) "
     SQLX$ = SQLX$ + " OR Cant_Entreg IS NULL) "
     
     SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
     Set consigtmp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     With consigtmp
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
         CI% = !Cod_Inte
         NUOC& = Val(Mid$(!Nro_Ocom, 4, 6))
         CAPEN = XVALO(!Cant_Entreg) - XVALO(!Cant_Rendid)
         Call REPLA(Y$, MKI$(CI%), 83, 2)
         Call REPLA(Y$, UNIMED$(CI%), 85, 4)
         Call REPLA(Y$, MKD$(0), 89, 8)
         Call REPLA(Y$, "MP " + TRIM$(CSTRING$(MKS$(CAPEN), 3, 12, 1, 2)) + " " + UNIMED$(CI%), 107, 12)
         Call REPLA(Y$, MKS$(NUOC&), 151, 4)
         Call REPLA(Y$, MKS$((-1) * i&), 189, 4)
         JJ& = JJ& + 1
         Call GRAREG("!BOLREDET", Y$, JJ&)
       .MoveNext
       Loop
     End With
     '\\\OT-06-0149-WAR-25/04/06///
1129 Call ABRECOMPRAS
     SQLX$ = "Select * from AUTOENT"
     SQLX$ = SQLX$ + " Where Stat_Aten < 8 "
     SQLX$ = SQLX$ + " AND  Stat_Aten <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
     SQLX$ = SQLX$ + " AND Acuen_Aten < Cant_Aten - 0.001 "
     SQLX$ = SQLX$ + " AND Npro_Aten = " & NC% & " "
     SQLX$ = SQLX$ + " AND Cod_Inte > 0 " 'EXCLUYE LAS DE 0/C's TEXTO LIBRE
     Set COMPTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With COMPTEMP
       '
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
         CAPEN = !Cant_Aten - !Acuen_Aten
         COEFUNI = 1 '!Coef_Unids
         If COEFUNI < 0.0001 Then COEFUNI = 1
         If CAPEN > 0.005 Then
           NUOC& = !Nume_Ocom
           FENTRESO% = CVINT(!Fentre_Sol)
           FESO2 = FENTRESO%
           If DIENTRE%(HOIX1, FESO2) <= DIANRE& Then
             CI% = !Cod_Inte
             Y$ = REGBLAN$("!BOLREDET")
             Call REPLA(Y$, MKI$(CI%), 83, 2)
             Call REPLA(Y$, UNIMED$(CI%), 85, 4)
             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 89, 8)
             Call REPLA(Y$, MKD$(CDbl(CAPEN * COEFUNI)), 163, 8)
             If ANUCA% = 1 Then
               Call REPLA(Y$, MKD$(0), 89, 8)
             End If
             '
             Call REPLA(Y$, MKI$(!NuOrd_Item), 98, 2)
               NANTEN& = 0
               On Error Resume Next
               NANTEN& = !Nume_Aten
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
     '
     '\\\OT-06-0149-WAR-FIN///
1130 If HAYOC% = 0 Then  '1129 antes de OT-06-0149
       Call COMUNI("No hay Ordenes de Compra Pendientes\para este Proveedor.")
     End If
     '
     Call SETULTREG("!BOLREDET", JJ&)
     Call CIERRARCH("!BOLREDET")
     '
1113 AVENTA$ = "BOLREDET"
     If ANUCA% = 1 Then AVENTA$ = "BOLRENUE"
     '
     VTB% = VENTABU%(AVENTA$ + "/NC/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(RASOCLI$(NC1%)) + " - " + REPLACAR$(FECHAREC$, "/", "-"))
     If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
       X$ = "": Y$ = "": YY$ = ""
       EBOL$ = "": PEDY$ = "": DESCRI$ = ""
       GoTo 1100
     End If
     '
     For YX& = 1 To ULTREG&("!BOLREDET")
       YY$ = REGLEIDO$("!BOLREDET", YX&)
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
     '///OT-8-639-15/09/08
1141 For YX& = 1 To ULTREG&("!BOLREDET")
        YY$ = REGLEIDO$("!BOLREDET", YX&)
        CI% = CVI(Mid$(YY$, 83, 2))
        CAN# = CVD(Mid$(YY$, 89, 8))
        CANRED# = ROUND#(CAN#, 1)
        If Abs((CAN# - CANRED#)) > 0.0005 Then
            Call MENSERR(24, "Las Cantidades Ingresadas Deben Ser Con Un Decimal.\Se Ha Ajustado el Articulo '" + CODEXT(CI%) + "'. Volver a Verificar y Continuar")
            Call REPLA(YY$, MKD$(CANRED#), 89, 8)
            Call GRAREG("!BOLREDET", YY$, YX&)
            GoTo 1113
         End If
     Next YX&
     '///OT-8-639-FIN

     CONTA% = ULTREG&("!BOLREDET")
     '
     BLOSOBRE% = 0
     If CONTROL$("BOLRECEP", "BLOSOBRE") = "SI" Then
       BLOSOBRE% = 1
       PORSOBRE = Val(CONTROL("BOLRECEP", "PORSOBRE"))
     End If
     '
1115 If CONTA% > 0 Then
       '
       For YX& = 1 To ULTREG&("!BOLREDET")
         YY$ = REGLEIDO$("!BOLREDET", YX&)
         CI% = CVI(Mid$(YY$, 83, 2))
         CAN# = CVD(Mid$(YY$, 89, 8))
         If CI% < 1 Then
           Call MENSERR(24, "Código no Válido.")
           GoTo 1113
         End If
         If DP1% = 99 Then
            If DEPOPRE%(CI%) < 1 Then
               Call MENSERR(24, "Falta Depósito Predeterminado\para el Codigo '" + TRIM$(CODEXT$(CI%)) + "'")
               GoTo 1113
            End If
         End If
         If ANUCA% < 1 Then
           If CAN# < 0.005 Then
             Call MENSERR(24, "Falta Cantidad.")
             GoTo 1113
           End If
         End If
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
             If CAN# > CVD(Mid$(YY$, 163, 8)) * (1 + PORSOBRE / 100) + 0.05 Then
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
       Call SETULTREG("!OBSERVOF", 0)
       Call CIERRARCH("!OBSERVOF")
       '
       VTB% = VENTABU%("OBSERVOF/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(RASOCLI$(NC1%)) + " - " + FECHAREC$)
       
       If VTB% < 0 Then
         If COMALTER%("Realmente Desea Cancelar ?\\Continuar\Cancelar") = 2 Then
           X$ = "": Y$ = "": YY$ = ""
           EBOL$ = "": PEDY$ = "": DESCRI$ = ""
           GoTo 1100
         End If
       End If
       '
1125   Call BLOQARCH("BOLRECEP")
       TINI = Int(Timer)
       While Int(Timer) < TINI + 5: Wend
       Call BLOQARCH("BOLRECEP")
       '
       Call ABRESTOCKS
       Set BolreTemp = Stocks.OpenRecordset("SELECT DISTINCT NumeBoRe from BOLRECEP ORDER BY NumeBoRe DESC")
       On Error Resume Next
       BolreTemp.MoveFirst
       'AQUI COMPARA SI EL NUMERO DE BOLETA DE RECEPCION  YA ESTA USADO EN LA TABLA BOLRECEP
       If Err < 1 Then
         If NROBOL& <= BolreTemp!NUMEBORE Then
           NROBOL& = 1 + BolreTemp!NUMEBORE
         End If
       End If
       On Error GoTo 0
       BolreTemp.Close
       '
       Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
       Call ABRESTOCKS
       Call ABRECOMPRAS
       '
1126   AINDI% = 0
       NORVENTA = ""
       For YX& = 1 To ULTREG&("!BOLREDET")
         YY$ = REGLEIDO$("!BOLREDET", YX&)
         CI% = CVI(Mid$(YY$, 83, 2))
         CAN# = CVD(Mid$(YY$, 89, 8))
         NUOC& = Abs(CVS(Mid$(YY$, 151, 4)))
         NUMOIT% = CVI(Mid$(YY$, 98, 2))   'EL Nº DE ORDEN DEL ITEM
         If CI% > 0 Then
           If CAN# > 0.005 Then
             CARECI = CAN#
             LOTECO$ = TRIM$(Mid$(YY$, 107, 12))
             '
             If CVD(Mid$(YY$, 163, 8)) >= 0.05 Then   ' solo si es comprado - no M.P. devuelto
               NATEN& = CVS(Mid$(YY$, 100, 4))
               If NATEN& = 0 Then 'O/C CERRADA'\\\OT-06-0149-WAR-26/04/06///
                 With OcomDeta
                   .FindFirst ("Nume_Ocom = " & NUOC& & " AND NuOrd_Item = " & NUMOIT% & " AND Stat_Ocom < 8 ")
                   If .NoMatch = False Then
                     .Edit
                     If IsNull(!ICONTAB) Then
                          IMPCONOC% = 0
                        Else
                          IMPCONOC% = !ICONTAB 'ASIGNO LA IMPUTACION CONTABLE
                     End If
                       COEFUNI = !Coef_Unids
                       If COEFUNI < 0.0001 Then COEFUNI = 1
                     !Cant_rec = !Cant_rec + CARECI / COEFUNI
                       CAPENDIEN = !Cant_Ocom - !Cant_rec
                       NUORSE& = !N_OrServ
'                   If CAPENDIEN >= 0.005 Then
'                       If COMALTER%("Esta O/C por Código " + TRIM$(CODEXT$(CI%)) + "\Queda con un Saldo Pendiente de " + TRIM$(FORMATNUM$(CAPENDIEN, "F10.1")) + " " + TRIM$(UNIMED$(CI%)) + ".\\Mantener Abierta\Cancelar Pendiente") = 2 Then
'                         !Stat_Ocom = 8
'                       End If
'                     Else
'                       !Stat_Ocom = 3
'                   End If
                   ' SE ELIMINO LO ANTERIOR PARA QUE NO HAYA INTERRUPCION
                   ' MIENTRAS ESTA GRABANDO. pROBLEMA DE NUMERO DE br
                     If CAPENDIEN < 0.05 Then
                       !Stat_Ocom = 3
                     End If
                     .Update
                   End If
                 End With
               ElseIf NATEN& > 0 Then '0/C ABIERTA (AE)
                 With DetoCabi
                   .FindFirst ("Nume_Ocom = " & NUOC& & " AND Cod_Inte = " & CI% & "  AND Stat_Ocom < 8 ") 'NuOrd_Item = " & NUMOIT% XQ ES EL Nº DE LA AE
                   If .NoMatch = False Then
                     .Edit
                     If IsNull(!ICONTAB) Then
                          IMPCONOC% = 0
                        Else
                          IMPCONOC% = !ICONTAB 'ASIGNO LA IMPUTACION CONTABLE
                     End If
                       COEFUNI = !Coef_Unids
                       If COEFUNI < 0.0001 Then COEFUNI = 1
                     !Cant_rec = !Cant_rec + CARECI / COEFUNI
                       CAPENDIEN = !Cant_Ocom - !Cant_rec
                       NUORSE& = !N_OrServ
                     If CAPENDIEN < 0.05 Then
                       !Stat_Ocom = 3
                     End If
                     .Update
                   End If
                 End With
                 '
                 With AutoEnt
                   .FindFirst ("Nume_Ocom = " & NUOC& & " AND NuOrd_Item = " & NUMOIT% & " AND Stat_Aten < 8  AND Nume_Aten= " & NATEN& & " ")
                   If .NoMatch = False Then
                     .Edit
                       COEFUNI = 1 ' !Coef_Unids
                       If COEFUNI < 0.0001 Then COEFUNI = 1
                     !Acuen_Aten = !Acuen_Aten + CARECI / COEFUNI
                       CAPENDIEN = !Cant_Aten - !Acuen_Aten
                       !Sald_Aten = CAPENDIEN
                     If CAPENDIEN < 0.05 Then
                       !Stat_Aten = 3
                     End If
                     .Update
                   End If
                 End With
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
             DESCRI$ = TRIM$(CODEXT$(CI%)) + " - " + TRIM$(DESCRIT$(CI%))
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
             '\\\OT-5-0413-EPB-21/06/05///
             Call CIERRARCH("INDIBOL")
             Call CIERRARCH("IBOLPEN")
             '\\\OT-5-0413-EPB-FIN///
             '
             NBR$ = TRIM$(Str$(NROBOL&))
             NORDE% = NORDE% + 1
             NORDES$ = TRIM$(Str$(NORDE%))
             IDENLO$ = "BR-" + NBR$ + "-" + NORDES$
             '
             Call REPLA(YY$, MKS$(NROBOL&), 1, 4)
             If TIMATEI% <> 2 Then Call REPLA(YY$, DESCRI$, 5, 64)
             '
             If TRIM$(Mid$(YY$, 69, 12)) = "" Then
               Call REPLA(YY$, IDENLO$, 69, 12)
             End If
             Call REPLA(YY$, MKI$(FF%), 81, 2)
             DEPX% = DP1%   ' : If DP1% = 99 Then DEPX% = DEPOPRE%(CI%)
             Call REPLA(YY$, Chr$(DEPX%), 97, 1)
             Call REPLA(YY$, MKI$(NC%), 105, 2)
             Call REPLA(YY$, REMI$, 119, 16)
             If CVD(Mid$(YY$, 163, 8)) >= 0.05 Then   ' solo si es comprado - no M.P. devuelto
               Call REPLA(YY$, MKS$(NUORSE&), 189, 4)
             End If
             Call GRAREG("!BOLREDET", YY$, YX&)
             '
''///**NUEVO ACA EDITO EL REGISTRO EN MACONSIG***///
'Call ABRESTOCKS
'With MaConsig
'  .FindFirst (" VAL(MID$(NRO_Ocom,4,6)) = " & NUOC& & " AND Cod_Inte = " & CI% & " ")
'  If .NoMatch = False Then
'    .Edit
'      !Cant_Rendid = !Cant_Rendid + CARECI / COEFUNI
'        CAPENDIEN = !Cant_Entreg - !Cant_Rendid
'      !Saldo_Pend = CAPENDIEN
'    .Update
'  End If
'End With
''////***HASTA ACA NUEVO****////
'
'
             With Bolrecep
              .AddNew
               !NUMEBORE = NROBOL&
               !Descrip = Left$(DESCRI$, 64)
               !IdenLote = Left$(IDENLO$, 12)
               !FechRecep = CVDAT(FF%)
               !Cod_Inte = CI%
               !Cod_Exte = Left$(CODEXT$(CI%), 16)
               !DescripItem = Left$(DESCRIT0$(CI%), 64)
               !Unidad = Left$(UNIMED$(CI%), 4)
               '\\\OT-08-0639-OD-11/09/08///
               !CantNom = CAN#
               '\\\OT-08-0639-OD-FIN///
               !DepoENT = DEPX%
               !NUMPROV = NC%
               !RaSo_Prov = Left$(RASOCLI$((-1) * NC%), 64)
                  LOTECO$ = TRIM$(LOTECO$)
                  While Len(LOTECO$) > 12
                    LOTECO$ = Mid$(LOTECO$, 2)
                  Wend
               !LoteCol_Prov = Left$(LOTECO$, 12)
               !NumConVig = Left$(NUCONVI$, 16)
               !NUMREM_PROV = Left$(REMI$, 16)
               !NumFac_Prov = ""
               !NumOC = NUOC&
               !ICONTAB = IMPCONOC% 'Imputación contable de la OC.
               !MoneVaUni = ""
               !Valo_Unit = 0
               !CanToApro = 0
               !CanToRech = 0
               !CanToFalta = 0
               !FechLibe = CVDAT(0)
               !NumeOrdSer = NUORSE&
               !Observa = " " 'OBSE$
               !Status = 0
               !UsuApro = " "
               !FeHoApro = CVDAT(0)
               '\\\OT-06-0149-WAR-26/04/06///
               !Nume_Aten = NATEN&
               '\\\OT-06-0149-WAR-FIN///
              .Update
             End With
             '
             If NUORSE& < 1 Then
               Call DESCACONSI(NC%, CI%, CAN#, NUOC&, "DEVOLU")
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
       Bolrecep.Close
       OcomEnca.Close
       OcomDeta.Close
'          For YX& = 1 To ULTREG&("!OBSERVOF")
'             YY$ = REGLEIDO$("!OBSERVOF", YX&)
'             YY$ = MKS$(NROBOL&) + YY$
'             Call REGAPP("OBSEBREC", YY$)
'          Next YX&
          '
'          Call CIERRARCH("BOLRECEP")
'          Call CIERRARCH("BOLREPEN")
'          Call CIERRARCH("INDIBOL")
'          Call CIERRARCH("IBOLPEN")
'          Call CIERRARCH("OBSEBREC")
'          '
'          Call FRESHECHO("IBOLPEN")
'          Screen.MousePointer = 1
'          '
       If TRIM$(FORIPRE$) <> "" Then
         '
         CANCELPRINT% = 0
         If CONTROL("BOLRECEP", "SUPRINT") = "SI" Then CANCELPRINT% = 1
         '
         TIPODOC$ = "Boleta de Recepción"
         DESTIDOC% = (-1) * Abs(NC1%)
         NUMEDOC$ = Str$(NROBOL&)
         FEX = FF%: FECHDOC$ = FECHATEX$(FEX)
         '
         CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = DEPOSAL$
         CODPARAM$(2) = "NRO-REMI": DOCPARAM$(2) = REMI$
         CODPARAM$(3) = "N-AUTENT": DOCPARAM$(3) = NORVENTA
         CODPARAM$(4) = "NRO-PASE": DOCPARAM$(4) = NUCONVI$
         '\\\OT-08-0684-OD-09/09/08///
         CODPARAM$(5) = "REF-MAT1": DOCPARAM$(5) = FACTU$ 'FACTURA
         CAPARDOC% = 5
         '\\\OT-08-0684-OD-FIN///

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
         CODTABU1$(11) = "UBI-MPRI"
         CACOLTAB1% = 11
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
                 If UNID$ = "" Then UNID$ = UNIMED$(CI%)
                 CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
                 KILAJE$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 10, 3, 2)
                 FEX = CVI(Mid$(YY$, 155, 2))
                 FESOLEN$ = FECHATEX$(FEX)
                 FEPLA$ = FEPLANO$(CI%)
                 NULOTE$ = TRIM$(Mid$(YY$, 107, 12))
                 If NULOTE$ = "" Then
                    NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
                 End If
                 NUOCO$ = FORMATNUM$(Abs(CVS(Mid$(YY$, 151, 4))), "F7,0")
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
                 TETABU1$(10, CAITAB1%) = FESOLEN$
                 '
                 TETABU1$(11, CAITAB1%) = ""
                 If DP1% = 99 Then
                    TETABU1$(11, CAITAB1%) = TRIM$(Str$(DEPOPRE%(CI%)))
                 End If
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
       End If
          '
1199 End If
     '
11999 Call CIERRARCH("BOLRECEP")
      Call CIERRARCH("*.*")
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
     If CVS(Mid$(X0$, 75, 4)) >= NUORSE& Then
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
   HOBA% = HORANUM(Time$)
   USUA% = USNBR% Mod 100
   DBAJA$ = MKI$(FEBA%) + MKI$(HOBA%) + Chr$(USUA%) + Chr$(1) + "B.Rec." + TRIM$(Str$(NROBOL&))
   '
   NUOROP% = Asc(Mid$(X0$, 67, 1))
   Call REPLA(X0$, MKD$(CAN#), 87, 8)
   Call REPLA(X0$, MKD$(SALRE#), 95, 8)
   Call REPLA(X0$, "3ROS  ", 151, 6)
   Call REPLA(X0$, MKI$(FEBA%), 157, 2)
   Call REPLA(X0$, MKI$(HORANUM(Time$)), 159, 2)
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

Private Sub RECEOCLI(NC2%, FR%, NBR&, NREMI$, DP2%, NUCONVI$)
   'MATERIAL TEXTO LIBRE
   DARX$ = TRIM$(CONTROL$("BOLRECEP", "DIANREC"))
     If DARX$ = "" Then
          DIANRE& = 10000
        Else
          DIANRE& = Val(DARX$)
     End If
     HOIX1 = HOY(HOS$)
     '
     ANUCA% = 0
     If CONTROL$("BOLRECEP", "ANUCANTI") = "SI" Then
       ANUCA% = 1
     End If
     '
     Call COPYSTRU("INDIOCOM", "INDIRETE")
     '
10   Screen.MousePointer = 11
     Call ABRECOMPRAS
     SQLX$ = "Select * from OCOMDETA"
     SQLX$ = SQLX$ + " Where Stat_Ocom < 8 "
     SQLX$ = SQLX$ + " AND  Stat_Ocom <> - 1 " 'EXCLUYE LAS O/C's NO APROB.
     SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
     SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC2% & " "
     SQLX$ = SQLX$ + " AND Cod_Inte = 0 " ' 0/C's TEXTO LIBRE
     Set COMPTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     With COMPTEMP
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
         Call REPLA(Y$, MKS(!Nume_Ocom), 1, 4)
         Call REPLA(Y$, TRIM$(Str$(!NuOrd_Item)) + ") " + DESIT$ + " (S=" + TRIM$(FORMATNUM$(!Cant_Ocom - !Cant_rec, "F8.1")) + ")", 5, 44)
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
     With OcomDeta
       NOC3& = Val(VDEVU$)
       NORIT& = Val(VALACT2$("!INDIRETE"))
       '
       SQLX$ = " NUME_OCOM = " & NOC3&
       SQLX$ = SQLX$ & " AND NuOrd_Item = " & NORIT&
       SQLX$ = SQLX$ + " AND Stat_Ocom < 8 "         ' NO ANULADAS
       SQLX$ = SQLX$ + " AND  Stat_Ocom <> - 1 "     ' EXCLUYE NO APROB.
       '
       .FindFirst SQLX$
       If .NoMatch = False Then
           '
           On Error Resume Next
           NROC& = !Nume_Ocom
           DECRICOR$ = TRIM$(!Descritem)
           DECRILAR$ = !Delibre
           NC% = !Npro_Ocom
           CANREN# = !Cant_Ocom - !Cant_rec
           IMPCONCO% = !ICONTAB
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
     If CONTROL$("BOLRECEP", "ANUCANTI") = "SI" Then
       ANUCA% = 1
     End If
     '
     With PRESOCLIB04
       .Label4 = NROC&
       .Label12 = NORIT&
       .Text1 = DECRICOR$
       .Text2 = DECRILAR$
       .LABEL5 = NC%
       .Label11 = RASOCLI(NC% * (-1))
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
'             Call REPLA(Y$, MKI$(!NuOrd_Item), 98, 2)
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
'         NUMOIT% = CVI(Mid$(YY$, 98, 2))   'EL Nº DE ORDEN DEL ITEM
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

Sub GENIBOLPE()
    '
    Screen.MousePointer = 11
    Call ABRESTOCKS
    '
    SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM BOLRECEP "
    SQLX$ = SQLX$ + " WHERE Status = 0 "
    SQLX$ = SQLX$ + " AND Cod_Inte > 0 "   ' SOLO MATERIAL CODIFICADO
    SQLX$ = SQLX$ + " ORDER BY NumeBoRe "
    '
    Set COBRECEPTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With COBRECEPTMP
      On Error Resume Next
      .MoveLast
      FIN& = .RecordCount
      If Err Then
        Call MENSERR(24, "No Existen Boletas de Recepcion Pendientes.")
        On Error GoTo 0
        JJ& = 0: K& = 0
        GoTo 90
      End If
      On Error GoTo 0
      .MoveFirst
      II& = 0
      JJ& = 0
      K& = 0
      REBLA$ = REGBLAN$("BOLREPEN")
      REBLB$ = REGBLAN$("IBOLPEN")
      '
      Do While (.EOF = False)
         II& = II& + 1
         Call TRACE(20, II&, FIN&)
         Y$ = REBLA$
         Call REPLA(Y$, MKS$(!NUMEBORE), 1, 4)
         Call REPLA(Y$, SAFETEXT$(!IdenLote), 69, 12)
         Call REPLA(Y$, MKI$(CVINT(!FechRecep)), 81, 2)
         Call REPLA(Y$, MKI$(!Cod_Inte), 83, 2)
         Call REPLA(Y$, SAFETEXT$(!Unidad), 85, 4)
         Call REPLA(Y$, MKD$(!CantNom), 89, 8)
         Call REPLA(Y$, Chr$(!DepoENT), 97, 1)
         Call REPLA(Y$, MKI$(!NUMPROV), 105, 2)
         Call REPLA(Y$, SAFETEXT$(!LoteCol_Prov), 107, 12)
         Call REPLA(Y$, SAFETEXT$(!NUMREM_PROV), 119, 16)
         Call REPLA(Y$, SAFETEXT$(!NumFac_Prov), 135, 16)
         Call REPLA(Y$, MKS$(!NumOC), 151, 4)
         Call REPLA(Y$, SAFETEXT$(!MoneVaUni), 155, 4)
         Call REPLA(Y$, MKS$(!Valo_Unit), 159, 4)
         Call REPLA(Y$, MKD$(!CanToApro), 163, 8)
         Call REPLA(Y$, MKD$(!CanToRech), 171, 8)
         Call REPLA(Y$, MKD$(!CanToFalta), 179, 8)
         Call REPLA(Y$, MKI$(CVINT(!FechLibe)), 187, 2)
           NAENT& = 0
           On Error Resume Next
           NAENT& = !Nume_Aent
           On Error GoTo 0
         Call REPLA(Y$, MKS$(NAENT&), 189, 4) '151, 4)'\\\OT-06-0241-WAR-14/06/06///
         Call REPLA(Y$, !Observa, 193, 64)
         JJ& = JJ& + 1
         Call GRAREG("BOLREPEN", Y$, JJ&)
         '
         If !NUMEBORE <> ANTNUM Then
           XX$ = REBLB$
           Call REPLA(XX$, MKS$(!NUMEBORE), 1, 4)
           FEX$ = FECHATEX$(CVINT(!FechRecep))
           Call REPLA(XX$, FEX$, 5, 8)
           Call REPLA(XX$, SAFETEXT$(!NUMREM_PROV), 15, 16)
           Call REPLA(XX$, ECOARCH$("PROVED1", MKI$(!NUMPROV)), 32, 29)
           K& = K& + 1
           Call GRAREG("IBOLPEN", XX$, K&)
           ANTNUM = !NUMEBORE
        End If
        '
      .MoveNext
      Loop
    End With
    '
90  Set COBRECEPTMP = Nothing
    Call SETULTREG("BOLREPEN", JJ&)
    Call CIERRARCH("BOLREPEN")
    Call SETULTREG("IBOLPEN", K&)
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
        TICOMPRO$ = TRIM$(CONTROL$("", "COMPBREC"))
        
1120    'Call CLEARVEN
        X$ = "": Y$ = "": z$ = "": YY$ = ""
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
                           If UNID$ = "" Then UNID$ = UNIMED$(CI%)
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
    SQLX$ = SQLX$ + " WHERE TRIM(NumRem_Prov) = '" & TRIM(NR1$) & "' "
    SQLX$ = SQLX$ + " AND NumProv = " & NC1% & ""
    Set EXITMP = Stocks.OpenRecordset(SQLX$, 4)
    '
    EXIREMI = 0
    '
    With EXITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Exit Function
      End If
      On Error GoTo 0
      EXIREMI = 1
    End With
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
    '
End Sub
'
Sub SHOWBORE(NPX&)
    '
    TIDOCUM$ = "Boleta de Recepción"
    NUDOCU$ = TRIM$(Str$(NPX&))
    NUDOCU$ = " " + NUDOCU$ + " -"
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        If InStr(XX$, NUDOCU$) > 4 Then
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
    Set COBRECEPTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With COBRECEPTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        Call MENSERR(24, "No Existen Boletas de Recepcion Pendientes.")
        On Error GoTo 0
        Exit Sub
      End If
      On Error GoTo 0
      JJ& = 0
      K& = 0
      TC& = 0
      Do While (.EOF = False)
            Y$ = REGBLAN("BOLREPEN")
            '
            Call REPLA(Y$, MKS$(!NUMEBORE), 1, 4)
            Call REPLA(Y$, !IdenLote, 69, 12)
            Call REPLA(Y$, MKI$(CVINT(!FechRecep)), 81, 2)
            Call REPLA(Y$, MKI$(!Cod_Inte), 83, 2)
            Call REPLA(Y$, !Unidad, 85, 4)
            Call REPLA(Y$, MKD$(!CantNom), 89, 8)
            Call REPLA(Y$, Chr$(!DepoENT), 97, 1)
            Call REPLA(Y$, MKI$(!NUMPROV), 105, 2)
            Call REPLA(Y$, !LoteCol_Prov, 107, 12)
            Call REPLA(Y$, !NUMREM_PROV, 119, 16)
            Call REPLA(Y$, !NumFac_Prov, 135, 16)
            Call REPLA(Y$, MKS$(!NumOC), 151, 4)
            Call REPLA(Y$, !MoneVaUni, 155, 4)
            Call REPLA(Y$, MKS$(!Valo_Unit), 159, 4)
            Call REPLA(Y$, MKD$(!CanToApro), 163, 8)
            Call REPLA(Y$, MKD$(!CanToRech), 171, 8)
            Call REPLA(Y$, MKD$(!CanToFalta), 179, 8)
            Call REPLA(Y$, MKI$(CVINT(!FechLibe)), 187, 2)
            Call REPLA(Y$, !Observa, 193, 64)
            '
            If !Cod_Inte <= 0 Then HAYMATERIALIB% = 1 'HAY MATERIAL TEXTO LIBRE
            '
            JJ& = JJ& + 1
            If !Cod_Inte > 0 Then
                 TC& = TC& + 1
                Call GRAREG("BOLREPEN", Y$, TC&)
            End If
            '
            XXZ$ = REGBLAN("BOREPELI")
            XXZ$ = Left$(Y$, 82) + Space$(174) + Mid$(Y$, 85)
            If !Cod_Inte <= 0 Then
                'Call REPLA(XXZ$, MKI$(!COD_INTE), 9, 2)
                Call REPLA(XXZ$, !Cod_Exte, 83, 16)
                Call REPLA(XXZ$, !DescripItem, 99, 64)
              Else
               CIXI% = !Cod_Inte
               Call REPLA(XXZ$, CODEXT$(CIXI%), 83, 16)
               Call REPLA(XXZ$, DESCRIT0$(CIXI%), 99, 64)
            End If
            '
            NUOCO1& = !NumOC
            Call ABRECOMPRAS
            With OcomDeta 'PARA OBTENER EL DESTINO Y GRUPO
                '
                SQLX$ = " NUME_OCOM = " & NUOCO1& & ""
                .FindFirst SQLX$
                If .NoMatch = False Then
                  '
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
              Call REPLA(XX$, !NUMREM_PROV, 15, 16)
              Call REPLA(XX$, ECOARCH$("PROVED1", MKI$(!NUMPROV)), 32, 29)
              K& = K& + 1
              Call GRAREG("IBOLPEN", XX$, K&)
              ANTNUM = !NUMEBORE
           End If
           '
      .MoveNext
      Loop
    End With
    Set COBRECEPTMP = Nothing
    Call SETULTREG("BOLREPEN", JJ&)
    Call CIERRARCH("BOLREPEN")
    Call SETULTREG("IBOLPEN", K&)
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
    Set OCMODIFTEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
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
            Call REPLA(XXY$, MKS(!Nume_Ocom), 1, 4) 'Nº O/C
            Call REPLA(XXY$, !Refe_OcoM, 5, 44) ' REFERENCIA
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
        Static CREALOTE%
        Dim REGBO&(2048)
        '
        If CREALOTE% = 0 Then
           CREALOTE% = (-1)
           If CONTROL$("", "RETRAZA") = "DECONSU" Then
              CREALOTE% = 1
           End If
        End If
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
        X$ = "": Y$ = "": z$ = "": YY$ = ""
        EBOL$ = "": PEDY$ = "": DESCRI$ = ""
                           'NBR$ = OBJPLA$("SELBOLPEN", "")
                           'If NBR$ = "" Then
                           '  On Error Resume Next
                           '  Movisto.Close
                           '  Bolrecep.Close
                           '  SalStock.Close
                           '  On Error GoTo 0
                           '  Exit Sub
                           'End If
        Call SELECHO("IBOLPEN")
        NBR1$ = TRIM$(VALACT1$("IBOLPEN"))
        If NBR1$ = "" Then
          Exit Sub
        End If
        NBR2$ = VALACT2$("IBOLPEN")
        NUREPRO$ = TRIM$(Mid$(NBR2$, 11, 16))
        '
                           'NROBOL& = CVS(Left$(NBR$, 4))
        NROBOL& = Val(NBR1$)
        If NROBOL& < 1 Then
            Exit Sub
        End If
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
                Call REPLA(Y$, RASOCLI$(NC1%), 9, 32)
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
11215   z$ = OBJPLA$("BOLREC-MUESTRW", Y$)
        If z$ = "" Then
            GoTo 1120
        End If
        DP1% = Asc(Mid$(z$, 121, 1))
        If DP1% < 1 Or ECOARCH$("ECODEP", Chr$(DP1%)) = "" Then
          Call MENSERR(24, "Falta Indicar Depósito de Entrada")
          GoTo 11215
        End If
        '
        Call ABRESTOCKS
        ABUS$ = "NumeBoRe = " & NROBOL & " AND NumRem_Prov = '" & NUREPRO$ & "' "
        With Bolrecep
          .FindFirst ABUS$
          If .NoMatch = True Then
            Call MENSERR(24, "Boleta de Recepción no Encontrada.")
            GoTo 1100
          End If
          '
          CIXBR% = !Cod_Inte
          RESP5% = 1  ' Material Codificado
          If CIXBR% = 0 Then RESP5% = 2    ' mATERIAL lIBRE
        End With
        '
        If RESP5% = 2 Then
           Call APROCLI(NC%, FF%, NROBOL&, NUREPRO$, DP1%)
           Screen.MousePointer = 1
           GoTo 1100
        End If
        'HASTA ACA NUEVO
        '
        For U& = 1 To ULTREG&("!BOLREDET")
          X$ = REGLEIDO$("!BOLREDET", U&)
          CI% = CVI(Mid$(X$, 83, 2))
          DPXX% = DP1%
          If DP1% = 99 Then DPXX% = DEPOPRE%(CI%)
          Call REPLA(X$, Chr$(DPXX%), 97, 1)
          Call GRAREG("!BOLREDET", X$, U&)
        Next U&
        Call CIERRARCH("!BOLREDET")
        '
1122    AVENTA$ = "APROBOL"
        If CONTROL$("BOLRECEP", "CERCAPRO") = "SI" Then AVENTA$ = "APROBOLG"
        VTB% = VENTABU%(AVENTA$ + "/NC/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(RASOCLI$(NC1%)) + " - " + FECHAREC$)
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
          If Abs(APRO#) < 0.005 Then
            If Abs(RECHA#) < 0.005 Then
              If Abs(FALTA#) < 0.005 Then
                 GoTo 1123                        ' NO SE PROCESO MOVIMIENTO
              End If
            End If
          End If
          '
          DPXX% = Asc(Mid$(X$, 97, 1))
          CIXI% = CVI(Mid$(X$, 83, 2))
          If DPXX% < 1 Then
             Call MENSERR(24, "Falta Depósito de Entrada\para Código '" + TRIM$(CODEXT$(CIXI%)) + "'")
             GoTo 1122
          End If
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
1124    If VTB% < 0 Then GoTo 1120
        '
        CODRECHA$ = "RD"
        REMIRECH$ = "RD." + TRIM$(Str$(NROBOL&))
        If HAYRECHA% > 0 Then
          '\\\OT-05-0723-WAR-23/11/05///
          MODORECHA% = 0
          If ULTREG&("TADESTI") > 0 Then
            If ULTREG&("TAMOTIV") > 0 Then
              MODORECHA% = 1
            End If
          End If
          '\\\OT-05-0723-WAR-FIN///
          If TRIM$(CONTROL$("", "REMIRECH")) <> "" Then
            CODRECHA$ = TRIM$(CONTROL$("", "CODRRECH"))
            If CODRECHA$ = "" Then CODRECHA$ = "RD"
          End If
        End If
        '
        QUEDAPEN% = 0
        HAYRECHA% = 0
        Call ABRESTOCKS
        NOCLOSE% = 1: NOARRAS% = 1
        FIN& = ULTREG&("!BOLREDET")
        '
        NOCLOSE% = 1: NOARRAS% = 1
        Label12.Visible = True
        '
        SQLX$ = " SELECT * FROM BOLRECEP "
        SQLX$ = SQLX$ + "WHERE Status = 0 "
        SQLX$ = SQLX$ + "AND NumeBoRe = " & NROBOL& & " AND NumRem_Prov = '" & NUREPRO$ & "' "
        Set COBRECEPTMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
        '
        For U& = 1 To FIN&
          Call TRACE(20, U&, FIN&)
          Label12.Caption = TRIM$(Str$(U&))
          DoEvents
          X$ = REGLEIDO$("!BOLREDET", U&)
          CI% = CVI(Mid$(X$, 83, 2))
            CIMOVI%(CI%) = 1
          REFX$ = Mid$(X$, 5, 64)
          LOTE$ = Mid$(X$, 69, 12)
          CANOM# = CVD(Mid$(X$, 89, 8))
          DEPXX% = Asc(Mid$(X$, 97, 1))
          APRO# = CVD(Mid$(X$, 163, 8))
          '///OT-8-639-15/09/08
          APRORED# = ROUND#(APRO#, 1)
          If Abs((APRO# - APRORED#)) > 0.0005 Then
             Call MENSERR(24, "Las Cantidades Aprobadas Deben Ser Con Un Decimal.\Se Ha Ajustado el Articulo '" + CODEXT(CI%) + "'. Volver a Verificar y Continuar")
             Call REPLA(X$, MKD$(APRORED#), 163, 8)
             Call GRAREG("!BOLREDET", X$, U&)
             GoTo 1122
          End If
          '///OT-8-639-FIN
          RECHA# = CVD(Mid$(X$, 171, 8))
          FALTA# = CVD(Mid$(X$, 179, 8))
          CERCAPRO$ = TRIM$(Mid$(X$, 241, 16))
          If Abs(APRO#) < 0.005 Then
            If Abs(RECHA#) < 0.005 Then
              If Abs(FALTA#) < 0.005 Then
                QUEDAPEN% = 1
                GoTo 1128                        ' NO SE PROCESO MOVIMIENTO
              End If
            End If
          End If
          '
          ABUS$ = "NumeBoRe = " & NROBOL & " "
          ABUS$ = ABUS$ + "AND IdenLote = '" & LOTE$ & "' "
          ABUS$ = ABUS$ + "AND Cod_Inte = " & CI%
          With COBRECEPTMP
            .FindFirst ABUS$
            If .NoMatch = False Then
              .Edit
              '\\\OT-08-0639-OD-11/09/08///
              !CanToApro = APRO#
              '\\\OT-08-0639-OD-FIN///
              !CanToRech = RECHA#
              !CanToFalta = FALTA#
              !Status = 1
              !FechLibe = Date
              !UsuApro = USERNAME   'NUEVO
              !FeHoApro = Now       'NUEVO
              !Observa = USERNAME   'nuevo
              !CerCal_Prov = CERCAPRO$
              '\\\OT-06-0149-WAR-27/04/06///
              NANTEN& = 0
              On Error Resume Next
              NANTEN& = !Nume_Aten
              On Error GoTo 0
              '\\\OT-06-0149-WAR-FIN///
              '
              .Update
              '
1127          DOPRI$ = "BR." + TRIM$(Str$(NROBOL&))
              DOPRILA$ = DOPRI$
              DOSELA$ = TRIM$(NUREPRO$)
              DOSEC$ = TRIM$(NUREPRO$)
              DEPX% = DEPXX%    ' DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(CI%)
              If CANOM# >= 0.05 Then
                CANTI = CANOM#
                REFE$ = "Rto." + TRIM$(NUREPRO$) + " - " + RASOCLI$(NC1%)
                PPX% = InStr(REFX$, "- BC D.")
                If PPX% > 0 Then
                    REFE$ = Mid$(REFX$, PPX% + 5)
                End If
                While Len(DOSEC$) > 10: DOSEC$ = Mid$(DOSEC$, 2): Wend
                If Len(DOSEC$) < 8 Then DOSEC$ = "Rt." + DOSEC$
                Call MOVISTOK(FF%, CI%, LOTE$, "BR", DOPRILA$, DOPRI$, DOSELA$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC1%, DEPX%, CANTI)
                If CREALOTE% = 1 Then
                   Call CREALOTIN(FF%, CI%, LOTE$, CANOM#)
                End If
              End If
              '
              If RECHA# >= 0.05 Then
                HAYRECHA% = 1
                '\\\OT-05-0723-WAR-23/11/05///
                If MODORECHA% = 1 Then
                     GoSub PROCERECHA
                     CANTI = RECHA# * (-1)
                     REFE$ = "(" + TRIM$(Str$(MOTI%)) + ") " + TMOTI$
                     CODRECHA$ = "TR"
                     Call MOVISTOK(FF%, CI%, LOTE$, CODRECHA$, REMIRECH$, REMIRECH$, DOSELA$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC1%, DEPX%, CANTI)
                     '
                     CODRECHA$ = "CH"
                     REMIRECH$ = DOPRI$ ' Nº BR
                     'Hago el movimiento de Entrada al Deposito Asociado al Destino
                     ZZZ$ = FILTERGETRECORD$("TADESTI", 1, Chr$(DEST%))
                     DEPODES% = Asc(Mid$(ZZZ$, 32, 1)) 'DEP. DESTINO
                     CANTI = RECHA#
                     REFE$ = REFX$
                     Call MOVISTOK(FF%, CI%, LOTE$, CODRECHA$, REMIRECH$, REMIRECH$, DOSELA$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC1%, DEPODES%, CANTI)
                     '\\\OT-05-0723-WAR-FIN///
                   Else
                     CANTI = RECHA# * (-1)
                     REFE$ = "RECHAZO - Rto." + TRIM$(NUREPRO$)
                     Call MOVISTOK(FF%, CI%, LOTE$, CODRECHA$, REMIRECH$, REMIRECH$, DOSELA$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC1%, DEPX%, CANTI)
                End If
                If CREALOTE% = 1 Then
                   Call ACSALOTE(CI%, LOTE$)
                End If
              End If
              '
              If Abs(FALTA#) >= 0.05 Then
                CANTI = FALTA# * (-1)
                REFE$ = "FALTANTE - Rto." + TRIM$(NUREPRO$)
                If CANTI > 0 Then
                  REFE$ = "SOBRANTE - Rto." + TRIM$(NUREPRO$)
                End If
                Call MOVISTOK(FF%, CI%, LOTE$, "BR", DOPRILA$, DOPRI$, DOSELA$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC1%, DEPX%, CANTI)
                If CREALOTE% = 1 Then
                   Call ACSALOTE(CI%, LOTE$)
                End If
              End If
              '
              ' aqui re-activa la Orden de Compra por el rechazado mas el faltante
              RECHFAL = RECHA# + FALTA#
              If RECHFAL > 0.005 Then
                NOCORI& = CVS(Mid$(X$, 151, 4))
                If NOCORI& > 0 Then
                  If NANTEN& <= 0 Then 'SI ES O/C cerrada'\\\OT-06-0149-WAR-27/04/06///
                    Call ABRECOMPRAS
                    ABUS$ = "Nume_Ocom = " & NOCORI& & " "
                    ABUS$ = ABUS$ + "AND Cod_Inte = " & CI% & " "
                    'ABUS$ = ABUS$ + "AND (Cant_Rec+0.005) >= " & Abs(CANOM#)
                    With OcomDeta
                      .FindFirst ABUS$
1135                  If .NoMatch = False Then
                        CFXUNI# = !Coef_Unids
                        If CFXUNI# < 0.005 Then CFXUNI# = 1
                        RECHAX# = RECHFAL / CFXUNI#
                        If !Cant_rec >= RECHAX# - 0.0005 Then
                            .Edit
                            !Cant_rec = !Cant_rec - RECHAX#
                            !Stat_Ocom = 0
                            .Update
                            GoTo 1145
                          Else
                            .FindNext ABUS$
                            GoTo 1135
                        End If
                      End If
                      '
                      Call MENSERR(24, "No fue Posible Actualizar la Cantidad Pendiente\en la Orden de Compra Número " + TRIM$(Str$(NOCORI&)))
                      '
1145                End With
                  End If
                End If
                '\\\OT-06-0149-WAR-27/04/06///
                If NANTEN& > 0 Then
                  '
                  NOCORI& = CVS(Mid$(X$, 151, 4))
                  If NOCORI& > 0 Then
                     'RE ACTIVA O/C ABIERTA
                    Call ABRECOMPRAS
                    ABUS$ = "Nume_Ocom = " & NOCORI& & " "
                    ABUS$ = ABUS$ + "AND Cod_Inte = " & CI% & " "
                    With DetoCabi
                      .FindFirst ABUS$
1136                  If .NoMatch = False Then
                        CFXUNI# = !Coef_Unids
                        If CFXUNI# < 0.005 Then CFXUNI# = 1
                        RECHAX# = RECHFAL / CFXUNI#
                        If !Cant_rec >= RECHAX# - 0.0005 Then
                            .Edit
                            !Cant_rec = !Cant_rec - RECHAX#
                            !Stat_Ocom = 0
                            .Update
                            GoTo 1137
                          Else
                            .FindNext ABUS$
                            GoTo 1136
                        End If
                      End If
                      '
                      Call MENSERR(24, "No fue Posible Actualizar la Cantidad Pendiente\en la Orden de Compra Abierta Número " + TRIM$(Str$(NOCORI&)))
                      '
1137                End With
                   'RE ACTIVA AUTORIZACIÓN DE ENTREGA
                   ABUS$ = "Nume_Ocom = " & NOCORI& & " "
                   ABUS$ = ABUS$ + "AND Cod_Inte = " & CI% & " "
                   ABUS$ = ABUS$ + "AND Nume_Aten = " & NANTEN& & " "
                   With AutoEnt
                     .FindFirst ABUS$
1155                 If .NoMatch = False Then
                       CFXUNI# = 1 ' !Coef_Unids
                       If CFXUNI# < 0.005 Then CFXUNI# = 1
                       RECHAX# = RECHFAL / CFXUNI#
                       If !Acuen_Aten >= RECHAX# - 0.0005 Then
                           .Edit
                           !Acuen_Aten = !Acuen_Aten - RECHAX#
                           !Sald_Aten = !Cant_Aten - !Acuen_Aten
                           !Stat_Aten = 0
                           .Update
                           GoTo 1165
                         Else
                           .FindNext ABUS$
                           GoTo 1155
                       End If
                     End If
                     '
                     Call MENSERR(24, "No fue Posible Actualizar la Cantidad Pendiente\en la Autorización de Entrega Número " + TRIM$(Str$(NANTEN&)))
                     '
1165               End With
                  End If
                End If
                '\\\OT-06-0149-WAR-FIN///
              End If
              ' ACA MARCA LA LISTAMAT CON EL Nº DE BR.
              ' ESTO ES PARA MARCAR EN LA LISTA DE MATERIALES LOS RECIBIDOS
              Call ABREORTRA
              CODEX$ = CODEXT$(CI%)
              CANAREP = APRO#
              If TRIM$(CODEX$) <> "" Then
                  SQLX$ = "CodExte = '" & CODEX$ & "' "
                  SQLX$ = SQLX$ + "AND (Nro_Ocom <> '' AND ISNULL(Nro_Ocom) = FALSE AND ISEMPTY(Nro_Ocom) = FALSE) "
                  SQLX$ = SQLX$ + "AND (Nro_BolRe = '' OR ISNULL(Nro_BolRe) = TRUE OR ISEMPTY(Nro_BolRe) = TRUE) "
                  With ListaMat
                    .FindFirst SQLX$
11271                  If .NoMatch = False Then
                       'FORMATEO EL NUMERO DE BR
                       N1$ = CStr(NROBOL)
                       Do While Len(N1$) < 6
                         N1$ = "0" + N1$
                       Loop
                       NBR2$ = "BR-" + N1$ + "-00"
                       '
                       If !CANTID <= CANAREP + 0.001 Then
                        .Edit
                        !nro_bolre = NBR2$
                        .Update
                        CANAREP = CANAREP - !CANTID
                      End If
                      .FindNext SQLX$
                      GoTo 11271
                    End If
                  End With
              End If
            End If
          End With
1128    Next U&
        'IMPRIMIR STICKER
         ' ESTO ES PARA IMPRIMIR LOS MATERIALES RECIBIDOS
   STIKAP$ = TRIM$(UCase$(CONTROL$("", "STIKAPRO")))
   If Left$(STIKAP$, 4) <> "STIK" Then
     If Left$(STIKAP$, 5) <> "STICK" Then
       GoTo 199
     End If
   End If
   '

                  Call ABREORTRA
                  SQLX$ = " SELECT * FROM LISTAMAT "
                  SQLX$ = SQLX$ + "WHERE Nro_BOLRE = '" & NBR2$ & "' "
                  Set MATRECTMP1 = OTrabajo.OpenRecordset(SQLX$, 4)
                  With MATRECTMP1
                    On Error Resume Next
                    .MoveFirst
                    If Err Then
                        On Error GoTo 0
                        Call MENSERR(24, "No se ha Podido Imprimir correctamente los STICKERS.")
                    End If
                    On Error GoTo 0
                    CRR1 = 1
                    Do While .EOF = False
                        '
                        NEV1$ = ""
                        GRU1$ = ""
                        CLI1$ = ""
                        DEV1$ = ""
                        DGR1$ = ""
                        DCR1$ = ""
                        PLA1$ = ""
                        MAT1$ = ""
                        LLA1$ = ""
                        '
                        NEV1$ = !Nro_Otrab
                        GRU1$ = !Nro_Grup
                        DCR1$ = !Descri_Cor
                        PLA1$ = !NPlano
                        MAT1$ = !Material
                        LLA1$ = !CLAVE
                        CAN1$ = FORMATNUM(!CANTID, "F8.1")
                        '
                        If NEV1$ <> "" Then
                             With OrdeTrab
                               SQLX$ = " Nro_Otrab = '" & NEV1$ & "' "
                               .FindFirst SQLX$
                               If .NoMatch = False Then
                                 DEV1$ = !Refe_Otrab
                                 NC% = !Nclie_Otrab
                                 CLI1$ = RASOCLI$(NC%)
                               End If
                             End With
                            '
                             If GRU1$ <> "" Then
                               With GrupOtra
                                 SQLX$ = " Nro_Otrab = '" & NEV1$ & "' "
                                 SQLX$ = SQLX$ + " AND Nro_Grup = '" & GRU1$ & "' "
                                 .FindFirst SQLX$
                                 If .NoMatch = False Then
                                   DGR1$ = !Deno_Grup
                                 End If
                               End With
                             End If
                            '
                            Call APROCLIB04.PRINSTIKER(NEV1$, GRU1$, CLI1$, DEV1$, DGR1$, DCR1$, PLA1$, MAT1$, LLA1$, CAN1$, CRR1)
                        End If
                        CRR1 = CRR1 + 1
                        .MoveNext
                    Loop
                  End With
        '
199     NOCLOSE% = 0: NOARRAS% = 0
        Call VESARRGLO
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
        Call FILTERDELETE("IBOLPEN", 1, MKS$(NROBOL&))
        '
        If HAYRECHA% = 1 Then
          If MODORECHA% = 0 Then 'SALTEO IMPRESIÓN DE REMITO DE RECHAZO
            Call REMIRECHA(NC1%, DEPX%, REMIRECH$)
          End If
        End If
        '
200     Call CIERRARCH("*.*")
        Call BAJALDISCO
        '
        If CREALOTE% = 1 Then
           If EXISTE%("TRAZAB04.EXE") > 0 Then
              Call CONECRUN("TRAZAB04/BORECEP" + TRIM$(NBR1$), "Registro de Trazabilidad de B/R's")
           End If
        End If
        '
        If EXISTE%("CALPRO02.EXE") > 0 Then
          ICALRE04.LABEL5 = TRIM$(Str$(NROBOL&))
          ICALRE04.Show 1
        End If
        GoTo 1120
        '
PROCERECHA:
     'Muestro el Objeto
     VDEF$ = Space$(256)
     Call REPLA(VDEF$, MKI$(CI%), 1, 2)
     Call REPLA(VDEF$, MKI$(FF%), 3, 2)
     Call REPLA(VDEF$, MKI$(NC%), 5, 2)
     Call REPLA(VDEF$, Chr$(0), 7, 1)
     Call REPLA(VDEF$, Chr$(0), 8, 1)
1130 OBJ$ = OBJPLA$("DETARECHA2", VDEF$)
     'Valido el OBJ$
     OBJANT$ = OBJ$
     If TRIM(OBJ$) <> "" Then
       DEST% = Asc(Mid$(OBJANT$, 7, 1))
       MOTI% = Asc(Mid$(OBJANT$, 8, 1))
       REFX$ = TRIM$(Mid$(OBJANT$, 9, 48))
       '
       If TRIM$(ECOARCH$("TADESTI", Chr(DEST%))) = "" Then
         Call MENSERR(24, "Destino de Rechazo no Válido.")
         VDEF$ = OBJANT$
         GoTo 1130
       End If
       If TRIM$(ECOARCH$("TAMOTIV", Chr(MOTI%))) = "" Then
         Call MENSERR(24, "Motivo de Rechazo no Válido.")
         VDEF$ = OBJANT$
         GoTo 1130
       End If
       If TRIM$(REFX$) = "" Then
         Call MENSERR(24, "Referencia de Rechazo no Válida.")
         VDEF$ = OBJANT$
         GoTo 1130
       End If
       '
       GoTo 1131
     End If
     Call MENSERR(24, "Es Mandatorio Procesar el Rechazo")
     GoTo 1130
     '
1131 TDEST$ = TRIM$(ECOARCH$("TADESTI", Chr(DEST%)))
     TMOTI$ = TRIM$(ECOARCH$("TAMOTIV", Chr(MOTI%)))
Return
'
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
        X$ = "": Y$ = "": z$ = "": YY$ = ""
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
              Call REPLA(Y$, RASOCLI$(NC1%), 7, 32)
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
11215   z$ = OBJPLA$("DECLADISCRE", Y$)
        If z$ = "" Then
          GoTo 1120
        End If
        '
        'VALIDACION DE DATOS
        PEDI1$ = Mid$(z$, 169, 18)
        LOTE1$ = Mid$(z$, 189, 18)
        TRACL1$ = Mid$(z$, 207, 48)
        PROCL1% = CVI(Mid$(z$, 187, 2))
        CONTA1$ = Mid$(z$, 103, 48)
        '
        If TRIM$(PEDI1$) = "" Then
          Call MENSERR(24, "Falta Indicar Número de Pedido")
          Y$ = z$
          GoTo 11215
        End If
        If TRIM$(LOTE1$) = "" Then
          Call MENSERR(24, "Falta Indicar Número de Lote")
          Y$ = z$
          GoTo 11215
        End If
        If TRIM$(TRACL1$) = "" Then
          Call MENSERR(24, "Falta Indicar Trazabilidad del Cliente")
          Y$ = z$
          GoTo 11215
        End If
        If PROCL1% <= 0 Then
          Call MENSERR(24, "Falta Indicar Producto")
          Y$ = z$
          GoTo 11215
        End If
        '
        'ACA GRABO DISCREPA
        Call BLANARCH("!DISCREPA")
        For U& = 1 To ULTREG&("!BOLREDET")
          XXX$ = REGLEIDO$("!BOLREDET", U&)
          CODI1% = CVI(Mid$(XXX$, 83, 2))
          CANOM1# = CVD(Mid$(XXX$, 89, 8))
          ZZ$ = REGBLAN("!DISCREPA")
          Call REPLA(ZZ$, MKI$(CODI1%), 1, 2)
          Call REPLA(ZZ$, MKS$(CANOM1#), 27, 4)
          Call REPLA(ZZ$, MKS$(CANOM1#), 31, 4)
          Call GRAREG("!DISCREPA", ZZ$, U&)
        Next U&
        Call SETULTREG("!DISCREPA", U&)
        Call CIERRARCH("!DISCREPA")
        '
1122    VTB% = VENTABU%("DISCREPA/NC/TITUPAN=Boleta de Recepción" + Str$(NROBOL&) + " - " + TRIM$(RASOCLI$(NC1%)) + " - " + FECHAREC$)
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
           CODPARAM$(2) = "RASO-CLI": DOCPARAM$(2) = RASOCLI$(NC1%)
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
1100    Call GENIBOLPE
        If ULTREG&("IBOLPEN") < 1 Then
          Exit Sub
        End If
        '
1120    Call FRESHECHO("IBOLPEN")
        Screen.MousePointer = 1
        Call SELECHO("IBOLPEN")
        NROBOL& = Val(VALACT1$("IBOLPEN"))
        If NROBOL& < 1 Then Exit Sub
        '
        If COMALTER%("Confirme la Anulación de la B.R. " + Str$(NROBOL&) + "\\Cancelar Transacción\Confirmar Anulación") <> 2 Then
          Exit Sub
        End If
        '
        SQLX$ = " SELECT * FROM BOLRECEP "
        SQLX$ = SQLX$ + "WHERE Status = 0 "
        SQLX$ = SQLX$ + "AND NumeBoRe = " & NROBOL& & " "
        Set COBRECEPTMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
        '
        FIN& = ULTREG&("BOLREPEN")
        Call COPYSTRU("BOLREDET", "BOLRESEL")
        '
        ACTU% = 0
        For i& = 1 To FIN&
          X$ = REGLEIDO$("BOLREPEN", i&)
          If CVS(Left$(X$, 4)) = NROBOL& Then
            CANOM# = CVD(Mid$(X$, 89, 8))
            APRO# = CVD(Mid$(X$, 163, 8))
            RECHA# = CVD(Mid$(X$, 171, 8))
            FALTA# = CVD(Mid$(X$, 179, 8))
            '
            If CANOM# < 0.05 Or APRO# > 0.05 Or RECHA# > 0.05 Or Abs(FALTA#) > 0.05 Then GoTo 1121
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
              ABUS$ = "Nume_Ocom = " & NOCORI& & " "
              ABUS$ = ABUS$ + "AND Cod_Inte = " & CI% & " "
              With OcomDeta
                .FindFirst ABUS$
1135            If .NoMatch = False Then
                  CFXUNI# = !Coef_Unids
                  If CFXUNI# < 0.005 Then CFXUNI# = 1
                  RECHAX# = CANOM# / CFXUNI#
                  If !Cant_rec >= RECHAX# - 0.0005 Then
                      .Edit
                      !Cant_rec = !Cant_rec - RECHAX#
                      !Stat_Ocom = 0
                      .Update
                      GoTo 1145
                    Else
                      .FindNext ABUS$
                      GoTo 1135
                  End If
                End If
                '
                Call MENSERR(24, "No fue Posible Actualizar la Cantidad Pendiente\en la Orden de Compra Número " + TRIM$(Str$(NOCORI&)))
                '
1145          End With
            End If
            GoTo 11471
            '
            '
1146        ' CASO DE O/COMPRA ABIERTA CON A/ENTREGA
            Call ABRECOMPRAS
            ABUS$ = "Nume_Aten = " & NAENT& & " "
            ABUS$ = ABUS$ + "AND Cod_Inte = " & CI% & " "
            With AutoEnt
              .FindFirst ABUS$
11461         If .NoMatch = False Then
                RECHAX# = CANOM#
                If !Acuen_Aten >= RECHAX# - 0.0005 Then
                    .Edit
                    !Acuen_Aten = !Acuen_Aten - RECHAX#
                    !Stat_Aten = 0
                    .Update
                    '
                    NOCORI& = !Nume_Ocom
                    ABUS2$ = "Nume_Ocom = " & NOCORI& & " "
                    ABUS2$ = ABUS$ + "AND Cod_Inte = " & CI% & " "
                    With DetoCabi
                      .FindFirst ABUS$
11462                 If .NoMatch = False Then
                         If !Cant_rec >= RECHAX# - 0.0005 Then
                             .Edit
                             !Cant_rec = !Cant_rec - RECHAX#
                             !Stat_Ocom = 0
                             .Update
                             GoTo 11463
                           Else
                             .FindNext ABUS$
                             GoTo 11462
                         End If
                      End If
11463               End With
                    GoTo 11469
                  Else
                    .FindNext ABUS$
                    GoTo 11461
                End If
              End If
              '
              Call MENSERR(24, "No fue Posible Actualizar la Cantidad Pendiente\en la Orden de Compra Número " + TRIM$(Str$(NOCORI&)))
              '
11469       End With
            GoTo 11471
            '
            '
11471       ABUS$ = "NumeBoRe = " & NROBOL & " "
            ABUS$ = ABUS$ + "AND IdenLote = '" & LOTE$ & "' "
            ABUS$ = ABUS$ + "AND Cod_Inte = " & CI%
            With COBRECEPTMP
              .FindFirst ABUS$
1147            If .NoMatch = False Then
                .Delete
                .FindNext ABUS$
                GoTo 1147
              End If
            End With
            '
            Call GRAREG("BOLREPEN", String$(512, 0), i&)
            ACTU% = 1
            '
          End If
1121    Next i&
        '
        Call GENINDIBOL
        '
        If ACTU% > 0 Then
           Call COMUNI("Boleta de Recepción Anulada")
        End If
        GoTo 1100
        '
End Sub

Sub GENINDIBOL()
        '\\\OT-06-0543-RG-12/12/06///
        Call APERBASDAT("STOCKS")
        SQLY$ = " SELECT * FROM BOLRECEP "
        SQLY$ = SQLY$ + " ORDER BY NumeBoRe "
        Set BRECEPTMP = Stocks.OpenRecordset(SQLY$, dbOpenSnapshot)
        '
        ANTNUM = 0
        i& = 0
        J& = 0
        With BRECEPTMP
          .MoveFirst
          Do While (BRECEPTMP.EOF = False)
            NROBOL& = !NUMEBORE
            If NROBOL& <> ANTNUM Then
               X$ = REGBLAN$("INDIBOL")
               FEX = CVINT(!FechRecep)
               REMI$ = !NUMREM_PROV
               ST& = !Status
               NC1% = (-1) * !NUMPROV
               '
               Call REPLA(X$, MKS$(NROBOL&), 1, 4)
               Call REPLA(X$, FECHATEX$(FEX), 5, 8)
               Call REPLA(X$, REMI$, 15, 16)
               Call REPLA(X$, RASOCLI$(NC1%), 32, 29)
               i& = i& + 1
               Call GRAREG("INDIBOL", X$, i&)
               If ST& = 0 Then
                  J& = J& + 1
                  Call GRAREG("IBOLPEN", X$, J&)
               End If
               ANTNUM = NROBOL&
            End If
          .MoveNext
          Loop
          Call SETULTREG("INDIBOL", i&)
          Call SETULTREG("IBOLPEN", J&)
        End With
        '
        Screen.MousePointer = 1
        '\\\OT-06-0543-RG-FIN///
End Sub

Sub RECPROP()
    '
    If PRECOCIE% = 0 Then
       PRECOCIE% = (-1)
       If CONTROL$("", "PRECOCIE") = "SI" Then
          PRECOCIE% = 1
       End If
    End If
    '
10  HOOI% = HOY(HO$)
    If ULTREG&("ORDEFABR") < 1 Then   ' por numero de O/F
        GoTo 100
    End If
    '
    modo% = COMALTER%("Puede Declarar Entradas a Stock:\  \a) por Cierre Parcial o Total de Orden de Fabricación\o\b) en Forma Directa por Armado de Conjunto\\Cierre de O./Fabricación\Armado de Conjunto\Cancelar")
    On modo% GoTo 20, 100
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
          KKL& = 0
       On Error GoTo RECUERR
       XY$ = REGLEIDO$("!CAREPRO", U&)
          KKL& = 10
       CANTI = CVD(Mid$(XY$, 23, 8))
       SCRAP = 0
       CI% = CVI(Left$(XY$, 2))
       CODICONJ$ = CODEXT$(CI%)
       CICONJ% = CI%
          KKL& = 20
       '
       '\\\OT-06-0165-WAR-08/05/06///
       'If CANTI < 0 Or SCRAP < 0 Or CANTI + SCRAP < 0.05 Then
       If SCRAP < 0 Then
       '\\\OT-06-0165-WAR-FIN///
          Call MENSERR(24, "Cantidades Incorrectas en Código " + TRIM$(CODICONJ$))
          GoTo 120
       End If
          KKL& = 30
       '
       NORFABRI$ = ""
       CACONJU# = CDbl(CANTI + SCRAP)
       VACONTROL% = 1
          KKL& = 40
       '
       Call SETULTREG("!MATEROF", 0)
       Call CIERRARCH("!MATEROF")
          KKL& = 50
       Call CARCOMPO(CODICONJ$, NORFABRI$, CACONJU#, VACONTROL%, 2)
          KKL& = 60
       '
       If VACONTROL% < 0 Then GoTo 120           ' cancelacion
       '
       '\\\OT-05-0730-WAR-10/11/05///
       CAITMO% = 0
          KKL& = 70
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
          KKL& = 80
       If PUEDEMOVI%(DP2%) < 0 Then GoTo 110 ' CANCELA SALIDA
          KKL& = 90
       '\\\OT-05-0730-WAR-FIN///
       '
       Nro& = CVS(Mid$(XX$, 85, 4))
       NROOF& = CVS(Mid$(XY$, 19, 4))
       If NROOF& < 1 Then NROOF& = Nro&
          KKL& = 100
       DOPRI$ = "OF-" + TRIM$(Str$(NROOF&))
       DOSEC$ = "PF-" + TRIM$(Str$(Nro&))
       LOTE$ = Mid$(XY$, 3, 16)
       If LOTE$ = "" Then
          LOTE$ = DOPRI$
       End If
          KKL& = 110
       REFE$ = Mid$(XX$, 55, 30)
       If TRIM$(REFE$) = "" Then
          REFE$ = "Informe de Produccion"
       End If
       If CANTI < (-0.005) Then
          REFE$ = "Desarmado de Conjunto"
       End If
          KKL& = 120
       '
       NC% = 0
       VUNI# = 0
       MONE$ = ""
       MONEII% = 0
       CMOV$ = "PF"
       '
       NUORIT% = 0
       DEPX% = DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(CICONJ%)
          KKL& = 130
       Call MOVISTOK(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPX%, CANTI)
          KKL& = 140
       '
       YY$ = REGBLAN$("RECPROP")
       Call REPLA(YY$, MKI$(FECHA%), 1, 2)
       Call REPLA(YY$, Chr$(DP1%), 3, 1)
       Call REPLA(YY$, Chr$(DP2%), 4, 1)
       Call REPLA(YY$, REFE$, 5, 32)
       Call REPLA(YY$, MKI$(CI%), 37, 2)
       Call REPLA(YY$, MKS$(CANTI), 39, 4)
       Call REPLA(YY$, LOTE$, 43, 16)
       Call REPLA(YY$, MKS$(Nro&), 59, 4)
       Call REPLA(YY$, MKI$(-1), 63, 2)
          KKL& = 150
       Call REGAPP("RECPROP", YY$)
          KKL& = 160
       COD$ = CODEXT$(CICONJ%)
       CODCON$ = COD$
       '
          KKL& = 170
       For U2& = 1 To ULTREG&("!MATEROF")
          ZZ$ = REGLEIDO$("!MATEROF", U2&)
          CI2% = CVI(Left$(ZZ$, 2))
          CAN = -1 * CVS(Mid$(ZZ$, 5, 4))
          '
          If TRIM$(UCase$(UNIMED$(CI2%))) = "GK" Then
             CAN = CAN / 1000
          End If
          '
          If CI2% > 0 Then
             If CI2% <= NUMAS% Then
                If Abs(CAN) >= 0.05 Then
                   LOTE$ = ""
                   CMOV$ = "CF"
                   DOPRI$ = "OF-" + TRIM$(Str$(NROOF&))
                   DOSEC$ = "PF-" + TRIM$(Str$(Nro&))
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
          KKL& = 180
       Call CIERRARCH("!MATEROF")
       '
       XY$ = REGLEIDO$("!CAREPRO", U&)
       Call REPLA(XY$, MKD$(0), 23, 8)
       Call GRAREG("!CAREPRO", XY$, U&)
          KKL& = 190
       '
       On Error GoTo 0
129 Next U&
    Call CIERRARCH("MOVISTO")
    Call CIERRARCH("RECPROP")
    '
180 Call CIERRARCH("*.*")
    Call BLANARCH("!CAREPRO")
    Screen.MousePointer = 1
    GoTo 100
    '
RECUERR:
    MENSA$ = "Se ha Producido un Error " & Err.Number & " al Procesar el Código '" & TRIM$(CODICONJ$) & "' en el Registro Número " & U& & " y la Línea Número " & KKL& & "."
    MENSA$ = MENSA$ + " Informe a FLEXOFT para Recibir Orientación Correctiva: 4524-1284 / 4522-2048 (Int.32)"
    MsgBox MENSA$, vbExclamation
    Call CIERRARCH("*.*")
    Call FINAL("")
    '
End Sub


Private Sub impresora_Click()
    Option2.Value = True
    Call Command3_Click
End Sub

Private Sub impresora2_Click()
   Me.Option2.Value = True
   Call Command5_Click

End Sub

Private Sub inspeccionyaprobacion_Click()
   Call Command2_Click
End Sub


Private Sub pantalla_Click()
    Me.Option1.Value = True
    Call Command3_Click
End Sub

Private Sub pantalla2_Click()
   Me.Option1.Value = True
   Call Command5_Click
End Sub


Private Sub pendientesderecepcion_Click()
   Call Command4_Click
End Sub

Private Sub RecepcionesconRechazos_Click()
   Call Command16_Click
End Sub

Private Sub recepcionesporrangdefecha_Click()
   Call Command11_Click
End Sub

Private Sub recepcionessinordendecompra_Click()
    Call Command15_Click
End Sub

Private Sub recepcionprovisoria_Click()
    Call Command7_Click
End Sub

Private Sub rechazoporrangodefechas_Click()
    Call Command18_Click
End Sub

Private Sub reimprimirremitodevolucion_Click()
    Call REIMPREDEV
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub


Private Sub seleccionpararechazo_Click()
   Call DEVOLMATE
End Sub

Private Sub Trazabilidad_Click()
   Call Command17_Click
End Sub

Private Sub verboletarecepcion_Click()
   Call Command12_Click
End Sub



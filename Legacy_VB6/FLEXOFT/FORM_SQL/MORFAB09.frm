VERSION 5.00
Begin VB.Form MORFAB09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0F8&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Oficina de Programacion"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8715
   ClipControls    =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "MORFAB09.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8715
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   4560
      TabIndex        =   79
      Top             =   5520
      Width           =   1695
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   1080
      TabIndex        =   78
      Top             =   5520
      Width           =   2655
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Despacho de Materiales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   105
      TabIndex        =   61
      Top             =   3150
      Width           =   3690
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   960
         ScaleHeight     =   585
         ScaleWidth      =   2340
         TabIndex        =   62
         Top             =   840
         Width           =   2400
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
            Picture         =   "MORFAB09.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Ver y Re-Imprimir Remito"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Re-Imprimir Remito"
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
            Left            =   720
            TabIndex        =   63
            Top             =   90
            Width           =   1440
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   66
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command42 
            Caption         =   "Inventarios"
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
            Left            =   210
            TabIndex        =   72
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command40 
            Caption         =   "Negativos"
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
            Left            =   2205
            TabIndex        =   71
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command39 
            Caption         =   "Consignación"
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
            Left            =   210
            TabIndex        =   70
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command38 
            Caption         =   "Críticos"
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
            Left            =   2205
            TabIndex        =   69
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command37 
            Caption         =   "Estadísticas"
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
            Left            =   210
            TabIndex        =   68
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command36 
            Caption         =   "Análisis A-B-C"
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
            Left            =   2205
            TabIndex        =   67
            Top             =   1890
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   600
         ScaleHeight     =   585
         ScaleWidth      =   2340
         TabIndex        =   64
         Top             =   360
         Width           =   2400
         Begin VB.CommandButton Command71 
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
            Picture         =   "MORFAB09.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Generar Remito"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Imprimir Remito segun O/Despacho"
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
            Left            =   680
            TabIndex        =   65
            Top             =   30
            Width           =   1785
         End
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Gestion de Compras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   3960
      TabIndex        =   49
      Top             =   3625
      Width           =   3690
      Begin VB.PictureBox Picture11 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   480
         ScaleHeight     =   585
         ScaleWidth      =   2580
         TabIndex        =   58
         Top             =   360
         Width           =   2640
         Begin VB.CommandButton Command35 
            Caption         =   "O/C"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   59
            ToolTipText     =   "Acceso a Gestion de Compras Industriales"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Acceso a Escritorio  Gestion de Compras"
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
            Left            =   675
            TabIndex        =   60
            Top             =   120
            Width           =   2100
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Ordenes de Producción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   3960
      TabIndex        =   32
      Top             =   120
      Width           =   3690
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   960
         ScaleHeight     =   585
         ScaleWidth      =   2340
         TabIndex        =   43
         Top             =   1320
         Width           =   2400
         Begin VB.CommandButton Command50 
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
            Left            =   1825
            Picture         =   "MORFAB09.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   77
            ToolTipText     =   "Revertir Cierre de O/P"
            Top             =   0
            Width           =   525
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
            Height          =   600
            Left            =   0
            Picture         =   "MORFAB09.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Cierre de Orden de Fabricación - Re-Apertura O/F's con Saldo"
            Top             =   0
            Width           =   530
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cierre Orden de Producción"
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
            Left            =   540
            TabIndex        =   45
            Top             =   105
            Width           =   1260
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   600
         ScaleHeight     =   585
         ScaleWidth      =   2340
         TabIndex        =   46
         Top             =   840
         Width           =   2400
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
            Picture         =   "MORFAB09.frx":0D72
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Secuencia de Procesos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Secuencia de Procesos"
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
            TabIndex        =   48
            Top             =   30
            Width           =   1560
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   36
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command31 
            Caption         =   "Análisis A-B-C"
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
            Left            =   2205
            TabIndex        =   42
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command30 
            Caption         =   "Estadísticas"
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
            Left            =   210
            TabIndex        =   41
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command28 
            Caption         =   "Críticos"
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
            Left            =   2205
            TabIndex        =   40
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command27 
            Caption         =   "Consignación"
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
            Left            =   210
            TabIndex        =   39
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command26 
            Caption         =   "Negativos"
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
            Left            =   2205
            TabIndex        =   38
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command25 
            Caption         =   "Inventarios"
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
            Left            =   210
            TabIndex        =   37
            Top             =   420
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture10 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   240
         ScaleHeight     =   585
         ScaleWidth      =   2340
         TabIndex        =   33
         Top             =   375
         Width           =   2400
         Begin VB.CommandButton Command33 
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
            Left            =   1800
            Picture         =   "MORFAB09.frx":107C
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Ver y Re-Imprimir Orden de Produccion"
            Top             =   0
            Width           =   560
         End
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
            Height          =   600
            Left            =   0
            Picture         =   "MORFAB09.frx":1386
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Generar Orden de Fabricación"
            Top             =   0
            Width           =   530
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar Ord. Produccion"
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
            Left            =   600
            TabIndex        =   35
            Top             =   30
            Width           =   1320
         End
      End
      Begin VB.Image Image3 
         Height          =   360
         Left            =   3120
         Picture         =   "MORFAB09.frx":1690
         Stretch         =   -1  'True
         ToolTipText     =   "Doble Click para Emitir Listados"
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Programacion de Materiales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   3960
      TabIndex        =   23
      Top             =   2350
      Width           =   3690
      Begin VB.PictureBox Picture9 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   360
         ScaleHeight     =   585
         ScaleWidth      =   2910
         TabIndex        =   55
         Top             =   360
         Width           =   2970
         Begin VB.CommandButton Command48 
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
            Left            =   2360
            Picture         =   "MORFAB09.frx":199A
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Asigna/Des-Asigna Lotes a S/F"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command34 
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
            Picture         =   "MORFAB09.frx":1AE4
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Asigna y Reserva Materiales en Stock y Genera Necesidades de Compras"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Re-Calculo General de Necesidades"
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
            Left            =   550
            TabIndex        =   57
            Top             =   105
            Width           =   1860
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   24
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command15 
            Caption         =   "Análisis A-B-C"
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
            Left            =   2205
            TabIndex        =   30
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Estadísticas"
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
            Left            =   210
            TabIndex        =   29
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Críticos"
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
            Left            =   2205
            TabIndex        =   28
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Consignación"
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
            Left            =   210
            TabIndex        =   27
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Negativos"
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
            Left            =   2205
            TabIndex        =   26
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Inventarios"
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
            Left            =   210
            TabIndex        =   25
            Top             =   420
            Width           =   1590
         End
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00000080&
      Height          =   6840
      Left            =   7800
      ScaleHeight     =   6780
      ScaleWidth      =   1950
      TabIndex        =   8
      Top             =   -45
      Width           =   2010
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
         Left            =   105
         Picture         =   "MORFAB09.frx":1DEE
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "Central de Fraccionamiento / Consumo"
         Top             =   1545
         Width           =   645
      End
      Begin VB.CommandButton AGenRep 
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
         Left            =   105
         Picture         =   "MORFAB09.frx":20F8
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   3270
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   12
         Top             =   4080
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   12000
         End
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
         Picture         =   "MORFAB09.frx":2402
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   190
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
         Height          =   640
         Left            =   105
         Picture         =   "MORFAB09.frx":254C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   860
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
         Picture         =   "MORFAB09.frx":2856
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   2535
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -270
         Picture         =   "MORFAB09.frx":2C98
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4350
         Width           =   1515
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0E0F8&
      Caption         =   "Solicitudes de Fabricación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2950
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3690
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   435
         ScaleHeight     =   585
         ScaleWidth      =   2685
         TabIndex        =   51
         Top             =   2150
         Width           =   2750
         Begin VB.CommandButton Command32 
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   2150
            Picture         =   "MORFAB09.frx":4BBA
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Listados de Balance de Requerimientos"
            Top             =   0
            Width           =   560
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
            Left            =   0
            Picture         =   "MORFAB09.frx":4EC4
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Genera Lista de Requerimientos por Solicitud de Fabricacion"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Requerimientos de Materiales"
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
            Left            =   705
            TabIndex        =   53
            Top             =   105
            Width           =   1320
         End
      End
      Begin VB.PictureBox Picture17 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1080
         ScaleHeight     =   585
         ScaleWidth      =   2100
         TabIndex        =   18
         Top             =   1200
         Width           =   2160
         Begin VB.CommandButton Command45 
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
            Picture         =   "MORFAB09.frx":500E
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Anulación de Solicitud de Fabricación"
            Top             =   15
            Width           =   530
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Anulación de S/Fabricación"
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
            Left            =   630
            TabIndex        =   19
            Top             =   105
            Width           =   1320
         End
      End
      Begin VB.PictureBox Picture16 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   720
         ScaleHeight     =   585
         ScaleWidth      =   2100
         TabIndex        =   16
         Top             =   765
         Width           =   2160
         Begin VB.CommandButton Command46 
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
            Picture         =   "MORFAB09.frx":5318
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Modificacion de Datos de Solicitud de Fabricacion"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Modificacion S/F"
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
            Left            =   555
            TabIndex        =   17
            Top             =   105
            Width           =   1500
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   345
         ScaleHeight     =   585
         ScaleWidth      =   2100
         TabIndex        =   14
         Top             =   375
         Width           =   2160
         Begin VB.CommandButton Command41 
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
            Picture         =   "MORFAB09.frx":5462
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Ingresar Solicitud de Fabricación"
            Top             =   0
            Width           =   530
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Nueva S/F"
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
            Left            =   700
            TabIndex        =   15
            Top             =   105
            Width           =   1320
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   1
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command23 
            Caption         =   "Inventarios"
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
            Left            =   210
            TabIndex        =   7
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Negativos"
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
            Left            =   2205
            TabIndex        =   6
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Consignación"
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
            Left            =   210
            TabIndex        =   5
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Críticos"
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
            Left            =   2205
            TabIndex        =   4
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Estadísticas"
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
            Left            =   210
            TabIndex        =   3
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Análisis A-B-C"
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
            Left            =   2205
            TabIndex        =   2
            Top             =   1890
            Width           =   1590
         End
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   3675
         Y1              =   1980
         Y2              =   1980
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10395
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu Orfab 
      Caption         =   "Ordenes de Fabricacion"
      Visible         =   0   'False
      Begin VB.Menu NvaOrd 
         Caption         =   "Nueva Orden"
         Visible         =   0   'False
      End
      Begin VB.Menu php0 
         Caption         =   "-"
      End
      Begin VB.Menu LanOrd 
         Caption         =   "Lanzar Orden"
      End
      Begin VB.Menu Cerrar 
         Caption         =   "Cerrar Orden"
      End
      Begin VB.Menu php1 
         Caption         =   "-"
      End
      Begin VB.Menu SusOrd 
         Caption         =   "Suspensión de Orden"
      End
      Begin VB.Menu AnuOrd 
         Caption         =   "Anulación de Orden"
      End
      Begin VB.Menu ModOrd 
         Caption         =   "Modificación de Orden"
      End
      Begin VB.Menu php2 
         Caption         =   "-"
      End
      Begin VB.Menu VerImp 
         Caption         =   "Ver y Re-Imprimir"
      End
   End
   Begin VB.Menu Edicion 
      Caption         =   "Edición"
      Visible         =   0   'False
      Begin VB.Menu SuOrd 
         Caption         =   "Suspender Orden"
      End
      Begin VB.Menu AnulOrd 
         Caption         =   "Anular Orden"
      End
      Begin VB.Menu ModiOrd 
         Caption         =   "Modificar Orden"
      End
      Begin VB.Menu php3 
         Caption         =   "-"
      End
      Begin VB.Menu VeryReimpi 
         Caption         =   "Ver y Re-Imprimir"
      End
   End
   Begin VB.Menu AsiMaPri 
      Caption         =   "Materiales"
      Visible         =   0   'False
      Begin VB.Menu MatReser 
         Caption         =   "Reserva de Materiales"
         Begin VB.Menu GenReservOC 
            Caption         =   "Generar Reservas por Orden de Fabricación"
         End
         Begin VB.Menu CanceReser 
            Caption         =   "Cancelación de Reservas"
         End
         Begin VB.Menu ConOrFaCu 
            Caption         =   "Control Segun Ordenes de Fabricación en Curso"
         End
         Begin VB.Menu DepuBdD 
            Caption         =   "Depuración de Base de Datos"
         End
         Begin VB.Menu php3a 
            Caption         =   "-"
         End
         Begin VB.Menu LiToMat 
            Caption         =   "Listado de Totales por Material"
            Begin VB.Menu LiToPan 
               Caption         =   "por Pantalla"
            End
            Begin VB.Menu LiToImpre 
               Caption         =   "por Impresión"
            End
         End
         Begin VB.Menu CoGrlRe 
            Caption         =   "Consulta General de Reservas por Orden de Fabricación"
         End
         Begin VB.Menu ReXNumOF 
            Caption         =   "Reservas por Numero de Orden de Fabricación"
         End
         Begin VB.Menu ReXMat 
            Caption         =   "Reservas por Material"
         End
      End
      Begin VB.Menu ValMate 
         Caption         =   "Vale de Materiales"
         Begin VB.Menu ValPrep 
            Caption         =   "Vale de Preparacion"
         End
         Begin VB.Menu ValTran 
            Caption         =   "Vale de Transferencia"
         End
      End
   End
   Begin VB.Menu AcDirec 
      Caption         =   "Accesos Directos"
      Visible         =   0   'False
      Begin VB.Menu InfoPro 
         Caption         =   "Informes de Producción"
      End
      Begin VB.Menu CarMaq 
         Caption         =   "Carga de Maquinas"
      End
      Begin VB.Menu php15 
         Caption         =   "-"
      End
      Begin VB.Menu MaestArt 
         Caption         =   "Maestro de Articulos"
      End
      Begin VB.Menu php16 
         Caption         =   "-"
      End
      Begin VB.Menu otraccdir 
         Caption         =   "Otros Accesos Directos"
      End
   End
   Begin VB.Menu ConsuLis 
      Caption         =   "Consultas y Listados"
      Visible         =   0   'False
      Begin VB.Menu ConGrBdD 
         Caption         =   "Consulta General de la Base de Datos"
      End
      Begin VB.Menu php4 
         Caption         =   "-"
      End
      Begin VB.Menu PenSecPro 
         Caption         =   "Pendientes por Sector Productivo"
         Begin VB.Menu Xpanta 
            Caption         =   "por Pantalla"
         End
         Begin VB.Menu Ximpres 
            Caption         =   "por Impresión"
         End
      End
      Begin VB.Menu PenNumOF 
         Caption         =   "Pendientes por Numero de Orden de Fabricación"
      End
      Begin VB.Menu PenCodPro 
         Caption         =   "Pendientes por Código de Producto"
         Begin VB.Menu porPantall 
            Caption         =   "por Pantalla"
         End
         Begin VB.Menu Ximpreso 
            Caption         =   "por Impresión"
         End
      End
      Begin VB.Menu php5 
         Caption         =   "-"
      End
      Begin VB.Menu ConSelEst 
         Caption         =   "Consulta Selectiva por Estado"
         Begin VB.Menu Xpant 
            Caption         =   "por Pantalla"
         End
         Begin VB.Menu Ximpre 
            Caption         =   "por Impresión"
         End
      End
      Begin VB.Menu php6 
         Caption         =   "-"
      End
      Begin VB.Menu VeryReimp 
         Caption         =   "Ver y Re-Imprimir"
      End
      Begin VB.Menu php24 
         Caption         =   "-"
      End
      Begin VB.Menu ProCar 
         Caption         =   "Programación de Carga"
         Begin VB.Menu OperPro 
            Caption         =   "Operaciones Programadas"
         End
         Begin VB.Menu CaToEqui 
            Caption         =   "Carga Total por Equipo"
         End
      End
   End
   Begin VB.Menu ConfGral 
      Caption         =   "Configuración"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "MORFAB09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CONCUMORFAB()
  '27/04/07 (OT-07-0172)
  '
  Call DESHASFECHA(DES%, HAS%, PERIO$)
  If PERIO$ = "" Then Exit Sub
  DESDEX = CVDAT(DES%)
  HASTAX = CVDAT(HAS%)
  '
  Call ABREORFAB
  '
  SQLX$ = " SELECT * FROM ORDEFABR "
  SQLX$ = SQLX$ + " WHERE FechGen >= " & DESDEX & " "
  SQLX$ = SQLX$ + " AND FechGen <= " & HASTAX & " "
  SQLX$ = SQLX$ + " AND StatuOrf <> 9 "
  SQLX$ = SQLX$ + " ORDER BY IdSubOr ASC"
  '
  'Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
  Dim ORDEFABTMP As ADODB.Recordset
  Set ORDEFABTMP = New ADODB.Recordset
  ORDEFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
  '
  With ORDEFABTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
       On Error GoTo 0
       Screen.MousePointer = 1
       Call MENSERR(24, "Ningun Registro Coincide con Registro SQL")
       Exit Sub
    End If
    '
    I& = 0
    Do While Not .EOF
       Y$ = REGBLAN("COCUMOF")
       FEMIS% = CVINT(!FechGen)
       Call REPLA(Y$, MKI$(FEMIS%), 1, 2)
       Call REPLA(Y$, !IDSUBOR, 3, 12)
       Call REPLA(Y$, MKI$(!Cod_Inte), 15, 2)
       Call REPLA(Y$, MKD$(!CanProy), 17, 8)
       Call REPLA(Y$, !DESTINO, 25, 8)
       FECHAENT% = CVINT(!FechEnSol)
       Call REPLA(Y$, MKI$(FECHAENT%), 33, 2)
       FECHACIE% = CVINT(!FechCierr)
       Call REPLA(Y$, MKI$(FECHACIE%), 35, 2)
       DIATRAZO% = DIENTRE%(FECHAENT%, FECHACIE%)
       Call REPLA(Y$, MKI$(DIATRAZO%), 37, 2)
       '
       I& = I& + 1
       Call GRAREG("COCUMOF", Y$, I&)
    .MoveNext
    Loop
    '
  End With
  Set ORDEFABTMP = Nothing
  '
  Call SETULTREG("COCUMOF", I&)
  '
  Call HEADERS("COCUMOF", "")
  Call HEADERS("COCUMOF", "Período: " + PERIO$)
  Call FINAL("*COCUMOF")
  '
  '27/04/07 (FIN)
End Sub

Sub REIMPRIETIQUETA()
    '
    ' ///OT-8-1037-12/11/08
    Call SELECHO("MASTER")
    CODEX$ = VALACT1$("MASTER")
    CIXI% = CODINT%(CODEX$)
    If CIXI% < 1 Then Exit Sub
    If ESTRAZABLE%(CIXI%) < 1 Then Exit Sub
    '
    Call ARMALILOT(CIXI%, DEPO%, FAJUI%)
    '
    Call FRESHECHO("!LILOTSE")
    If ULTREG("!LILOTSE") < 1 Then
       Call MENSERR(24, "No Hay Lotes Disponibles\para la Operacion Solicitada.")
       Exit Sub
    End If
    '
    Call SELECHO("!LILOTSE")
    VDEVU$ = VALACT1$("!LILOTSE")
    IDREGIS$ = VDEVU$
    '
    CONDI$ = "IDENLOTE='" & TRIM$(IDREGIS$) & "'"
    DPX% = XVALO(VALOBADA("LOTINGRE", "DepoUbic", CONDI$))
    If DPX% > 0 Then Call PRINSTIKAPRO(CIXI%, IDREGIS$, DPX%)
    ' ///OT-8-1037-FIN
    '
End Sub

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

Private Sub AnulOrd_Click()
Call Command45_Click
End Sub

Private Sub AnuOrd_Click()
Call Command45_Click
End Sub

Private Sub CanceReser_Click()
Call RESERV07.Command4_Click
End Sub





Private Sub Cerrar_Click()
Call Command46_Click
End Sub

Private Sub CoGrlRe_Click()
'
RESERV07.Option1 = True
RESERV07.Option5 = True
Call RESERV07.Command2_Click
'

End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub



Private Sub Command14_Click()
   '
   Command14.Enabled = False
   If DERACCE%("STPOFAB", "Configuracion de Funcionamiento de O/Fabricacion") = 2 Then
      OPORFA09.Show 1
   End If
   Command14.Enabled = True
   '
End Sub

Private Sub Command19_Click()
   Command19.Enabled = False
   '
   Call ABREORFAB
   FLEXCONN.Execute (FILTSQL$("UPDATE ORDEFABR SET NuOrProd = '' where NuOrProd IS NULL"))
   Call ACTUNUORPRO        ' ajusta el numero correlativo a 4 posiciones
   '
5  CONDI$ = "IDSUBOR <> '' "
   CONDI$ = CONDI$ + " AND STATUORF < 3 "
   CONDI$ = CONDI$ + " AND CanApro < CanProy "
   CONDI$ = CONDI$ + " AND (NuOrProd = '' OR NuOrProd IS NULL) "
   CONDI$ = CONDI$ + " ORDER BY IDSUBOR ASC"
   Call CARGALISEL("!ECORDEF", "ORDEFABR", "IDSUBOR/A12;REFEREN/A64", CONDI$)
   Call FRESHECHO("!ECORDEF")
   '
10 Call SELECHO("!ECORDEF")
   NORFA$ = VALACT1$("!ECORDEF")
   If NORFA$ = "" Then GoTo 99
   '
   FFEE% = HOY(HOS$)
   '
   CONDI$ = "IdSubOr ='" & NORFA$ & "'"
   '
   LOTE$ = VALOBADA("ORDEFABR", "IdenLote", CONDI$)
   NUMOP$ = VALOBADA("ORDEFABR", "DESTINO", CONDI$)
   CLI% = XVALO(VALOBADA("ORDEFABR", "Nume_Cli", CONDI$))
   CCII% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", CONDI$))
   '
   If LOTE$ = "" Then
      Call MENSERR(24, "Solicitud de Fabricacion Sin Lote Asignado")
      GoTo 10
   End If
   '
   If NUMOP$ = "" Then
      Call MENSERR(24, "Solicitud de Fabricacion Sin O-P Asignada")
      GoTo 10
   End If
   '
   If CLI% < 1 Then
      Call MENSERR(24, "Solicitud de Fabricacion Sin Cliente Asignado")
      GoTo 10
   End If
   '
   CONDI$ = "CodIConj = " & CCII%
   If CantRegistros("FORMUPAT", CONDI$) <= 1 Then
      Call MENSERR(24, "Codigo de Articulo a Producir Debe Tener Formula Patron")
      GoTo 10
   End If
   '
   If FEREFOR%(CCII%) <= FECHANUM("01/01/80") Then
      Call MENSERR(24, "Imposible Continuar.\Formula Patron No Aprobada")
      GoTo 10
   End If
   '
   CONDI$ = "Destino = '" & NORFA$ & "' AND Kilo_Reque > 0"
   QTR% = CantRegistros("BALAREQMO", CONDI$)
   If QTR% < 1 Then
      Call MENSERR(24, "No Hay Materiales Requeridos\para la Solicitud de Fabricacion '" & NORFA$ & "'")
      GoTo 10
   End If
   '
   If NORFA$ <> "" Then
      CONDI$ = "Destino = '" & NORFA$ & "'"
      SQLX$ = "SELECT IDENREQUE,Cod_Inte,Sum(Kilo_Reque) as Sureque,Sum(Kilo_Asig) as Suasig FROM BALAREQMO WITH(NOLOCK)"
      SQLX$ = SQLX$ + " WHERE " & CONDI$
      SQLX$ = SQLX$ + " GROUP BY IDENREQUE,COD_INTE"
      '
      Dim BALRE As ADODB.Recordset
      Set BALRE = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With BALRE
        Do While Not .EOF
          CIXI% = XVALO(!Cod_Inte)
          If CRITREP%(CIXI%) <> 3 Then
            If (XVALO(!SUREQUE) - XVALO(!SUASIG)) > FRACMIN(CIXI%) Then
              TTXX$ = "Imposible Lanzamiento de la Orden de Producción\para la Solicitud de Fabricación '" + NORFA$ + "'."
              TTXX$ = TTXX$ + "\   \Falta Asignación de Materiales para '" + CODEXT$(CIXI%) + "'\(" + DESCRIT0$(CIXI%) + ")."
              Call MENSERR(24, TTXX$)
              GoTo 10
            End If
          End If
        .MoveNext
        Loop
      End With
      BALRE.Close
      Set BALRE = Nothing
      ' aqui se verificó el balance general
      '
      FALMAT% = 0
      CONDI$ = "Destino = '" & NORFA$ & "'"
      CONDI$ = CONDI$ + " AND Kilo_Asig > 0 "
      SQLX$ = "SELECT * FROM BALAREQMO WITH(NOLOCK)"
      SQLX$ = SQLX$ + " WHERE " & CONDI$
      SQLX$ = SQLX$ + " ORDER BY IDENREQUE ASC"
      Set BALRE = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With BALRE
        Do While Not .EOF
          CIXI% = XVALO(!Cod_Inte)
          If CRITREP%(CIXI%) = 3 Then GoTo 19
            IDLO$ = SAFETEXT$(!idenLote)
            CANASI = XVALO(!Kilo_Asig)
               CONDX$ = "Cod_Inte = " & CIXI%
               CONDX$ = CONDX$ + " AND IdenLote = '" & IDLO$ & "'"
               CONDX$ = CONDX$ + " AND LIBERA = 'APRO'"
            SALOT = 0
            QTRX& = CantRegistros("LOTINGRE", CONDX$)
            If QTRX& > 0 Then
               Call ACSALOTE(CIXI%, IDLO$)
               SALOT = XVALO(VALOBADA("LOTINGRE", "CanAlta", CONDX$))
               SALOT = SALOT - XVALO(VALOBADA("LOTINGRE", "CanConsu", CONDX$))
            End If
            If SALOT < CANASI Then
               FALMAT% = 1
               TTXX$ = "Imposible Lanzamiento de la Orden de Producción\para la Solicitud de Fabricación '" + NORFA$ + "'."
               TTXX$ = TTXX$ + "\   \Lote Asignado '" + IDLO$ + "' para '" + CODEXT$(CIXI%) + "'\(" + DESCRIT0$(CIXI%) + ")\Inexistente, Insuficiente o No Liberado (en Cuarentena)."
               OPX% = COMALTER%(TTXX$ + "\\Continuar\Cancelar")
               If OPX% = 2 Then GoTo 25
            End If
19      .MoveNext
        Loop
      End With
      '
25    BALRE.Close
      Set BALRE = Nothing
      If FALMAT% < 1 Then
         '
         'Obtiene el proximo numero de la orden de produccion segun la fecha
         NORPRO$ = NUMIDEORPROD$(FFEE%)
         '
         CONDI$ = "IDSUBOR = '" & NORFA$ & "' "
         Call ACTUREGISTRO("ORDEFABR", "NuOrProd", CONDI$, NORPRO$, RAFE&)
         Call COMUNI("Solicitud de Fabricación '" + NORFA$ + "'\fue Lanzada como Orden de Produccion '" + NORPRO$ + "'.")
         '
         CONDI$ = "NuOrProd = '" & NORFA$ & "'"
         SOLFAB$ = VALOBADA("ORDEFABR", "IDSUBOR", CONDI$)
         REMIRECE$ = TOTREMITREC$(SOLFAB$)
         '
         FILTX$ = TRIM$(NORPRO$) & ";" & TRIM$(REMIRECE$)
         '
         Call STDFORM("BR-ALE-X", FILTX$)
         If STDSCREEN.APRO <> "OK" Then
            GoTo 99
         End If
         '
         GoTo 5
         '
      End If
      GoTo 10
      '
   End If
   '
99 Command19.Enabled = True
   '
End Sub

Private Sub Command2_Click()
   '
   Command2.Enabled = False
   OPP% = COMALTER%("Opciones de Operación:\\Secuencia Automatica\Paso a Paso\Imprimir Batch Record")
   If OPP% < 1 Or OPP% > 3 Then GoTo 99
   '
   ' para poder referenciar la O/Produccion anterior
   CONDI$ = "NUORPROD <> '' order by NuOrProd "
   Call CARGALISEL("!ECORDEF", "ORDEFABR", "NUORPROD/A12;REFEREN/A64", CONDI$)
   Call FRESHECHO("!ECORDEF")
   '
   Call ABRECONEXION
   If OPP% = 2 Then
      If DERACCE%("BRECPASO", "Batch Record Paso a Paso") < 2 Then GoTo 99
      SECUPROC09.Show 1
      GoTo 99
   End If
   '
   CONDI$ = "StatuOrf < 3 "   ' para que no este anulada ni cumplida
      If OPP% = 3 Then CONDI$ = "StatuOrf <= 3 "   ' para que en la impresion aparezcan las cumplidas
   CONDI$ = CONDI$ + "AND CanApro < (CanProy - 0.005) "
   CONDI$ = CONDI$ + " and NuOrProd <> '' "
   CONDI$ = CONDI$ + " ORDER BY NuOrProd "
   '
   Call CARGALISEL("ORDPEN", "ORDEFABR", "NUORPROD/A16;REFEREN/A48", CONDI$, "DISTINCT")
   Call SELECHO("ORDPEN/Indice de Ordenes de Produccion Pendientes")
   VDEVU$ = VALACT1$("ORDPEN")
   If VDEVU$ <> "" Then
      SECUPROC09.Label32 = VDEVU$
      AUTOPAGE% = 1
      '
      If SECUPROC09.List1.ListCount > 0 Then
         SECUPROC09.List1.ListIndex = 0
         If OPP% = 2 Then
              If DERACCE%("BRECAUTO", "Batch Record Secuencia Automatica") < 2 Then GoTo 99
              Call SECUPROC09.List1_DblClick
           ElseIf OPP% = 3 Then
              If DERACCE%("BRECPRIN", "Impresion de Batch Record") < 2 Then GoTo 99
              Call SECUPROC09.Command2_Click
         End If
      End If
      AUTOPAGE% = 0
      '
   End If
   '
99 Command2.Enabled = True
   '
End Sub

Private Sub Command24_Click()
    '
    Command24.Enabled = False
    '
10  Call COPYSTRU("ECORDEF", "ECORPEN")
    CONDI$ = "IDSUBOR <> '' AND STATUORF < 2 AND (NUORPROD = '' OR NUORPROD IS NULL)ORDER BY IDSUBOR ASC"
    Call CARGALISEL("!ECORPEN", "ORDEFABR", "IDSUBOR/A12;REFEREN/A64", CONDI$)
    '
    Call SELECHO("!ECORPEN")
    NORFA$ = TRIM$(VALACT1$("!ECORPEN"))
    DESTIINDI$ = TRIM$(VALACT2$("!ECORPEN"))
    If NORFA$ <> "" Then
       TTXX$ = "¿ Confirma Refrescar Lista de Requerimientos\de la Solicitud de Fabricacion '" + NORFA$ + "' ?"
       If COMALTER%(TTXX$ + "\\No, Continuar\Si, Refrescar") = 2 Then
          Call GENERALISTAREQ(NORFA$, "REFRESH")
          Call ASIGNAMAT(NORFA$)
       End If
    End If
    '
90  Screen.MousePointer = 1
    Command24.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call ACTUNUORPRO
    ' CIERRAOF09.Show 1
    CIERREOP10.Show 1
    GoTo 99
    '
    '
    OPX% = COMALTER%("Opciones de Trabajo con O/F's:\\Cierre de O/F's\Re-Apertura\Anular Cierre de O/F")
    If OPX% = 1 Then
       If EXISTE%("TRAZAB07.EXE") > 0 Then
          If DERACCE%("CIERROFT", "Cierre de O/F's con Trazabilidad") = 2 Then
             Call CONECRUN("TRAZAB07/ORFCIE", "Cierre de O/F's con Trazabilidad")
          End If
          GoTo 99
        End If
        Call ORFCIE
      ElseIf OPX% = 2 Then
        If DERACCE%("REAPEROF", "Re-Apertura O/F's con Saldo") >= 2 Then
          Call REAPEROF
        End If
      ElseIf OPX% = 3 Then
        If DERACCE%("ANUCIEOF", "Anulación Cierre Parcial de O/F") >= 2 Then
          Call ANUCIEOF
        End If
    End If
    
99    Command3.Enabled = True

End Sub

Private Sub Command29_Click()
    Call HELPONLINE("ORDEFABR")
End Sub


Private Sub Command4_Click()
File1.Path = LUDAT$
File1.Pattern = "*.sqr"
'
For I% = 1 To File1.ListCount
TTXX$ = LOADFILE$(LUDAT$ & File1.List(I% - 1))
HUBOCA% = 0
'
10 PPXX& = InStr(TTXX$, "BLIPACK")
   If PPXX& > 0 Then
     Mid$(TTXX$, PPXX&, 7) = "PHARMAX"
     HUBOCA% = 1
     GoTo 10
   End If
   '
   If HUBOCA% > 0 Then
      Call SAVEFILE(LUDAT$ & File1.List(I% - 1), TTXX$)
   End If
Next I%
'
End Sub

Private Sub Command42_Click()
    Command42.Enabled = False
    Call ORFLAN
    Command42.Enabled = True
End Sub

Private Sub Command32_Click()
   Command32.Enabled = False
   '
   ' LISTADOS DE REQUERIMIENTOS
   Call CARGALISEL("!ECORDEF", "ORDEFABR", "IDSUBOR/A12;REFEREN/A64", "IDSUBOR <> '' AND STATUORF < 3 AND CanApro < CanProy ORDER BY IDSUBOR ASC")
   Call FRESHECHO("!ECORDEF")
   '
   Call CIERRARCH("*.*")
10 Call SELECHO("!ECORDEF")
   NLIPIX$ = VALACT1$("!ECORDEF")
   '
   ' OJO!!! SI SE MODIFICA ESTA RUTINA(RECUOCPEN), TAMBIEN SE DEBE ACTUALIZAR EN
   ' EN EL PROYECTO DECAOT09, FORMULARIO MENCAOT09.
   Call RECUOCPEN(NLIPIX$)
   '
   If TRIM$(NLIPIX$) <> "" Then
25      OPX% = COMALTER%("Opciones de Listados de Requerimientos:\\Listado Plano\Balance de Detalle\Balance Acumulado")
        Select Case OPX%
          Case 1
            Call FINAL("?LISREQUERI(" + TRIM$(NLIPIX$) + "%)")
            GoTo 25
          Case 2
            Call FINAL("?BALREQUERI(" + TRIM$(NLIPIX$) + "%)")
            GoTo 25
          Case 3
            Call FINAL("?ACUREQUERI(" + TRIM$(NLIPIX$) + "%)")
            GoTo 25
        End Select
        GoTo 10
   End If
   '
   Command32.Enabled = True
End Sub

Private Sub Command33_Click()
   '
   Command33.Enabled = False
   '
   CONDI$ = "NUORPROD <> '' order by NuOrProd "
   Call CARGALISEL("!ECORDEF", "ORDEFABR", "NUORPROD/A12;REFEREN/A64", CONDI$)
   Call FRESHECHO("!ECORDEF")
   '
   Call SELECHO("!ECORDEF")
   NORFA$ = VALACT1$("!ECORDEF")
   If NORFA$ = "" Then GoTo 99
   '
   CONDI$ = "NuOrProd = '" & NORFA$ & "'"
   SOLFAB$ = VALOBADA("ORDEFABR", "IDSUBOR", CONDI$)
   REMIRECE$ = TOTREMITREC$(SOLFAB$)
   '
   FILTX$ = TRIM$(NORFA$) & ";" & TRIM$(REMIRECE$)
   '
   Call STDFORM("BR-ALE-X", FILTX$)
   If STDSCREEN.APRO <> "OK" Then
      GoTo 99
   End If
   '
99 Command33.Enabled = True
   '
End Sub

Private Sub Command34_Click()
    Command34.Enabled = False
    RECALMAT09.Show 1
    Command34.Enabled = True
End Sub

Private Sub Command35_Click()
    Command35.Enabled = False
    Call CONECRUN("MORCOM09", "Escritorio de Gestion de Compras")
    Command35.Enabled = True
End Sub

Private Sub Command41_Click()
    Command41.Enabled = False
    If DERACCE%("GENORFA09", "Ordenes de Fabricacion") >= 2 Then
       GENORFA09.Show 1
    End If
    Command41.Enabled = True
End Sub

Private Sub Command50_Click()
   '
   If DERACCE%("RECIEOP", "Revertir Cierre de O/P") >= 2 Then
     CONDI$ = "StatuOrf = 3 "   ' para que no este anulada ni cumplida
     CONDI$ = CONDI$ + " and NuOrProd <> '' "
     CONDI$ = CONDI$ + " ORDER BY NuOrProd "
    
     Call CARGALISEL("ORDCER", "ORDEFABR", "NUORPROD/A16;REFEREN/A48", CONDI$, "DISTINCT")
     SELECHO ("ORDCER/Indice de Ordenes de Produccion Cerradas")
     NOP$ = VALACT1$("ORDCER")
     '
     If NOP$ <> "" Then
       CONDI$ = "NUORPROD = '" & NOP$ & "'"
       Call ACTUREGISTRO("ORDEFABR", "STATUORF", CONDI$, 1, REG&)
       Call COMUNI("OP " & NOP$ & " Re-abierta Exitosamente")
     End If
   End If
   '
End Sub

Private Sub Command6_Click()
   '
   Screen.MousePointer = 11
   '
   Call CONECRUN("MDESMA09/REIMPRE", "")
   '
   Screen.MousePointer = 1
99 Command6.Enabled = True

End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   '
   CENTRAPES09.Show 1
   '
   Command7.Enabled = True
End Sub

Private Sub Command71_Click()
   Command71.Enabled = False
   '
   Call CONECRUN("MDESMA09/REMITO", "")
   '
99 Screen.MousePointer = 1
   Command71.Enabled = True
End Sub

Private Sub Command45_Click()
    Command45.Enabled = False
    If DERACCE%("ANULASF", "Anulacion de S/F") >= 2 Then
       Call ORFANU
       Call GENLISORF
    End If
99  Command45.Enabled = True
End Sub

Private Sub Command46_Click()
   Command46.Enabled = False
   '
   If DERACCE%("MODIOF", "Modificacion de Datos Solicitud de Fabricacion") < 2 Then
      GoTo 99
   End If
   '
   'GENERACION DE ECORDEF
   Call ABRECONEXION
   CONDI$ = "IDSUBOR <> '' "
   CONDI$ = CONDI$ + " AND (nuorprod ='' or nuorprod is null) "
   CONDI$ = CONDI$ + " AND STATUORF < 3 "
   CONDI$ = CONDI$ + " AND CanApro < CanProy ORDER BY IDSUBOR ASC"
   Call CARGALISEL("!ECORDEF", "ORDEFABR", "IDSUBOR/A12;REFEREN/A64", CONDI$)
   Call FRESHECHO("!ECORDEF")
   Call CIERRARCH("*.*")
   '
10 Call SELECHO("!ECORDEF")
   OF$ = VALACT1("!ECORDEF")
   If OF$ = "" Then GoTo 99
   
   DATOS$ = String$(150, 0)
   CONDI$ = "IdSubOr ='" & OF$ & "'"
   '
   LOTE$ = VALOBADA("ORDEFABR", "IdenLote", CONDI$)
   NUMOP$ = VALOBADA("ORDEFABR", "DESTINO", CONDI$)
   CAN# = XVALO(VALOBADA("ORDEFABR", "CanProy", CONDI$))
   REMI$ = VALOBADA("ORDEFABR", "NumRem_Prov", CONDI$)
   OCOM$ = VALOBADA("ORDEFABR", "Nume_Ocom", CONDI$)
   CODITEM% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", CONDI$))
   CLI% = XVALO(VALOBADA("ORDEFABR", "Nume_Cli", CONDI$))
   FESOLE% = CVINT(VALOBADA("ORDEFABR", "FechEnSol", CONDI$))
   '
   Call REPLA(DATOS$, OF$, 1, 16)
   Call REPLA(DATOS$, NUMOP$, 17, 16)
   Call REPLA(DATOS$, LOTE$, 33, 16)
   Call REPLA(DATOS$, MKD(CAN#), 49, 8)
   Call REPLA(DATOS$, MKI$(CLI%), 57, 2)
   Call REPLA(DATOS$, REMI$, 59, 16)
   Call REPLA(DATOS$, OCOM$, 75, 16)
   Call REPLA(DATOS$, MKI$(CODITEM%), 91, 2)
   Call REPLA(DATOS$, MKI$(FESOLE%), 93, 2)
   '
20 RESULT$ = OBJPLA$("MODIOF", DATOS$)
   If RESULT$ = "" Then GoTo 10
   '
   DATOS$ = RESULT$
   'TOMAR DATOS DEL OBJETO
   NUMOP$ = Mid(RESULT$, 17, 16)
   LOTE$ = Mid(RESULT$, 33, 16)
   CAN# = CVD(Mid(RESULT$, 49, 8))
   CLI% = CVI(Mid(RESULT$, 57, 2))
   REMI$ = Mid(RESULT$, 59, 16)
   OCOM$ = Mid(RESULT$, 75, 16)
   CODITEM% = CVI(Mid$(RESULT$, 91, 2))
   FESOLE% = CVI(Mid$(RESULT$, 93, 2))
   '
   CONDI1$ = "Destino ='" & TRIM$(NUMOP$) & "' AND IdSubOr <> '" & OF$ & "'"
   If TRIM(NUMOP$) <> "" Then
     If CantRegistros("ORDEFABR", CONDI1$) > 0 Then
        OPT1% = COMALTER%("Nro de O.P " & TRIM$(NUMOP$) & " Existente. Desea Continuar?\\No, Cancelar\Si, Continuar")
        If OPT1% <> 2 Then GoTo 20
     End If
   End If
   '
   'VALIDA QUE EL ARTICULO NO ESTE DADO DE BAJA
   STAT% = STATUIT%(CODITEM%)
   If STAT% < 1 Then
      COMUNI ("Imposible Continuar. Codigo De La S/F Se Encuentra Dado de Baja")
      GoTo 20
   End If
   '
   'VALIDACION QUE NO DIFIERA CANTIDAD DEL LOTE STANDARD
   LOTSTD# = LOTESTAN(CODITEM%)
   If Abs(CAN# - LOTSTD#) > XVALO(CNTDC$(CODITEM%)) Then
      COMUNI ("La Cantidad No Coincide Con La Registrada Como Lote Stantard (" & LOTSTD & ")")
      GoTo 20
   End If
   '
   If CLI% < 1 Or RASOCLI$(CLI%) = "" Then
     Call MENSERR(24, "Cliente No Valido")
     GoTo 20
   End If
   '
   If TRIM$(OCOM$) = "" Then
     COMUNI ("Falta Ingresar Nro de O/Compra")
     GoTo 20
   End If
   '
   'MODIFICACION DE REFERENCIA
   CI% = CODITEM%
   REFNUE$ = "OP." & TRIM$(NUMOP$)
   REFNUE$ = REFNUE$ & " - L." & TRIM$(LOTE$)
   REFNUE$ = REFNUE$ & " - " & FANTCLI$(CLI%)
   REFNUE$ = REFNUE$ & " - " & DESCRIT0$(CI%)
   '
   Call ACTUREGISTRO("ORDEFABR", "DESTINO", CONDI$, NUMOP$, REG&)
   Call ACTUREGISTRO("ORDEFABR", "IDENLOTE", CONDI$, LOTE$, REG&)
   Call ACTUREGISTRO("ORDEFABR", "NumRem_Prov", CONDI$, REMI$, REG&)
   Call ACTUREGISTRO("ORDEFABR", "Nume_Ocom", CONDI$, OCOM$, REG&)
   Call ACTUREGISTRO("ORDEFABR", "REFEREN", CONDI$, Left(REFNUE$, 64), REG&)
   Call ACTUREGISTRO("ORDEFABR", "Nume_Cli", CONDI$, CLI%, REG&)
   Call ACTUREGISTRO("ORDEFABR", "FechEnSol", CONDI$, FETEXCOR1$(FESOLE%), REG&)
   '
   CANTIANT# = XVALO(VALOBADA("ORDEFABR", "CanProy", CONDI$))
   If Abs(CAN# - CANTIANT#) >= 0.5 Then
     If DERACCE%("MODCANOF", "Modificacion Cantidad Proyectada de S/Fabric.") >= 2 Then
       TTXX$ = "¿ Realmente Desea Modificar la Cantidad\de esta Solicitud de Fabricación ?"
       OPXX% = COMALTER%(TTXX$ + "\\No, Mantener la Anterior\Si, Modificar Cantidad")
       If OPXX% = 2 Then
         Screen.MousePointer = 11
         CAN1$ = FORMATNUM$(CAN#, "F13." + CNTDC$)
         Call ACTUREGISTRO("ORDEFABR", "CanProy", CONDI$, CAN1$, REG&)
         Call GENERALISTAREQ(OF$, "REFRESH")
         Call ASIGNAMAT(OF$)
         Screen.MousePointer = 1
       End If
     End If
   End If
   '
   Call GENLISORF
   Screen.MousePointer = 1
   COMUNI ("Modificación Registrada")
   '
99 Command46.Enabled = True
   '
End Sub


Private Sub ConfGral_Click()
Call Command14_Click
End Sub



Private Sub ConOrFaCu_Click()
Call RESERV07.Command7_Click
End Sub

Private Sub DepuBdD_Click()
Call RESERV07.Command5_Click
End Sub

Private Sub Form_Activate()
   WindowState = 0
   Openforms = DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Openforms = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    'Show
    'APLIXYZ$ = TRIM$(Command$)
    '
    VERANTER$ = "ORDEFABR"
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    'Esta actualizacion evita el error de Valor Nulo cuando se lee el campo Priorid
    Call ABRECONEXION
    Call ACTUREGISTRO("ORDEFABR", "Priorid", "Priorid IS NULL", "0", REG&)
    '
    ' CAMBIA BARRA POR GUION EN LAS ORDENES DE BLIPACK
    Call AJUFORPROD_BLI
    '
    If APLINVO$ = "OPRESERV" Then
      RESERV07.Show 1
      Call FINAL("")
    End If
    '
    If UCase$(Left$(APLINVO$, 6)) = "ORFNUE" Then
       APLIXYZ$ = APLINVO$
    End If
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Call GRATITFOR(Caption)
    '
    '
    'CODIGO PARA COLGAR MODULO
    FF% = FECHANUM("30/10/11")
    If FF% <= HOY(HO$) Then
        Call Form_Load
    End If
    '
    If InStr(EPAC$, "BOREN") > 0 Then
       Command12.Visible = True
    End If
    '
    Call VERIDEST
    If InStr(UCase$(EPAC$), "BORENS") > 0 Then
       Picture1.Width = Picture1.Width + Command4.Width
       Picture1.Left = Picture1.Left - Command4.Width / 2
    End If
    '
'    If EXISTE%("REPARA02.EXE") < 1 Then
'      If EXISTE%("REPABO02.EXE") < 1 Then
'        Command3.Enabled = False
'        Label5.Enabled = False
'      End If
'    End If
    '
    Openforms = DoEvents
    Show
    
    '
    ' ESTO ES PARA QUE GENERE AUTOMÁTICAMENTE UNA O/F
    If APLIXYZ$ <> "" Then
       If UCase$(Left$(APLIXYZ$, 6)) = "ORFNUE" Then
          AppActivate Caption
          SendKeys Chr$(1)
          Call Command41_Click
          Exit Sub
       End If
    End If
    '
End Sub
'
Sub AJUFORPROD_BLI()
   '
   ' ESTA RUTINA CAMBIA EL FORMATO nnnn/nnnn DE BLIPACK POR nnnn-nnnn PARA EVITAR EL ERROR DE GRABAR
   ' EL NUMERO TRUNCADO EN EL BATCH RECORD.
   '
   Call ABREORFAB
   CONDI$ = "NUORPROD LIKE '%/%'"
   If CantRegistros("ORDEFABR", CONDI$) > 0 Then
      Dim ORFA As ADODB.Recordset
      Set ORFA = New ADODB.Recordset
      SQLX$ = "SELECT * FROM ORDEFABR WHERE " + CONDI$
      Set ORFA = FLEXCONN.Execute(SQLX$)
      With ORFA
         Do While Not .EOF
            IDOR$ = SAFETEXT$(!NUORPROD)
            NORPRO$ = REPLACAR$(IDOR$, "/", "-")
            Call ACTUREGISTRO("ORDEFABR", "NUORPROD", "NUORPROD = '" & IDOR$ & "'", NORPRO$, RAFE&)
         .MoveNext
         Loop
      End With
      ORFA.Close
      Set ORFA = Nothing
      '
      CONDI$ = "NUORPROD <> '' order by NuOrProd "
      Call CARGALISEL("!ECORDEF", "ORDEFABR", "NUORPROD/A12;REFEREN/A64", CONDI$)
      Call FRESHECHO("!ECORDEF")
   End If
   '
End Sub

Sub VALEMATER(MODOVA%)
   '
   Dim ORFASEL$(8192)
   Dim CANECE(32767)
   '
   If MODOVA% < 1 Or MODOVA% > 2 Then
      MODOVA% = ALTERNA%("Vale de Preparacion\Vale de Transferencia")
   End If
   If MODOVA% < 1 Or MODOVA% > 2 Then Exit Sub
   '
   If MODOVA% = 2 Then
     If DERACCE%("VALETRAN", "Vale de Transferencia de Stocks") < 2 Then
       Exit Sub
     End If
   End If
   '
10 CAOSEL% = 0
   XX$ = REGSEL$("SECPROD/Sector Productivo Destino")
   If Len(XX$) <= 4 Then
      Exit Sub
   End If
   '
   SECCI = CVS(Left$(XX$, 4))
   SecProd$ = Mid$(XX$, 5, 32)
   DEPODEST% = Asc(Mid$(XX$, 81, 1))
   If DEPODEST% < 1 Or TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPODEST%))) = "" Then
     Call MENSERR(24, "Falta Definir Depósito Asociado\al Sector '" + TRIM$(SecProd$) + "'.")
     GoTo 10
   End If
   '
   DepOrig% = Val(CONTROL$("", "DEPOCEN"))
   If DepOrig% < 1 Then
      'Call MENSERR(24, "Falta Definir Depósito Central")
      GoTo 13
   End If
   '
   EPAC$ = EMPREAC$
   If InStr(EPAC$, "DOSIVAC") > 0 Then
13    Call SELECHO("TADEPOSI/Depósito Origen de Materiales")
      DepOrig% = Val(VALACT1$("TADEPOSI"))
      If DepOrig% < 1 Then Exit Sub
   End If
   If DepOrig% = DEPODEST% Then
      Call MENSERR(24, "Operación Incorrecta.\Deposito Origen igual a Destino.")
      GoTo 10
   End If
   DEPOSI$ = TRIM$(ECOARCH$("TADEPOSI", Chr$(DepOrig%)))
   Call BLANARCH("!VALEMAT")
   '
   Screen.MousePointer = 11
   For U& = 1 To NUMAS%
      CANECE(U&) = 0
   Next U&
   '
   'SELECCIONAR ORDENES LANZADAS DE ESE SECTOR
   Call GENORDEFDAT(1)
11 FIN& = ULTREG&("ORDEFABR")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("ORDEFABR", U&)
      If Asc(Mid$(XX$, 108, 1)) = 1 Then
         If CVS(Mid$(XX$, 113, 4)) = SECCI Then
            CAOSEL% = CAOSEL% + 1
            ORFASEL$(CAOSEL%) = Mid$(XX$, 69, 12)
         End If
      End If
   Next U&
   '
   'CARGAR MATERIALES RESERVADOS PARA O.F'S
   Call RESERV07.GENRESDAT
   FIN& = ULTREG&("RESERVA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("RESERVA", U&)
      ORFA$ = Mid$(XX$, 3, 12)
      CI% = CVI(Left$(XX$, 2))
      CANTI = CVD(Mid$(XX$, 19, 8)) - CVD(Mid$(XX$, 27, 8))
      If CANTI < 0 Then CANTI = 0
      '
      If TRIM$(UCase$(UNIMED$(CI%))) = "GK" Then
         CANTI = CANTI / 1000
      End If
      '
      For KK% = 1 To CAOSEL%
         If ORFA$ = ORFASEL$(KK%) Then GoTo 15
      Next KK%
      GoTo 19
      '
15    If CI% > 0 Then
         If CI% <= NUMAS% Then
            CANECE(CI%) = CANECE(CI%) + CANTI
         End If
      End If
      '
19 Next U&
   '
   FIN& = NUMAS%
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      CI% = U&
      II1$ = CODEXT$(CI%)
      If CANECE(CI%) > 0 Then
        If Abs(SEPRODUC(CI%) - SECCI) > 0.0005 Then
          STODEST = STODEPOS(CI%, DEPODEST%)
          STOCENT = STODEPOS(CI%, DepOrig%)
          NENETA = CANECE(CI%) - STODEST
          If NENETA < 0 Then NENETA = 0
          If NENETA > 0.05 Then
             VALEMATX = NENETA
             If VALEMATX > STOCENT Then VALEMATX = STOCENT
             If VALEMATX < 0 Then VALEMATX = 0
             '
             ZZ$ = REGBLAN$("VALEMAT")
             Call REPLA(ZZ$, MKI$(CI%), 1, 2)
             Call REPLA(ZZ$, MKS$(CANECE(CI%)), 3, 4)
             Call REPLA(ZZ$, MKS$(STODEST), 7, 4)
             Call REPLA(ZZ$, MKS$(NENETA), 11, 4)
             Call REPLA(ZZ$, MKS$(STOCENT), 15, 4)
             Call REPLA(ZZ$, MKS$(VALEMATX), 19, 4)
             Call REPLA(ZZ$, MKS$(SECCI), 23, 4)
             Call REPLA(ZZ$, "......", 27, 6)
             Call REPLA(ZZ$, UBISTO$(CI%), 33, 16)
             Call REGAPP("!VALEMAT", ZZ$)
          End If
        End If
      End If
   Next U&
   Call CIERRARCH("!VALEMAT")
   '
   Screen.MousePointer = 1
   '
   If MODOVA% = 1 Then
      TITUPA$ = "Vale de PREPARACION" + " - de " + DEPOSI$ + " a " + SecProd$
      VTB% = VENTABU%("VALEMAP" + "/TITUPAN=" + TITUPA$)
      If VTB% < 0 Then GoTo 10
      '
      Screen.MousePointer = 11
      JJ& = 0
      For U& = 1 To ULTREG&("!VALEMAT")
         ZZ$ = REGLEIDO$("!VALEMAT", U&)
         JJ& = U&
         Call GRAREG("VALEMAT", ZZ$, JJ&)
      Next U&
      Call CIERRARCH("!VALEMAT")
      Call SETULTREG("VALEMAT", JJ&)
      Call CIERRARCH("VALEMAT")
      Call HEADERS("VALEMAT", "")
      SECCX$ = TRIM$(CSTRING$(MKS$(SECCI), 3, 8, 3, 2))
      Call HEADERS("VALEMAT", "Seccion: " + SECCX$ + " - " + SecProd$)
      Screen.MousePointer = 1
      Call CONECRUN("*PRIVALMA", "Vale de Preparacion de Materiales")
      GoTo 10
   End If
   '
25 TITUPA$ = "Vale de TRANSFERENCIA" + " - de " + DEPOSI$ + " a " + SecProd$
   VTB% = VENTABU%("VALEMAT" + "/TITUPAN=" + TITUPA$)
   If VTB% < 0 Then GoTo 10
   '
   ' BLOQUE AGREGADO PARA CONTROL DE NEGATIVOS
   CAITMO% = 0
   For YX& = 1 To ULTREG&("!VALEMAT")
      YY$ = REGLEIDO$("!VALEMAT", YX&)
      CI% = CVI(Mid$(YY$, 1, 2))
      CAN# = CVS(Mid$(YY$, 19, 4))
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
   If PUEDEMOVI%(DepOrig%) < 0 Then GoTo 25 ' CANCELA SALIDA
   ' FIN CONTROL DE NEGATIVOS
   '
   If COMALTER%("Confirme la Transferencia a Realizar\\Cancelar\Confirmar Transferencia") <> 2 Then GoTo 10
   '
   REGXX& = ULTREG&("TRANSFE")
   NRO& = 1
   If REGXX& > 0 Then
       NRO& = 1 + CVS(Mid$(REGLEIDO$("TRANSFE", REGXX&), 45, 4))
   End If
   If NRO& <= REGXX& Then
       Screen.MousePointer = 11
       For U& = 1 To REGXX&
           X0$ = REGLEIDO$("TRANSFE", U&)
           If CVS(Mid$(X0$, 45, 4)) >= NRO& Then
               NRO& = 1 + CVS(Mid$(X0$, 45, 4))
           End If
       Next U&
       Screen.MousePointer = 1
   End If
   '
   ' GRABAR TRANSFERENCIA
   '
   CMOV$ = TRIM$(CONTROL$("", "CODVALMA"))
   If CMOV$ = "" Then CMOV$ = "TR"
   '
   Call BLOQARCH("MOVISTO")
   Call BLOQARCH("TRANSFE")
   '
   Screen.MousePointer = 11
   NC% = 0
   VUNI# = 0
   MONE$ = ""
   LOTE$ = ""
   FECHA% = HOY(HOS$)
   DOPRI$ = "TR." + TRIM$(Str$(NRO&))
   DOSEC$ = "TR." + Mid$(HO$, 8, 1) + Mid$(HO$, 4, 2) + Left$(HO$, 2)
                             While Len(DOSEC$) < 10
                               DOSEC$ = DOSEC$ + " "
                             Wend
                              '
   NORI% = 0
   For U& = 1 To ULTREG&("!VALEMAT")
      XX$ = REGLEIDO$("!VALEMAT", U&)
      If CVS(Mid$(XX$, 19, 4)) >= 0.05 Then
         '
         CI1% = CVI(Mid$(XX$, 1, 2))
         CAN = CVS(Mid$(XX$, 19, 4))
         REFE$ = "TR.a " + SecProd$
         '
         CANTI = (-1) * CAN
         NORFI% = NORI% + 1
         Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DepOrig%, CANTI)
         '
         REFE$ = "TR.de " + DEPOSI$
         CANTI = CAN
         Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DEPODEST%, CANTI)
         '
         XXX$ = REGBLAN$("TRANSFE")
         Call REPLA(XXX$, MKI$(FECHA%), 1, 2)
         Call REPLA(XXX$, Chr$(DepOrig%), 3, 1)
         Call REPLA(XXX$, Chr$(DEPODEST%), 4, 1)
         Call REPLA(XXX$, MKI$(CI1%), 5, 2)
         Call REPLA(XXX$, MKI$(CI1%), 7, 2)
         Call REPLA(XXX$, LOTE$, 9, 16)
         Call REPLA(XXX$, LOTE$, 25, 16)
         Call REPLA(XXX$, MKS$(CAN), 41, 4)
         Call REPLA(XXX$, MKS$(NRO&), 45, 4)
         Call REGAPP("TRANSFE", XXX$)
         '
      End If
   Next U&
   '
   Call CIERRARCH("TRANSFE")
   Call BAJALDISCO
   '
   FORIPRE$ = CONTROL$("", "FORVALMA")
   If FORIPRE$ <> "" Then
      '
      FECHDOC$ = HOS$
      NUMEDOC$ = TRIM$(Str$(NRO&))
      While Len(NUMEDOC$) < 6
         NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Vale de Materiales"
      '
      CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = TRIM$(Str$(DepOrig%)) + " - " + DEPOSI$
      CODPARAM$(2) = "DESTINO": DOCPARAM$(2) = CSTRING$(MKS$(SECCI), 3, 8, 3, 2) + " - " + SecProd$
      CAPARDOC% = 2
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "UNIMED"
      CODTABU1$(4) = "CANTIDAD"
      CODTABU1$(5) = "CAKILOS"
      CACOLTAB1% = 5
      '
      DESTIDOC% = 0
      CAITAB1% = 0
      CAITCRU% = 0
      '
      For U& = 1 To ULTREG&("!VALEMAT")
         XX$ = REGLEIDO$("!VALEMAT", U&)
         If CVS(Mid$(XX$, 19, 4)) >= 0.05 Then
            '
            CICRU% = CVI(Mid$(XX$, 1, 2))
            CACOL1$ = Str$(CVS(Mid$(XX$, 19, 4)))
            KICOL1$ = Str$(CVS(Mid$(XX$, 19, 4)) * PESUNI(CICRU%))
            '
            CAITAB1% = CAITAB1% + 1
            TETABU1$(1, CAITAB1%) = CODEXT$(CICRU%)
            TETABU1$(2, CAITAB1%) = DESCRIT0$(CICRU%)
            TETABU1$(3, CAITAB1%) = UNIMED$(CICRU%)
            TETABU1$(4, CAITAB1%) = CACOL1$
            TETABU1$(5, CAITAB1%) = KICOL1$
            '
         End If
      Next U&
      '
      If CAITAB1% < 1 Then
         Screen.MousePointer = 1
         GoTo 10
      End If
      '
      Call PRINTDOC("FORVALMA")   ' Vale de Materiales
      Call BAJALDISCO
      '
   End If
   '
   Screen.MousePointer = 1
   GoTo 10
   '
End Sub

Sub VERORFA()
    '
    Call ABREORFAB
    '
3   Call SELECHO("ECORDEF/Indice General de Ordenes de Fabricación (Pendientes-Cumplidas-Anuladas)")
    NORFA$ = TRIM$(VALACT1$("ECORDEF"))
    If NORFA$ = "" Then GoTo 99
    '
    TIDOCUM$ = "Orden de Fabricación " + NORFA$
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 5
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Documento no Encontrado")
    GoTo 99
    '
5   For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    REDOCU$ = REGACT$("PDOCLST")
    URE& = ULTREG&("PDOCSPL")
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    GoTo 99
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      DCC& = CVS(Left$(REGLEIDO$("PDOCSPL", L&), 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      TBUF$ = REGLEIDO$("PDOCSPL", KKL&)
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  SPOOLSTRING$ = TPSP$
    FIN& = Len(TPSP$)
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7070
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7070
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    UI& = UI& + 1
    XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
    If InStr(XX$, TIDOCUM$) > 4 Then
      If InStr(XX$, NUDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 7
      End If
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    ANCHPAGINA = EDITFORM.Picture1(PAGINACTIVA%).Width
    EDITFORM.HScroll1.Min = 0
    EDITFORM.HScroll1.Value = 0
    EDITFORM.HScroll1.Max = 1
    If ANCHPAGINA > EDITFORM.Frame1.Width Then
      EDITFORM.HScroll1.Max = (ANCHPAGINA - EDITFORM.Frame1.Width) / 100
    End If
    EDITFORM.Command3.Enabled = False
    'EDITFORM.Command5.Enabled = False
    EDITFORM.Show 1
    '
    GoTo 3
    '
99  End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
End Sub

Sub ORFNUE(REPIM%)                           ' NUEVA ORDEN DE FABRICACION
    '
    ' REPIM%=1 ==> nueva orden de fabricacion
    ' REPIM%=2 ==> repite impresion de O/FAB
    '
    Static YADEPU%
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    Dim CILOTE%(256), CALOTE#(256), CALOTE1#(256), PUNMAT$(256), PUNOPE$(256)
    Dim CISUMI%(512), CASUMI#(512)
    '
    MONUME$ = "AUTO"                   ' modo de numeracion
    '
    Call CIERRARCH("*.*")
    Call FRESHECHO("ORFAEMI")
    '
100 GoSub ARMAPRIMERA
    X1$ = OBJPLA$("ORFANUE-1", X0$)
    If X1$ = "" Then GoTo 999
    '
    GoSub VERICODI
    If CIOF% < 1 Then GoTo 100
    CIXI% = CIOF%
    SPX$ = String$(4, 0): SPY$ = ""
    '
    GoSub AJUCANORF
    GoSub ARMASEGUNDA
    TMT% = TIMATE%(CIOF%)   ' PARA PONER TIPO DE MATERIAL EN POSICION 257,1
    XX22$ = X1$ + Chr$(TMT%)
102 XX22$ = OBJPLA$("ORFANUE-2", XX22$)
    REFOF$ = TRIM$(Mid$(XX22$, 258, 30))
    X2$ = Left$(XX22$, 256)
    If X2$ = "" Then
       If APLIXYZ$ <> "" Then
          Call FINAL("")
          Exit Sub
       End If
       GoTo 100
    End If
    '
    GoSub VALIDA1                   ' validacion de entrada
    If VALIDA% < 1 Then
       Call MENSERR(24, "Datos Incompletos o no Validos")
       GoTo 102
    End If
    '
103 XACCE$ = ""
    Call CIERRARCH("*.*")
    '
    TERENUME$ = ""
    If REPIM% <> 2 Then
       '
       NROPROX& = 1
       Call ABREORFAB
       SQLX$ = "SELECT NumeOrFa FROM ORDEFABR "
       SQLX$ = SQLX$ + "ORDER BY NumeOrFa DESC;"
       '
       Dim ORFABTEMP As ADODB.Recordset
       Set ORFABTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
       On Error Resume Next
       With ORFABTEMP
         .MoveFirst
         If Err < 1 Then
           NROPROX& = 1 + !NUMEORFA
         End If
       End With
       On Error GoTo 0
       '
       If NROPROX& <> NRO& Then
          TERENUME$ = "FLEXOFT ha Re-Numerado la O/F\como Orden Número '" + TRIM$(Str$(NROPROX&)) + "'."
          NRO& = NROPROX&
       End If
    End If
    '
    GoSub GRAORFNUE
    Call CIERRARCH("ORFANUE")
    GoSub GRAORDEFA
    Call CIERRARCH("ORDEFABR")
    Call CIERRARCH("ORFANUE")
    Call CIERRARCH("INDIORF")
    Call CIERRARCH("ECORDEP")
    Call CIERRARCH("ECORDEF")
    '
    Call COMUNI("Orden de Fabricación '" + NORFA$ + "' Generada.")
    '
    If APLIXYZ$ <> "" Then
       Call FINAL("")
       Exit Sub
    End If
    '
    GoTo 100
    '
' ************************************* subrutinas asociadas
ARMAPRIMERA:
    '
    NRO& = 1
    Call ABREORFAB
    '
    SQLX$ = "SELECT NumeOrFa FROM ORDEFABR "
    SQLX$ = SQLX$ + "ORDER BY NumeOrFa DESC;"
    'Set ORFABTEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim ORFABTEMP1 As ADODB.Recordset
    Set ORFABTEMP1 = FLEXCONN.Execute(FILTSQL$(SQLX$))
    On Error Resume Next
    With ORFABTEMP1
      .MoveFirst
      If Err < 1 Then
        NRO& = 1 + !NUMEORFA
      End If
    End With
   On Error GoTo 0
   Set ORFABTEMP1 = Nothing

    '
    X0$ = REGBLAN$("ORFANUE")
    Call REPLA(X0$, MKS$(NRO&), 1, 4)
    Call REPLA(X0$, MKI$(HO%), 69, 2)
    Call REPLA(X0$, MKI$(HO%), 109, 2)
    Call REPLA(X0$, MKI$(HO%), 113, 2)
    '
    If APLIXYZ$ <> "" Then
       CODICAN6$ = Mid$(APLIXYZ$, 8)
       PPXX% = InStr(CODICAN6$, Chr$(124))
       If PPXX% < 1 Then PPXX% = 1 + Len(CODICAN$)
       CODI6$ = Left$(CODICAN6$, PPXX% - 1)
       CANT6 = XVALO(Mid$(CODICAN6$, PPXX% + 1))
       Call REPLA(X0$, MKI$(CODINT%(CODI6$)), 71, 2)
    End If
    '
    Return
    '
'**************************************************************************
VERICODI:                                                         ' esta OK
    CIOF% = 0: UMED$ = "U.": LOSTA = 1
    CI% = CVI(Mid$(X1$, 71, 2))
    If CI% < 1 Then
         Call MENSERR(24, "Falta Código de Producto.")
       Else
         CIOF% = CI%
         UMED$ = UNIMED$(CI%)
         LOSTA = LOREPOMA(CI%)
         If LOSTA < 1 Then LOSTA = LOTESTAN(CI%)
    End If
    Return
    '
'**************************************************************************
BORENACCE:
    XACCE$ = ""
    RGACC& = 0
    For U& = 1 To ULTREG&("BOREACCE")
       YCC$ = REGLEIDO$("BOREACCE", U&)
       If CVI(Left$(YCC$, 2)) = CIOF% Then
          RGACC& = U&
          YACCE$ = YCC$
          GoTo 710
       End If
    Next U&
    '
    YACCE$ = String$(128, 0)
    Call REPLA(YACCE$, MKI$(CIOF%), 1, 2)
    Call REPLA(YACCE$, Space$(28), 3, 28)
    Call REPLA(YACCE$, Space$(40), 63, 40)
    Call REPLA(YACCE$, Space$(18), 105, 16)
    '
710 Call REPLA(YACCE$, TALLE$(CIOF%), 3, 12)
    Call REPLA(YACCE$, COLOR$(CIOF%), 15, 16)
    '
    RFF$ = RECFILT$("ESTRUNUE", 1, MKI$(CIOF%))
    If RFF$ <> "" Then
       REGEST& = CVS(Left$(RFF$, 4))
       CIHIL% = CVI(Mid$(REGLEIDO$("ESTRUNUE", REGEST&), 3, 2))
       Call REPLA(YACCE$, TIPOHIL$(CIHIL%), 63, 12)
       Call REPLA(YACCE$, TITULO$(CIHIL%), 75, 12)
       Call REPLA(YACCE$, COLOR$(CIHIL%), 87, 16)
    End If
    '
    XACCE$ = YACCE$
    If REPIM% <> 2 Then
       XACCE$ = OBJPLA$("BORENACCE", YACCE$)
    End If
    '
    If XACCE$ <> "" Then
       PROVHIL$ = TEL2CLI$((-1) * CVI(Mid$(XACCE$, 103, 2)))
       PARTHIL$ = Mid$(XACCE$, 105, 16)
       '
       LARCUTE$ = CSTRING$(Mid$(XACCE$, 31, 4), 3, 8, 1, 1)
       ACHCUTE$ = CSTRING$(Mid$(XACCE$, 35, 4), 3, 8, 1, 1)
       LARMATE$ = CSTRING$(Mid$(XACCE$, 39, 4), 3, 8, 1, 1)
       ACHMATE$ = CSTRING$(Mid$(XACCE$, 43, 4), 3, 8, 1, 1)
       LARCUCO$ = CSTRING$(Mid$(XACCE$, 47, 4), 3, 8, 1, 1)
       ACHCUCO$ = CSTRING$(Mid$(XACCE$, 51, 4), 3, 8, 1, 1)
       LARMACO$ = CSTRING$(Mid$(XACCE$, 55, 4), 3, 8, 1, 1)
       ACHMACO$ = CSTRING$(Mid$(XACCE$, 59, 4), 3, 8, 1, 1)
       '
       If RGACC& < 1 Then RGACC& = 1 + ULTREG&("BOREACCE")
       Call GRAREG("BOREACCE", XACCE$, RGACC&)
    End If
    Call CIERRARCH("BOREACCE")
    '
    Return

BORENCOMPLE:
    Call BLANARCH$("!BCOMPLEM")
    HAYCOMP% = 0
    For U& = 1 To ULTREG&("BCOMPLEM")
       YCC$ = REGLEIDO$("BCOMPLEM", U&)
       If CVI(Left$(YCC$, 2)) = CIOF% Then
          CACOM = CVS(Mid$(YCC$, 49, 4)) * CAN#
          Call REPLA(YCC$, MKS$(CACOM), 53, 4)
          CACOM = CVS(Mid$(YCC$, 57, 4)) * CAN#
          Call REPLA(YCC$, MKS$(CACOM), 61, 4)
          CACOM = CVS(Mid$(YCC$, 69, 4)) * CAN#
          Call REPLA(YCC$, MKS$(CACOM), 73, 4)
          Call REGAPP("!BCOMPLEM", YCC$)
          HAYCOMP% = 1
       End If
    Next U&
    '
    If HAYCOMP% = 0 Then
       YCC$ = REGBLAN$("BCOMPLEM")
       Call REPLA(YCC$, MKI$(CIOF%), 1, 2)
       Call REPLA(YCC$, "Cuerpo", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Manga", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Puño", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Cuello", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Tiras", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Capucha", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Bolsillo", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Presilla", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
    End If
    '
    REFCOMPXX$ = ""
    TITU$ = " - " + TRIM$(CSTRING$(MKD$(CAN#), 3, 8, 0, 2)) + " Prendas " + TRIM$(CODEXT$(CIOF%)) + " - " + DESCRIT0$(CIOF%)
    '
    VTB% = 0
    If REPIM% <> 2 Then
       VTB% = VENTABU%("BCOMPLEM/TITUPAN=Orden de Complementos:" + TITU$)
    End If
    '
    If VTB% >= 0 Then
       If HAYCOMP% > 0 Then
           J& = 0
           For U& = 1 To ULTREG&("BCOMPLEM")
             YCC$ = REGLEIDO$("BCOMPLEM", U&)
             If CVI(Left$(YCC$, 2)) <> CIOF% Then
                J& = J& + 1
                Call GRAREG("BCOMPLEM", YCC$, J&)
             End If
           Next U&
         Else
           J& = ULTREG&("BCOMPLEM")
       End If
       For U& = 1 To ULTREG&("!BCOMPLEM")
          YCC$ = REGLEIDO$("!BCOMPLEM", U&)
          CACOM = CVS(Mid$(YCC$, 53, 4)) / CAN#
          Call REPLA(YCC$, MKS$(CACOM), 49, 4)
          CACOX = CVS(Mid$(YCC$, 61, 4)) / CAN#
          Call REPLA(YCC$, MKS$(CACOX), 57, 4)
          CACOX = CVS(Mid$(YCC$, 73, 4)) / CAN#
          Call REPLA(YCC$, MKS$(CACOX), 69, 4)
          Call GRAREG("!BCOMPLEM", YCC$, U&)
          J& = J& + 1
          Call GRAREG("BCOMPLEM", YCC$, J&)
          '
          If CACOM > 0 Then
             If REFCOMPXX$ <> "" Then REFCOMPXX$ = REFCOMPXX$ + " - "
             REFCOMPXX$ = REFCOMPXX$ + TRIM$(Mid$(YCC$, 3, 14))
          End If
       Next U&
       Call SETULTREG("BCOMPLEM", J&)
    End If
    Return
    '
'**************************************************************************
AJUCANORF:                                                        ' esta OK
    CAMRP = 0
    CANORF = LOSTA
    LOECON = LOTESTAN(CIOF%)
    If MODORFA% = 2 Then
        CAMRP = CVD(Mid$(REGACT$("ORFAEMI"), 49, 8))
        CANORF = CAMRP
        If CAMRP < LOECON Then
          CANORF = LOECON
          Call MENSERR(24, "La Cantidad Pedida es Inferior\a un Lote Económico.\  \Ajustada a un Lote Standard Económico.")
        End If
      ElseIf MODORFA% = 1 Then
        CANORF = LOECON
        If CANORF < 1 Then CANORF = 1
    End If
    '
    If LINDIVIS(CIOF%) >= 1 Then
       CANORF1 = (1 + Int((CANORF - 1) / LINDIVIS(CIOF%))) * LINDIVIS(CIOF%)
       If Abs(CANORF - CANORF1) > 0.5 Then
          CANORF = CANORF1
          Call MENSERR(24, "Cantidad Ajustada a Múltiplo\de Lote Indivisible.")
       End If
    End If
    '
    If LOSTA < CANORF Then
      XLMX$ = MKI$(CI%) + MKS$(CANORF)
      XLMY$ = OBJPLA$("CARLOREPOMA", XLMX$)
      If XLMY$ <> "" Then
        LOSTA = CVS(Mid$(XLMY$, 3, 4))
        If LOSTA < 1 Then LOSTA = 1
        'Master.FindFirst ("CODINT = " & CIOF%)
        SQLX$ = "UPDATE MASTER SET LOREPOS =" & LOSTA & ""
        SQLX$ = SQLX$ + " WHERE CODINT = " & CIOF% & ""
         FLEXCONN.Execute (FILTSQL$(SQLX$))
        'If Master.NoMatch = False Then
          'Master.Edit
          'Master!LOREPOS = LOSTA
          'Master.Update
        'End If
      End If
    End If
    Return
    '
'**************************************************************************
ARMASEGUNDA:
    NRO& = CVS(Left$(X1$, 4))
    REOFE& = 0
    FENTRESOL% = HO%
    If MODORFA% = 2 Then
        FENTRESOL% = CVI(Mid$(REGACT$("ORFAEMI"), 57, 2))
        REOFE& = REGNUM&("ORFAEMI")
    End If
    '
    FENTREGA% = HO%
    If MODORFA% = 2 Then FENTREGA% = FENTRESOL%
    If FENTREGA% < HO% Then
        FEX = FENTREGA%
        Call MENSERR(24, "Imposible Cumplir la Fecha de\Entrega Solicitada del " + FECHATEX$(FEX) + ".\  \Entrega Re-Programada para el " + HOS$ + ".")
        FENTREGA% = HO%
    End If
    '
    DESTI$ = "STOCK": If TIMATE%(CIOF%) = 4 Then DESTI$ = "SEMIELA"
    'If SPY$ <> "" Then Destino$ = SPY$
    '
    If APLIXYZ$ <> "" Then
       CANORF = CANT6
       DESTI$ = "VENTA"
    End If
    '
    Call REPLA(X1$, UMED$, 73, 4)
    Call REPLA(X1$, MKD$(CDbl(CANORF)), 89, 8)
    Call REPLA(X1$, MKS$(LOSTA), 97, 4)
    Call REPLA(X1$, DESTI$, 101, 8)
    Call REPLA(X1$, MKI$(FENTREGA%), 109, 2)
    Call REPLA(X1$, SPX$, 115, 4)
    '
    Return
    '
VALIDA1:
    VALIDA% = (-1)
    X3$ = X2$: X1$ = X2$: X0$ = X2$
    If CVS(Mid$(X3$, 1, 4)) < 1 Then Return
    If CVI(Mid$(X3$, 69, 2)) < 1 Then Return
    If CVI(Mid$(X3$, 71, 2)) < 1 Then Return
    '
    VALIDA% = 0
    If TRIM$(Mid$(X3$, 73, 4)) = "" Then
        Call REPLA(X3$, "UNID", 73, 4)
    End If
    CANOR = CVD(Mid$(X3$, 89, 8))
    If CANOR < 1 Then Return
    If CVS(Mid$(X3$, 97, 4)) < 1 Then Call REPLA(X3$, MKS$(CANOR), 97, 4)
    If CVI(Mid$(X3$, 109, 2)) < HO% Then
        Call COMUNI("Fecha de Entrega Proyectada\es Anterior a la Fecha Actual")
        Return
    End If
    FENTRE% = CVI(Mid$(X3$, 109, 2))
    '
    CAN# = CVD(Mid$(X3$, 89, 8))
    CANORF = CAN#
    If FRALOMAX% < 0 Then
       If CAN# > CVS(Mid$(X3$, 97, 4)) Then
          CALOTX = CAN#
          Call REPLA(X3$, MKS$(CALOTX), 97, 4)
       End If
    End If
    '
    DESTI$ = Mid$(X3$, 101, 8)
    DESTJ$ = ECOARCH$("DESTFAB", DESTI$)
    If TRIM$(DESTJ$) <> "" Then
        DESTI$ = DESTI$ + " - " + DESTJ$
    End If
    '
    If Asc(Mid$(X3$, 111, 1)) > 9 Then
      PRIO% = Asc(Mid$(X3$, 111, 1)) Mod 10
      Call REPLA(X3$, Chr$(PRIO%), 111, 1)
    End If
    PRIO% = Asc(Mid$(X3$, 111, 1))
    FLAN% = CVI(Mid$(X3$, 113, 2))
    Status% = 0
    If FLAN% <= HO% Then
       If HAYDISTO%(CIOF%, NRO&, CANORF) = 1 Then
          Status% = 1
       End If
    End If
    Call REPLA(X3$, Chr$(Status%), 112, 1)
    '
    CI% = CVI(Mid$(X3$, 71, 2))
    DESCRI$ = TRIM$(CODEXT$(CI%)) + " - " + TRIM$(DESCRIT$(CI%))
    If REFOF$ <> "" Then DESCRI$ = TRIM$(REFOF$) + " - " + DESCRI$
    Call REPLA(X3$, DESCRI$, 5, 64)
    '
    TORFANUE$ = X3$     ' prepara registro para ORFANUE.DAT
    '
    FGEN = HO%
    STS$ = TRIM$(Str$(Status%)): While Len(STS$) < 3: STS$ = " " + STS$: Wend
    PRS$ = TRIM$(Str$(PRIO%)): While Len(PRS$) < 3: PRS$ = " " + PRS$: Wend
    SecProd = CVS(SPX$)
    FCIE = 0
    '
    X$ = REGBLAN$("INDIORF")
    Call REPLA(X$, MKS$(NRO&), 1, 4)
    Call REPLA(X$, FECHATEX$(FGEN), 5, 8)
    If Status% > 0 Then Call REPLA(X$, STS$, 13, 3)
    If PRIO% > 0 Then Call REPLA(X$, PRS$, 16, 3)
    Call REPLA(X$, ECOARCH$("SECPROD", MKS$(SecProd)), 21, 20)
    If REFOF$ <> "" Then
       Call REPLA(X$, REFOF$, 21, 20)
    End If
    If FCIE <> 0 Then Call REPLA(X$, FECHATEX$(FCIE), 43, 8)
    '
    TINDIORF$ = X$      ' prepara registro para INDIORF.DAT
    VALIDA% = 1
    Return
    '
GRAORFNUE:
    Call REPLA(TINDIORF$, MKS$(NRO&), 1, 4)
    '
    Z$ = RECFILT$("INDIORF", 1, MKS$(NRO&))
    If Len(Z$) >= 4 Then
          REGXX& = CVS(Left$(Z$, 4))
          Call GRAREG("INDIORF", TINDIORF$, REGXX&)
        Else
          Call REGAPP("INDIORF", TINDIORF$)
    End If
Return
'
GRAORDEFA:
    Call ABREORFAB
    '
    If REPIM% = 2 Then
       FLEXCONN.Execute (FILTSQL$("DELETE  FROM ORDEFABR WHERE NumeOrfa = " & NRO & ""))
    End If
    '
    NBRS$ = TRIM$(Str$(NRO&)): While Len(NBRS$) < 6
                                  NBRS$ = "0" + NBRS$
                               Wend
    NBRS$ = "OF-" + NBRS$ + "-1"
      NORFA$ = NBRS$
      II% = 1
      '
      SQLX$ = " SELECT * FROM OrdeFabr "
      Dim ORDEFFFABR As ADODB.Recordset
      Set ORDEFFFABR = New ADODB.Recordset
      ORDEFFFABR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
      'With OrdeFabr
      With ORDEFFFABR
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
          !NUMEORFA = NRO&
          !REFEREN = AJUSTI$(DESCRI$, 64)
          !IDSUBOR = NORFA$
          !FechGen = Now
          !Cod_Inte = CIOF%
          !COD_EXTE = CODEXT$(CIOF%)
          !Descrip = DESCRIT0$(CIOF%)
          !Unidad = UNIMED$(CIOF%)
          !CanProy = CAN#
          If Len(DESTI$) > 8 Then
          DESTI$ = TRIM$(Left(DESTI$, 8))
          End If
          !DESTINO = DESTI$
          !FechEnSol = CVDAT(FENTRE%)
          !Priorid = PRIO%
          !STATUORF = Status%
          !FechLan = CVDAT(FLAN%)
          !SecciPro = SecProd
          !FechCierr = CVDAT(0)
          !CanApro = 0
          !CanRech = 0
          !IdenRetra = ""
          !FechProyEntre = CVDAT(0)
          !FechTempEntre = CVDAT(0)
          !FechRecal = CVDAT(0)
          !Observa = ""
        .Update
      End With
      Set ORDEFFFABR = Nothing
      '
      CANS$ = " - (" + TRIM$(Str$(CAN#)) + ")"
      DESCRI1$ = Left$(DESCRI$, 52 - Len(CANS$)) + CANS$
      ECORDEF$ = AJUSTI$(NORFA$, 12) + DESCRI1$
      '
      Call FILTERPUTRECORD("ECORDEF", 1, NORFA$, ECORDEF$)
      If Status% = 1 Then
        Call FILTERPUTRECORD("ECORDEP", 1, NORFA$, ECORDEF$)
      End If
      '
Return
    '
999 Call CIERRARCH("*.*")
    Call FRESHALL
    Screen.MousePointer = 1
    If APLIXYZ$ <> "" Then
       Call FINAL("")
    End If
End Sub
'
Private Sub GenReservOC_Click()
   Call RESERV07.Command3_Click
End Sub



Private Sub LanOrd_Click()
   Call Command42_Click
End Sub

Private Sub LiToImpre_Click()
'
RESERV07.Option6 = True
RESERV07.Option4 = True
Call RESERV07.Command2_Click
'

End Sub

Private Sub LiToPan_Click()
'
RESERV07.Option6 = True
RESERV07.Option5 = True
Call RESERV07.Command2_Click
'
End Sub



Private Sub MaestArt_Click()
   '
   Call CONECRUN("ARCHIG07/AMASTO", "Padron Maestro de Articulos")
   '
End Sub

Private Sub NvaOrd_Click()
   Call Command41_Click
End Sub

Private Sub ReXMat_Click()
   '
   RESERV07.Option3 = True
   RESERV07.Option5 = True
   Call RESERV07.Command2_Click
   '
End Sub

Private Sub ReXNumOF_Click()
'
RESERV07.Option2 = True
RESERV07.Option5 = True
Call RESERV07.Command2_Click
'

End Sub

Private Sub ValPrep_Click()
Call VALEMATER(1)
End Sub

Private Sub ValTran_Click()
Call VALEMATER(2)
End Sub

Sub REPRINORFA()
   '
   Call ORFNUE(2)
   '
End Sub
'
Sub ANUCIEOF()
   '
   ' 1.- Selección del Producto
10 Call SELECHO("MASTER")
   If VALACT1$("MASTER") = "" Then Exit Sub
   CODD$ = VALACT1$("MASTER")
   CIXI% = CODINT%(CODD$)
   If CIXI% < 1 Then Exit Sub
   '
   ' 2.- Selección de la O/F
   Call APERBASDAT("STOCKS")
   SQLX$ = " SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
   SQLX$ = SQLX$ + "AND CodiMovi = 'OF' "
   SQLX$ = SQLX$ + "AND CantIngre > 0 "
   SQLX$ = SQLX$ + "ORDER BY FeReMovi ASC "
   'Set Movisel = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
   Dim MOVISEL As ADODB.Recordset
   Set MOVISEL = New ADODB.Recordset
   MOVISEL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   '
   With MOVISEL
      On Error Resume Next
      .MoveFirst '
      If Err Then
         On Error GoTo 0
         Call MENSERR(24, "No Hay Cierres de O/F Reversibles\para el Código Elegido.")
         GoTo 10
      End If
      On Error GoTo 0
      '
      JJ& = 0
      Do While .EOF = False
         Y$ = REGBLAN("ECORDEP")
         NOFA$ = !DoPriCor
         FERE$ = FECHATEX$(CVINT(!fechMov))
         CANX$ = FORMATNUM$(!CantIngre, "F10.1")
         If Not IsNull(!FeReMovi) Then DIVE$ = FORMATNUM$(CDbl(!FeReMovi) - 0.00001, "F18.6")
         TTXX$ = FERE$ + "    " + CANX$ + "  " + DIVE$
         Call REPLA(Y$, NOFA$, 1, 12)
         Call REPLA(Y$, TTXX$, 13, 52)
         '
         JJ& = JJ& + 1
         Call GRAREG("!ECORDEP", Y$, JJ&)
         '
      .MoveNext
      Loop
   End With
   Set MOVISEL = Nothing
   '
   Call SETULTREG("!ECORDEP", JJ&)
   Call CIERRARCH("!ECORDEP")
   Call FRESHECHO("!ECORDEP")
   Screen.MousePointer = 1
   '
   RSEL$ = REGSEL$("!ECORDEP")
   If Len(RSEL$) <= 4 Then GoTo 10
   '
   NOFASE$ = TRIM$(Left$(RSEL$, 12))
   CANUL = XVALO(Mid$(RSEL$, 25, 10))
   If COMALTER%("¿ Confirma la Anulación de Este Cierre ?\\No, Conservarlo\Si, Anularlo") <> 2 Then GoTo 10
   '
  
   Screen.MousePointer = 11
   FREMOX$ = TRIM$(Mid$(RSEL$, 37, 18))
   SQLX$ = " SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + "WHERE DoPriCor = '" & NOFASE$ & "' "
   SQLX$ = SQLX$ + "AND CDBL(FeReMovi) >= " & FREMOX$ & " "
   SQLX$ = SQLX$ + "ORDER BY FeReMovi ASC, NuorItem ASC"
   'Set MOVISELE = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    Dim MOVISELE As ADODB.Recordset
    Set MOVISELE = New ADODB.Recordset
    MOVISELE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

   '
   Screen.MousePointer = 11
   With MOVISELE
      On Error Resume Next
      .MoveFirst '
      If Err Then
         On Error GoTo 0
         Call MENSERR(24, "No Hay Movimientos de Cierre Reversibles\para la O/F Elegida.")
         GoTo 10
      End If
      On Error GoTo 0
      '
      JJ& = 0: NUORAN% = 0
      Do While .EOF = False
         If !NUORITEM < NUORAN% Then GoTo 39
         '.Edit
         REFA$ = !ReferCor
         CIXI% = !Cod_Inte
         NUORAN% = !NUORITEM
         !ReferCor = "Anul." + REFA$
         !CantIngre = 0
         !cantsalid = 0
         .Update
         If CIXI% > 0 Then
            CIMOVI%(CIXI%) = 1
            IDLOTX$ = TRIM$(!idenLote)
            If IDLOTX$ <> "" Then
               Call ACSALOTE(CIXI%, IDLOTX$)
            End If
         End If
      .MoveNext
      Loop
   End With
   '
39 Set MOVISELE = Nothing
   '
'3.- Actualiza saldo pendiente de la Orden de Fabricacion
   Call ABREORFAB
   SQLX$ = "SELECT * FROM OrdeFabr"
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NOFASE$ & "'"
   Dim ORDEFFFAB As ADODB.Recordset
   Set ORDEFFFAB = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With ORDEFFFAB
     '.FindFirst "IdSubOr = '" & NOFASE$ & "'"
     'If .NoMatch = False Then
       '.Edit
     On Error Resume Next
     If Not (.EOF And .BOF) Then
      On Error GoTo 0
      !CanApro = !CanApro - CANUL
      !STATUORF = 1                ' MARCA COMO EN CURSO
      .Update
     End If
     On Error GoTo 0
   End With
   ORDEFFFAB.Close
   Set ORDEFFFAB = Nothing
   '
'   Call VESARRGLO
   Screen.MousePointer = 1
   '
   Call COMUNI("Movimiento Anulado")
   GoTo 10
   
End Sub
'
Sub REAPEROF()
    '
    REBLA$ = REGBLAN$("ECORDEP")
    '
    Call ABREORFAB
    SQLX$ = " SELECT DISTINCT IdSubOr, StatuOrf, CanApro, CanProy FROM ORDEFABR "
    SQLX$ = SQLX$ + "WHERE StatuOrf = 3 "   ' para que no este anulada ni cumplida
    SQLX$ = SQLX$ + "AND CanApro < (CanProy - 0.005) "
    SQLX$ = SQLX$ + " ORDER BY IdSubOr "
    '
    'Set ORDEFABTMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim ORDEFABTMP As ADODB.Recordset
    Set ORDEFABTMP = New ADODB.Recordset
    ORDEFABTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

    With ORDEFABTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        On Error GoTo 0
        Call MENSERR(24, "No Existen Ordenes de Fabricación\que Puedan Re-Abrirse.")
        GoTo 99
      End If
      On Error GoTo 0
      JJ& = 0
      Do While (.EOF = False)
        Y$ = REBLA$
        Call REPLA(Y$, !IDSUBOR, 1, 12)
        Call REPLA(Y$, ECOARCH$("ECORDEF", !IDSUBOR), 13, 52)
        JJ& = JJ& + 1
        Call GRAREG("!ECORDEP", Y$, JJ&)
      .MoveNext
      Loop
    End With
    Set ORDEFABTMP = Nothing
    '
90  Call SETULTREG("!ECORDEP", JJ&)
    Call CIERRARCH("!ECORDEP")
    Call FRESHECHO("!ECORDEP")
    Screen.MousePointer = 1
    '
    Call SELECHO("!ECORDEP")
    NOFA$ = TRIM$(VALACT1$("!ECORDEP"))
    If NOFA$ <> "" Then
      If COMALTER%("Confirme la Re-Apertura de la " + NOFA$ + "\\Si, Re-Abrir\No, Cancelar") = 1 Then
        Call ABREORFAB
        PUDOABRIRSE% = 0
        '.FindFirst "Idsubor = '" & NOFA & "' AND StatuOrf = 3 "
        SQLX$ = "SELECT Statuorf FROM OrdeFabr "
        SQLX$ = SQLX$ + " WHERE Idsubor = '" & NOFA & "' AND StatuOrf = 3"
        Dim ESTADORDEN As ADODB.Recordset
        Set ESTADORDEN = New ADODB.Recordset
        ESTADORDEN.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

        With ESTADORDEN
          On Error Resume Next
          'If .NoMatch = False Then
            '.Edit
            .MoveFirst
            If Err < 1 Then
              Do While Not .EOF
                !STATUORF = 1
                .Update
                Y$ = REBLA$
                Call REPLA(Y$, NOFA$, 1, 12)
                Call REPLA(Y$, ECOARCH$("ECORDEF", NOFA$), 13, 52)
                Call FILTERPUTRECORD("ECORDEP", 1, NOFA$, Y$)
                PUDOABRIRSE% = 1
                '.FindNext "Idsubor = '" & NOFA$ & "' AND StatuOrf = 3 "
                .MoveNext
              Loop
            End If
            On Error GoTo 0
        End With
        Set ESTADORDEN = Nothing
        If PUDOABRIRSE% = 0 Then
            Call COMUNI("No Fue Posible Re-Abrir esta O/F")
          ElseIf PUDOABRIRSE% = 1 Then
            Call COMUNI(NOFA$ + " Re-Abierta.\No se Re-Generaron Reservas ni Operaciones")
        End If
      End If
    End If
    '
99  'OrdeFabr.Close
    '
End Sub
'
Function NUMIDEORPROD$(FF%)
    '
    FEXX$ = FECHATEX$(FF%)
    '
    IDENORPROD$ = Mid$(FEXX$, 7, 2) & Mid$(FEXX$, 4, 2)
    '
    VALOCAMP$ = "MAX(CAST(SUBSTRING(NuOrProd, 6,6)AS INTEGER))"
    CONDI$ = "NuOrProd LIKE '" & IDENORPROD$ & "%'"
    '
    PRONUMORPRO% = 1 + XVALO(VALOBADA("OrdeFabr", VALOCAMP$, CONDI$, "NOMESS"))
    '
    NUMORDPROD$ = TRIM$(Str$(PRONUMORPRO%))
    While Len(NUMORDPROD$) < 4: NUMORDPROD$ = "0" + NUMORDPROD$: Wend
    NUMIDEORPROD$ = IDENORPROD$ & "-" & NUMORDPROD$
    '
End Function


Function TOTREMITREC$(SF$)
    '
    TOTREMITREC$ = ""
    BRLOTE$ = ""
    '
    Call ABREORFAB
    CONDI$ = "Destino = '" & SF$ & "'"
    SQLX$ = "SELECT * FROM BALAREQMO "
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    '
    Dim BALRE As ADODB.Recordset
    Set BALRE = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With BALRE
      Do While Not .EOF
         BRLOTE$ = SAFETEXT(!idenLote)
         If BRLOTE$ <> "" Then
            CONDI$ = "IDENLOTE = '" & BRLOTE$ & "'"
            REMITREC$ = VALOBADA("LOTINGRE", "DOCUINGRE", CONDI$, "NOMESS")
            If InStr(TOTREMITREC$, REMITREC$) < 1 Then
                TOTREMITREC$ = TOTREMITREC$ & REMITREC$ & " / "
            End If
         End If
        .MoveNext
      Loop
    End With
    BALRE.Close
    Set BALRE = Nothing
    '
    REMITOSDESF$ = TOTREMITREC$
    '
End Function

Sub RECUOCPEN(NLIPIX$)
   '
   SQLX$ = "SELECT * FROM BALAREQMO "
   SQLX$ = SQLX$ + " WHERE Kilo_Asig > 0 "
   SQLX$ = SQLX$ + " AND IDENREQUE LIKE '" & NLIPIX$ & "%' "
   SQLX$ = SQLX$ + " AND (IDENLOTE LIKE 'OC%' OR IDENLOTE LIKE 'NP%')"
   Dim BALRE As ADODB.Recordset
   Set BALRE = New ADODB.Recordset
   BALRE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With BALRE
      Do While Not .EOF
        IDOC$ = SAFETEXT$(!idenLote)
        CIXI% = XVALO(!Cod_Inte)
        SALCOMPRA = SALDOCOM(IDOC$, CIXI%)
        If SALCOMPRA < XVALO(!Kilo_Asig) Then
           !Kilo_Asig = SALCOMPRA
           .Update
        End If
      .MoveNext
      Loop
   End With
   '
   BALRE.Close
   Set BALRE = Nothing
   '
End Sub
'
Function NroRemi(NROX$)
      Dim NRO$(10)
      NROY$ = NROX$
      LETRAFA$ = "X"
      TIDOCOM$ = "RT"
      While Len(NROY$) > 0
         If Asc(Left$(NROY$, 1)) < 48 Or Asc(Left$(NROY$, 1)) > 57 Then
              NROY$ = Mid$(NROY$, 2)
            Else
              GoTo 101
         End If
      Wend
      '
101   If NROY$ = "" Then NROY$ = "1-1"
      SUC% = 0
      C1% = 1: For c% = 1 To 10:  NRO$(c%) = "": Next c%
      '
      For c% = 1 To Len(NROY$)
         C2% = Asc(Mid$(NROY$, c%, 1))
         If C2% = 32 Then GoTo 102
         If (C2% < 48 Or C2% > 57) Then C1% = C1% + 1: GoTo 102
         NRO$(C1%) = NRO$(C1%) + Chr$(C2%)
102   Next c%
      '
      If C1% = 1 Then
         SUC% = 0
         VD# = XVALO(NRO$(1))
         NRO$(2) = NRO$(1)
         NRO$(1) = "0"
      End If
      '
      If XVALO(NRO$(2)) < 1 Or Len(NRO$(1)) > 4 Or Len(NRO$(2)) > 8 Then
          Call MENSERR(24, "Numero de Documento no Permitido")
          Exit Function
      End If
      '
      While Len(NRO$(1)) < 4: NRO$(1) = "0" + NRO$(1): Wend
      While Len(NRO$(2)) < 8: NRO$(2) = "0" + NRO$(2): Wend
      '
      
      PVCL% = PIVACLI%(-1 * NC%)
      NROY$ = TIDOCOM$ + "-" + LETRAFA$ + "-" + NRO$(1) + "-" + NRO$(2)
      NroRemi = NROY$
End Function

Private Sub Command48_Click()
    Command48.Enabled = False
10  OPX% = ALTERNA%("Asignacion de Lote a S/F\Des-Asignar Lotes Asignados a S/F\Des-Asignar Recepciones\Des-Asignacion de Materiales")
    If OPX% = 1 Then Call ASIGMANLOT: GoTo 10
    If OPX% = 2 Then Call DESASIGMANLOT: GoTo 10
    If OPX% = 3 Then Call DESASIGLOT: GoTo 10
    If OPX% = 4 Then Call DESASIGMAT: GoTo 10
    Command48.Enabled = True
End Sub

Sub DESASIGMANLOT()
    '
    'LOGICA DE LA RUTINA
    'PRESENTA VENTANA DE SELECCION DE SOLICITUD DE FABRICACION NO CERRADA
    'FILTRA EN BALAREQMO POR SF = ELEGIDA Y CANTIDAD ASIGNADA > 0
    'PRESENTA VENTA DE SELECCION DE CODIGOS Y LOTES PARA DESASIGNAR
    'BORRA EL REGISTRO DEL BALAREQMO
    'VALIDAR SI LA MISMA OF QUE EL COD_DESTI DEL LOTINGRE PRESENTE MENSAJE DE ADVERTENCIA
    'REALIZA UN ACSALOTE
    '
    '
    If DERACCE%("DESALOT", "Desasignacion de Lote") < 2 Then Exit Sub
    '
    CONDI$ = "IDSUBOR <> '' "
'    CONDI$ = CONDI$ + " AND (STATUORF < 3 OR STATUORF = 9) "
    CONDI$ = CONDI$ + " AND STATUORF < 3 "
'    CONDI$ = CONDI$ + " AND (NuOrProd = '' OR NuOrProd IS NULL) "
    CONDI$ = CONDI$ + " ORDER BY IDSUBOR ASC"
    Call CARGALISEL("!ECORDEF2", "ORDEFABR", "IDSUBOR/A12;REFEREN/A64", CONDI$)
    Call FRESHECHO("!ECORDEF2")
    '
10  Call SELECHO("!ECORDEF2/Des-Asignacion de Lotes a S/F")
    NORFA$ = VALACT1$("!ECORDEF2")
    If NORFA$ = "" Then Exit Sub
    '
    CONDI1$ = "IDSUBOR = '" & NORFA$ & "' AND (NuOrProd = '' OR NuOrProd IS NULL)"
    If CantRegistros("ORDEFABR", CONDI1$, "NOMESS") < 1 Then
       Call MENSERR(24, "Impósible  Des-Asignar Lotes a Solicitud de Fabricación: " & NORFA$ & "\ Tiene Orden de Producción ")
       GoTo 10
    End If
    '
'20  CONDI$ = "DESTINO = '" & NORFA$ & "' AND IDENLOTE <> '' AND (KILO_REQUE = 0 or KILO_REQUE is null)"
20  CONDI$ = "DESTINO = '" & NORFA$ & "' AND IDENLOTE <> '' AND Kilo_Asig > 0 "
    REGG& = CantRegistros&("BALAREQMO", CONDI$)
    If REGG& < 1 Then
        Call MENSERR(24, "No Existen Lotes Asignados Para la " & NORFA$)
        CONDII$ = "COD_DESTI = '" & NORFA$ & "' AND CANALTA > CANCONSU"
        Call ACTUREGISTRO("LOTINGRE", "COD_DESTI", CONDII$, "", RAFE&, "NOMESS")
        If RAFE& > 0 Then
           Call COMUNI(Str$(RAFE&) & " Registros Recuperados en 'LOTINGRE'")
        End If
        GoTo 10
    End If
    CONDI$ = CONDI$ & "ORDER BY COD_INTE"
    '
    CONDI2$ = "BALAREQMO.DESTINO = '" & NORFA$ & "' "
    CONDI2$ = CONDI2$ + " AND BALAREQMO.IDENLOTE <> '' "
    CONDI2$ = CONDI2$ + " AND (BALAREQMO.KILO_REQUE = 0 or BALAREQMO.KILO_REQUE is null)"
    CONDI2$ = CONDI2$ + " ORDER BY BALAREQMO.COD_INTE "
    '
    'Call CARGALISEL("!DESALOT", "BALAREQMO+LOTINGRE/BALAREQMO.IDENLOTE = LOTINGRE.IDENLOTE", "BALAREQMO.COD_INTE/C16;LOTINGRE.DOCUINGRE/A18;BALAREQMO.IDENLOTE/A16;LOTINGRE.LOTECOL_PROV/A16;BALAREQMO.KILO_ASIG/F15.3", CONDI2$)
    Call CARGALISEL("!DESALOT", "BALAREQMO+LOTINGRE/BALAREQMO.IDENLOTE = LOTINGRE.IDENLOTE", "BALAREQMO.COD_INTE/C16;BALAREQMO.IDENLOTE/A16;LOTINGRE.DOCUINGRE/A18;LOTINGRE.LOTECOL_PROV/A16;BALAREQMO.KILO_ASIG/F15.3", CONDI2$)
    'Call CARGALISEL("!DESALOT", "BALAREQMO+ORDEFABR/BALAREQMO.IDENLOTE = ORDEFABR.IDSUBOR", "BALAREQMO.COD_INTE/C16;BALAREQMO.IDENLOTE/A36;ORDEFABR.IDENLOTE/A16;BALAREQMO.KILO_ASIG/F15.3", CONDI2$, "NOCLEAR")
    Call CARGALISEL("!DESALOT", "BALAREQMO+ORDEFABR/BALAREQMO.IDENLOTE = ORDEFABR.IDSUBOR", "BALAREQMO.COD_INTE/C16;BALAREQMO.IDENLOTE/A36;ORDEFABR.IDENLOTE/A16;BALAREQMO.KILO_ASIG/F15.3", CONDI2$, "NOCLEAR")
    '
    Call FRESHECHO("!DESALOT")
    Call SELECHO("!DESALOT/Des-Asignacion de Lotes a " & NORFA$)
    CI% = CODINT%(VALACT1$("!DESALOT"))
    If CI% > 0 Then
        LOTEDESA$ = TRIM$(Mid$(VALACT2("!DESALOT"), 1, 16))
'        ' Se modifico los parametros de lectura
'        LOTEDESA$ = TRIM$(Mid$(VALACT2("!DESALOT"), 15, 16))
        '
        'VALIDACION SI EL LOTE ESTA ASIGNADO ORIGINALMENTE A LA SF
        CONDI3$ = "IDENLOTE = '" & LOTEDESA$ & "'"
        SOLFA$ = VALOBADA("LOTINGRE", "COD_DESTI", CONDI3$, "NOMESS")
        If UCase$(TRIM$(SOLFA$)) = UCase$(TRIM$(NORFA$)) Then
            TXTX$ = "El Lote que está Intentando Desasignar \"
            TXTX$ = TXTX$ + "Ha sido Asignado Originalmente a la " & NORFA$ & " \"
            TXTX$ = TXTX$ + "Está Seguro de Continuar"
            RESP% = COMALTER%(TXTX$ & "\\No, Cancelar\Si, Continuar")
            If RESP% < 2 Then GoTo 10
        End If
        '
        CONDI$ = "COD_INTE = " & CI% & " AND IDENLOTE = '" & LOTEDESA$ & "' And DESTINO = '" & NORFA$ & "'"
        'COMENTADO POR QUE SE PONEN EN 0 Y LUEGO SE TERMINAN BORRANDO . AHORA SE BORRAN DIRECTAMENTE
'        Call ACTUREGISTRO("BALAREQMO", "KILO_ASIG", CONDI$, 0, REGG&, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "CANTIDAD", CONDI$, 0, REGG&, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "MARQASI", CONDI$, 0, REGG&, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "POSICION", CONDI$, "", REGG&, "NOMESS")
'        Call ACTUREGISTRO("BALAREQMO", "IDENLOTE", CONDI$, "", REGG&, "NOMESS")
'        '
'        CONDI$ = "COD_INTE = " & CI% & " AND IDENLOTE = ''AND (KILO_REQUE = 0 or KILO_REQUE is null)"
        Call BORRAREGISTROS("BALAREQMO", CONDI$, REGG&, "NOMESS")
        '
        CONDI$ = "COD_INTE = " & CI% & " And DESTINO = '" & NORFA$ & "'"
        KISALD = XVALO(VALOBADA("BALAREQMO", Abs(KILO_REQUE - Kilo_Asig), CONDI$, "NOMESS"))
        Call ACTUREGISTRO("BALAREQMO", "kilo_saldo", CONDI$, KISALD, REGG&, "NOMESS")
        '
        CONDI$ = "COD_INTE = " & CI% & " AND IDENLOTE = '" & LOTEDESA$ & "'"
        Call ACTUREGISTRO("LOTINGRE", "CANRESER", CONDI$, 0, REGG&, "NOMESS")
        CANNS = XVALO(VALOBADA("LOTINGRE", "ABS(CANALTA-CANCONSU - CANRESER)", CONDI$, "NOMESS"))
        Call ACTUREGISTRO("LOTINGRE", "CANSALDO", CONDI$, CANNS, REGG&, "NOMESS")
        Call ACTUREGISTRO("LOTINGRE", "COD_DESTI", CONDI$, "", REGG&, "NOMESS")
        Call ACSALOTE(CI%, LOTEDESA$, "NOMESS")
        Call MENSERR(24, "Des-Asignación de Lote " & LOTEDESA$ & " Procesada")
        GoTo 20
    Else
        GoTo 10
    End If
 
End Sub

Sub ASIGMANLOT()
    '
    'LOGICA DE RUTINA
    '*****************
    'PRESENTA LISTA DE SELECCION DE SF NO CERRADAS
    'VERIFICACION SI TIENE  COMPONENTES  REQUERIDOS
    'VALIDACIONES QUE NO TENGA ORDEN DE PRODUCCION
    'PRESENTA LISTA DE MATERIALES REQUERIDOS PARA LA SF
    'VERIFICA SI EXISTEN LOTES DISPONIBLES EN LOTINGRE
    'PRESENTA LISTA DE LOTES DISPONIBLES
    'VALIDA QUE EL DESTINO SEA EL MISMO QUE LA SF ACTIVA
    'REALIZA ACSALOTE
    'AGREGA REGISTROS EN BALAREQMO
    'REALIZA UN ACSALOTE
    
    '
    CONDI$ = "IDSUBOR <> '' "
    CONDI$ = CONDI$ + " AND STATUORF < 3 "
    
'    CONDI$ = CONDI$ + " AND (NuOrProd = '' OR NuOrProd IS NULL) "
'    CONDI$ = CONDI$ + " AND Referen <> '' "
'    CONDI$ = CONDI$ + " AND Referen IS NOT NULL "
    CONDI$ = CONDI$ + " ORDER BY IDSUBOR ASC"
    Call CARGALISEL("!ECORDEF", "ORDEFABR", "IDSUBOR/A12;REFEREN/A64", CONDI$)
    Call FRESHECHO("!ECORDEF")
    '
10  Call SELECHO("!ECORDEF/Asignacion de Lotes a S/F")
    NORFA$ = VALACT1$("!ECORDEF")
    If NORFA$ = "" Then Exit Sub
    '
20  CONDI$ = "DESTINO = '" & NORFA$ & "' AND IDENLOTE = ''"
    REGG& = CantRegistros&("BALAREQMO", CONDI$)
    If REGG& < 1 Then
        Call MENSERR(24, "No Existen Asignaciones Pendientes Para " & NORFA$)
        GoTo 10
    End If
    '
    CONDI1$ = "IDSUBOR = '" & NORFA$ & "' AND (NuOrProd = '' OR NuOrProd IS NULL)"
    If CantRegistros("ORDEFABR", CONDI1$, "NOMESS") < 1 Then
       Call MENSERR(24, "Impósible  Asignar Lotes a O/Fabricación: " & NORFA$ & "\ Tiene Orden de Producción ")
       GoTo 10
    End If
    '
    CONDI2$ = "IDSUBOR = '" & NORFA$ & "' AND (Referen = '' OR Referen IS NULL)"
    If CantRegistros("ORDEFABR", CONDI2$, "NOMESS") > 0 Then
       Call MENSERR(24, "Impósible  Asignar Lotes a O/Fabricación: " & NORFA$ & "\ Falta Referencia")
       GoTo 10
    End If

30  CONDI$ = "DESTINO = '" & NORFA$ & "' GROUP BY COD_INTE"
    Call CARGALISEL("!ASIGLOT", "BALAREQMO", "COD_INTE/C16+MASTER/A40;SUM(KILO_REQUE)-SUM(KILO_ASIG)/F15.3", CONDI$)
    Call FRESHECHO("!ASIGLOT")
    Call SELECHO("!ASIGLOT/Asignacion de Lotes a Manual a " & NORFA$)
    CI% = CODINT%(VALACT1$("!ASIGLOT"))
    If CI% < 1 Then GoTo 10
    'CONDI1$ = "COD_INTE = '" & CI% & "' AND COD_DESTI = '' AND (CANALTA - CANCONSU - CANRESER) > 0 AND LIBERA <> 'RECH'"
    'ESTO PERMITE QUE UN LOTE PREASIGNADO PERO CON SALDO, PUEDA ASIGNARSE A OTRA S/F
    CONDI1$ = "COD_INTE = '" & CI% & "' AND (CANALTA - CANCONSU - CANRESER) > 0 AND LIBERA <> 'RECH'"
    REGG1& = CantRegistros&("LOTINGRE", CONDI1$)
    CONDI2$ = "COD_INTE = '" & CI% & "' AND STATUORF < 3 AND ABS(CANPROY - CANAPRO) > 0"
    REGG2& = CantRegistros&("ORDEFABR", CONDI2$)
    If REGG1& + REGG2& < 1 Then
        Call MENSERR(24, "No Existen Lotes Disponibles del Codigo " & CODEXT$(CI%) & "\para Asignar a la " & NORFA$)
        GoTo 20
    End If
    '
    'MODIFIQUE POSICION DE DOCUINGRE PARA QUE SE VISUALICE TOTALMENTE LA SF POR QUE SUPERABA EL ANCHO DE LA VENTANA
    Call CARGALISEL("!ASIGLOT2", "LOTINGRE", "COD_INTE/C16;IDENLOTE/A16;DOCUINGRE/A14;LoteCol_Prov/A16;(CANALTA - CANCONSU - CANRESER)/F15.3;Cod_Desti/A16", CONDI1$)
    Call CARGALISEL("!ASIGLOT2", "ORDEFABR", "COD_INTE/C16;IDSUBOR/A16;IDENLote/A34;(CANPROY - CANAPRO)/F31.3", CONDI2$, "NOCLEAR")
    Call FRESHECHO("!ASIGLOT2")
40  Call SELECHO("!ASIGLOT2/Asignacion de Lotes Manual a " & NORFA$)
    CI% = CODINT%(VALACT1$("!ASIGLOT2"))
    
    If CI% > 0 Then
        'VALIDACION QUE NO DIFIERA EL DESTINO
        DEST$ = TRIM$(UCase$(Mid$(VALACT2("!ASIGLOT2"), 66, 16)))
        If TRIM$(UCase$(NORFA$)) <> DEST$ And DEST$ <> "" Then
           Call MENSERR(24, "Impósible Efectuar Asignación.\El Lote Elegido Posee Otro Destino.\Para Utilizarlo Deberà Fraccionarlo")
           GoTo 40
        End If
         'SE COMENTA POR QUE PIERDE SENTIDO ACUTALIZAR EL DESTINO EN EL LOTINGRE SI EL MISMO NO SE CAMBIA
        LOTEASIGN$ = TRIM$(Mid$(VALACT2("!ASIGLOT2"), 1, 16))
'        CONDI$ = "COD_INTE = " & CI% & " AND IDENLOTE = '" & LOTEASIGN$ & "'"
'        Call ACTUREGISTRO("LOTINGRE", "COD_DESTI", CONDI$, NORFA$, REGG&, "NOMESS")

'        CANSALOT$ = XVALO(VALOBADA("LOTINGRE", "Abs(CanAlta - CANCONSU - CANRESER)", CONDI$, "NOMESS"))
'        Call ACTUREGISTRO("LOTINGRE", "CANSALDO", CONDI$, CANSALOT$, REGG&, "NOMESS")
        Call ACSALOTE(CI%, LOTEASIGN$, "NOMESS")
        GoTo 45    '    GoTo 10
      Else
        GoTo 10
    End If
    '
45  'AQUI AGREGAR A BALAREQMO
    Dim BALTEMP As ADODB.Recordset
    Set BALTEMP = New ADODB.Recordset
    BALTEMP.Open FILTSQL$("SELECT * FROM BALAREQMO WHERE IDENREQUE='X'"), FLEXCONN, adOpenKeyset, adLockOptimistic
    '
    ' ESTO PARA CONTEMPLAR LOTES EN STOCK MAS EN PROCESO
    SALNOASI = XVALO(VALOBADA("LOTINGRE", "CANALTA - CANCONSU", "COD_INTE = " & CI% & " AND IDENLOTE = '" & LOTEASIGN$ & "'"))
    SALNOASI = SALNOASI + XVALO(VALOBADA("ORDEFABR", "CANPROY - CANAPRO", "IDSUBOR = '" & LOTEASIGN$ & "'"))
    SALNOASI = SALNOASI - XVALO(VALOBADA("BALAREQMO", "SUM(KILO_ASIG)", "COD_INTE = " & CI% & " AND IDENLOTE = '" & LOTEASIGN$ & "'"))
    ' LA SIGUIENTE INTERVENCION ES PARA EL RECLAMO DE PLUCHINSKI DE NO PODER RE-ASIGNAR
    ' EL LOTE IR-399-1 DEL BERN0512 EN FECHA 10 DE SETIEMBRE. OCURRIA QUE ESE LOTE ESTABA
    ' ASIGNADO A LAS SF's 370 Y 372, QUE AUNQUE CERRADAS MANTENIAN EL MATERIAL RESERVADO
    ' CUANDO SE DECIDA, COMENTAR LA ULTIMA LINEA ANTERIOR Y DESCOMENTAR LAS DOS SIGUIENTES
    ' ... CON ESO SE CORRIGE - EPB 11/SET/2010
    ' CANRESE = XVALO(VALOBADA("BALAREQMO INNER JOIN ORDEFABR ON BALAREQMO.DESTINO = ORDEFABR.IDSUBOR ", "SUM(KILO_ASIG)", "BALAREQMO.COD_INTE = " & CI% & " AND BALAREQMO.IDENLOTE = '" & LOTEASIGN$ & "' AND ORDEFABR.STATUORF < 3"))
    ' SALNOASI = SALNOASI - CANRESE
    If SALNOASI < 0 Then
       Call MENSERR(24, "Lote '" & LOTEASIGN$ & "' Totalmente Consumido o Reservado.")
       GoTo 10
    End If
    '
'    ' AQUI NETEAR DE LO YA ASUGNADO
    CONDI$ = "Destino = '" & NORFA$ & "' AND COD_INTE = " & CI%
    SQLX$ = "SELECT IDENREQUE,Cod_Inte,Sum(Kilo_Reque) as Sureque,Sum(Kilo_Asig) as Suasig FROM BALAREQMO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    SQLX$ = SQLX$ + " GROUP BY IDENREQUE,COD_INTE"
    '
    Dim BALAREQ As ADODB.Recordset
    Set BALAREQ = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With BALAREQ
      Do While Not .EOF
         CIXI% = XVALO(!Cod_Inte)
         If XVALO(!SUREQUE) > XVALO(!SUASIG) Then
            CANTU = XVALO(!SUREQUE) - XVALO(!SUASIG)
            If CANTU > SALNOASI Then CANTU = SALNOASI
            SALNOASI = SALNOASI - CANTU
            CODE$ = CODEXT$(CIXI%)
            '
            'Se agrega Registro en Balareqmo Informando el Pedido Generado
            NROPED$ = LOTEASIGN$
               CONDI$ = "IDENREQUE = '" & SAFETEXT$(!IdenReque) & "' AND Kilo_Reque > 0"
            BALTEMP.AddNew   ' YA CONVERTIDO
                   On Error Resume Next
            BALTEMP!CodiEmpr = CodiEmp%
                   On Error GoTo 0
            BALTEMP!CODICOMPRO = SAFETEXT(VALOBADA("BALAREQMO", "CODICOMPRO", CONDI$, "NOMESS"))
            BALTEMP!NUMECOMPRO = XVALO(VALOBADA("BALAREQMO", "NUMECOMPRO", CONDI$, "NOMESS"))
            BALTEMP!CODSUCU = XVALO(VALOBADA("BALAREQMO", "CODSUCU", CONDI$, "NOMESS"))
            BALTEMP!NUORITEM = XVALO(VALOBADA("BALAREQMO", "NUORITEM", CONDI$, "NOMESS"))
            BALTEMP!IdenReque = SAFETEXT(VALOBADA("BALAREQMO", "IDENREQUE", CONDI$, "NOMESS"))
            BALTEMP!Cod_Inte = XVALO(VALOBADA("BALAREQMO", "Cod_Inte", CONDI$, "NOMESS"))
            BALTEMP!idenLote = NROPED$
            BALTEMP!Kilo_Asig = CANTU ' XVALO(!KILO_SALDO)
            BALTEMP!FeReMovi = CVDAT(HOY(HOS))
            BALTEMP!fechMov = CVDAT(HOY(HOS))
            BALTEMP!CREADOPOR = "FLEX"
            BALTEMP!DESTINO = SAFETEXT(VALOBADA("BALAREQMO", "Destino", CONDI$, "NOMESS"))
            BALTEMP.Update
         End If
      .MoveNext
      Loop
    End With
    '
    Call ACSALOTE(CI%, LOTEASIGN$, "NOMESS")
    '
    Call COMUNI("Asignación de Lote " & LOTEASIGN$ & " Procesada.")
    GoTo 30
    '
End Sub

Sub DESASIGLOT()
    '
    Call COMUNI("Esta Transacción No Está Disponible \ Para Reasignar un Nuevo Destino \Debe Realizar Fraccionamiento de Lote")
    Exit Sub
''    If DERACCE%("DESALOT", "Desasignacion de Lote") < 2 Then Exit Sub
''    '
''    CONDI$ = "IDSUBOR <> '' "
''    CONDI$ = CONDI$ + " AND (NuOrProd = '' OR NuOrProd IS NULL) "
''    CONDI$ = CONDI$ + " AND Referen <> '' "
''    CONDI$ = CONDI$ + " AND Referen IS NOT NULL "
''    CONDI$ = CONDI$ + " ORDER BY IDSUBOR ASC"
''    Call CARGALISEL("!ECORDEF2", "ORDEFABR", "IDSUBOR/A12;REFEREN/A64", CONDI$)
''    Call FRESHECHO("!ECORDEF2")
''    '
''10  Call SELECHO("!ECORDEF2/Des-Asignacion de Lotes a S/F Anuladas")
''    NORFA$ = VALACT1$("!ECORDEF2")
''    If NORFA$ = "" Then Exit Sub
''    '
''20  CONDI$ = "COD_DESTI = '" & NORFA$ & "'"
''    REGG& = CantRegistros&("LOTINGRE", CONDI$)
''    If REGG& < 1 Then
''        Call MENSERR(24, "No Existen Lotes Asignados Para la " & NORFA$)
''        GoTo 10
''    End If
''    CONDI$ = CONDI$ & "ORDER BY COD_INTE"
''    Call CARGALISEL("!DESALOT", "LOTINGRE", "COD_INTE/C16;IDENLOTE/A16;DOCUINGRE/A18;CANALTA/F15.3", CONDI$)
''    Call FRESHECHO("!DESALOT")
''    Call SELECHO("!DESALOT/Des-Asignacion de Lotes a " & NORFA$)
''    CI% = CODINT%(VALACT1$("!DESALOT"))
''    IDLO$ = TRIM$(Mid$(VALACT2$("!DESALOT"), 1, 16))
'''    ' Se modifico los parametros de lectura
'''    IDLO$ = TRIM$(Mid$(VALACT2$("!DESALOT"), 15, 16))
''    '
''    If CI% > 0 Then
''        CONDI1$ = "DESTINO='" & NORFA$ & "' AND IDENLOTE='" & IDLO$ & "'"
''        REGG& = CantRegistros&("BALAREQMO", CONDI1$)
''        If REGG& > 0 Then
''            Call MENSERR(24, "Material Ya Asignado en Balance de Requerimientos.\Utilizar la Opción Des-Asignar Lotes Asignados a S/F")
''            GoTo 10
''        End If
''        CONDI1$ = "COD_DESTI='" & NORFA$ & "' AND IDENLOTE='" & IDLO$ & "'"
''        Call ACTUREGISTRO("LOTINGRE", "COD_DESTI", CONDI1$, "", REGG&, "NOMESS")
''        GoTo 20
''    Else
''        GoTo 10
''    End If
End Sub

Sub ACTUNUORPRO()
    '
    Call ABREORFAB
    SQLX$ = "SELECT * FROM ORDEFABR "
    SQLX$ = SQLX$ & "WHERE NUORPROD <> '' AND LEN(NUORPROD)< 9"
    '
    Dim NUORPTMP As ADODB.Recordset
    Set NUORPTMP = FLEXCONN.Execute(SQLX$)
    '
    With NUORPTMP
       Do While Not NUORPTMP.EOF
          NUMEVIE$ = SAFETEXT(!NUORPROD)
          SOLFA$ = !IDSUBOR
          If NUMEVIE$ = "0000" Then
             GoSub RECUNUMOR
   '                MsgBox SOLFA$ + vbCrLf + NORPRO$
             If NORPRO$ <> "" Then
                CONDI$ = "IDSUBOR = '" & SOLFA$ & "'"
                Call ACTUREGISTRO("ORDEFABR", "NUORPROD", CONDI$, NORPRO$, REGS&, "NOMESS")
                NUMEVIE$ = NORPRO$
             End If
          End If
          '
          PREFIJ$ = TRIM$(Mid$(NUMEVIE$, 1, 5))
          If Mid$(PREFIJ$, 5, 1) = "/" Then
             NUMORDPROD$ = TRIM$(Mid$(NUMEVIE$, 6))
             While Len(NUMORDPROD$) < 4: NUMORDPROD$ = "0" + NUMORDPROD$: Wend
             NUMENUE$ = PREFIJ$ & NUMORDPROD$
             CONDI$ = "IDSUBOR = '" & SOLFA$ & "'"
             Call ACTUREGISTRO("ORDEFABR", "NUORPROD", CONDI$, NUMENUE$, REGS&, "NOMESS")
          End If
       NUORPTMP.MoveNext
       Loop
    End With
    NUORPTMP.Close
    Set NUORPTMP = Nothing
    '
Exit Sub
'
RECUNUMOR:
    CADENA$ = "fue Lanzada como Orden de Produccion "
    Screen.MousePointer = 11
    JJ& = 0
    NORPRO$ = ""
    For U& = 1 To ULTREG&("MENSAJES")
       XX$ = REGLEIDO$("MENSAJES", U&)
       If InStr(1, UCase(XX$), UCase(CADENA$)) > 1 Then
          XX1$ = REGLEIDO$("MENSAJES", U& - 1)
          If InStr(XX1$, SOLFA$) > 70 Then
             PPXX% = InStr(UCase$(XX$), "PRODUCCION '")
             If PPXX% > 0 Then
                NORPRO$ = TRIM$(Mid$(XX$, PPXX% + 12, 12))
                NORPRO$ = REPLACAR$(NORPRO$, ".", "")
                NORPRO$ = TRIM$(REPLACAR$(NORPRO$, "'", ""))
                Return
             End If
          End If
       End If
    Next U&
Return

End Sub
Sub DESASIGMAT()

    Call COMUNI("Esta Transacción No Está Disponible \ Para Reasignar un Nuevo Destino \Debe Realizar Fraccionamiento de Lote")
    Exit Sub

    
'''    If DERACCE%("DESMATE", "Desasignacion de Material") < 2 Then Exit Sub
'''    '
'''10  LOTE$ = LOTSELEC(CI%)
'''    '
'''    If TRIM(LOTE$) = "" Then Exit Sub
'''    '
'''    CONDI1$ = "IDENLOTE='" & LOTE$ & "'"
'''    REGG& = CantRegistros&("BALAREQMO", CONDI1$)
'''    If REGG& > 0 Then
'''         Call MENSERR(24, "Material Ya Asignado en Balance de Requerimientos.\Utilizar la Opción Des-Asignar Lotes Asignados a S/F")
'''         GoTo 10
'''    End If
'''
'''    RESP% = COMALTER%("¿Desea Realmente Dejar este Lote de Libre Disponibilidad?\\No, Cancelar\Si, Continuar")
'''    If RESP% = 2 Then
'''       CONDI$ = "Cod_inte = " & CI% & " AND IDENLOTE ='" & LOTE$ & "'"
'''       Call ACTUREGISTRO("LOTINGRE", "COD_DESTI", CONDI$, "", REG&)
'''       Call COMUNI("Lote Des-Asignado Correctamente")
'''    End If
    '
End Sub

VERSION 5.00
Begin VB.Form INFOPATRI 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe Patrimonial"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9810
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   210
      Sorted          =   -1  'True
      TabIndex        =   46
      Top             =   1365
      Visible         =   0   'False
      Width           =   4110
   End
   Begin VB.CommandButton Command6 
      Caption         =   "C-Flow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   730
      Left            =   8925
      Picture         =   "INFOPATRI.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Flujo de Caja Proyectado"
      Top             =   5460
      Width           =   730
   End
   Begin VB.CommandButton Command5 
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
      Height          =   730
      Left            =   8925
      Picture         =   "INFOPATRI.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Listado por Rango de Fechas"
      Top             =   3990
      Width           =   730
   End
   Begin VB.CommandButton Command4 
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
      Height          =   730
      Left            =   8925
      Picture         =   "INFOPATRI.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Consulta de Informe Patrimonial por Fecha"
      Top             =   3255
      Width           =   730
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Crédito - Debito Fiscal IVA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   4560
      TabIndex        =   38
      Top             =   210
      Width           =   4125
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Crédito Fiscal:"
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
         Left            =   -105
         TabIndex        =   42
         Top             =   1065
         Width           =   2325
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   2415
         TabIndex        =   41
         Top             =   945
         Width           =   1455
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Debito Fiscal:"
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
         Left            =   -120
         TabIndex        =   40
         Top             =   630
         Width           =   2325
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   2400
         TabIndex        =   39
         Top             =   510
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cálculo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   730
      Left            =   8925
      Picture         =   "INFOPATRI.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Re-Calcula Resúmenes Mensuales"
      Top             =   2520
      Width           =   730
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
      Height          =   730
      Left            =   8925
      Picture         =   "INFOPATRI.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1785
      Width           =   730
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Situación al:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   240
      TabIndex        =   31
      Top             =   240
      Width           =   4125
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   320
         Left            =   3120
         TabIndex        =   35
         Top             =   440
         Width           =   735
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Height          =   320
         Left            =   1080
         TabIndex        =   34
         Top             =   440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hora:"
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
         Left            =   1380
         TabIndex        =   33
         Top             =   525
         Width           =   1590
      End
      Begin VB.Label Label2 
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
         Index           =   13
         Left            =   -1260
         TabIndex        =   32
         Top             =   525
         Width           =   2220
      End
   End
   Begin VB.CommandButton Command2 
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
      Height          =   730
      Left            =   8925
      Picture         =   "INFOPATRI.frx":106A
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Impresión - Vuelco de Pantalla"
      Top             =   4725
      Width           =   730
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Resumen Patrimonial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   4560
      TabIndex        =   27
      Top             =   4305
      Width           =   4125
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   2400
         TabIndex        =   29
         Top             =   510
         Width           =   1455
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total General:"
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
         Left            =   -120
         TabIndex        =   28
         Top             =   630
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Inventario de Materiales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   4560
      TabIndex        =   24
      Top             =   3150
      Width           =   4125
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total de Stocks:"
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
         Left            =   -120
         TabIndex        =   26
         Top             =   630
         Width           =   2325
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   2415
         TabIndex        =   25
         Top             =   525
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pedidos Pendientes de Entrega"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   4560
      TabIndex        =   21
      Top             =   1890
      Width           =   4125
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   2400
         TabIndex        =   23
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total P.P.E.:"
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
         Left            =   -120
         TabIndex        =   22
         Top             =   735
         Width           =   2325
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cuentas Corrientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2020
      Left            =   240
      TabIndex        =   14
      Top             =   4200
      Width           =   4125
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Deudores por Ventas:"
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
         Left            =   -120
         TabIndex        =   20
         Top             =   600
         Width           =   2325
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Height          =   320
         Left            =   2400
         TabIndex        =   19
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Acreedores Comerciales:"
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
         Left            =   -120
         TabIndex        =   18
         Top             =   960
         Width           =   2325
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   2400
         TabIndex        =   17
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Neto:"
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
         Left            =   -120
         TabIndex        =   16
         Top             =   1560
         Width           =   2325
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   2400
         TabIndex        =   15
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.CommandButton command1 
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
      Height          =   730
      Left            =   8925
      Picture         =   "INFOPATRI.frx":16D4
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Abandona la Aplicación"
      Top             =   320
      Width           =   730
   End
   Begin VB.CommandButton Command7 
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
      Height          =   730
      Left            =   8925
      Picture         =   "INFOPATRI.frx":181E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1050
      Width           =   730
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Disponibilidades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2520
      Left            =   240
      TabIndex        =   3
      Top             =   1575
      Width           =   4125
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   2400
         TabIndex        =   13
         Top             =   1890
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Disponibilidades:"
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
         Left            =   -120
         TabIndex        =   12
         Top             =   1995
         Width           =   2325
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   2400
         TabIndex        =   11
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valores en Cartera:"
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
         Left            =   -120
         TabIndex        =   10
         Top             =   1320
         Width           =   2325
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   2400
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldos Bancarios:"
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
         Left            =   -120
         TabIndex        =   8
         Top             =   960
         Width           =   2325
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   2415
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Efectivo en Caja:"
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
         Left            =   -120
         TabIndex        =   4
         Top             =   600
         Width           =   2325
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Avance de Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4560
      TabIndex        =   0
      Top             =   5460
      Width           =   4125
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   315
         ScaleHeight     =   75
         ScaleWidth      =   3465
         TabIndex        =   1
         Top             =   420
         Width           =   3530
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   12000
         End
      End
   End
End
Attribute VB_Name = "INFOPATRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click()
   HOII% = HOY(HO$)
   DESX% = FECHANUM("01" + Mid$(HO$, 3))
   VDEF$ = MKI$(DESX%) + MKI$(HOII%)
   OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
   If Len(OBX$) >= 4 Then
      DES% = CVI(Left$(OBX$, 2))
      HAS% = CVI(Mid$(OBX$, 3, 2))
      If HAS% >= DES% Then
         Screen.MousePointer = 11
         JJ& = 0
         FIN& = ULTREG&("INFOPAT")
         Call COPYSTRU("INFOPAT", "INFOPAL")
         For U& = 1 To FIN&
            Call TRACE(20, U&, FIN&)
            X$ = REGLEIDO$("INFOPAT", U&)
            FF% = CVI(Left$(X$, 2))
            If FF% >= DES% Then
               If FF% <= HAS% Then
                  JJ& = JJ& + 1
                  Call GRAREG("INFOPAL", X$, JJ&)
               End If
            End If
         Next U&
         Call SETULTREG("INFOPAL", JJ&)
         Call CIERRARCH("INFOPAL")
         Call CIERRARCH("*.*")
         Screen.MousePointer = 1
         If JJ& > 0 Then
              Call CONECRUN("*SOINFPA", "Listado de Informe Patrimonial")
            Else
               Call MENSERR(24, "No Hay Informes Registrados\en el rango de Fechas Elegido.")
         End If
      End If
   End If
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call CALFLUCAPRO
   Command6.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    HOI% = HOY(HOS$)
    Label21.Caption = HOS$
    Label22.Caption = Left$(Time$, 5)
    Show
    Refresh
    '
End Sub
'
Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub
'
Private Sub Command1_Click()
    If Command3.Enabled = False Then
        HOI% = HOY(HOS$)
        Label21.Caption = HOS$
        Label22.Caption = Left$(Time$, 5)
        Label1.Caption = ""
        Label4.Caption = ""
        Label6.Caption = ""
        Label8.Caption = ""
        Label15.Caption = ""
        Label13.Caption = ""
        Label9.Caption = ""
        Label23.Caption = ""
        Label25.Caption = ""
        Label12.Caption = ""
        Label17.Caption = ""
        Label20.Caption = ""
        Command3.Enabled = True
        Call FRESHALL
        Exit Sub
    End If
    '
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command2_Click()
   PRINTFORM
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call CALINFPATRI
    Call GUARDINFPAT
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Call SELECHO("SELFECHA")
   FEINF$ = VALACT1$("SELFECHA")
   FF% = FECHANUM(FEINF$)
   If FF% > 0 Then
      For U& = 1 To ULTREG&("INFOPAT")
         X$ = REGLEIDO$("INFOPAT", U&)
         If CVI(Left$(X$, 2)) = FF% Then
            Call MUEINFPAT(X$)
            Exit Sub
         End If
      Next U&
      Call MENSERR(24, "No Hay Informe Registrado\para la Fecha Elegida")
   End If
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   OPINFOPAT.Show 1
   Command14.Enabled = True
End Sub

Sub CALINFPATRI()
   '
   HOI% = HOY(HOS$)
   Dim CUECA%(10), CUEBA%(10)
   '
   CUECAJA1$ = AJUSTI$(Control$("", "CUECAJA1"), 12)
   CUECAJA2$ = AJUSTI$(Control$("", "CUECAJA2"), 12)
   CUECAJA3$ = AJUSTI$(Control$("", "CUECAJA3"), 12)
   CUECAJA4$ = AJUSTI$(Control$("", "CUECAJA4"), 12)
   CUEBANC1$ = AJUSTI$(Control$("", "CUEBANC1"), 12)
   CUEBANC2$ = AJUSTI$(Control$("", "CUEBANC2"), 12)
   CUEBANC3$ = AJUSTI$(Control$("", "CUEBANC3"), 12)
   CUEBANC4$ = AJUSTI$(Control$("", "CUEBANC4"), 12)
   '
   For U& = 1 To ULTREG&("ECUECON")
      XJ$ = REGLEIDO$("ECUECON", U&)
      If Left$(XJ$, 12) = CUECAJA1$ Then CUECA%(1) = CVI(Mid$(XJ$, 41, 2))
      If Left$(XJ$, 12) = CUECAJA2$ Then CUECA%(2) = CVI(Mid$(XJ$, 41, 2))
      If Left$(XJ$, 12) = CUECAJA3$ Then CUECA%(3) = CVI(Mid$(XJ$, 41, 2))
      If Left$(XJ$, 12) = CUECAJA4$ Then CUECA%(4) = CVI(Mid$(XJ$, 41, 2))
      If Left$(XJ$, 12) = CUEBANC1$ Then CUEBA%(1) = CVI(Mid$(XJ$, 41, 2))
      If Left$(XJ$, 12) = CUEBANC2$ Then CUEBA%(2) = CVI(Mid$(XJ$, 41, 2))
      If Left$(XJ$, 12) = CUEBANC3$ Then CUEBA%(3) = CVI(Mid$(XJ$, 41, 2))
      If Left$(XJ$, 12) = CUEBANC4$ Then CUEBA%(4) = CVI(Mid$(XJ$, 41, 2))
   Next U&
   '
   SALCA# = 0: SALBA# = 0
   NUCAJ& = ULTREG&("SUCAJA0")
   For I& = 1 To NUCAJ&
      Call TRACE(20, I&, NUCAJ&)
      SS0$ = REGLEIDO$("SUCAJA0", I&)
      CUE% = CVI(Left$(SS0$, 2))
      FFI% = CVI(Mid$(SS0$, 11, 2))
      IMX# = CVD(Mid$(SS0$, 13, 8))
      If FFI% <= HOI% Then
         For JJ% = 1 To 4
            If CUE% = CUECA%(JJ%) Then
               SALCA# = SALCA# + IMX#
               GoTo 19
            End If
            If CUE% = CUEBA%(JJ%) Then
               SALBA# = SALBA# + IMX#
               GoTo 19
            End If
         Next JJ%
      End If
19 Next I&
   '
   NUCAJ& = ULTREG&("DIARIO0")
   For I& = 1 To NUCAJ&
      Call TRACE(20, I&, NUCAJ&)
      SS0$ = REGLEIDO$("DIARIO0", I&)
      CUE% = CVI(Left$(SS0$, 2))
      FFI% = CVI(Mid$(SS0$, 11, 2))
      IMX# = CVD(Mid$(SS0$, 13, 8))
      If FFI% <= HOI% Then
         If Asc(Mid$(SS0$, 23, 1)) <> 1 Then
            For JJ% = 1 To 4
               If CUE% = CUECA%(JJ%) Then
                  SALCA# = SALCA# + IMX#
                  GoTo 29
               End If
               If CUE% = CUEBA%(JJ%) Then
                  SALBA# = SALBA# + IMX#
                  GoTo 29
               End If
            Next JJ%
         End If
      End If
29 Next I&
   '
   INFOPATRI.Label1.Caption = TRIM$(CSTRING$(MKD$(SALCA#), 4, 12, 2, 2))
   INFOPATRI.Label1.Refresh
   INFOPATRI.Label4.Caption = TRIM$(CSTRING$(MKD$(SALBA#), 4, 12, 2, 2))
   INFOPATRI.Label4.Refresh
   GoSub 99
   '
   SALDEUTO# = 0
   NUSAL& = ULTREG&("SVD202")
   For I& = 1 To NUSAL&
      Call TRACE(20, I&, NUSAL&)
      SS0$ = REGLEIDO$("SVD202", I&)
      SS2$ = Mid$(SS0$, 3, 2)
      SS4$ = Mid$(SS0$, 9, 2)
      SS5$ = Mid$(SS0$, 11, 8)
      '
      FVII% = CVI(SS4$)
      If FVII% > 0 Then
         TI% = Asc(Mid$(SS2$, 1, 1))
         If TI% > 6 Then TI% = 2
         SALDEU# = CVD(SS5$)
         If TI% > 3 Then SALDEU# = (-1) * SALDEU#
         '
         If Abs(SALDEU#) >= 0.005 Then
            SALDEUTO# = SALDEUTO# + SALDEU#
         End If
       End If
   Next I&
   '
   INFOPATRI.Label15.Caption = TRIM$(CSTRING$(MKD$(SALDEUTO#), 4, 12, 2, 2))
   INFOPATRI.Label15.Refresh
   GoSub 99
   '
   SALACTO# = 0
   Dim TICUEPRO%(32767)
   '
   FIN& = ULTREG&("PROVED1")
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("PROVED1", I&)
      NCI% = CVI(Left$(X$, 2))
      If NCI% > 0 Then
         If NCI% <= 32767 Then
            X$ = REGLEIDO$("PROVED2", I&)
            If Asc(Mid$(X$, 119, 1)) = 0 Then
               TICUEPRO%(NCI%) = 1
            End If
         End If
      End If
   Next I&
   '
   FINI% = COMESACT%
   If Val(Left$(HOS$, 2)) <= 9 Then
      FINI% = COMESANT%
   End If
   '
   DEFIVA# = 0: CREIVA# = 0
   '
   NUSAL& = ULTREG&("CUECORR")
   For I& = 1 To NUSAL&
      Call TRACE(20, I&, NUSAL&)
      SS0$ = REGLEIDO$("CUECORR", I&)
      NCI% = CVI(Mid$(SS0$, 5, 2))
      If NCI% > 0 Then
         If NCI% < 32767 Then
            If TICUEPRO%(NCI%) = 1 Then
               SS8$ = Mid$(SS0$, 139, 60)
               '
               For V% = 1 To 6
                  FVI% = CVI(Mid$(SS8$, 10 * V% - 9, 2))
                  If FVI% > 0 Then
                     SALDEU# = CVD(Mid$(SS8$, 10 * V% - 7, 8))
                     If Abs(SALDEU#) >= 0.005 Then
                        SALACTO# = SALACTO# + Int(100 * SALDEU# + 0.5) / 100
                     End If
                  End If
               Next V%
               '
               If CVI(Left$(SS0$, 2)) >= FINI% Then
                  NF$ = UCase$(Mid$(SS0$, 7, 18))
                  If InStr(NF$, "ANUL") < 1 Then
                     TIX$ = UCase$(Mid$(SS0$, 7, 2))
                     If TIX$ <> "OP" Then
                        If TIX$ <> "AS" Then
                           SS7$ = Mid$(SS0$, 51, 64)
                           IIVA# = CVD(Mid$(SS7$, 9, 8))
                           IIVA# = IIVA# + CVD(Mid$(SS7$, 17, 8))
                           IIVA# = IIVA# + CVD(Mid$(SS7$, 25, 8))
                           IIVA# = IIVA# + CVD(Mid$(SS7$, 41, 8))
                           SIG% = 1
                           If TIX$ = "NC" Or TIX$ = "DC" Or TIX$ = "DP" Then
                              SIG% = (-1)
                           End If
                           CREIVA# = CREIVA# + (Abs(IIVA#)) * SIG%
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
   Next I&
   '
   INFOPATRI.Label13.Caption = TRIM$(CSTRING$(MKD$(SALACTO#), 4, 12, 2, 2))
   INFOPATRI.Label13.Refresh
   '
   INFOPATRI.Label9.Caption = TRIM$(CSTRING$(MKD$(SALDEUTO# - SALACTO#), 4, 12, 2, 2))
   INFOPATRI.Label9.Refresh
   '
   INFOPATRI.Label25.Caption = TRIM$(CSTRING$(MKD$(CREIVA#), 4, 12, 2, 2))
   INFOPATRI.Label25.Refresh
   GoSub 99
   '
   SUCHECAR# = 0
   FIN& = ULTREG&("BADACHE")
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("BADACHE", I&)
      If CVI(Mid$(X$, 133, 2)) < 1 Then
         SUCHECAR# = SUCHECAR# + CVD(Mid$(X$, 49, 8))
      End If
   Next I&
   '
   INFOPATRI.Label6.Caption = TRIM$(CSTRING$(MKD$(SUCHECAR#), 4, 12, 2, 2))
   INFOPATRI.Label6.Refresh
   GoSub 99
   '
   '
   NUSAL& = ULTREG&("FACTUR")
   For I& = 1 To NUSAL&
      Call TRACE(20, I&, NUSAL&)
      F0$ = REGLEIDO$("FACTUR", I&)
      F3$ = Mid$(F0$, 7, 2)
      FII% = CVI(F3$)
      '
      If CVI(F3$) >= FINI% Then
         F1$ = Left$(F0$, 2)
         F6$ = Mid$(F0$, 13, 2)
         TI% = Asc(F6$): If TI% < 0 Or TI% > 6 Then TI% = 2
         SIG% = 1: If TI% > 3 Then SIG% = (-1)
         '
         F8$ = Mid$(F0$, 23, 8)
         F9$ = Mid$(F0$, 31, 8)
         F12$ = Mid$(F0$, 49, 8)
         F13$ = Mid$(F0$, 57, 8)
         '
         DEFIVA# = DEFIVA# + (Abs(CVD(F8$) + CVD(F9$))) * SIG%
         '
      End If
   Next I&
   '
   INFOPATRI.Label23.Caption = TRIM$(CSTRING$(MKD$(DEFIVA#), 4, 12, 2, 2))
   INFOPATRI.Label23.Refresh
   GoSub 99
   '
   FIN& = NUMAS%
   VASTO# = 0
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      CIII% = I&
      ST = STOCKACT(CIII%)
      COU = COSUNI(CIII%)
      VASTO# = VASTO# + ST * COU
   Next I&
   '
   INFOPATRI.Label17.Caption = TRIM$(CSTRING$(MKD$(VASTO#), 4, 12, 2, 2))
   INFOPATRI.Label17.Refresh
   GoSub 99
   '
   FIN& = ULTREG&("PEDDETP")
   VAPED# = 0
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("PEDDETP", I&)
      SALPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
      If SALPEN > 0 Then
         NC% = CVI(Mid$(X$, 7, 2))
         VAPED# = VAPED# + (SALPEN * CVS(Mid$(X$, 83, 4))) * ((100 - PORDESCLI(NC%)) / 100)
      End If
   Next I&
   '
   INFOPATRI.Label12.Caption = TRIM$(CSTRING$(MKD$(VAPED#), 4, 12, 2, 2))
   INFOPATRI.Label12.Refresh
   GoSub 99
   '
   Exit Sub
   '
99 TODISPO# = SALCA# + SALBA# + SUCHECAR#
   INFOPATRI.Label8.Caption = TRIM$(CSTRING$(MKD$(TODISPO#), 4, 12, 2, 2))
   INFOPATRI.Label8.Refresh
   '
   RESUPATRI# = TODISPO# + SALDEUTO# - SALACTO#
   RESUPATRI# = RESUPATRI# - DEFIVA + CREIVA#
   RESUPATRI# = RESUPATRI# + VASTO#
   RESUPATRI# = RESUPATRI# + VAPED#
   '
   INFOPATRI.Label20.Caption = TRIM$(CSTRING$(MKD$(RESUPATRI#), 4, 12, 2, 2))
   INFOPATRI.Label20.Refresh
   Return
   '
End Sub
'
Sub GUARDINFPAT()
   '
   Screen.MousePointer = 11
   Z$ = String$(128, 0)
   FF% = FECHANUM(Label21.Caption)
   If FF% < 1 Then Exit Sub
   HG% = HORANUM%(Label22.Caption)
   '
   Call REPLA(Z$, MKI$(FF%), 1, 2)
   Call REPLA(Z$, MKI$(HG%), 3, 2)
   Call REPLA(Z$, MKD$(Val(Label1.Caption)), 9, 8)
   Call REPLA(Z$, MKD$(Val(Label4.Caption)), 17, 8)
   Call REPLA(Z$, MKD$(Val(Label6.Caption)), 25, 8)
   Call REPLA(Z$, MKD$(Val(Label8.Caption)), 33, 8)
   Call REPLA(Z$, MKD$(Val(Label15.Caption)), 41, 8)
   Call REPLA(Z$, MKD$(Val(Label13.Caption)), 49, 8)
   Call REPLA(Z$, MKD$(Val(Label23.Caption)), 57, 8)
   Call REPLA(Z$, MKD$(Val(Label25.Caption)), 65, 8)
   Call REPLA(Z$, MKD$(Val(Label12.Caption)), 73, 8)
   Call REPLA(Z$, MKD$(Val(Label17.Caption)), 81, 8)
   Call REPLA(Z$, MKD$(Val(Label20.Caption)), 121, 8)
   '
   For U& = 1 To ULTREG&("INFOPAT")
      X$ = REGLEIDO$("INFOPAT", U&)
      If CVI(Left$(X$, 2)) = FF% Then
         Call GRAREG("INFOPAT", Z$, U&)
         GoTo 99
      End If
   Next U&
   Call REGAPP("INFOPAT", Z$)
   '
99 Call CIERRARCH("INFOPAT")
   Screen.MousePointer = 1
   '
End Sub
'
Sub MUEINFPAT(X$)
   '
   Command3.Enabled = False
   '
   FEX = CVI(Left$(X$, 2))
   HG% = CVI(Mid$(X$, 3, 2))
   Label21.Caption = FECHATEX$(FEX)
   Label22.Caption = Left$(HORATEX$(HG%), 5)
   '
   SALCA# = CVD(Mid$(X$, 9, 8))
   SALBA# = CVD(Mid$(X$, 17, 8))
   Label1.Caption = TRIM$(CSTRING$(MKD$(SALCA#), 4, 12, 2, 2))
   Label1.Refresh
   Label4.Caption = TRIM$(CSTRING$(MKD$(SALBA#), 4, 12, 2, 2))
   Label4.Refresh
   GoSub 99
   '
   SALDEUTO# = CVD(Mid$(X$, 41, 8))
   Label15.Caption = TRIM$(CSTRING$(MKD$(SALDEUTO#), 4, 12, 2, 2))
   Label15.Refresh
   GoSub 99
   '
   SALACTO# = CVD(Mid$(X$, 49, 8))
   DEFIVA# = CVD(Mid$(X$, 57, 8))
   CREIVA# = CVD(Mid$(X$, 65, 8))
   Label13.Caption = TRIM$(CSTRING$(MKD$(SALACTO#), 4, 12, 2, 2))
   Label13.Refresh
   '
   Label9.Caption = TRIM$(CSTRING$(MKD$(SALDEUTO# - SALACTO#), 4, 12, 2, 2))
   Label9.Refresh
   '
   Label25.Caption = TRIM$(CSTRING$(MKD$(CREIVA#), 4, 12, 2, 2))
   Label25.Refresh
   GoSub 99
   '
   SUCHECAR# = CVD(Mid$(X$, 25, 8))
   Label6.Caption = TRIM$(CSTRING$(MKD$(SUCHECAR#), 4, 12, 2, 2))
   Label6.Refresh
   GoSub 99
   '
   Label23.Caption = TRIM$(CSTRING$(MKD$(DEFIVA#), 4, 12, 2, 2))
   Label23.Refresh
   GoSub 99
   '
   VASTO# = CVD(Mid$(X$, 81, 8))
   Label17.Caption = TRIM$(CSTRING$(MKD$(VASTO#), 4, 12, 2, 2))
   Label17.Refresh
   GoSub 99
   '
   VAPED# = CVD(Mid$(X$, 73, 8))
   Label12.Caption = TRIM$(CSTRING$(MKD$(VAPED#), 4, 12, 2, 2))
   Label12.Refresh
   GoSub 99
   '
   Exit Sub
   '
99 TODISPO# = SALCA# + SALBA# + SUCHECAR#
   Label8.Caption = TRIM$(CSTRING$(MKD$(TODISPO#), 4, 12, 2, 2))
   Label8.Refresh
   '
   RESUPATRI# = TODISPO# + SALDEUTO# - SALACTO#
   RESUPATRI# = RESUPATRI# - DEFIVA + CREIVA#
   RESUPATRI# = RESUPATRI# + VASTO#
   RESUPATRI# = RESUPATRI# + VAPED#
   '
   Label20.Caption = TRIM$(CSTRING$(MKD$(RESUPATRI#), 4, 12, 2, 2))
   Label20.Refresh
   Return
   '
End Sub

Sub CALFLUCAPRO()    ' calculo flujo de caja proyectado
   '
   List1.Clear
   HOI% = HOY(HOS$)
   Dim CUECA%(16), CUECAJA$(16), DENCA$(16), SALCA#(16)
   '
   Dim TIS$(10)
   TDAT$ = "AS,VC,FC,ND,RB,DC,NC,AS"
   For JJ% = 0 To 7
      TIS$(JJ%) = Mid$(TDAT$, 3 * JJ% + 1, 2)
   Next JJ%
   '
   Dim TICUEPRO%(32767)
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("PROVED1")
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("PROVED1", I&)
      NCI% = CVI(Left$(X$, 2))
      If NCI% > 0 Then
         If NCI% <= 32767 Then
            X$ = REGLEIDO$("PROVED2", I&)
            If Asc(Mid$(X$, 119, 1)) = 0 Then
               TICUEPRO%(NCI%) = 1
            End If
         End If
      End If
   Next I&
   '
   Screen.MousePointer = 11
   CUECAJA$(1) = AJUSTI$(Control$("", "CUECAJA1"), 12)
   CUECAJA$(2) = AJUSTI$(Control$("", "CUECAJA2"), 12)
   CUECAJA$(3) = AJUSTI$(Control$("", "CUECAJA3"), 12)
   CUECAJA$(4) = AJUSTI$(Control$("", "CUECAJA4"), 12)
   CUECAJA$(5) = AJUSTI$(Control$("", "CUEBANC1"), 12)
   CUECAJA$(6) = AJUSTI$(Control$("", "CUEBANC2"), 12)
   CUECAJA$(7) = AJUSTI$(Control$("", "CUEBANC3"), 12)
   CUECAJA$(8) = AJUSTI$(Control$("", "CUEBANC4"), 12)
   '
   FIN& = ULTREG&("ECUECON")
   Screen.MousePointer = 11
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XJ$ = REGLEIDO$("ECUECON", U&)
      XJ1$ = Left$(XJ$, 12)
      For JJ% = 1 To 8
         If CUECAJA$(JJ%) = XJ1$ Then
            CUECA%(JJ%) = CVI(Mid$(XJ$, 41, 2))
            DENCA$(JJ%) = Mid$(XJ$, 13, 28)
         End If
      Next JJ%
   Next U&
   '
   NUCAJ& = ULTREG&("SUCAJA0")
   Screen.MousePointer = 11
   For I& = 1 To NUCAJ&
      Call TRACE(20, I&, NUCAJ&)
      SS0$ = REGLEIDO$("SUCAJA0", I&)
      CUE% = CVI(Left$(SS0$, 2))
      FFI% = CVI(Mid$(SS0$, 11, 2))
      IMX# = CVD(Mid$(SS0$, 13, 8))
      If FFI% <= HOI% Then
         For JJ% = 1 To 8
            If CUE% = CUECA%(JJ%) Then
               SALCA#(JJ%) = SALCA#(JJ%) + IMX#
               GoTo 19
            End If
         Next JJ%
      End If
19 Next I&
   '
   NUCAJ& = ULTREG&("DIARIO0")
   Screen.MousePointer = 11
   For I& = 1 To NUCAJ&
      Call TRACE(20, I&, NUCAJ&)
      SS0$ = REGLEIDO$("DIARIO0", I&)
      CUE% = CVI(Left$(SS0$, 2))
      FFI% = CVI(Mid$(SS0$, 11, 2))
      IMX# = CVD(Mid$(SS0$, 13, 8))
      If FFI% <= HOI% Then
         If Asc(Mid$(SS0$, 23, 1)) <> 1 Then
            For JJ% = 1 To 8
               If CUE% = CUECA%(JJ%) Then
                  SALCA#(JJ%) = SALCA#(JJ%) + IMX#
                  GoTo 29
               End If
            Next JJ%
         End If
      End If
29 Next I&
   '
   For JJ% = 1 To 8
      If CUECA%(JJ%) > 0 Then
         Call AGREFLUCA(HOI%, "CAJA-BANCOS", DENCA$(JJ%), SALCA#(JJ%))
      End If
   Next JJ%
   '
   FIN& = ULTREG&("BADACHE")
   Screen.MousePointer = 11
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("BADACHE", I&)
      If CVI(Mid$(X$, 133, 2)) < 1 Then
         FEMO% = CVI(Mid$(X$, 59, 2))
         If FEMO% < HOI% Then FEMO% = HOI%
         REFE$ = TRIM$(Left$(X$, 18))
         If TRIM$(Mid$(X$, 19, 14)) <> "" Then
            REFE$ = REFE$ + " - " + TRIM$(Mid$(X$, 19, 14))
         End If
         If TRIM$(Mid$(X$, 33, 8)) <> "" Then
            REFE$ = REFE$ + " - " + TRIM$(Mid$(X$, 33, 8))
         End If
         IMPO# = CVD(Mid$(X$, 49, 8))
         If Abs(IMPO#) > 0.005 Then
            Call AGREFLUCA(FEMO%, "CH.CARTERA", REFE$, IMPO#)
         End If
      End If
   Next I&
   '
   FIN& = ULTREG&("CHEQUEMI")
   Screen.MousePointer = 11
   FLIMEM% = HOI% - 32 * 13
   '
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("CHEQUEMI", I&)
      EMIS = CVI(Mid$(X$, 31, 2))
      VENCI% = CVI(Mid$(X$, 49, 2))
      '
      If CVI(Mid$(X$, 45, 2)) <> 999 Then ' excluye cheques terceros
         If VENCI% >= HOI% Then
            If EMIS >= FLIMEM% Then
               BANCO$ = Mid$(X$, 15, 16)
               If InStr(BANCO$, "EFECT") < 1 Then
                  PROV% = CVI(Left$(X$, 2))
                  DESTI$ = TRIM$(RASOCLI$((-1) * PROV%))
                  NUMERO$ = TRIM$(Mid$(X$, 3, 12))
                  IMPO# = CVD(Mid$(X$, 33, 8))
                  FEMO% = VENCI%
                  REFE$ = TRIM$(BANCO$)
                  If NUMERO$ <> "" Then REFE$ = REFE$ + " - " + NUMERO$
                  If DESTI$ <> "" Then REFE$ = REFE$ + " - " + DESTI$
                  If Abs(IMPO#) > 0.005 Then
                     Call AGREFLUCA(FEMO%, "CH.PROPIO", REFE$, (-1) * IMPO#)
                  End If
               End If
            End If
         End If
      End If
   Next I&
   '
   NUSAL& = ULTREG&("SVD202")
   Screen.MousePointer = 11
   For I& = 1 To NUSAL&
      Call TRACE(20, I&, NUSAL&)
      SS0$ = REGLEIDO$("SVD202", I&)
      SS2$ = Mid$(SS0$, 3, 2)
      SS4$ = Mid$(SS0$, 9, 2)
      SS5$ = Mid$(SS0$, 11, 8)
      '
      FVII% = CVI(SS4$)
      If FVII% > 0 Then
         IMPO# = CVD(SS5$)
         If Abs(IMPO#) > 0.005 Then
            NC% = CVI(Left$(SS0$, 2))
            GCLI% = GRUCOCLI%(NC%)
            If GCLI% <> 9 Then      ' excluye incobrables
               TI% = Asc(Mid$(SS2$, 1, 1))
               If TI% > 6 Then TI% = 2
               If TI% > 3 Then IMPO# = (-1) * IMPO#
               NRO& = CVS(Mid$(SS0$, 5, 4))
               DESTI$ = "": If NC% > 0 Then DESTI$ = RASOCLI$(NC%)
               '
               FEMO% = FVII%
               If FEMO% < HOI% Then FEMO% = HOI%
               REFE$ = TIS$(TI%) + "." + TRIM$(Str$(NRO&))
               If DESTI$ <> "" Then REFE$ = REFE$ + " - " + DESTI$
               Call AGREFLUCA(FEMO%, "DEUDORES C.C.", REFE$, IMPO#)
            End If
         End If
      End If
   Next I&
   '
   Screen.MousePointer = 11
   NUSAL& = ULTREG&("CUECORR")
   For I& = 1 To NUSAL&
      Call TRACE(20, I&, NUSAL&)
      SS0$ = REGLEIDO$("CUECORR", I&)
      NCI% = CVI(Mid$(SS0$, 5, 2))
      If NCI% > 0 Then
        If NCI% < 32767 Then
          If TICUEPRO%(NCI%) = 1 Then
            SS8$ = Mid$(SS0$, 139, 60)
            NF$ = UCase$(Mid$(SS0$, 7, 18))
            If InStr(NF$, "ANUL") < 1 Then
               '
               For V% = 1 To 6
                  FVI% = CVI(Mid$(SS8$, 10 * V% - 9, 2))
                  If FVI% > 0 Then
                     SALDEU# = CVD(Mid$(SS8$, 10 * V% - 7, 8))
                     If Abs(SALDEU#) >= 0.005 Then
                        IMPO# = (-1) * SALDEU#
                        FEMO% = FVI%
                        If FEMO% < HOI% Then FEMO% = HOI%
                        REFE$ = TRIM$(NF$)
                        DESTI$ = TRIM$(RASOCLI$((-1) * NCI%))
                        If DESTI$ <> "" Then REFE$ = REFE$ + " - " + DESTI$
                        Call AGREFLUCA(FEMO%, "PROVED.C.C.", REFE$, IMPO#)
                     End If
                  End If
               Next V%
               '
               If CVI(Left$(SS0$, 2)) >= FINI% Then
                     TIX$ = UCase$(Mid$(SS0$, 7, 2))
                     If TIX$ <> "OP" Then
                        If TIX$ <> "AS" Then
                           SS7$ = Mid$(SS0$, 51, 64)
                           IIVA# = CVD(Mid$(SS7$, 9, 8))
                           IIVA# = IIVA# + CVD(Mid$(SS7$, 17, 8))
                           IIVA# = IIVA# + CVD(Mid$(SS7$, 25, 8))
                           IIVA# = IIVA# + CVD(Mid$(SS7$, 41, 8))
                           SIG% = 1
                           If TIX$ = "NC" Or TIX$ = "DC" Or TIX$ = "DP" Then
                              SIG% = (-1)
                           End If
                           CREIVA# = CREIVA# + (Abs(IIVA#)) * SIG%
                        End If
                     End If
               End If
             End If
           End If
         End If
      End If
   Next I&
   '
   FIN& = ULTREG&("PEDDETP")
   VAPED# = 0
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("PEDDETP", I&)
      SALPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
      If SALPEN > 0 Then
         NC% = CVI(Mid$(X$, 7, 2))
         VAPED# = (SALPEN * CVS(Mid$(X$, 83, 4))) * ((100 - PORDESCLI(NC%)) / 100)
         ALIV = ALIVA: If PIVACLI%(NC%) = 2 Then ALIV = ALIVA + ALIVARNI
         VAPED# = VAPED# * (100 + ALIV) / 100
         FPED% = CVI(Mid$(X$, 25, 2))
         If FPED% < HOI% Then FPED% = HOI%
         NROPED& = CVS(Left$(X$, 4))
         COPAGPED% = Asc(Mid$(X$, 29, 1))
         Call CALVENPED(NC%, FPED%, NROPED&, VAPED#, COPAGPED%)
      End If
   Next I&
   '
   '
   WW& = 0: ACUMU# = 0
   X0$ = REGBLAN$("FLUCAPRO")
   FIN& = List1.ListCount
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = List1.List(U& - 1)
      X$ = X0$
      Call REPLA(X$, MKI$(Val(Left$(Z$, 6))), 1, 2)
      Call REPLA(X$, Mid$(Z$, 9, 14), 3, 14)
      Call REPLA(X$, Mid$(Z$, 25, 48), 17, 48)
      Call REPLA(X$, MKD$(Val(Mid$(Z$, 73, 12))), 65, 8)
      Call REPLA(X$, MKD$(Val(Mid$(Z$, 85, 12))), 73, 8)
      ACUMU# = ACUMU# + Val(Mid$(Z$, 73, 12)) - Val(Mid$(Z$, 85, 12))
      Call REPLA(X$, MKD$(ACUMU#), 81, 8)
      WW& = WW& + 1
      Call GRAREG("FLUCAPRO", X$, WW&)
   Next U&
   Call SETULTREG("FLUCAPRO", WW&)
   Call CIERRARCH("FLUCAPRO")
   '
   Call HEADERS("FLUCAPRO", "")
   Call HEADERS("FLUCAPRO", "Generado el: " + HOS$ + " - " + Time$ + " Hs.")
   '
   Screen.MousePointer = 1
   MODO% = COMALTER%("Generado Flujo de Caja Proyectado.\   \Elija su Opción de Emisión:\\Pantalla\Impresora\Cancelar")
   Call FRESHALL
   If MODO% = 1 Then
        Call CONECRUN("*COFLUCA", "Flujo de Caja Proyectado (Pantalla)")
      ElseIf MODO% = 2 Then
        Call CONECRUN("*LIFLUCA", "Flujo de Caja Proyectado (Pantalla)")
   End If
   '
End Sub
         
Sub AGREFLUCA(FF%, RUBRO$, CONCE$, IMPO#)
   '
   Z$ = Space$(96)
   Call REPLA(Z$, CSTRING$(MKI$(FF%), 1, 6, 0, 2), 1, 6)
   Call REPLA(Z$, RUBRO$, 9, 14)
   Call REPLA(Z$, CONCE$, 25, 48)
   If IMPO# >= 0 Then
        Call REPLA(Z$, CSTRING$(MKD$(IMPO#), 3, 12, 2, 2), 73, 12)
      Else
        Call REPLA(Z$, CSTRING$(MKD$(Abs(IMPO#)), 3, 12, 2, 2), 85, 12)
   End If
   List1.AddItem Z$
   '
End Sub

Sub CALVENPED(NCLI%, FPED%, NROPED&, TOTPED#, COPAGPED%)
    '
    Static FVI%(6), PLAII%(6), IM#(6)
    COPAGI% = COPAGPED%
    If TRIM$(DECONPA$(COPAGPED%)) = "" Then
       COPAGI% = CPAGCLI%(NCLI%)
    End If
    '
    LU$ = LUDAT$
200 Call BUCONPA(LU$, COPAGI%, COPAG$, FREF%, PLA$)
    
    ' determina vencimientos segun total factura y condicion de pago
    '
300 FF = FPED%: If FF < 1 Then FF = FECHANUM("01/01/80")
    FFS$ = FECHATEX$(FF)
    PLA$ = Left$(PLA$ + String$(6, 0), 6)
    If (FREF% < 1 Or FREF% > 4) Then FREF% = 1
    '
    If FREF% = 2 Then FREF% = 1   ' AUN NO FUNCIONA SEMANA DE FACTURA
    If FREF% <> 3 Then GoTo 305
    If Val(Left$(FFS$, 2)) <= 15 Then FEBAS$ = "15" + Mid$(FFS$, 3)
    If Val(Left$(FFS$, 2)) > 15 Then FEBAS$ = "31" + Mid$(FFS$, 3)
305 If FREF% = 1 Then FEBAS$ = FFS$
    If FREF% = 4 Then FEBAS$ = "31" + Mid$(FFS$, 3)
    '
    For I% = 1 To 6: FVI%(I%) = 0: PLAII%(I%) = 0: Next I%
    For I% = 1 To 6
       PLAI% = Asc(Mid$(PLA$, I%, 1)): PLAII%(I%) = PLAI%
       If I% > 1 Then If (PLAII%(I%) < PLAII%(I% - 1) Or PLAII%(I%) <= 0) Then GoTo 310
       FEBAI = FECHANUM(FEBAS$)
       FEVEI = FEBAI + PLAI%
       While DIENTRE%(FEBAI, FEVEI) < PLAI%
          FEVEI = FEVEI + 1
       Wend
       FVI%(I%) = FEVEI: NUPAG% = I%
    Next I%

310 FVII% = FVI%(NUPAG%)
    If NUPAG% < 1 Then FVII% = FPED%

    ' determina importes segun numero de pagos

400 If NUPAG% < 1 Then NUPAG% = 1
    IM#(1) = (Int(100 * TOTPED# / NUPAG% + 0.5)) / 100
    For H% = 2 To NUPAG% - 1
       IM#(H%) = IM#(1)
    Next H%
    IM#(NUPAG%) = TOTPED#
    For H% = 1 To NUPAG% - 1
       IM#(NUPAG%) = IM#(NUPAG%) - IM#(H%)
    Next H%

410 For I% = 1 To NUPAG%
       If IM#(I%) >= 0.005 Then
          NRO& = NROPED&
          DESTI$ = ""
          If NCLI% > 0 Then DESTI$ = RASOCLI$(NC%)
          '
          FEMO% = FVI%(I%)
          REFE$ = "PC." + TRIM$(Str$(NRO&))
          If DESTI$ <> "" Then REFE$ = REFE$ + " - " + DESTI$
          Call AGREFLUCA(FEMO%, "PEDIDO CLIENTE", REFE$, IM#(I%))
       End If
    Next I%
    '
End Sub

VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form FDIRMAIL09 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libreta de Direcciones de Mails"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9270
   Icon            =   "FDIRMAIL09.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   255
      Left            =   3240
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame7 
      Caption         =   "Alternativa (Empresa)"
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
      Height          =   1050
      Left            =   120
      TabIndex        =   54
      Top             =   3810
      Width           =   3900
      Begin VB.ListBox List3 
         Height          =   645
         Left            =   120
         TabIndex        =   55
         Top             =   320
         Width           =   3690
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Alternativa (Nombre)"
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
      Height          =   1100
      Left            =   105
      TabIndex        =   28
      Top             =   2730
      Width           =   3900
      Begin VB.ListBox List2 
         Height          =   645
         Left            =   105
         TabIndex        =   37
         Top             =   315
         Width           =   3690
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear"
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
      Left            =   8310
      Picture         =   "FDIRMAIL09.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Procesa Mails Rechazados"
      Top             =   3810
      Width           =   750
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   8295
      Picture         =   "FDIRMAIL09.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Generar y Enviar Mail Circular"
      Top             =   4515
      Width           =   750
   End
   Begin VB.Frame Frame3 
      Caption         =   "Identificación"
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
      Height          =   2535
      Left            =   105
      TabIndex        =   15
      Top             =   105
      Width           =   7995
      Begin VB.CommandButton Command6 
         Height          =   330
         Left            =   2835
         Picture         =   "FDIRMAIL09.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Refresca - Quita Marca de Baja"
         Top             =   1995
         Width           =   435
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1155
         TabIndex        =   3
         Text            =   " "
         ToolTipText     =   "Doble-Click para Enviar un Mail"
         Top             =   420
         Width           =   6690
      End
      Begin VB.Frame Frame5 
         Caption         =   "Nro"
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
         Left            =   6300
         TabIndex        =   23
         Top             =   1575
         Width           =   1440
         Begin VB.CommandButton Command27 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   160
            Left            =   1155
            Picture         =   "FDIRMAIL09.frx":0A68
            TabIndex        =   33
            ToolTipText     =   "Selecciona Contacto - Lista Completa"
            Top             =   315
            Width           =   175
         End
         Begin VB.CommandButton Command26 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   160
            Left            =   1155
            Picture         =   "FDIRMAIL09.frx":0D72
            TabIndex        =   32
            ToolTipText     =   "Selecciona Contacto - Lista Completa"
            Top             =   450
            Width           =   175
         End
         Begin VB.CommandButton Command16 
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
            Height          =   285
            Left            =   885
            TabIndex        =   0
            ToolTipText     =   "Selecciona Contacto - Lista Completa"
            Top             =   315
            Width           =   175
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   150
            TabIndex        =   24
            Top             =   315
            Width           =   750
         End
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1155
         TabIndex        =   2
         Text            =   " "
         Top             =   1155
         Width           =   6690
      End
      Begin VB.TextBox Text23 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   4725
         TabIndex        =   6
         Text            =   " "
         Top             =   1995
         Width           =   1485
      End
      Begin VB.TextBox Text22 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   3675
         TabIndex        =   5
         Text            =   " "
         Top             =   840
         Width           =   4170
      End
      Begin VB.TextBox Text21 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1155
         TabIndex        =   4
         Top             =   840
         Width           =   1170
      End
      Begin VB.TextBox Text18 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   1155
         TabIndex        =   7
         Text            =   " "
         Top             =   1680
         Width           =   2115
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "F.Baja:"
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
         Left            =   315
         TabIndex        =   35
         Top             =   2100
         Width           =   750
      End
      Begin VB.Label Label19 
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
         Height          =   285
         Left            =   1155
         TabIndex        =   34
         Top             =   1995
         Width           =   1170
      End
      Begin VB.Label Label2 
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
         Height          =   285
         Left            =   5040
         TabIndex        =   27
         Top             =   1680
         Width           =   1170
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Motivo Baja:"
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
         Left            =   2835
         TabIndex        =   26
         Top             =   2100
         Width           =   1800
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Dir-elec:"
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
         Left            =   105
         TabIndex        =   25
         Top             =   525
         Width           =   960
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Empresa:"
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
         Left            =   105
         TabIndex        =   22
         Top             =   1260
         Width           =   960
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "Persona:"
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
         Left            =   2625
         TabIndex        =   19
         Top             =   945
         Width           =   960
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Fecha Alta:"
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
         Left            =   2835
         TabIndex        =   18
         Top             =   1785
         Width           =   2115
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Apelativo:"
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
         Left            =   105
         TabIndex        =   17
         Top             =   945
         Width           =   960
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Origen:"
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
         Left            =   105
         TabIndex        =   16
         Top             =   1785
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Observaciones"
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
      Height          =   1005
      Left            =   120
      TabIndex        =   39
      Top             =   4935
      Width           =   5475
      Begin VB.TextBox Text1 
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
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   315
         Width           =   5265
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Lista de Listas"
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
      Height          =   2150
      Left            =   4110
      TabIndex        =   38
      Top             =   2730
      Width           =   3990
      Begin VB.ListBox List5 
         Height          =   1635
         Left            =   130
         Style           =   1  'Checkbox
         TabIndex        =   53
         Top             =   360
         Width           =   3690
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Search"
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
      Left            =   8310
      Picture         =   "FDIRMAIL09.frx":107C
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Búsqueda Secuencial por Texto"
      Top             =   2010
      Width           =   750
   End
   Begin VB.CommandButton Command24 
      Caption         =   "Erase"
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
      Left            =   8310
      Picture         =   "FDIRMAIL09.frx":197E
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Marca Contacto de Baja"
      Top             =   2610
      Width           =   750
   End
   Begin VB.CommandButton Command21 
      Caption         =   "New"
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
      Left            =   8310
      Picture         =   "FDIRMAIL09.frx":1C88
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Alta de Nueva Ficha de Seguimiento"
      Top             =   810
      Width           =   750
   End
   Begin VB.CommandButton Command3 
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
      Height          =   600
      Left            =   8310
      Picture         =   "FDIRMAIL09.frx":1DD2
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   210
      Width           =   750
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   8190
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   735
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Sort"
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
      Left            =   8310
      Picture         =   "FDIRMAIL09.frx":1F1C
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Ordenamiento Alfabético"
      Top             =   3210
      Width           =   750
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H80000000&
      Caption         =   "Varios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   5670
      TabIndex        =   20
      Top             =   4935
      Width           =   2430
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
         Height          =   580
         Left            =   105
         Picture         =   "FDIRMAIL09.frx":2226
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   315
         Width           =   540
      End
      Begin VB.CommandButton Command25 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   580
         Left            =   650
         Picture         =   "FDIRMAIL09.frx":2668
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Imprime Ficha del Contacto"
         Top             =   315
         Width           =   540
      End
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
         Height          =   580
         Left            =   1195
         Picture         =   "FDIRMAIL09.frx":2CD2
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Listados por Impresora"
         Top             =   315
         Width           =   540
      End
      Begin VB.CommandButton Command20 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   580
         Left            =   1740
         TabIndex        =   30
         ToolTipText     =   "Leer e Importar"
         Top             =   315
         Width           =   540
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
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
      Left            =   8310
      Picture         =   "FDIRMAIL09.frx":2FDC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Guarda los Cambios Realizados"
      Top             =   1410
      Width           =   750
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   2535
      Left            =   105
      TabIndex        =   41
      Top             =   2835
      Visible         =   0   'False
      Width           =   7890
      Begin VB.ListBox List4 
         Height          =   2010
         Left            =   3570
         TabIndex        =   45
         Top             =   315
         Width           =   5685
      End
      Begin VB.FileListBox File1 
         Height          =   2040
         Left            =   1785
         TabIndex        =   44
         Top             =   315
         Width           =   1695
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   210
         TabIndex        =   43
         Top             =   840
         Width           =   1485
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   210
         TabIndex        =   42
         Top             =   315
         Width           =   1485
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   630
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Label Label10 
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
      Height          =   270
      Left            =   1575
      TabIndex        =   52
      Top             =   6250
      Width           =   7575
   End
   Begin VB.Label Label9 
      Caption         =   "Importando de:"
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
      TabIndex        =   51
      Top             =   6300
      Width           =   1590
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   8330
      TabIndex        =   50
      Top             =   4120
      Width           =   705
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   8330
      TabIndex        =   49
      Top             =   3885
      Width           =   700
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   7665
      Picture         =   "FDIRMAIL09.frx":32E6
      Top             =   5430
      Width           =   2010
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Apel:"
      Height          =   225
      Left            =   210
      TabIndex        =   14
      Top             =   1785
      Width           =   960
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Apel:"
      Height          =   225
      Left            =   210
      TabIndex        =   13
      Top             =   1260
      Width           =   960
   End
End
Attribute VB_Name = "FDIRMAIL09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cancelflag


Private Sub Command16_Click()
   If TRIM$(FORSELEC.Text1) = "" Then
     FORSELEC.ABUSQUE = TRIM$(Label26)
   End If
   Call SELECHO("LIDIRMAI")
   VDEVU$ = VALACT1$("LIDIRMAI")
   If VDEVU$ <> "" Then
     Label26 = TRIM$(VDEVU$)
   End If
End Sub

Private Sub Command18_Click()
   '
   Command18.Enabled = False
   '
10 OPX% = COMALTER%("Seleccione de Listado:\\Todos los Contactos\Filtrado por Grupo")
   If OPX% < 1 Or OPX% > 2 Then GoTo 99
   '
   If OPX% = 1 Then
     Call CONECRUN("*LIDIRMAI", "") 'TODOS
     GoTo 10
   ElseIf OPX% = 2 Then
     Call CONTAFILTRA                'FILTRADO POR GRUPO
     GoTo 10
   End If
   '
99 Command18.Enabled = True
   '
End Sub
Sub CONTAFILTRA()
   '
   Call SELECHO("LISLISTA")
   NUM& = XVALO(VALACT1$("LISLISTA"))
   If NUM& < 1 Then Exit Sub
   TIPO$ = VALACT2$("LISLISTA")
   '
   'COPIO LA ESTRUCTURA
   Call COPYSTRU("LIDIRMAI", "LIDRMFIL")
   '
   Call SETULTREG("LIDRMFIL", 0)
   
   'RECORRO ARCHIVO LIDIRMAI
   For U& = 1 To ULTREG&("LIDIRMAI")
       XXX$ = REGLEIDO$("LIDIRMAI", U&)
       PERLIST$ = Mid$(XXX$, 495, 16)
       If InStr(PERLIST$, Chr(NUM&)) > 0 Then
         JJ& = JJ& + 1
         Call GRAREG("LIDRMFIL", XXX$, JJ&)
       End If
   Next U&
   '
   Call SETULTREG("LIDRMFIL", JJ&)
   Call CIERRARCH("LIDRMFIL")
   '
   Call HEADERS("LIDRMFIL", "")
   Call HEADERS("LIDRMFIL", "Filtro: " + TIPO$)
   Call FINAL("*LIDRMFIL")
   '

End Sub

Private Sub Command2_Click()
   '
   Static ARGUME$
   Command2.Enabled = False
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
10 ARGUME$ = TRIM$(UCase$(InputBox$("Buscar", "Búsqueda Cliente por Texto", ARGUME$)))
   '
   If ARGUME$ <> "" Then
     Screen.MousePointer = 11
     JJ& = 0
     For U& = 1 To ULTREG&("LIDIRMAI")
       XXX$ = REGLEIDO$("LIDIRMAI", U&)
       If InStr(UCase$(Mid$(XXX$, 3)), ARGUME$) > 0 Then
         JJ& = JJ& + 1
         Call GRAREG("!LIDIRMAI", XXX$, JJ&)
       End If
     Next U&
     '
     Call SETULTREG("!LIDIRMAI", JJ&)
     Call CIERRARCH("!LIDIRMAI")
     Call BAJALDISCO
     Call FRESHECHO("!LIDIRMAI")
     '
20   If JJ& > 0 Then
       Call SELECHO("!LIDIRMAI")
       NPX& = Val(VALACT1$("!LIDIRMAI"))
       If NPX& > 0 Then
         Label26 = Str$(NPX&)
       End If
     End If
     '
   End If
   Screen.MousePointer = 1
   Command2.Enabled = True
End Sub

Private Sub Command20_Click()

   Command20.Enabled = False
'Call IMPORAUTO
'Command20.Enabled = True
'Exit Sub
   HOII% = HOY(HOS$)
   List1.Clear
   For U& = 1 To ULTREG&("LIDIRMAI")
     XX$ = REGLEIDO$("LIDIRMAI", U&)
     List1.AddItem TRIM$(Mid$(XX$, 105, 60))
   Next U&
   '
   OPX% = COMALTER%("Seleccione Origen de Importación:\\Planilla Excel\DIRMAI.PRN\Otros Origenes")
   If OPX% < 1 Or OPX% > 3 Then GoTo 99
   If OPX% = 1 Then
     IMPORMAIL.Show 1
     GoTo 99
   End If
   '
   Call COMUNI("Opcion Deshabilitada por el Momento")
   GoTo 99
   '
   Stop
   N1% = FILEOPEN%("C:\FLEXOFT\DIRMAI.PRN", 1, 0)
   Line Input #N1%, A$
   While Not EOF(N1%)
     Line Input #N1%, A$
     APENO$ = ""
     DIREL$ = TRIM$(Mid$(A$, 191, 40))
     APELA$ = "Sr": If TRIM$(APENO$) = "" Then APELA$ = "Sres."
     EMPRA$ = TRIM$(Left$(A$, 50))
     ORIGE$ = "Rev.Farmac."
     GoSub ALTANUE
   Wend
   Call CIERRARCH("*.*")
   Close
   Exit Sub

1 'Importa de BSEGCOM.dat
   Dim BFLE1 As String * 1024
   ARIMPO$ = Left$(LUDAT$, 2) + "\SEGUICOM\BSEGCOM.DAT"
   N1% = FILEOPEN%(ARIMPO$, 0, 1024)
   For U& = 1 To LOF(N1%) / 1024
     Get #N1%, U&, BFLE1$
     APENO$ = TRIM$(Mid$(BFLE1$, 449, 32))
     DIREL$ = TRIM$(Mid$(BFLE1$, 65, 64))
     APELA$ = "Sr": If TRIM$(APENO$) = "" Then APELA$ = "Sres."
     EMPRA$ = TRIM$(Mid$(BFLE1$, 3, 62))
     ORIGE$ = "BSEGCOM"
     GoSub ALTANUE
   Next U&
   '
   '
   '
2 'Importa de BCONTAC.dat - Agenda Telefonica
   Dim BFLE2 As String * 1024
   ARIMPO$ = Left$(LUDAT$, 2) + "\AGETEL\BCONTACT.DAT"
   N1% = FILEOPEN%(ARIMPO$, 0, 1024)
   For U& = 1 To LOF(N1%) / 1024
     Get #N1%, U&, BFLE2$
     APENO$ = TRIM$(Mid$(BFLE2$, 89, 56))
     APELA$ = "Sr": If TRIM$(APENO$) = "" Then APELA$ = "Sres."
     EMPRA$ = TRIM$(Mid$(BFLE2$, 145, 64))
     ORIGE$ = "AGETEL"
     DIREL$ = TRIM$(Mid$(BFLE2$, 217, 64)): GoSub ALTANUE
     DIREL$ = TRIM$(Mid$(BFLE2$, 281, 64)): GoSub ALTANUE
     DIREL$ = TRIM$(Mid$(BFLE2$, 345, 64)): GoSub ALTANUE
   Next U&
   '
   '
   '
   '
   '
7  List4.Clear
   Dim BUFLE As String * 512
   ILISTA& = 0
   For UNILO% = 1 To Drive1.ListCount
     LETRA$ = TRIM$(UCase$(Drive1.List(UNILO% - 1)))
     PPXX% = InStr(LETRA$, ":")
     If PPXX% > 1 Then
       LETRA$ = Left$(LETRA$, PPXX%)
       On Error Resume Next
       Dir1.Path = LETRA$ + "\"
       If Err Then GoTo 9
       On Error GoTo 0
       List4.AddItem LETRA$ + "\"
       '
8     Dir1.Path = List4.List(ILISTA&)
       For KKI% = 1 To Dir1.ListCount
         TTXX$ = Dir1.List(KKI% - 1) + "\"
         List4.AddItem TTXX$
         List4.ListIndex = List4.ListCount - 1
       Next KKI%
       '
       If ILISTA& < List4.ListCount - 1 Then
         ILISTA& = ILISTA& + 1
         GoTo 8
       End If
     End If
9  On Error GoTo 0
   Next UNILO%
   '
   For U& = 1 To List4.ListCount
     RUTAX$ = List4.List(U& - 1)
     BACON$ = "BADACLIE"
11   If EXISTE%(RUTAX$ + BACON$ + ".DAT") Then
       N1% = FILEOPEN%(RUTAX$ + BACON + ".DAT", 0, 512)
       For REGU& = 1 To LOF(N1%) / 512
         Get #N1%, REGU&, BUFLE$
         APENO$ = TRIM$(Mid$(BUFLE$, 385, 48))
         DIREL$ = TRIM$(Mid$(BUFLE$, 305, 48))
         APELA$ = "Sr": If TRIM$(APENO$) = "" Then APELA$ = "Sres."
         EMPRA$ = TRIM$(Mid$(BUFLE$, 3, 62))
         ORIGE$ = "B.D.CLIENTES"
         GoSub ALTANUE
       Next REGU&
       Close #N1%
     End If
     If BACON$ = "BADACLIE" Then
       BACON$ = "BADAPROV"
       GoTo 11
     End If
   Next U&
   '
90 Call CIERRARCH("LIDIRMAI")
   Call CIERRARCH("*.*")
   GoTo 99
   '

ALTANUE:
   DIREL$ = LCase$(TRIM$(DIREL$))
   PPXX% = InStr(DIREL$, " ")
   RESTO$ = ""
   If PPXX% > 0 Then
      DIREL$ = TRIM$(Left$(DIREL$, PPXX% - 1))
      RESTO$ = TRIM$(Mid$(DIREL$, PPXX% + 1))
   End If
   PPXX% = InStr(DIREL$, "/")
   If PPXX% > 0 Then DIREL$ = TRIM$(Left$(DIREL$, PPXX%))
   If Left$(DIREL$, 1) = Chr$(39) Then
     DIREL$ = TRIM$(Mid$(DIREL$, 2))
   End If
   '
   If TRIM$(DIREL$) <> "" Then
     If InStr(DIREL$, "@") > 0 Then
       '
       POARRO% = InStr(DIREL$, "@")
       If InStr(POARRO% + 1, DIREL$, "@") > 0 Then
         GoTo 89
         'PARA EVITAR CARGAR DIRECCIONES MALAS
       End If
       '
       For KKU& = 1 To List1.ListCount
         If List1.List(KKU& - 1) = DIREL$ Then GoTo 89
       Next KKU&
       '
       If InStr(UCase$(APENO$), "MAILER-DAEMON") > 0 Then GoTo 89
       If InStr(UCase$(APENO$), "UNDISCLOSED-RECIPIEN") > 0 Then GoTo 89
       '
       List1.AddItem DIREL$
       RENUE$ = REGBLAN$("LIDIRMAI")
       '
       PPXX% = InStr(APENO$, "@")
       If PPXX% > 0 Then APENO$ = TRIM$(Left$(APENO$, PPXX% - 1))
       If APENO$ <> "" Then
         APENO$ = LCase$(APENO$)
         APENO$ = REPLACAR$(APENO$, "_", " ")
         '
         Call REPLA(APENO$, UCase$(Left$(APENO$, 1)), 1, 1)
         For KKI% = 2 To Len(APENO$)
           If Asc(Mid$(APENO$, KKI% - 1, 1)) < 48 Then
             Call REPLA(APENO$, UCase$(Mid$(APENO$, KKI%, 1)), KKI%, 1)
           End If
         Next KKI%
       End If
       '
       JJ& = 1 + ULTREG("LIDIRMAI")
       Call REPLA(RENUE$, MKS$(JJ&), 1, 4)
       MUESTRA$ = ""
       If TRIM$(APENO$) <> "" Then MUESTRA$ = MUESTRA$ + " " + TRIM$(APENO$)
       If TRIM$(EMPRA$) <> "" Then
         If MUESTRA$ <> "" Then
           MUESTRA$ = MUESTRA$ + " -"
         End If
         MUESTRA$ = MUESTRA$ + " " + EMPRA$
       End If
       '
       Call REPLA(RENUE$, TRIM$(MUESTRA$), 5, 60)
       Call REPLA(RENUE$, APENO$, 65, 40)
       Call REPLA(RENUE$, DIREL$, 105, 60)
       Call REPLA(RENUE$, APELA$, 165, 8)
       Call REPLA(RENUE$, EMPRA$, 173, 64)
       Call REPLA(RENUE$, ORIGE$, 237, 16)
       Call REPLA(RENUE$, MKI$(HOII%), 253, 2)
       Call REGAPP("LIDIRMAI", RENUE$)
       '
     End If
   End If
   '
89 Return
'
99 Command20.Enabled = True
End Sub

Sub IMPORAUTO()
   '
   Dim ABUSQUE$(32)
   ABUSQUE$(1) = "MAIL DELIVERY SERVICE"
   ABUSQUE$(2) = "MAIL DELIVERY SUBSYSTEM"
   ABUSQUE$(3) = "MAIL DELIVERY SYSTEM"
   ABUSQUE$(4) = "MAILER-DAEMON"
   ABUSQUE$(5) = "SYSTEM ADMINIS"
   ABUSQUE$(6) = "MAIL ADMINIS"
   ABUSQUE$(7) = "POSTMASTER"
   ABUSQUE$(8) = "THE POST OFFICE"
   ABUSQUE$(9) = "INTERNET MAIL DELIVERY"
   CABUS% = 9
   '
   HOII% = HOY(HOS$)
   List1.Clear
   For U& = 1 To ULTREG&("LIDIRMAI")
     XX$ = REGLEIDO$("LIDIRMAI", U&)
     List1.AddItem AJUSTI$(TRIM$(Mid$(XX$, 105, 60)), 60) + TRIM$(Mid$(XX$, 265, 60))
   Next U&
   '
   OPX% = COMALTER%("Seleccione Origen de Importación:\\Barrida Automática\Mails Recibidos\Planilla Excel")
   If OPX% < 1 Or OPX% > 3 Then GoTo 99
   If OPX% = 2 Then
       Height = 7000
       GoSub IMPMAIRECEP
       GoTo 99
     ElseIf OPX% = 3 Then
       IMPORMAIL.List5.Clear
       For JJ& = 0 To List5.ListCount - 1
          IMPORMAIL.List5.AddItem List5.List(JJ&)
       Next JJ&
       IMPORMAIL.Text18 = Text18
       IMPORMAIL.Show 1
       GoTo 99
   End If
   '
   ' IMPORTA DE BSEGCOM
Dim BFLE1 As String * 1024
Height = 7000
'   ARIMPO$ = "\SEGUICOM\BSEGCOM.DAT"
'   N1% = FILEOPEN%(ARIMPO$, 0, 1024)
'   For U& = 1 To LOF(N1%) / 1024
'     Label10 = U& & " - " & LOF(N1%) / 1024 & " - \SEGUICOM\BSEGCOM.DAT"
'     Label10.Refresh
'     Get #N1%, U&, BFLE1$
'     APENO$ = TRIM$(Mid$(BFLE1$, 449, 32))
'     DIREL$ = TRIM$(Mid$(BFLE1$, 65, 64))
'     APELA$ = "Sr": If TRIM$(APENO$) = "" Then APELA$ = "Sres."
'     EMPRA$ = TRIM$(Mid$(BFLE1$, 3, 62))
'        EMPRA$ = REPLACAR$(EMPRA$, "!", " ")
'     FECHBAJ% = CVI(Mid$(BFLE1$, 973, 2))
'        If FECHBAJ% < 8225 Then FECHBAJ% = 0
'     MOTIBAJ$ = TRIM$(Mid$(BFLE1$, 975, 16))
'     ORIGE$ = "BSEGCOM"
'     GoSub ALTANUE
'   Next U&
'   '
   '
   ' BARRIDA SOBRE BASES DE DATOS CLIENTES
7  List4.Clear
   Dim BUFLE As String * 512
   ILISTA& = 0
   For UNILO% = 1 To Drive1.ListCount
     LETRA$ = TRIM$(UCase$(Drive1.List(UNILO% - 1)))
     PPXX% = InStr(LETRA$, ":")
     If PPXX% > 1 Then
       LETRA$ = Left$(LETRA$, PPXX%)
       On Error Resume Next
       Dir1.Path = LETRA$ + "\"
       If Err Then GoTo 9
       On Error GoTo 0
       List4.AddItem LETRA$ + "\"
       '
8      Dir1.Path = List4.List(ILISTA&)
       Label10 = UNILO% & " - " & Drive1.ListCount & " - " & Dir1.Path
       Label10.Refresh
       For KKI% = 1 To Dir1.ListCount
         TTXX$ = Dir1.List(KKI% - 1) + "\"
         List4.AddItem TTXX$
         List4.ListIndex = List4.ListCount - 1
       Next KKI%
       '
       If ILISTA& < List4.ListCount - 1 Then
         ILISTA& = ILISTA& + 1
         GoTo 8
       End If
     End If
9  On Error GoTo 0
   Next UNILO%
   '
   For U& = 1 To List4.ListCount
     RUTAX$ = List4.List(U& - 1)
     Label10 = U& & " - " & List4.ListCount & " - " & RUTAX$
     Label10.Refresh
     ' IMPORTA DE BADACLIE Y BADAPROVE
     BACON$ = "BADACLIE"
     FECHBAJ% = 0
     MOTIBAJ$ = ""
     ORIGE$ = UCase$(RUTAX$)
     PPXX% = InStr(ORIGE$, "\DATOS")
     If PPXX% > 0 Then
       ORIGE$ = Left$(ORIGE$, PPXX% - 1)
     End If
10   PPXX% = InStr(ORIGE$, "\")
     If PPXX% > 0 Then
       ORIGE$ = Mid$(ORIGE$, PPXX% + 1)
       GoTo 10
     End If
     If TRIM$(ORIGE$) = "" Then
       ORIGE$ = "B.D.CLIENTES"
     End If
11   If EXISTE%(RUTAX$ + BACON$ + ".DAT") Then
       N1% = FreeFile
       On Error Resume Next
       Open RUTAX$ + BACON$ + ".DAT" For Random Access Read Write Shared As #N1% Len = 512
       If Err Then
           On Error GoTo 0
           GoTo 15
       End If
       On Error GoTo 0
       For REGU& = 1 To LOF(N1%) / 512
         Get #N1%, REGU&, BUFLE$
         APENO$ = TRIM$(Mid$(BUFLE$, 385, 48))
         DIREL$ = TRIM$(Mid$(BUFLE$, 305, 48))
         APELA$ = "Sr": If TRIM$(APENO$) = "" Then APELA$ = "Sres."
         EMPRA$ = TRIM$(Mid$(BUFLE$, 3, 62))
         GoSub ALTANUE
       Next REGU&
       Close #N1%
     End If
15   If BACON$ = "BADACLIE" Then
       BACON$ = "BADAPROV"
       GoTo 11
     End If
     '
     ' IMPORTA DE BADACLIE.MDB Y BADAPROVE.MDB
     ORIDAT$ = "BADACLIEN"
     If EXISTE%(RUTAX$ + "BADACLIEN.MDB") Then
        Set Badaclie = OpenDatabase(RUTAX$ + "BADACLIEN.MDB")
        Set Deudor12 = Badaclie.OpenRecordset("DEUDOR12", dbOpenDynaset)
        '
18      On Error Resume Next
        With Deudor12
           .MoveFirst
           If Err Then
              On Error GoTo 0
              GoTo 19
           End If
           On Error GoTo 0
           '
           Do While Not .EOF
             APENO$ = "": DIREL$ = "": EMPRA$ = ""
             On Error Resume Next
             APENO$ = TRIM$(!Ctac_Cli)
             DIREL$ = TRIM$(!Mail_Cli)
             APELA$ = "Sr": If TRIM$(APENO$) = "" Then APELA$ = "Sres."
             EMPRA$ = TRIM$(!Raso_Cli)
             On Error GoTo 0
             GoSub ALTANUE
           .MoveNext
           Loop
19      End With
        Set Deudor12 = Nothing
        Badaclie.Close
     End If
     '
     If ORIDAT$ = "BADACLIEN" Then
        ORIDAT$ = "BADAPROVE"
        If EXISTE%(RUTAX$ + "BADAPROVE.MDB") Then
           Set Badaclie = OpenDatabase(RUTAX$ + "BADAPROVE.MDB")
           Set Deudor12 = Badaclie.OpenRecordset("PROVED12", dbOpenDynaset)
           GoTo 18
        End If
     End If
     '
   Next U&
   '
50 ' importa de LIDIRMAIVIE - SOLO LOS DE EMPRESAS
   Call COPYSTRU("LIDIRMAI", "LIDIRVIE")
   FIN& = ULTREG&("LIDIRVIE")
   For KUK& = 1 To FIN&
      Call TRACE(20, KUK&, FIN&)
      REVIE$ = REGLEIDO$("LIDIRVIE", KUK&)
      APENO$ = TRIM$(Mid$(REVIE$, 65, 40))
      DIREL$ = TRIM$(Mid$(REVIE$, 105, 60))
      APELA$ = TRIM$(Mid$(REVIE$, 165, 8))
      EMPRA$ = TRIM$(Mid$(REVIE$, 173, 64))
      ORIGE$ = TRIM$(Mid$(REVIE$, 237, 16))
      HOII% = CVI(Mid$(REVIE$, 253, 2))
      FECHBAJ% = CVI(Mid$(REVIE$, 255, 2))
      MOTIBAJ$ = TRIM$(Mid$(REVIE$, 257, 8))
      If FECHBAJ% < 8225 Then
        If MOTIBAJ$ = "" Then
          GoSub ALTANUE
        End If
      End If
   Next KUK&
   '
   Height = 6450
   Call COMUNI("Fin de la Importación.\Importados " & List1.ListCount & " Registros")
   Exit Sub
   '
ALTANUE:
   OPENFORMS = DoEvents
   DIREL$ = LCase$(TRIM$(DIREL$))
   PPXX% = InStr(DIREL$, " ")
   RESTO$ = ""
   If PPXX% > 0 Then
      DIREL$ = TRIM$(Left$(DIREL$, PPXX% - 1))
      RESTO$ = TRIM$(Mid$(DIREL$, PPXX% + 1))
   End If
   PPXX% = InStr(DIREL$, "/")
   If PPXX% > 0 Then DIREL$ = TRIM$(Left$(DIREL$, PPXX%))
   If Left$(DIREL$, 1) = Chr$(39) Then
     DIREL$ = TRIM$(Mid$(DIREL$, 2))
   End If
   '
   If TRIM$(DIREL$) <> "" Then
     If InStr(DIREL$, "@") > 0 Then
       '
       POARRO% = InStr(DIREL$, "@")
       If InStr(POARRO% + 1, DIREL$, "@") > 0 Then
         GoTo 89
         'PARA EVITAR CARGAR DIRECCIONES MALAS
       End If
       '
       DIREL$ = LCase$(TRIM$(DIREL$))
       For KKU& = 1 To List1.ListCount
         If TRIM$(Left$(List1.List(KKU& - 1), 60)) = DIREL$ Then GoTo 89
         If TRIM$(Mid$(List1.List(KKU& - 1), 61, 60)) = DIREL$ Then GoTo 89
       Next KKU&
       '
       If InStr(UCase$(APENO$), "MAILER-DAEMON") > 0 Then GoTo 89
       If InStr(UCase$(APENO$), "UNDISCLOSED-RECIPIEN") > 0 Then GoTo 89
       '
       RENUE$ = REGBLAN$("LIDIRMAI")
       '
       PPXX% = InStr(APENO$, "@")
       If PPXX% > 0 Then APENO$ = TRIM$(Left$(APENO$, PPXX% - 1))
       If APENO$ <> "" Then
         APENO$ = LCase$(APENO$)
         APENO$ = REPLACAR$(APENO$, "_", " ")
         '
         Call REPLA(APENO$, UCase$(Left$(APENO$, 1)), 1, 1)
         For KKI% = 2 To Len(APENO$)
           If Asc(Mid$(APENO$, KKI% - 1, 1)) < 48 Then
             Call REPLA(APENO$, UCase$(Mid$(APENO$, KKI%, 1)), KKI%, 1)
           End If
         Next KKI%
       End If
       '
       JJ& = 1 + ULTREG("LIDIRMAI")
       Call REPLA(RENUE$, MKS$(JJ&), 1, 4)
       MUESTRA$ = ""
       If TRIM$(APENO$) <> "" Then MUESTRA$ = MUESTRA$ + " " + TRIM$(APENO$)
       If TRIM$(EMPRA$) <> "" Then
         If MUESTRA$ <> "" Then
           MUESTRA$ = MUESTRA$ + " -"
         End If
         MUESTRA$ = MUESTRA$ + " " + EMPRA$
       End If
       '
       Call REPLA(RENUE$, TRIM$(MUESTRA$), 5, 60)
       Call REPLA(RENUE$, APENO$, 65, 40)
       Call REPLA(RENUE$, DIREL$, 105, 60)
       Call REPLA(RENUE$, APELA$, 165, 8)
       Call REPLA(RENUE$, EMPRA$, 173, 64)
       Call REPLA(RENUE$, ORIGE$, 237, 16)
       Call REPLA(RENUE$, MKI$(HOII%), 253, 2)
       Call REPLA(RENUE$, MKI$(FECHBAJ%), 255, 2)
       Call REPLA(RENUE$, MOTIBAJ$, 257, 8)
       Call REPLA(RENUE$, DIREL$, 265, 60)
       Call REPLA(RENUE$, MKI$(0), 511, 2)
       Call REGAPP("LIDIRMAI", RENUE$)
       '
       List1.AddItem DIREL$
       List1.Refresh
       '
       Text5 = DIREL$
       Text21 = APELA$
       Text22 = APENO$
       Text4 = EMPRA$
       Text18 = ORIGE$
       OPENFORMS = DoEvents
       '
     End If
   End If
   '
89 Return
   '
IMPMAIRECEP:
  '
4 'Importa de Libreta de Mails Recibidos
   MAPISession1.DownLoadMail = False
   MAPISession1.LogonUI = False
   MAPISession1.Action = 1
   MAPIMessages1.SessionID = MAPISession1.SessionID
   '
   MAPIMessages1.Fetch
   CAMENSA& = MAPIMessages1.MsgCount
   '
   For U& = 1 To CAMENSA&
     MAPIMessages1.MsgIndex = U& - 1
     APENO$ = TRIM$(MAPIMessages1.MsgOrigDisplayName)
        Label10 = U& & " - " & CAMENSA& & " - " & APENO$
        Label10.Refresh
     DIREL$ = TRIM$(MAPIMessages1.MsgOrigAddress)
     APELA$ = "Sr": If TRIM$(APENO$) = "" Then APELA$ = "Sres."
     EMPRA$ = ""
     ORIGE$ = "Mails Recibidos"
     '
     '
     ' PARA DESCARTAR LOS NO DESEADOS
     ORENVI$ = TRIM$(MAPIMessages1.MsgOrigDisplayName)
     ASUNTO$ = TRIM$(LCase$(MAPIMessages1.MsgSubject))
     ' borra confirmación de lectura
     ATENTI% = 0
     If Left$(ASUNTO$, 6) = "leído:" Or Left$(ASUNTO$, 9) = "no leído:" Or Left$(ASUNTO$, 5) = "read:" Then
          ATENTI% = 1: GoTo 60
        ElseIf InStr(UCase$(ASUNTO$), "SPAM") > 0 Then
          ATENTI% = 1: GoTo 60
        ElseIf InStr(UCase$(ASUNTO$), "PHARMACY") > 0 Then
          ATENTI% = 1: GoTo 60
        ElseIf InStr(UCase$(ASUNTO$), "GENERIC") > 0 Then
          ATENTI% = 1: GoTo 60
        ElseIf InStr(UCase$(ASUNTO$), "DRUGS") > 0 Then
          ATENTI% = 1: GoTo 60
        ElseIf InStr(UCase$(ASUNTO$), "UNDELIVERABLE MAIL") > 0 Then
          ATENTI% = 1: GoTo 60
     End If
     '
     '
     GoSub ALTANUE
     '
     CARECIP& = MAPIMessages1.RecipCount
     For UU& = 1 To CARECIP&
       MAPIMessages1.RecipIndex = UU& - 1
       APENO$ = TRIM$(MAPIMessages1.RecipDisplayName)
       DIREL$ = TRIM$(MAPIMessages1.RecipAddress)
       APELA$ = "Sres": If TRIM$(APENO$) = "" Then APELA$ = "Sres."
       EMPRA$ = ""
       ORIGE$ = "Mails Recibidos"
       GoSub ALTANUE
     Next UU&
60 Next U&
   '
Return
   '
99 Height = 6450
   End Sub
'
Private Sub Command24_Click()
   Command24.Enabled = False
   MOTI$ = InputBox$("Motivo de Baja")
   If MOTI$ <> "" Then
     HOII% = HOY(HOS$)
     Label19 = HOS$
     Text23 = MOTI$
     Call GRAFICON
   End If
   Command24.Enabled = True
End Sub

Private Sub Command26_Click()
   UKU& = REGNUM&("LIDIRMAI")
   If UKU& > 1 Then
     XX$ = REGLEIDO$("LIDIRMAI", UKU& - 1)
     Label26 = TRIM$(Str$(CVS(Left$(XX$, 4))))
   End If
End Sub

Private Sub Command27_Click()
   UKU& = REGNUM&("LIDIRMAI")
   If UKU& < ULTREG&("LIDIRMAI") Then
     XX$ = REGLEIDO$("LIDIRMAI", UKU& + 1)
     Label26 = TRIM$(Str$(CVS(Left$(XX$, 4))))
   End If
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   SELDESTI09.List5.Clear
   For JJ& = 0 To List5.ListCount - 1
     SELDESTI09.List5.AddItem List5.List(JJ&)
   Next JJ&
   SELDESTI09.Check1.Value = 0
   SELDESTI09.Label1 = ""
   SELDESTI09.Show 1
   '
   If SELDESTI09.Label1 <> "OK" Then GoTo 99
   '
   Top = Screen.Height
   ARMANEW09.Show 1
   Call CENTRAFORM(Me)
   '
99 Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Visible = False
   OPENFORMS = DoEvents
   '
   Dim DIRMA$()
   'Command5.Enabled = False
   '
   Dim ABUSQUE$(32)
   ABUSQUE$(1) = "MAIL DELIVERY SERVICE"
   ABUSQUE$(2) = "MAIL DELIVERY SUBSYSTEM"
   ABUSQUE$(3) = "MAIL DELIVERY SYSTEM"
   ABUSQUE$(4) = "MAILER-DAEMON"
   ABUSQUE$(5) = "SYSTEM ADMINIS"
   ABUSQUE$(6) = "MAIL ADMINIS"
   ABUSQUE$(7) = "POSTMASTER"
   ABUSQUE$(8) = "THE POST OFFICE"
   ABUSQUE$(9) = "INTERNET MAIL DELIVERY"
   CABUS% = 9
   'ABUS$ = TRIM$(UCase$(InputBox$("Buscar Remitentes", , "MAIL")))
   'If ABUS$ = "" Then GoTo 99
   '
   Screen.MousePointer = 11
   '
   CANDIR& = ULTREG&("LIDIRMAI")
   ReDim DIRMA$(CANDIR&)
   For UKK& = 1 To CANDIR&
       XXX$ = REGLEIDO$("LIDIRMAI", UKK&)
       XX1$ = LCase$(TRIM$(Mid$(XXX$, 105, 60)))
       DIRMA$(UKK&) = XX1$
   Next UKK&
   '
   MAPISession1.DownLoadMail = False
   MAPISession1.LogonUI = False
   MAPISession1.Action = 1
   MAPIMessages1.SessionID = MAPISession1.SessionID
   '
   MAPIMessages1.Fetch
   CAMENSA& = MAPIMessages1.MsgCount
   Screen.MousePointer = 1
   '
   Label6 = CAMENSA&
   For U& = 1 To CAMENSA&
10   If U& > MAPIMessages1.MsgCount Then GoTo 99
     Label1 = U&
     Label6 = MAPIMessages1.MsgCount
     OPENFORMS = DoEvents
     MAPIMessages1.MsgIndex = U& - 1
     ORENVI$ = TRIM$(MAPIMessages1.MsgOrigDisplayName)
     ASUNTO$ = TRIM$(LCase$(MAPIMessages1.MsgSubject))
     ' borra confirmación de lectura
     ATENTI% = 0
     If Left$(ASUNTO$, 6) = "leído:" Or Left$(ASUNTO$, 9) = "no leído:" Or Left$(ASUNTO$, 5) = "read:" Then
          ATENTI% = 1: GoTo 60
        ElseIf InStr(UCase$(ASUNTO$), "SPAM") > 0 Then
          If InStr(UCase$(ORENVI$), "CESSI") < 1 Then
            ATENTI% = 1: GoTo 60
          End If
        ElseIf InStr(UCase$(ASUNTO$), "PHARMACY") > 0 Then
          ATENTI% = 1: GoTo 60
        ElseIf InStr(UCase$(ASUNTO$), "GENERIC") > 0 Then
          ATENTI% = 1: GoTo 60
        ElseIf InStr(UCase$(ASUNTO$), "DRUGS") > 0 Then
          ATENTI% = 1: GoTo 60
        ElseIf InStr(UCase$(ASUNTO$), "UNDELIVERABLE MAIL") > 0 Then
          MODOBO% = 0
          FORLIMPIA.Text1 = ORENVI$
          FORLIMPIA.Text2 = Left$(MAPIMessages1.MsgNoteText, 1024)
          GoTo 31
     End If
     '
     MODOBO% = 0
     '
     For KKI% = 1 To CABUS%
       If Left$(UCase$(ORENVI$), Len(ABUSQUE$(KKI%))) = ABUSQUE$(KKI%) Then
         ATENTI% = 0
         MODOBO% = 0
         FORLIMPIA.Text1 = ORENVI$
         FORLIMPIA.Text2 = Left$(MAPIMessages1.MsgNoteText, 1024)
         PINI% = 1
         '
12       PPYY% = InStr(FORLIMPIA.Text2, "rfc822;")
         If PPYY% > 0 Then
           PPYY% = PPYY% + 7
           For PPXX% = PPYY% To PPYY% + 128
             If Asc(Mid$(FORLIMPIA.Text2, PPXX%, 1)) < 45 Then
               FORLIMPIA.Text2.SelStart = PPYY%
               FORLIMPIA.Text2.SelLength = PPXX% - PPYY%
               TEBUX$ = Mid$(FORLIMPIA.Text2, PPYY%, PPXX% - PPYY%)
               FORLIMPIA.Text3 = TEBUX$
               MODOBO% = 1
               GoTo 50
             End If
           Next PPXX%
         End If
         '
         PINI% = 1
20       PPXX% = InStr(PINI%, FORLIMPIA.Text2, "<")
         If PPXX% > 0 Then
           FORLIMPIA.Text2.SelStart = PPXX%
           PPYY% = InStr(PPXX%, FORLIMPIA.Text2, ">")
           If PPYY% <= PPXX% Then GoTo 30
           If PPYY% > PPXX% Then
             FORLIMPIA.Text2.SelLength = PPYY% - PPXX% - 1
           End If
           TEBUX$ = Mid$(FORLIMPIA.Text2, PPXX% + 1, FORLIMPIA.Text2.SelLength)
           If InStr(UCase$(TEBUX$), "FLEXOFT") > 0 Then
             PINI% = PPYY%
             GoTo 20
           End If
           If InStr(TEBUX$, "@") < 1 Then
             PINI% = PPYY%
             GoTo 20
           End If
           FORLIMPIA.Text3 = TEBUX$
           MODOBO% = 1
           GoTo 50
         End If
         '
30       PPYY% = InStr(FORLIMPIA.Text2, "; Failed")
         If PPYY% > 0 Then
           PPYY% = PPYY% - 1
           For PPXX% = PPYY% To 1 Step -1
             If Asc(Mid$(FORLIMPIA.Text2, PPXX%, 1)) < 45 Then
               FORLIMPIA.Text2.SelStart = PPXX%
               FORLIMPIA.Text2.SelLength = PPYY% - PPXX%
               TEBUX$ = Mid$(FORLIMPIA.Text2, PPXX% + 1, PPYY% - PPXX%)
               FORLIMPIA.Text3 = TEBUX$
               MODOBO% = 1
               GoTo 50
             End If
           Next PPXX%
         End If
         '
31       PINI% = 1
32       PPYY% = InStr(PINI%, FORLIMPIA.Text2, "@")
         If PPYY% > 0 Then
           For PPXX% = PPYY% - 1 To 1 Step -1
             If Asc(Mid$(FORLIMPIA.Text2, PPXX%, 1)) < 45 Then
               INI% = PPXX% + 1
               GoTo 33
             End If
           Next PPXX%
           INI% = 1
33         For PPXX% = PPYY% + 1 To Len(FORLIMPIA.Text2)
             If Asc(Mid$(FORLIMPIA.Text2, PPXX%, 1)) < 45 Or Asc(Mid$(FORLIMPIA.Text2, PPXX%, 1)) = 58 Then
               FIN% = PPXX% - 1
               GoTo 34
             End If
           Next PPXX%
           FIN% = Len(FORLIMPIA.Text2)
34         TEBUX$ = Mid$(FORLIMPIA.Text2, INI%, 1 + FIN% - INI%)
           If InStr(UCase$(TEBUX$), "FLEXOFT") > 0 Then
              PINI% = FIN% + 1
              GoTo 32
           End If
           FORLIMPIA.Text3 = TEBUX$
           MODOBO% = 1
           GoTo 50
         End If
         '
50       If MODOBO% = 1 Then
           HOII% = HOY(HOS$)
           TBXX$ = LCase$(TRIM$(TEBUX$))
           If TBXX$ <> "" Then
             JJ& = 0
             For UKK& = 1 To CANDIR&  ' ULTREG&("LIDIRMAI")
               If DIRMA$(UKK&) = TBXX$ Then
                  XXX$ = REGLEIDO$("LIDIRMAI", UKK&)
                  'XX1$ = TRIM$(Mid$(XXX$, 105, 60))
                  'If XX1$ = TRIM$(TEBUX$) Then
                  Call REPLA(XXX$, MKI$(HOII%), 255, 2)
                  Call REPLA(XXX$, "MAL MAIL", 257, 8)
                  Call GRAREG("LIDIRMAI", XXX$, UKK&)
               End If
             Next UKK&
             ATENTI% = 1
           End If
           GoTo 60
         End If
         '
         FORLIMPIA.Text3 = ""
         FORLIMPIA.Show 1
         '
60       If ATENTI% = 1 Then
            MAPIMessages1.Delete
            If U& <= MAPIMessages1.MsgCount Then
              GoTo 10
            End If
          ElseIf ATENTI% = 2 Then
            GoTo 99
         End If
         GoTo 89
         '
       End If
     Next KKI%
     '
89 Next U&
   '
99 MAPISession1.Action = 2  ' !!!
   Command5.Visible = True
End Sub

Private Sub Command6_Click()
   Label19 = ""
   Text23 = ""
End Sub


'Private Sub Command7_Click()
'   Dim FLEXCONN As ADODB.Connection
'
'   LUDA$ = "X:\DOSIVAC_SQL\FLEXSQL\DATOS\"
'   XI% = EXISTE%(LUDA$ + "CONNSQLS.DRV")
'   If XI% = 1 Then
''   On Error Resume Next
'   ESTADODECONEXION% = FLEXCONN.State
'   'If ESTADODECONEXION% = 1 Then
'   '   Exit Sub
'   'End If
'   '
'   CONNSQL$ = TRIM$(LOADFILE$(LUDA$ + "CONNSQLS.DRV"))
'   If CONNSQL$ <> "" Then
'      ' VERIFICA LA CONEXION
'      Set FLEXCONN = New ADODB.Connection
'      FLEXCONN.ConnectionString = CONNSQL$   ' "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Data Source=FLEXSERVER;Initial Catalog=" + conecta$ + "; Connect Timeout=999"
'      On Error Resume Next
'      FLEXCONN.Open
'      If Err Then
'         MsgBox "Imposible Abrir Conexión a Servidor SQL con " + CONNSQL$, vbCritical
'         Close: End
'      End If
'      On Error GoTo 0
'      '
'    End If
'   End If
'End Sub
'
Private Sub Form_Load()
   Call AUTOEXEC
   ABC$ = UCase$(LUDAT$)
   LUDAT$ = "F:\LIDIRMAI\"
   If InStr(ABC$, "ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB") > 0 Then ' \ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB") > 0 Then
     LUDAT$ = "F:\LIDIRMAI\"
     MsgBox LUDAT$
   End If
   '
   
   Label35 = NOMTER
   HOII% = HOY(HOS$)
   Label34 = HOS$
   Text12 = FECHATEX$(DIPOST(HOII%))
   Call cargalist5
End Sub

Sub cargalist5()
     FIN& = ULTREG&("LISLISTA")
     For U& = 1 To ULTREG&("LISLISTA")
       XX$ = REGLEIDO$("LISLISTA", U&)
       DESCRIPCI$ = TRIM$(Mid$(XX$, 3, 64))
       List5.AddItem DESCRIPCI$
     Next U&
End Sub
'
Private Sub Command1_Click()
   '
   If Val(Label26) < 1 Then
     Command16.SetFocus
     Exit Sub
   End If
   '
   Command1.Enabled = False
   If Val(Label26) > 0 Then
     Call GRAFICON
   End If
   Command1.Enabled = True
   '
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   OPX% = ALTERNA%("Configuración General\Lista de Listas\Lista de E-Mail")
   If OPX% < 1 Or OPX% > 3 Then GoTo 99
   If OPX% = 1 Then
    Call CONECRUN("FLSETUP", "Configuración de Funcionamiento")
    Call FINAL("")
   ElseIf OPX% = 2 Then
    Call CARLSTLST ' CARGA DE LISTAS
    Call FINAL("")
   ElseIf OPX% = 3 Then
    Call CARLSTMAIL ' CARGA DE E-Mail-Usuario-Pass-SMTP
   End If
99 Command14.Enabled = True
End Sub
Sub CARLSTLST()

   Call TRALOCAL("LISLISTA", "")
10 VTB% = VENTABU%("LISLISTA") ' recorre la ventana validando que tengan
   If VTB% >= 0 Then           ' descripcion todos los registos grabados
     FIN& = ULTREG&("!LISLISTA")
     For U& = 1 To ULTREG&("!LISLISTA")
       XX$ = REGLEIDO$("!LISLISTA", U&)
       DESCRIPCI$ = TRIM$(Mid$(XX$, 3, 61))
       If DESCRIPCI$ = "" And U& <= FIN& Then
         Call COMUNI("Falta Ingresar Descripción ")
         GoTo 10
       End If
       UKI% = U&
       Call REPLA(XX$, MKI$(UKI%), 1, 2) ' numera los registros
       Call GRAREG("!LISLISTA", XX$, U&) ' grabo el numero de registro
     Next U&
     Call TRACENTRAL("LISLISTA", "")
   End If
End Sub

Sub CARLSTMAIL()
  Call TRALOCAL("LISMAIL", "")
10 VTB% = VENTABU%("LISMAIL")
   If VTB% >= 0 Then
     FIN& = ULTREG&("!LISMAIL")
     For U& = 1 To FIN&
       XX$ = REGLEIDO$("!LISMAIL", U&)
       MAIL$ = TRIM$(Mid$(XX$, 1, 64))
       USSUU$ = TRIM$(Mid$(XX$, 65, 32))
       PASSW$ = TRIM$(Mid$(XX$, 97, 16))
       SMTPP$ = TRIM$(Mid$(XX$, 113, 32))
       If MAIL$ = "" Or USSUU$ = "" Or SMTPP = "" And U& <= FIN& Then
         Call COMUNI("Datos Incompletos. Verificar")
         GoTo 10
       End If
       '
       UKI% = U&
       Call REPLA(XX$, MAIL$, 1, 64)
       Call REPLA(XX$, USSUU$, 65, 32)
       Call REPLA(XX$, PASSW$, 97, 16)
       Call REPLA(XX$, SMTPP$, 113, 32)
       Call GRAREG("!LISMAIL", XX$, U&)
     Next U&
     Call TRACENTRAL("LISMAIL", "")
   End If
  End Sub

Private Sub Command21_Click()
   Label26 = 1 + ULTREG&("LIDIRMAI")
   HOII% = HOY(HOS$)
   Label2 = HOS$
   Label19 = ""
   On Error Resume Next
   Text5.SetFocus
   On Error GoTo 0
End Sub

Private Sub Command25_Click()
   ANCHAC = Width
   VISIAC = Picture1.Visible
   Image2.Visible = False
   Picture1.Visible = False
   Width = 10800
   On Error Resume Next
   PRINTFORM
   If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Error de Impresión de Ficha")
   End If
   On Error GoTo 0
   Image2.Visible = True
   Picture1.Visible = VISIAC
   Width = ANCHAC
End Sub

Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command9_Click()
   '
   Command9.Enabled = False
   Call FILESORT("LIDIRMAI", "", 2, 2, "ASC")
   Command9.Enabled = True
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Label19_DblClick()
   '
   If Val(Label26) > 0 Then
     If TRIM$(Label19) = "" Then
       HOII% = HOY(HOS$)
       Label19 = HOS$
       Text23 = "B.Manual"
       Call GRAFICON
     End If
   End If
   '
End Sub

Private Sub Label26_Change()
   '
   Call CARFICHA
   '
End Sub

Private Sub List2_Click()
   List2.ToolTipText = TRIM$(Mid$(List2.TEXT, 25, 48))
End Sub

Private Sub List2_DblClick()
   Call ACTAGRUP
End Sub

Private Sub Text1_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text18_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text21_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text22_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text23_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Sub GRAFICON()
   '
5  IFICHA& = Val(Label26)
   REG& = REGBUS&("LIDIRMAI", 1, MKS$(IFICHA&))
   If REG& < 1 Then REG& = 1 + ULTREG&("LIDIRMAI")
   REGI% = REG&
   '
   If REG& <= ULTREG("LIDIRMAI") Then
        RELE$ = REGLEIDO$("LIDIRMAI", REG&)
      Else
        RELE$ = MKS$(REG&) + Space$(510)
   End If
   '
   APENO$ = TRIM$(Text22)
   EMPRA$ = TRIM$(Text4)
   DIREL$ = TRIM$(Text5)
   APELA$ = TRIM$(Text21)
   ORIGE$ = TRIM$(Text18)
   OBSER$ = Text1
   MUESTRA$ = ""
   If TRIM$(APENO$) <> "" Then MUESTRA$ = MUESTRA$ + " " + TRIM$(APENO$)
   If TRIM$(EMPRA$) <> "" Then
     If MUESTRA$ <> "" Then
       MUESTRA$ = MUESTRA$ + " -"
     End If
     MUESTRA$ = MUESTRA$ + " " + EMPRA$
   End If
   '
   Call REPLA(RELE$, MUESTRA$, 5, 60)
   Call REPLA(RELE$, APENO$, 65, 40)
   Call REPLA(RELE$, DIREL$, 105, 60)
   Call REPLA(RELE$, APELA$, 165, 8)
   Call REPLA(RELE$, EMPRA$, 173, 64)
   Call REPLA(RELE$, ORIGE$, 237, 16)
   FEXI% = FECHANUM(Label2)
   If TRIM$(Label2) = "" Then FEXI% = HOY(HOS$)
   Call REPLA(RELE$, MKI$(FEXI%), 253, 2)
   FEXI% = 0
   If TRIM$(Label19) <> "" Then FEXI% = FECHANUM(Label19)
   Call REPLA(RELE$, MKI$(FEXI%), 255, 2)
   MBAJ$ = ""
   If FEXI% > 0 Then MBAJ$ = Text23
   Call REPLA(RELE$, MBAJ$, 257, 8)
   Call REPLA(RELE$, OBSER$, 265, 230)
      PERLIST$ = ""
      For KKI% = 1 To List5.ListCount
         If List5.Selected(KKI% - 1) = True Then PERLIST$ = PERLIST$ + Chr$(KKI%)
      Next KKI%
      PERLIST$ = PERLIST$ + String$(16, 0)
   Call REPLA(RELE$, PERLIST$, 495, 16)
   '
   Call GRAREG("LIDIRMAI", RELE$, REG&)
   '
   Call CIERRARCH("LIDIRMAI")
   Call FRESHECHO("LIDIRMAI")
   '
End Sub

Sub CARFICHA()
   '
5  IFICHA& = Val(Label26)
   REG& = REGBUS&("LIDIRMAI", 1, MKS$(IFICHA&))
   If REG& < 1 Or REG& > ULTREG&("LIDIRMAI") Then
       Text1 = ""
       Text2 = ""
       Text3 = ""
       Text4 = ""
       Text5 = ""
       Text9 = ""
       Text16 = ""
       Text15 = ""
       Text13 = ""
       Text12 = ""
       Text25 = ""
       Text14 = ""
       Text11 = ""
       Text20 = ""
       Text21 = ""
       Text22 = ""
       Text19 = ""
       Text23 = ""
       Text26 = ""
       Text17 = ""
       Text24 = ""
       Text18 = ""
       '
       For KKI% = 1 To List5.ListCount
         List5.Selected(KKI% - 1) = False
       Next KKI%
       '
       HOII% = HOY(HOS$)
     Else
       RELE$ = REGLEIDO$("LIDIRMAI", REG&)
       Text4 = TRIM$(Mid$(RELE$, 173, 64))
       If InStr(Text4, "!") > 0 Then
         Text4 = REPLACAR$(Text4, "!", " ")
         Call REPLA(RELE$, Text4, 173, 64)
         Call GRAREG("LIDIRMAI", RELE$, REG&)
       End If
       '
       Text5 = TRIM$(Mid$(RELE$, 105, 60))
       Text21 = TRIM$(Mid$(RELE$, 165, 8))
       Text22 = TRIM$(Mid$(RELE$, 65, 40))
       Text18 = TRIM$(Mid$(RELE$, 237, 16))
       '
       FEX = CVI(Mid$(RELE$, 253, 2))
       Label2 = FECHATEX$(FEX)
       FEX = CVI(Mid$(RELE$, 255, 2))
       Label19 = FECHATEX$(FEX)
       Text23 = TRIM$(Mid$(RELE$, 257, 8))
       Text1 = RTrim$(Mid$(RELE$, 265, 230))
       '
       PERLIST$ = Mid$(RELE$, 495, 16)
       For KKI% = 1 To List5.ListCount
         List5.Selected(KKI% - 1) = False
         If InStr(PERLIST$, Chr$(KKI%)) > 0 Then
           List5.Selected(KKI% - 1) = True
         End If
       Next KKI%
       '
   End If
   List5.ListIndex = (-1)
   '
End Sub

Sub EXPORAGE()
   '
   On Error GoTo 90
   N1% = FreeFile
   Open "A:EXPORLIB.CSV" For Output As #N1%
   Print #N1%, "Nombre,Dirección electrónica"
   On Error GoTo 0
   '
   Screen.MousePointer = 11
   For U& = 1 To ULTREG&("BCONTACT")
     RELE$ = REGLEIDO$("BCONTACT", U&)
     TX2$ = TRIM$(Mid$(RELE$, 89, 56))
     TX4$ = TRIM$(Mid$(RELE$, 145, 64))
     TTXX$ = TRIM$(TX2$)
     If TRIM$(TX2$) <> "" And TRIM$(TX4$) <> "" Then
         TTXX$ = TTXX$ + " - "
       Else
         TTXX$ = TTXX$ + " "
     End If
     '
     TTYY$ = TRIM$(TTXX$ + TRIM$(TX4$))
     TTZZ$ = REPLACAR$(TTYY$, Chr$(34), "")
     TTXX$ = REPLACAR$(TTZZ$, ",", " ")
     '
     Call REPLA(RELE$, TTXX$, 3, 78)
     Call GRAREG("BCONTACT", RELE$, U&)
     '
     EMA$ = TRIM$(Mid$(RELE$, 217, 64))
     If EMA$ <> "" Then
       If Asc(Left$(EMA$, 1)) <> 39 Then
         AX$ = TTXX$ + "," + EMA$
         Print #N1%, AX$
         JJ& = JJ& + 1
       End If
     End If
   Next U&
   '
   Close #N1%
   Call CIERRARCH("*.*")
   MsgBox Str$(JJ&) + " Direcciones Exportadas"
   Screen.MousePointer = 1
   Exit Sub
   '
90 Resume 91
91 On Error GoTo 0
   '
End Sub

Sub IMPORAGE()
   '
   Screen.MousePointer = 11
   Dim RAEMP$(4096)
   For U& = 1 To ULTREG&("BCONTACT")
     X$ = REGLEIDO$("BCONTACT", U&)
     RAEMP$(U&) = UCase$(TRIM$(Mid$(X$, 145, 64)))
     CAEMP& = U&
   Next U&
   '
   FIN& = ULTREG&("!BCONTACT")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     X$ = REGLEIDO$("!BCONTACT", U&)
     RASOE$ = UCase$(TRIM$(Mid$(X$, 145, 64)))
     EMAI$ = LCase$(TRIM$(Mid$(X$, 217, 64)))
     For JJ& = 1 To CAEMP&
       If RAEMP$(JJ&) = RASOE$ Then
         Y$ = REGLEIDO$("BCONTACT", JJ&)
         If InStr(Y$, EMAI$) > 0 Then GoTo 19 ' YA ESTA
         If TRIM$(Mid$(Y$, 217, 64)) = "" Then
             Call REPLA(Y$, EMAI$, 217, 64)
           ElseIf TRIM$(Mid$(Y$, 281, 64)) = "" Then
             Call REPLA(Y$, EMAI$, 281, 64)
           ElseIf TRIM$(Mid$(Y$, 345, 64)) = "" Then
             Call REPLA(Y$, EMAI$, 345, 64)
         End If
         Call GRAREG("BCONTACT", Y$, JJ&)
         GoTo 19
       End If
    Next JJ&
    '
    CAEMP& = CAEMP& + 1
    RAEMP$(CAEMP&) = UCase$(TRIM$(Mid$(X$, 145, 64)))
    KKI% = CAEMP&
    Call REPLA(X$, MKI$(KKI%), 1, 2)
    Call REGAPP("BCONTACT", X$)
19 Next U&
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
End Sub

Sub IMPOROUT()
   '
   On Error GoTo 90
   N1% = FreeFile
   Open "A:\OUTLOOK.CSV" For Input As #N1%
   Line Input #N1%, A$
   On Error GoTo 0
   '
   REG& = 0
   While Not EOF(N1%)
     REG& = REG& + 1
     REGI% = REG&
     RELE$ = MKI$(REGI%) + Space$(1022)
     '
     Line Input #N1%, TTXX$
     PPXX% = InStr(TTXX$, ",")
     If PPXX% > 0 Then
        EMPRE$ = Left$(TTXX$, PPXX% - 1)
        TTXX$ = Mid$(TTXX$, PPXX% + 1)
        PPXX% = InStr(TTXX$, ",")
        If PPXX% > 0 Then
          EMAI$ = Left$(TTXX$, PPXX% - 1)
          '
          Call REPLA(RELE$, EMPRE$, 3, 78)
          Call REPLA(RELE$, "Sr.", 81, 8)
          Call REPLA(RELE$, EMPRE$, 145, 64)
          Call REPLA(RELE$, EMAI$, 217, 64)
          Call GRAREG("BCONTACT", RELE$, REG&)
        End If
     End If
     '
   Wend
   '
   Call CIERRARCH("BCONTACT")
   Call FRESHECHO("BCONTACT")
   Exit Sub
   '
90 Resume 91
91 On Error GoTo 0
   '
End Sub

Private Sub Text4_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Private Sub Text5_DblClick()
   '
   NPX& = Val(Label26)
   If NPX& > 0 Then
     Call CREAMAIL(NPX&)
     '
     ATENTI% = 1
     MAILMESS.Label26 = Label26
     MAILMESS.Show 1
   End If
   '
End Sub

Private Sub Text5_GotFocus()
   If Val(Label26) < 1 Then Command16.SetFocus
End Sub

Sub ACTAGRUP()
   Call TRALOCAL("TAGRUCON", "")
   VTB% = VENTABU%("TAGRUCON")
   If VTB% >= 0 Then
     For U& = 1 To ULTREG&("!TAGRUCON")
       XX$ = REGLEIDO$("!TAGRUCON", U&)
       UKI% = U&
       Call REPLA(XX$, MKI$(UKI%), 1, 2)
       Call GRAREG("!TAGRUCON", XX$, U&)
     Next U&
     Call TRACENTRAL("TAGRUCON", "")
     If ULTREG&("TAGRUCON") > 255 Then
       Call SETULTREG("TAGRUCON", 255)
     End If
   End If
End Sub


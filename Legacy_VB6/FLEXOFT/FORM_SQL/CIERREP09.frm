VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CIERREP09 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11550
   Icon            =   "CIERREP09.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   11550
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame15 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Lote de Repuestos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   80
      TabIndex        =   26
      Top             =   7470
      Width           =   10560
      Begin VB.CommandButton Command4 
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
         Height          =   840
         Left            =   9720
         Picture         =   "CIERREP09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   300
         Width           =   660
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
         Height          =   285
         Left            =   5520
         TabIndex        =   27
         Top             =   840
         Width           =   175
      End
      Begin VB.Label Label12 
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
         Left            =   3360
         TabIndex        =   35
         Top             =   420
         Width           =   5145
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   450
         TabIndex        =   34
         Top             =   855
         Width           =   810
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   33
         Top             =   855
         Width           =   810
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
         Left            =   1320
         TabIndex        =   32
         Top             =   855
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
         Left            =   1320
         TabIndex        =   31
         Top             =   420
         Width           =   1800
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -510
         TabIndex        =   30
         Top             =   420
         Width           =   1770
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
         Left            =   3960
         TabIndex        =   29
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command1 
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
      Height          =   900
      Left            =   10680
      Picture         =   "CIERREP09.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5820
      Width           =   765
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "TRABAJOS REALIZADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   80
      TabIndex        =   19
      Top             =   5445
      Width           =   10540
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7350
         Locked          =   -1  'True
         TabIndex        =   42
         ToolTipText     =   "Doble Clic Para Desplegar Lista de Técnicos"
         Top             =   400
         Width           =   3090
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   890
         Left            =   70
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   360
         Width           =   7140
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Realizó:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7350
         TabIndex        =   41
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7350
         TabIndex        =   23
         Top             =   945
         Width           =   3090
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario Activo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   7350
         TabIndex        =   20
         Top             =   735
         Width           =   1320
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      Caption         =   "REPUESTOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3015
      Left            =   80
      TabIndex        =   16
      Top             =   2400
      Width           =   10545
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   2100
         Left            =   120
         TabIndex        =   52
         Top             =   1200
         Width           =   10365
         _ExtentX        =   18283
         _ExtentY        =   3704
         _Version        =   393216
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Roboto Mono"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton command15 
         Caption         =   "..."
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
         Left            =   6240
         Picture         =   "CIERREP09.frx":091E
         TabIndex        =   36
         Top             =   280
         Width           =   175
      End
      Begin VB.Frame Frame7 
         Caption         =   "                                                                                                                  "
         Height          =   2120
         Left            =   8760
         TabIndex        =   25
         Top             =   840
         Width           =   15
      End
      Begin VB.Frame Frame5 
         Caption         =   "                                                                                                                  "
         Height          =   2160
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   15
      End
      Begin VB.Frame Frame2 
         Caption         =   "                                                                                                                  "
         Height          =   2120
         Left            =   7320
         TabIndex        =   22
         Top             =   840
         Width           =   15
      End
      Begin VB.Frame Frame3 
         Caption         =   "                                                                                                                  "
         Height          =   2120
         Left            =   1950
         TabIndex        =   21
         Top             =   840
         Width           =   15
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1830
         ItemData        =   "CIERREP09.frx":0C28
         Left            =   70
         List            =   "CIERREP09.frx":0C2A
         TabIndex        =   17
         Top             =   1000
         Width           =   10395
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5765
         TabIndex        =   39
         Top             =   280
         Width           =   495
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito de Consumo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   38
         Top             =   360
         Width           =   2610
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6490
         TabIndex        =   37
         Top             =   280
         Width           =   3975
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Código         Descripción                                        Cant. Lote"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   75
         TabIndex        =   18
         ToolTipText     =   "Doble Click Para Editar Repuestos"
         Top             =   680
         Width           =   10395
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00CDDADA&
      Caption         =   "CODIGO - MODELO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2235
      Left            =   80
      TabIndex        =   5
      Top             =   100
      Width           =   10545
      Begin VB.Frame Frame9 
         BackColor       =   &H00CDDADA&
         Caption         =   "Técnico Externo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   1
         Left            =   1110
         TabIndex        =   45
         Top             =   200
         Width           =   5175
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   210
            Width           =   375
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Talón Nro"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Left            =   2880
            TabIndex        =   50
            Top             =   225
            Width           =   720
         End
         Begin VB.Label Label17 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   3675
            TabIndex        =   49
            Top             =   210
            Width           =   1335
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Técnico:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   -40
            TabIndex        =   48
            Top             =   225
            Width           =   720
         End
         Begin VB.Label Label14 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Left            =   1155
            TabIndex        =   47
            Top             =   210
            Width           =   1575
         End
      End
      Begin VB.Label TEPRUEBA 
         AutoSize        =   -1  'True
         Caption         =   "TEPRUEBA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   0
         TabIndex        =   53
         Top             =   0
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Accesorios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   4920
         TabIndex        =   15
         Top             =   945
         Width           =   1245
      End
      Begin VB.Label Label28 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   5655
         TabIndex        =   14
         Top             =   1200
         Width           =   4725
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Serie"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   5760
         TabIndex        =   13
         Top             =   900
         Width           =   1245
      End
      Begin VB.Label Label26 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7125
         TabIndex        =   12
         Top             =   840
         Width           =   3255
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   9255
         TabIndex        =   11
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Talón Nro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   7320
         TabIndex        =   10
         Top             =   330
         Width           =   1800
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Defectos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   30
         TabIndex        =   9
         Top             =   1155
         Width           =   1005
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1110
         TabIndex        =   8
         Top             =   840
         Width           =   3840
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Equipo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   30
         TabIndex        =   7
         Top             =   900
         Width           =   1005
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   1110
         TabIndex        =   6
         Top             =   1200
         Width           =   3840
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   10680
      ScaleHeight     =   5745
      ScaleWidth      =   1995
      TabIndex        =   3
      Top             =   0
      Width           =   2025
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
         Height          =   650
         Left            =   120
         Picture         =   "CIERREP09.frx":0C2C
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Registro de Armado EYM"
         Top             =   3120
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   100
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   43
         Top             =   5000
         Visible         =   0   'False
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   44
            Top             =   0
            Width           =   12000
         End
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
         Height          =   650
         Left            =   120
         Picture         =   "CIERREP09.frx":0F36
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Desvio de Consumos"
         Top             =   3930
         Width           =   650
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
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
         Left            =   120
         Picture         =   "CIERREP09.frx":1080
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Cierra y Abondona la Aplicación"
         Top             =   80
         Width           =   650
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
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
         Height          =   650
         Left            =   110
         Picture         =   "CIERREP09.frx":11CA
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   800
         Width           =   650
      End
   End
End
Attribute VB_Name = "CIERREP09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub CARGAR_DETALLE_EN_CARDETPE(NUPEDIDO&)
     Call SETULTREG("!CARDETPE", 0)
     JJ& = 0
     REBLA$ = REGBLAN$("CARDETPE")
     
     SQLX$ = "SELECT * FROM PEDDETA WITH(NOLOCK)"
     SQLX$ = SQLX$ + "WHERE NroPed = " & NUPEDIDO& & " "
     SQLX$ = SQLX$ + "ORDER BY NuorItem ASC "
     Set PEDDETTEMP = READSET(SQLX$)
     
     YAMONE% = 0
     Do While Not PEDDETTEMP.EOF
         If YAMONE% = 0 Then
           YAMONE% = 1
           MONEMI% = PEDDETTEMP!MONEEMIS
           MONEMA% = MONEMI%
           If MONEMI% < 1 Then MONEMI% = 1
               MODPED07.MONEMIS = TRIM$(DEMONECO$(MONEMI%))
               MODPED07.TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
           End If
         X$ = REBLA$
         CIXI% = XVALO(PEDDETTEMP!CODIINT)
         Call REPLA(X$, MKI$(CIXI%), 1, 2)
         Call REPLA(X$, PEDDETTEMP!tcolor, 3, 16)
         
         Call REPLA(X$, PEDDETTEMP!ttalle, 19, 8)
           TLAR = PEDDETTEMP!TLargo
         Call REPLA(X$, MKS$(TLAR), 27, 4)
           TCHO = PEDDETTEMP!TAncho
         Call REPLA(X$, MKS$(TCHO), 31, 4)
           TESP = PEDDETTEMP!TEspesor
         Call REPLA(X$, MKS$(TESP), 35, 4)
         Call REPLA(X$, MKS$(PREL), 39, 4)
           PDESCU = 0
           If Abs(XVALO(PEDDETTEMP!PreList)) >= 0.005 Then
              PDESCU = 100 * (XVALO(PEDDETTEMP!PreList) - XVALO(PEDDETTEMP!PreUnid)) / XVALO(PEDDETTEMP!PreList)
           End If
           If Abs(PDESCU) < 0.05 Then PDESCU = 0
         Call REPLA(X$, MKS$(PDESCU), 69, 4)
           '
         PREL = PREL * (100 - PDESCU) / 100
         PDESCU = 0
         Call REPLA(X$, MKS$(PREL), 39, 4)
         Call REPLA(X$, MKS$(PDESCU), 69, 4)
           '
           CAN = PEDDETTEMP!CanPed
         Call REPLA(X$, MKS$(CAN), 47, 4)
           FSE% = CVINT(PEDDETTEMP!FeSolEnt)
         Call REPLA(X$, MKI$(FSE%), 63, 2) ' FECHA SOLICITADA DE ENTREGA
           NLISSTA = XVALO(PEDDETTEMP!lista_de_item)
           Call REPLA(X$, MKS$(NLISSTA), 85, 4)
           
           IDIT$ = PEDDETTEMP!TPedItem
         Call REPLA(X$, IDIT$, 93, 16)
         Call REPLA(X$, MKS$(0), 117, 4)   ' ESTO SE ACTUALIZA MAS ABAJO
         Call REPLA(X$, MKI$(0), 121, 2)
         
         
           SUCOD$ = ""
           If IsNull(PEDDETTEMP!SubCod) = False Then SUCOD$ = PEDDETTEMP!SubCod
         Call REPLA(X$, SUCOD$, 109, 8)
         JJ& = JJ& + 1
         Call GRAREG("!CARDETPE", X$, JJ&)
         PEDDETTEMP.MoveNext
       Loop
     PEDDETTEMP.Close
     Set PEDDETTEMP = Nothing
     
     Call SETULTREG("!CARDETPE", JJ&)
     Call CIERRARCH("CARDETPE")

End Sub

'Sub MOSTRARREMITORP(NROREP&, NCLI_TECNICO, REMI$)
'    REMI$ = VALOBADA("REPARA", "NUMREMITO", "NUMREPA = " & NROREP&, "NOMESS")
'    '
'    CODXT$ = ""
'    CI% = XVALO(VALOBADA("REPARA", "COD_INTE", "NUMREPA = " & NROREP&, "NOMESS"))
'    If CI% > 0 Then
'       CODXT$ = CODEXT$(CI%)
'    End If
'    '
'
'    If TRIM$(REMI$) = "" Then
'       PRONREM& = 1 + XVALO(VALOBADA("REPARA", "max(cast(substring(NUMREMITO, 9, 8)as int))", "NUMREPA > 0", "NOMESS"))
'       REP$ = TRIM$(FORMATNUM$(PRONREM&, "F8.0"))
'       While Len(REP$) < 8: REP$ = "0" + REP$: Wend
'       REMI$ = "RP-0001-" & REP$
'    End If
'    '
'    'NCLIE = XVALO(VALOBADA("REPARA", "NUME_CLI", "NUMREPA = " & NROREP&, "NOMESS"))
'
'    FREP = HOY(HO$)
'    FILTX$ = TRIM$(Str$(NROREP&)) & ";" & TRIM$(REMI$) & ";" & TRIM$(Str$(FREP)) & ";" & TRIM$(CODXT$)
'    Call STDFORM("REMX-MDT", FILTX$)
'    If STDSCREEN.APRO <> "OK" Then
'       GoTo 15
'    End If
'    'grabo fecha de reparacion y numero de reparacion en tanumse
'     NUSERI$ = VALOBADA("REPARA", "NUMEROSERIE", "NUMREPA = " & NROREP&, "NOMESS")
'     FE% = FREP
'     FE1% = DIASPOST(FE%, 180)
'     FEVEGREP = FETEXCOR1$(FE1%)
'     SQLX$ = "SELECT * FROM REPARA "
'      SQLX$ = SQLX$ + " WHERE  NUMREPA = " & NROREP&
'      Dim RS2 As ADODB.Recordset
'      Set RS2 = New ADODB.Recordset
'27              On Error Resume Next
'      RS2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'       '
'      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'             On Error GoTo 0
'             Call CapturaErrorOpen
'             GoTo 27
'      End If
'      On Error GoTo 0
'
'      With RS2
'          Do While Not .EOF
'            !Status = 5
'            !FECHENT = CVDAT(FE%)
'            .Update
'            .MoveNext
'          Loop
'      End With
'      RS2.Close
'      Set RS2 = Nothing
'      '
'      SQLX$ = "SELECT * FROM TANUMSE "
'      SQLX$ = SQLX$ + " WHERE  NUMEROSERIE = '" & NUSERI$ & "'"
'      Dim RS3 As ADODB.Recordset
'      Set RS3 = New ADODB.Recordset
'28          On Error Resume Next
'      RS3.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'    '
'      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'             On Error GoTo 0
'             Call CapturaErrorOpen
'             GoTo 28
'      End If
'      On Error GoTo 0
'
'      With RS3
'       Do While Not .EOF
'        !FevenGaRep = CVDAT(FE1%)
'        !UnumRep = NROREP&
'        .Update
'        .MoveNext
'       Loop
'      End With
'      RS3.Close
'      Set RS3 = Nothing
'      '
'      CUERPO$ = "Equipo Remitido" & vbCrLf
'      CUERPO$ = CUERPO$ + "Número de Reparación: " & NROREP& & vbCrLf '
'      CUERPO$ = CUERPO$ + "Número de Remito: " & REMI$ & vbCrLf
'      CUERPO$ = CUERPO$ + "Cliente: " & rasocli$(NCLIE)
'      Call ANOTANUSE(NUSERI$, CUERPO$, , "NOMESS")
'
'     'BORRO LOS REGISTROS DEL REPUREPA DE ESTA ORDEN (ANTES SE BORRABA EN EL CIERRE, SE TRASLADO AQUI POR QUE LUEGO DEL CIERRE SE PUEDE MODIFICAR EL PRESUPESTO
'     CONDI$ = "NUMREPA = " & NROREP& & ""
'     Call BORRAREGISTROS("REPUREPA", CONDI$, REGAF&, "NOMESS")
'
'End Sub

Sub Command1_Click()
    Command1.Enabled = False
    '
    Call ABREORFAB
    NUREP& = XVALO(Label25)
    NUMEREPA$ = "OR-" & TRIM$(Str$(NUREP&))
    '
    If ECOARCH$("TADEPOSI", Chr$(XVALO(Label13))) = "" Then
       Call COMUNI("Falta Elegir Depósito de Consumo")
       Command1.Enabled = True
       Exit Sub
    End If
    '
    If TRIM$(Text1) = "" And XVALO(Text5) < 1 Then ' SI TIENE TECNICO EXTERNO NO VALIDA USUARIO DE REPARACION
       Call COMUNI("Falta Elegir Usuario Realizó Reparación")
       Command1.Enabled = True
       Exit Sub
    End If
    '
    '***** EYM ******
    EPAC$ = TRIM$(EMPREAC$)
    If InStr(EPAC$, "EYM") > 0 Then
      Call CierreRepaEym(NUREP&, OK%)
      Command1.Enabled = True
      If OK% < 0 Then Exit Sub 'PARA QUE SE QUEDE EN EL FORMULARIO DE CIERRE
      Call Command2_Click      'SE CIERRA
      Exit Sub
    End If
    '***** EYM ******
    
    'SE AGREGA PARA QUE SE PUEDA CERRAR LA OR SIN NECESIDAD DE CONTAR CON LOS COMPONENTES
    'Y ASI DEPURAR LAS O/REPARACION VIEJAS
    'SE QUITA POR QUE SI NO SE CONSUME.
'    If DERACCE%("CIERSINC/NOMESS", "CIERRE SIN CONTROL FINAL") > 1 Then GoTo 10
    '
    
    Dim MALCONSU%(32767): Dim DESCRIMAT$(32767)
    DepDest% = XVALO(CONTROL$("", "DEPREDES"))

    CUERPO4$ = ""
   'SE CONSIDERA QUE SI ES CON TEC.EXT.
    TECEXT& = TECNICO_REPA&(NUREP&)
    If TECEXT& > 0 Then
       CUERPO4$ = CUERPO4$ + "Técnico Externo: " & TECEXT&
       DepDest% = XVALO(CONTROL$("", "DPREPPAÑ"))
    End If
    
    FECHA% = HOY(HO$)
    CMOV$ = "RP"
    NCOMP& = ProNroComprobante(CMOV$)
    'NCOMP& = 1 + XVALO(VALOBADA("MOVISTO", "MAX(DOPRINUM)", "CODIMOVI='RP'", "NOMESS"))
    DOPRI$ = CMOV$ & "." & NCOMP&
    
    DOSEC$ = "CR." & NCOMP&
    REFE$ = "Cons. de Repar. " & NUREP& ' NO CAMBIAR TEXTO LONGITUD SE USA UN MID PARA TOMAR EL NRO DE REPARACION
    For i& = 1 To List1.ListCount ' VALIDACIONES DEL LOTE
        TTXX$ = List1.List(i& - 1)
        CODIEX$ = Mid$(TTXX$, 1, 16)
        CI% = CODINT(CODIEX$)
        If CI% > 0 Then
          ' CANTIDAD RESERVADA
          CONDI$ = "IDSUBOR= '" & NUMEREPA$ & "' AND COD_INTE = " & CI% & ""
          CANTIRESER# = XVALO(VALOBADA("RESERVA", "SUM(CANTRES)", CONDI$, "NOMESS"))
          ' CANTIDAD INGRESADA EN MOVISTO
          CONDI$ = "Cod_Inte = " & CI% & " AND DOSECCOR = '" & TRIM$(NUMEREPA$) & "' And DepoMovi = " & DepDest%
          CANTING = XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", CONDI$, "NOMESS"))
          ' CANTIDAD SALIDA MOVISTO
          CONDI$ = "CodiMovi = 'TR' AND Cod_Inte = " & CI% & " AND DOPRICOR = '" & TRIM$(NUMEREPA$) & "' AND DepoMovi = " & DepDest%
          CANTISL = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", CONDI$, "NOMESS"))
          CANTOTA# = XVALO(CANTING) - XVALO(CANTISL)
          '
          If CANTIRESER# - CANTOTA# > 0.05 Then
             Call COMUNI("Material Insuficiente para Código: " & CODEXT$(CI%))
             MALCONSU%(CI%) = 1  'CARGO VECTOR PARA GUARDAR INFORMACION DE LOS ITEMS Q´NO FUERON TRANSFERIDOS
             DESCRIMAT$(CI%) = TRIM$(Mid$(TTXX$, 17, 40))
             'command1.Enabled = True
'             Exit Sub SOLO AVISA Y PERMITE CONTINUAR
          End If
          '
          If ESTRAZABLE%(CI%) = 1 Then
              If TRIM$(Mid$(TTXX$, 72, 12)) = "" Then
                 Call COMUNI("Falta Lote Para Código: " & CODEXT$(CI%))
                 If MALCONSU%(CI%) < 1 Then MALCONSU%(CI%) = 2: DESCRIMAT$(CI%) = TRIM$(Mid$(TTXX$, 17, 40)) 'CARGO VECTOR PARA GUARDAR INFORMACION DE LOS ITEMS Q´NO FUERON TRANSFERIDOS
                  'command1.Enabled = True
                  'Exit Sub SOLO AVISA Y PERMITE CONTINUAR
              End If
          End If
        End If
20  Next i&
    '
    If XVALO(Text5) > 0 Then
      NCLI_TECNICO = NCLI_ASOC_TECNICO(NUREP&)
    End If
    'CARGO EL TEXTO PARA EL ANOTADOR DE LOS ITEMS QUE SE CONSUMIERON DE MANERA NO NORMALIZADA (SIN STOCK TRANSFERIDO O SIN LOTE)
    PUSOTITULO% = 0: CUERPO3$ = ""
    For L% = 0 To UBound(MALCONSU%)
'       If L% = 1 Then
'         OPX% = COMALTER%("Confirma el Cierre de La O.Reparación\Sin Importar los Avisos sobre Faltantes (No Transferidos)\\Cancelar\Cerrar O.Reparación")
'         If OPX% < 2 Then Command1.Enabled = True: Exit Sub
'       End If
       '
       If MALCONSU%(L%) = 1 Then
          If PUSOTITULO% = 0 Then
             CUERPO2$ = "Items No Consumidos de Manera Normalizada" + vbCrLf
             PUSOTITULO% = 1
          End If
          CUERPO3$ = CUERPO3$ + DESCRIMAT$(L%) + vbCrLf
          CUERPO3$ = CUERPO3$ + "Material Insuficiente " + vbCrLf & String$(56, "-") & vbCrLf
       ElseIf MALCONSU%(L%) = 2 Then
          If PUSOTITULO% = 0 Then
             CUERPO2$ = "Items No Consumidos de Manera Normalizada" + vbCrLf
             PUSOTITULO% = 1
          End If
          CUERPO3$ = CUERPO3$ + DESCRIMAT$(L%) + vbCrLf
          CUERPO3$ = CUERPO3$ + "Falta Lote" + vbCrLf & String$(56, "-") & vbCrLf
       End If
       If L% >= 32767 Then Exit For
    Next L%
    '
    
    Screen.MousePointer = 11
    coefi = Val(CONTROL$("REPARA", "COEFPREA")) ' COEFICIENTE SETEADO EN SETUP
    If coefi < 0.05 Then coefi = 1
    '
    ACUMUPRECIUNIT# = 0
    PIVA% = PIVACLI%(NCLI_TECNICO)
    
    TECEXT& = TECNICO_REPA&(NUREP&)
    ESTAENGARANTIA% = XVALO(VALOBADA("REPARA", "Status_Gar", "NUMREPA = " & NUREP& & "", "NOMESS"))
    If TECEXT& > 0 And ESTAENGARANTIA% = 0 Then GoTo 55
    
    For i& = 1 To List1.ListCount
         ZZ$ = List1.List(i& - 1)
         CODIEX$ = Mid$(ZZ$, 1, 16)
         CI% = CODINT(CODIEX$)
         CANTI = XVALO(Mid$(ZZ$, 64, 8))
         'PRECIUNIT# = COSUNI(CI%) * COEFI
         PRECIUNIT# = XVALO(VALOBADA("REPUREPA", "PRUN_NETO", "NUMREPA = " & NUREP& & " AND COD_INTE = " & CI%, "NOMESS"))
         'PARA EL CASO CLIENTE TIERRA DEL FUEGO, TRAE EL COSTO CALCULADO
         If PIVA% = 7 And PRECIUNIT# <= 0 Then
            PRECIUNIT# = COSCALC#(CI%)
            If PRECIUNIT# <= 0 Then PRECIUNIT# = COSUNI(CI%)
         End If
         'FIN TIERRA DEL FUEGO

         'DESCU$ = TRIM$(VALOBADA("REPUREPA", "PORDESC", "NUMREPA = " & NUREP& & " AND COD_INTE = " & CI%, "NOMESS"))
         'PRECIUNIT# = ImporteConDescuento#(PRECIUNIT#, DESCU$)
         LOTEX$ = ""
         DEPX% = DepDest%
         If ESTRAZABLE%(CI%) = 1 Then
          LOTEX$ = TRIM$(Mid$(ZZ$, 74, 12))
          DEPX% = DEPOLOT%(CI%, LOTEX$)
         End If
         '
          Call MOVISTOK(FECHA%, CI%, LOTEX$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, PRECIUNIT#, MONEII%, NCLI_TECNICO, DEPX%, (-1) * CANTI, , , , , , PRECIUNIT#, , , NUREP&)
          Call ACSALOTE(CI%, LOTEX$, "NOMESS")
          ACUMUPRECIUNIT# = ACUMUPRECIUNIT# + PRECIUNIT#
         '
40  Next i&

55  Call CIERRARCH("!MATEROF")
    
    ' ORDEN CERRADA Y MODIFICADA - DA POR CUMPLIDO EL CONSUMO DE LOS NUEVOS REPUESTOS
    Call ACTUREGISTRO("REPUREPA", "MARCAITEMAGREGADO", "NUMREPA=" & CStr(NUREP&), 0, REGAF&)
    
'10: CONDI$ = "NUMREPA = " & NUREP& & ""
     'HAY QUE BORRAR EN EL REMITO
'    Call BORRAREGISTROS("REPUREPA", CONDI$, REGAF&, "NOMESS")
    '
60  SQLX$ = "SELECT Trab_Efect, REPARACION, STATUS, UsRealizaRepa, FechaCierreRepa FROM REPARA"
    SQLX$ = SQLX$ + " WHERE NUMREPA = " & NUREP&
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With rs
      Do While Not .EOF
        '!TRAB_EFECT = Text3
        !REPARACION = Text3
        !Status = 4
        !UsRealizaRepa = Left$(TRIM$(Text1), 32)
        !FechaCierreRepa = CVDAT(HOY(HO$))
        .Update
        .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    ' BORRA RESERVAS
    CONDI$ = "IdSubOr = '" & NUMEREPA$ & "'"
    Call BORRAREGISTROS("Reserva", CONDI$, RAFE&, "NOMESS")
    '
    NUMESER$ = TRIM$(Label26)
    '
    RACLI$ = VALOBADA("REPARA", "RASO_CLI", "NUMREPA = " & NUREP& & "", "NOMESS")
    CUERPO$ = "Cierre de Orden de Reparación" & vbCrLf
    CUERPO$ = CUERPO$ + "Número de Reparación: " & NUREP& & vbCrLf
    CUERPO$ = CUERPO$ + "Cliente: " & RACLI$
    If TRIM$(CUERPO2$) <> "" Then
       CUERPO$ = CUERPO$ & vbCrLf & vbCrLf
       CUERPO$ = CUERPO$ + CUERPO2$ '& vbCrLf
       CUERPO$ = CUERPO$ + CUERPO3$ '& vbCrLf
       CUERPO$ = CUERPO$ + CUERPO4$
    End If
    
    Call ANOTANUSE(NUMESER$, CUERPO$, , "NOMESS")
    '
    '=========================================================================
    TECEXT& = TECNICO_REPA&(NUREP&)
    ESTAENGARANTIA% = XVALO(VALOBADA("REPARA", "Status_Gar", "NUMREPA = " & NUREP& & "", "NOMESS"))
    If TECEXT& > 0 Then                'SI TIENE TECNICO EXTERNO
       NCLI_TECNICO = NCLI_ASOC_TECNICO(NUREP&) ' lo lee mas arriba
       Call ACTUALIZA_GARANT_UNUREP(NUREP&, NUMESER$)   'ACTUALIZA EN TANUMSE
       If ESTAENGARANTIA% > 0 Then 'SI ESTA EN GARANTIA
          Call CIERREP09.GENERA_MOV_CBLE(NUREP&) 'GENERA EL MOVIMIENTO POR MAS QUE SEA T.FUEGO POR QUE LA FACTURA SE VA A ANULAR.
          'SI EL CLIENTE ASOCIADO AL TECNICO ES P.IVA = 7 SE GENERA EL PEDIDO.
          PIV% = XVALO(VALOBADA("DEUDOR12", "PIVA_CLI", "NUME_CLI = " & NCLI_TECNICO, "NOMESS"))
          If PIV% = 7 Then
             Call GENERAPEDIDO(NUREP&, DOPRI$, NPED&, OKX%)
             If NPED& > 0 Then Call CIERREP09.CARGAELPEDIDO(NPED&)
             If OKX% > 0 Then
               Call ACTUREGISTRO("REPARA", "STATUS", "NUMREPA = " & NUREP&, 5, REGAF&, "NOMESS")
               CUERPO4$ = vbCrLf & "Se Generó Orden de Venta N°: " & NPED& & vbCrLf
               CUERPO4$ = CUERPO4$ + " En Forma Autómatica a Técnico Tierra del Fuego"
               Call ANOTANUSE(NUMESER$, CUERPO4$, , "NOMESS")
               Call ACTUREGISTRO("REPARA", "NROPED", "NUMREPA = " & NUREP&, NPED&, REGAF&, "NOMESS")  'GUARDO EL NRO DE PEDIDO ASOCIADO.
               If NPED& > 0 Then Call COMUNI("Atención !!! \Se Ha Generado el Pedido N°: " & NPED&)
             End If
          Else ' SI NO ES DE TIERRA DEL FUEGO
             Call CIERREP09.PRI_REMI_TEC_EXTERNO(NCLI_TECNICO, DOPRI$)
             'Call MOSTRARREMITORP(NUREP&, NCLI_TECNICO)
             Call ACTUREGISTRO("REPARA", "NumRemito", "NUMREPA = " & NUREP&, DOPRI$, REGAF&, "NOMESS")
             Call ACTUREGISTRO("REPARA", "STATUS", "NUMREPA = " & NUREP&, 5, REGAF&, "NOMESS")
             Call COMUNI("Atención !!! \Se Ha Generado el Remito N°: " & DOPRI$)
         End If
       Else 'SI NO ESTA EN GARANTIA
         Call GENERAPEDIDO(NUREP&, DOPRI$, NPED&, OKX%)
         If NPED& > 0 Then Call CIERREP09.CARGAELPEDIDO(NPED&)
         If OKX% > 0 Then
               Call ACTUREGISTRO("REPARA", "STATUS", "NUMREPA = " & NUREP&, 5, REGAF&, "NOMESS")
               CUERPO4$ = vbCrLf & "Se Generó Orden de Venta N°: " & NPED& & vbCrLf
               CUERPO4$ = CUERPO4$ + " En Forma Autómatica - No en Garantía "

               Call ANOTANUSE(NUMESER$, CUERPO4$, , "NOMESS")
               Call ACTUREGISTRO("REPARA", "NROPED", "NUMREPA = " & NUREP&, NPED&, REGAF&, "NOMESS")  'GUARDO EL NRO DE PEDIDO ASOCIADO.
               If NPED& > 0 Then Call COMUNI("Atención !!! \Se Ha Generado el Pedido N°: " & NPED&)
          End If
       End If
   Else 'SI NO TIENE TECNICO Y SI ESTA EN GARANTIA
       'SE COMENTO POR QUE EN ESTE CASO SE REALIZA CUANDO SE HACE EL REMITO
'      If ESTAENGARANTIA% > 0 Then 'SI NO TIENE TECNICO PERO ESTA EN GARANTIA GENERA EL MOVIMIENTO CONTABLE
'         Call CIERREP09.GENERA_MOV_CBLE(NUREP&)
'      End If
   End If
  '=========================================================================
    Screen.MousePointer = 1
    Command1.Enabled = True
    Call Command2_Click ' SE CIERRA
End Sub
Sub CARGAELPEDIDO(NUPEDIDO&)
     Dim RSX As New ADODB.Recordset
     SQLX$ = "SELECT MAX(NROPED) AS M FROM PEDENCA WITH(NOLOCK)"
     Set rs = FLEXCONN.Execute(SQLX$)
     AAA = rs!m
     'SQLX$ = SQLX$ + " WHERE NROPED = 1"
'     Set RSX = READSET(SQLX$)
'     If RSX.EOF Then
'        RSX.Close
'        Set RSX = Nothing
'        Exit Sub
'     End If
'     AAA = RSX!m
     
     SQLX$ = "SELECT * FROM PEDENCA WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE NroPed = " & NUPEDIDO&
     SQLX$ = SQLX$ + " AND CodiEmpr = " & CodiEmp%
     Set RSX = FLEXCONN.Execute(SQLX$)
'     If RSX.EOF Then
'        RSX.Close
'        Set RSX = Nothing
'        Exit Sub
'     End If
            '
            FEX = CVINT(RSX!FECHEMIS)
            MODPED07.Text9.TEXT = FECHATEX$(FEX)
            MODPED07.Text11.TEXT = TRIM$(Str$(NUPEDIDO&))
              NCLIE = XVALO(RSX!NROCLIE)
            MODPED07.LTEXT1 = TRIM$(Str$(NCLIE))
            MODPED07.Text1.TEXT = TRIM$(Str$(NCLIE))
            '
              NCLIEN = NCLIE
            MODPED07.Text2.TEXT = rasocli$(NCLIEN)
            MODPED07.Text3.TEXT = DOMICLI$(NCLIEN)
            MODPED07.Text4.TEXT = CPOSCLI$(NCLIEN)
            MODPED07.Text5.TEXT = LOCACLI$(NCLIEN)
            MODPED07.Text7.TEXT = CSTRING$(MKS$(LIMICRE(NCLIEN)), 3, 13, 2, 2)
            MODPED07.Text12.TEXT = TRIM$(POSIVAC$(NCLIEN)) + " - CUIT " + CuitCli$(NCLIEN)
              CREDISPO# = LIMICRE(NCLIEN) - SALDCLI#(NCLIEN) - CHEPEACRE#(NCLIEN)
            MODPED07.Text8.ForeColor = RGB(0, 0, 0)
              If CREDISPO# < 0 Then MODPED07.Text8.ForeColor = RGB(255, 0, 0)
            MODPED07.Text8.TEXT = CSTRING$(MKD$(Abs(CREDISPO#)), 3, 13, 2, 2)
            '
            MODPED07.Text10.TEXT = RSX!NROOCOM
            MODPED07.TEXT21(0).TEXT = TRIM$(Str$(RSX!ListPre))
            MODPED07.TEXT21(1).TEXT = TRIM$(Str$(RSX!CondPag))
              If XVALO(MODPED07.TEXT21(1).TEXT) < 1 Then MODPED07.TEXT21(1).TEXT = Str$(CPAGCLI%(NCLIEN))
            MODPED07.TEXT21(2).TEXT = TRIM$(Str$(RSX!VENDED))
              If XVALO(MODPED07.TEXT21(2).TEXT) < 1 Then MODPED07.TEXT21(2).TEXT = Str$(VENDCLI%(NCLIEN))
              If XVALO(MODPED07.TEXT21(2).TEXT) < 1 Then MODPED07.Eco21(2).TEXT = "GERENCIA"
            POCOMI = XVALO(RSX!PorComVend)
              If Left$(UCase$(MODPED07.Label21(4)), 5) = "LETRA" Then
                   MODPED07.Text44 = Chr$(64 + POCOMI)
                 Else
                   MODPED07.Text44 = TRIM(FORMATNUM$(POCOMI, "F4.1"))
              End If
            MODPED07.TEXT21(3).TEXT = TRIM$(Str$(XVALO(RSX!TIPOINGRE)))
            MODPED07.TXTFLETE = SAFETEXT$(RSX!FLETE_TEX)       '
            MODPED07.PORDESCU.TEXT = SAFETEXT$(RSX!TDescApli)
            MODPED07.LENTREGA(0) = SAFETEXT$(RSX!DOMIENT)
            MODPED07.LENTREGA(1) = SAFETEXT$(RSX!LOCAENT)
            MODPED07.LENTREGA(2) = SAFETEXT$(RSX!TranspEnt)
            '
            MODPED07.Text19.TEXT = SAFETEXT$(RSX!NTranspo)
            MODPED07.Text20.TEXT = SAFETEXT$(RSX!NSucent)
            '
            MODPED07.Text13.TEXT = SAFETEXT$(RSX!Observa)
            MODPED07.Label11.ForeColor = RGB(0, 0, 255)
            MODPED07.Label11 = "Venta"
            '
            'MOTIVO DE PEDIDO
            MODPED07.Text16 = XVALO(RSX!OV_Motivo)
            NUMOT% = XVALO(MODPED07.Text16)
            If EXISTE(LUDAT$ + "\TAMOTPRE.RFS") > 0 Then
               MODPED07.Label28 = ECOARCH$("TAMOTPRE", Chr$(NUMOT%))
            End If
            '
            If RSX!VenRep = "R" Then
               MODPED07.Label11.ForeColor = RGB(255, 0, 0)
               MODPED07.Label11 = "Reparación"
               MODPED07.Label11.ToolTipText = "Doble-Click para 'Venta'"
            End If
            '
            
            If InStr(EPAC$, "GABAL") > 0 Then
               DESCURETIRA$ = TRIM$(SAFETEXT$(RSX!Dto_Ret_Cte))
               If DESCURETIRA$ <> "" Then
                  Call MODPED07.Image1_Click
                  MODPED07.Check2.Value = 1
                  MODPED07.Label31 = DESCURETIRA$
               End If
              DESCURETIRA$ = TRIM$(SAFETEXT$(RSX!DtoxCant))
               If DESCURETIRA$ <> "" Then
                  Call MODPED07.Image1_Click
                  MODPED07.Text18 = DESCURETIRA$
                  MODPED07.Labeldto.Visible = True
                  MODPED07.Text17.Visible = True
                  MODPED07.Text18.Visible = True
               End If
            End If
            TEXTOFLETE$ = SAFETEXT$(RSX!FLETE_TEX)
            MODPED07.TXTFLETE = TEXTOFLETE$
     RSX.Close
     Set RSX = Nothing
     
     '
     Call SETULTREG("!CARDETPE", 0)
     JJ& = 0
     REBLA$ = REGBLAN$("CARDETPE")
     
     SQLX$ = "SELECT * FROM PEDDETA "
     SQLX$ = SQLX$ + "WHERE NroPed = " & NUPEDIDO& & " "
     '\\\OT-06-0520-RG-01/12/06///
     SQLX$ = SQLX$ + "AND Status < 8 "
     SQLX$ = SQLX$ + "AND CodiEmpr = " & CodiEmp%
     '\\\OT-06-0520-RG-FIN///
     SQLX$ = SQLX$ + "ORDER BY NuorItem ASC "
     'Set PEDDETTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     'jandy sql
'     Dim PEDDETTEMP As ADODB.Recordset
'     Set PEDDETTEMP = New ADODB.Recordset
'     PEDDETTEMP.CursorLocation = adUseClient
'27   On Error Resume Next
'     PEDDETTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'        On Error GoTo 0
'        Call CapturaErrorOpen
'        GoTo 27
'     End If
'     On Error GoTo 0
'     'jandy esto asi estaba
'
'     YAMONE% = 0
'     On Error Resume Next
'     PEDDETTEMP.MoveFirst
'     If Err < 1 Then
      Set PEDDETTEMP = FLEXCONN.Execute(SQLX$)
      Do While Not PEDDETTEMP.EOF
         If YAMONE% = 0 Then
           YAMONE% = 1
           MONEMI% = PEDDETTEMP!MONEEMIS
           MONEMA% = MONEMI%
           If MONEMI% < 1 Then MONEMI% = 1
           'If MONEMI% = 1 Then
               MODPED07.MONEMIS = TRIM$(DEMONECO$(MONEMI%))
               MODPED07.TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
           '  Else
           '    MONEMI% = 2
           '    MODPED07.MONEMIS = "Dólares"
           'End If
         End If
         X$ = REBLA$
           CIXI% = PEDDETTEMP!CODIINT
         Call REPLA(X$, MKI$(CIXI%), 1, 2)
         Call REPLA(X$, SAFETEXT$(PEDDETTEMP!tcolor), 3, 16)
         
         Call REPLA(X$, SAFETEXT$(PEDDETTEMP!ttalle), 19, 8)
           TLAR = SAFETEXT$(PEDDETTEMP!TLargo)
         Call REPLA(X$, MKS$(TLAR), 27, 4)
           TCHO = XVALO(PEDDETTEMP!TAncho)
         Call REPLA(X$, MKS$(TCHO), 31, 4)
           TESP = XVALO(PEDDETTEMP!TEspesor)
         Call REPLA(X$, MKS$(TESP), 35, 4)
           'PREL = PEDDETTEMP!PreUnid * TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
           VENTANACARGA$ = TRIM$(CONTROL$("", "VENCARPE"))
         If UCase$(VENTANACARGA$) <> "CARDETLO" Then
           PREL = PEDDETTEMP!PreList * TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
         ElseIf UCase$(VENTANACARGA$) = "CARDETLO" Then
           'VENTANA CARDETLO ()VENTANA DE CARGA CON LOTES DE ARRASTRE DE PRESUPUESTO AHP
           PREL = XVALO(PEDDETTEMP!PreUnid) * TICAMONE(MONEMA%) / TICAMONE(MONEMI%)
         End If
         Call REPLA(X$, MKS$(PREL), 39, 4)
           PDESCU = 0
           If Abs(XVALO(PEDDETTEMP!PreList)) >= 0.005 Then
              PDESCU = 100 * (XVALO(PEDDETTEMP!PreList) - XVALO(PEDDETTEMP!PreUnid)) / XVALO(PEDDETTEMP!PreList)
           End If
           If Abs(PDESCU) < 0.05 Then PDESCU = 0
         Call REPLA(X$, MKS$(PDESCU), 69, 4)
           '
           ' SI NO HAY DESCUENTOS PREDEFINIDOS, LO APLICA AL PRECIO
           If CantRegistros("DESCOMBI") < 1 Then
              PREL = PREL * (100 - PDESCU) / 100
              PDESCU = 0
              Call REPLA(X$, MKS$(PREL), 39, 4)
              Call REPLA(X$, MKS$(PDESCU), 69, 4)
           End If
           '
           CAN = XVALO(PEDDETTEMP!CanPed)
         Call REPLA(X$, MKS$(CAN), 47, 4)
           FSE% = CVINT(PEDDETTEMP!FeSolEnt)
         Call REPLA(X$, MKI$(FSE%), 63, 2) ' FECHA SOLICITADA DE ENTREGA
           'CAENTAN = PEDDETTEMP!CantEnt
           'FULENTAN% = CVINT(PEDDETTEMP!FeUltEnt)
           NLISSTA = XVALO(PEDDETTEMP!lista_de_item)
           Call REPLA(X$, MKS$(NLISSTA), 85, 4)
           
           IDIT$ = SAFETEXT$(PEDDETTEMP!TPedItem)
         Call REPLA(X$, IDIT$, 93, 16)
         Call REPLA(X$, MKS$(0), 117, 4)   ' ESTO SE ACTUALIZA MAS ABAJO
         Call REPLA(X$, MKI$(0), 121, 2)
         
         'esto agregado por ahp
         NPROV% = XVALO(PEDDETTEMP!Num_Prov)
         Call REPLA(X$, MKI$(NPROV%), 123, 2)
         'fin agregado ahp
         
           SUCOD$ = ""
           If IsNull(PEDDETTEMP!SubCod) = False Then SUCOD$ = PEDDETTEMP!SubCod
         Call REPLA(X$, SUCOD$, 109, 8)
         JJ& = JJ& + 1
         Call GRAREG("!CARDETPE", X$, JJ&)
       PEDDETTEMP.MoveNext
       Loop
'     End If
     PEDDETTEMP.Close
     Set PEDDETTEMP = Nothing
     
     Call SETULTREG("!CARDETPE", JJ&)
     Call CIERRARCH("CARDETPE")
     On Error GoTo 0
     '
     '
'     Erase CANENTAN()

'     Call ABRESTOCKS
'     SQLX$ = "SELECT * FROM MOVISTO with(nolock)"
'     SQLX$ = SQLX$ + " WHERE CodiMovi = '" & CMO$ & "' "
'     SQLX$ = SQLX$ + " AND NuPedCli = " & NUPEDIDO& & " "
     'Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     '
     'jandy sql
'     Dim MOVITMP As ADODB.Recordset
'     Set PEDDETTEMP = New ADODB.Recordset
'     PEDDETTEMP.CursorLocation = adUseClient
'28   On Error Resume Next
'     Set PEDDETTEMP = READSET(SQLX$)
'     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'        On Error GoTo 0
'        Call CapturaErrorOpen
'        GoTo 28
'     End If
'     On Error GoTo 0
     
     'jandy esto asi estaba
'     On Error Resume Next
       
    MODPED07.MARCONOTA.Visible = True
    On Error Resume Next
    MODPED07.Command6.SetFocus
    SendKeys "{ENTER}"
    MODPED07.Label25 = "" 'QUITO LOS DESCUENTOS
    MODPED07.PORDESCU = ""
    Call T_Espera(1)
    SendKeys "{ESC}"
    SendKeys "{ENTER}"
    Call T_Espera(1)
    Call MODPED07.PRIPEDIDO(1) 'IMPRIMO
    Call SETULTREG("!CARDETPE", 0) 'VUELO EL CARDETPE
    On Error GoTo 0
    Screen.MousePointer = 1

End Sub


Sub PRI_REMI_TEC_EXT(NCLI, NroRep&, RET_NROREMIT$)
    PRINTERPORT$ = TRIM$(CONTROL$(IDENTER$, "PORTREM"))
    
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
    'ASIGNO EL FORMULARIO CONFIGURADO - SI NO TIENE CONFIGURADO FRM LE VUELVO A ASIGNAR EL GENERAL
    If TRIM$(FORIPRE$) <> "" Then
      '
      PREFIREM$ = TRIM$(CONTROL$("", "PREFIREM"))
      If PREFIREM$ <> "*" Then
          NUPUVEN$ = PREFIREM$
      End If
      '
      Call NUREMIFA(NROPRO&, NUPUVEN$)
      NROREMIT$ = "RT" & "-X-" & FORMATOCON0$(XVALO(NUPUVEN$), 4) & "-" & FORMATOCON0$(NROPRO&, 8)
      RET_NROREMIT$ = NROREMIT$
      '
'    '
'    'GRABO EL MOVIMIENTO DE STOCK
    REMI$ = TRIM$(NROREMIT$)
    NROO& = Val(Mid$(REMI$, 11))
    REMICOR1$ = "RT." & SAFETEXT$(NROO&)
    FF% = HOY(HO$)
    DEPO% = XVALO(CONTROL$("", "DPREPPAÑ"))
    REFEMO$ = rasocli$(NCLI)

    For i& = 1 To ULTREG&("!VALEMAT")
        NUORIT% = i&
        YY$ = REGBLAN$("VALEMAT")
        YY$ = REGLEIDO$("!VALEMAT", i&)
        CI% = CVI(Mid$(YY$, 1, 2))
        CAN# = Str$(CVS(Mid$(YY$, 19, 4)))
        If CAN# > 0 Then
          LOTE$ = TRIM$(Mid$(YY$, 49, 16))
          'VUNIST# = CVS(Mid$(YY$, 159, 4))
          Call MOVISTOK(FF%, CI%, LOTE$, "RT", REMI$, REMICOR1$, "", "", REFEMO$, NUORIT%, VUNIST#, 1, NCLI, DEPO%, -1 * CAN#, 0, 0, "", "", "", 0, 0, , 0)
        End If
    Next i&
    Call ACTUREGISTRO("REPARA", "STATUS", "NUMREPA = " & NroRep&, 5, REGAF&, "NOMESS")
      
      
      Call LIMPIATETAB
      
      FEX = FF%
      FECHDOC$ = FECHATEX$(FEX)
      NUMEDOC$ = NROREMIT$
      TIPODOC$ = "Remito"
      '
      CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = "" ' TRIM$(Text8.TEXT)
      CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = "" 'TRIM$(Text15.TEXT)
      CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = "" 'TRIM$(Text12.TEXT)
      CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = "" 'TRIM$(Text17.TEXT)
      CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = "" 'TRIM$(Text14.TEXT)
      CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = "" 'TRIM$(Text13.TEXT)
      CODPARAM$(7) = "DEPOSAL ": DOCPARAM$(7) = "" 'TRIM$(DEPOSAL$)
      'CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = TRIM$(NROREMIT$)
      CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = Mid$(NROREMIT$, 6) ' SE IMPRIME FORMATO 0000-00000000
      CODPARAM$(9) = "NRO-OCOM": DOCPARAM$(9) = "" 'TRIM$(Text20.TEXT)
      CODPARAM$(10) = "PATE-TRA": DOCPARAM$(10) = "" 'TRIM$(Text18.TEXT)
      CODPARAM$(11) = "N-AUTENT": DOCPARAM$(11) = "" 'TRIM$(NORVENTA)
      CODPARAM$(12) = "COND-TRA": DOCPARAM$(12) = "" 'TRIM$(Text19)
      CODPARAM$(13) = "LOCA-TRA": DOCPARAM$(13) = "" 'TRIM$(Text21)
      CODPARAM$(14) = "NRO-AENT": DOCPARAM$(14) = "" 'TRIM$(Text22)

      '\\\OT-07-0007-RG-08/01/07///
      HORA$ = HORAENT$(NCLI)
      CODPARAM$(15) = "Z-PLANO": DOCPARAM$(15) = "" 'TRIM$(HORA$)
      CODPARAM$(16) = "IMP-TOTA": DOCPARAM$(16) = "" 'TRIM$(FORMATNUM$(ACUMLAIMPOX#, "F14.2"))
      CODPARAM$(17) = "DOCU-APL": DOCPARAM$(17) = "" 'TRIM$(STRNPEDI$)
      CODPARAM$(18) = "DESTINO": DOCPARAM$(18) = TRIM$(rasocli$(NCLI)) + " - Almacén " + Label11
      
      '\\\OT-07-0007-RG-FIN///
      CAPARDOC% = 18
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "CANTIDAD"
      CODTABU1$(4) = "NRO-LOTE"
      CODTABU1$(5) = "UNIMED"
      CODTABU1$(6) = "N-OCOMPR"
      CODTABU1$(7) = "ORD-ITEM"
      CODTABU1$(8) = "REF-MAT1"
      CODTABU1$(9) = "REF-MAT2"
      CODTABU1$(10) = "PRE-NETO"
      CODTABU1$(11) = "NET-ITEM"
      CODTABU1$(12) = "REF-MAT6"
      CODTABU1$(13) = "REF-MAT5"
      CACOLTAB1% = 13
      '
      DESTIDOCNUE = NCLI
      CAITAB1% = 0
      '



    For YX& = 1 To ULTREG&("!VALEMAT")
        YY$ = REGLEIDO$("!VALEMAT", YX&)
        CI% = CVI(Mid$(YY$, 1, 2))
        CAN# = Str$(CVS(Mid$(YY$, 19, 4)))
        LOTE$ = TRIM$(Mid$(YY$, 49, 16))
'        NPEDI& = CVS(Mid$(YY$, 151, 4))
        CAC# = 0
        If CI% > 0 And CAN# > 0 Then
          
              '
              UNID$ = Unimed$(CI%)
              If UNID$ = "" Then UNID$ = Unimed$(CI%)
              CANS$ = CSTRING$(MKD$(CAN#), 3, 12, 1, 2)
              If CAC# >= 0.05 Then
                CANS$ = CSTRING$(MKD$(CAC#), 3, 12, 1, 2)
                Call REPLA(CANS$, "(*)", 1, 3)
              End If
              '
              '
              CAITAB1% = CAITAB1% + 1
              CODDX$ = CODEXT$(CI%)
              If TRIM$(Mid$(YY$, 97, 8)) <> "" Then 'VALIDACION POR QUE SI NO QUEDA EL GUION SOLO
                CODDX$ = CODEXT$(CI%) + "-" + TRIM$(Mid$(YY$, 97, 8)) ' agrega subcodigo
              End If
              If REPLAORI% = 1 Then
                CODDX$ = CODORIG$(CI%, NCLI)
              End If
              '
              TETABU1$(1, CAITAB1%) = CODDX$
              TETABU1$(2, CAITAB1%) = DESCRIT0$(CI%)
              If Frame15.Visible = True Then
                TETABU1$(2, CAITAB1%) = Mid$(YY$, 5, 64)
              End If
              '
              TETABU1$(3, CAITAB1%) = CANS$
              TETABU1$(4, CAITAB1%) = NULOTE$
              TETABU1$(5, CAITAB1%) = UNID$
              TETABU1$(6, CAITAB1%) = NUOCO$
              TETABU1$(7, CAITAB1%) = Str$(CAITAB1%)
              TETABU1$(8, CAITAB1%) = NUOCOCLI$(Val(NUOCO$), CI%)
              TETABU1$(9, CAITAB1%) = NUPLANO$(CI%) '
              '
              CORIG$ = TRIM$(CODORIG$(CI%, NCLI))
              'If CORIG$ = TRIM$(CODDX$) Then CORIG$ = ""
              If CORIG$ = CODEXT$(CI%) Then CORIG$ = ""
                'If CORIG$ <> CODEXT$(CI%) Then CORIG$ = CODEXT$(CI%) 'ESTA LINEA HABRIA QUE QUITAR
              TETABU1$(12, CAITAB1%) = CORIG$
              '
              TETABU1$(13, CAITAB1%) = SAFETEXT$(ORD&) ' NUMERO DE ORDEN INGRESADO EN VENTANA DE ARCON NEUQUEN.
              '
              Call CARLINSEP         'Carga linea de separacion entre codigo y codigo
        End If
      Next YX&
      '
      'presenta frm de impresion fantasma.
      MODOIMPRE07.Check7.Caption = "Imprime Remito"
      MODOIMPRE07.Label1.Caption = "" 'LIMPIA EL LABEL DONDE PONE OK, SI ACEPTA
      MODOIMPRE07.Check8.Value = 1
      Screen.MousePointer = 1
      CANCELPRINT% = 0
      MODOIMPRE07.Show 1
      If MODOIMPRE07.Label1.Caption = "OK" Then
         If MODOIMPRE07.Check8.Value > 0 Then CANCELPRINT% = 1
      End If
      '
      Call PRINTDOC("@" + FORIPRE$)   ' Remito de Facturacion

      '
    End If
    '
End Sub

Sub PRI_REMI_TEC_EXTERNO(NCLI, REMI$)
    PRINTERPORT$ = TRIM$(CONTROL$(IDENTER$, "PORTREM"))
    
    FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
    'ASIGNO EL FORMULARIO CONFIGURADO - SI NO TIENE CONFIGURADO FRM LE VUELVO A ASIGNAR EL GENERAL
    If TRIM$(FORIPRE$) <> "" Then
      '
      PREFIREM$ = TRIM$(CONTROL$("", "PREFIREM"))
      If PREFIREM$ <> "*" Then
          NUPUVEN$ = PREFIREM$
      End If
      '
      REFEMO$ = rasocli$(NCLI)
      Call LIMPIATETAB
      FF% = HOY(HO$)
      FEX = FF%
      FECHDOC$ = FECHATEX$(FEX)
      NUMEDOC$ = REMI$
      TIPODOC$ = "Rem-Repa"
      '
      CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = "" ' TRIM$(Text8.TEXT)
      CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = "" 'TRIM$(Text15.TEXT)
      CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = "" 'TRIM$(Text12.TEXT)
      CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = "" 'TRIM$(Text17.TEXT)
      CODPARAM$(5) = "DOMI-TRA": DOCPARAM$(5) = "" 'TRIM$(Text14.TEXT)
      CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = "" 'TRIM$(Text13.TEXT)
      CODPARAM$(7) = "DEPOSAL ": DOCPARAM$(7) = "" 'TRIM$(DEPOSAL$)
      'CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = TRIM$(NROREMIT$)
      CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = Mid$(NROREMIT$, 6) ' SE IMPRIME FORMATO 0000-00000000
      CODPARAM$(9) = "NRO-OCOM": DOCPARAM$(9) = "" 'TRIM$(Text20.TEXT)
      CODPARAM$(10) = "PATE-TRA": DOCPARAM$(10) = "" 'TRIM$(Text18.TEXT)
      CODPARAM$(11) = "N-AUTENT": DOCPARAM$(11) = "" 'TRIM$(NORVENTA)
      CODPARAM$(12) = "COND-TRA": DOCPARAM$(12) = "" 'TRIM$(Text19)
      CODPARAM$(13) = "LOCA-TRA": DOCPARAM$(13) = "" 'TRIM$(Text21)
      CODPARAM$(14) = "NRO-AENT": DOCPARAM$(14) = "" 'TRIM$(Text22)

      '\\\OT-07-0007-RG-08/01/07///
      HORA$ = HORAENT$(NCLI)
      CODPARAM$(15) = "Z-PLANO": DOCPARAM$(15) = "" 'TRIM$(HORA$)
      CODPARAM$(16) = "IMP-TOTA": DOCPARAM$(16) = "" 'TRIM$(FORMATNUM$(ACUMLAIMPOX#, "F14.2"))
      CODPARAM$(17) = "DOCU-APL": DOCPARAM$(17) = "" 'TRIM$(STRNPEDI$)
      CODPARAM$(18) = "DESTINO": DOCPARAM$(18) = TRIM$(rasocli$(NCLI)) + " - Almacén " + Label11
      
      '\\\OT-07-0007-RG-FIN///
      CAPARDOC% = 18
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "CANTIDAD"
      CODTABU1$(4) = "NRO-LOTE"
      CODTABU1$(5) = "UNIMED"
      CODTABU1$(6) = "N-OCOMPR"
      CODTABU1$(7) = "ORD-ITEM"
      CODTABU1$(8) = "REF-MAT1"
      CODTABU1$(9) = "REF-MAT2"
      CODTABU1$(10) = "PRE-NETO"
      CODTABU1$(11) = "NET-ITEM"
      CODTABU1$(12) = "REF-MAT6"
      CODTABU1$(13) = "REF-MAT5"
      CACOLTAB1% = 13
      '
      DESTIDOCNUE = NCLI
      CAITAB1% = 0
      '
      CMO$ = "RP"
      CONDI$ = "(CodiMovi = '" & CMO$ & "' OR CodiMovi = 'DV') "
      CONDI$ = CONDI$ + " AND DoPriLAR = '" & REMI$ & "' AND CANTSALID > 0 AND CODIMOVI= 'RP'"
      SQLX$ = "SELECT COD_INTE,NuorItem,DESCRI_LAR, CantSalid, cantingre, NUPEDCLI, ValoUnitIva, Nume_Clie, FECHMOV,IDENLOTE FROM MOVISTO WITH(NOLOCK)"
      SQLX$ = SQLX$ + " WHERE " & CONDI$ & "  ORDER BY NuOrItem"
      Set rs = READSET(SQLX$)
      With rs
          Do While Not .EOF
                CI% = XVALO(!COD_INTE)
                CAN# = XVALO(!CANTSALID)
                LOTE$ = SAFETEXT$(!idenlote)
                CAC# = 0
                If CI% > 0 And CAN# > 0 Then
                      UNID$ = Unimed$(CI%)
                      If UNID$ = "" Then UNID$ = Unimed$(CI%)
                      CANS$ = CSTRING$(MKD$(CAN#), 3, 12, 1, 2)
                      If CAC# >= 0.05 Then
                        CANS$ = CSTRING$(MKD$(CAC#), 3, 12, 1, 2)
                        Call REPLA(CANS$, "(*)", 1, 3)
                      End If
                      '
                      '
                      CAITAB1% = CAITAB1% + 1
                      CODDX$ = CODEXT$(CI%)
                      If TRIM$(Mid$(YY$, 97, 8)) <> "" Then 'VALIDACION POR QUE SI NO QUEDA EL GUION SOLO
                        CODDX$ = CODEXT$(CI%) + "-" + TRIM$(Mid$(YY$, 97, 8)) ' agrega subcodigo
                      End If
                      If REPLAORI% = 1 Then
                        CODDX$ = CODORIG$(CI%, NCLI)
                      End If
                      '
                      TETABU1$(1, CAITAB1%) = CODDX$
                      TETABU1$(2, CAITAB1%) = DESCRIT0$(CI%)
'                      If Frame15.Visible = True Then
'                        TETABU1$(2, CAITAB1%) = Mid$(YY$, 5, 64)
'                      End If
                      '
                      TETABU1$(3, CAITAB1%) = CANS$
                      TETABU1$(4, CAITAB1%) = NULOTE$
                      TETABU1$(5, CAITAB1%) = UNID$
                      TETABU1$(6, CAITAB1%) = NUOCO$
                      TETABU1$(7, CAITAB1%) = Str$(CAITAB1%)
                      TETABU1$(8, CAITAB1%) = NUOCOCLI$(Val(NUOCO$), CI%)
                      TETABU1$(9, CAITAB1%) = NUPLANO$(CI%) '
                      '
                      CORIG$ = TRIM$(CODORIG$(CI%, NCLI))
                      'If CORIG$ = TRIM$(CODDX$) Then CORIG$ = ""
                      If CORIG$ = CODEXT$(CI%) Then CORIG$ = ""
                        'If CORIG$ <> CODEXT$(CI%) Then CORIG$ = CODEXT$(CI%) 'ESTA LINEA HABRIA QUE QUITAR
                      TETABU1$(12, CAITAB1%) = CORIG$
                      '
                      TETABU1$(13, CAITAB1%) = SAFETEXT$(ORD&) ' NUMERO DE ORDEN INGRESADO EN VENTANA DE ARCON NEUQUEN.
                      '
                      Call CARLINSEP         'Carga linea de separacion entre codigo y codigo
                End If
                .MoveNext
          Loop
      End With
              '
      'presenta frm de impresion fantasma.
      MODOIMPRE07.Check7.Caption = "Imprime Remito"
      MODOIMPRE07.Label1.Caption = "" 'LIMPIA EL LABEL DONDE PONE OK, SI ACEPTA
      MODOIMPRE07.Check8.Value = 1
      Screen.MousePointer = 1
      CANCELPRINT% = 0
      MODOIMPRE07.Show 1
      If MODOIMPRE07.Label1.Caption = "OK" Then
       If MODOIMPRE07.Check8.Value > 0 Then CANCELPRINT% = 1
      End If
      '
      Call PRINTDOC("@" + FORIPRE$)   ' Remito de Facturacion
        
      '
    End If
    '
End Sub



Sub GENERA_MOV_CBLE(NUREP&)
   'GENERA EL ASIENTO DE REPARACION Y MAMO DE OBRA DE LA REPARACION
   'CONTROLES DE CONFIGURACION DE CUENTAS CONTABLES
   CTACONTRA$ = TRIM$(CONTROL("REPARA", "CTCONREP")) 'cuenta configurada en setup de reparaciones
   CTAREP$ = TRIM$(CONTROL("REPARA", "CTBREPUE")) 'cuenta configurada en setup de reparaciones
   
   CTACONTRAPARTIDA% = CUECOINT%(CTACONTRA$)
   CTAREPU% = CUECOINT%(CTAREP$)
   '
   SUMREPUX = SUMREPU(NUREP&)  'REPUESTOS
   'MOBRAX = SUMMOBRA(NUREP&)   'MANO DE OBRA en GARANTIA
   TOTREPYMOBRA = SUMREPUX  'EN EL CASO QUE SEA CON TECNICO DEBERIA SOLO TENER 1
   If TOTREPYMOBRA <= 0 Then Exit Sub
   
   FEHOY% = HOY(HO$)
   
   Call CARDATEJ
   NUASI$ = NUPROA1$(HOY(HO$))
   '
   OBS$ = "ASIENTO AUTOMÁTICO DE ORDEN DE REPARACIÓN N°: " & NUREP&
   
   On Error GoTo ERROR_GUARDAR
   FLEXCONN.BeginTrans

   Dim EncaSiGra As ADODB.Recordset
   Set EncaSiGra = New ADODB.Recordset
   strsql = "Select * from EncaBeAsien where NuIDAsien = '" & NUASI$ & "'"
   EncaSiGra.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With EncaSiGra
         .AddNew   ' YA CONVERTIDO
         !StatAsien = (-1)
          On Error Resume Next
          !CodiEmpr = CodiEmp%
          On Error GoTo 0
          !NUIDASIEN = NUASI$
          If Text4 = "" Then Text4 = " "
          !Refeasien = "ASIENTO POR REPARACIÓN N°: " & NUREP& & " (EN GARANTIA)"
             FECHAU% = HOY(HO$)
          !FECHASIEN = CVDAT(FECHAU%)
             PPXX% = InStr(NUASI$, "-")
          !OrFeAsien = Val(Mid$(NUASI$, PPXX% + 1))
          !SucuAsien = UNINEGO%
          !NuDfAsien = 0
          !TIPOASIEN = 11
          !StatAsien = 1  ' activo
          !EditAsien = -1
          !QPasAsien = 3
          !DebeAsien = TOTREPYMOBRA
          !Observa = OBS$
          !Nume_Usuar = USNBR%
          !Niv_Clasi = 0
          !FeReAsien = Now
          .Update
   End With
   
   NICLA% = 0
   ORPASE% = 0
   '
   Set DetasiGra = New ADODB.Recordset
   strsql = "Select * from DetaAsien where NuIDAsien = '" & NUASI$ & "'"
   DetasiGra.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   With DetasiGra
     For K2% = 1 To 2
          Select Case K2%
            Case 1
              .AddNew   ' YA CONVERTIDO
              On Error Resume Next ' ASIENTO DE TOTAL CONTRAPARTIDA DE CUENTA CONTABLE DE REPARACION
              !CodiEmpr = CodiEmp%
              On Error GoTo 0
              !NUIDASIEN = NUASI$
              ORPASE% = ORPASE% + 1
              !OrNuPase = ORPASE%
              !RefePase = Left$(DENOCUE$(CTACONTRAPARTIDA%), 24)
              FECHAU% = FEHOY%
              !FECHASIEN = CVDAT(FECHAU%)
              AA = CVDAT(FECHAU%)
              PPXX% = InStr(NUASI$, "-")
              !OrFeAsien = Val(Mid$(NUASI$, PPXX% + 1))
              !SucuAsien = 0
              !NuDfAsien = 0
              !StatPase = 1
              !CUECOEXT = CTACONTRA$
              !CUECOINT = CTACONTRAPARTIDA%
              !IDebePase = 0
              !IHabePase = TOTREPYMOBRA
              !MARCONSI = ""
              !CenCos = ""
              !Niv_Clasi = 0
              !DATRETPER = ""
             .Update
             
            Case 2
              If SUMREPUX <= 0 Then GoTo 80 ' NEXT
              .AddNew   ' YA CONVERTIDO
              On Error Resume Next 'ASIENTO DE REPUESTOS
              !CodiEmpr = CodiEmp%
              On Error GoTo 0
              !NUIDASIEN = NUASI$
              ORPASE% = ORPASE% + 1
              !OrNuPase = ORPASE%
              !RefePase = Left$(DENOCUE$(CTAREPU%), 24)
              FECHAU% = FEHOY%
              !FECHASIEN = CVDAT(FECHAU%)
              PPXX% = InStr(NUASI$, "-")
              !OrFeAsien = Val(Mid$(NUASI$, PPXX% + 1))
              !SucuAsien = 0
              !NuDfAsien = 0
              !StatPase = 1
              !CUECOEXT = CTAREP$
              !CUECOINT = CTAREPU%
              !IDebePase = SUMREPUX
              !IHabePase = 0
              !MARCONSI = ""
              !CenCos = ""
              !Niv_Clasi = 0
             !DATRETPER = ""
             .Update
             
'            Case 3
'              If MOBRAX <= 0 Then GoTo 80 ' NEXT
'              .AddNew   ' YA CONVERTIDO
'              On Error Resume Next 'ASIENTO DE MANO DE OBRA
'              !Codiempr = CodiEmp%
'              On Error GoTo 0
'              !NUIDASIEN = NUASI$
'              ORPASE% = ORPASE% + 1
'              !OrNuPase = ORPASE%
'              !RefePase = Left$(DENOCUE$(CTAMOBRA%), 24)
'              FECHAU% = FEHOY%
'              !FECHASIEN = CVDAT(FECHAU%)
'              PPXX% = InStr(NUASI$, "-")
'              !OrFeAsien = Val(Mid$(NUASI$, PPXX% + 1))
'              !SucuAsien = 0
'              !NuDfAsien = 0
'              !StatPase = 1
'              !CUECOEXT = CTAMOB$
'              !CUECOINT = CTAMOBRA%
'              !IDebePase = 0
'              !IHabePase = MOBRAX
'              !MARCONSI = ""
'              !CenCos = ""
'              !Niv_Clasi = 0
'             !DATRETPER = ""
'             .Update
          End Select
          '
80        Next K2%
          DetasiGra.Close
          Set DetasiGra = Nothing
   End With
   '
   'GRABO EL NUMERO DE ASIENTO EN REPARA
   CONDI$ = "NUMREPA = " & NUREP&
   Call ACTUREGISTRO("REPARA", "NUIDASIEN", CONDI$, NUASI$, REGAF&, "NOMESS")
   '
ERROR_GUARDAR:
    If Err <> 0 Then
       FLEXCONN.RollbackTrans
       On Error Resume Next
       DetasiGra.Close
       Set DetasiGra = Nothing
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       NUSERI$ = VALOBADA("REPARA", "NUMEROSERIE", "NUMREPA = " & NUREP&, "NOMESS")
       CUERPO$ = "No se Generó El Asiento Automático " & vbCrLf
       CUERPO$ = "Reparación N°:  " & NUREP& & vbCrLf
       Call ANOTANUSE(NUSERI$, CUERPO$, , "NOMESS")
       Exit Sub
    Else
       FLEXCONN.CommitTrans
    End If
90
End Sub

Private Sub Command11_Click()
    
   Command11.Enabled = False
   
10 LOTE$ = LOTSELEC$(CI%)
   '
   If CI% < 1 And TRIM(LOTE$) = "" Then
       GoTo 99
    End If

   VDEF$ = String$(64, 0)
   DEP% = DEPOLOT%(CI%, LOTE$)
   Call REPLA(VDEF$, MKI$(CI%), 1, 2)
   Call REPLA(VDEF$, LOTE$, 3, 12)
   Call REPLA(VDEF$, MKI$(DEP%), 20, 1)
   '
20 OBX$ = OBJPLA$("COLODECA", VDEF$)
   '
   If OBX$ = "" Then GoTo 99
   '
   DEPOCONSUM% = Asc(Mid$(OBX$, 20, 1))
   CANTI = CVS(Mid$(OBX$, 15, 4))
   '
   If CI% < 1 Then
      Call COMUNI("Falta Ingresar Código")
      GoTo 10
   End If
   '
   If TRIM$(LOTE$) = "" And ESTRAZABLE%(CI%) = 1 Then
      Call COMUNI("Falta Ingresar Lote")
      GoTo 10
   End If
   '
   If XVALO(CANTI) < 0.05 Then
      Call COMUNI("Falta Ingresar Cantidad")
      VDEF$ = OBX$
      GoTo 20
   End If
   '
   If DEPOCONSUM% < 1 Or ECOARCH$("TADEPOSI", Chr$(DEPOCONSUM%)) = "" Then
     Call MENSERR(24, "Depósito de Consumo No Válido")
     VDEF$ = OBX$
     GoTo 20
   End If
   '
   If TRIM$(LOTE$) <> "" Then
    If DEPOCONSUM% <> DEP% Then
       Call COMUNI("Depósito Seleccionado Difiere de Ubicación de Lote")
       Call REPLA(OBX$, MKI$(DEP%), 20, 1)
       VDEF$ = OBX$
       GoTo 20
    End If
   End If
   '
   '
   If TRIM$(LOTE$) <> "" Then
     CONDI$ = "Cod_Inte =" & CI% & " AND IdenLote ='" & LOTE$ & "'"
     CANTX = XVALO(VALOBADA("LOTINGRE", "CanSaldo", CONDI$))
     If CANTX < XVALO(CANTI) Then
       Call COMUNI("Cantidad Solicitada Supera Saldo de Lote Seleccionado")
       GoTo 20
     End If
   End If
   '
   '
   NUREP& = XVALO(Label25)
   NUMEREPA$ = "OR-" & TRIM$(Str$(NUREP&))

   EPAC$ = TRIM$(EMPREAC$)
   If InStr(EPAC$, "EYM") > 0 Then
      CI1% = CI%
      Call CargarGrillaCierreEym(CI1%, CANTI, LOTE$)
      Command11.Enabled = True
      Exit Sub
   End If

   coefi = Val(CONTROL$("REPARA", "COEFPREA")) ' COEFICIENTE SETEADO EN SETUP
   If coefi < 0.05 Then coefi = 1
   PRECIUNIT# = COSUNI(CI%) * coefi
   '
   FECHA% = HOY(HO$)
   CMOV$ = "RP"
   NCOMP& = ProNroComprobante(CMOV$)
   DOPRI$ = CMOV$ & "." & NCOMP&
   DOSEC$ = "CR." & NCOMP&
   REFE$ = "Cons. de Repar. " & NUREP&

   Call MOVISTOK(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, PRECIUNIT#, MONEII%, NC%, DEPOCONSUM%, (-1) * CANTI)
   Call ACSALOTE(CI%, LOTE$, "NOMESS")
  
   Screen.MousePointer = 1
   Call COMUNI("Consumo Realizado")
99 Command11.Enabled = True
  
End Sub

Private Sub Command15_Click()
    '
    If DERACCE%("MODEPRE", "Modificación de Depósito de Reparaciones") > 1 Then
        Call SELECHO("TADEPOSI")
        If VALACT1$("TADEPOSI") <> "" Then
            Label13 = TRIM$(VALACT1$("TADEPOSI"))
            Label6 = ECOARCH$("TADEPOSI", Chr$(XVALO(Label13)))
        End If
    End If
    '
End Sub

Sub Command2_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Unload Me
 End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    CI% = CODINT(Label8)
    If CI% < 1 Then Call MENSERR(24, "Falta seleccionar Item"): GoTo 99

    NUVAL$ = TRIM$(UCase(Label4))
    If NUVAL = "" Then Call MENSERR(24, "Falta Ingresar el Lote")
    '
    TTXX$ = List1.List(List1.ListIndex)
    Call REPLA(TTXX$, AJUSTI$(NUVAL$, 12), 74, 12)
    List1.List(List1.ListIndex) = TTXX$
99  Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   If TRIM$(Label8) <> "" Then
    Call SELECHO("!LILOTSE")
    TTXX$ = List1.List(List1.ListIndex)
    If VALACT1$("!LILOTSE") <> "" Then
       VDEV2$ = VALACT2("!LILOTSE")
       CANTI = Mid$(TTXX$, 64, 8)
       CANTILOTE = FORMATNUM$(Mid$(VDEV2$, 21, 12), "F8.1")
       If XVALO(CANTI) > XVALO(CANTILOTE) Then
         Call MENSERR(24, "Cantidad Solicitada Supera\Saldo Total del Lote")
         Call LIMPIAR
         Exit Sub
       End If

       Label4 = VALACT1$("!LILOTSE")

    End If
   End If
End Sub


Private Sub Form_Load()
    Label13 = ""
    DEPOSI% = XVALO(CONTROL$("", "DEPREDES"))
    If DEPOSI% > 0 Then
      Label13 = DEPOSI%
    End If
    
    EPAC$ = TRIM$(EMPREAC$)
    MSF.Visible = False
    If InStr(EPAC$, "EYM") > 0 Then
          MSF.Visible = True
          MSF.Top = 705
          MSF.Height = 2100
          MSF.Left = 75
          MSF.Width = 10365
          Campos$ = "Código/A16;Descripción/A24;Cant/F10.1;Lote/A14;CI/A0"
          Call CARGA_ENCABEZADO(Me.MSF, Campos$, Me)
          Command6.Visible = True
    End If
    
End Sub

Private Sub Frame4_DblClick()
    Call Label9_DblClick
End Sub

Private Sub Label13_Change()
    Label6 = ECOARCH$("TADEPOSI", Chr$(XVALO(Label13)))
End Sub

Private Sub Label25_Change()
    '
    NUREP& = XVALO(Label25)
    NUMEREPA$ = "OR-" & TRIM$(Str$(NUREP&))
    '
    CONDI$ = "NUMREPA = " & NUREP& & ""
    Label22 = VALOBADA("REPARA", "DESCRIPCION", CONDI$, "NOMESS")
    Label26 = VALOBADA("REPARA", "NUMEROSERIE", CONDI$, "NOMESS")
    Label20 = VALOBADA("REPARA", "DEFEC_INFOR", CONDI$, "NOMESS")
    Label28 = VALOBADA("REPARA", "ACCESORIOS", CONDI$, "NOMESS")
    ESTEMHS# = XVALO(VALOBADA("REPARA", "Estima_Horas", CONDI$, "NOMESS"))
    Text2 = FORMATNUM$(ESTEMHS#, "F5.2")
    Text1 = VALOBADA("REPARA", "UsRealizaRepa", CONDI$, "NOMESS")
    
    'Text3 = VALOBADA("REPARA", "TRAB_EFECT", CONDI$, "NOMESS")
    REPAX$ = VALOBADA("REPARA", "REPARACION", CONDI$, "NOMESS")
    If TRIM$(REPAX$) = "" Then REPAX$ = VALOBADA("REPARA", "TRAB_EFECT", CONDI$, "NOMESS")
    Text3 = REPAX$
    
    Label2 = USERNAME$
    Label13 = ""
    DEPOSI% = XVALO(CONTROL$("", "DEPREDES"))
    If DEPOSI% > 0 Then
       Label13 = DEPOSI%
       Label6 = ECOARCH$("TADEPOSI", Chr$(XVALO(Label13)))
    End If
    Text5 = XVALO(VALOBADA("REPARA", "TEC_EXT", CONDI$, "NOMESS"))
    Label17.Caption = VALOBADA("REPARA", "TALON_TEC", CONDI$, "NOMESS")
    '
    List1.Clear
    ReparacionCerradaConItemsAgregados% = 0
    '
'    SEGUIR CARGAR A PARTIR DE ACA LOS CODIGOS DE LOS ACCEORIOS Y VER COMO AGREGAR EL LOTE Y QUE SE REALICE EL MOVIMIENTO
'    JUNTO CON LOS OTROS MOVIMIENTOS.
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "EYM") > 0 Then

        Campos$ = "Cod_Inte,CODIGO,DESCRIPCION,IdenLote,PrUn_Neto"
        SQLX$ = "SELECT " & Campos$
        SQLX$ = SQLX$ + " FROM ACCESORIOSCOTIZ WITH(NOLOCK) WHERE NUMREPA = " & XVALO(Label25)
        Set rs1 = READSET(SQLX$)
        With rs1
          Do While Not .EOF
             REG& = REG& + 1
             CIERREP09.MSF.Rows = REG& + 1
             CI1% = XVALO(!COD_INTE)
             CIERREP09.MSF.TextMatrix(REG&, 1) = SAFETEXT$(!Codigo)
             CIERREP09.MSF.TextMatrix(REG&, 2) = SAFETEXT$(!Descripcion)
             CIERREP09.MSF.TextMatrix(REG&, 3) = 1
             CIERREP09.MSF.TextMatrix(REG&, 4) = SAFETEXT$(!idenlote)
             CIERREP09.MSF.TextMatrix(REG&, 5) = CI1%
            .MoveNext
          Loop
        End With
        Exit Sub
    End If
    
    SQLX$ = "SELECT * FROM RESERVA  WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NUMEREPA$ & "'"
    SQLX$ = SQLX$ + " ORDER BY NUORITEM "
    Dim REPURTMP As ADODB.Recordset
    Set REPURTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With REPURTMP
      Do While Not .EOF
       CI% = XVALO(!COD_INTE)
       CANTI = XVALO(!CANTRES)
       LOTE$ = SAFETEXT$(!idenlote)
       
       Call REPLA(TTXX$, CODEXT$(CI%), 1, 16)
       Call REPLA(TTXX$, DESCRIT$(CI%), 17, 40)
       Call REPLA(TTXX$, FORMATNUM(CANTI, "F8.1"), 64, 8)
       Call REPLA(TTXX$, LOTE$, 74, 12)
       
       List1.AddItem TTXX$
       .MoveNext
      Loop
    End With
    REPURTMP.Close
    Set REPURTMP = Nothing
    
    
End Sub

Sub ActualizarRegArmado(NUSERIE$, CI1%, LOTE$, ValidarSerieImpo_o_Sistema%)
'    SerieSistema_o_Importado% ' 2 IMPORTADO, 1 SERIE DE SISTEMA
    
    If ValidarSerieImpo_o_Sistema% < 1 Then 'SI ES 0 NO VALIDO ANTERIORMENTE EN LA LLAMADA DE LA FUNCION
        If NumeroRegistroSegunSerieEym&(NUSERIE$) > 0 Then
           SerieSistema_o_Importado% = 2 'es serie importado
        Else
           SerieSistema_o_Importado% = 1  'es serie del sistema
        End If
     Else
        SerieSistema_o_Importado% = ValidarSerieImpo_o_Sistema% 'ES POR QUE VINO CON EL VALOR
     End If
     
'     ACA TOMAR LOS VALORES Y ASIGNAR LA TABLA CAMPO ....

     If SerieSistema_o_Importado% = 2 Then
           TABLA$ = "DETAREGARMA"
           CAMPO$ = "CODIGO"
           CampoLoteX$ = "IDENTIFICACION"
     Else
           TABLA$ = "TRAZAXSERIE"
           CAMPO$ = "COD_EXTE"
           CampoLoteX$ = "IDENLOTE"
     End If
     
     '
     'VALIDO POR LAS DUDAS SI EL CODIGO ESTA MAS DE UNA VEZ EN EL REG.ARM.
     CONDI$ = CAMPO$ & " = '" & CODEXT$(CI1%) & "'"
     CONDI$ = CONDI$ + " AND NUMEROSERIE = '" & NUSERIE$ & "'"
     If CantRegistros&(TABLA$, CONDI$, "NOMESS") > 0 Then
        'TOMA EL ID DE LA PRIMER OCURRENCIA PARA ACTUALIZAR SOLO UNO
        ID& = XVALO(VALOBADA(TABLA$, TABLA$ & "_ID", CONDI$, "NOMESS"))
        CONDI$ = TABLA$ & "_ID" & " = " & ID&
        LOTEANTERIOR$ = VALOBADA(TABLA$, CampoLoteX$, CONDI$, "NOMESS") 'LEO EL LOTE EXISTENTE
        Call ACTUREGISTRO(TABLA$, CampoLoteX$, CONDI$, LOTE$, REGAF&, "NOMESS") 'GRABO EL NUEVO LOTE
     End If
     '
     
     CUERPO$ = "Reemplazo Codigo: " & CODEXT$(Abs(CI1%)) & vbCrLf ' ABSOLUTO POR SI ESTA DADO DE BAJA
     CUERPO$ = CUERPO$ + "Lote Anterior:  " & LOTEANTERIOR$ & vbCrLf
     CUERPO$ = CUERPO$ + "Lote Nuevo:  " & LOTE$
     Call ANOTANUSE(NUSERIE$, CUERPO$, , "NOMESS")

End Sub

Sub CierreRepaEym(NroRep&, OK%)
   OK% = 0
   Dim CODICANT(32767)
   Dim LOTEX$(32767)
   'valido los lotes
   NUSERIE$ = Label26
   For REG& = 1 To MSF.Rows - 1
      CI1% = MSF.TextMatrix(REG&, 5)
      CANTI = XVALO(MSF.TextMatrix(REG&, 3))
      LOTE$ = MSF.TextMatrix(REG&, 4)
      
      If TRIM$(LOTE$) <> "" And ESTRAZABLE%(CI1%) = 1 Then
          CONDI$ = "Cod_Inte =" & CI1% & " AND IdenLote ='" & LOTE$ & "'"
          CANTX = XVALO(VALOBADA("LOTINGRE", "CanSaldo", CONDI$))
          If LOTEX$(CI1%) = LOTE$ Then CODICANT(CI1%) = CODICANT(CI1%) + CANTI 'SI SE REPITE UN LOTE LO ACUMULA
          If CANTX < CODICANT(CI1%) Then  'VALIDA LA CANTIDA DEL LOTE CONTRA LA CANTIDAD ACUMULADA
               Call COMUNI("Cantidad Acumulada del Lote: " & LOTE$ & " Supera el Saldo Actual")
               Call SELECCIONARFILA(Me.MSF, REG&)
               OK% = -1
               Exit Sub
          End If
           
          If CANTX < XVALO(CANTI) Then
              Call COMUNI("Cantidad Solicitada Supera Saldo de Lote Seleccionado")
              Call SELECCIONARFILA(Me.MSF, REG&)
              OK% = -1
              Exit Sub
           End If
       End If
       LOTEX$(CI1%) = LOTE$
       CODICANT(CI1%) = CANTI
    Next REG&
    '
    SerieSistema_o_Importado% = 0 ' por que si en la funcion de validacion de componentes si es 0 vuelve a consultar
    If NumeroRegistroSegunSerieEym&(NUSERIE$) > 0 Then
       SerieSistema_o_Importado% = 2 'es serie importado
    Else
       SerieSistema_o_Importado% = 1  'es serie del sistema
    End If
    'VALDO SI EL COMPONENTE EXISTE EN REGISTRO DE ARMADO
    If ValidacionExisteComponenteCriticoEnRegistroDeArmado%(NUSERIE$, MSF, 5, 4, SerieSistema_o_Importado%) < 1 Then
       Exit Sub
    End If
    
    'CONSUMO
    FECHA% = HOY(HO$)
    CMOV$ = "RP"
    REFE$ = "Cons. de Repar. " & NroRep&
    DEPOCONSUM% = XVALO(Me.Label13)
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans

    For REG& = 1 To MSF.Rows - 1
      CI1% = MSF.TextMatrix(REG&, 5)
      CANTI = XVALO(MSF.TextMatrix(REG&, 3))
      LOTE$ = MSF.TextMatrix(REG&, 4)
      NCOMP& = ProNroComprobante(CMOV$)
      DOPRI$ = CMOV$ & "." & NCOMP&
      DOSEC$ = "CR." & NCOMP&
      Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC%, DEPOCONSUM%, (-1) * CANTI)
      Call ACSALOTE(CI1%, LOTE$, "NOMESS")
      Call ActualizarRegArmado(NUSERIE$, CI1%, LOTE$, SerieSistema_o_Importado%)
     Next REG&

'    GRABO EL ESTADO EN LA REPARACION COMO CERRADA
    '
60   SQLX$ = "SELECT Trab_Efect, REPARACION, STATUS, UsRealizaRepa, FechaCierreRepa FROM REPARA"
     SQLX$ = SQLX$ + " WHERE NUMREPA = " & NroRep&
    
     Dim rs As ADODB.Recordset
     Set rs = New ADODB.Recordset
     rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
     With rs
      Do While Not .EOF
        !REPARACION = Text3
        !Status = 4
        !UsRealizaRepa = Left$(TRIM$(Text1), 32)
        !FechaCierreRepa = CVDAT(HOY(HO$))
        .Update
        .MoveNext
      Loop
     End With
     rs.Close
     Set rs = Nothing
    '
     CUERPO$ = "Se Generó Cierre de Reparación " & vbCrLf
     CUERPO$ = CUERPO$ + "Reparación N°:  " & NroRep& & vbCrLf
     Call ANOTANUSE(NUSERIE$, CUERPO$, , "NOMESS")

ERROR_GUARDAR:
    If Err <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       Exit Sub
    Else
       FLEXCONN.CommitTrans
    End If

End Sub

Function ValidacionExisteComponenteCriticoEnRegistroDeArmado%(NUSERIE$, MSF As MSFlexGrid, ColCodint%, ColLote%, Optional ValidarSerieImpo_o_Sistema%)
     'viene como parametro el nro de columna donde esta el codint y el lote
     If ValidarSerieImpo_o_Sistema% < 1 Then 'SI ES 0 NO VALIDO ANTERIORMENTE EN LA LLAMADA DE LA FUNCION
        If NumeroRegistroSegunSerieEym&(NUSERIE$) > 0 Then
           SerieSistema_o_Importado% = 2 'es serie importado
       Else
           SerieSistema_o_Importado% = 1  'es serie del sistema
        End If
     Else
        SerieSistema_o_Importado% = ValidarSerieImpo_o_Sistema% 'ES POR QUE VINO CON EL VALOR
     End If
     
     If SerieSistema_o_Importado% = 2 Then
       TABLA$ = "DETAREGARMA"
       CAMPO$ = "CODIGO"
     Else
       TABLA$ = "TRAZAXSERIE"
       CAMPO$ = "COD_EXTE"
     End If
     
     ValidacionExisteComponenteCriticoEnRegistroDeArmado% = 1
     For REG& = 1 To MSF.Rows - 1
      CI1% = MSF.TextMatrix(REG&, ColCodint%)
      LOTE$ = MSF.TextMatrix(REG&, ColLote%)
      Codigo$ = CODEXT$(CI1%)
'      If SerieSistema_o_Importado% = 1 Then '
      CONDI$ = CAMPO$ & " = '" & Codigo$ & "'"
'      End If
      If CantRegistros&(TABLA$, CONDI$, "NOMESS") < 1 Then
         OPX% = COMALTER%("Componente " & Codigo$ & " Inexistente en Registro de Armado\\Cancelar\Continuar")
         If OPX% < 2 Then 'SI CANCELA INTERRUME EL FOR
            Exit For      'SI ELIGE CONTINUAR SIGUE RECORRIENDO Y PRESENTA MENSAJE SI CORRESPONDE
         End If           'EN LOS DOS CASOS DEVUELVE - 1
         ValidacionExisteComponenteCriticoEnRegistroDeArmado% = -1
      End If
     Next REG&
End Function


Private Sub Command6_Click()
   NSERIE$ = Label26
   CONDI$ = "NUMEROSERIE = '" & NSERIE$ & " '"
   NORFA$ = VALOBADA("TANUMSE", "IDSUBOR", CONDI$, "NOMESS")
   FRMREGARMADO.Label13(2) = NSERIE$
   FRMREGARMADO.Show 1
   
End Sub

Private Sub MSF_DblClick()
      'SI VA AL BOTON DE INFORMAR CONSUMOS AGREGA UNA FILA SIEMPRE
      'EN ESTE CASO LEE EL CODIGO Y COMPLETA EL FORMULARIO DE ASIGNACION DE LOTES CON EL CODIGO
      'SI ELIGE LOTE PARA ESTE CODIGO AGREGA UNA FILA AL FINAL DEL MSF
      'Y BORRA LA FILA EN LA CUAL SE HIZO DOBLE CLIC
      
      REG& = MSF.MouseRow
      CantidadRegistros% = MSF.Rows - 1
      If MSF.Rows > 1 Then
        If REG& >= 1 Then
            Call SELECCIONARFILA(Me.MSF, REG&)
            SELCODLOT07.Text1 = MSF.TextMatrix(REG&, 1)
            Command11_Click 'llama a boton de seleccion de lotes
            If CantidadRegistros% < MSF.Rows - 1 Then 'VALIDO SI SE AGREGO UNA FILA ES QUE ACEPTO, SI NO AGREGO SALIO SIN ACEPTAR
              Call cmdEliminaItem(MSF, REG&) 'ELIMINO EL CODIGO EN EL CUAL SE HIZO DOBLE CLIC
            End If
        End If
      End If
      
End Sub

Private Sub MSF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
      REG& = MSF.MouseRow
      If MSF.Rows > 1 Then
        If REG& >= 1 Then
            Call SELECCIONARFILA(Me.MSF, REG&)
            OPX% = COMALTER("Realmente Desea Eliminar el Registro Seleccionado\\Cancelar\Sí, Eliminar")
            If OPX% = 2 Then
              Call cmdEliminaItem(MSF, REG&)
            End If
        End If
      End If
    End If
    
End Sub


Sub CargarGrillaCierreEym(CI1%, CANTI, LOTE$)
    If CI1% < 1 Then Exit Sub
    If CANTI < 1 Then Exit Sub
    
    Codigo$ = CODEXT$(CI1%)
    DESXRIP$ = DESCRIT0$(CI1%)
    
    MSF.Rows = MSF.Rows + 1
    REG& = MSF.Rows - 1
    MSF.TextMatrix(REG&, 1) = Codigo$
    MSF.TextMatrix(REG&, 2) = DESXRIP$
    MSF.TextMatrix(REG&, 3) = FORMATNUM$(CANTI, "F10.1")
    MSF.TextMatrix(REG&, 4) = LOTE$
    MSF.TextMatrix(REG&, 5) = CI1%
End Sub

Private Sub Label8_Change()
     CIXI% = CODINT%(Label8)
     If CIXI% > 0 Then
        Call CREALOTES(CIXI%)
        Call ARMALISLOTES(CIXI%, , , , 1)
     End If

End Sub

Private Sub Label9_DblClick()
   Exit Sub
   Call SETULTREG("!MATEROF", 0)
   '
   If List1.ListCount > 0 Then
      For i& = 1 To (List1.ListCount)
         TX$ = List1.List(i& - 1)
         CI% = CODINT%(Mid$(TX$, 1, 16))
         CANTI = XVALO(Mid$(TX$, 64, 8))
         LOTE$ = Mid$(TX$, 74, 12)
         '
         Call REPLA(ZZ$, MKI$(CI%), 1, 2)
         Call REPLA(ZZ$, MKS$(CANTI), 5, 4)
         Call REPLA(ZZ$, LOTE$, 17, 12)
         Call GRAREG("!MATEROF", ZZ$, i&)
      Next i&
   End If
   ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=Repuestos para Reparación"
   '
10 VTB% = VENTABU%("EDITREP1" + ATRI$)
   If VTB% < 0 Then
     Exit Sub
   End If
   '
   For U& = 1 To ULTREG&("!MATEROF")
     ZZ$ = REGLEIDO$("!MATEROF", U&)
     CII% = CVI(Left$(ZZ$, 2))
     LOTE$ = TRIM$(Mid$(ZZ$, 17, 12))
     CANTI = CVS(Mid$(ZZ$, 5, 4))
     If XVALO(CANTI) < 0.005 Then
        Call COMUNI("Código: " & CODEXT$(CII%) & " Sin Cantidad")
        GoTo 10
     End If
     If LOTE$ <> "" Then
        CONDI$ = "Cod_Inte=" & CII% & " AND IDENLOTE='" & LOTE$ & "'"
        If CantRegistros("LOTINGRE", CONDI$) = 0 Then
            Call COMUNI("Código: " & CODEXT$(CII%) & " - Lote Erroneo '" & LOTE$ & "'")
            Call REPLA(ZZ$, Space(12), 17, 12)
            Call GRAREG("!MATEROF", ZZ$, U&)
        End If
     End If
   Next U&
   '
   '
   List1.Clear
   
   For U& = 1 To ULTREG&("!MATEROF")
     ZZ$ = REGLEIDO$("!MATEROF", U&)
     TTXX$ = Space$(128)
     CII% = CVI(Left$(ZZ$, 2))
     Call REPLA(TTXX$, CODEXT$(CII%), 1, 16)
     Call REPLA(TTXX$, DESCRIT$(CII%), 17, 40)
     CANTI = CVS(Mid$(ZZ$, 5, 4))
     Call REPLA(TTXX$, FORMATNUM(CANTI, "F8.1"), 64, 8)
     LOTE$ = Mid$(ZZ$, 17, 12)
     Call REPLA(TTXX$, LOTE$, 74, 12)
     
     List1.AddItem TTXX$
     '
   Next U&
   CONDI$ = "MARBORRA = 1" ' borro los que estaban marcados con 1
   Call BORRAREGISTROS("REPUREPA", CONDI$, REGAF&, "NOMESS")

   Call LIMPIAR
   Call CIERRARCH("!MATEROF")
   Screen.MousePointer = 1

End Sub


Private Sub List1_Click()
     Exit Sub
     Call LIMPIAR

     TTXX$ = List1.List(List1.ListIndex)
     If ESTRAZABLE%(CODINT%(Mid$(TTXX$, 1, 16))) < 1 Then
        Call COMUNI$("Código No Trazable")
        Exit Sub
     End If
     TTXX$ = List1.List(List1.ListIndex)
     Label8 = Mid$(TTXX$, 1, 16)
     Label12 = TRIM$(Mid$(TTXX$, 17, 56))
     Label5 = Mid$(TTXX$, 64, 8)
     Label4 = Mid$(TTXX$, 72, 12)

End Sub

Private Sub text2_LostFocus()
   Text2 = FORMATNUM$(Text2, "F5.2")
End Sub

Sub LIMPIAR()
     Label8 = ""
     Label12 = ""
     Label5 = ""
     Label4 = ""
End Sub

Private Sub Text1_DblClick()
    Campos$ = "Código/f6;Nombre/A32"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = "TABLA DE TÉCNICOS"
    
    Dim obj As New RECORXET
    Set rs1 = obj.RS_TECNICOS_MDT()
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 8000
50  FRMZELECHO.Show 1

    Text1 = TRIM(RESP_ZELE_VECT(2))

End Sub

Private Sub Text5_Change()
   If XVALO(Text5) > 0 Then
     Frame9(1).Visible = True
     Label14 = NOMBRE_TEC_EXT$(XVALO(Text5))
     Text1 = Label14
     
   Else
     Frame9(1).Visible = False
   End If

End Sub



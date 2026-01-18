VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form ASINUMSE 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignación de Números de Serie a Pedidos"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   9135
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2520
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   1170
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
      Left            =   7920
      TabIndex        =   0
      Top             =   310
      Width           =   175
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00CDDADA&
      Caption         =   "SELECCIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4005
      Left            =   120
      TabIndex        =   10
      Top             =   4515
      Width           =   7965
      Begin VB.Frame Frame8 
         BackColor       =   &H00CDDADA&
         Height          =   1785
         Left            =   5805
         TabIndex        =   50
         Top             =   2040
         Width           =   1815
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
            Height          =   285
            Left            =   240
            TabIndex        =   56
            Top             =   360
            Width           =   1320
         End
         Begin VB.Label Label55 
            BackStyle       =   0  'Transparent
            Caption         =   "Solicitado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   240
            TabIndex        =   55
            Top             =   130
            Width           =   885
         End
         Begin VB.Label Label66 
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
            Height          =   285
            Left            =   240
            TabIndex        =   54
            Top             =   860
            Width           =   1320
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Asignado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   53
            Top             =   620
            Width           =   810
         End
         Begin VB.Label Label100 
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
            Height          =   285
            Left            =   240
            TabIndex        =   52
            Top             =   1360
            Width           =   1320
         End
         Begin VB.Label Label111 
            BackStyle       =   0  'Transparent
            Caption         =   "Falta"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   240
            TabIndex        =   51
            Top             =   1120
            Width           =   570
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Asignar y Transferir"
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
         Left            =   3840
         TabIndex        =   49
         Top             =   2160
         Width           =   1740
      End
      Begin VB.CommandButton Command6 
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
         Height          =   250
         Index           =   0
         Left            =   3360
         TabIndex        =   41
         Top             =   2160
         Width           =   175
      End
      Begin VB.CommandButton Command6 
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
         Height          =   250
         Index           =   1
         Left            =   3360
         TabIndex        =   40
         Top             =   2430
         Width           =   175
      End
      Begin VB.CommandButton Command6 
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
         Height          =   250
         Index           =   2
         Left            =   3360
         TabIndex        =   39
         Top             =   2700
         Width           =   175
      End
      Begin VB.CommandButton Command6 
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
         Height          =   250
         Index           =   3
         Left            =   3360
         TabIndex        =   38
         Top             =   2970
         Width           =   175
      End
      Begin VB.CommandButton Command6 
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
         Height          =   250
         Index           =   4
         Left            =   3360
         TabIndex        =   37
         Top             =   3240
         Width           =   175
      End
      Begin VB.CommandButton Command6 
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
         Height          =   250
         Index           =   5
         Left            =   3360
         TabIndex        =   36
         Top             =   3510
         Width           =   175
      End
      Begin VB.CommandButton Command2 
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
         Left            =   5340
         TabIndex        =   22
         Top             =   1380
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Left            =   4920
         TabIndex        =   21
         Top             =   1380
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Asignar y Transferir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5280
         TabIndex        =   3
         Top             =   840
         Width           =   2340
      End
      Begin VB.CommandButton Command4 
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
         Height          =   270
         Left            =   3720
         TabIndex        =   2
         Top             =   825
         Width           =   175
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccionar Lotes Para Item Trazable Activo "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Index           =   1
         Left            =   240
         TabIndex        =   48
         Top             =   2280
         Width           =   1275
      End
      Begin VB.Label Label23 
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
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   47
         Top             =   2160
         Width           =   1755
      End
      Begin VB.Label Label23 
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
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   46
         Top             =   2430
         Width           =   1755
      End
      Begin VB.Label Label23 
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
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   45
         Top             =   2700
         Width           =   1755
      End
      Begin VB.Label Label23 
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
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   44
         Top             =   2970
         Width           =   1755
      End
      Begin VB.Label Label23 
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
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   43
         Top             =   3240
         Width           =   1755
      End
      Begin VB.Label Label23 
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
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   42
         Top             =   3510
         Width           =   1755
      End
      Begin VB.Label Label11 
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
         Height          =   285
         Left            =   210
         TabIndex        =   26
         Top             =   315
         Width           =   420
      End
      Begin VB.Label Label10 
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
         Height          =   285
         Left            =   3675
         TabIndex        =   25
         Top             =   315
         Width           =   3960
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
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
         Left            =   680
         TabIndex        =   24
         Top             =   390
         Width           =   810
      End
      Begin VB.Label Label8 
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
         Height          =   285
         Left            =   1575
         TabIndex        =   23
         Top             =   315
         Width           =   2010
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
         Height          =   285
         Left            =   5520
         TabIndex        =   20
         Top             =   1380
         Width           =   2100
      End
      Begin VB.Label Label5 
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
         Height          =   285
         Left            =   1560
         TabIndex        =   19
         Top             =   1380
         Width           =   1860
      End
      Begin VB.Label Label4 
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
         Height          =   285
         Left            =   1080
         TabIndex        =   18
         Top             =   1380
         Width           =   420
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "A Depósito"
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
         Left            =   4035
         TabIndex        =   17
         Top             =   1455
         Width           =   825
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "De Depósito"
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
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   1460
         Width           =   915
      End
      Begin VB.Label Label1 
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
         Height          =   285
         Left            =   1575
         TabIndex        =   4
         Top             =   825
         Width           =   2145
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Número Serie"
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
         Index           =   0
         Left            =   495
         TabIndex        =   11
         Top             =   900
         Width           =   1050
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "DETALLE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3750
      Left            =   120
      TabIndex        =   7
      Top             =   660
      Width           =   7965
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
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
         Height          =   210
         ItemData        =   "ASINUMSE.frx":0000
         Left            =   200
         List            =   "ASINUMSE.frx":0002
         MousePointer    =   1  'Arrow
         TabIndex        =   1
         Top             =   3480
         Width           =   6945
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   3060
         Left            =   240
         TabIndex        =   57
         Top             =   600
         Width           =   7500
         _ExtentX        =   13229
         _ExtentY        =   5398
         _Version        =   393216
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Número de Serie"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   33
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label77 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   32
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label76 
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label37 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   195
         TabIndex        =   30
         Top             =   315
         Width           =   7425
      End
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   8520
      Left            =   8160
      ScaleHeight     =   8490
      ScaleWidth      =   1350
      TabIndex        =   5
      Top             =   -20
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   120
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   34
         Top             =   6000
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command26 
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
         Picture         =   "ASINUMSE.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3960
         Width           =   690
      End
      Begin VB.CommandButton Command43 
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
         Picture         =   "ASINUMSE.frx":0446
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Desasignar N.Serie a Pedido"
         Top             =   3120
         Width           =   690
      End
      Begin VB.CommandButton Command27 
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
         Picture         =   "ASINUMSE.frx":0590
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Acceso Directo a Transferencia de Equipos"
         Top             =   2370
         Width           =   690
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
         Height          =   690
         Left            =   120
         Picture         =   "ASINUMSE.frx":089A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Consulta de Asignacion de N/Serie"
         Top             =   1620
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
         Picture         =   "ASINUMSE.frx":0BA4
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir"
         Top             =   120
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
         Picture         =   "ASINUMSE.frx":0CEE
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   870
         Width           =   690
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "ASINUMSE.frx":0FF8
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "ASINUMSE.frx":122C
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "Roboto Mono"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   840
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
      Height          =   285
      Left            =   6480
      TabIndex        =   9
      Top             =   315
      Width           =   1440
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Número de Pedido"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4815
      TabIndex        =   8
      Top             =   329
      Width           =   1545
   End
End
Attribute VB_Name = "ASINUMSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TXTRAZA_ASINUMSE$

Private Sub Command1_Click()
     '
     NPSEL& = XVALO(Label28)
     Call SELECHO("INDIANSE")
     '
     Label28 = XVALO(VALACT1("INDIANSE"))
     If Label28 = 0 Then Label28 = NPSEL&
     '
     'VALIDO SI EL PEDIDO TIENE MOTIVO DE VENTA
     If VALIDAOV%(XVALO(Label28)) < 1 Then
       Call BLANFORMU
       Exit Sub
     End If
     '
     TXTRAZA_ASINUMSE$ = ""
     Call CARLISPED 'Carga la lista con el pedido seleccionado
     '
End Sub

Private Sub Command11_Click()
     '
     Call CIERRARCH("*.*")
     Unload Me
     '
End Sub

Private Sub Command2_Click()
    '
    Call SELECHO("TADEPOSI")
    VDEV$ = VALACT1$("TADEPOSI")
    If VDEV$ <> "" Then
       Text1.TEXT = VDEV$
    End If
    '
End Sub

Private Sub Command23_Click()
    '
    OPX% = ALTERNA%("Pedidos Pendientes de Asignacion\Asignacion Por Numero de Serie\Numero de Serie por Deposito\Consulta por Numero de Serie")
    '
    Select Case OPX%
        Case 1: Call FINAL("?PEDPENASIG")
        Case 2: Call FINAL("?ASIGNUMSE")
        Case 3: Call FINAL("?NUMSEPORDEP")
        Case 4: Call FINAL("?TODOSNUMEROS")
    End Select
    '
End Sub

Private Sub Command26_Click()
    '
    FSTASNUSE.Show 1
    '
End Sub

Private Sub Command27_Click()
    Command27.Enabled = False
    TRANEQUI07.Show 1
    Command27.Enabled = True
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NPED& = XVALO(Label28)
    'ORDENITEM% = XVALO(Mid$(List1.TEXT, 70, 6))
    ORDENITEM% = XVALO(Label11)
    Ci% = CODINT%(Label8)
    NUMSE$ = TRIM$(Label1)
    DEPODE% = XVALO(Text1)
    DEPORI% = XVALO(Label4)
    '
    If Ci% < 1 Then
         Call MENSERR(24, "Codigo No Válido")
         GoTo 99
      ElseIf NUMSE$ = "" Then
         Call MENSERR(24, "Falta Numero de Serie.")
         GoTo 99
      ElseIf DEPODE% < 1 Then
         Call MENSERR(24, "Falta Depósito Destino.")
         GoTo 99
      ElseIf ECOARCH$("TADEPOSI", Chr$(DEPODE%)) = "" Then
         Call MENSERR(24, "Depósito Destino no Válido.")
         GoTo 99
      ElseIf DEPORI% < 1 Then
         Call MENSERR(24, "Falta Depósito Origen.")
         GoTo 99
      ElseIf ECOARCH$("TADEPOSI", Chr$(DEPORI%)) = "" Then
         Call MENSERR(24, "Depósito Origen no Válido.")
         GoTo 99
    End If
    '

    Call TRANEQUI
    '
    Call PRINTASINUMSE
    '
    TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 17 = " & Now & vbCrLf
    Call CARLISPED
'    TX$ = List1.TEXT
'    i& = List1.ListIndex
'    Call REPLA(TX$, NUMSE$, 48, 16)
'    List1.List(i&) = TX$
'    List1.Refresh
    Call SETULTREG("!INUSEDIS", 0)
    Label1 = ""
    Label11 = ""
    Label4 = ""
    Label8 = ""
    Label10 = ""
    Text1 = ""

99  Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    '
    Ci% = CODINT(Label8)
    jj& = 0
    If Ci% > 0 Then
      If ESTRAZABLE(Ci%) < 2 Then
         Call MENSERR(24, "Código '" & CODEXT(Ci%) & "' No Trazable. Verifique !")
         Exit Sub
      End If
      '
      If ULTREG&("!INUSEDIS") > 0 Then
         Call SELECHO("!INUSEDIS")
         TTXX$ = TRIM$(VALACT1$("!INUSEDIS"))
         If TTXX$ <> "" Then
            Label1 = TTXX$
15          Label4.Caption = XVALO(VALOBADA("TANUMSE", "DepoNumse", "NumeroSerie = '" & TTXX$ & "'"))
            If XVALO(Label4.Caption) < 1 Then
               Call RECUDEPO(TTXX$, OKX%)
               If OKX% < 1 Then Call FINAL("")
               Call COMUNI("Recuperado Depósito del Equipo")
               GoTo 15
            End If
         End If
      Else
         Call MENSERR(24, "Código " & CODEXT$(Ci%) & " Sin Números de  Serie para Asignar")
      End If
    End If
End Sub

Sub RECUDEPO(NUSEX$, OKX%)
    OKX% = 0
    NORFAORI$ = SAFETEXT$(VALOBADA("TANUMSE", "IdSubOr", "NumeroSerie = '" & NUSEX$ & "'", "NOMESS"))
    If NORFAORI$ = "" Then
       Call MENSERR(24, "Imposible Recuperar - Consulte")
       Exit Sub
    End If
    '
    DEPOCI = XVALO(VALOBADA("MOVISTO", "DEPOMOVI", "IDENLOTE = '" & NORFAORI$ & "' AND CantIngre >0 AND DepoMovi > 0"))
    If DEPOCI < 1 Then DEPOCI = XVALO(CONTROL$("", "DEPOCIE"))
    '
    Call ACTUREGISTRO("MOVISTO", "DEPOMOVI", "IDENLOTE = '" & NORFAORI$ & "' AND CantIngre >0 AND DepoMovi = 0", DEPOCI, RAFE&)
    Call ACTUREGISTRO("TANUMSE", "DEPONUMSE", "IDSUBOR = '" & NORFAORI$ & "' AND DEPONUMSE = 0", DEPOCI, RAFE&)
    If RAFE > 0 Then OKX% = 1
    '
End Sub
'
Private Sub Command43_Click()

      Call COPYSTRU("INDINUSE", "DESAGNUM")
      TTXX$ = REGBLAN$("DESAGNUM")
      Call SETULTREG("DESAGNUM", 0)
      '
10    Call SELECHO("MASTER")
      Ci% = CODINT%(VALACT1$("MASTER"))
      If Ci% < 1 Then Exit Sub
      If ESTRAZABLE(Ci%) < 2 Then
         Call MENSERR(24, "Código '" & CODEXT(Ci%) & "' No Trazable. Verifique !")
         Exit Sub
      End If
      jj& = 0
      SQLX$ = "SELECT * FROM TANUMSE "
      SQLX$ = SQLX$ + " WHERE Cod_Inte = " & Ci% & ""
      SQLX$ = SQLX$ + " AND (NumRemito = '' OR NumRemito IS NULL)"
      Dim NUSERTMP As ADODB.Recordset
      Set NUSERTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With NUSERTMP
         Do While Not .EOF
         NUSERR$ = SAFETEXT$(!NUMEROSERIE)
         NPED& = XVALO(VALOBADA("PEDDETA", "NroPed", "NUMEROSERIE = '" & NUSERR$ & "' AND Status < 8 "))
         NusePed$ = TRIM$(Str(NPED&))
         If NPED& > 0 Then
              Ci% = XVALO(!COD_INTE)
              Call REPLA(TTXX$, NUSERR$, 1, 16)
              Call REPLA(TTXX$, CODEXT$(Ci%) + "  " + Left$(DESCRIT(Ci%), 20) + " - NºPed. " & NusePed$, 17, 47)
              jj& = jj& + 1
              Call GRAREG("DESAGNUM", TTXX$, jj&)
         End If
         .MoveNext
         Loop
      End With
      NUSERTMP.Close
      Set NUSERTMP = Nothing
      Call SETULTREG("DESAGNUM", jj&)
      Call CIERRARCH("DESAGNUM")

      If ULTREG&("DESAGNUM") > 0 Then
        Call SELECHO("DESAGNUM")
        NUMSE$ = TRIM$(VALACT1$("DESAGNUM"))
        TXTX$ = TRIM$(VALACT2$("DESAGNUM"))
        If NUMSE$ = "" Then GoTo 10
        '
        NN% = InStr(1, TXTX$, "NºPed.") + 6
        NPED& = XVALO(Mid$(TXTX$, NN%))
        CONDI$ = "NUMEROSERIE = '" & NUMSE$ & "'"
        Call ACTUREGISTRO("PEDDETA", "NUMEROSERIE", CONDI$, "", REGAF&, "NOMESS") ' desasigna del pedido
        Call ACTUREGISTRO("PEDDETA", "NroRemi", CONDI$, "", REGAF&, "NOMESS")
        '
        HOII% = HOY(HOS$)
        TTXX$ = Space$(96)
        Call REPLA(TTXX$, HOS$ + "  " + "Desasignación de N/Serie", 1, 56)
        Call REPLA(TTXX$, USERTER$, 69, 24)
        Call REPLA(TTXYZ$, " - Código: ", 10, 11)
        Call REPLA(TTXYZ$, CODEXT$(Ci%), 22, 16)
        Call REPLA(TTXYZ$, "- N/Serie: ", 39, 11)
        Call REPLA(TTXYZ$, NUMSE$, 51, 12)
        TTXYZ1$ = TTXYZ1$ + TTXYZ$ + Chr$(10) + Chr$(13)
        Call AGREGA_ANOTAPED(NPED&, TTXX$)
        TTXYZ1$ = Left$(TTXYZ1$, (Len(TTXYZ1$) - 2))
        Call AGREGA_ANOTAPED(NPED&, TTXYZ1$)
        Call AGREGA_ANOTAPED(NPED&, String$(88, "="))
        '
        'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
        CUERPO$ = "Desasignación de Pedido: " & NPED&
        Call ANOTANUSE(NUMSE$, CUERPO$)
        '
        Call MENSERR(24, "Número de Serie Desasignado")
        '
        NPSEL& = XVALO(ASINUMSE.Label28)
        '
        Call CARLISPED 'Carga la lista con el pedido seleccionado
        '
      Else
        Call MENSERR(24, "No Hay Números de Serie para Desasignar")
         
      End If
End Sub

Private Sub Command7_Click()
   'NO LO ACTUALIZO POR QUE NO EXITE UN BOTON COMMAND7 LO COMENTO PARA COMPILACION
   Command7.Enabled = False
'   Screen.MousePointer = 11
'
'   Call ABREPEDCLI
'    '
'   TTXYZ1$ = ""
'   VUELTA% = 0
'   For i& = 0 To List1.ListCount - 1
'    '
'    If XVALO(Mid$(List1.List(i&), 78, 1)) > 0 Then
'       '
'       NPED& = XVALO(Label28)
'       ORDENITEM% = XVALO(Mid$(List1.List(i&), 70, 6))
'       Ci% = CODINT%(TRIM$(Mid$(List1.List(i&), 1, 16)))
'       NUMSE$ = TRIM$(Mid$(List1.List(i&), 48, 16))
'       DEPONU% = XVALO(Mid$(List1.List(i&), 80, 2))
'       DepOrig% = XVALO(Mid$(List1.List(i&), 83, 2))
'       '
'       CONDICION$ = " NroPed = " & NPED& & " AND CodiInt = " & Ci% & " AND NuOrItem = " & ORDENITEM%
'       '
'       If NUMSE$ = "" Then GoTo 10 ' SI DESASIGNA NO VALIDA
'       Valiped& = XVALO(VALOBADA("PEDDETA", "NroPed", "NUMEROSERIE = '" & NUMSE$ & "'"))
'       If Valiped& = 0 Then ' validacion que otro usuaario no haya utilizado el mismo numero de serie
'10       VUELTA% = VUELTA% + 1
'         Call ACTUREGISTRO("PEDDETA", "NUMEROSERIE", CONDICION$, NUMSE$, REGAF&, "NOMESS")
'         If DEPONU% > 0 Then
'            '
'            FECHH% = HOY(HO$)
'            NUMCOM$ = ProNroComprobante("TR")
'            NUMCOM$ = "TR." + TRIM$(NUMCOM$)
'            REF$ = "Transferencia de Deposito"
'            '
'            'Registra la salida del articulo del deposito original
'            Call MOVISTOK(FECHH%, Ci%, NUMSE$, "TR", NUMCOM$, NUMCOM$, "", "", REF$, ORDENITEM%, 0, 0, 0, DepOrig%, (-1) * 1, NPED&)
'            'Registra el ingreso del articulo al deposito nuevo
'            Call MOVISTOK(FECHH%, Ci%, NUMSE$, "TR", NUMCOM$, NUMCOM$, "", "", REF$, ORDENITEM%, 0, 0, 0, DEPONU%, 1, NPED&)
'            '
'            'REALIZA LA TRANSFERENCIA ACTUALIZANDO EL DEPOSITO EN TANUMSE
'            Call ACTUREGISTRO("TANUMSE", "DEPONUMSE", "NUMEROSERIE = '" & NUMSE$ & "'", DEPONU%, REGAF&)
'            '
'         End If
'         '
'         Call REPLA(TTXYZ$, " - Código: ", 10, 11)
'         Call REPLA(TTXYZ$, CODEXT$(Ci%), 22, 16)
'         Call REPLA(TTXYZ$, "- N/Serie: ", 39, 11)
'         If TRIM$(NUMSE$) = "" Then NUMSE$ = "S/N"
'         Call REPLA(TTXYZ$, NUMSE$, 51, 12)
'         TTXYZ1$ = TTXYZ1$ + TTXYZ$ + Chr$(10) + Chr$(13)
'         '
'         TX$ = List1.List(i&) ' pongo en 0 los items leidos para evitar doble grabacion
'         Call REPLA(TX$, "0", 78, 1)
'         List1.List(i&) = TX$
'       Else
'         Call MENSERR(24, "El Número de Serie: " & NUMSE$ & " Ocupado en Pedido: " & Valiped&)
'       End If
'    End If
'    '
'   Next i&
'   If VUELTA% > 0 Then ' SI GENERO ALGUNA MODIFICACION ANOTA EN EL ANOTADOR DE PEDIDOS
'     HOII% = HOY(HOS$)
'     TTXX$ = Space$(96)
'     Call REPLA(TTXX$, HOS$ + "  " + "Asignación de N/Serie", 1, 56)
'     Call REPLA(TTXX$, USERTER$, 69, 24)
'     Call AGREGA_ANOTAPED(NPED&, TTXX$)
'     TTXYZ1$ = Left$(TTXYZ1$, (Len(TTXYZ1$) - 2))
'     Call AGREGA_ANOTAPED(NPED&, TTXYZ1$)
'     Call AGREGA_ANOTAPED(NPED&, String$(88, "="))
'     Call MENSERR(24, "Números de Series Asignados")
'   End If
'    '
   Command7.Enabled = True
   Screen.MousePointer = 1
  '
End Sub

Sub TRANEQUI()
   '
   Screen.MousePointer = 11
   'Call ABREPEDCLI
   Call ABRECONEXION
   '
   TTXYZ1$ = ""
   VUELTA% = 0
   
   NPED& = XVALO(Label28)
   Ci% = CODINT%(Label8)
   NUMSE$ = TRIM$(Label1)
   DEPONU% = XVALO(Text1)
   DepOrig% = XVALO(Label4)
   NUORD% = XVALO(Label11)
   '
   CONDICION$ = " NroPed = " & NPED&
   CONDICION$ = CONDICION$ & " AND CodiInt = " & Ci%
   CONDICION$ = CONDICION$ & " AND NuOrItem = " & NUORD% & ""
   CONDICION$ = CONDICION$ & " AND (NUMEROSERIE = '' or NUMEROSERIE IS NULL)"
   '
   Valiped& = XVALO(VALOBADA("PEDDETA", "NroPed", "NUMEROSERIE = '" & NUMSE$ & "'"))
   If Valiped& > 0 Then ' validacion que otro usuaario no haya utilizado el mismo numero de serie
       Call MENSERR(24, "Imposible Completar Transaccion Solicitada.\  \El Número de Serie: " & NUMSE$ & " Ocupado en Pedido: " & Valiped&)
       Call FINAL("")
   End If
   '
   'Call ACTUREGISTRO("PEDDETA", "NUMEROSERIE", CONDICION$, NUMSE$, REGAF&, "NOMESS")
   SQLX$ = "SELECT NUMEROSERIE FROM PEDDETA WHERE " + CONDICION$
   Dim rstPEDDETA As New ADODB.Recordset
   Set rstPEDDETA = New ADODB.Recordset
25 On Error Resume Next
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 4 = " & Now & vbCrLf
   rstPEDDETA.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 5 = " & Now & vbCrLf

   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
   End If
   On Error GoTo 0
   '
   With rstPEDDETA
     If Not (.BOF And .EOF) Then
        !NUMEROSERIE = NUMSE$
        .Update
     End If
   End With
   rstPEDDETA.Close
   Set rstPEDDETA = Nothing
   '
   '
   FECHH% = HOY(HO$)
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 6 = " & Now & vbCrLf
   NUMCOM$ = ProNroComprobante("TR")
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 7 = " & Now & vbCrLf
   NUMCOM$ = "TR." + TRIM$(NUMCOM$)
   REF$ = "Transferencia de Deposito"
   '
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 8 = " & Now & vbCrLf
   'Registra la salida del articulo del deposito original
   Call MOVISTOK(FECHH%, Ci%, NUMSE$, "TR", NUMCOM$, NUMCOM$, "", "", REF$, ORDENITEM%, 0, 0, 0, DepOrig%, (-1) * 1, NPED&)
   'Registra el ingreso del articulo al deposito nuevo
   Call MOVISTOK(FECHH%, Ci%, NUMSE$, "TR", NUMCOM$, NUMCOM$, "", "", REF$, ORDENITEM%, 0, 0, 0, DEPONU%, 1, NPED&)
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 9 = " & Now & vbCrLf
   '
   'REALIZA LA TRANSFERENCIA ACTUALIZANDO EL DEPOSITO EN TANUMSE
   '
   'Call ACTUREGISTRO("TANUMSE", "DEPONUMSE", "NUMEROSERIE = '" & NUMSE$ & "'", DEPONU%, REGAF&)
   SQLX$ = "SELECT DEPONUMSE FROM TANUMSE WHERE NUMEROSERIE = '" & NUMSE$ & "'"
   Dim rstTANUMSE As New ADODB.Recordset
   rstTANUMSE.CursorLocation = adUseServer
   Set rstTANUMSE = New ADODB.Recordset
35 On Error Resume Next
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 10 = " & Now & vbCrLf

   rstTANUMSE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 11 = " & Now & vbCrLf
   
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 35
   End If
   On Error GoTo 0
   '
   With rstTANUMSE
     If Not (.BOF And .EOF) Then
        !DEPONUMSE = DEPONU%
        .Update
     End If
   End With
   rstTANUMSE.Close
   Set rstTANUMSE = Nothing
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 12 = " & Now & vbCrLf
   '
   If TRIM$(NUMSE$) = "" Then NUMSE$ = "S/N"
   TTXYZ$ = "-Código: "
   TTXYZ$ = TTXYZ$ + CODEXT$(Ci%)
   TTXYZ$ = TTXYZ$ + " - N/Serie: "
   TTXYZ$ = TTXYZ$ + " " & NUMSE$
   
   TTXX$ = "Asignación de N/Serie" & vbCrLf
   TTXX$ = TTXX$ & TTXYZ$
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 13 = " & Now & vbCrLf
   Call AGREGA_ANOTAPED_NEW(NPED&, TTXX$)
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 14 = " & Now & vbCrLf
   
   'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
   '
   CONDI$ = " NroPed = " & NPED& & " And CodiInt = " & Ci% & " And NUMEROSERIE = '" & NUMSE$ & "'"
   NCLIE = XVALO(VALOBADA("PEDDETA", "NROCLIE", CONDI$, "NOMESS"))
   CONDI$ = " NroPed = " & NPED&
   NUMOT% = NUMOTIVO%(NPED&)
   '
   CUERPO$ = "-Asignación de Pedido: " & NPED&
   If NUMOT% > 0 Then
     DESCMOTI$ = ECOARCH$("TAMOTPRE", Chr$(NUMOT%))
     CUERPO$ = CUERPO$ & " Motivo: " & NUMOT% & "(" & TRIM$(DESCMOTI$) & ")"
   End If
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 15 = " & Now & vbCrLf
   CUERPO$ = CUERPO$ & " - Cliente: " & NCLIE
   CUERPO$ = CUERPO$ & " - " & rasocli(NCLIE)
   Call ANOTANUSE(NUMSE$, CUERPO$)
   TXTRAZA_ASINUMSE$ = TXTRAZA_ASINUMSE$ & "PUNTO 16 = " & Now & vbCrLf
   Screen.MousePointer = 1

End Sub

Private Sub Form_Load()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    CAMPOS$ = "CI/A0;Código/A20;Descripción/A32;Número de Serie/A16;Lote/A16;NORITEM/A0"
    Call CARGA_ENCABEZADO(Me.MSF, CAMPOS$, Me)

End Sub

Private Sub Label28_Click()

End Sub

Private Sub Label4_Change()
    Label5.Caption = TRIM$(ECOARCH$("TADEPOSI", MKI$(XVALO(Label4.Caption))))
End Sub

Private Sub Label8_Change()
    Label10 = ""
    Ci% = CODINT%(Label8)
    If Ci% > 0 Then
       Label10 = DESCRIT0$(Ci%)
    End If
    '
    Call SETULTREG("!INUSEDIS", 0)
    TTXX$ = REGBLAN$("INUSEDIS")
    '
    jj& = 0
    If Ci% > 0 Then
      '
      If ESTRAZABLE(Ci%) >= 2 Then
         SQLX$ = "SELECT NUMEROSERIE, CODIGO, DESCRIPCION FROM TANUMSE INNER JOIN MASTER WITH(NOLOCK) ON (TANUMSE.COD_INTE = MASTER.CODINT) "
         SQLX$ = SQLX$ + " WHERE TANUMSE.Cod_Inte = " & Ci% & ""
         SQLX$ = SQLX$ + " AND (TANUMSE.NumRemito = '' OR TANUMSE.NumRemito IS NULL)"
         SQLX$ = SQLX$ + " AND DepoNumse >= 0" ' PARA QUE TRAIGA SOLO LOS TERMINADOS.
         Set NUSERTMP = READSET(SQLX$)
         
         With NUSERTMP
           Do While Not .EOF
              NUSERR$ = SAFETEXT$(!NUMEROSERIE)
              NusePed$ = VALOBADA("PEDDETA", "NUMEROSERIE", "NUMEROSERIE = '" & NUSERR$ & "'")
              If NusePed$ = "" Then
                 Call REPLA(TTXX$, NUSERR$, 1, 16)
                 'Call REPLA(TTXX$, CODEXT$(CI%) + "  " + DESCRIT(CI%), 17, 47)
                 COD$ = SAFETEXT$(!Codigo): DESC$ = SAFETEXT$(!Descripcion)
                 Call REPLA(TTXX$, COD$ + "  " + DESC$, 17, 47)
                 jj& = jj& + 1
                 Call GRAREG("!INUSEDIS", TTXX$, jj&)
              End If
           .MoveNext
           Loop
         End With
         NUSERTMP.Close
         Set NUSERTMP = Nothing
      End If
    End If
    Call SETULTREG("!INUSEDIS", jj&)
    Call FRESHECHO("!INUSEDIS")
    '
End Sub

Private Sub MSF_Click()
    REG& = MSF.MouseRow
    COLUM% = MSF.MouseCol
    If REG& = 0 Or REG& > MSF.Rows - 1 Then Exit Sub
    
    If TRIM$(MSF.TextMatrix(REG&, 4)) <> "" Then
      Call MENSERR(24, "No Válido Des-Asignar.\Utilice Función Correspondiente.")
      Exit Sub
    End If
    Label8 = TRIM$(MSF.TextMatrix(REG&, 2))
    Label11 = TRIM$(MSF.TextMatrix(REG&, 3))
    Label1 = ""
    Text1 = TRIM$(Str$(DEPOEXPE%))

End Sub

Private Sub MSF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
      Command4_Click
    End If

End Sub


Private Sub Text1_Change()
    Label6.Caption = TRIM$(ECOARCH$("TADEPOSI", MKI$(XVALO(Text1))))
End Sub

Sub CARLISPED() 'Carga la lista con el pedido seleccionado
    '
    NPSEL& = XVALO(ASINUMSE.Label28)
    Label1 = "": Label11 = ""
    Label4 = ""
    Text1 = ""
    MSF.Rows = 1
    'ASINUMSE.List1.Clear
    '
    'Call CARGALISTA(ASINUMSE.List1, "PEDDETA", "CodiInt/C16+MASTER/A29; NUMEROSERIE/A16; NuOrItem/F5", "NroPed = " & NPSEL& & " AND Status < 8  AND ((CanPed - CantEnt) >= 0.05) ORDER BY NuOrItem ASC")
    REG& = 0
    SQLX$ = "SELECT CODIINT, CodiExt,Descrip,NUMEROSERIE,IdenLote,NuOrItem FROM PEDDETA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NroPed = " & NPSEL& & " AND Status < 8  AND ((CanPed - CantEnt) >= 0.05) ORDER BY NuOrItem ASC"
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        REG& = REG& + 1
        MSF.Rows = REG& + 1
         
        MSF.TextMatrix(REG&, 1) = XVALO(!CODIINT)
        MSF.TextMatrix(REG&, 2) = SAFETEXT$(!CodiExt)
        MSF.TextMatrix(REG&, 3) = SAFETEXT$(!Descrip)
        MSF.TextMatrix(REG&, 4) = SAFETEXT$(!NUMEROSERIE)
        MSF.TextMatrix(REG&, 5) = SAFETEXT$(!IDENLOTE)
        MSF.TextMatrix(REG&, 6) = XVALO(!NuOrItem)
        .MoveNext
      Loop
    End With
    '
End Sub

Sub PRINTASINUMSE()
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FOASNUSE"))
    If FORIPRE$ = "" Then Exit Sub
    '
    NUMEDOC$ = Label1.Caption
    TIPODOC$ = "Asignacion Numero de Serie"
    '
    'CODIGO DE BARRAS DE NUMEROS DE SERIE
    NUMSE$ = TRIM$(Label1)
    'Convierte al formato del codigo de barras
    NUMSE1$ = CONVIERTOCODIGOBARRAS$(NUMSE$)
    '
    FECHH = HOY(HO$)
    CODPARAM$(1) = "REF-MAT1": DOCPARAM$(1) = Label1.Caption     'Numero de serie
    CODPARAM$(2) = "COD-MAT": DOCPARAM$(2) = Label8.Caption      'codigo del articulo
    CODPARAM$(3) = "REF-MAT2": DOCPARAM$(3) = Text1.TEXT      'numero de Deposito de destino
    CODPARAM$(4) = "REF-MAT3": DOCPARAM$(4) = TRIM$(ECOARCH$("TADEPOSI", MKI$(XVALO(Text1.TEXT))))   'Nombre Deposito de destino
    CODPARAM$(5) = "FECH-EMI": DOCPARAM$(5) = FECHATEX(FECHH)
    CODPARAM$(6) = "N-PLANO": DOCPARAM$(6) = NUMSE1$    'Numero de serie (CODIGO DE BARRAS)
    CAPARDOC = 6
    '
    Call PRINTDOC("@" + FORIPRE$)
    '
End Sub
Sub BLANFORMU()
    Label1 = ""
    Label11 = ""
    Label4 = ""
    Label8 = ""
    Text1 = ""
    Label28 = ""
    'List1.Clear
    MSF.Rows = 1
End Sub



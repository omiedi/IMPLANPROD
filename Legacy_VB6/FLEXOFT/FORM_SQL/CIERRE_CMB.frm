VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CIERRE_CMB 
   BackColor       =   &H00808080&
   Caption         =   "Informar Cierres de Orden de Fabricación "
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   13680
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   13680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TXTTOTALCIERRES 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   " "
      Top             =   4200
      Width           =   960
   End
   Begin VB.TextBox TXTCANTICORTES 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   " "
      Top             =   3840
      Width           =   945
   End
   Begin VB.TextBox lblidregistro 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   43
      Text            =   " "
      ToolTipText     =   "ID (Tabla Movisto)"
      Top             =   3840
      Width           =   945
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Caption         =   "Completar Solo Para el Caso de Bobinas Laminadas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   7680
      TabIndex        =   25
      Top             =   1560
      Width           =   5055
      Begin VB.TextBox TXTMBB 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   64
         Text            =   " "
         Top             =   1650
         Width           =   840
      End
      Begin VB.TextBox TXTKBB 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   " "
         Top             =   1650
         Width           =   840
      End
      Begin VB.CommandButton Command5 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   7.5
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4800
         TabIndex        =   62
         ToolTipText     =   "Bobina Base"
         Top             =   730
         Width           =   175
      End
      Begin VB.CommandButton Command2 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   7.5
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4800
         TabIndex        =   61
         ToolTipText     =   "Lotes Disponibles de Bobina Base"
         Top             =   1175
         Width           =   175
      End
      Begin VB.TextBox TXTCODIGOBOBINABASE 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         TabIndex        =   59
         Top             =   720
         Width           =   2760
      End
      Begin VB.TextBox TXTLECTURA 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         TabIndex        =   57
         Top             =   1175
         Width           =   2760
      End
      Begin VB.TextBox TXTCANTIDADACERRAR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2040
         TabIndex        =   26
         Text            =   " "
         Top             =   285
         Width           =   2760
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "K.Bob. Base:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   2760
         TabIndex        =   66
         Top             =   1680
         Width           =   1365
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "M.Bob. Base:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   65
         Top             =   1680
         Width           =   1485
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Código Bob.Base:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   60
         Top             =   760
         Width           =   1965
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote Bobina Base:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   58
         Top             =   1175
         Width           =   1965
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad a Cerrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1965
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   1620
      Left            =   4080
      TabIndex        =   22
      Top             =   1560
      Width           =   3375
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         TabIndex        =   56
         Text            =   " "
         Top             =   840
         Width           =   1605
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Balanza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Lee el Peso de La Balanza"
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Metros"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   55
         Top             =   840
         Width           =   765
      End
      Begin VB.Label LblKgImpresos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   1560
         TabIndex        =   24
         ToolTipText     =   "Kg. Con Tara Descontada"
         Top             =   360
         Width           =   1605
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   10320
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13590
      Begin VB.CommandButton Command1 
         Caption         =   "O.K."
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
         Left            =   12840
         Picture         =   "CIERRE_CMB.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Genera Transacción de Cierre de Orden de Fabricación Activa"
         Top             =   2990
         Width           =   650
      End
      Begin VB.CommandButton Command4 
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   7.5
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3300
         TabIndex        =   41
         ToolTipText     =   "Ordenes de Fabricación (Todas)"
         Top             =   360
         Width           =   175
      End
      Begin VB.TextBox TXTNROCOTIZA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12800
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   " "
         Top             =   1080
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   " "
         Top             =   1080
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00808080&
         Height          =   1620
         Left            =   170
         TabIndex        =   28
         Top             =   1560
         Width           =   3765
         Begin VB.CommandButton CMDCM 
            Caption         =   "..."
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
            Left            =   735
            TabIndex        =   33
            Top             =   960
            Width           =   175
         End
         Begin VB.CommandButton CMDDE 
            Caption         =   "..."
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
            Left            =   735
            TabIndex        =   29
            Top             =   360
            Visible         =   0   'False
            Width           =   175
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "DEPÓSITO DE CONSUMO"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   3495
         End
         Begin VB.Label LBLDEPOCONS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   120
            TabIndex        =   35
            Top             =   960
            Width           =   585
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   960
            TabIndex        =   34
            Top             =   960
            Width           =   2505
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "DEPÓSITO DE ENTRADA"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   120
            Width           =   2655
         End
         Begin VB.Label LBLDEPOENT 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   960
            TabIndex        =   30
            Top             =   360
            Width           =   2505
         End
      End
      Begin VB.TextBox TXTPENDIENTE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   " "
         Top             =   1080
         Width           =   1185
      End
      Begin VB.TextBox Text14 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1230
         LinkTimeout     =   0
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   " "
         Top             =   1080
         Width           =   1080
      End
      Begin VB.TextBox Text13 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12800
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   " "
         Top             =   720
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox Text12 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   " "
         Top             =   720
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   " "
         Top             =   360
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.CommandButton Command3 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   7.5
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3100
         TabIndex        =   1
         ToolTipText     =   "Ordenes de Fabricación Abiertas"
         Top             =   360
         Width           =   175
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Ingresar Código de Producto Terminado"
         Top             =   720
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cotización N°:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10900
         TabIndex        =   40
         Top             =   1080
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ord. N°:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   8520
         TabIndex        =   38
         Top             =   1080
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   " Pendiente O.F.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5400
         TabIndex        =   21
         Top             =   1080
         Width           =   1605
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.O.F..:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   165
         TabIndex        =   19
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant.del Pedido:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   11150
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   1605
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   10320
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   3210
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pedido N°:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8400
         TabIndex        =   14
         Top             =   720
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8280
         TabIndex        =   13
         Top             =   420
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3360
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   4890
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1200
         TabIndex        =   6
         ToolTipText     =   "Seleccione Orden de Fabricación Asociada"
         Top             =   360
         Width           =   1890
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -150
         TabIndex        =   5
         Top             =   795
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de O/F: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -120
         TabIndex        =   4
         Top             =   420
         Width           =   1335
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3480
         TabIndex        =   3
         Top             =   360
         Width           =   4770
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "CIERRE_CMB.frx":030A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CIERRE_CMB.frx":053E
      TabIndex        =   9
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   1650
      Left            =   0
      TabIndex        =   42
      Top             =   3840
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   2910
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   5786388
      ForeColorFixed  =   16777215
      BackColorSel    =   10468997
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   1890
      Left            =   0
      TabIndex        =   49
      Top             =   5400
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   3334
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   5786388
      ForeColorFixed  =   16777215
      BackColorSel    =   10468997
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF3 
      Height          =   1845
      Left            =   0
      TabIndex        =   50
      Top             =   7560
      Width           =   13665
      _ExtentX        =   24104
      _ExtentY        =   3254
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   5786388
      ForeColorFixed  =   16777215
      BackColorSel    =   10468997
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   0
      Top             =   4000
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Caption         =   "Cierres Realizados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   53
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H009FBE85&
      Caption         =   "Necesidad de Consumo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   52
      Top             =   7320
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H009FBE85&
      Caption         =   "Material En linea "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   51
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Cierres"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   11160
      TabIndex        =   48
      Top             =   4200
      Width           =   1245
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cant. de Cierre:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   10920
      TabIndex        =   46
      Top             =   3915
      Width           =   1815
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Id Registro:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   11280
      TabIndex        =   45
      Top             =   3915
      Width           =   1335
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuMarcadeCierre 
         Caption         =   "Marcar Como Cerradas Ordenes de Fabircación"
      End
      Begin VB.Menu mnuAnulOrofa 
         Caption         =   "Anular Cierre de Orden de Fabricaciòn."
      End
   End
   Begin VB.Menu mnuCtaListado 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuLoteDespbobFab 
         Caption         =   "Lotes de Bobinas fabricadas Disponibles"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver "
      Begin VB.Menu mnuCierresRealizados 
         Caption         =   "Ver Foto de Situación  de Cierre (Cierres Realizados)"
      End
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
      Begin VB.Menu mnuOrigScrap 
         Caption         =   "Informar Origen de Scrap"
      End
      Begin VB.Menu mnuReaperturaOF 
         Caption         =   "Re-Apertura O/F's con Saldo"
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuración"
      Begin VB.Menu mnuTabAux 
         Caption         =   "Tablas Auxiliares"
      End
      Begin VB.Menu mnuTablaConfScrap 
         Caption         =   "Configuración Tabla de Scrap"
      End
   End
   Begin VB.Menu mnuAccesDirect 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuUbic 
         Caption         =   "Ubicaciones"
      End
   End
End
Attribute VB_Name = "CIERRE_CMB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMDCM_Click()
    CMDCM.Enabled = False
    Call SELECHO("TADEPOSI")
    VDEV$ = VALACT1$("TADEPOSI")
    If VDEV$ <> "" Then
      Me.LBLDEPOCONS = VDEV$
    End If
    CMDCM.Enabled = True
End Sub

Private Sub CMDDE_Click()
    CMDDE.Enabled = False
    Call SELECHO("TADEPOSI")
    VDEV$ = VALACT1$("TADEPOSI")
    If VDEV$ <> "" Then
      LBLDEPOENT = VDEV$
    End If
    CMDDE.Enabled = True
End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
    NORFA$ = TRIM(Label4)
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    
    If STATORFA%(NORFA$) >= 3 Then
       Call COMUNI(" Imposible Cerrar Orden de Fabricación Se encuentra Cerrada o Anulada")
       GoTo 99
    End If
    '
    If TRIM(NORFA$) = "" Then
       Call COMUNI("Orden de Fabricación No Válida")
       GoTo 99
    End If
    '
    If STATORFA%(NORFA$) > 1 Then
       Call COMUNI("Imposible Cerrar, Orden de Fabricación Seleccionada Se Encuentra Cerrada")
       GoTo 99
    End If
    
    CANTI = XVALO(TXTCANTIDADACERRAR)
    If CANTI < 1 Then
       Call COMUNI("Cantidad Ingresada No Válida")
       Call PINTATEXT(TXTCANTIDADACERRAR)
       GoTo 99
    End If
'    If MaximoPosibleACerrar&(CIFAB%, NORFA$) < CANTI Then
'
'       MJE$ = "La Cantidad Ingresada: " & CANTI & " +  La Cantidad Yá Cerrada\Supera la Cantidad Proyectada + La Tolerancia Permitida"
'       OPX% = COMALTER%(MJE$ & "\\Cancelar Cierre\Generar Cierre (Con Derechos)")
'       If OPX% = 2 Then
'          If FIRMAELECTRONICAOK% < 1 Then
'             GoTo 50 'SI NO APRUEBA FIRMA ELECT... NO CONTINUA
'          Else
'             Call COMUNI("Firma Electrónica Aprobada")
'          End If
'       End If
'
'       If OPX% < 2 Then
'50        TXTPENDIENTE = CANT_OF_PENDIENTEDECIERRE(CIFAB%, NORFA$) 'actualizo por las dudas por si se realizo un cierre desde otra terminal
'          Call PINTATEXT(TXTCANTIDADACERRAR)
'          GoTo 99
'       End If
'    End If
    If XVALO(LblKgImpresos) < 0.05 Then
       Call COMUNI("Cantidad de Pesada No Válida")
       'Call PINTATEXT(LblKgImpresos)
       GoTo 99
    End If
    
    If XVALO(Me.LBLDEPOENT) < 1 Then
       Call COMUNI("Depósito Ingreso No Válido")
       GoTo 99
    End If
    If XVALO(Me.LBLDEPOCONS) < 1 Then
       Call COMUNI("Depósito Consumo No Válido")
       GoTo 99
    End If
    '
    CIFAB% = CODINT%(Text10)
    If CIFAB% < 1 Then
       Call COMUNI("Código No Válido")
       GoTo 99
    End If
    '
    CFAMILIA$ = TRIM$(UCase$(CODFAMIL$(CIFAB%))) 'TELAS PLANAS Y TELAS TUBULARES

    PesoDelMaestro = PESMETRO(CIFAB%)
    MetrosDelMaestro = PESUNI(CIFAB%)
    If PesoDelMaestro <= 0 Then
       Call COMUNI("Falta Configurar Peso en el Maestro de Artículos Para el Item Activo ")
       GoTo 99
    End If
    If MetrosDelMaestro <= 0 Then
       Call COMUNI("Falta Configurar Metros en el Maestro de Artículos Para el Item Activo ")
       GoTo 99
    End If

    pesoReal = XVALO(LblKgImpresos)
    metrosReal = XVALO(Text2)
    
    If metrosReal <= 0 Then
       Call COMUNI("Falta Ingresar los Metros Reales de la Bobina Activa ")
       GoTo 99
    End If
    '
    If CFAMILIA$ = "PL" Or CFAMILIA$ = "TL" Or CFAMILIA$ = "R1" Or CFAMILIA$ = "R2" Then 'TELAS LAMINADAS, TELAS TUBULARES LAMINDAS
       LOTE$ = UCase$(TRIM$(Me.TXTLECTURA)) 'lote pertenece a la bobina base             'TELAS REFILADAS R1 R2
       If LOTE$ = "" Then
        Call COMUNI("Falta Ingresar Lote de Bobina Base ")
        GoTo 99
       End If
       If TRIM$(TXTCODIGOBOBINABASE) = "" Then
        Call COMUNI("Lote Ingresado No Válido")
        GoTo 99
       End If
       codielem% = CODINT%(TXTCODIGOBOBINABASE)
       If Es_componente%(CIFAB%, codielem%) < 1 Then
        Call COMUNI("El Lote Ingresado No Corresponde a un Componente de La Bobina a Informar el Cierre")
        GoTo 99
       End If
       CantidDisponibleLote% = XVALO(VALOBADA("MOVISTO", "Sum(CANTINGRE-CANTSALID)", "IDENLOTE= '" & LOTE$ & "'", "NOMESS"))
       If CantidDisponibleLote% < 1 Then
        Call COMUNI("Lote de Bobina Base No Vàlido o Sin Saldo")
        GoTo 99
       End If
       
       If XVALO(Me.TXTKBB) < 1 Then
        Call COMUNI("Lote de Bobina Base Falta Informar Kilos")
        GoTo 99
       End If

       If XVALO(Me.TXTMBB) < 1 Then
        Call COMUNI("Lote de Bobina Base Falta Informar Metros")
        GoTo 99
       End If

       
    End If
       
       
    
   CFAMILIA$ = TRIM$(UCase$(CODFAMIL$(CIFAB%))) 'TELAS PLANAS Y TELAS TUBULARES
   
   If CFAMILIA$ = "PL" Or CFAMILIA$ = "TL" Or CFAMILIA$ = "R1" Or CFAMILIA$ = "R2" Then 'TELAS LAMINADAS, TELAS TUBULARES LAMINDAS
      'busco el consumo x bobina configurado en tabla MATCONLA segun codigo
      KGBobBase# = XVALO(TXTKBB)
      KGBobLaminada# = XVALO(LblKgImpresos)
      ConsumoMaterialLaminado# = KGBobLaminada# - KGBobBase#:  If ConsumoMaterialLaminado# < 0 Then ConsumoMaterialLaminado# = 0
'      ConsumoMatLamXBobina = XVALO(VALOBADA("DATASQL", "CAMPO4", "CAMPO1 = '" & Text10 & "' AND LISTA = 'MATCONLA'", "NOMESS"))
'      If ConsumoMatLamXBobina <= 0 Then
'        Call COMUNI("Imposible Realizar Cierre, Falta Registrar el Consumo del Material Laminado Para el Código de La Orden de Fabricación Activa ")
'        GoTo 99
'      End If
   End If
    
    coeficienteSegunMaestro = PesoDelMaestro / MetrosDelMaestro
    coeficienteSegunDatosPantalla = pesoReal / metrosReal
    resultadoDiferenciaAbs = Abs(coeficienteSegunMaestro - coeficienteSegunDatosPantalla)
    resultadoSoportado = coeficienteSegunMaestro * 1.05 'aca cambie a 5% en vez de 1.15 22-04-22
    
    If (coeficienteSegunMaestro + resultadoDiferenciaAbs) > resultadoSoportado Then
'    If (Abs(coeficienteSegunMaestro) * 1.15) < coeficienteSegunDatosPantalla Then
        TXXK$ = "El Peso Registrado Excede La tolerancia de 5% Dada Por el Coeficiente Resultante\ de la Configuración del Maestro Vs. Los Datos en Pantalla\O.F.: " & Label4 & " Peso Kg.: " & LblKgImpresos & " Metros: " & Text2
        OPX% = COMALTER%("" & TXXK$ & "\\Canelar\Cerrar Igual")
        If OPX% < 2 Then GoTo 99 'si cancela se va
        'si continua valida derechos
        If COTICONT.FIRMAELECTRONICAOK% <= 0 Then
            GoTo 99
        End If
    End If
    
    If (XVALO(TXTPENDIENTE) - 1) < 4 Then
       Call COMUNI("Atención !!!!!\La Orden de Fabricación se Está por Acabar\Realizado Este Cierre,Restan: " & XVALO(TXTPENDIENTE) - 1)
    End If
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans

    
    'MODO CIERREFACIL
    FEC$ = FECHATEX$(HOY(HO$))
    STG$ = NORFA$ & "|"
    STG$ = STG$ & CANTI & "|"
    STG$ = STG$ & XVALO(Me.LBLDEPOENT) & "|"
    STG$ = STG$ & XVALO(Me.LBLDEPOCONS) & "|"
    STG$ = STG$ & "0"
    
    CIERFACIL.STRINGORFA = STG$
    Call ORFCIE("", 0, "")
    
    'BUSCO EL VALOR MAXIMO DE LA OF ACTIVA 'OF-000001-001'
    SQLX$ = "SELECT MAX(CAST(substring(IDENLOTE,10,3) AS INT)) as IdentiLote "
    SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " where DOPRILAR = '" & NORFA$ & "'"
    SQLX$ = SQLX$ + " AND CANTINGRE > 0 AND COD_INTE = " & CIFAB%
    SQLX$ = SQLX$ + " AND CODIMOVI = 'OF' "
    Set RS0 = READSET(SQLX$)
    SiguienteIdentificador& = 1 + XVALO(RS0!IdentiLote)
    
    'BUSCO EL ID DEL MOVIMIENTO GENEERADO
    SQLX$ = "SELECT TOP 1 Movisto_id FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE IDENLOTE = '" & NORFA$ & "'  AND DOPRILAR = '" & NORFA$ & "'"
    SQLX$ = SQLX$ + " AND CANTINGRE > 0 AND COD_INTE = " & CIFAB%
    SQLX$ = SQLX$ + " ORDER BY MOVISTO_ID DESC"
    Set rs = READSET(SQLX$)
    IdentifMovisto& = XVALO(rs!MOVISTO_ID)
    '
    'formateo el identificador del lote ej 002 y luego concateno a la of
    'GUARDO EL PESO , LOS METROS Y EN IDENLOTE EL NUMERO DE ORFA ACTIVO CON EL IDENTIFICADOR UNIVOCO EJ: OF-000001-005
    SigIdentLote$ = FORMATOCON0$(SiguienteIdentificador&, 3)
    LoteNorfa$ = Left$(NORFA$, 9) & SigIdentLote$ 'ej: 'OF-000001' & '001' va a quedar OF-000001001 p mantner los 12 caracteres
    SQLX$ = "UPDATE MOVISTO SET METROS = '" & XVALO(FORMATNUM$(metrosReal, "F12.2")) & "'"
    SQLX$ = SQLX$ + " ,  PESOKG = '" & FORMATNUM$(pesoReal, "F16.5") & "'"
    SQLX$ = SQLX$ + " ,  IDENLOTE = '" & LoteNorfa$ & "'"
    SQLX$ = SQLX$ + " WHERE MOVISTO_ID = " & IdentifMovisto&
    FLEXCONN.Execute (SQLX$)
    '
    'esto es para actualizar el consumo con el peso informado de la bobina con una relacion 1 a 1
    'pertenece al consumo de z extrusion
    
    SQLX$ = "SELECT COD_INTE,MOVISTO_ID FROM MOVISTO WITH(NOLOCK) WHERE DOPRILAR =  '" & NORFA$ & "' "
    SQLX$ = SQLX$ + " AND CODIMOVI = 'CF' AND MOVISTO_ID > " & IdentifMovisto& 'BUSCAR LOS CONSUMOS POSTERIORES AL CIERRE
    SQLX$ = SQLX$ + " ORDER BY MOVISTO_ID"
    
    Set rs = READSET(SQLX$) 'SOLO DEBERIA TENER UN SOLO COMPONENTE, PERO RECORRO POR LAS DUDAS Y ME QUEDO SOLO
    Do While Not rs.EOF       'LOS QUE PERTENECEN A LA FAMILIA .....
      CI1% = XVALO(rs!cod_inte)
      idMov& = XVALO(rs!MOVISTO_ID)
      'en todos los consumos le agrega en el campo docuorig el numero de LoteNorfa$ (OF-000001001)
      'esto se usa para cuando se anula un cierre se puedan borrar los consumos relacionados a dicho LoteNorfa$
      SQLX$ = "UPDATE MOVISTO SET DocuOrig = '" & LoteNorfa$ & "'"
      SQLX$ = SQLX$ + " WHERE DOPRILAR =  '" & NORFA$ & "' AND CODIMOVI = 'CF' AND MOVISTO_ID = " & idMov&
      FLEXCONN.Execute (SQLX$)

      CFAMILIA$ = TRIM$(UCase$(CODFAMIL$(CI1%)))
      If CFAMILIA$ = "M1" Then 'MATEERIAL DE CONSUMO BOBINA BASICA
        'aca agregar por eltipo de familia actualizar el consumo
        SQLX$ = "UPDATE MOVISTO SET CANTSALID = '" & FORMATNUM$(pesoReal, "F16.5") & "'"
        SQLX$ = SQLX$ + " ,PESOKG ='" & FORMATNUM$(pesoReal, "F16.5") & "'" 'EL PESO ES IGUAL P Q U.MED ES KG
        SQLX$ = SQLX$ + " WHERE DOPRILAR =  '" & NORFA$ & "' AND CODIMOVI = 'CF' AND COD_INTE = " & CI1%
        SQLX$ = SQLX$ + " AND MOVISTO_ID = " & idMov&
        FLEXCONN.Execute (SQLX$)
      ElseIf CFAMILIA$ = "M2" Then 'ConsumoMatLamXBobina CONSUME SEGUN UNA TABLA
'          SQLX$ = "UPDATE MOVISTO SET CANTSALID = '" & FORMATNUM$(ConsumoMatLamXBobina, "F16.5") & "'"
'          SQLX$ = SQLX$ + " ,PESOKG ='" & FORMATNUM$(ConsumoMatLamXBobina, "F16.5") & "'" 'EL PESO ES IGUAL P Q U.MED ES KG
'          si modifica que consuma por diferencia entre bobina base y laminada
          SQLX$ = "UPDATE MOVISTO SET CANTSALID = '" & FORMATNUM$(ConsumoMaterialLaminado#, "F16.5") & "'"
          SQLX$ = SQLX$ + " ,PESOKG ='" & FORMATNUM$(ConsumoMaterialLaminado#, "F16.5") & "'" 'EL PESO ES IGUAL P Q U.MED ES KG
          SQLX$ = SQLX$ + " WHERE DOPRILAR =  '" & NORFA$ & "' AND CODIMOVI = 'CF' AND COD_INTE = " & CI1%
          SQLX$ = SQLX$ + " AND MOVISTO_ID = " & idMov&
          FLEXCONN.Execute (SQLX$)
          'ACA ABAJO ES PARA LAS BOBINAS LAMINADAS QUE CONSUMEN BOBINA BASE.o
      ElseIf CFAMILIA$ = "TP" Or CFAMILIA$ = "TT" Or CFAMILIA$ = "PL" Or CFAMILIA$ = "TL" Or CFAMILIA$ = "PA" Then 'TELAS PLANAS O T TUBULARES O LAMINADAS  'CAMBIO DE DEPOSITO PARA QUE LA BOBINA SE CONSUMA DEL DEPOSITO DE BOBINAS
          DEPOBOBINA% = XVALO(Control("", "DEPOCIE")) 'TAMBIEN ES PARA AGREGARLE EL LOTE A LAS BOBINAS DE BASE CUANDO SE TRATA DE UNA BOBINA LAMINADA
          If InStr(EPAC$, "CAMBACERES") > 0 And InStr(EPAC$, "QUILMES") < 1 Then
            DEPOBOBINA% = XVALO(LBLDEPOCONS) ' se cambia para que en cambaceres si son telas planas o tubulares se consuma del 3 que es el que le pone al elegir la o.f.
          End If
          SQLX$ = "UPDATE MOVISTO SET DEPOMOVI = " & DEPOBOBINA% & ", IDENLOTE =  '" & LOTE$ & "'"
          SQLX$ = SQLX$ + " ,PESOKG ='" & FORMATNUM$(XVALO(TXTKBB), "F16.5") & "'"
          SQLX$ = SQLX$ + " ,METROS ='" & FORMATNUM$(XVALO(TXTMBB), "F16.5") & "'"
          SQLX$ = SQLX$ + " WHERE DOPRILAR =  '" & NORFA$ & "' AND CODIMOVI = 'CF' AND COD_INTE = " & CI1%
          SQLX$ = SQLX$ + " AND MOVISTO_ID = " & idMov& 'PARA REFORZAR CUAL ES EL REGISTRO A MODIFICAR
          FLEXCONN.Execute (SQLX$)
      End If
      rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    
    If ESTRAZABLE%(CIFAB%) = 1 Then
      'creo el lote
      FE% = HOY(HO$)
      Call CREALOTIN(FE%, FE%, FE%, CIFAB%, LoteNorfa$, "", "", CDbl(1))
      Call ACSALOTE(CIFAB%, LoteNorfa$, "NOMESS")
    End If
    Call PRINTETIQORFA_CAMBACERES(NORFA$, LoteNorfa$, pesoReal, metrosReal, 1)

ERROR_GUARDAR:
    If Err <> 0 Then
       FLEXCONN.RollbackTrans
       MsgBox Err.Description
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       GoTo 99
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Grabación Realizado Exitosamente")

    End If
    Label4.Caption = ""
    Call BLANFORMU
    
99  Command1.Enabled = True
    Set CTP = Nothing
    

End Sub
 Sub PRINTETIQORFA_CAMBACERES(NORFA$, LoteNorfa$, pesoReal, metrosReal, CANCOPIAS&)
 
 '**** OJO TAMBIEN SE USA EN QUILMES DESDE EL FORMULARIO DE INFUBIPT   ******************
        EPAC$ = UCase$(TRIM$(EMPREAC$))
        Call SETPRINPORT("RESTORE")
        PRINTERPORT$ = ""

        CI1% = CODIFAB%(NORFA$)
        If CI1% < 1 Then
           Call COMUNI("IMPOSIBLE CONTINUAR NO SE PUDO OBTENER EL CÓDIGO/PERTENECIENTE A LA ORDEN DE FABRICACIÓN ACTIVA")
           Exit Sub
        End If
        PIMPRE$ = TRIM$(Control(IDENTER$, "PORTETIQ")) 'configurado en opvalemat (Pto.O/Etiqueta OF:)
        If PIMPRE$ = "" Then PIMPRE$ = Control$("*", "PORTETIQ")
        Call SETPRINPORT(PIMPRE$)
        
        Codigo$ = CODEXT$(CI1%)
        Descripcion$ = DESCRIT0$(CI1%)
       
        NumeroProveedor% = 0
        If InStr(EPAC$, "QUILMES") Then
           RazonSocialProveedor$ = "QUILMES"
        Else
           RazonSocialProveedor$ = "CAMBACERES"
        End If
        CONDI$ = "CODIMOVI = 'OF' AND CANTINGRE > 0 AND DOPRILAR = '" & NORFA$ & "' AND IDENLOTE = '" & LoteNorfa$ & "'"
        'METROSREALES$ = XVALO(VALOBADA("MOVISTO", "METROS", CONDI$, "NOMESS"))
        METROSREALES$ = TRIM$(FORMATNUM$(metrosReal, "F10.1"))
'        PESOENKG# = XVALO(VALOBADA("MOVISTO", "PESOKG", CONDI$, "NOMESS"))
        PESOENKG$ = TRIM$(FORMATNUM$(pesoReal, "F10.1"))
        
        FECHAGENERACION% = CVINT(VALOBADA("MOVISTO", "FECHMOV", CONDI$, "NOMESS"))
        FORIPRE$ = TRIM$(Control("ORDEFABR", "ETIQUEOF"))
        If EXISTE%(LUDAT$ + FORIPRE$ + ".SQR") = 0 Then Exit Sub

        IMAG$ = "IMAGORFA_QR.JPG"
        IDENTIFICADOR$ = "QR_ORDFABRA"
        ANCHOX% = 0
        ALTOX% = 0

        On Error Resume Next
        Kill LUCOM$ & IMAG$
        
        TEXTOQR$ = Codigo$ & "|" & NORFA$ & "|" & LoteNorfa$ & "|" & PESOENKG$ & "|" & METROSREALES$

        On Error GoTo 0
        'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
        QRX$ = TRIM$(IMAGENCODBARRA$(LUCOM$, IMAG$, TEXTOQR$, "QR_CODE", ANCHOX%, ALTOX%))
        Call T_Espera(3)
        Image1.Picture = LoadPicture(LUCOM$ & IMAG$)
        Call GRABAR_QR(LUCOM$ & IMAG$, IDENTIFICADOR$)

'        FILTX$ = LOTE$ & ";" & Codigo$ & ";" & Descripcion$ & ";" & RazonSocialProveedor$ & ";" & TRIM$(FORMATNUM$(METROSREALES$, "F12.1")) & ";" & PESOENKG$ & ";" & NORFA$ & ";" & USERNAME$ & ";" & FECHATEX$(FECHAGENERACION%)
        FILTX$ = LoteNorfa$ & ";" & Codigo$ & ";" & Descripcion$ & ";" & RazonSocialProveedor$ & ";" & METROSREALES$ & ";" & PESOENKG$

        Call STDFORM(FORIPRE$, FILTX$, "PRINT", "", 0, CANCOPIAS&)

        Call SETPRINPORT("RESTORE")
        PRINTERPORT$ = ""

 End Sub
Sub BLANFORMU()
    Text6 = ""
    Label21 = ""
    Text12 = ""
    Text1 = ""
    Text13 = ""
    TXTCANTIDADACERRAR = ""
    LblKgImpresos = ""
    TXTPENDIENTE = ""
    Label3 = ""
    Label2 = ""
    Text10 = ""
    Text14 = ""
    BALANZA.TXTPESO = ""
    PREBALANZA.LblKgImpresos = ""
    TXTLECTURA = ""
    TXTCODIGOBOBINABASE = ""
    TXTNROCOTIZA = ""
    Text2 = ""
    Me.TXTKBB = ""
    Me.TXTMBB = ""
    
End Sub
Private Sub Command11_Click()
        LblKgImpresos = 0
        PREBALANZA.FORMULARIO_NAME = Me.Name
        PREBALANZA.Show 1
        PESOLEIDO# = XVALO(PREBALANZA.LblKgImpresos)
        If PESOLEIDO# <= 0 Then
           Call COMUNI("Peso Ingresado No Válido")
           Exit Sub
        End If
        LblKgImpresos = PESOLEIDO#
End Sub



Private Sub Command2_Click()
   Command2.Enabled = False
   If TRIM$(TXTCODIGOBOBINABASE) = "" Then
      Call COMUNI("Falta Seleccionar Código de Bobina Base")
      GoTo 99
   End If
   '
   
    TituloLista$ = "Lotes Disponibles"
    Campos$ = "Lote/A14;Cantidad/F11.0"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    SQLX$ = "SELECT  IDENLOTE, SUM(CANTINGRE - CANTSALID)  AS RESULTADO"
    SQLX$ = SQLX$ + " FROM MOVISTO WHERE COD_EXTE = '" & TRIM$(TXTCODIGOBOBINABASE) & "'"
    SQLX$ = SQLX$ + " GROUP BY IDENLOTE "
    SQLX$ = SQLX$ + " HAVING SUM(CANTINGRE - CANTSALID)  > 0"
    SQLX$ = SQLX$ + " ORDER BY IDENLOTE"

    Set Rst = FLEXCONN.Execute(SQLX$)
    Call Carga_MSF_Recordset(Rst, Campos$, FRMZELECHO.MSF2, 1)
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 4800
    FRMZELECHO.Show 1
    
    LOTE$ = TRIM(RESP_ZELE_VECT(1))
    TXTLECTURA = LOTE$
    
99 Command2.Enabled = True
End Sub

Private Sub Command3_Click()
    '
    Label4 = OrdenFabricacionAbiertas$
    
    Exit Sub
    CIXI% = CODINT%(TRIM$(Text10))
    If CIXI% <= 0 Then
        Call MENSERR(24, "Primero debe Seleccionar un Articulo")
        Exit Sub
    End If
    '
    CIRANG = CODIGORANGO&(CIXI%)
    CONDI$ = "COD_INTE = " & CIXI%
    CONDI$ = CONDI$ + " AND STATUORF < 3"
    Reg& = CantRegistros&("ORDEFABR", CONDI$)
    If Reg& < 1 Then
        Call MENSERR(24, "No Existen Ordenes de Fabricación\Para este Articulo")
        Exit Sub
    End If
    '
    Call SELECHO("!ECORDEF1/Ordenes de Fabricación")
    If VALACT1$("!ECORDEF1") <> "" Then
     Label4 = VALACT1$("!ECORDEF1")
     Label2 = VALACT2$("!ECORDEF1")
    End If
    '
End Sub




Private Sub Command8_Click()
    Call CIERRARCH("*.*")
    Unload Me
End Sub




Private Sub Command4_Click()
    Label4 = OrdenFabricacionAbiertas$(1)
End Sub

Private Sub Command5_Click()
10 Call SELECHO("MASTER")
   If VALACT1$("MASTER") = "" Then Exit Sub
   CODD$ = VALACT1$("MASTER")
   TXTCODIGOBOBINABASE = CODD$
   
End Sub

Private Sub Form_Load()
   
   DEPOCONSUMO% = 1
   Me.LBLDEPOCONS = ""
   Me.LBLDEPOENT = Control("", "DEPOCIE")
End Sub


Private Sub Label4_Change()
'   MSF1.Rows = 1
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   Me.Height = 4470
   Call BLANFORMU
   Call Centrar(Me)
   NORFA$ = TRIM$(Label4)
   If NORFA$ = "" Then Exit Sub
   Label2 = RESP_ZELE_VECT(2)
   CI1% = CODIFAB%(NORFA$)
   Text2.TEXT = PESUNI(CI1%)
   TXTCODINT = CI1%
   Text10 = CODEXT$(CI1%)
   
   '*******CAMBACERES *******+
   'SEGUN LA FAMILIA LE ASIGNA UN DEPOSITO DE CONSUMO
   CFAMILIA$ = TRIM$(UCase$(CODFAMIL$(CI1%))) 'TELAS PLANAS Y TELAS TUBULARES
   If CFAMILIA$ = "TP" Or CFAMILIA$ = "TT" Then
      LBLDEPOCONS = "5"  'deposito de consumo de extrusora
   ElseIf CFAMILIA$ = "PL" Or CFAMILIA$ = "TL" Or CFAMILIA$ = "R1" Or CFAMILIA$ = "R2" Then 'TELAS LAMINADAS, TELAS TUBULARES LAMINDAS
      LBLDEPOCONS = "6"  'deposito de consumo de laminado antes 6
      If InStr(EPAC$, "CAMBACERES") > 0 And InStr(EPAC$, "QUILMES") < 1 Then LBLDEPOCONS = "3" 'si es camba = 3
   Else
      LBLDEPOCONS = XVALO(Control("", "DEPODEST"))
   End If
   
   
   Label3 = DESCRIT0$(CI1%)
   CIFAB% = CODINT%(Text10)
   TIP% = TIMATE%(CI1%)
   
   TXTCANTIDADACERRAR.Locked = False 'por defecto editable
   If TIP% = 41 Then 'si es bobina no se edita solo se cierran de a 1
      TXTCANTIDADACERRAR = "1"
      TXTCANTIDADACERRAR.Locked = True
   End If
   NROPED& = NroPedidoSegunOf&(NORFA$)
   NC = cliente_SEGUN_PEDIDO(NROPED&)
   Text6 = NC
   Text12 = NROPED&
   CANTI = CantiProyectadaOF(NORFA$)
   Text14 = CANTI
   Text1 = NumeroOrdenPedidoSegunOrdenFabricacion%(NORFA$)
   CONDI$ = "NROPED = " & NROPED& & " AND  CODIINT = " & CIFAB%
   Text13 = XVALO(VALOBADA("PEDDETA", "SUM(CanPed)", CONDI$, "NOMESS"))
   Me.TXTPENDIENTE = CANT_OF_PENDIENTEDECIERRE(CIFAB%, NORFA$)
 
   TXTNROCOTIZA = NROCOTIZA(XVALO(NROPED&), XVALO(Text1))

End Sub


Private Sub lbldepogeneral_Change()
'   DEPO% = XVALO(lbldepogeneral)
'   Label1 = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPO%)))

End Sub

Private Sub LBLDEPOCONS_Change()
   Label10.Caption = TRIM$(ECOARCH$("TADEPOSI", MKI$(XVALO(Me.LBLDEPOCONS))))

End Sub

Private Sub LBLDEPOENT_Change()
   Label7.Caption = TRIM$(ECOARCH$("TADEPOSI", MKI$(XVALO(Me.LBLDEPOENT))))

End Sub


Private Sub mnuAnulOrofa_Click()
        'modulo correcto
        If DERACCE%("ANUCIEOF", "Anulación Cierre Parcial de O/F") >= 2 Then
          Call OFABRI07.ANUCIEOF_CMB
        End If

End Sub

Private Sub mnuLoteDespbobFab_Click()

    'lotes disponibles solo para cambaceres bobinas fabricadas
    Campos$ = "CI/A0;Código/A16;Descripción/A40;Lote/A18;Disponiblidad/F12.0;"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    J& = 0
    
    
    SQLX$ = "SELECT COD_INTE,IDENLOTE, SUM(CANTINGRE - CANTSALID)as SALDO,COD_EXTE,DESCRIP FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE IDENLOTE LIKE 'OF%'"
    SQLX$ = SQLX$ + " GROUP BY COD_INTE,IDENLOTE,COD_EXTE,DESCRIP"
    SQLX$ = SQLX$ + " Having Sum(CANTINGRE - CANTSALID) > 0"
    SQLX$ = SQLX$ + " ORDER BY COD_EXTE"
   
    Set rs = READSET(SQLX$)
    With rs
    Do While Not .EOF
        CI1% = XVALO(!cod_inte)
        Codigo$ = SAFETEXT$(!Cod_Exte)
        DESXRI$ = SAFETEXT$(!Descrip)
        LOTE$ = SAFETEXT$(!idenlote)
        SALDO = XVALO(!SALDO)
        J& = J& + 1
        GRILLAGRAL.GRID.Rows = J& + 1
        GRILLAGRAL.GRID.TextMatrix(J&, 1) = CI1%
        GRILLAGRAL.GRID.TextMatrix(J&, 2) = Codigo$
        GRILLAGRAL.GRID.TextMatrix(J&, 3) = DESXRI$
        GRILLAGRAL.GRID.TextMatrix(J&, 4) = LOTE$
        GRILLAGRAL.GRID.TextMatrix(J&, 5) = FORMATNUM$(SALDO, "F10.1")
       .MoveNext
    Loop
   End With
   rs.Close
   Set rs = Nothing
   '
   GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
   GRILLAGRAL.mnuMenu1.Visible = True
   GRILLAGRAL.GRID.ZOrder 0
   GRILLAGRAL.TXTBUSCAR = ""
   GRILLAGRAL.SinEfectoClickEnGrilla% = 1
   GRILLAGRAL.Caption = "Bobinas Fabricadas Disponibles "
   GRILLAGRAL.Show 1
   
   Dim MSF As msFlexGrid
   If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "BOFABDIS", "Bobinas Fabricadas Disponibles", "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
   End If

End Sub





Private Sub mnuMarcadeCierre_Click()
   If DERACCE%("CIEMASOF", "Cierre Automático de Ordenes de Fabricación") < 2 Then Exit Sub
   CIEMASOF.Command3.Visible = False
   CIEMASOF.BTNMARCACIERREOF.Visible = True
   CIEMASOF.Frame2.Visible = False
   CIEMASOF.Frame3.Visible = True
   CIEMASOF.Caption = "Marca Masiva de Cierre de Ordenes de Fabricación Seleccionadas"
   CIEMASOF.Show 1

End Sub


Private Sub mnuOrigScrap_Click()
   FRMSCRAP.Show 1
End Sub

Private Sub mnuReaperturaOF_Click()
    
    If DERACCE%("REAPEROF", "Re-Apertura O/F's con Saldo") >= 2 Then
      Call OFABRI07.REAPEROF
      GoTo 99
    End If
99
End Sub

Private Sub mnuTabAux_Click()
  Call GENERAR_ESTRUCTURAS_SQL
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set rs1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(rs1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.MSF2)
  Set obj = Nothing
  rs1.Close
  Set rs1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  '

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  IDLIST& = XVALO(RESP_ZELE_VECT(1))
  If InStr(EPAC$, "ENVAPLAST") > 0 Then
    LIXTA$ = TRIM$(RESP_ZELE_VECT(2))
    If LIXTA$ = "MATERIALES ENVAPLAST" Then
       Call COMUNI("Atención !!! Editar Tabla de Familias Desde Maestro de Artículos")
       Exit Sub
    End If
  End If

  Call ABM_TABLA(IDLIST&)
End Sub


Private Sub mnuTablaConfScrap_Click()
      CIERRECTP.mnuTablaConfScrap_Click
End Sub

Private Sub Text10_Change()
   TXTCODIGOBOBINABASE = ""

   CI1% = CODINT%(TRIM$(Text10))
   If CI1% > 0 Then
     CF$ = TRIM$(UCase$(CODFAMIL$(CI1%)))
     If CF$ = "R1" Or CF$ = "R2" Then  'CASO DE TELAS PLANAS O LAMINADAS -REFILADAS-
        CONDIFAMILIA$ = "FAMILIA = 'TL' OR FAMILIA = 'PL'"
     Else 'CASO TELAS LAMINADAS TUBULARES O PLANAS
        CONDIFAMILIA$ = "FAMILIA = 'TP' OR FAMILIA = 'TT'"
     End If
     CI2% = ComponenteSegunCodiConjunYFamilia%(CI1%, CONDIFAMILIA$)
     If CI2% > 0 Then
        TXTCODIGOBOBINABASE = CODEXT$(CI2%)
     End If
   End If
End Sub

Private Sub Text2_LostFocus()
   Text2.TEXT = TRIM$(FORMATNUM$(XVALO(Text2), "f12.2"))
End Sub

Private Sub TEXT3_Change()

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
     LecturaEtiqueta$ = TRIM$(TXTLECTURA)
     LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, "'", "-")
     LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, "]", "|")
     If LecturaEtiqueta$ = "" Then Exit Sub

End Sub


Private Sub Text6_Change()
   Label21 = rasocli$(XVALO(Text6))
End Sub


Private Sub TXTLECTURA_Change()
   CIBobinaBase% = CodigoIntSegunLoteDelMovisto(TRIM$(TXTLECTURA))
   If CIBobinaBase% > 0 Then
      TXTCODIGOBOBINABASE = CODEXT$(CIBobinaBase%)
      CONDI$ = "IDENLOTE =  '" & TRIM$(TXTLECTURA) & "' AND COD_INTE = " & CIBobinaBase% & " AND (CODIMOVI = 'OF' OR CODIMOVI = 'BR')"
      TXTKBB.TEXT = TRIM$(FORMATNUM$(XVALO(VALOBADA("MOVISTO", "PESOKG", CONDI$, "NOMESS")), "F12.2"))
      TXTMBB.TEXT = TRIM$(FORMATNUM$(XVALO(VALOBADA("MOVISTO", "METROS", CONDI$, "NOMESS")), "F12.2"))

   Else
      TXTCODIGOBOBINABASE = ""
      Me.TXTKBB = ""
      Me.TXTMBB = ""
   End If

End Sub

Private Sub TXTLECTURA_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        LecturaEtiqueta$ = TRIM$(TXTLECTURA)
        LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, "'", "-")
        LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, "]", "|")
        If LecturaEtiqueta$ = "" Then Exit Sub
        
        Dim strArray() As String
        LECTURA$ = LecturaEtiqueta$
        On Error Resume Next
        strArray = Split(LECTURA$, "|")
        Codigo$ = strArray(0)
        NORFA$ = TRIM$(strArray(1))
        LOTE$ = TRIM$(strArray(2))
        On Error GoTo 0
     End If
End Sub



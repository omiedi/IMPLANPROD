VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CIERRECTP 
   BackColor       =   &H009FBE85&
   Caption         =   "Informar Cierres de Orden de Fabricación "
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   13695
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   13695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H009FBE85&
      Height          =   975
      Left            =   10800
      TabIndex        =   25
      Top             =   2220
      Width           =   2775
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
         Height          =   600
         Left            =   2040
         Picture         =   "CIERRECTP.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Genera Transacción de Cierre de Orden de Fabricación Activa"
         Top             =   200
         Width           =   650
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
         Left            =   600
         TabIndex        =   26
         Text            =   " "
         Top             =   360
         Width           =   1320
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Cant. a Cerrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   27
         Top             =   120
         Width           =   2085
      End
   End
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
      TabIndex        =   49
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
      TabIndex        =   46
      Text            =   " "
      Top             =   3480
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
      TabIndex        =   45
      Text            =   " "
      ToolTipText     =   "ID (Tabla Movisto)"
      Top             =   3840
      Width           =   945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Situación de Cierre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   1500
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H009FBE85&
      Height          =   1400
      Left            =   3960
      TabIndex        =   22
      Top             =   1800
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
         TabIndex        =   59
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Lee el Peso de La Balanza"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Metros:"
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
         Left            =   720
         TabIndex        =   60
         Top             =   960
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
      Left            =   12000
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H009FBE85&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Index           =   0
      Left            =   50
      TabIndex        =   2
      Top             =   0
      Width           =   13590
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1310
         Left            =   7350
         MaxLength       =   128
         MultiLine       =   -1  'True
         TabIndex        =   61
         Top             =   1890
         Width           =   3285
      End
      Begin VB.TextBox TXTIDMOVTALLER 
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
         Left            =   12200
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   " "
         Top             =   960
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cierre Final bolsas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   3360
         Picture         =   "CIERRECTP.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Genera Transacción de Cierre de Orden de Fabricación Activa"
         Top             =   960
         Width           =   1845
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
         TabIndex        =   43
         ToolTipText     =   "Ordenes de Fabricación (Todas)"
         Top             =   240
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
         Height          =   300
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   " "
         Top             =   1320
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
         TabIndex        =   38
         Text            =   " "
         Top             =   960
         Width           =   720
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H009FBE85&
         Height          =   1400
         Left            =   45
         TabIndex        =   29
         Top             =   1800
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
            TabIndex        =   34
            Top             =   960
            Visible         =   0   'False
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
            TabIndex        =   30
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   35
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
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
         Top             =   960
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
         Top             =   960
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
         Height          =   300
         Left            =   12800
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   " "
         Top             =   600
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
         Top             =   600
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
         Top             =   240
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
         Top             =   240
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
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Ingresar Código de Producto Terminado"
         Top             =   600
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Anexo:"
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
         Left            =   7320
         TabIndex        =   62
         Top             =   1560
         Width           =   765
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ID MOVTALLER"
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
         Left            =   10440
         TabIndex        =   58
         Top             =   960
         Visible         =   0   'False
         Width           =   1485
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
         Left            =   7560
         TabIndex        =   41
         Top             =   1320
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
         TabIndex        =   39
         Top             =   960
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
         Top             =   960
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
         Top             =   960
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
         Left            =   11145
         TabIndex        =   17
         Top             =   600
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
         Top             =   240
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
         Top             =   600
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
         Top             =   300
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
         Top             =   600
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
         Top             =   240
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
         Top             =   675
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
         Top             =   300
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
         Top             =   240
         Width           =   4770
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "CIERRECTP.frx":0614
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CIERRECTP.frx":0848
      TabIndex        =   9
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   1650
      Left            =   0
      TabIndex        =   44
      Top             =   3480
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
      TabIndex        =   51
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
      TabIndex        =   52
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
   Begin VB.Label Label1 
      BackColor       =   &H009FBE85&
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
      TabIndex        =   55
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
      TabIndex        =   54
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
      TabIndex        =   53
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
      TabIndex        =   50
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
      TabIndex        =   48
      Top             =   3555
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
      TabIndex        =   47
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
      Begin VB.Menu mnuMatProceso 
         Caption         =   "Materiales en Proceso"
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
      Begin VB.Menu mnuDepFabCerradas 
         Caption         =   "Depuraciòn de Consumos de Orden de Fabricaciones Cerradas"
      End
      Begin VB.Menu mnuSaldoDepoCorte 
         Caption         =   "Saldos en Depòsito de Corte"
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
Attribute VB_Name = "CIERRECTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Function EnSector(CI1%, DESIDSECTOR%, HASSECTOR%)
    SQLX$ = "SELECT CodiElem,IDCOORDUBIC,SUM(CANTREAL * CORTE )AS CANTI, SUM(CONSUMO) AS CONSU,COORDUBIC FROM INFOFAB I"
    SQLX$ = SQLX$ + " inner join MASTER MAS WITH(NOLOCK) ON MAS.Codint = I.CodiElem"
    SQLX$ = SQLX$ + " INNER  join ORDEFABR O ON I.IdSubOr = O.IDSUBOR"
    SQLX$ = SQLX$ + "  Where CODIELEM = " & CI1%
    SQLX$ = SQLX$ + " And MAS.TIMATE <> 49 And MAS.Status >= 0 And O.StatuOrf < 3"
    SQLX$ = SQLX$ + "  AND (I.IdRegPadre < 1 OR IdRegPadre IS NULL) "
    If DESIDSECTOR% = 30 And HASSECTOR% = 50 Then 'TALLERES
        SQLX$ = SQLX$ + " AND substring(i.COORDUBIC,3,1) = ';'"
        SQLX$ = SQLX$ + " AND convert(int,(left(I.COORDUBIC,2))) >= 30"
        SQLX$ = SQLX$ + " AND convert(int,(left(I.COORDUBIC,2))) <= 50"
    Else
        SQLX$ = SQLX$ + "  AND IDCOORDUBIC >= " & DESIDSECTOR%
        SQLX$ = SQLX$ + "  AND IDCOORDUBIC <= " & HASSECTOR%
    End If
    SQLX$ = SQLX$ + " AND (I.cantreal - I.CONSUMO) > 0"
    SQLX$ = SQLX$ + " GROUP BY IDCOORDUBIC,CodiElem,COORDUBIC"
    Set Rst = READSET(SQLX$)
    EnSector = XVALO(Rst!CANTI) - XVALO(Rst!CONSU)
    If EnSector < 0 Then EnSector = 0
    Rst.Close
    Set Rst = Nothing
    
End Function

Sub IgualarCantidadConsumidaConConsumoInfofab(NORFA$)
   'EN ESTE CASO ACTUALIZA TODA LA OF EN EL INFOFAB CONSUMO = CANTREAL
       SQLX$ = "UPDATE INFOFAB SET CONSUMO = CANTREAL WHERE  CANTREAL > CONSUMO AND IDSUBOR = '" & NORFA$ & "'"
       FLEXCONN.Execute (SQLX$)
       

End Sub

Function validarOrdenFabricacionCortes(NORFA$, CANTI)
    'PARA QUE ESTO FUNCIONE CORRECTAMENTE DEBE TENER EL CORTE EN LA FORMULA SOLO UNA TELA
    ORFA$ = Mid$(NORFA$, 2)
    ORFA$ = "%" & ORFA$
    SALDOTRANSFERIDO = 0
    'EL SALDO DEL COMPONENTE TELA DEBEN SER METROS
    SQLX$ = "SELECT SUM(CANTINGRE-CANTSALID) AS SALDO  FROM MOVISTO WITH(NOLOCK)  WHERE  DOSECLAR LIKE '" & ORFA$ & "'"
    SQLX$ = SQLX$ + " AND DEPOMOVI > 1 AND CODIMOVI <> 'OF'"
    Set rs = READSET(SQLX$)
    If Not rs.EOF Then
      SALDOTRANSFERIDO = XVALO(rs!SALDO)
    End If
    '
    'AHORA CALCULO LA NECESIDAD
    CIFAB% = CODIFAB(NORFA$)
    NECESIDAD = 0
    USOBRUTOX = XVALO(VALOBADA("FORMULAS", "USOBRUTO", "CodIConj=" & CIFAB%, "NOMESS"))
'    SQLX$ = "SELECT (usobruto) AS NECESI FROM FORMULAS F WITH(NOLOCK) INNER JOIN MASTER M ON F.CODIELEM  = M.CODINT"
'    SQLX$ = SQLX$ + " WHERE CodIConj=" & CIFAB%
'    SQLX$ = SQLX$ + " AND M.UMEDIDA = 'M.'"
'    Set RstFormula = READSET(SQLX$)
'    J& = 0
'    With RstFormula
'        If Not (.EOF And .BOF) Then
'           NECESIDAD = XVALO(!NECESI)
'        End If
'    End With
'    RstFormula.Close
'    Set RstFormula = Nothing
    '
    If USOBRUTOX > 0 Then
       Resultado = SALDOTRANSFERIDO / USOBRUTOX
    End If
    'EJEMPLO SI EL SALDO ES DE 100  MTS Y EL USO ES DE 1 M , RESULTADO VA A SER 100
    'SI LA CANTIDAD EN U. DE CORTES ES <= 100 DA POSITIVO SI ES MAYOR DA NEGATIVO
    validarOrdenFabricacionCortes = Resultado - CANTI
End Function

Private Sub mnuAnulOrofa_Click()
        'modulo correcto
        If DERACCE%("ANUCIEOF", "Anulación Cierre Parcial de O/F") >= 2 Then
          Call OFABRI07.ANUCIEOF_CMB
        End If

End Sub


Private Sub LBLDEPOCONS_DblClick()
    Call SELECHO("ECODEP")
    DPX1% = Val(VALACT1$("ECODEP"))
    If DPX1% > 0 Then
        DP1% = DPX1%
        LBLDEPOCONS = TRIM$(str$(DP1%))
    End If

End Sub

Private Sub LBLDEPOENT_DblClick()
    Call SELECHO("ECODEP")
    DPX1% = Val(VALACT1$("ECODEP"))
    If DPX1% > 0 Then
        DP1% = DPX1%
        LBLDEPOENT = TRIM$(str$(DP1%))
    End If

End Sub

Private Sub mnuReaperturaOF_Click()
    
    If DERACCE%("REAPEROF", "Re-Apertura O/F's con Saldo") >= 2 Then
      Call OFABRI07.REAPEROF(1)
      GoTo 99
    End If
99
End Sub
Sub IgualarCantidadConsumidaConConsumoInfofabConteplast(MSF As msFlexGrid)
   
   For i& = 1 To MSF.Rows - 1
      If MSF.TextMatrix(i&, 0) = "*" Then
         CI1% = XVALO(MSF.TextMatrix(i&, 1))
         NORFA$ = TRIM$(MSF.TextMatrix(i&, 4))
         Call IgualarCantidadConsumidaConConsumoInfofab(NORFA$)
      End If
   Next i&
End Sub
Sub AjustarSegunOf(MSF As msFlexGrid)
   
   EPAC$ = TRIM$(UCase$(EMPREAC$))
    
   If InStr(EPAC$, "CONTEPLAST") > 0 Then DEPO% = 2
   For i& = 1 To MSF.Rows - 1
      If MSF.TextMatrix(i&, 0) = "*" Then
        CI1% = XVALO(MSF.TextMatrix(i&, 1))
         
        SQLX$ = "SELECT M.COD_INTE,M.COD_EXTE,M.DESCRIP,SUM(CANTINGRE-CANTSALID) AS RESULTADO ,O.IDSUBOR"
        SQLX$ = SQLX$ + " FROM MOVISTO M WITH(NOLOCK)"
        SQLX$ = SQLX$ + " INNER JOIN ORDEFABR O ON (O.IdSubOr  = M.DoSecLar     OR O.IDSUBOR = M.DOPRICOR)"
        'SQLX$ = SQLX$ + " where  (left(DOSECLAR,3) = 'OF-'  or  left(DOPRICOR,3)   = 'OF-' ) AND DEPOMOVI = 2"
        SQLX$ = SQLX$ + " where (DOSECLAR LIKE '_F-______-1'  OR DOSECLAR LIKE 'AJ.%') AND DEPOMOVI = 2"
        SQLX$ = SQLX$ + " AND O.StatuOrf = 3"
        
        SQLX$ = SQLX$ + " AND M.Cod_Inte  = " & CI1%
        SQLX$ = SQLX$ + " GROUP BY M.COD_INTE,M.COD_EXTE,M.DESCRIP,O.IDSUBOR "
        SQLX$ = SQLX$ + " HAVING SUM(CANTINGRE-CANTSALID) <> 0"
        SQLX$ = SQLX$ + " ORDER BY M.COD_EXTE"
        Set Rst = FLEXCONN.Execute(SQLX$)
        With Rst
            Do While Not .EOF
                NORFA$ = SAFETEXT$(!idsubor)
                Call AjusteA0SegunCodigoOfabricacionDeposito(CI1%, NORFA$, DEPO%, HOY(HO$))
                .MoveNext
            Loop
        End With
                
        Rst.Close
        Set Rst = Nothing
        On Error GoTo 0
      End If
   Next i&
   
End Sub
Sub PrintEtiqProdTerminado()
        Dim CTP As New Conteplast
        NROPED& = XVALO(Text12)
        Codigo$ = TRIM$(Text10)
        CIFAB% = CODINT(Codigo$) 'PRODUCTO TERMINADO
        NROOF$ = Label4
        CANTI = XVALO(TXTCANTIDADACERRAR)
        NC = XVALO(Text6)
        CODIBIGBAG$ = TRIM$(CTP.CODIGOBIGBAG(CIFAB%))
        If CODIBIGBAG$ = "" Then CODIBIGBAG$ = "-"
        Call PRINTETIQPT(CODEXT$(CIFAB%), NROOF$, CANTI, NC, NROPED&, CIFAB%, CODIBIGBAG$)
        
End Sub

Sub PRINTETIQPT(Codigo$, NROOF$, CANTI, NC, NROPED&, CIFAB%, CODIBIGBAG$)
     '
     FORIPRE$ = TRIM$(Control("ETIQPRTE", "ETPROTER"))
     If FORIPRE$ = "" Then
       Call COMUNI("Falta Configurar Etiqueta de Cierre de Orden de Fabricación\Configurar Formulario desde la Ficha Técnica > Congiguración General")
       Exit Sub
     End If
      
     IMAG$ = "QR_ETIQPTERM.JPG"
     IDENTIFICADOR$ = "QR_ETIQPTERM"

     On Error Resume Next
     Kill "C:\FLEXOFT\" & IMAG$
     TEXTOQR$ = Codigo$ & "|" & NROOF$ & "|" & CANTI & "|" & NC & "|" & CIFAB%

     On Error GoTo 0
     'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
     QRX$ = TRIM$(IMAGENCODBARRA$("C:\FLEXOFT\", IMAG$, TEXTOQR$, "QR_CODE", ANCHOX%, ALTOX%))
     Call T_Espera(3)
'     Image1.Picture = LoadPicture("C:\FLEXOFT\" & IMAG$)
     Call GRABAR_QR("C:\FLEXOFT\" & IMAG$, IDENTIFICADOR$)
     '
     FILTX$ = TRIM$(str$(NC)) & ";" & NROOF$ & ";" & TRIM$(str$(CANTI)) & ";" & TRIM$(Codigo$) & ";" & TRIM$(CODIBIGBAG$)
     Call STDFORM(FORIPRE$, FILTX$)

End Sub


Sub RegistrarFotoCierre(NORFA$, CANTI, RetIdentificador)
    CIOF% = CODIFAB%(NORFA$)
    CF$ = TRIM$(UCase$(CODFAMIL$(CIOF%)))
    If CF$ = "BO" Or CF$ = "CT" Or CF$ = "SA" Then
       'SI ES BOSA O CORTES SE VA POR QUE HABRIA QUE MODIFICAR DE DONDE LEE LOS DATOS DE LA PANTALLA DE DETACIERRECTP
       'POR QUE LOS DATOS SON DIFERENTES.
       Exit Sub
    End If
    If IsFormLoaded(DetaCierreCtp) Then
       Unload DetaCierreCtp 'CIERRO EL FORMULARIO PARA QUE SE PRODUZCAN LOS EVENTOS CHANGE SI NO CAMBIAN LOS DATOS
    End If

    DetaCierreCtp.TXTACERRAR = CANTI
   'ESTE BLOQUE SIEMPRE AL FINAL
    DetaCierreCtp.Text10 = NORFA$
    'DetaCierreCtp.Show 1
    Call CREATABLA_FOTODETACIERRE
    
'    Campos$ = "Código/A18;N.Op./f5;Desc.Oper/A24;CAnt.Cortes Originales/f22;Fact.Corte/F12;CORTE/F10.2;Lote/a16;Peso/F10.2;Cod.Rango/A12;Saldo/F10.2"
'    Call CARGA_ENCABEZADO(MSF2, Campos$, Me)
    
    SQLX$ = "SELECT * FROM FOTODETACIERRE2"
    SQLX$ = SQLX$ + " WHERE FOTODETACIERRE2_ID < 1"
20  RNDX1$ = RANDSTRING$(-10)
    CONDI$ = "IDMOVISTO =" & XVALO(RNDX1$)
    If CantRegistros("FOTODETACIERRE2", CONDI$, "NOMESS") > 0 Then GoTo 20
    If CantRegistros("FOTODETACIERRE3", CONDI$, "NOMESS") > 0 Then GoTo 20
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With rs
        For Reg& = 1 To DetaCierreCtp.MSF2.Rows - 1
                .AddNew
                !CodXElem = DetaCierreCtp.MSF2.TextMatrix(Reg&, 1)
                !NUMEOPER = XVALO(DetaCierreCtp.MSF2.TextMatrix(Reg&, 2))
                !descoper = DetaCierreCtp.MSF2.TextMatrix(Reg&, 3)
                !CantCortes = XVALO(DetaCierreCtp.MSF2.TextMatrix(Reg&, 4))
                !FACTORCORTE = XVALO(DetaCierreCtp.MSF2.TextMatrix(Reg&, 5))
                !corte = XVALO(DetaCierreCtp.MSF2.TextMatrix(Reg&, 6))
                !idenlote = DetaCierreCtp.MSF2.TextMatrix(Reg&, 7)
                !Peso = XVALO(DetaCierreCtp.MSF2.TextMatrix(Reg&, 8))
                !CODRANGO = DetaCierreCtp.MSF2.TextMatrix(Reg&, 9)
                !SALDO = XVALO(DetaCierreCtp.MSF2.TextMatrix(Reg&, 10))
                !IDMOVISTO = XVALO(RNDX1$)
                .Update
        Next Reg&
    End With
    rs.Close
    Set rs = Nothing
    '
'    Campos$ = "Código/A16 '  1"
'    Campos$ = Campos$ + ";Descripción/a48"             '  2"
'    Campos$ = Campos$ + ";Componente/A24"               '  3"
'    Campos$ = Campos$ + ";Cielem/A0"                    '  4"
'    Campos$ = Campos$ + ";Moduxb/A0"                   '  5"
'    Campos$ = Campos$ + ";CORTE/A0"                  '  6"
'    Campos$ = Campos$ + ";Nec.M./A0"                       '  7"
'    Campos$ = Campos$ + ";Cod.Rango/A12"                '  8"
'    Campos$ = Campos$ + ";Nec.Cortes/f10.1"            '  9
'    Campos$ = Campos$ + ";Disponibilidad/f14"              '  10
'    Campos$ = Campos$ + ";Falta/f10"                       '  11
'
'    Call CARGA_ENCABEZADO(MSF3, Campos$, Me)
    SQLX$ = "SELECT * FROM FOTODETACIERRE3"
    SQLX$ = SQLX$ + " WHERE FOTODETACIERRE3_ID < 1"
    Dim rs3 As ADODB.Recordset
    Set rs3 = New ADODB.Recordset
    rs3.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With rs3
        For Reg& = 1 To DetaCierreCtp.MSF3.Rows - 1
                .AddNew
                !Codigo = DetaCierreCtp.MSF3.TextMatrix(Reg&, 1)
                
                !Descripcion = Left$(DetaCierreCtp.MSF3.TextMatrix(Reg&, 2), 64)
                !COMPONENTE = DetaCierreCtp.MSF3.TextMatrix(Reg&, 3)
                !CODINTELEM = XVALO(DetaCierreCtp.MSF3.TextMatrix(Reg&, 4))
                !moduxbb = XVALO(DetaCierreCtp.MSF3.TextMatrix(Reg&, 5))
                !corte = XVALO(DetaCierreCtp.MSF3.TextMatrix(Reg&, 6))
                !NECES_MTS = XVALO(DetaCierreCtp.MSF3.TextMatrix(Reg&, 7))
                !CODRANGO = DetaCierreCtp.MSF3.TextMatrix(Reg&, 8)
                !NECESIDAD = XVALO(DetaCierreCtp.MSF3.TextMatrix(Reg&, 9))
                !DISPONIBILIDAD = XVALO(DetaCierreCtp.MSF3.TextMatrix(Reg&, 10))
                !FALTA = XVALO(DetaCierreCtp.MSF3.TextMatrix(Reg&, 11))
                !IDMOVISTO = XVALO(RNDX1$)
                .Update

        Next Reg&
    End With
    rs3.Close
    Set rs = Nothing
    RetIdentificador = XVALO(RNDX1$)
    
    
End Sub

Sub UBICARPROCUTOTERMINADO(NORFA$, CANTI, NC, NP&, NUORDENPED%, FEENT%, Optional IDREGMOVISTO&)
   INFUBIPT.Label4 = NORFA$
   
   Call INFUBIPT.CARGARCASILLEROS(INFUBIPT.MSF3, 1)
   Call SELECCIONARFILA(INFUBIPT.MSF3, 1)
   INFUBIPT.Show 1
End Sub

Function ValidacionOcultaDeBolsasCTP%(NORFA$, CantiaCerrar)
    DetaCierreCtp.TXTACERRAR = XVALO(CantiaCerrar)
    ValidacionOcultaDeBolsasCTP% = 0
    'ESTE BLOQUE SIEMPRE AL FINAL
    DetaCierreCtp.Text10 = NORFA$
    For i& = 1 To DetaCierreCtp.MSF3.Rows - 1
       If XVALO(DetaCierreCtp.MSF3.TextMatrix(i&, 5)) > 0 Then 'columna Faltan
            ValidacionOcultaDeBolsasCTP% = (-1)
            XXX$ = "O.F.: " & NORFA$
            XXX$ = XXX$ & " - Cierre x " & CantiaCerrar
            Call COMUNI(XXX$ & "\ Faltan Transferir Material")
            Exit For
       End If
    Next i&
End Function

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
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    CIFAB% = CODINT%(Text10)
    If CIFAB% < 1 Then
       Call COMUNI("Código No Válido")
       GoTo 99
    End If

    NORFA$ = TRIM(Label4)
    If TRIM(NORFA$) = "" Then
       Call COMUNI("Orden de Fabricación No Válida")
       GoTo 99
    End If
    '
    If STATORFA%(NORFA$) > 1 Then
       Call COMUNI("Imposible Cerrar, Orden de Fabricación Seleccionada Se Encuentra Cerrada")
       GoTo 99
    End If
    
    CODFAMI$ = TRIM$(CODFAMIL$(CIFAB%))
    
    CANTI = XVALO(TXTCANTIDADACERRAR)
    If CANTI < 1 Then
       Call COMUNI("Cantidad Ingresada No Válida")
       Call PINTATEXT(TXTCANTIDADACERRAR)
       GoTo 99
    End If
    '
    If MaximoPosibleACerrar&(CIFAB%, NORFA$) < CANTI Then
    
       MJE$ = "La Cantidad Ingresada: " & CANTI & " +  La Cantidad Yá Cerrada\Supera la Cantidad Proyectada + La Tolerancia Permitida"
       OPX% = COMALTER%(MJE$ & "\\Cancelar Cierre\Generar Cierre (Con Derechos)")
       If OPX% = 2 Then
          If FIRMAELECTRONICAOK% < 1 Then
             GoTo 50 'SI NO APRUEBA FIRMA ELECT... NO CONTINUA
          Else
             Call COMUNI("Firma Electrónica Aprobada")
          End If
       End If
       
       If OPX% < 2 Then
50        TXTPENDIENTE = CANT_OF_PENDIENTEDECIERRE(CIFAB%, NORFA$) 'actualizo por las dudas por si se realizo un cierre desde otra terminal
          Call PINTATEXT(TXTCANTIDADACERRAR)
          GoTo 99
       End If
    End If
    
    
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
    '
    If InStr(EPAC$, "QUILMES") > 0 Then 'SI LA UNIDAD DE MEDIDA SON METROS TEXT2 (METROS) DEBE SER IGUAL A LA CANTIDAD A CERRAR
      If UCase$(Unimed$(CIFAB%)) = "M." Then
         Text2 = TXTCANTIDADACERRAR
      End If
    End If

    
    Dim CTP As New Conteplast
    'ACA CONSIDERA VALIDCION PARA BOLSONES
    Call CREARTABLASOLICOTI
    If CantRegistros&("SOLICOTI", "CODIMASTER = '" & TRIM$(Text10) & "'", "NOMESS") > 0 Then
        If CTP.ValidarMatDeConsumoTransferidoALinea(CIFAB%, XVALO(TXTNROCOTIZA), CANTI, NORFA$) < 0 Then
            OPX% = COMALTER%("Atención !!!! Al Cerrar ésta Orden, Se Registrará en la Auditoría del Sistema\Las Alertas Presentadas\Responsabilizando al Usuario Activo\\Cancelar Cierre\Generar Cierre Sin Respetar Alertas")
            If OPX% < 2 Then GoTo 99
        End If
    End If
    '
    
    If InStr(EPAC$, "CONTEPLAST") > 0 Then
        If Left(TRIM$(Text10), 3) = "CTP" Then    'VALIDACION PARA BOLSAS Y CORTES (solicito Lautaro 09/08/24)  Or InStr(CODFAMI$, "CT") > 0
          If ValidacionOcultaDeBolsasCTP(NORFA$, CANTI) < 0 Then
             OPX% = COMALTER%("Atención !!!! Al Cerrar ésta Orden, Se Registrará en la Auditoría del Sistema\Las Alertas Presentadas\Responsabilizando al Usuario Activo\\Cancelar Cierre\Generar Cierre Sin Respetar Alertas")
             If OPX% < 2 Then GoTo 99
          End If
        End If
    End If
    
    FORIPRE$ = TRIM$(Control("ETIQPRTE", "ETPROTER"))
     If FORIPRE$ = "" Then
       Call COMUNI("Falta Configurar Etiqueta de Cierre de Orden de Fabricación\Configurar Formulario desde la Ficha Técnica > Configuración General")
       GoTo 99
     End If

     Me.Height = 4200
     Call Centrar(Me)
     If InStr(CODFAMI$, "CT") > 0 Then
         If validarOrdenFabricacionCortes(NORFA$, CANTI) < 0 Then
            TXXK$ = "No hay Material Transferido Suficiente"
            OPX% = COMALTER%("" & TXXK$ & "\\Canelar\Cerrar Igual")
            If OPX% < 2 Then GoTo 99 'si cancela se va
         End If
     End If
'TODO ESTTE BLOQUE IF PARA QUILMES ANTES NO ESTABA SE AGREGO ESPECIALMENTE
    If InStr(EPAC$, "QUILMES") > 0 Then
      'If InStr(CODFAMI$, "TP") > 0 Or InStr(CODFAMI$, "TT") > 0 Then
      If InStr(CODFAMI$, "BO") < 1 Then 'CAMBIO REALIZADO EL 14 02 24
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
        On Error Resume Next
        coeficienteSegunMaestro = PesoDelMaestro / MetrosDelMaestro
        coeficienteSegunDatosPantalla = pesoReal / metrosReal
        resultadoDiferenciaAbs = Abs(coeficienteSegunMaestro - coeficienteSegunDatosPantalla)
        resultadoSoportado = coeficienteSegunMaestro * 1.05 'aca cambie a 5% en vez de 1.15 22-04-22
        On Error GoTo 0
        If (coeficienteSegunMaestro + resultadoDiferenciaAbs) > resultadoSoportado Then
    '    If (Abs(coeficienteSegunMaestro) * 1.15) < coeficienteSegunDatosPantalla Then
            TXXK$ = "El Peso Registrado Excede La tolerancia de 5% Dada Por el Coeficiente Resultante\ de la Configuración del Maestro Vs. Los Datos en Pantalla\O.F.: " & Label4 & " Peso Kg.: " & LblKgImpresos & " Metros: " & Text2
            OPX% = COMALTER%("" & TXXK$ & "\\Canelar\Cerrar Igual")
            If OPX% < 2 Then GoTo 99 'si cancela se va
            'si continua valida derechos
'            If COTICONT.FIRMAELECTRONICAOK% <= 0 Then
'                GoTo 99
'            End If
        End If
        
      End If
      
    End If
    
    
    'MODO CIERREFACIL
    FEC$ = FECHATEX$(HOY(HO$))
    STG$ = NORFA$ & "|"
    STG$ = STG$ & CANTI & "|"
    STG$ = STG$ & XVALO(Me.LBLDEPOENT) & "|"
    STG$ = STG$ & XVALO(Me.LBLDEPOCONS) & "|"
    STG$ = STG$ & "0"
    Call ACTUREGISTRO("ORDEFABR", "ANEXOS", "IDSUBOR = '" & NORFA$ & "'", Text3, REGAF&, "NOMESS")
    CIERFACIL.STRINGORFA = STG$
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans
    If InStr(EPAC$, "CONTEPLAST") > 0 Then
       Call RegistrarFotoCierre(NORFA$, CANTI, RetIdentificador) 'foto de sitacion de cierre,Devuelve un identific Random
    End If
    'GENERO EL CIERRE DE OF Y RECUPERO EL ID CREADO
    CODFAMI$ = TRIM$(CODFAMIL$(CIFAB%))
    NoConsumirComponentes% = 0
    
    If (InStr(CODFAMI$, "BB") > 0 Or InStr(CODFAMI$, "EE") Or InStr(CODFAMI$, "CC")) And InStr(EPAC$, "CONTEPLAST") > 0 Then 'ESTA ES A PARTIR DE LA CUAL SE VA CONSIDRERAR DIFERENTE LA DEVOLUCION DE LA BOBINA
       NoConsumirComponentes% = 1
    End If
    If InStr(CODFAMI$, "CT") > 0 And InStr(EPAC$, "CONTEPLAST") > 0 Then NoConsumirComponentes% = 1
'ATENCION !!! LA LINEA SIGGUIENTE Call AgregarKgMovisto( SIEMPRE DEBE ESTAR DEBAJO DE ORFCIE SIN OTRA TRANSACCION EN MEDIO
    Call ORFCIE("", 0, "", NoConsumirComponentes%)
    
    Call AgregarKgMovisto(NORFA$, CANTI, XVALO(LblKgImpresos), Ret_idMov&) 'DEVUELVE EN Ret_idMov& EL ID GENERADO EN MOVISTO
'FIN BLOQUE QUE DEBE ESTAR SIEMPRE UNO DEBAJO DE OTRO.

   If InStr(CODFAMI$, "CT") > 0 And InStr(EPAC$, "CONTEPLAST") > 0 Then
       Call GENERACONSUMO_TELAS_CONTEPLAST(NORFA$, CIFAB%, CANTI, XVALO(Me.LBLDEPOCONS))
       GoTo 80
    End If
'    On Error GoTo ERROR_GUARDAR
'    FLEXCONN.BeginTrans

    
    If XVALO(TXTIDMOVTALLER) > 0 Then 'aca tiene el movtaller_id
        SQLX$ = "UPDATE MOVTALLER SET IDMOVISTO = " & XVALO(Ret_idMov&)
        SQLX$ = SQLX$ + " WHERE MOVTALLER_ID = " & XVALO(TXTIDMOVTALLER)
        FLEXCONN.Execute (SQLX$)
    End If
    '

    '''quilmes
    If InStr(EPAC$, "QUILMES") > 0 Then 'MODIF.25_01_24 SI SON ESTAS FAMILIAS IMPRIME LA ETIQUETA IGUAL QUE RIVADAVIA
    
        CODFAMI$ = TRIM$(CODFAMIL$(CIFAB%))
        
        If InStr(CODFAMI$, "R1") > 0 Or InStr(CODFAMI$, "R2") > 0 Or InStr(CODFAMI$, "TP") > 0 Or InStr(CODFAMI$, "TT") > 0 Or InStr(CODFAMI$, "PL") > 0 Or InStr(CODFAMI$, "TL") > 0 Or InStr(CODFAMI$, "EE") > 0 Or InStr(CODFAMI$, "CT") > 0 Then
    
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
            CONDI$ = "MOVISTO_ID = " & IdentifMovisto&
            'metrosReal = XVALO(PESUNI(CIFAB%)) ' LOS METROS CONFIGURADOS EN EL MAESTRO
            SQLX$ = "UPDATE MOVISTO SET METROS = " & FORMATNUM$(metrosReal, "F12.2")
            SQLX$ = SQLX$ + " ,  PESOKG = " & FORMATNUM$(LblKgImpresos, "F16.5")
            SQLX$ = SQLX$ + " ,  IDENLOTE = '" & LoteNorfa$ & "'"
            SQLX$ = SQLX$ + " WHERE MOVISTO_ID = " & IdentifMovisto&
            FLEXCONN.Execute (SQLX$)
        '   'LO IMPRIME CUANDO VA A LA PANTALLA DE UBICACIONES
'            Call CIERRE_CMB.PRINTETIQORFA_CAMBACERES(NROOF$, LoteNorfa$, pesoReal, metrosReal, 1)
            If ESTRAZABLE%(CIFAB%) = 1 Then
              'creo el lote
              FE% = HOY(HO$)
              Call CREALOTIN(FE%, FE%, FE%, CIFAB%, LoteNorfa$, "", "", CDbl(1))
              Call ACSALOTE(CIFAB%, LoteNorfa$, "NOMESS")
            End If

        End If
    End If
   
    
    ''''
    
    If RetIdentificador > 0 Then 'PARA QUE ENTRES SOLO CUANDO RESGISTRO FOTO DE CIERRE
      'CON Ret_idMov& VINCULA LA FOTO DE CIERRE CON EL REGISTRO DEL MOVISTO
      SQLX$ = "UPDATE FOTODETACIERRE2 SET IDMOVISTO = " & Ret_idMov& & " WHERE IDMOVISTO= " & RetIdentificador
      FLEXCONN.Execute (SQLX$)
      SQLX$ = "UPDATE FOTODETACIERRE3 SET IDMOVISTO = " & Ret_idMov& & " WHERE IDMOVISTO= " & RetIdentificador
      FLEXCONN.Execute (SQLX$)
    End If
    '
    'CINTILLA
    If TRIM$(UCase$(CODFAMIL$(CIFAB%))) = "CC" Then
       DEPOCINTILLA% = XVALO(Control("CONTEPLA", "DEPOCINT"))
       Call CTP.RecuperarLoteCintilla(NORFA$, DEPOCINTILLA%)
    End If

   'RECORRE EL CONSUMO REALIZADO Y REEMPLAZA (EN EL MOVISTO) X LOS CODIGOS TRANSFERIDOS QUE CORRESPONDAN A LOS COIGOS RANGOS
    NROCOTI& = XVALO(TXTNROCOTIZA)
    
    IDCOORDUBIC% = 0
    IDMOVTALLER = XVALO(TXTIDMOVTALLER) 'EN EL CASO QUE ES DE TALLER TIENE VALOR EN ESTE TXT
    If IDMOVTALLER > 0 Then
       IDCOORDUBIC% = 400
    End If

    If (XVALO(Mid$(NORFA$, 4, 6)) >= 2341 Or XVALO(Mid$(NORFA$, 4, 6)) = 2325) And (InStr(CODFAMI$, "BB") > 0 Or InStr(CODFAMI$, "EE") Or InStr(CODFAMI$, "CC")) And InStr(EPAC$, "CONTEPLAST") > 0 Then 'ESTA ES A PARTIR DE LA CUAL SE VA CONSIDRERAR DIFERENTE LA DEVOLUCION DE LA BOBINA
       Call ConsumoDeCortesDelInfofab(NORFA$, CANTI, CIFAB%, Ret_idMov&, NROCOTI&)
       GoTo 80
    Else
       Call CTP.RecorrerOperacionesSegunOf(NORFA$, CANTI, NROCOTI&, IDCOORDUBIC%)  'APLICA EN INFOFAB EN EL CONSUMO
    End If
'    If Left(TRIM$(Text10), 3) = "CTP" Then ' REEMPLAZA LOS CODIGOS RANGO X LOS DE TELA
'       Call CTP.ConsumoSegunLosCodigosTransferidosCTP(NORFA$, XVALO(Me.LBLDEPOCONS)) 'RECORRE EL CONSUMO REALIZADO Y REEMPLAZA X LOS CODIGOS TRANSFERIDOS QUE CORRESPONDAN A LOS COIGOS RANGOS
'    End If
    '

    Call CTP.ConsumoSegunLosCodigosTransferidosCTP(NORFA$, XVALO(Me.LBLDEPOCONS))
'    Call RecuperarLoteDeConsumosdeOrdenesdeFabricacion(0, NORFA$, XVALO(Me.LBLDEPOCONS))
    DEPOESLINGA% = XVALO(Control("CONTEPLA", "DEPOESLI"))
    Call RecuperarLoteDeConsumosdeOrdenesdeFabricacion(0, NORFA$, DEPOESLINGA%)
   

    DEPOLINER% = XVALO(Control("CONTEPLA", "DEPLINER"))
    Call ConsumoLiner(CIFAB%, NORFA$, DEPOLINER%, CANTI)
    
80  NC = XVALO(Text6)
    NP& = XVALO(Text12)
    NUORDENPED% = XVALO(Text1)
    FEENT% = CTP.FechaEntregaSegunNroPedyOrdenPedido(NP&, NUORDENPED%)
    '
'    If XVALO(Mid$(NORFA$, 4, 6)) >= 2341 And InStr(CODFAMI$, "BB") > 0 And InStr(EPAC$, "CONTEPLAST") > 0 Then 'ESTA ES A PARTIR DE LA CUAL SE VA CONSIDRERAR DIFERENTE LA DEVOLUCION DE LA BOBINA
'       Call ACTUALIZAORFA(CIFAB%, NORFA$, 0, HOY(HO$), "") ' SE ACTUALIZA EL CIERRE DE OF Y SI SE CUMPLE SE MARCA COMO CERRADA.
'    End If
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
    
    'Call PrintEtiqProdTerminado
    FE% = HOY(HO$)
'    Call ctp.PRINTETIQUETACTP(CIFAB%, NORFA$, "", CANTI, FECHATEX$(FE%), CANTIPAQ)
     If XVALO(FEENT%) < 1 Then FEENT% = FE%
    Call UBICARPROCUTOTERMINADO(NORFA$, CANTI, NC, NP&, NUORDENPED%, FEENT%)
    
    'ACA DEBERIA ASIGNAR LOS LOTES.
    'INFORMAR UBICACION.
    
    Label4.Caption = ""
    Call BLANFORMU
99  Command1.Enabled = True
    Set CTP = Nothing
    
'  SELECT C.CODIGO, C.DESCRIPCION, C.COMPONENTE,C.CODINTELEM, C.MODUXBB,C.CORTE, SUM(C.MODUXBB * C.CORTE * 5) AS NECESIDAD_EN_M
', SUM(C.MODUXBB * 5) AS NECESIDAD_DE_CORTES FROM COTI1 C  WITH(NOLOCK)  WHERE CODINTELEM > 0  And NROCOTIZACION = 3955
'GROUP BY CODIGO, C.DESCRIPCION, C.COMPONENTE,C.CODINTELEM, C.MODUXBB,C.CORTE

'CODIGO               DESCRIPCION                                COMPONENTE             CODINTELEM  MODUXBB     CORTE   NECESIDAD_EN_M  NECESIDAD_DE_CORTES

'CA E 6.5 1250 BL     ESLINGA BLANCA 6.5cm para 1250Kg. (M.)     ESLINGAS [CM]               2045    4          120     2400            20
'CA TPL 101 210BL     TELA PLANA LAMINADA 101cm 210g BLANCA (M.) LATERALES (90-90-120)        647    2          129     1290            10
'CA TPL 101 210BL     TELA PLANA LAMINADA 101cm 210g BLANCA (M.) UPANEL (90-90-120)           647    1          343     1715            5
'CA TPL 95 74 BL      TELA PLANA LAMINADA 95cm 74g BLANCA (M.)   POLLERA (90-90-120)          621    1          367     1835            5
'CA TTL 57 72 BL      TELA TUBULAR LAMINADA 57cm 72g BLANCA (M.) VÁLVULAS DE DESCARGA [CM]    675    1          78      390             5
'CINT 080             CINTILLA 1.4/2.0 x 80 cm                   CINTILLA DS [CM]             971    1          145     725             5
'CINT 080             CINTILLA 1.4/2.0 x 80 cm                   CINTILLA POLLERA [CM]        971    1          80      400             5
'CINT 080             CINTILLA 1.4/2.0 x 80 cm                   CINTILLAS D [CM]             971    1          80      400             5


End Sub

Sub ConsumoDeCortesDelInfofab(NORFA$, CANTI_ACERRAR, CIFAB%, idMov&, NROCOTI&)
       'idMov& es el que se genero cuando se dio de alta en el movisto el movimiento OF (INGRESO DE P.TERMINADO)
        If NROCOTI& < 1 Then Exit Sub
        Dim CTP As New Conteplast
        If CANTI_ACERRAR < 1 Then CANTI_ACERRAR = 1
        SQLX$ = "SELECT C.CODIGO, C.DESCRIPCION, C.COMPONENTE,C.CODINTELEM, C.MODUXBB,C.CORTE"
        SQLX$ = SQLX$ + ", SUM(C.MODUXBB * C.CORTE * "
        SQLX$ = SQLX$ & CANTI_ACERRAR
        SQLX$ = SQLX$ + ") AS NECESIDAD_EN_M" 'ESTO EN PRINCIPIO NO LO UTILIZO LO DEJO POR LAS DUDAS
        
        SQLX$ = SQLX$ + ", SUM(C.MODUXBB * "
        SQLX$ = SQLX$ & CANTI_ACERRAR
        SQLX$ = SQLX$ + ") AS NECESIDAD_DE_CORTES"
        
        SQLX$ = SQLX$ + " FROM COTI1 C  WITH(NOLOCK)"
        SQLX$ = SQLX$ + "  WHERE CODINTELEM > 0 "
        SQLX$ = SQLX$ + " And NROCOTIZACION = " & NROCOTI&
        SQLX$ = SQLX$ + " GROUP BY CODIGO, C.DESCRIPCION, C.COMPONENTE,C.CODINTELEM, C.MODUXBB,C.CORTE"

        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            Descrioper$ = SAFETEXT$(!COMPONENTE)
            Descrioper$ = TRIM$(CTP.SectorSinParentesisConCorchetes$(Descrioper$))

            CantCortesCalculadosNecesarios = XVALO(!NECESIDAD_DE_CORTES) 'LA CANTIDAD TOTAL QUE PRECISO PARA CADA CORTE
            CONDI$ = "DESCOPER = '" & Descrioper$ & "' AND codiconj = " & CIFAB%
            NUMEOPER% = XVALO(VALOBADA("SECOPER", "NUMEOPER", CONDI$, "NOMESS"))
'            APLICA EL CONSUMO EN INFOFAB
            APLICONSUMO% = CTP.NuevaAPLICARCONSUMO%(NORFA$, Descrioper$, CantCortesCalculadosNecesarios, NUMEOPER%, idMov&, 100)
            .MoveNext
          Loop
        End With
        rs.Close: Set rs = Nothing
        
   'CON ESTO OBTENGO LA NECESIDAD
'    SQLX$ = "SELECT  C.CODIGO, C.DESCRIPCION, C.COMPONENTE,C.CODINTELEM, C.MODUXBB,C.CORTE, SUM(C.MODUXBB * C.CORTE * "
'    SQLX$ = SQLX$ & CANTI_ACERRAR
'    SQLX$ = SQLX$ + ") AS NECESIDAD_EN_M"
'    SQLX$ = SQLX$ + ", SUM(C.MODUXBB * "
'    SQLX$ = SQLX$ & CANTI_ACERRAR
'    SQLX$ = SQLX$ + ") AS NECESIDAD_DE_CORTES"
'
'    SQLX$ = SQLX$ + " FROM COTI1 C  WITH(NOLOCK)"
'    SQLX$ = SQLX$ + " Where CODINTCONJ = " & CIFAB%
'    SQLX$ = SQLX$ + "  And CODINTELEM > 0 "
'    SQLX$ = SQLX$ + " And NROCOTIZACION = " & NROCOTI&
'
'    SQLX$ = SQLX$ + " and LEFT(COMPONENTE,8) <> 'CINTILLA'"
'    SQLX$ = SQLX$ + " and LEFT(COMPONENTE,4) <> 'HILO'"
'
'    SQLX$ = SQLX$ + " GROUP BY  CODIGO, C.DESCRIPCION, C.COMPONENTE,C.CODINTELEM, C.MODUXBB,C.CORTE"
'    REG& = 0
'
'    Set rs = READSET(SQLX$)
'    With rs
'      Do While Not .EOF
'        REG& = REG& + 1
'        MSF3.Rows = REG& + 1
'        MSF3.TextMatrix(REG&, 1) = SAFETEXT$(!Codigo)
'        MSF3.TextMatrix(REG&, 2) = SAFETEXT$(!Descripcion)
'        MSF3.TextMatrix(REG&, 3) = SAFETEXT$(!COMPONENTE)
'        MSF3.TextMatrix(REG&, 4) = XVALO(!CODINTELEM)
'        MSF3.TextMatrix(REG&, 5) = XVALO(!MODUXBB)
'        MSF3.TextMatrix(REG&, 6) = XVALO(!CORTE)
'        MSF3.TextMatrix(REG&, 7) = FORMATNUM$(XVALO(!NECESIDAD_EN_M), "F9.1")
'        MSF3.TextMatrix(REG&, 8) = SAFETEXT$(!CODRANGO)
'        MSF3.TextMatrix(REG&, 9) = FORMATNUM$(XVALO(!NECESIDAD_DE_CORTES), "F9.0")
'        MSF3.TextMatrix(REG&, 10) = ""
'        Call COLOREAR_TEXTO_SOLO_UNA_AFILA(Me.MSF3, vbRed, REG&, 11)
'        MSF3.TextMatrix(REG&, 11) = FORMATNUM$(XVALO(!NECESIDAD_DE_CORTES), "F9.0") 'POR DEFECO FALTA = NECESIDAD
'                                                                    ', LUEGO SI HAY DISPONIBLIDAD SE SOBRE ESCRIBE
'
'
'        .MoveNext
'      Loop
'    End With
'    rs.Close: Set rs = Nothing

   
End Sub


Function FIRMAELECTRONICAOK%()
     FIRMAELEC.Text1 = ""
     FIRMAELEC.Text2 = ""
     FIRMAELEC.Label1 = ""
     FIRMAELECTRONICAOK% = -1
     FIRMAELEC.Show 1

     If FIRMAELEC.ESTADO > 0 Then
        FIRMAELECTRONICAOK% = 1
        Exit Function
     Else
        If NIVCONFI%(XVALO(FIRMAELEC.Label1)) < 3 Then
         Call COMUNI("Usuario Con Nivel de Clasificación Inválido Para Esta Operación")
         FIRMAELECTRONICAOK% = -1
         Exit Function
        End If
     End If

End Function
Sub ConsumoLiner(CIFAB%, NORFA$, DEPOLINER%, CANTI)
   Dim CTP As New Conteplast
   If CTP.FuncionLINER_SI_NO(CIFAB%) <> "SI" Then Exit Sub
   '
   CODICONJ$ = CODEXT(CIFAB%)
   familia$ = Control("CONTEPLA", "FAMILINE")
   FECONSUM% = HOY(HO$)
   CMOV$ = "CF"
   REFE$ = "CONSUMO " & CODICONJ$
   DOPRI$ = NORFA$
   DOSEC$ = "PF-" + TRIM$(Mid$(NORFA$, 4))
   REFE$ = "CONSUMO " + CODICONJ$
   NUORIT% = 0

   'BUSCO LOS LINERS TRANFERIDOS Y CONSUME DE LOS QUE ESTAN
   SQLX$ = " SELECT SUM(M.CANTINGRE - M.CANTSALID) AS RESUL, M.IDENLOTE,M.DOSECLAR,M.COD_INTE "
   SQLX$ = SQLX$ + " from MOVISTO M WITH(NOLOCK) INNER JOIN MASTER MA ON MA.CODINT = M.COD_INTE"
   SQLX$ = SQLX$ + " where M.DOSECLAR = '" & NORFA$ & "'"
   SQLX$ = SQLX$ + " AND M.DEPOMOVI = " & DEPOLINER%
   SQLX$ = SQLX$ + " AND MA.FAMILIA = '" & familia$ & "'"
   SQLX$ = SQLX$ + " GROUP BY M.IDENLOTE, M.DOSECLAR ,M.COD_INTE"
   Set rs = READSET(SQLX$)
   With rs
     Do While Not .EOF
        CI2% = XVALO(!cod_inte)
        LOTE$ = SAFETEXT$(!idenlote)
        SALDO = XVALO(!RESUL)
        NUORIT% = NUORIT% + 1
        If SALDO > 0 Then
           If SALDO < CANTI Then
              CAN = SALDO
           Else
              CAN = CANTI
           End If
           CANTI = CANTI - CAN
           Call MOVISTOK(FECONSUM%, CI2%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, 0, 0, 0, DEPOLINER%, CAN * (-1))
           If CANTI < 1 Then Exit Do
        End If
        .MoveNext
     Loop
   End With
   rs.Close
   Set rs = Nothing
   Screen.MousePointer = 1

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
    TXTIDMOVTALLER = ""
    Text3 = ""
        
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

   DetaCierreCtp.TXTACERRAR = XVALO(TXTCANTIDADACERRAR)
   
   'ESTE BLOQUE SIEMPRE AL FINAL
   DetaCierreCtp.Text10 = ""
   DetaCierreCtp.Text10 = Label4
   DetaCierreCtp.Show 1
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
   NORFA$ = Label4.Caption
   CieFinBo.Text10 = NORFA$
   
   CieFinBo.Show 1
End Sub

Private Sub Form_Load()
   EPAC$ = TRIM$(UCase$(EMPREAC$))
    
   If InStr(EPAC$, "CONTEPLAST") < 1 Then Command2.Visible = False
   If InStr(EPAC$, "QUILMES") < 1 Then
      Label11.Visible = False
      Text2.Visible = False
   End If
   DEPOCORTE% = XVALO(Control("", "DEPODEST"))
   DEPOPRODTER% = XVALO(Control("", "DEPOCIE"))

   Me.LBLDEPOCONS = DEPOCORTE%
   Me.LBLDEPOENT = DEPOPRODTER%
   Call CREATABLA_FOTODETACIERRE
End Sub


Private Sub Label4_Change()
'   MSF1.Rows = 1
   Command5.Visible = False 'BOTON DE CIERRE FINAL PARA BOLSAS
   Me.Height = 4200
   Call Centrar(Me)
   EPAC$ = UCase$(TRIM$(EMPREAC$))
   NORFA$ = TRIM$(Label4)

   If NORFA$ = "" Then Exit Sub
   CI1% = CODIFAB%(NORFA$)
   FAM$ = TRIM$(UCase$(CODFAMIL$(CI1%)))
   If InStr(EPAC$, "QUILMES") > 0 Then
    If FAM$ = "TI" Then 'TELAS IMPRESAS
       CieSolBol.Label4 = NORFA$
       'SI ES TELA IMPRESA  SE VA AL OTRO FRM PARA BOLSAS
       CieSolBol.Show 1
       Exit Sub
    End If
   End If
   
   TXTCANTIDADACERRAR.Enabled = True
   TXTCANTIDADACERRAR = 0
   Label2 = RESP_ZELE_VECT(2)
   Dim CTP As New Conteplast
   TXTIDMOVTALLER = ""
   If CTP.SaldoTallerSegunOF(NORFA$) > 0 Then 'VALIDA SI TIENE MATERIAL EN TALLER.
      IDX = EntregasTallerPendientesID(NORFA$)
      If IDX < 1 Then
         Call COMUNI("Esta O.Fabricaciòn Posee Entregas Pendientes de Taller\Debe Seleccionar una de Estas")
         Exit Sub
      End If
      TXTIDMOVTALLER = IDX
      CantiSeelecACerrarTaller = RESP_ZELE_VECT(4)
      TXTCANTIDADACERRAR = XVALO(CantiSeelecACerrarTaller)
      TXTCANTIDADACERRAR.Enabled = False
   End If
   Set CTP = Nothing
   '
'   CI1% = CODIFAB%(NORFA$)
'   FAM$ = TRIM$(UCase$(CODFAMIL$(CI1%)))
   If FAM$ = "CC" Then
      Me.LBLDEPOCONS = XVALO(Control("CONTEPLA", "DEPOCINT"))
      Me.LBLDEPOENT = XVALO(Control("", "DEPODEST"))  'DEPOCORTE%
   End If
   '
   If InStr(EPAC$, "QUILMES") > 0 Then
     Text2 = XVALO(PESUNI(CI1%)) ' LOS METROS CONFIGURADOS EN EL MAESTRO
     If FAM$ = "BO" Then 'si son bolsas el deposito de cierre es el de prod.term 6
       Me.LBLDEPOENT = XVALO(Control("", "DEPOCIE"))   'DEPOSITO PROD TERMINADO
       Me.LBLDEPOCONS = XVALO(Control("CONTEPLA", "DEBOAIMP")) 'Dep.Bobinas Para Impresión (Quilmes)(FRM ConfigDepo)
     Else
'     If FAM$ = "CC" Or FAM$ = "EE" Then 'ESLINGAS O CINTILLAS
      Me.LBLDEPOCONS = XVALO(Control("CONTEPLA", "DECOMAPR")) ' DEPOSITO DE MATERIA PRIMA (FRM configdepo)
      'Me.LBLDEPOENT = XVALO(CONTROL("", "DEPOCIE"))   'DEPOSITO PROD TERMINADO
      Me.LBLDEPOENT = XVALO(Control("CONTEPLA", "DEPOCINT")) 'el resto cierra en deposito 1 (FRM configdepo)
     End If
   End If
   
   TXTCODINT = CI1%
   Text10 = CODEXT$(CI1%)
   Label3 = DESCRIT0$(CI1%)
   CIFAB% = CODINT%(Text10)
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
   Text3 = VALOBADA("ORDEFABR", "ANEXOS", "IDSUBOR = '" & NORFA$ & "'", "NOMESS")

   TXTNROCOTIZA = NROCOTIZA(XVALO(NROPED&), XVALO(Text1))
   If TXTNROCOTIZA < 1 And Text12 < 1 Then
     If TRIM$(UCase$(CODFAMIL$(CI1%))) = "BO" Then
        Command5.Visible = True 'BOTON DE CIERRE FINAL PARA BOLSAS
     End If
   End If
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

Private Sub lblidregistro_Change()
        Campos$ = "Código/A18;N.Op./f5;Desc.Oper/A24;Cant.Cort. Orig./f12;Fact.Corte/F12;CORTE/F10.2;Lote/a16;Peso/F10.2;Cod.Rango/A12;Saldo/F10.2"
        Call CARGA_ENCABEZADO(MSF2, Campos$, Me)
        Call PoneEnCeroNull("INFOFAB", "CONSUMO", REGAF&, "NOMESS")

        SQLX$ = "SELECT  * FROM FOTODETACIERRE2  WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE IDMOVISTO = " & XVALO(lblidregistro)
        Reg& = 0

        Set rs = READSET(SQLX$)
        With rs
          Do While Not .EOF
            Reg& = Reg& + 1
            MSF2.Rows = Reg& + 1
            MSF2.TextMatrix(Reg&, 1) = SAFETEXT$(!CodXElem)
            MSF2.TextMatrix(Reg&, 2) = XVALO(!NUMEOPER)
            MSF2.TextMatrix(Reg&, 3) = SAFETEXT$(!descoper)
            MSF2.TextMatrix(Reg&, 4) = FORMATNUM$(XVALO(!CantCortes), "F9")
            MSF2.TextMatrix(Reg&, 5) = XVALO(!FACTORCORTE)
            MSF2.TextMatrix(Reg&, 6) = XVALO(!corte)
            MSF2.TextMatrix(Reg&, 7) = SAFETEXT$(!idenlote)
            MSF2.TextMatrix(Reg&, 8) = FORMATNUM$(XVALO(!Peso), "F9.1")
            MSF2.TextMatrix(Reg&, 9) = SAFETEXT$(!CODRANGO)
            MSF2.TextMatrix(Reg&, 10) = SAFETEXT$(!SALDO)
            .MoveNext
          Loop
        End With
        rs.Close: Set rs = Nothing
        
        '
        Campos$ = "Código/A16 '  1"
        Campos$ = Campos$ + ";Descripción/a48"             '  2"
        Campos$ = Campos$ + ";Componente/A24"               '  3"
        Campos$ = Campos$ + ";Cielem/A0"                    '  4"
        Campos$ = Campos$ + ";Moduxb/A0"                   '  5"
        Campos$ = Campos$ + ";CORTE/A0"                  '  6"
        Campos$ = Campos$ + ";Nec.M./A0"                       '  7"
        Campos$ = Campos$ + ";Cod.Rango/A12"                '  8"
        Campos$ = Campos$ + ";Nec.Cortes/f10.1"            '  9
        Campos$ = Campos$ + ";Disponibilidad/f14"              '  10
        Campos$ = Campos$ + ";Falta/f10"                       '  11
        Call CARGA_ENCABEZADO(MSF3, Campos$, Me)
        
        SQLX$ = "SELECT  * FROM FOTODETACIERRE3  WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE IDMOVISTO = " & XVALO(lblidregistro)
        Reg& = 0
        Set rs1 = READSET(SQLX$)
        With rs1
          Do While Not .EOF
            Reg& = Reg& + 1
            MSF3.Rows = Reg& + 1
            MSF3.TextMatrix(Reg&, 1) = SAFETEXT$(!Codigo)
            MSF3.TextMatrix(Reg&, 2) = SAFETEXT$(!Descripcion)
            MSF3.TextMatrix(Reg&, 3) = SAFETEXT$(!COMPONENTE)
            MSF3.TextMatrix(Reg&, 4) = XVALO(!CODINTELEM)
            MSF3.TextMatrix(Reg&, 5) = XVALO(!moduxbb)
            MSF3.TextMatrix(Reg&, 6) = XVALO(!corte)
            MSF3.TextMatrix(Reg&, 7) = FORMATNUM$(XVALO(!NECES_MTS), "F9.1")
            MSF3.TextMatrix(Reg&, 8) = SAFETEXT$(!CODRANGO)
            MSF3.TextMatrix(Reg&, 9) = FORMATNUM$(XVALO(!NECESIDAD), "F9.0")
            MSF3.TextMatrix(Reg&, 10) = FORMATNUM$(XVALO(!DISPONIBILIDAD), "F9.0")
            Call COLOREAR_TEXTO_SOLO_UNA_AFILA(Me.MSF3, vbRed, Reg&, 11)
            MSF3.TextMatrix(Reg&, 11) = FORMATNUM$(XVALO(!FALTA), "F9.0") 'POR DEFECO FALTA = NECESIDAD
            .MoveNext
          Loop
        End With
        rs1.Close: Set rs1 = Nothing


End Sub

Private Sub mnuCierresRealizados_Click()
    NORFA$ = TRIM$(Label4)
    If NORFA$ = "" Then
      Call COMUNI("Debe Seleccionar Una Orden de Fabricación")
      Exit Sub
    End If
    
    
    TXTTOTALCIERRES = CantiAprobadaOF(NORFA$)

    Campos$ = "CI/A0;" '1
    Campos$ = Campos$ + "Código/A16;" '2
    Campos$ = Campos$ + "Descripción/A32;" '3
    Campos$ = Campos$ + "F.Emis./D8;" '4
    Campos$ = Campos$ + "Cantidad./F9.1;" '5
    Campos$ = Campos$ + "Movisto_id/A0;" '6
    Campos$ = Campos$ + "Ubicación/A12;" '7
    Campos$ = Campos$ + "Cant.Pq./f8;" '8
    Campos$ = Campos$ + "ID/A10;" '9
    Call CARGA_ENCABEZADO(MSF1, Campos$, Me)

    Dim CTP As New Conteplast
    SQLX$ = "SELECT  COD_INTE,COD_EXTE,DESCRIP,FECHMOV, CANTINGRE,MOVISTO_ID,IdCalidad "
    SQLX$ = SQLX$ + " FROM MOVISTO  WITH(NOLOCK)  "
    SQLX$ = SQLX$ + " WHERE CODIMOVI = 'OF' AND CANTINGRE > 0" 'PARA QUE NO TRAIGA LOS CIERRES ANULADOS
    SQLX$ = SQLX$ + " AND DOPRILAR = '" & NORFA$ & "'"
    SQLX$ = SQLX$ + " ORDER BY MOVISTO_ID DESC"
    
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        J& = J& + 1
        MSF1.Rows = J& + 1
        MSF1.TextMatrix(J&, 1) = XVALO(!cod_inte)
        MSF1.TextMatrix(J&, 2) = SAFETEXT$(!Cod_Exte)
        MSF1.TextMatrix(J&, 3) = SAFETEXT$(!Descrip)
        MSF1.TextMatrix(J&, 4) = FECHATEX$(CVINT(!FechMov))
        MSF1.TextMatrix(J&, 5) = FORMATNUM$(XVALO(!CANTINGRE), "F9.1")
        MSF1.TextMatrix(J&, 6) = XVALO(!MOVISTO_ID)
        MSF1.TextMatrix(J&, 7) = CTP.CordenadasUbicaciónSegunMovistoid(XVALO(!MOVISTO_ID))
        MSF1.TextMatrix(J&, 8) = XVALO(!IdCalidad)
        MSF1.TextMatrix(J&, 9) = XVALO(!MOVISTO_ID)
        .MoveNext
      Loop
    End With
    
    rs.Close
    Set rs = Nothing
    Set CTP = Nothing
    
    Me.Height = 10185
    Call Centrar(Me)

End Sub

Private Sub mnuDepFabCerradas_Click()
   
10  Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
    If PERIO$ = "" Then Exit Sub

    DESFE = FETEXCOR1$(DESDEX%)
    HASFE = FETEXCOR1$(HASTAX%, "HASTA")
    '
20  Campos$ = "CI/A0;Còdigo/A16;Descripciòn/A32;O.Fabric./A16"
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    J& = 0
    SQLX$ = "select o.Cod_Inte,o.Cod_Exte,o.Descrip,o.idsubor from ordefabr o WITH(NOLOCK)"
    SQLX$ = SQLX$ + " inner join infofab  i on o.idsubor = i.IdSubOr"
    SQLX$ = SQLX$ + " Where o.StatuOrf = 3 And i.CantReal > i.consumo"
    SQLX$ = SQLX$ + " AND o.FechCierr >= '" & DESFE & "' and o.FechCierr <='" & HASFE & "'"
    SQLX$ = SQLX$ + " group by o.Cod_Inte,o.Cod_Exte,o.Descrip,o.idsubor"
    SQLX$ = SQLX$ + " order by idsubor "
    Set rs = READSET(SQLX$)
    
    With rs
    Do While Not .EOF
        CI1% = XVALO(!cod_inte)
        Codigo$ = SAFETEXT$(!Cod_Exte)
        DESXRI$ = SAFETEXT$(!Descrip)
        NORFA$ = SAFETEXT$(!idsubor)
        'FE% = CVINT(!FechCierr)
        
        J& = J& + 1
        GRILLAGRAL.GRID.Rows = J& + 1
        GRILLAGRAL.GRID.TextMatrix(J&, 1) = CI1%
        GRILLAGRAL.GRID.TextMatrix(J&, 2) = Codigo$
        GRILLAGRAL.GRID.TextMatrix(J&, 3) = DESXRI$
        GRILLAGRAL.GRID.TextMatrix(J&, 4) = NORFA$
'        GRILLAGRAL.GRID.TextMatrix(j&, 5) = FECHATEX$(FE%)
       .MoveNext
    Loop
   End With
   rs.Close
   Set rs = Nothing
   '
   GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
   GRILLAGRAL.mnuMenu1.Visible = True
   GRILLAGRAL.mnuMenu2.Caption = "Igualar Consumo vs Cortes Informados"
   GRILLAGRAL.mnuMenu2.Visible = True
   GRILLAGRAL.GRID.ZOrder 0
   GRILLAGRAL.TXTBUSCAR = ""
   GRILLAGRAL.Caption = "Depuraciòn de O.F. Cerradas (" & PERIO$ & ")"
30 Erase RESP_ZELE_VECT()

   GRILLAGRAL.Show 1
   '
    
   NORFA$ = RESP_ZELE_VECT(4)
   If NORFA$ <> "" Then
       Call VerConsumoComponentesSegunInfofabConteplast(NORFA$)
       GoTo 30
   End If
   Erase RESP_ZELE_VECT()
    

   Dim MSF As msFlexGrid
   If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "DepOFCer", "Depuraciòn de O.F. Cerradas", "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
        GoTo 20
   End If
   
   If GRILLAGRAL.MENU2 > 0 Then
        'VALIDAR QUE ESTE ACTIVA LA GRILLA GRID SI NO AVISAR
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2

        Call IgualarCantidadConsumidaConConsumoInfofabConteplast(MSF)
        
        For i& = 1 To MSF.Rows - 1 ' PONGO EN BLANCO LAS CELDAS NUEVAMENTE
          Call COLOREAR_GRILLA_SOLAFILA(MSF, vbWhite, i&)
         Next i&
        GRILLAGRAL.MENU2 = 0
        Screen.MousePointer = 1
        GoTo 20
   End If
End Sub

Function VerConsumoComponentesSegunInfofabConteplast(NORFA$)

    Campos$ = "CI/A0;Còdigo/A16;Cant.Transf./F9.1;Consumida/F9.1;Lote/A16;Corte/f10.1;Factor/F10.1"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = "Consumos Informados en Cantidad de Cortes O.F.: " & NORFA$

    Dim obj As New RECORXET
    Set rs1 = obj.RS_ConsInfofab(NORFA$)
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2, 1)
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    FRMZELECHO.Width = 12500
    FRMZELECHO.Show 1

End Function
    


Private Sub mnuMarcadeCierre_Click()
   If DERACCE%("CIEMASOF", "Cierre Automático de Ordenes de Fabricación") < 2 Then Exit Sub
   CIEMASOF.Command3.Visible = False
   CIEMASOF.BTNMARCACIERREOF.Visible = True
   CIEMASOF.Frame2.Visible = False
   CIEMASOF.Frame3.Visible = True
   CIEMASOF.Caption = "Marca Masiva de Cierre de Ordenes de Fabricación Seleccionadas"
   CIEMASOF.Show 1

End Sub


Private Sub mnuMatProceso_Click()

'        NORFA$ = OrdenFabricacionAbiertas$

        
        Campos$ = "CI/A0;CODIGO/A16;Descripción/A40;En Corte(M)/F11.1;Accesorios/F9.1;Impresiòn/F9.1;Semiterminados/F12.1;Talleres/F9.1;Linea/F9.1"
        Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL)
        Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL)
        DEPOCORTE% = Control("", "DEPODEST")
        Dim CTP As New Conteplast
        J& = 0
        SQLX$ = "SELECT M.Cod_Inte, M.COD_EXTE, SUM(M.CANTINGRE-CANTSALID)AS RESUL FROM MOVISTO M  WITH(NOLOCK) "
        SQLX$ = SQLX$ + " inner join MASTER MAS ON MAS.Codint = M.COD_INTE"
        SQLX$ = SQLX$ + " INNER  join ORDEFABR O ON (M.DOPRICOR = O.IDSUBOR OR M.DOSECLAR = O.IDSUBOR )"
        SQLX$ = SQLX$ + " WHERE DEPOMOVI = 2 "
        If NORFA$ <> "" Then
            SQLX$ = SQLX$ + " AND (DOSECLAR = '" & TRIM(NORFA$) & "' or  DOPRICOR = '" & TRIM(NORFA$) & "')"
        Else
            SQLX$ = SQLX$ + " AND (LEFT(DOSECLAR,3) = 'OF-' or  LEFT(DOPRICOR,3) = 'OF-')"
        End If
        SQLX$ = SQLX$ + " And MAS.TIMATE <> 49 "
        SQLX$ = SQLX$ + " And MAS.Status >= 0   "
        SQLX$ = SQLX$ + " And MAS.FAMILIA ='TE'  "
        SQLX$ = SQLX$ + " GROUP BY M.COD_EXTE,M.Cod_Inte ORDER BY M.COD_INTE"
'        SQLX$ = "SELECT M.Cod_Inte, M.COD_EXTE, SUM(M.CANTINGRE-CANTSALID)AS RESUL"
'        SQLX$ = SQLX$ + " FROM MOVISTO M WITH(NOLOCK)"
'        SQLX$ = SQLX$ + " inner join MASTER MAS ON MAS.Codint = M.COD_INTE"
'        SQLX$ = SQLX$ + " INNER  join ORDEFABR O ON M.DOSECLAR = O.IDSUBOR"
'        SQLX$ = SQLX$ + " Where m.DEPOMOVI = 2 And MAS.TIMATE <> 49 And MAS.Status >= 0"
'        SQLX$ = SQLX$ + " And O.StatuOrf < 3 and (CODIMOVI = 'TR' OR CODIMOVI = 'CF')"
'        SQLX$ = SQLX$ + " And MAS.FAMILIA ='TE'"
'        If NORFA$ <> "" Then SQLX$ = SQLX$ + " AND (DOSECLAR = '" & TRIM(NORFA$) & "' or  DOPRICOR = '" & TRIM(NORFA$) & "')"
'        SQLX$ = SQLX$ + " GROUP BY M.COD_EXTE,M.Cod_Inte"
'        SQLX$ = SQLX$ + " ORDER BY M.COD_INTE"
        Set rs = READSET(SQLX$)
    
        With rs
        Do While Not .EOF
            CI1% = XVALO(!cod_inte)
            Codigo$ = SAFETEXT$(!Cod_Exte)
            DESXRI$ = DESCRIT0$(CI1%)
            J& = J& + 1
            GRILLAGRAL.GRID.Rows = J& + 1
            GRILLAGRAL.GRID.TextMatrix(J&, 1) = CI1%
            GRILLAGRAL.GRID.TextMatrix(J&, 2) = Codigo$
            GRILLAGRAL.GRID.TextMatrix(J&, 3) = DESXRI$
            If CI1% = 623 Then
            A = 1
            End If
            GRILLAGRAL.GRID.TextMatrix(J&, 5) = FORMATNUM$(XVALO(EnSector(CI1%, 2, 2)), "F10.1") 'ACCESORIOS
            GRILLAGRAL.GRID.TextMatrix(J&, 6) = FORMATNUM$(XVALO(EnSector(CI1%, 1, 1)), "F10.1") 'IMPRESION
            GRILLAGRAL.GRID.TextMatrix(J&, 7) = FORMATNUM$(XVALO(EnSector(CI1%, 3, 3)), "F10.1") 'SEMITERMINADOS
            GRILLAGRAL.GRID.TextMatrix(J&, 8) = FORMATNUM$(XVALO(EnSector(CI1%, 30, 50)), "F10.1") 'TALLERES
            GRILLAGRAL.GRID.TextMatrix(J&, 9) = FORMATNUM$(XVALO(EnSector(CI1%, 100, 100)), "F10.1") 'LINEA
            ACUMULADOENSECTORES = XVALO(GRILLAGRAL.GRID.TextMatrix(J&, 5)) + XVALO(GRILLAGRAL.GRID.TextMatrix(J&, 6)) + XVALO(GRILLAGRAL.GRID.TextMatrix(J&, 7)) + XVALO(GRILLAGRAL.GRID.TextMatrix(J&, 8)) + XVALO(GRILLAGRAL.GRID.TextMatrix(J&, 9))
            If (XVALO(!RESUL) - ACUMULADOENSECTORES) < 0 Then ACUMULADOENSECTORES = 0
            A = (XVALO(!RESUL))
            GRILLAGRAL.GRID.TextMatrix(J&, 4) = FORMATNUM$(XVALO(!RESUL) - ACUMULADOENSECTORES, "F10.1")
            If XVALO(GRILLAGRAL.GRID.TextMatrix(J&, 4)) <= 0 Then GRILLAGRAL.GRID.TextMatrix(J&, 4) = 0
            .MoveNext
         Loop
        End With
        rs.Close
        Set rs = Nothing

        GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
        GRILLAGRAL.mnuMenu1.Visible = True
'        GRILLAGRAL.mnuMenu2.Caption = "Ajustar Consumos De Ordenes de Fabricaciòn Seleccionadas"
'        GRILLAGRAL.mnuMenu2.Visible = True
        GRILLAGRAL.GRID.ZOrder 0
        GRILLAGRAL.TXTBUSCAR = ""
        GRILLAGRAL.Caption = "Materiales en Sector de Corte"
30      Erase RESP_ZELE_VECT()

      GRILLAGRAL.Show 1
      '
      If GRILLAGRAL.MENU1 > 0 Then
            Set MSF = GRILLAGRAL.GRID
            If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
            'Call CARGA_TABLA_IMPRE(MSF, Campos$, "SALDNORF", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
            Call FINAL("?" & RET_NAMECONS$)
            GRILLAGRAL.MENU1 = 0
            Screen.MousePointer = 1
      End If

  
      Erase RESP_ZELE_VECT()


End Sub


Private Sub mnuOrigScrap_Click()
   FRMSCRAP.Show 1
End Sub


Private Sub mnuSaldoDepoCorte_Click()
      OPX% = COMALTER%("Mostrar Saldos en Depòsito de Corte\\Cancelar\O.F. Abiertas\O.F. Cerradas\Todas")
      If OPX% = 2 Then
         AbiertasoCerradas% = 1: ABIERTAOCERRA$ = "Abiertas"
      ElseIf OPX% = 3 Then
         AbiertasoCerradas% = 3: ABIERTAOCERRA$ = "Cerradas"
      ElseIf OPX% = 4 Then
         AbiertasoCerradas% = 0: ABIERTAOCERRA$ = "Todas"
      Else
         Exit Sub
      End If
      Campos$ = "CI/A0;CODIGO/A20;Descripción/A40;Metros/F9.1;Kg./f9.1"
      Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL)
      Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL)
      DEPOCORTE% = Control("", "DEPODEST")
      Dim CTP As New Conteplast
      Call CTP.CargaMatTranfTodasLasOrdenes(GRILLAGRAL.GRID, DEPOCORTE%, AbiertasoCerradas%)

      
      GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
      GRILLAGRAL.mnuMenu1.Visible = True
      GRILLAGRAL.mnuMenu2.Caption = "Ajustar Consumos De Ordenes de Fabricaciòn Seleccionadas"
      GRILLAGRAL.mnuMenu2.Visible = True
      GRILLAGRAL.GRID.ZOrder 0
      GRILLAGRAL.TXTBUSCAR = ""
      GRILLAGRAL.Caption = "Saldos de Materiales Transferidos al Depòsito de Corte"
30    Erase RESP_ZELE_VECT()

      GRILLAGRAL.Show 1
      '
    
      CI1% = RESP_ZELE_VECT(1)
      Codigo$ = RESP_ZELE_VECT(2)
      
      If CI1% > 0 Then
        OPX% = COMALTER%("Opciones de Trabajo\\Cancelar\Ver Ordenes de Fabricaciòn\Informe Sector Corte")
        If OPX% = 1 Then GoTo 30
        If OPX% = 3 Then InfSecCorte.Show 1: GoTo 30
        
        
'       Call VerConsumoComponentesSegunInfofabConteplast(NORFA$)
       SQLX$ = "SELECT M.COD_INTE,M.COD_EXTE,M.DESCRIP,SUM(CANTINGRE-CANTSALID) AS RESULTADO ,O.IDSUBOR"
       SQLX$ = SQLX$ + " ,M.IDENLOTE FROM MOVISTO M WITH(NOLOCK)"
       SQLX$ = SQLX$ + " INNER JOIN ORDEFABR O ON (O.IdSubOr  = M.DoSecLar     OR O.IDSUBOR = M.DOPRICOR)"
       'SQLX$ = SQLX$ + " where  (left(DOSECLAR,3) = 'OF-'  or  left(DOPRICOR,3)   = 'OF-' ) AND DEPOMOVI = 2"
       SQLX$ = SQLX$ + " where (DOSECLAR LIKE '_F-______-1'  OR DOSECLAR LIKE 'AJ.%') AND DEPOMOVI = 2"
       
       If AbiertasoCerradas% > 0 Then
          SQLX$ = SQLX$ + " AND O.StatuOrf = " & AbiertasoCerradas%
       End If
       SQLX$ = SQLX$ + " AND M.Cod_Inte  = " & CI1%
       SQLX$ = SQLX$ + " GROUP BY M.COD_INTE,M.COD_EXTE,M.DESCRIP,O.IDSUBOR,M.IDENLOTE "
       SQLX$ = SQLX$ + " having SUM(CANTINGRE-CANTSALID)  <> 0 "
       SQLX$ = SQLX$ + " ORDER BY O.IDSUBOR,m.idenlote"
       Campos$ = "CI/A0;Còdigo/A16;Descripciòn/A32;Cant.Transf./F9.1;O.F./A12;Lote/A12"
       Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
       Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
        FRMZELECHO.Caption = "Ordenes de Fabricaciòn " & ABIERTAOCERRA$ & " Con Reservas del Còdigo: " & Codigo$
        Set Rst = FLEXCONN.Execute(SQLX$)
        Call Carga_MSF_Recordset(Rst, Campos$, FRMZELECHO.MSF2, 1)
        On Error Resume Next
        Rst.Close
        Set Rst = Nothing
        On Error GoTo 0
        FRMZELECHO.Width = 16500


       Set Rst = Nothing
       FRMZELECHO.Width = 12500
       FRMZELECHO.Show 1
       GoTo 30
      End If
      
      Dim MSF As msFlexGrid
      If GRILLAGRAL.MENU1 > 0 Then
            Set MSF = GRILLAGRAL.GRID
            If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
            Call CARGA_TABLA_IMPRE(MSF, Campos$, "SALDNORF", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
            Call FINAL("?" & RET_NAMECONS$)
            GRILLAGRAL.MENU1 = 0
            Screen.MousePointer = 1
      End If

   If GRILLAGRAL.MENU2 > 0 Then
        'VALIDAR QUE ESTE ACTIVA LA GRILLA GRID SI NO AVISAR
        
        If AbiertasoCerradas% <> 3 Then
           Call COMUNI("Este Proceso Solo Puedo Correrse para Ordenes de Fabricaciòn Cerradas")
           GoTo 30
        End If
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
         
        Call AjustarSegunOf(MSF)
        
        For i& = 1 To MSF.Rows - 1 ' PONGO EN BLANCO LAS CELDAS NUEVAMENTE
          Call COLOREAR_GRILLA_SOLAFILA(MSF, vbWhite, i&)
         Next i&
        GRILLAGRAL.MENU2 = 0
        Screen.MousePointer = 1
        GoTo 30
   End If

      Erase RESP_ZELE_VECT()


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

 Sub mnuTablaConfScrap_Click()
        Unload VENTABNEW
        VENTANA.Inicializar = 0
        VENTANA.CampEditables = "1;3;5;7"
        VENTANA.AgregaRegistro = 1
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        VENTANA.HabilitaInsertar = 1
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = "CONFIGURACIÓN DE  TABLAS SCRAP"
        VENTANA.ColEcoCampo(1) = "CAMPO1;CAMPO2;DATASQL;LISTA  = 'SECTSCRP'"
        VENTANA.ColEcoCampo(3) = "CAMPO1;CAMPO2;DATASQL;LISTA  = 'CATESCRP'"
        VENTANA.ColEcoCampo(5) = "CODIGO;DESCRIPCION;MASTER;FAMILIA  = 'SC'"
        
        Campos$ = "C.Sector/F10;" '0
        Campos$ = Campos$ + "Descripción/A32;"      '1
        Campos$ = Campos$ + "C.Categoría/F10;"      '2
        Campos$ = Campos$ + "Descripción/A64;"   '3
        Campos$ = Campos$ + "Cód.Art./A10;"       '4
        Campos$ = Campos$ + "Descripción/A64;"        '5
        Campos$ = Campos$ + "N°Dep./F8;"     '6
        Campos$ = Campos$ + "CI/A0;" '7
        Campos$ = Campos$ + "ID/A0" '8
'HASTA CA LLEGUE FALTA CONFIGURAR EL RECORDSET Y GENEAR LA TABLA DONDE SE VA A GUARDAR LA INFO
        VENTANA.Campox = Campos$
        VENTANA.Inicializar = 1
        Dim obj As New RECORXET
        Set rs1 = obj.RS_CONFIGSCRAP()
        Call Carga_MSF_Recordset(rs1, Campos$, VENTABNEW.MSF)
        Set obj = Nothing
        rs1.Close
        Set rs1 = Nothing
'        Call numerargrilla(VENTABNEW.MSF, 0, 0, 0, 0)
        Dim MSF1 As msFlexGrid
        Set MSF1 = VENTABNEW.MSF

30      VENTABNEW.Show 1
        If VENTABNEW.APROBADO% < 1 Then Exit Sub
        '
        'PRIMERO VALIDO
        For Reg& = 1 To MSF1.Rows - 1
                  CodSectorX = XVALO(MSF1.TextMatrix(Reg&, 1))
                  DescSectorX$ = Left$(MSF1.TextMatrix(Reg&, 2), 32)
                  If TRIM$(DescSectorX$) = "" Then
                     Call COMUNI("Atención Sector Inválidoen Fila: " & Reg&)
                     GoTo 30
                  End If
                  
                  CodCategX = XVALO(MSF1.TextMatrix(Reg&, 3))
                  DescCategX$ = Left$(MSF1.TextMatrix(Reg&, 4), 64)
                  If TRIM$(DescCategX$) = "" Then
                     Call COMUNI("Atención Categoría Inválida en Fila: " & Reg&)
                     GoTo 30
                  End If

                  CodArtX = Left$(MSF1.TextMatrix(Reg&, 5), 24)
                  DescArtX$ = Left$(MSF1.TextMatrix(Reg&, 6), 64)
                  If TRIM$(DescArtX$) = "" Then
                     Call COMUNI("Atención Artículo Inválido en Fila: " & Reg&)
                     GoTo 30
                  End If
                  
                  If XVALO(MSF1.TextMatrix(Reg&, 7)) > 32767 Then
                     Call COMUNI("Atención Depósito Inválido en Fila: " & Reg&)
                     GoTo 30
                  End If
                  
                  DepIngreX% = XVALO(MSF1.TextMatrix(Reg&, 7))
                  On Error Resume Next
                  Refedepo$ = TRIM$(ECOARCH$("TADEPOSI", Chr$(DepIngreX%)))
                  On Error GoTo 0
                  If Refedepo$ = "" Then
                     Call COMUNI("Atención Depósito Inválido en Fila: " & Reg&)
                     GoTo 30
                  End If
        Next Reg&

        'SI PASO VALIDACION GRABA
        'TRANSACCIÓN GRABACION
        On Error GoTo ERROR_GUARDAR
        FLEXCONN.BeginTrans

        CONDI$ = " CONFSCRAP_ID > 0"
        SQLX$ = "UPDATE CONFSCRAP SET MARBORRA = 1 WHERE " & CONDI$
        FLEXCONN.Execute (SQLX$)
        
        SQLX$ = "SELECT * FROM CONFSCRAP"
        SQLX$ = SQLX$ + " WHERE CONFSCRAP_ID < 1"
        FE% = HOY(HO$)
        FEC$ = FETEXCOR1$(FE%)
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

        With rs
              For Reg& = 1 To MSF1.Rows - 1
                  .AddNew
                  !CodSector = XVALO(MSF1.TextMatrix(Reg&, 1))
                  !DescSector = Left$(MSF1.TextMatrix(Reg&, 2), 32)
                  !CodCateg = XVALO(MSF1.TextMatrix(Reg&, 3))
                  !DescCateg = Left$(MSF1.TextMatrix(Reg&, 4), 64)
                  !CodArt = Left$(MSF1.TextMatrix(Reg&, 5), 24)
                  !DescArt = Left$(MSF1.TextMatrix(Reg&, 6), 64)
                  !DepIngre = XVALO(MSF1.TextMatrix(Reg&, 7))
                  !CiArt = CODINT%(Left$(MSF1.TextMatrix(Reg&, 5), 24))
                  !FECHA = FEC$
                  !CODIEMPR = CodiEmp%
                  !MARBORRA = 0
                  .Update
              Next Reg&
        End With
        
        rs.Close
        Set rs = Nothing
        
        CONDI$ = CONDI$ + " AND  MARBORRA = 1  "
        SQLX$ = "DELETE FROM CONFSCRAP WHERE " & CONDI$
        FLEXCONN.Execute (SQLX$)
        
        Set VENTANA = Nothing
        
ERROR_GUARDAR:
    If Err <> 0 Then
       FLEXCONN.RollbackTrans
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Grabación Realizado Exitosamente")
    End If
    Unload VENTABNEW
End Sub




Private Sub mnuUbic_Click()
   
   INFUBIPT.Show 1
End Sub

Private Sub MSF1_Click()
    Reg& = MSF1.MouseRow
    J& = MSF1.MouseCol
    If Reg& = 0 Or Reg& > MSF1.Rows Then Exit Sub

    Me.TXTCANTICORTES = XVALO(MSF1.TextMatrix(Reg&, 5))
    lblidregistro = XVALO(MSF1.TextMatrix(Reg&, 6))

End Sub

Private Sub Text2_Change()
' EPAC$ = UCase$(TRIM$(EMPREAC$))
' NORFA$ = Label4
' If NORFA$ <> "" Then
'    CI1% = CODIFAB%(NORFA$)
'    If UCase$(Unimed$(CI1%)) = "U." Then
'       If InStr(EPAC$, "QUILMES") > 0 Then
'          TXTCANTIDADACERRAR = Text2
'       End If
'    End If
' End If
 
End Sub

 
Private Sub Text6_Change()
   Label21 = rasocli$(XVALO(Text6))
End Sub

Sub cargarTransfDeOf_DeSalidasDeposito(NORFA$, DEPO%, MSF As msFlexGrid)


    Reg& = 0
    SQLX$ = "Select sum(cantsalid - cantingre) as CANTI,idenlote,cod_exte, DepoMovi,CODIMOVI,DESCRIP,ID_OPERACION"
    SQLX$ = SQLX$ + " from movisto WITH(NOLOCK) "
    SQLX$ = SQLX$ + " where doseclar = '" & NORFA$ & "'"
    SQLX$ = SQLX$ + " AND DepoMovi = " & DEPO%
    SQLX$ = SQLX$ + " AND CODIMOVI = 'TR'"
    SQLX$ = SQLX$ + " group by idenlote,cod_exte ,DepoMovi, CODIMOVI,DESCRIP,ID_OPERACION"
    SQLX$ = SQLX$ + " order by idenlote"
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
         Reg& = Reg& + 1
         MSF.Rows = Reg& + 1
         MSF.TextMatrix(Reg&, 1) = SAFETEXT$(!Cod_Exte)
         MSF.TextMatrix(Reg&, 2) = SAFETEXT$(!Descrip)
         MSF.TextMatrix(Reg&, 3) = SAFETEXT$(!idenlote)
         MSF.TextMatrix(Reg&, 4) = XVALO(!CANTI)
         MSF.TextMatrix(Reg&, 5) = XVALO(!CANTI)
         MSF.TextMatrix(Reg&, 6) = XVALO(!ID_OPERACION)
         MSF.TextMatrix(Reg&, 7) = SAFETEXT$(!ID_OPERACION)
        .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
End Sub
Sub AgregarKgMovisto(NORFA$, CANTI, KG, Ret_IdMovisto)
   'ESTO SE UTILIZA PARA QEU LUEGO DE GENERAR UNA ORDEN DE FABRICACION SE GRABE EL PESO
   'SIRVE PARA REALIZARLO INMEDIATAMENTE LUEGO DE LA GENERACION DE LA OF
   CONDI$ = "CODIMOVI = 'OF' AND CANTINGRE =  " & CANTI
   CONDI$ = CONDI$ + " AND DOPRILAR = '" & NORFA$ & "'"
   CONDI$ = CONDI$ + " AND (PESOKG = 0 OR PESOKG IS NULL)"
   CONDI$ = CONDI$ + " ORDER BY MOVISTO_ID  DESC"
   ID& = XVALO(VALOBADA("MOVISTO", "TOP 1 MOVISTO_ID", CONDI$, "NOMESS"))
   Ret_IdMovisto = ID& 'retorono de id
   CONDI$ = "MOVISTO_ID =" & ID&
   SQLX$ = "UPDATE MOVISTO SET PESOKG = '" & FORMATNUM$(KG, "F16.5") & "' WHERE " & CONDI$
   FLEXCONN.Execute (SQLX$)
End Sub

Sub AjusteA0SegunCodigoOfabricacionDeposito(CI1%, NORFA$, DEPO%, FEAJU%)
 If DERACCE%("AJUCERO", "Ajuste de Stocks a Cero") < 2 Then
     GoTo 99
   End If
   '
10 If FEAJU% < 1 Then
   Call SELECHO("SELFECHA")
   sFechaCorte$ = VALACT1$("SELFECHA")
   iFechaCorte% = FECHANUM(sFechaCorte$)
   If iFechaCorte% < 0 Then
     Call MENSERR(24, "Fecha de Ajuste no Válida")
     GoTo 10
   End If
   Else
      iFechaCorte% = FEAJU%
   End If
   If iFechaCorte% = 0 Then
    GoTo 99
   End If
   '
   
20 If DEPO% < 1 Then 'SI NO VIENE PARAMETRIZADO LO PIDE
        Call SELECHO("TADEPOSI")
        iDeposito% = XVALO(VALACT1$("TADEPOSI"))
        If iDeposito% <= 0 Then
          GoTo 10
        End If
   Else
        iDeposito% = DEPO%
   End If
   '
   sDeposito$ = FORMATNUM$(iDeposito%, "F3.0") + " " + ECOARCH$("TADEPOSI", Chr$(iDeposito%))
'   If COMALTER%("Esta Seguro de Ajustar a Cero\Todos los Stocks del Depósito Nro. : " & sDeposito$ & "\ al " & sFechaCorte$ & " \\No, Cancelar\Si, Ajustar a Cero") = 2 Then
        '
        Call ABRESTOCKS
        
        ' *** FILTRO POR DEPOSITO
        Call CARDEPOS
        iDepositoPrincipal% = 0
        If iDeposito% = COSU%(1) Then iDepositoPrincipal% = 1
        If iDepositoPrincipal% = 1 Then
            sFiltroDeposito$ = " AND (MOVISTO.DepoMovi = " & CStr(iDeposito%) & " OR MOVISTO.DepoMovi = 0) "
        Else
            sFiltroDeposito$ = " AND MOVISTO.DepoMovi = " & CStr(iDeposito%) & " "
        End If
        
        ' *** FILTRO POR ARTICULO
        If CI1% < 1 Then
            iPorArticulo% = ALTERNA%("Un Articulo\Sin Filtro Por Articulo")
            sFiltroArticulo$ = ""
            If iPorArticulo% = 1 Then
                Call SELECHO("MASTER")
                iArticulo% = CODINT(VALACT1$("MASTER"))
                If iArticulo% > 0 Then
                    sFiltroArticulo$ = " AND MASTER.CODINT = " & CStr(iArticulo%)
                End If
            End If
        Else
           iPorArticulo% = CI1%
           sFiltroArticulo$ = " AND MASTER.CODINT = " & CStr(iPorArticulo%)
        End If
        
        ' *** FILTRO POR STATUS DE ARTICULO
        sFiltroStatus$ = ""
'        If iPorArticulo% > 1 Then
'            iItemsBaja% = ALTERNA%("Solo Articulos De Baja\Todos Los Articulos")
'            If iItemsBaja% = 1 Then sFiltroStatus$ = " AND MASTER.STATUS < 0 "
'        End If
        
        Screen.MousePointer = 11
        
        ' *** PROXIMO NUMERO DE AJUSTE Y FORMATEO DEL NUMERO DE COMPROBANTE
        sCodMovimiento$ = TRIM$(Control$("", "CODAJUST"))
        If sCodMovimiento$ = "" Then sCodMovimiento$ = "AJ"
        NRO& = ProNroComprobante(sCodMovimiento$)
        sDocSecundarioCorto$ = "AJ." + Mid$(HOS$, 7, 2) + Mid$(HOS$, 4, 2) + Left$(HOS$, 2)
        sDocSecundarioLargo$ = NORFA$
        sDocPrimarioCorto$ = "AS." + TRIM$(str$(NRO&))
        sDocPrimarioLargo$ = sDocPrimarioCorto$

        ' ***********************************************
        ' *** AJUSTE DE STOCK A CERO PARA "" NO TRAZABLES ""
        ' ***********************************************
        Dim RstAjusteNoTrazables As ADODB.Recordset
        ConsultaSql$ = "SELECT SUM(CANTINGRE) AS INGRESOS, SUM(CANTSALID) AS SALIDAS, MASTER.CODINT, MOVISTO.COD_INTE, MOVISTO.DEPOMOVI "
        ConsultaSql$ = ConsultaSql$ & " FROM MASTER INNER JOIN MOVISTO WITH(NOLOCK) ON MASTER.CODINT = MOVISTO.COD_INTE WHERE MASTER.TRAZABLE = 0 "
        ConsultaSql$ = ConsultaSql$ & " AND MOVISTO.FECHMOV <= '" & FETEXCOR1$(iFechaCorte%) & "'"      ' FECHA DE CORTE
        ConsultaSql$ = ConsultaSql$ & sFiltroDeposito$                                                  ' DEPOSITO
        If sFiltroStatus$ <> "" Then ConsultaSql$ = ConsultaSql$ & sFiltroStatus$                       ' STATUS ARTICULO
        If sFiltroArticulo$ <> "" Then ConsultaSql$ = ConsultaSql$ & sFiltroArticulo$                   ' UN ARTICULO
        ConsultaSql$ = ConsultaSql$ + " AND MOVISTO.DOSECLAR = '" & NORFA$ & "'"
        ConsultaSql$ = ConsultaSql$ + " AND MOVISTO.IDENLOTE <> ''"
        ConsultaSql$ = ConsultaSql$ & " GROUP BY MOVISTO.COD_INTE, MOVISTO.DEPOMOVI, MASTER.CODINT "
        ConsultaSql$ = ConsultaSql$ & " ORDER BY MOVISTO.COD_INTE, MOVISTO.DEPOMOVI"
        
        Set RstAjusteNoTrazables = READSET(ConsultaSql$)
        
        iFecha% = iFechaCorte%
        iNroOrden% = 0
        With RstAjusteNoTrazables
            Do While Not .EOF
                iCodigo% = !CODINT
                iDeposito% = !depomovi
                
                dIngresos# = XVALO(!INGRESOS)
                dSalidas# = XVALO(!SALIDAS)
                dSaldo# = dIngresos# - dSalidas#
                dCantidadAjuste# = (-1) * dSaldo#   ' AJUSTA POR EL VALOR CONTRARIO
                
                If Abs(dCantidadAjuste) >= 0.005 Then
                    iNroOrden% = iNroOrden% + 1
                    dValorUnitario# = 0
                    sReferencia$ = "Inventario de Stocks a Cero"
                    iMoneda% = 0
                    iCliente% = 0
                    SLOTE$ = ""
                    Call MOVISTOK(iFecha%, iCodigo%, SLOTE$, sCodMovimiento$, sDocPrimarioLargo$, sDocPrimarioCorto$, sDocSecundarioLargo$, sDocSecundarioCorto$, sReferencia$, iNroOrden%, dValorUnitario#, iMoneda%, iCliente%, iDeposito%, dCantidadAjuste#)
                End If
                
                .MoveNext
            Loop
        End With
        
        On Error Resume Next
        RstAjusteNoTrazables.Close
        Set RstAjusteNoTrazables = Nothing
        On Error GoTo 0
        
        ' ***********************************************
        ' *** AJUSTE DE STOCK A CERO PARA "" TRAZABLES ""
        ' ***********************************************
        Dim RstAjusteTrazables As ADODB.Recordset
        ConsultaSql$ = "SELECT SUM(CANTINGRE) AS INGRESOS, SUM(CANTSALID) AS SALIDAS, MASTER.CODINT, MOVISTO.COD_INTE, MOVISTO.IDENLOTE, MOVISTO.DEPOMOVI "
        ConsultaSql$ = ConsultaSql$ & " FROM MASTER INNER JOIN MOVISTO WITH(NOLOCK) ON MASTER.CODINT = MOVISTO.COD_INTE WHERE MASTER.TRAZABLE > 0 "
        ConsultaSql$ = ConsultaSql$ & " AND MOVISTO.FECHMOV <= '" & FETEXCOR1$(iFechaCorte%) & "'"                          ' FECHA CORTE
        ConsultaSql$ = ConsultaSql$ & sFiltroDeposito$                                                                      ' DEPOSITO
        If sFiltroStatus$ <> "" Then ConsultaSql$ = ConsultaSql$ & sFiltroStatus$                                           ' STATUS ARTICULO
        If sFiltroArticulo$ <> "" Then ConsultaSql$ = ConsultaSql$ & sFiltroArticulo$ ' UN ARTICULO
        If NORFA$ <> "" Then
           ConsultaSql$ = ConsultaSql$ + " AND (MOVISTO.DOSECLAR = '" & NORFA$ & "' or MOVISTO.DOPRILAR = '" & NORFA$ & "' )"
        End If
        ConsultaSql$ = ConsultaSql$ + " AND MOVISTO.IDENLOTE <> ''"
        ConsultaSql$ = ConsultaSql$ & " GROUP BY MOVISTO.COD_INTE, MOVISTO.IDENLOTE, MOVISTO.DEPOMOVI, MASTER.CODINT "
        ConsultaSql$ = ConsultaSql$ & " ORDER BY MOVISTO.COD_INTE, MOVISTO.IDENLOTE, MOVISTO.DEPOMOVI"
        
      Set RstAjusteTrazables = READSET(ConsultaSql$)
        
        With RstAjusteTrazables
            Do While Not .EOF
                iCodigo% = !CODINT
                SLOTE$ = TRIM$(!idenlote)
                
                iDeposito% = !depomovi
                
                dIngresos# = XVALO(!INGRESOS)
                dSalidas# = XVALO(!SALIDAS)
                dSaldo# = dIngresos# - dSalidas#
                If dSaldo# > 0 Then ' si es > 0 va a ser un descuento, por eso valido que no sea superior al saldo del lote
                  SaldoDelLote# = SaldoLoteConDoseclar(CI1%, SLOTE$, iDeposito%, iFechaCorte%, , NORFA$)
                  If SaldoDelLote# >= 0 Then
                    If dSaldo# > SaldoDelLote# + 0.05 Then
                     dSaldo# = SaldoDelLote#
                     If dSaldo# <= 0 Then GoTo 80
                    End If
                  End If
                  
                End If
                dCantidadAjuste# = (-1) * dSaldo#   ' AJUSTA POR EL VALOR CONTRARIO
                
                If Abs(dCantidadAjuste) > 0.005 Then
                    iNroOrden% = iNroOrden% + 1
                    dValorUnitario# = 0
                    sReferencia$ = "Aj.Transf.Dep.2 " & NORFA$
                    iMoneda% = 0
                    iCliente% = 0
                    Call MOVISTOK(iFecha%, iCodigo%, SLOTE$, sCodMovimiento$, sDocPrimarioLargo$, sDocPrimarioCorto$, sDocSecundarioLargo$, sDocSecundarioCorto$, sReferencia$, iNroOrden%, dValorUnitario#, iMoneda%, iCliente%, iDeposito%, dCantidadAjuste#)
                End If
                
80             .MoveNext
            Loop
        End With
        
        On Error Resume Next
        RstAjusteTrazables.Close
        Set RstAjusteTrazables = Nothing
        On Error GoTo 0
    
        Screen.MousePointer = 1
' Else
'     GoTo 20
'   End If
   '
'   Call COMUNI("Ajuste a Cero Finalizado Correctamente")
   Screen.MousePointer = 1
99
'\\\OT-05-0704-WAR-FIN///
End Sub



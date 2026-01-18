VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CALICIER 
   BackColor       =   &H00C0D0D0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes de Calidad | Cierres de O-Fabricación"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0D0D0&
      Caption         =   "INFORMES DE CALIDAD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4560
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   9060
      Begin VB.TextBox TXTREFEREN 
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
         Height          =   330
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   720
         Width           =   1490
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
         Height          =   320
         Left            =   8730
         TabIndex        =   49
         Top             =   850
         Width           =   175
      End
      Begin VB.TextBox TXTFECHEMISION 
         Alignment       =   2  'Center
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
         Height          =   330
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   840
         Width           =   1050
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
         Height          =   320
         Left            =   6760
         TabIndex        =   36
         Top             =   330
         Width           =   175
      End
      Begin VB.TextBox TXTPENDCALIOF 
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
         Height          =   330
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   2160
         Width           =   1000
      End
      Begin VB.TextBox TXTPENDCALIDAD 
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
         Height          =   330
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   2160
         Width           =   1000
      End
      Begin VB.TextBox TXTRETRABAJO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7440
         TabIndex        =   41
         Top             =   2160
         Width           =   1000
      End
      Begin VB.TextBox TXTORDENCIERRE 
         Alignment       =   2  'Center
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
         Height          =   330
         Left            =   8220
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   360
         Width           =   690
      End
      Begin VB.TextBox TXTNORFA 
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
         Height          =   330
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   315
         Width           =   1490
      End
      Begin VB.TextBox TXTPIN 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7020
         TabIndex        =   32
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox TXTDEPOCONSU 
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
         Height          =   330
         Index           =   1
         Left            =   6780
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2640
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.TextBox TXTDEPOENTRADA 
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
         Height          =   330
         Index           =   1
         Left            =   2300
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   2640
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.CommandButton Command4 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   2100
         TabIndex        =   28
         Top             =   2655
         Visible         =   0   'False
         Width           =   175
      End
      Begin VB.TextBox TXTDEPOENTRADA 
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
         Height          =   330
         Index           =   0
         Left            =   1630
         TabIndex        =   1
         Top             =   2640
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.CommandButton Command3 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Left            =   6585
         TabIndex        =   26
         Top             =   2655
         Visible         =   0   'False
         Width           =   175
      End
      Begin VB.TextBox TXTDEPOCONSU 
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
         Height          =   330
         Index           =   0
         Left            =   6120
         TabIndex        =   2
         Top             =   2640
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.TextBox TXTOBSERVA 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   1125
         MaxLength       =   128
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   3240
         Width           =   7605
      End
      Begin VB.TextBox TXTCIERRAPOR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5280
         TabIndex        =   0
         Top             =   2160
         Width           =   1000
      End
      Begin VB.TextBox TXTPENDIENTE 
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
         Height          =   330
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1680
         Width           =   1000
      End
      Begin VB.TextBox TXTCANTICERRADA 
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
         Height          =   330
         Left            =   3135
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1680
         Width           =   1000
      End
      Begin VB.TextBox TXTCANTIPROY 
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
         Height          =   330
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1680
         Width           =   1000
      End
      Begin VB.TextBox TXTDESCRIPCION 
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
         Height          =   330
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1200
         Width           =   3975
      End
      Begin VB.TextBox TXTCODIGO 
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
         Height          =   300
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Orden de Compra"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4150
         TabIndex        =   51
         Top             =   660
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
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
         Index           =   14
         Left            =   7000
         TabIndex        =   48
         Top             =   960
         Width           =   585
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pend. Cierre OF ""Activa"""
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   9
         Left            =   2100
         TabIndex        =   46
         Top             =   2080
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Pend. de Calidad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   3
         Left            =   210
         TabIndex        =   44
         Top             =   2085
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Rechazadas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   6420
         TabIndex        =   42
         Top             =   2280
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Orden"
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
         Index           =   13
         Left            =   7300
         TabIndex        =   40
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cierres de O-Fabricación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4150
         TabIndex        =   38
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Identificador (PIN)"
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
         Index           =   12
         Left            =   5520
         TabIndex        =   33
         Top             =   1290
         Width           =   1425
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito Entrada"
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
         Index           =   11
         Left            =   -120
         TabIndex        =   29
         Top             =   2760
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito Consumo"
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
         Index           =   10
         Left            =   4440
         TabIndex        =   27
         Top             =   2760
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Observa- ciónes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   8
         Left            =   310
         TabIndex        =   24
         Top             =   3195
         Width           =   705
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Aprobadas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   4300
         TabIndex        =   22
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pendientes de Cierre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   6
         Left            =   4370
         TabIndex        =   20
         Top             =   1620
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad Cerrada"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Index           =   5
         Left            =   2250
         TabIndex        =   18
         Top             =   1605
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad a Fabricar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   4
         Left            =   230
         TabIndex        =   16
         Top             =   1605
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
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
         Index           =   1
         Left            =   100
         TabIndex        =   15
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label Label2 
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
         Height          =   210
         Index           =   0
         Left            =   560
         TabIndex        =   13
         Top             =   810
         Width           =   585
      End
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0D0D0&
      Caption         =   "HISTORIAL DE INFORMES DE LA OF's ACTIVA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   4800
      Width           =   9060
      Begin MSFlexGridLib.MSFlexGrid MSF2 
         Height          =   2475
         Left            =   150
         TabIndex        =   23
         Top             =   360
         Width           =   8760
         _ExtentX        =   15452
         _ExtentY        =   4366
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   8421504
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.TextBox Text7 
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
      Left            =   6960
      TabIndex        =   6
      Text            =   " "
      Top             =   600
      Width           =   1065
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   8460
      Left            =   9360
      ScaleHeight     =   8430
      ScaleWidth      =   1035
      TabIndex        =   5
      Top             =   0
      Width           =   1065
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
         Left            =   120
         Picture         =   "CALICIER.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   2280
         Width           =   650
      End
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
         Left            =   120
         Picture         =   "CALICIER.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Anulación Cierre de Orden de Fabricación"
         Top             =   3120
         Width           =   645
      End
      Begin VB.CommandButton BOTREC 
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
         Picture         =   "CALICIER.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Graba - Genera Cierre de Orden de Fabaricación"
         Top             =   3840
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
         Picture         =   "CALICIER.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salir"
         Top             =   200
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "CALICIER.frx":0A68
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CALICIER.frx":0C9C
      TabIndex        =   10
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nro. O/F.:"
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
      Left            =   5880
      TabIndex        =   7
      Top             =   210
      Width           =   1335
   End
End
Attribute VB_Name = "CALICIER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub CIERRECALIDAD(CI1%, NORFA$, CANTI, FEMI%)
    Call CREATABLA_CONTCALIDAD
    '
    SQLX$ = "SELECT * FROM CONTCALIDAD"
    SQLX$ = SQLX$ + " WHERE "
    
End Sub

Sub CREATABLA_CONTCALIDAD()
   Call ABRECONEXION
   Call CREACAMPO("", "CONTCALIDAD", "IdSubOr", 10, 12, 1)
   Call CREACAMPO("", "CONTCALIDAD", "OBSERVACIONES", 12, 0)
   Call CREACAMPO("", "CONTCALIDAD", "FechGen", 8, 0)
   Call CREACAMPO("", "CONTCALIDAD", "FECHCIERRE", 8, 0)
   Call CREACAMPO("", "CONTCALIDAD", "Cod_Inte", 3, 0, 1)
   Call CREACAMPO("", "CONTCALIDAD", "Cod_Exte", 10, 16)
   Call CREACAMPO("", "CONTCALIDAD", "Descrip", 10, 64)
   Call CREACAMPO("", "CONTCALIDAD", "CANTICONTROL", 7, 8)
   Call CREACAMPO("", "CONTCALIDAD", "IDCONTROL", 4, 0)
   Call CREACAMPO("", "CONTCALIDAD", "ORDEN", 3, 0)
   Call CREACAMPO("", "CONTCALIDAD", "USUARIO", 3, 0)
   Call CREACAMPO("", "CONTCALIDAD", "MARBORRA", 3, 0)
End Sub

Sub LimpiarObjetos()

    TXTNORFA = ""
    TXTORDENCIERRE = ""
    TXTREFEREN = ""
    TXTCODIGO = ""
    TXTDESCRIPCION = ""
    TXTPIN = ""
    TXTCANTIPROY = ""
    TXTCANTICERRADA = ""
    TXTPENDIENTE = ""
    TXTPENDCALIDAD = ""
    TXTPENDCALIOF = ""
    TXTCIERRAPOR = ""
    TXTRETRABAJO = ""
    TXTDEPOENTRADA(0) = ""
    TXTDEPOCONSU(0) = ""
    TXTOBSERVA = ""
    
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

Private Sub BOTREC_Click()

    BOTREC.Enabled = False
    If TRIM(TXTNORFA) = "" Then
        Call COMUNI("Falta Seleccionar Orden de Fabricación")
        GoTo 99
    End If
    '
    CANTI = XVALO(TXTCIERRAPOR) + XVALO(TXTRETRABAJO)
    If CANTI < 0.1 Then
        Call COMUNI("Falta Informar Cantidad a Cerrar y/o Por Retrabajo.")
        GoTo 99
    End If
    '
    If (XVALO(TXTCIERRAPOR) + XVALO(TXTRETRABAJO)) > XVALO(TXTPENDCALIOF) Then
        Call COMUNI("La Cantidad Aprobada más el Retrabajo\No Puede Superar el Pendiente del Cierre Activo.")
        GoTo 99
    End If
    '
    NumeroOrden% = XVALO(TXTORDENCIERRE)
    CodigoInterno% = CODINT%(TXTCODIGO)
    '
    Call BLOQUEAR_CONTROLES(Me, 2, 1, 1)
    
    FechaCalidad% = FECHANUM(TXTFECHEMISION)
    Call GrabaInformeDeCalidad(CodigoInterno%, FechaCalidad%, TXTNORFA, NumeroOrden%, TXTPIN, XVALO(TXTCIERRAPOR), XVALO(TXTRETRABAJO), TXTOBSERVA)
    Call ListaDeFabricacionesSinAprobar
    Call LimpiarObjetos
    
    Call BLOQUEAR_CONTROLES(Me, 1, 1, 1)
    '
99  BOTREC.Enabled = True

End Sub

Sub GrabaInformeDeCalidad(CI1%, FechaCalidad%, OrdenDeFabricacion$, NumeroOrden%, PIN$, CantidadAprobada#, CantidadRechazada#, Observaciones$)
    
    USERN% = USNBR% Mod 100
    
    ' *** REGISTRA EL INFORME DE CALIDAD
    Informe& = XVALO(VALOBADA("INFOCALIDAD", "MAX(NUMEROINFORME)", "NUMEROINFORME > 0")) + 1
    Dim RstCalidad As ADODB.Recordset
    Set RstCalidad = New ADODB.Recordset
    '
    ConsultaSql$ = "SELECT * FROM INFOCALIDAD WHERE NUMEROINFORME=0"
    RstCalidad.Open ConsultaSql$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With RstCalidad
        .AddNew
        !NumeroInforme = Informe&
        !FechaInforme = CVDAT(FechaCalidad%)
        !FechaReal = CVDAT(HOY(NO$))
        !NUMEROUSUARIO = USERN%
        !DocumentoOrigen$ = OrdenDeFabricacion$
        !NumeroDeControl$ = PIN$
        !CantidadAprobada = CantidadAprobada#
        !CantidadRechazada = CantidadRechazada#
        !Observaciones = Left$(Observaciones$, 128)
        !EstadoInforme = 1
        .Update
    End With
    RstCalidad.Close
    Set RstCalidad = Nothing
    
    
    '*** PROCESA LAS CANTIDADES APROBADAS
    If CantidadAprobada# > 0.05 Then
    
            On Error GoTo ProcesaCantidadesAprobadas
            
            Screen.MousePointer = 11
            Dim RstMOVITRANSTO As ADODB.Recordset
            Set RstMOVITRANSTO = New ADODB.Recordset
            RstMOVITRANSTO.Open "SELECT * FROM MOVITRANSTO WHERE DOPRILAR='" & OrdenDeFabricacion$ & "' AND OrdenRegistroMovimiento='" & CStr(NumeroOrden%) & "' AND CODIMOVI='OF'", FLEXCONN, adOpenKeyset, adLockPessimistic
            With RstMOVITRANSTO
                If Not (.EOF And .BOF) Then
                    '
                    '*** JUNTA LOS DATOS Y GENERA EL MOVIMIENTO DE STOCK DEL ARTICULO FABRICADO
                    REFERENCIA$ = TRIM$(SAFETEXT$(!ReferCor))
                    FechaMovimiento% = CVINT(!FechMov)
                    CodigoInterno% = XVALO(!cod_inte)
                    IdentificadorLote$ = SAFETEXT$(!idenlote)
                    CodigoMovimiento$ = SAFETEXT$(!CODIMOVI)
                    DocumentoPrimarioLargo$ = TRIM$(SAFETEXT$(!doprilar))
                    DocumentoPrimarioCorto$ = TRIM$(SAFETEXT$(!DoPriCor))
                    DocumentoSecundarioLargo$ = TRIM$(SAFETEXT$(!DOSECLAR))
                    DocumentoSecundarioCorto$ = TRIM$(SAFETEXT$(!DoSecCor))
                    DEPOSITO% = XVALO(!depomovi)
                    CantidadEntrada# = XVALO(!CANTINGRE)
                    CantidadSalida# = XVALO(!CANTSALID)
                    Cantidad# = CantidadEntrada# - CantidadSalida#
                    CoeficienteCierreSaldo = CantidadAprobada# / Cantidad#
                    OrdenItem% = 1
                    Moneda% = 1
                    codigoExterno$ = SAFETEXT$(!Cod_Exte)
                    
                    If DEPOSITO% = 99 Then DEPOSITO% = DEPOPRE%(CI2%)
                    If DEPOSITO% = 0 Then DEPOSITO% = 1
                    Call MOVISTOK(FechaMovimiento%, CodigoInterno%, IdentificadorLote$, CodigoMovimiento$, DocumentoPrimarioLargo$, DocumentoPrimarioCorto$, DocumentoSecundarioLargo$, DocumentoSecundarioCorto$, REFERENCIA$, OrdenItem%, ValorUnitario#, Moneda%, NumeroCliente%, DEPOSITO%, CantidadAprobada#, , , , , , , , , Informe&)
                    '
                    '*** ACTUALIZA LA CANTIDAD PENDIENTE DEL PRODUCTO FABRICADO SOBRE LA TABLA DE MOVIMIENTOS TRANSITORIOS
                    !CANTINGRE = !CANTINGRE - CantidadAprobada#
                    .Update
                    
                    '*** JUNTA LA INFORMACION PARA REGISTRAR EL CONSUMO DE LOS COMPONENTES DEL PROD. FABRICADO
                    OrdenItem% = 0
                    Dim RstComponentes As ADODB.Recordset
                    Set RstComponentes = New ADODB.Recordset
                    RstComponentes.Open "SELECT * FROM MOVITRANSTO WHERE DOPRILAR='" & OrdenDeFabricacion$ & "' AND OrdenRegistroMovimiento='" & CStr(NumeroOrden%) & "' AND CODIMOVI='CF'", FLEXCONN, adOpenKeyset, adLockPessimistic
                    With RstComponentes
                        Do While Not .EOF
                            '
                            DocumentoPrimarioLargo$ = !doprilar
                            DocumentoPrimarioCorto$ = !DoPriCor
                            DocumentoSecundarioLargo$ = !DOSECLAR
                            DocumentoSecundarioCorto$ = !DoSecCor
                            CodigoMovimiento$ = !CODIMOVI
                            DEPOSITO% = !depomovi
                            REFERENCIA$ = !ReferCor
                            FechaConsumo% = CVINT(!FechMov)
                            CodigoInterno% = !cod_inte
                            IdentificadorLote$ = !idenlote
                            CantidadSalida# = !CANTSALID
                            Cantidad# = CantidadSalida# * CoeficienteCierreSaldo
                            Moneda = 1
                            '
                            If TRIM$(UCase$(Unimed$(CI2%))) = "GK" Then
                              Cantidad# = Cantidad# / 1000
                            End If
                            '
                            If CodigoInterno% > 0 Then
                             If CodigoInterno% <= NUMAS% Then
                               If Abs(Cantidad#) >= 0.05 Then
                                 
                                 '*** GENERA EL MOVIMIENTO DE STOCK DE LOS COMPONENTES
                                 If DEPOSITO% = 99 Then DEPOSITO% = DEPOPRE%(CodigoInterno%)
                                 If DEPOSITO% = 0 Then DEPOSITO% = 1
144                              OrdenItem% = OrdenItem% + 1
                                 Call MOVISTOK(FechaConsumo%, CodigoInterno%, IdentificadorLote$, CodigoMovimiento$, DocumentoPrimarioLargo$, DocumentoPrimarioCorto$, DocumentoSecundarioLargo$, DocumentoSecundarioCorto$, REFERENCIA$, OrdenItem%, ValorUnitario#, Moneda%, NumeroCliente%, DEPOSITO%, Cantidad#, , , , , , , , , Informe&)
                                 '
                                 '*** ACTUALIZA MOVIMIENTO TRANSITORIO DE LAS RESERVAS
                                 !CANTSALID = !CANTSALID - Abs(Cantidad#)
                                 .Update
                                 '
                                 '*** ACTUALIZA LA RESERVA DE MATERIALES
                                 SALCON# = (-1) * Cantidad#
                                 SQLX$ = "SELECT * FROM Reserva "
                                 SQLX$ = SQLX$ + " WHERE IdSubOr = '" & DocumentoPrimarioLargo$ & "' "
                                 SQLX$ = SQLX$ + " AND Cod_Inte = " & CodigoInterno% & " "
                                 Dim RESETEMP As ADODB.Recordset
                                 Set RESETEMP = New ADODB.Recordset
                                 RESETEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                                 On Error Resume Next
                                 With RESETEMP
                                   .MoveFirst
                                   If Err < 1 Then
                                     On Error GoTo 0
                                     Do While Not .EOF
                                        CARESE# = XVALO(!CANTRES)
                                        CACOANT# = XVALO(!CANTCONS)
                                        If CARESE# - CACOANT# >= SALCON# Then
                                             !CANTCONS = CACOANT# + SALCON#
                                             SALCON# = 0
                                           Else
                                             SALCON# = SALCON# - (CARESE# - CACOANT#)
                                             !CANTCONS = XVALO(!CANTRES)
                                        End If
                                        .Update
                                        .MoveNext
                                     Loop
                                   End If
                                   On Error GoTo 0
                                 End With
                                 RESETEMP.Close
                                 Set RESETEMP = Nothing
                               
                               End If
                             End If
                            End If
                            '
                            .MoveNext
146                     Loop
                        '
                        RstComponentes.Close
                        Set RstComponentes = Nothing
                        '
                    End With
                    '
                    
                    '*** ACTUALIZA LA TABLA DE ORDENES DE FABRICACION
                    CIERRA% = 0
                    SQLX$ = " SELECT * FROM  ORDEFABR "
                    SQLX$ = SQLX$ + " WHERE IdSubOr = '" & DocumentoPrimarioLargo$ & "'"
                    Dim ABREORFFFAB As ADODB.Recordset
                    Set ABREORFFFAB = New ADODB.Recordset
                    ABREORFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                    On Error Resume Next
                    With ABREORFFFAB
                      .MoveFirst
                      If Err < 1 Then
                         On Error GoTo 0
                        !CanApro = !CanApro + CantidadAprobada#
                        If Abs(!CANPROY - !CanApro) < 0.05 Then
                          !StatuOrf = 3         ' MARCA COMO CERRADA
                          CIERRA% = 1
                        End If
                        !Fechcierr = CVDAT(FechaMovimiento%)
                        .Update
                      End If
                      On Error GoTo 0
                    End With
                    Set ABREORFFFAB = Nothing
                    '
                    If CIERRA% = 1 Then
                       ConsultaSql$ = "DELETE FROM Reserva WHERE IdSubOr = '" & NORFA$ & "'"
                       FLEXCONN.Execute ConsultaSql$
                       ConsultaSql$ = "DELETE FROM OPERFAB WHERE IdSubOr = '" & NORFA$ & "'"
                       FLEXCONN.Execute ConsultaSql$
                    End If
                End If
            End With
            RstMOVITRANSTO.Close
            Set RstMOVITRANSTO = Nothing
            
            '*** ELIMINA LOS MOVIMIENTOS TRANSITORIOS QUE NO TIENEN CANTIDAD. O SEA, QUE YA FUERON PASADOS A MOVISTO
            ConsultaSql$ = "DELETE FROM MOVITRANSTO WHERE (CantSalid + CantIngre < 0.05)"
            FLEXCONN.Execute ConsultaSql$
            
            Screen.MousePointer = 1
            
            On Error GoTo 0
    End If
    
    '*** PROCESA LAS CANTIDADES RECHAZADAS PARA RETRABAJO
    If CantidadRechazada# > 0.05 Then
        
            Dim RstRechaMOVITRANSTO As ADODB.Recordset
            Set RstRechaMOVITRANSTO = New ADODB.Recordset
            RstRechaMOVITRANSTO.Open "SELECT * FROM MOVITRANSTO WHERE DOPRILAR='" & OrdenDeFabricacion$ & "' AND OrdenRegistroMovimiento='" & CStr(NumeroOrden%) & "' AND CODIMOVI='OF'", FLEXCONN, adOpenKeyset, adLockPessimistic
            With RstRechaMOVITRANSTO
                If Not (.EOF And .BOF) Then
                    '
                    '*** ACTUALIZA LA CANTIDAD PENDIENTE DE APROBACION POR CALIDAD EN LA TABLA DE MOV TRANSITORIOS
                    Cantidad# = !CANTINGRE
                    CoeficienteRechazoSaldo = CantidadRechazada# / Cantidad#
                    !CANTINGRE = !CANTINGRE - CantidadRechazada#
                    !CantidadRechazada = XVALO(!CantidadRechazada) + CantidadRechazada#
                    .Update
                    
                    Dim RstComponent As ADODB.Recordset
                    Set RstComponent = New ADODB.Recordset
                    RstComponent.Open "SELECT * FROM MOVITRANSTO WHERE DOPRILAR='" & OrdenDeFabricacion$ & "' AND OrdenRegistroMovimiento='" & CStr(NumeroOrden%) & "' AND CODIMOVI='CF'", FLEXCONN, adOpenKeyset, adLockPessimistic
                    With RstComponent
                        Do While Not .EOF
                            '
                            Cantidad# = !CANTSALID
                            !CANTSALID = Cantidad# * CoeficienteRechazoSaldo
                            .Update
                            .MoveNext
                        Loop
                    End With
                    RstComponent.Close
                    Set RstComponent = Nothing
                 End If
            End With
            RstRechaMOVITRANSTO.Close
            Set RstRechaMOVITRANSTO = Nothing
    End If
    
    Call COMUNI("Informe de Calidad Registrado")
    
    Exit Sub
    
ProcesaCantidadesAprobadas:
    MensajeError$ = Err.Description
    On Error GoTo 0: Screen.MousePointer = 1
    Call MENSERR(24, "Error al Registrar Cantidades Aprobadas\" & MensajeError$)
    
    Exit Sub
        
'    Call ENRAPRODUCTO_TERMINADO(CI1%, NORFA$, CANTI, FECHACIERRE%, XVALO(Me.TXTDEPOENTRADA), "Fabricación Propia - Cierre Calidad")
'    '
'    CAN# = CANTI
'    Call CARCOMPO(CODEXT$(CI1%), NORFA$, CAN#, 0, 1)
'
'    Call CONSUCOMPONETESXOF(CI1%, NORFA$, CANTI, FECHACIERRE%, XVALO(Me.TXTDEPOCONSU))
'    '
'    'ACTUALIZA LA RESERVA
'    FIN& = ULTREG&("!MATEROF")
'    For U& = 1 To FIN&
'      Call TRACE(20, U&, FIN&)
'      ZZ$ = REGLEIDO$("!MATEROF", U&)
'      CICOMPO% = CVI(Left$(ZZ$, 2))
'      CANCONSUM = CVS(Mid$(ZZ$, 5, 4))
'      If TRIM$(UCase$(UNIMED$(CICOMPO%))) = "GK" Then CANCONSUM = CANCONSUM / 1000
'      '
'      FECHA% = FECHACIERRE%
'      Call ACTURESERVA(CICOMPO%, NORFA$, CANCONSUM)
'    Next U&
'    '
'    '
'    TXTOVARIABLE$ = "Cierre de O/Fabricación Desde Calidad"
'    NORFAX$ = TRIM$(NORFA$)
'    TTXYZ$ = Space$(76)
'    FE% = FECHACIERRE%
'    Call REPLA(TTXYZ$, FECHATEX$(FE%), 1, 8)
'    Call REPLA(TTXYZ$, TXTOVARIABLE$, 11, 56)
'    Call REPLA(TTXYZ$, USERTER$, 69, 24)
'    Call AGREGA_ANOTAORFAB(NORFAX$, TTXYZ$)
'    '
'    TTXYZ$ = Space$(76)
'    Call REPLA(TTXYZ$, " - Cantidad Aprob.= " & CANTI, 11, 56)
'    Call AGREGA_ANOTAORFAB(NORFA$, TTXYZ$)
'    Call AGREGA_ANOTAORFAB(NORFA$, String$(88, "="))
'    '
'    Call ACTUALIZAORFA(CI1%, NORFA$, CANTI, FECHACIERRE%, Me.TXTOBSERVA)
    '
    
    
End Sub
Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command16_Click()

End Sub

Private Sub Command2_Click()

    Command2.Enabled = False
    
50  Call SELECHO("!ORFABAP/Ordenes de Fabricación Pendientes de Aprobación")
    IDENTIFICADOR$ = TRIM$(VALACT1$("!ORFABAP"))
    If IDENTIFICADOR$ = "" Then GoTo 99
    
    PosicionCaracterControl% = InStr(1, IDENTIFICADOR$, "|")
    OrdenFabricacion$ = TRIM$(Mid$(IDENTIFICADOR$, 1, PosicionCaracterControl% - 1))
    OrdenCierre& = XVALO(Mid$(IDENTIFICADOR$, PosicionCaracterControl% + 1))
    ReferenciaOc$ = TRIM$(Mid$(VALACT2$("!ORFABAP"), 17, 16))
    
    TXTNORFA = ""
    TXTORDENCIERRE = CStr(OrdenCierre&)
    TXTREFEREN = ReferenciaOc$
    TXTNORFA = OrdenFabricacion$
    
99  Command2.Enabled = True

End Sub

Private Sub Command29_Click()
  '
  Call OPNOIN("Ayuda FLEXOFT en Línea")
  '
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call SELECHO("TADEPOSI")
    If VALACT1$("TADEPOSI") <> "" Then
      Me.TXTDEPOCONSU(0) = TRIM$(VALACT1$("TADEPOSI"))
    End If
99  Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    '
    Command4.Enabled = False
    Call SELECHO("TADEPOSI")
    If VALACT1$("TADEPOSI") <> "" Then
      Me.TXTDEPOENTRADA(0) = TRIM$(VALACT1$("TADEPOSI"))
    End If
    
99  Command4.Enabled = True
    '
End Sub

Private Sub Command45_Click()
   Command45.Enabled = False
   Call OFABRI07.ANUCIEOF
   Call TXTNORFA_Change
   Command45.Enabled = True

End Sub

Private Sub Command6_Click()
  '
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
     Text6.TEXT = VDEV$
  End If
   '
End Sub

Private Sub Command7_Click()
  '
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
     Text5.TEXT = VDEV$
  End If
  '
End Sub

Private Sub Command5_Click()
  '
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
     TXTFECHEMISION.TEXT = VDEV$
  End If
  '
End Sub

Sub Form_Load()
    '
    Call ABRECALIDAD
    Screen.MousePointer = 11
    
    HOII% = HOY(HOS$)
    TXTFECHEMISION = FECHATEX$(HOII%)
    Call ListaDeFabricacionesSinAprobar
    Campos$ = "Fecha/d8;Documento/A15;Referencia/A28;Dep./F4.0;Cant.Apro./F10.1"
    Call CARGA_ENCABEZADO(Me.MSF2, Campos$, CALICIER)
    '
    Screen.MousePointer = 1
    '
End Sub

Private Sub Text5_DblClick()
  '
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
     Text5.TEXT = VDEV$
  End If
  '
End Sub



Private Sub TXTCIERRAPOR_KeyPress(KeyAscii As Integer)

    ' ACEPTA SOLO NUMEROS Y SEPARADOR DECIMAL
    CARACTER$ = Chr$(KeyAscii)
    If IsNumeric(CARACTER$) = False And KeyAscii <> 8 And KeyAscii <> 46 Then
        KeyAscii = 0
    End If
    
    ' ACEPTA SOLO 1 SEPARADOR DECIMAL
    If CARACTER$ = "." And InStr(1, TXTRETRABAJO, ".", vbTextCompare) > 0 Then
        KeyAscii = 0
    End If
    
End Sub


Private Sub TXTDEPOCONSU_Change(Index As Integer)
   DEPOMO% = XVALO(Me.TXTDEPOCONSU(0))
   Me.TXTDEPOCONSU(1) = ECOARCH("TADEPOSI", Chr$(DEPOMO%))

End Sub

Private Sub TXTDEPOCONSU_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      Me.TXTOBSERVA.SetFocus
      Call POSICIONARALFINALTXT(TXTOBSERVA)
   End If

End Sub

Private Sub TXTDEPOENTRADA_Change(Index As Integer)

   DEPOMO% = XVALO(Me.TXTDEPOENTRADA(0))
   Me.TXTDEPOENTRADA(1) = ECOARCH("TADEPOSI", Chr$(DEPOMO%))

End Sub


Private Sub TXTDEPOENTRADA_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii = 13 Then
      TXTDEPOCONSU(0).SetFocus
   End If

End Sub


Private Sub TXTNORFA_Change()
    
    On Error GoTo CargaOFSinAprobar
    
    NORFA$ = TRIM$(TXTNORFA)
    TXTCIERRAPOR = "": TXTOBSERVA = ""
        
    If NORFA$ = "" Then GoTo 99
    
    SQLX$ = "SELECT Cod_Inte, Cod_Exte, Descrip, FechMov, CantIngre"
    SQLX$ = SQLX$ + " FROM MOVITRANSTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE DOPRILAR = '" & NORFA$ & "' AND OrdenRegistroMovimiento=" & TXTORDENCIERRE
    Dim RstMOVITRANSTO As ADODB.Recordset
    Set RstMOVITRANSTO = READSET(SQLX$)
    With RstMOVITRANSTO
        If Not (.EOF And .BOF) Then
            TXTCODIGO = SAFETEXT$(!Cod_Exte)
            TXTDESCRIPCION = SAFETEXT$(!Descrip)
            .MoveNext
        End If
    End With
        
    RstMOVITRANSTO.Close
    Set RstMOVITRANSTO = Nothing
    
    ConsultaSql$ = "IDSUBOR='" & NORFA$ & "'"
    CantidadProyectada# = XVALO(VALOBADA("ORDEFABR", "CanProy", ConsultaSql$))
    CantidadCerrada# = XVALO(VALOBADA("ORDEFABR", "CanApro", ConsultaSql$))
    TXTCANTIPROY = TRIM$(FORMATNUM$(CantidadProyectada#, "F12.1"))
    TXTCANTICERRADA = TRIM$(FORMATNUM$(CantidadCerrada#, "F12.1"))
    PENDIENTEX = CantidadProyectada# - CantidadCerrada#
    If PENDIENTEX < 0 Then PENDIENTEX = 0
    TXTPENDIENTE = TRIM$(FORMATNUM$(PENDIENTEX, "F12.1"))
    PendientesDeCalidad# = XVALO(VALOBADA("MOVITRANSTO", "SUM(CANTINGRE)", "DOPRILAR='" & NORFA$ & "'"))
    TXTPENDCALIDAD = TRIM$(FORMATNUM$(PendientesDeCalidad#, "F12.1"))
    OrdenMovimiento% = XVALO(TXTORDENCIERRE)
    PendientesDeCalidad# = XVALO(VALOBADA("MOVITRANSTO", "SUM(CANTINGRE)", "DOPRILAR='" & NORFA$ & "' AND OrdenRegistroMovimiento='" & OrdenMovimiento% & "'"))
    TXTPENDCALIOF = TRIM$(FORMATNUM$(PendientesDeCalidad#, "F12.1"))
    
    Call POSICIONARALFINALTXT(TXTOBSERVA)
    
    On Error GoTo 0

99  Exit Sub

CargaOFSinAprobar:

    MensageError$ = Err.Description
    On Error GoTo 0
    Call MENSERR(24, "Problema al Carga Datos de la OF's" & vbCrLf & MensageError$)
    
    On Error Resume Next
    RstMOVITRANSTO.Close
    Set RstMOVITRANSTO = Nothing
    On Error GoTo 0
    Exit Sub

End Sub

Sub ListaDeFabricacionesSinAprobar()

    '*** CARGO LA LISTA DE SELECCION ORFABAP DESDE LA TABLA MOVITRANSTO CON LA OF PENDIENTES DE APROBACION
    Call SETULTREG("!ORFABAP", 0)
    
    On Error GoTo ListaPendientesAprobacion
    
    Dim RstMOVITRANSTO As ADODB.Recordset
    Set RstMOVITRANSTO = READSET("SELECT * FROM MOVITRANSTO WITH(NOLOCK) WHERE CodiMovi='OF'")
    
    With RstMOVITRANSTO
        Do While Not .EOF
            ReferenciaOc$ = SAFETEXT$(!DOCUORIG)
            DocumentoPrimario$ = SAFETEXT$(!doprilar)
            DescripcionCodigo$ = SAFETEXT$(!Descrip)
            CantidadIngresada# = XVALO(!CANTINGRE)
            NumeroOrden& = XVALO(!OrdenRegistroMovimiento)
            IDENTIFICADOR$ = TRIM$(DocumentoPrimario$) & "|" & CStr(NumeroOrden&)
            
            Call REPLA(TTXX$, IDENTIFICADOR$, 1, 16)
            Call REPLA(TTXX$, DocumentoPrimario$, 18, 16)
            Call REPLA(TTXX$, ReferenciaOc$, 36, 16)
            Call REPLA(TTXX$, DescripcionCodigo$, 52, 24)
            Call REPLA(TTXX$, FORMATNUM$(CantidadIngresada#, "F10.1"), 78, 10)
            Call REGAPP("!ORFABAP", TTXX$)
            .MoveNext
        Loop
    End With
   
    RstMOVITRANSTO.Close
    Set RstMOVITRANSTO = Nothing
                
    
    Call CIERRARCH("!ORFABAP")
    Call FRESHECHO("!ORFABAP")
    On Error GoTo 0
    
    Exit Sub

ListaPendientesAprobacion:

    MensageError$ = Err.Description
    On Error GoTo 0
    Call MENSERR(24, "Problema al Generar Lista de OF's Sin Aprobar" & vbCrLf & MensageError$)
    
    On Error Resume Next
    RstMOVITRANSTO.Close
    Set RstMOVITRANSTO = Nothing
    Call SETULTREG("!ORFABAP", 0)
    Call CIERRARCH("!ORFABAP")
    Call FRESHECHO("!ORFABAP")
    On Error GoTo 0
    Exit Sub

End Sub

Private Sub TXTRETRABAJO_KeyPress(KeyAscii As Integer)

    ' ACEPTA SOLO NUMEROS Y SEPARADOR DECIMAL
    CARACTER$ = Chr$(KeyAscii)
    If IsNumeric(CARACTER$) = False And KeyAscii <> 8 And KeyAscii <> 46 Then
        KeyAscii = 0
    End If
    
    ' ACEPTA SOLO 1 SEPARADOR DECIMAL
    If CARACTER$ = "." And InStr(1, TXTRETRABAJO, ".", vbTextCompare) > 0 Then
        KeyAscii = 0
    End If
    
End Sub



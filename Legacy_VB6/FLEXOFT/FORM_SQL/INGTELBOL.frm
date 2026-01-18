VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form INGTELBOL 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe de Asignación de Telas para Bolsas "
   ClientHeight    =   7890
   ClientLeft      =   1050
   ClientTop       =   855
   ClientWidth     =   15720
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   15720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Transferir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   14280
      TabIndex        =   43
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   1305
      Left            =   120
      TabIndex        =   27
      Top             =   840
      Width           =   15495
      Begin VB.CommandButton Command4 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7200
         Picture         =   "INGTELBOL.frx":0000
         TabIndex        =   59
         Top             =   900
         Width           =   300
      End
      Begin VB.TextBox TXTPESOORIG 
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
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   915
         Width           =   1365
      End
      Begin VB.TextBox TXTNUMEOPER 
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
         Left            =   14160
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   516
         Width           =   456
      End
      Begin VB.TextBox TXTMETROS 
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
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   915
         Width           =   1125
      End
      Begin VB.CommandButton command15 
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
         Height          =   330
         Left            =   1710
         Picture         =   "INGTELBOL.frx":030A
         TabIndex        =   46
         Top             =   920
         Width           =   175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   14280
         TabIndex        =   42
         Top             =   915
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
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
         Height          =   330
         Left            =   15240
         TabIndex        =   40
         ToolTipText     =   "Despliega Lista de Selección"
         Top             =   120
         Width           =   160
      End
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
         Height          =   330
         Left            =   14160
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   150
         Width           =   1056
      End
      Begin VB.TextBox TXTNECESIDAD 
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
         Left            =   10920
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   915
         Width           =   1650
      End
      Begin VB.TextBox TXTPESO 
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
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   520
         Width           =   1125
      End
      Begin VB.TextBox TXTLOTE 
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
         Left            =   10935
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   520
         Width           =   1630
      End
      Begin VB.TextBox TXTDESCRIELEM 
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
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   150
         Width           =   4170
      End
      Begin VB.TextBox TXTCODIGOELEM 
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
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   150
         Width           =   1530
      End
      Begin VB.TextBox TXTLECTURA 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   120
         Width           =   2355
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Peso Original:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   4375
         TabIndex        =   58
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
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
         Height          =   330
         Left            =   1920
         TabIndex        =   56
         Top             =   555
         Width           =   2325
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "N° Operación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   11
         Left            =   12720
         TabIndex        =   55
         Top             =   510
         Width           =   1335
      End
      Begin VB.Label lbldepogeneral 
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
         Height          =   330
         Left            =   1200
         TabIndex        =   53
         Top             =   555
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep. General"
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
         Left            =   -720
         TabIndex        =   52
         Top             =   600
         Width           =   1890
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Metros:"
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
         Index           =   10
         Left            =   7650
         TabIndex        =   51
         Top             =   960
         Width           =   735
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
         Height          =   330
         Left            =   1920
         TabIndex        =   49
         Top             =   915
         Width           =   2325
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dep. Corte"
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
         Left            =   -720
         TabIndex        =   48
         Top             =   960
         Width           =   1890
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
         Height          =   330
         Left            =   1200
         TabIndex        =   47
         Top             =   915
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desc.Operación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   12360
         TabIndex        =   41
         Top             =   150
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Necesid.(mts.):"
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
         Index           =   8
         Left            =   9600
         TabIndex        =   38
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Peso Actual:"
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
         Index           =   3
         Left            =   7230
         TabIndex        =   36
         Top             =   525
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
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
         Index           =   2
         Left            =   10440
         TabIndex        =   34
         Top             =   525
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción:"
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
         Index           =   9
         Left            =   7245
         TabIndex        =   32
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código Elemento:"
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
         Index           =   5
         Left            =   3480
         TabIndex        =   30
         Top             =   150
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   ">>>>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   28
         Top             =   156
         Width           =   612
      End
   End
   Begin VB.TextBox txtcodintelem 
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
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7440
      Width           =   1530
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "Cód.Interno"
      Top             =   7440
      Width           =   1140
   End
   Begin VB.TextBox TXTCODINT 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   7440
      Width           =   1155
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   15495
      Begin VB.CommandButton Command14 
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2640
         TabIndex        =   60
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox TXTFTECNICA 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   14205
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   420
         Width           =   1140
      End
      Begin VB.TextBox TXTNROCOTIZA 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   13060
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   420
         Width           =   1140
      End
      Begin VB.TextBox TXTNROPED 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   11920
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   420
         Width           =   1140
      End
      Begin VB.TextBox TXTCODIGO 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   2860
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   420
         Width           =   2235
      End
      Begin VB.TextBox TXTCANAPRO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   10790
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   " "
         Top             =   420
         Width           =   1140
      End
      Begin VB.TextBox TXTCANPROY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   9650
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   " "
         Top             =   420
         Width           =   1140
      End
      Begin VB.TextBox TXTDESCRIPCION 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   5100
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   " "
         Top             =   420
         Width           =   4575
      End
      Begin VB.TextBox Text14 
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
         Left            =   9650
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "Cant.Proyec."
         Top             =   150
         Width           =   1140
      End
      Begin VB.TextBox Text2 
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
         Index           =   0
         Left            =   5100
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "Descripción"
         Top             =   150
         Width           =   4575
      End
      Begin VB.TextBox TXTOF 
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   315
         Width           =   1170
      End
      Begin VB.CommandButton Command6 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2370
         TabIndex        =   24
         ToolTipText     =   "Despliega Lista de Selección"
         Top             =   315
         Width           =   160
      End
      Begin VB.TextBox Text14 
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
         Index           =   6
         Left            =   14205
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "F.Técnica"
         Top             =   150
         Width           =   1140
      End
      Begin VB.TextBox Text14 
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
         Index           =   5
         Left            =   13060
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "Cotización"
         Top             =   150
         Width           =   1140
      End
      Begin VB.TextBox Text14 
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
         Index           =   3
         Left            =   11920
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Pedido"
         Top             =   150
         Width           =   1140
      End
      Begin VB.TextBox Text13 
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
         Left            =   2860
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Código"
         Top             =   150
         Width           =   2235
      End
      Begin VB.TextBox Text14 
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
         Index           =   2
         Left            =   10790
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "Cant.Aprob."
         Top             =   150
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Orden de Fabricación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF0 
      Height          =   2205
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   15465
      _ExtentX        =   27279
      _ExtentY        =   3889
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1440
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "INGTELBOL.frx":0614
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   3240
      OleObjectBlob   =   "INGTELBOL.frx":0848
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2445
      Left            =   120
      TabIndex        =   44
      Top             =   4920
      Width           =   15465
      _ExtentX        =   27279
      _ExtentY        =   4313
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      Redraw          =   -1  'True
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Consumos  (O.F.Activa)"
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
      Index           =   7
      Left            =   120
      TabIndex        =   45
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CI Elemento:"
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
      Index           =   6
      Left            =   3360
      TabIndex        =   23
      Top             =   7440
      Width           =   1575
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "INGTELBOL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function OperacionesDelCodigoLeido$(CIELEM%, Optional ret_CantRegistros&)
   CI1% = XVALO(TXTCODINT)
   Dim ctp As New Conteplast
   OperacionesDelCodigoLeido$ = ""
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "Operación/A30;Código./A18;Descripción/A32;Formulario/A12;Necesidad;U.Med./A6; CI/f6;Nop/F4", FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Operación/A30;Código./A18;Descripción/A32;Formulario/A12;Necesidad;U.Med./A6; CI/f6;Nop/F4", FRMZELECHO)
   REG& = 0
   SQLX$ = "SELECT * FROM FORMULAS WITH(NOLOCK) WHERE CODICONJ =" & CI1%
   SQLX$ = SQLX$ + " AND (CODIELEM = " & CIELEM%
   CIRANGO& = CODIGORANGO&(CIELEM%)
   If CIRANGO& > 0 Then
      SQLX$ = SQLX$ + " OR CODIELEM = " & CIRANGO&
   End If
   SQLX$ = SQLX$ + " )"
   SQLX$ = SQLX$ + " ORDER BY NUORIT"
   Set rs = READSET(SQLX$)
   With rs
     Do While Not .EOF
         REG& = REG& + 1
         CIELEM% = XVALO(!codielem)
         FRMZELECHO.MSF2.Rows = REG& + 1
         NOMBRESECTOR$ = SAFETEXT$(!COMPONENTE)
         NOP% = ctp.NumeOperXNombre(CI1%, NOMBRESECTOR$)
         FRMZELECHO.MSF2.TextMatrix(REG&, 1) = NOMBRESECTOR$
         FRMZELECHO.MSF2.TextMatrix(REG&, 2) = SAFETEXT$(!CodXElem)
         FRMZELECHO.MSF2.TextMatrix(REG&, 3) = SAFETEXT$(DESCRIT0$(CIELEM%))
         FRMZELECHO.MSF2.TextMatrix(REG&, 4) = ctp.FRMSECTOR(NOMBRESECTOR$)
         FRMZELECHO.MSF2.TextMatrix(REG&, 5) = XVALO(!usobruto) * XVALO(TXTCANPROY)
         FRMZELECHO.MSF2.TextMatrix(REG&, 6) = SAFETEXT$(!UMED_FORM)
         FRMZELECHO.MSF2.TextMatrix(REG&, 7) = CIELEM%
         FRMZELECHO.MSF2.TextMatrix(REG&, 8) = NOP%
         
         'SI LA UNIDAD DE CONSUMO ES
'         If TRIM$(UCase$(SAFETEXT$(!UMED_FORM))) = "CM" Then FRMZELECHO.msf2.TextMatrix(REG&, 6) = XVALO(!USOBRUTO) / 100
         
       .MoveNext
     Loop
   End With
   rs.Close
   Set rs = Nothing
   If REG& > 1 Then 'si son mas de un registro muestra la lista
     FRMZELECHO.Show 1
     OperacionesDelCodigoLeido$ = RESP_ZELE_VECT(1)
     TXTNUMEOPER = RESP_ZELE_VECT(8)
     Me.TXTNECESIDAD = RESP_ZELE_VECT(5)
     Exit Function
   End If
   'si es un solo registro llega aqui y asigna directamente
   OperacionesDelCodigoLeido$ = NOMBRESECTOR$
   Me.TXTNECESIDAD = FRMZELECHO.MSF2.TextMatrix(REG&, 5)
   ret_CantRegistros& = REG&
   'EL CONSUMO TOTAL DEBERIA SER EL CORTE POR MODXBB

End Function

Private Sub Command1_Click()
    command1.Enabled = False
    If MSF0.Rows > 1 Then
       OPX% = COMALTER%("Hay Datos Cargados en Grilla Que No Se Han Transferido\\Cancelar\Transferir Primero Datos Ingresados")
       If OPX% = 2 Then
         Command3_Click
         GoTo 99
       Else
         GoTo 99
       End If
    End If
    If XVALO(txtcodintelem) < 1 Then
       Call COMUNI("No Hay Ninguna Lectura de Bobina Activa")
       GoTo 99
    End If
    
    ORD% = MSF0.Rows
    MSF0.Rows = MSF0.Rows + 1
    
    MSF0.TextMatrix(ORD%, 1) = TRIM$(Me.TXTCODIGOELEM)
    MSF0.TextMatrix(ORD%, 2) = TRIM$(DESCRIT0$(XVALO(Me.txtcodintelem)))
    MSF0.TextMatrix(ORD%, 3) = TRIM$(TRIM$(Me.TXTLOTE))
    MSF0.TextMatrix(ORD%, 4) = TRIM$(TRIM$(Me.TXTPESO))
    MSF0.TextMatrix(ORD%, 5) = TRIM$(TRIM$(Me.TXTNECESIDAD))
    MSF0.TextMatrix(ORD%, 6) = TRIM$(TRIM$(Me.TXTMETROS))
    MSF0.TextMatrix(ORD%, 7) = TRIM$(TRIM$(Me.TXTNUMEOPER))
    
    Call BLANFORMU2
99  command1.Enabled = True

End Sub

Private Sub Command14_Click()
   TXTOF = OrdenFabricacionAbiertas$(1)
End Sub

Private Sub Command4_Click()
   IDLOT$ = TRIM$(TXTLOTE.TEXT)
   If IDLOT$ <> "" Then
     On Error Resume Next
     CONSUMLOT07.Label4 = TXTCODIGOELEM
     If Err Then Exit Sub
     On Error GoTo 0
     CONSUMLOT07.Label1 = IDLOT$
     CONSUMLOT07.Show 1
   End If

End Sub

Private Sub Label13_Change()
   DEPOGEN% = XVALO(Label13)
   Label6 = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPOGEN%)))
End Sub

Private Sub lbldepogeneral_Change()
   DEPOCOR% = XVALO(lbldepogeneral)
   Label3 = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPOCOR%)))

End Sub

Private Sub MSF0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
      REG& = MSF0.MouseRow
      If MSF0.Rows > 1 Then
        If REG& >= 1 Then
            Call SELECCIONARFILA(Me.MSF0, REG&)
            OPX% = COMALTER("Realmente Desea Eliminar el Registro Seleccionado\\Cancelar\Sí, Eliminar")
            If OPX% = 2 Then
              Call cmdEliminaItem(MSF0, REG&)
            End If
        End If
      End If
    End If
End Sub

Sub BLANFORMU2()
        TXTCODIGOELEM = ""
        TXTPESO = ""
        TXTDESCRIELEM = ""
        TXTNECESIDAD = ""
        TXTMETROS = ""
        TXTLOTE = ""
        Text1 = ""
        TXTLECTURA = ""
        txtcodintelem = ""
End Sub
Private Sub Command15_Click()
       Call SELECHO("TADEPOSI")
        If VALACT1$("TADEPOSI") <> "" Then
            Label13 = TRIM$(VALACT1$("TADEPOSI"))
        End If
End Sub

Private Sub Command2_Click()
    Text1 = OperacionesDelCodigoLeido(XVALO(txtcodintelem))
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    'CONSUMO
    FECHA% = HOY(HO$)
    CMOV$ = "TR"
    
    DEPOCONSUM% = XVALO(Me.Label13)
    DEPOALMACEN% = XVALO(Me.lbldepogeneral)
    For REG& = 1 To MSF0.Rows - 1
          CI1% = CODINT%(MSF0.TextMatrix(REG&, 1))
          CANTI = XVALO(MSF0.TextMatrix(REG&, 6))
          LOTE$ = MSF0.TextMatrix(REG&, 3)
          If CI1% < 1 Then
              Call MENSERR(24, "Falta Código o Código Inválido en Fila: " & REG&)
              GoTo 99
          End If
          If CANTI < 0.05 Then
              Call MENSERR(24, "Falta Cantidad en Fila: " & REG&)
              GoTo 99
          End If
          If LOTE$ = "" Then
              Call MENSERR(24, "Falta Lote en Fila: " & REG&)
              GoTo 99
          End If
    Next REG&
    
'    On Error GoTo ERROR_GUARDAR
'    FLEXCONN.BeginTrans
    
    For REG& = 1 To MSF0.Rows - 1
      CI1% = CODINT%(MSF0.TextMatrix(REG&, 1))
      CANTI = XVALO(MSF0.TextMatrix(REG&, 6))
      LOTE$ = UCase$(MSF0.TextMatrix(REG&, 3))
      NCOMP& = ProNroComprobante(CMOV$)
      DOPRI$ = "TR." + TRIM$(str$(NCOMP&))
      DOSEC$ = TRIM$(TXTOF)
      NUMEOPERACION% = XVALO(MSF0.TextMatrix(REG&, 7))
      DESCRIDEPO$ = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPOALMACEN%)))
      REFE$ = "Tr.de " & TRIM$(DESCRIDEPO$) & "-" & TRIM$(TXTOF)
      Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC%, DEPOALMACEN%, (-1) * CANTI, , , , , , , , , , NUMEOPERACION%)   'GUARDO EL NUMERO DE OPERACION DE CADA ITEMS TRANSFERIDO

      DESCRIDEPO$ = ECOARCH$("TADEPOSI", Chr$(DEPOCONSUM%))
      REFE$ = "Tr.A " & TRIM$(DESCRIDEPO$) & "-" & TRIM$(TXTOF)
      NCOMP& = ProNroComprobante(CMOV$)

      Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC%, DEPOCONSUM%, CANTI, , , , , , , , , , NUMEOPERACION%)  'GUARDO EL NUMERO DE OPERACION DE CADA ITEMS TRANSFERIDO
      Call ACSALOTE(CI1%, LOTE$, "NOMESS")
    Next REG&
   
'ERROR_GUARDAR:
'    If Err <> 0 Then
'       FLEXCONN.RollbackTrans
'       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
'       GoTo 99
'    Else
'       FLEXCONN.CommitTrans
'       GoTo 99
'    End If
    DEPOCORTE% = XVALO(Control("", "DEPODEST"))
    
    Dim ctp As New Conteplast
    Call ctp.CargaMatTranfXOrden(MSF1, TRIM$(TXTOF), DEPOCONSUM%)
    Set ctp = Nothing

    MSF0.Rows = 1

    
99 Screen.MousePointer = 1
   Command3.Enabled = True
End Sub

Private Sub Command6_Click()
   TXTOF = OrdenFabricacionAbiertas$
End Sub

Private Sub Form_Load()
      Campos$ = "Código/A16;Descripción/A32;Lote/A16;Peso/F10.2;Neces.M./F12.2;Bobina M./f12.2;NADA/a0"
      Call CARGA_ENCABEZADO(Me.MSF0, Campos$, Me)
      
      Call CARGA_ENCABEZADO(MSF1, "CODIGO/A20;Descripción/A40;LOTE/A16;Cons.Mts./f10.1;CI/f6;Op.N°/F6;Kg.Transf/F9.2", Me)
      
      DEPOMATBOL% = XVALO(Control("CONTEPLA", "DEMATBOL"))
      If DEPOMATBOL% < 1 Then
           Call COMUNI("Debe Configurar Depósito de Bolsas")
           Exit Sub
      End If
      lbldepogeneral = Control$("", "DEPOCEN")
      Label13 = DEPOMATBOL%
End Sub


Sub CARGAMSF0(NORFA$)
   MSF0.Rows = 1
   '
16 SQLX$ = "SELECT IDTEXT,DOPRICOR, CANTINGRE, CODIMOVI,COD_INTE  FROM  ORFAPEN WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE (CODIMOVI = 'OF') AND IDSITUACION = 'GenCierre' AND DOPRICOR = '" & TRIM$(NORFA$) & "'"
   SQLX$ = SQLX$ + " ORDER BY FEREMOVI "
   Set rs = READSET(SQLX$)
   With rs
   Do While Not .EOF
     Call TRACE(20, U&, FIN&)
     CIII% = XVALO(!cod_inte)
     DOPRI$ = SAFETEXT$(!DoPriCor)
     CODI$ = VectArticulos.Codigo(CIII%)
     DESCRI$ = VectArticulos.Descripcion(CIII%)
     CANTIX$ = FORMATNUM$(XVALO(!CANTINGRE), "F12.1")
      J& = J& + 1

     MSF0.Rows = J& + 1
     MSF0.TextMatrix(J&, 1) = CIII%
     MSF0.TextMatrix(J&, 2) = SAFETEXT$(!IDTEXT)
     MSF0.TextMatrix(J&, 3) = CODI$
     MSF0.TextMatrix(J&, 4) = DESCRI$
     MSF0.TextMatrix(J&, 5) = CANTIX$
 
20   .MoveNext
   Loop
   End With
   rs.Close
   Set rs = Nothing
   
   '
   Screen.MousePointer = 1
   '

End Sub
'CUANDO SE LEE CON LECTORA ESCRIBE EL CODIGO Y EL LOTE
'TRAE EL PESO VER SI ESTE VA A QUEDAR SIEMPRE ACTUALIZADO
'TENER EN CUENTA QUE SIEMPRE VA A  SERE UN CODIGO REAL
'TAL VEZ NO HAGA FALTA MOSTRAR LAS OPERACIONES O SOLO PARA VISUALIZARLAS
'Y QUE LAS VALIDACIONES SEAN SOBRE EL TOTAL DE LOS CODIGOS Y SI NO QUE SEAN PARTE DE LOS RANGO.
'UNA VEZ QUE SE ACETPA BAJAR A LA GRILLA Y REALIZAR EL MOVIMIENTO DE STOCK COMO TRANSFERENCIA
'SI SE QUITA VERIFICAR QUE NO HAYA CONSUMOS Y EN ESE CASO REALIZAR UN MOVIMIENTO INVERSO O QUE ANULE DIRECTAMENTE. O PODRIA SER QUE NO TENGA MOVIMIENTOS POSTERIORES.

'
Sub DevoluciondeBobinaaOtraOrdFab(MSF1, REG&, NORFA$)
'
' Se tranfiere del deposito de mate a deposito general calculando en consumo segun los cortes informados.
    
    Dim ctp As New Conteplast
    DEPOCORTE% = XVALO(Label13)
    CodigoProducto$ = TRIM$(MSF1.TextMatrix(REG&, 1))
    NOP% = XVALO(MSF1.TextMatrix(REG&, 6))

    CI1% = XVALO(MSF1.TextMatrix(REG&, 5))
    METROSTRANSFERIDOS = XVALO(MSF1.TextMatrix(REG&, 4))
    LOTE$ = MSF1.TextMatrix(REG&, 3)
    
    CIELEM% = CODINT%(TXTCODIGOELEM)
    CICONJ% = Me.TXTCODINT
    'ESTO ES POR EL KILOMBO QUE HICIERON QUE CAMBIARON CODIGOS DESPUES DE ESTAR GENERADAS LAS ETIQUETAS
    If CIELEM% < 1 Then
       CIELEM% = XVALO(VALOBADA("BOLRECEP", "COD_INTE", "IDENLOTE = '" & LOTE$ & "'", "NOMESS"))
       TXTCODIGOELEM = CODEXT$(CIELEM%)
    End If
    '
'    'VALIDAR SI CORRESPONDE EL CODIGO  A LA O.F.
'    If ctp.ValidarCodigosOF(CICONJ%, CIELEM%) <= 0 Then
'        Call COMUNI("Atención El Código Seleccionado No Corresponde a la Orden de Fabriación Seleccionada")
'        Exit Sub
'     End If
    
    DEPOPAÑOL% = XVALO(lbldepogeneral)
    'VALIDO SI REALMENTO NO SE INGRESO A DEPOSITO GENERAL NUEVAMENTE
    CONDI$ = "COD_INTE = " & CI1%
    CONDI$ = CONDI$ + " AND IDENLOTE = '" & LOTE$ & "'"
    CONDI$ = CONDI$ + " AND DOSECLAR = '" & TRIM(NORFA$) & " '"
    CONDI$ = CONDI$ + " AND DEPOMOVI = " & DEPOCORTE%
    CANTREG& = CantRegistros&("MOVISTO", CONDI$, "NOMESS")
    If CANTREG& Mod 2 = 0 Then
       Call COMUNI("Atención!! La Bobina Seleccionada Ya ha tenido Ingreso y Salida Para La Orden de Fabricación Activa")
       Screen.MousePointer = 1
       Exit Sub
    End If
    
'   PesoPorMetro# = PESMETRO(CI1%)
    PesoPorMetro# = CoeficientePesoPorMetroSegunLote#(LOTE$)
    If PesoPorMetro# < 0.0005 Then
        Call COMUNI("Falta Configurar Desde la Ficha del Artículo\El Peso por Metro Lineal\del Producto " & CodigoProducto$ & "")
        GoTo 99
    End If
    '

    MetroInforma# = XVALO(ctp.MetInformados(NORFA$, LOTE$))
    SALDOENMETROS# = METROSTRANSFERIDOS - MetroInforma#
    
    SaldoBobinaEnKilos# = SALDOENMETROS# * PesoPorMetro#
    
    If SaldoBobinaEnKilos# < 0.05 Then
      Call COMUNI("El Saldo Teórico es = 0 Imposible Realizar La Devolución")
      GoTo 99
      
    End If
    
    'TRANSFERENCIA
    FECHA% = HOY(HO$)
    CMOV$ = "TR"

    DEPOCONSUM% = XVALO(Me.Label13)
    DEPOALMACEN% = XVALO(Me.lbldepogeneral)

    NCOMP& = ProNroComprobante(CMOV$)
    DOPRI$ = "TR." + TRIM$(str$(NCOMP&))
    DOSEC$ = NORFA$
      
    DESCRIDEPO$ = ECOARCH$("TADEPOSI", Chr$(DEPOCONSUM%))
    REFE$ = "Tr.A " & DESCRIDEPO$ & "-" & TXTOF
    NCOMP& = ProNroComprobante(CMOV$)
    Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC%, DEPOCONSUM%, (-1) * SALDOENMETROS#, , , , , , , , , , NOP%)
    Call T_Espera(1 / 10) 'PARA QUE HAYA DIFERENCIA ENTRE UNA GRABACION Y LA OTRA
    DESCRIDEPO$ = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPOALMACEN%)))
    REFE$ = "Tr.A " & DESCRIDEPO$ & "-" & TRIM$(TXTOF)
    Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC%, DEPOALMACEN%, SALDOENMETROS#, , , , , , , , , , NOP%)
    '
    Call ctp.CargaMatTranfXOrden(Me.MSF1, NORFA$, DEPOCONSUM%)
    Set ctp = Nothing

   
99  Exit Sub
    
End Sub

Private Sub MSF1_DblClick()
    REG& = MSF1.MouseRow
    COL& = MSF1.MouseCol
    If REG& = 0 Or REG& >= MSF1.Rows Then Exit Sub
    If COL& > 0 Then Exit Sub
    IDLOT$ = TRIM$(MSF1.TextMatrix(REG&, 3))
   If IDLOT$ <> "" Then
     On Error Resume Next
     CONSUMLOT07.Label4 = TRIM$(MSF1.TextMatrix(REG&, 1))
     If Err Then Exit Sub
     On Error GoTo 0
     CONSUMLOT07.Label1 = IDLOT$
     CONSUMLOT07.Show 1
   End If
End Sub

Private Sub MSF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
      Dim ctp As New Conteplast
      REG& = MSF1.MouseRow
      SaldoEnOrigen = XVALO(MSF1.TextMatrix(REG&, 7)) 'Kg.Transf/F9.2
      If MSF1.Rows > 1 Then
        If REG& >= 1 Then
            Call SELECCIONARFILA(Me.MSF1, REG&)
            OPX% = COMALTER("Realmente Desea Retornar La Bobina Seleccionada\\Cancelar\Retornar a Depósito Original (Pesada en Balanza)")
            If OPX% = 2 Then
            
              Call ctp.DevoluciondeBobina(MSF1, REG&, TXTOF, XVALO(Label13), XVALO(lbldepogeneral), SaldoOrigen)
              Call ctp.CargaMatTranfXOrden(MSF1, TXTOF, XVALO(Label13))

'            ElseIf OPX% = 3 Then
'              Call DevoluciondeBobinaaOtraOrdFab(MSF1, REG&, TXTOF)
            End If
            Screen.MousePointer = 1
        End If
      End If
    End If
End Sub

Private Sub Text1_Change()
     Dim ctp As New Conteplast
     NUMEOPER% = ctp.NumeOperXNombre(XVALO(TXTCODINT), Text1)
     TXTNUMEOPER = NUMEOPER%
End Sub

Private Sub TXTCODINT_Change()
   CI1% = XVALO(TXTCODINT)
   Dim ctp As New Conteplast
   TXTFTECNICA = ctp.NroFTecnicaXCodigo(CI1%)

End Sub

Private Sub txtcodintelem_Change()
   CI1% = XVALO(txtcodintelem)
   TXTDESCRIELEM = DESCRIT0$(CI1%)

End Sub

Private Sub TXTLECTURA_KeyPress(KeyAscii As Integer)
'CUANDO SE LEE CON LECTORA ESCRIBE EL CODIGO Y EL LOTE
'TRAE EL PESO VER SI ESTE VA A QUEDAR SIEMPRE ACTUALIZADO ver como guardar el saldo, se va a quedar en la tabla bolrecep y el consumo va a ser en mts.
'
'TENER EN CUENTA QUE SIEMPRE VA A  SERE UN CODIGO REAL
'TAL VEZ NO HAGA FALTA MOSTRAR LAS OPERACIONES O SOLO PARA VISUALIZARLAS
'Y QUE LAS VALIDACIONES SEAN SOBRE EL TOTAL DE LOS CODIGOS Y SI NO QUE SEAN PARTE DE LOS RANGO.
'UNA VEZ QUE SE ACETPA BAJAR A LA GRILLA Y REALIZAR EL MOVIMIENTO DE STOCK COMO TRANSFERENCIA
'SI SE QUITA VERIFICAR QUE NO HAYA CONSUMOS Y EN ESE CASO REALIZAR UN MOVIMIENTO INVERSO O QUE ANULE DIRECTAMENTE. O PODRIA SER QUE NO TENGA MOVIMIENTOS POSTERIORES.

   If KeyAscii = 13 Then
     LecturaEtiqueta$ = TRIM$(TXTLECTURA)
     LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, "'", "-")
     LecturaEtiqueta$ = REPLACAR$(LecturaEtiqueta$, "]", "|")
     If LecturaEtiqueta$ = "" Then Exit Sub
     n% = InStr(1, LecturaEtiqueta$, "|")
     If n% < 1 Then Exit Sub
     Dim ctp As New Conteplast

     Dim strArray() As String
     LECTURA$ = LecturaEtiqueta$
     On Error Resume Next
     strArray = Split(LECTURA$, "|")
     TXTCODIGOELEM = strArray(0)
     LOTE$ = TRIM$(strArray(1))
     On Error GoTo 0
     '
     CIELEM% = CODINT%(TXTCODIGOELEM)
     CICONJ% = Me.TXTCODINT
     'ESTO ES POR EL KILOMBO QUE HICIERON QUE CAMBIARON CODIGOS DESPUES DE ESTAR GENERADAS LAS ETIQUETAS
     If CIELEM% < 1 Then
       CIELEM% = XVALO(VALOBADA("BOLRECEP", "COD_INTE", "IDENLOTE = '" & LOTE$ & "'", "NOMESS"))
       TXTCODIGOELEM = CODEXT$(CIELEM%)
     End If
     
     Me.txtcodintelem = CIELEM%
     DEPOPAÑOL% = XVALO(lbldepogeneral)
     If ctp.ValidarCodigosOF(CICONJ%, CIELEM%) <= 0 Then
        Call COMUNI("Atención El Código Seleccionado No Corresponde a la Orden de Fabriación Seleccionada")
        Exit Sub
     End If
     CANTMETROSBOBINA = SaldoLote(CIELEM%, LOTE$, DEPOPAÑOL%)

     Me.TXTMETROS = CANTMETROSBOBINA
'     PesoPorMetro# = PESMETRO(CIELEM%)
     PesoPorMetro# = CoeficientePesoPorMetroSegunLote#(LOTE$)
     If PesoPorMetro# < 0.0005 Then
        Call COMUNI("Falta Configurar Desde la Ficha del Artículo\El Peso por Metro Lineal\del Producto " & CodigoProducto$ & "")
        Exit Sub
     Else
        PESOSALDOBOBINA = PesoPorMetro# * CANTMETROSBOBINA
'        PESOSALDOBOBINA = PesoPorMetro# * CANTMETROSBOBINA
     End If
     '
     TXTLOTE = LOTE$
     TXTPESO = FORMATNUM$(PESOSALDOBOBINA, "F10.2")
     Text1 = OperacionesDelCodigoLeido$(CIELEM%, retornoCantReg&)
     TXTPESOORIG = FORMATNUM$(PesoOriginalSegúnLote#(LOTE$), "F10.2")
     On Error GoTo 0
   End If
End Sub

Private Sub TXTOF_Change()
    NORFA$ = TRIM$(TXTOF)
    CI1% = CODIFAB%(NORFA$)
    TXTCODINT = CI1%
    TXTCODIGO = CODEXT$(CI1%)
    TXTDESCRIPCION = DESCRIT0$(CI1%)
    Me.TXTCANAPRO = CantiAprobadaOF(NORFA$)
    Me.TXTCANPROY = CantiProyectadaOF(NORFA$)
    TXTNROPED = NroPedidoSegunOf&(NORFA$)
    TXTNROCOTIZA = NROCOTIZA(XVALO(TXTNROPED))
    TXTLECTURA.Locked = True
    MSF0.Rows = 1
    MSF1.Rows = 1
    If NORFA$ <> "" Then TXTLECTURA.Locked = False
    Dim ctp As New Conteplast
    DEPOCORTE% = XVALO(Label13)

    Call ctp.CargaMatTranfXOrden(MSF1, TRIM$(NORFA$), DEPOCORTE%)
    Set ctp = Nothing

    

 
End Sub

Private Sub TXTPESO_Change()
'   CI1% = XVALO(txtcodintelem)
'   If CI1% > 0 Then
'     PesoLeido# = XVALO(txtPeso)
'     PesoPorMetro# = PESMETRO(CI1%)
'     SaldoBobinaEnMetrosLineales# = PesoLeido# / PesoPorMetro#
'     Me.TXTMETROS = FORMATNUM$(SaldoBobinaEnMetrosLineales#, "F12.1")
'   Else
'     Me.TXTMETROS = ""
'   End If

End Sub

VERSION 5.00
Begin VB.Form DESPMAT09 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Despacho de Materiales"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8070
      TabIndex        =   9
      Top             =   6735
      Width           =   1305
   End
   Begin VB.TextBox Text12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   1260
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2595
      Width           =   8100
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1275
      TabIndex        =   5
      Top             =   2160
      Width           =   8085
   End
   Begin VB.TextBox Text10 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8070
      TabIndex        =   8
      Top             =   6390
      Width           =   1305
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1500
      TabIndex        =   7
      Top             =   6390
      Width           =   5295
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
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
      Left            =   225
      TabIndex        =   54
      Top             =   3450
      Width           =   9165
      Begin VB.Frame Frame8 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   2160
         TabIndex        =   57
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame7 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   5640
         TabIndex        =   56
         Top             =   360
         Width           =   15
      End
      Begin VB.Frame Frame5 
         Caption         =   "Frame3"
         Height          =   2200
         Left            =   7080
         TabIndex        =   55
         Top             =   360
         Width           =   15
      End
      Begin VB.ListBox List1 
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
         Height          =   2040
         ItemData        =   "DESPMAT09.frx":0000
         Left            =   420
         List            =   "DESPMAT09.frx":0002
         MousePointer    =   1  'Arrow
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   58
         Top             =   560
         Width           =   8355
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"DESPMAT09.frx":0004
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   420
         TabIndex        =   59
         Top             =   330
         Width           =   8340
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   4740
      TabIndex        =   45
      Top             =   165
      Width           =   4620
      Begin VB.CommandButton Command13 
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
         Left            =   1920
         TabIndex        =   2
         Top             =   345
         Width           =   175
      End
      Begin VB.CommandButton Command7 
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
         Left            =   1920
         TabIndex        =   4
         ToolTipText     =   "Seleccione aquí si recibe de una DROGUERIA"
         Top             =   1200
         Width           =   175
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
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         ToolTipText     =   "Seleccione aquí si recibe de un LABORATORIO"
         Top             =   765
         Width           =   175
      End
      Begin VB.Label Label34 
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
         Left            =   1260
         TabIndex        =   71
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label Label33 
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
         Left            =   1260
         TabIndex        =   70
         Top             =   765
         Width           =   660
      End
      Begin VB.Label Label32 
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
         Left            =   1260
         TabIndex        =   69
         Top             =   345
         Width           =   660
      End
      Begin VB.Label Label20 
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
         Left            =   2220
         TabIndex        =   61
         Top             =   345
         Width           =   2205
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   345
         TabIndex        =   60
         Top             =   420
         Width           =   885
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Laboratorio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   49
         Top             =   855
         Width           =   1080
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   285
         TabIndex        =   48
         Top             =   1305
         Width           =   945
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
         Left            =   2220
         TabIndex        =   47
         Top             =   1215
         Width           =   2220
      End
      Begin VB.Label Label3 
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
         Left            =   2220
         TabIndex        =   46
         Top             =   765
         Width           =   2205
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Origen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   255
      TabIndex        =   44
      Top             =   690
      Width           =   4275
      Begin VB.CommandButton Command14 
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
         Left            =   1650
         TabIndex        =   0
         Top             =   300
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
         Height          =   285
         Left            =   3915
         TabIndex        =   1
         Top             =   300
         Width           =   175
      End
      Begin VB.Label Label37 
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
         Left            =   990
         TabIndex        =   72
         Top             =   300
         Width           =   660
      End
      Begin VB.Label Label26 
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
         Left            =   975
         TabIndex        =   64
         Top             =   690
         Width           =   3120
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   45
         TabIndex        =   63
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label13 
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
         Left            =   2910
         TabIndex        =   53
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1350
         TabIndex        =   52
         Top             =   345
         Width           =   1545
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Seleccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1605
      Left            =   240
      TabIndex        =   36
      Top             =   7065
      Width           =   9180
      Begin VB.TextBox Text3 
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
         Left            =   4800
         TabIndex        =   10
         Top             =   315
         Width           =   3495
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Del"
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
         Left            =   8415
         TabIndex        =   17
         Top             =   1215
         Width           =   600
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Ins"
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
         Left            =   8415
         TabIndex        =   16
         Top             =   975
         Width           =   600
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Add"
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
         Left            =   8415
         TabIndex        =   15
         Top             =   735
         Width           =   600
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
         Left            =   3045
         TabIndex        =   12
         Top             =   345
         Width           =   175
      End
      Begin VB.TextBox Text7 
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
         Left            =   900
         TabIndex        =   11
         Top             =   345
         Width           =   2160
      End
      Begin VB.CommandButton Command8 
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
         Left            =   2400
         TabIndex        =   14
         Top             =   1125
         Width           =   175
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
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
         Left            =   6840
         TabIndex        =   13
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lectura >>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3240
         TabIndex        =   74
         Top             =   345
         Width           =   1545
      End
      Begin VB.Label Label11 
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
         Left            =   4800
         TabIndex        =   37
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -720
         TabIndex        =   51
         Top             =   375
         Width           =   1545
      End
      Begin VB.Label Label30 
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
         Left            =   900
         TabIndex        =   50
         Top             =   690
         Width           =   7390
      End
      Begin VB.Label Label36 
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
         Left            =   900
         TabIndex        =   43
         Top             =   1125
         Width           =   1575
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Id.Bulto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -510
         TabIndex        =   42
         Top             =   1185
         Width           =   1335
      End
      Begin VB.Label Label24 
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
         Left            =   3120
         TabIndex        =   40
         Top             =   1140
         Width           =   1140
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ubic:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1950
         TabIndex        =   39
         Top             =   1215
         Width           =   1140
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cant."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3840
         TabIndex        =   38
         Top             =   1230
         Width           =   930
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cons."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5800
         TabIndex        =   41
         Top             =   1230
         Width           =   1005
      End
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   8820
      Left            =   9585
      ScaleHeight     =   8760
      ScaleWidth      =   1320
      TabIndex        =   22
      Top             =   -20
      Width           =   1380
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
         Height          =   645
         Left            =   105
         Picture         =   "DESPMAT09.frx":009B
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Anulación  Orden de Despacho"
         Top             =   2640
         Width           =   690
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Kdx"
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
         Left            =   105
         Picture         =   "DESPMAT09.frx":03A5
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Ficha Tipo Kardex de Movimientos Históricos"
         Top             =   1815
         Width           =   690
      End
      Begin VB.CommandButton Command15 
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
         Height          =   690
         Left            =   105
         Picture         =   "DESPMAT09.frx":06AF
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1005
         Width           =   690
      End
      Begin VB.CommandButton Command9 
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
         Left            =   3255
         TabIndex        =   35
         Top             =   0
         Width           =   175
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
         Height          =   285
         Left            =   1680
         TabIndex        =   32
         Top             =   1050
         Width           =   6675
      End
      Begin VB.TextBox Text2 
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
         Left            =   7035
         TabIndex        =   31
         Top             =   525
         Width           =   1320
      End
      Begin VB.CommandButton Command2 
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
         Height          =   690
         Left            =   105
         Picture         =   "DESPMAT09.frx":0AF1
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Graba - Aprueba la Transacción"
         Top             =   6960
         Width           =   690
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   630
         TabIndex        =   23
         Top             =   7845
         Width           =   690
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   24
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
         Height          =   690
         Left            =   105
         Picture         =   "DESPMAT09.frx":0DFB
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Salir"
         Top             =   195
         Width           =   690
      End
      Begin VB.Label Label10 
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
         Left            =   1680
         TabIndex        =   34
         Top             =   0
         Width           =   1650
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   33
         Top             =   45
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia:"
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
         Left            =   0
         TabIndex        =   30
         Top             =   1095
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Consumo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5355
         TabIndex        =   29
         Top             =   570
         Width           =   1545
      End
      Begin VB.Label Label18 
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
         Left            =   1680
         TabIndex        =   28
         Top             =   525
         Width           =   1650
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ubicacion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   210
         TabIndex        =   27
         Top             =   570
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2730
         TabIndex        =   26
         Top             =   570
         Width           =   1545
      End
      Begin VB.Label Label2 
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
         Left            =   4515
         TabIndex        =   25
         Top             =   525
         Width           =   1335
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "DESPMAT09.frx":0F45
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   8085
         Width           =   1515
      End
   End
   Begin VB.Label Label38 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cant. Palet:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7020
      TabIndex        =   73
      Top             =   6810
      Width           =   1350
   End
   Begin VB.Label Label31 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Despacho de Materiales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   285
      TabIndex        =   68
      Top             =   180
      Width           =   3945
   End
   Begin VB.Label Label29 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Producto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   330
      TabIndex        =   67
      Top             =   2610
      Width           =   1410
   End
   Begin VB.Label Label28 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Concepto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   270
      TabIndex        =   66
      Top             =   2220
      Width           =   1410
   End
   Begin VB.Label Label27 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Bultos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7425
      TabIndex        =   65
      Top             =   6465
      Width           =   855
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Propiedad de: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   255
      TabIndex        =   62
      Top             =   6465
      Width           =   1410
   End
End
Attribute VB_Name = "DESPMAT09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private LOTE$
Private Codigo$

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command10_Click()
'     If TRIM$(Text7) <> "" Then
'     If TRIM$(Label36) <> "" Then
   CIXI% = CODINT%(Text7)
   If CIXI% < 1 Then
        COMUNI ("Código No Válido")
        Text3 = ""
        Text7 = ""
        Label36 = ""
        Exit Sub
   End If
   LOTE$ = TRIM$(Label36)
   If CantRegistros&("LOTINGRE", "COD_INTE = " & CIXI% & " AND IDENLOTE = '" & LOTE$ & "'") < 1 Then
        COMUNI ("Lote Ingresado No Válido")
        Text3 = ""
        Text7 = ""
        Label36 = ""
        Exit Sub
   End If
      'VALIDA CANTIDAD
      If XVALO(Text5) < FRACMIN(CIXI%) Then
         COMUNI ("Cantidad a Consumir no Válida")
         Exit Sub
      End If
      'VALIDA SI SUPERA A LOTE
      If XVALO(Text5) > XVALO(Label11) Then
         COMUNI ("Cantidad Consumida es Mayor a la Disponible en Lote")
         Exit Sub
      End If
      '
      CONDI$ = "IDENLOTE = '" & Label36 & "' AND COD_INTE = " & CODINT%(Text7)
      DEPLOT% = XVALO(VALOBADA("LOTINGRE", "DEPOUBIC", CONDI$, "NOMESS"))
      If DEPLOT% <> XVALO(Label37) Then
        Call MENSERR(24, "El Lote Seleccionado no se Encuentra en el Deposito de Origen Seleccionado")
        GoTo 99
      End If
      '

      'VALIDA SI YA SE SELECCIONO ENTONCES:
      'SE PREGUNTA SI NO SUPERA A LOTE
      'Y SE SUMA LA CANTIDAD
      For I = 0 To List1.ListCount
        COD$ = UCase$(TRIM(Mid(List1.List(I), 1, 16)))
        LOTE$ = UCase$(TRIM(Mid(List1.List(I), 51, 16)))
        If COD$ = UCase$(TRIM(Text7)) And LOTE$ = UCase$(TRIM$(Label36)) Then
           XX$ = List1.List(I)
           CANTACT# = XVALO(Mid(List1.List(I), 67, 12))
           CantCons# = XVALO(Text5)
           CANTOT# = CANTACT# + CantCons#
           If CANTOT# > XVALO(Label11) Then
              COMUNI ("Cantidad Consumida es Mayor a la Disponible en Lote")
              Call LIMPCONMAT
              Text7 = ""
              Text3 = ""
              Exit Sub
           End If
           CANTOTAL$ = FORMATNUM(CANTOT#, "F12." & CNTDC)
           Call REPLA(XX$, CANTOTAL$, 67, 12)
           List1.List(I) = XX$
           'Call COMUNI("Ya se utilizo el Lote Seleccionado")
           GoTo 99
        End If
      Next I
      '
      'SI NO SE SELECCIONO ENTONCES SE AGREGA
      LINEA$ = ""
      '
      Call REPLA(LINEA$, Text7, 1, 16)    'CODIGO
      Call REPLA(LINEA$, Label30, 18, 32) 'DESCR
      Call REPLA(LINEA$, Label36, 51, 16) 'LOTE
      CONS# = XVALO(Text5)
      Call REPLA(LINEA$, FORMATNUM(CONS#, "F12." & CNTDC), 67, 12) 'CANTIDAD
      
      List1.AddItem LINEA$
'    End If
'  End If
  '
99 Call LIMPCONMAT
   Text7 = ""
   Text3 = ""
  '
End Sub

Private Sub Command11_Click()
  pos% = List1.ListIndex
  '
'     If TRIM$(Text7) <> "" Then
'     If TRIM$(Label36) <> "" Then
   CIXI% = CODINT%(Text7)
   If CIXI% < 1 Then
        COMUNI ("Código No Válido")
        Text3 = ""
        Text7 = ""
        Label36 = ""
        Exit Sub
    End If
    '
    LOTE$ = TRIM$(Label36)
    If CantRegistros&("LOTINGRE", "COD_INTE = " & CIXI% & " AND IDENLOTE = '" & LOTE$ & "'") < 1 Then
        COMUNI ("Lote Ingresado No Válido")
        Text3 = ""
        Text7 = ""
        Label36 = ""
        Exit Sub
    End If
    '
    'VALIDA CANTIDAD
    If XVALO(Text5) < FRACMIN(CIXI%) Then
       COMUNI ("Cantidad a Consumir no Válida")
       Exit Sub
    End If
    'VALIDA SI SUPERA A LOTE
    If XVALO(Text5) > XVALO(Label11) Then
       COMUNI ("Cantidad Consumida es Mayor a la Disponible en Lote")
       Exit Sub
    End If
    '
    CONDI$ = "IDENLOTE = '" & Label36 & "' AND COD_INTE = " & CODINT%(Text7)
    DEPLOT% = XVALO(VALOBADA("LOTINGRE", "DEPOUBIC", CONDI$, "NOMESS"))
    If DEPLOT% <> XVALO(Label37) Then
      Call MENSERR(24, "El Lote Seleccionado no se Encuentra en el Deposito de Origen Seleccionado")
      GoTo 99
    End If
    '
    
    'VALIDA SI YA SE SELECCIONO ENTONCES:
    'SE PREGUNTA SI NO SUPERA A LOTE
    'Y SE SUMA LA CANTIDAD
    For I = 0 To List1.ListCount
      COD$ = UCase$(TRIM(Mid(List1.List(I), 1, 16)))
      LOTE$ = UCase$(TRIM(Mid(List1.List(I), 51, 16)))
      If COD$ = UCase$(TRIM(Text7)) And LOTE$ = UCase$(TRIM$(Label36)) Then
         XX$ = List1.List(I)
         CANTACT# = XVALO(Mid(List1.List(I), 66, 14))
         CantCons# = XVALO(Text5)
         CANTOT# = CANTACT# + CantCons#
         If CANTOT# > XVALO(Label11) Then
            COMUNI ("Cantidad Consumida es Mayor a la Disponible en Lote")
            Exit Sub
         End If
         CANTOTAL$ = FORMATNUM(CANTOT#, "F12." & CNTDC)
         Call REPLA(XX$, CANTOTAL$, 66, 14)
         List1.List(I) = XX$
         'Call COMUNI("Ya se utilizo el Lote Seleccionado")
         GoTo 99
      End If
    Next I
    '
    'SI NO SE SELECCIONO ENTONCES SE AGREGA
    LINEA$ = ""
    '
    Call REPLA(LINEA$, Text7, 1, 16)    'CODIGO
    Call REPLA(LINEA$, Label30, 18, 32) 'DESCR
    Call REPLA(LINEA$, Label36, 51, 16) 'LOTE
    CONS# = XVALO(Text5)
    Call REPLA(LINEA$, FORMATNUM(CONS#, "F12." & CNTDC), 67, 12) 'CANTIDAD
    
    List1.AddItem LINEA$, pos%
'    End If
'  End If
  '
99 Call LIMPCONMAT
   Text7 = ""
   Text3 = ""
  '
End Sub

Private Sub COMMAND12_Click()
   If List1.List(ListIndex) <> "" Then
      List1.RemoveItem (List1.ListIndex)
   End If
End Sub

Private Sub Command13_Click()
  Command13.Enabled = False
  '
  Call SELECHO("TADEPOSI")
  NPROV$ = VALACT1$("TADEPOSI")
  If NPROV$ <> "" Then
     Label32 = NPROV$
  End If
  '
  Command13.Enabled = True
End Sub

Private Sub Command14_Click()
  Command14.Enabled = False
  '
  Call SELECHO("TADEPOSI")
  NPROV$ = VALACT1$("TADEPOSI")
  If NPROV$ <> "" Then
     Label37 = NPROV$
  End If
  '
  Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
   Command14.Enabled = False
   CFGRECEP.Show 1
   Command14.Enabled = True
End Sub

Private Sub Command16_Click()
 FIKARDEX09.Show 1
End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    '
    DEPOTRANS% = XVALO(CONTROL$("BOLRECEP", "DEPTRANS")) ' DEPOSITO MATERIAL EN TRANSITO
    '
    If XVALO(Label37) = DEPOTRANS% Then
        COMUNI ("Imposible Transferir desde Deposito de Material en Transito")
       GoTo 99
    End If
    '
    If XVALO(Label32) = DEPOTRANS% Then
        COMUNI ("Imposible Transferir Hacia Deposito de Material en Transito")
       GoTo 99
    End If
    '
    If XVALO(Label37) < 1 Then
       COMUNI ("Falta Ingresar Deposito Origen")
       GoTo 99
    End If
    '
    If XVALO(Label33) < 1 And XVALO(Label34) < 1 And XVALO(Label32) < 1 Then
       Call COMUNI("Debe Seleccionar un Destino")
       GoTo 99
    End If
    '
    If TRIM(Text11) = "" Then
       COMUNI ("Falta Ingresar 'Concepto'")
       GoTo 99
    End If
    '
    If TRIM(Text8) = "" Then
       COMUNI ("Falta Ingresar 'Propiedad'")
       Text8.SetFocus
       GoTo 99
    End If
    '
    If XVALO(Text10) < 1 Then
       COMUNI ("Falta Ingresar 'Cant. Bultos'")
       Text10.SetFocus
       GoTo 99
    End If
    '
    If XVALO(Text1) < 1 Then
       COMUNI ("Falta Ingresar 'Cant. Palet'")
       GoTo 99
    End If
    '
    If List1.ListCount < 1 Then
       COMUNI ("Falta Ingresar Algun Producto")
       GoTo 99
    End If
    '
    DEPDESTI% = XVALO(Label32)                 ' Deposito DESTINO (central de pesadas)
    DEPORIGI% = XVALO(Label37)
    If DEPDESTI% = DEPORIGI% Then
        Call MENSERR(24, "Imposible Transferir.\Coinciden Deposito ORIGEN y DESTINO.")
        GoTo 99
    End If
    '
    If DEPDESTI% > 0 Then
        REGI$ = FILTERGETRECORD("TADEPOSI", 1, Chr$(DEPDESTI%))
        PROVE% = CVI(Mid(REGI$, 61, 2))                   ' Nro Proveedor
        If PROVE% > 0 Then
          If TRIM$(RASOCLI$(-1 * PROVE%)) = "" Then
            Call MENSERR(24, "Error de Configuración.\Proveedor no Válido para Depósito " & DEPDESTINO%)
            GoTo 99
          End If
          If ECOARCH$("TADEPOSI", Chr$(DEPOTRANS%)) = "" Then
            Call MENSERR(24, "Debe Definir Depósito de Material en Transito")
            GoTo 99
          End If
        End If
        'VALIDACION DE PLANTA
        RegOrig$ = FILTERGETRECORD("TADEPOSI", 1, Chr$(DEPORIGI%))
        PlantOrig$ = TRIM$(Mid(RegOrig$, 57, 4)) 'PLANTA ASIGNADA A DEPOSITO ORIGEN
        If ECOARCH$("PLANTPRO", PlantOrig$) = "" Then
            Call MENSERR(24, "Debe Definir Planta para Depósito de Origen")
            GoTo 99
        End If
        PlantDest$ = TRIM$(Mid(REGI$, 57, 4)) 'PLANTA ASIGNADA A DEPOSITO DESTINO
        If ECOARCH$("PLANTPRO", PlantDest$) = "" Then
            Call MENSERR(24, "Debe Definir Planta para Depósito de Destino")
        GoTo 99
        End If

        
    End If
    '
    'Obtener Datos
    NumeroOD$ = ProNroComp("OD")
    NumODCor$ = "OD." & NumeroOD$
    Do While Len(NumeroOD$) < 6
       NumeroOD$ = "0" & NumeroOD$
    Loop
    NumODLar$ = "OD-" & NumeroOD$ & "-1"
    FECHA% = FECHANUM(Label13)
    Destino$ = TRIM$(Label20) + TRIM(Label2) + TRIM(Label3)
    DESCRI$ = Label13 + "  -  " + Destino$ + "  -  " + Text11 'Descripcion
    '
    'Grabar en EncaRemi
    Dim ENCA As ADODB.Recordset
    Set ENCA = New ADODB.Recordset
    ENCA.Open "SELECT * FROM ENCAREMI", FLEXCONN, adOpenKeyset, adLockPessimistic
    With ENCA
      .AddNew
      !DoSecLar = NumODLar$
      !DOPRILAR = ""
      !Descripcion = Left$(DESCRI$, 64)
      !fechMov = FETEXCOR$(FECHA%)
      !Propiedad = Left$(SAFETEXT(Text8), 64)
      !Concepto = Left$(SAFETEXT(Text11), 64)
      !Producto = Left$(SAFETEXT(Text12), 64)
      !Bultos = XVALO(Text10)
      !Origen = XVALO(Label37)
      !Deposito = XVALO(Label32)
      !Nume_Clie = XVALO(Label33)
      !Nume_Prov = XVALO(Label34)
      !CantPalet = XVALO(Text1)
      .Update
      .Close
    End With
    
    'Se Graba en Movisto - Detalle
    For I = 0 To List1.ListCount - 1
      XX$ = List1.List(I)
      CI% = CODINT%(Mid$(XX$, 1, 16))
      LOTE$ = TRIM$(Mid$(XX$, 51, 16))
      CAN# = XVALO(Mid$(XX$, 67, 12))  'CANTIDAD
      DEPO% = XVALO(Label37)
      Call MOVISTOK(FECHA%, CI%, LOTE$, "OD", "", "", NumODLar$, NumODCor$, Text11, 0, 0, 0, 0, DEPO%, (-1) * CAN#)

      ' EN CASO QUE EL DESTINO SEA UN DEPOSITO DE BLIPACK
      If DEPDESTI% > 0 Then
        '
        REFEREN$ = "Despacho de Material"
        DEPDESTIXX% = DEPDESTI%
'        If PROVE% > 0 Then           'PROVE% = proveedor asignado al deposito de destino
        If PlantOrig$ <> PlantDest$ Then ' SI LAS PLANTAS ASIGNADAS DIFIEREN PASA A DEPOSITO EN TRANSITO
           DEPDESTIXX% = DEPOTRANS%  'Si el deposito tiene proveedor asignado, el deposito de destino pasa a ser el deposito de transito
           REFEREN$ = "Material en Transito"
        End If
        ' SE MOVIO FUERA DEL IF PROVE% PARA QUE SI GRABE LA TRANSFERENCIA DE TODOS MODOS
        Call MOVISTOK(FECHA%, CI%, LOTE$, "OD", "", "", NumODLar$, NumODCor$, REFEREN$, 0, 0, 0, 0, DEPDESTIXX%, CAN#)
        'Actualiza Deposito en Lotingre
        CONDI$ = "IDENLOTE='" & LOTE$ & "' and Cod_inte =" & CI%
        Call ACTUREGISTRO("LOTINGRE", "DEPOUBIC", CONDI$, DEPDESTIXX%, REG&)
        UBI$ = ""
        'SI ES MATERIAL EN TRANSITO ASIGNA A LA UBICACION 3 ASTERISCOS
        If DEPDESTIXX% = DEPOTRANS% Then UBI$ = "***"
        Call ACTUREGISTRO("LOTINGRE", "UBICFIS", CONDI$, UBI$, REG&)

      End If
      '
    Next I
    Call COMUNI("Orden de Despacho Procesada")
    Call limpiar
    '
    Call ENVIMAILODES(NumODLar$)
    '
99  Screen.MousePointer = 1
    Command2.Enabled = True
    '
End Sub

Private Sub Command3_Click()
    '
    Command3.Enabled = False
    '
    If DERACCE%("ANUODES", "Anulacion de Orden de Despacho") < 2 Then
       GoTo 99
    End If
    '
5   Call CARSELMESA
    '
10  Screen.MousePointer = 1
    Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
       GoTo 99
    End If
    CORRESQ$ = MESANO$(MESASEL$, DES%, HAS%)
    '
    Call FRESHECHO("!SELREMI")
    Screen.MousePointer = 11
    '
    DES1$ = FETEXCOR1$(DES%)
    HAS1$ = FETEXCOR1$(HAS%)
    '
    SQLX$ = "FechMov >= '" & DES1$ & "' "
    SQLX$ = SQLX$ + " AND FechMov <= '" & HAS1$ & "' "
    SQLX$ = SQLX$ + " AND CodiMovi =  'OD'  "
    '
    Call CARGALISEL("DESPMES", "MOVISTO", "DoSecLar/A16;DoPriCor/A16;FECHMOV/D8", SQLX$, "DISTINCT")
    '
    AA$ = "Despacho de Materiales del " & FECHATEX(DES%) & " al " & FECHATEX(HAS%)
    Call SELECHO("DESPMES/" & AA$)
    OD$ = TRIM$(VALACT1$("DESPMES"))
    If OD$ = "" Then GoTo 5
    '
    CONDI$ = SQLX$ & " AND DOSECLAR ='" & OD$ & "'"
    Call BORRAREGISTROS("MOVISTO", CONDI$, REG&)
    Call COMUNI("Orden de Despacho Eliminada Correctamente")
    GoTo 5
    '
99  Screen.MousePointer = 1
    
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
  Command4.Enabled = False
  '
  Call SELECHO("DEUDOR1")
  NPROV$ = VALACT1$("DEUDOR1")
    
  If NPROV$ <> "" Then
     Label33 = NPROV$
  End If
  '
  Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    '
    Call SELECHO("MASTER")
    If VALACT1$("MASTER") <> "" Then
       Text7 = VALACT1$("MASTER")
       A = CODINT(Text7)
    End If
    '
End Sub

Private Sub Command6_Click()
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
      Label13 = VDEV$
   End If
End Sub

Private Sub Command7_Click()
  Command5.Enabled = False
  '
  Call SELECHO("PROVED1")
  NPROV$ = VALACT1$("PROVED1")
    
  If NPROV$ <> "" Then
     Label34 = NPROV$
  End If
  '
  Command5.Enabled = True
End Sub

Private Sub Command8_Click()
    '
    Call SELECHO("!LILOTSE")
    If VALACT1$("!LILOTSE") <> "" Then
       Label36.Caption = VALACT1$("!LILOTSE")
    End If
    Text5.SetFocus
    '
End Sub

Private Sub Form_Load()
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)

   Label13 = FECHATEX(HOY(HO$))
End Sub


Private Sub Label36_Change()
  '
   CIXXI% = CODINT(TRIM(Text7))
   Label24 = UBILOTE$(CIXXI%, Label36, 1)
   Label11 = TRIM$(FORMATNUM(SALNOCONS(Label36, CIXXI%), "F13." & CNTDC$(CIXXI%)))
   Text5 = Label11
   'Si no trae informacion la lectura de la etiqueta
   'se muestra el total del lote
'   If XVALO(Text5) < 1 Then
'      Text5 = Label11
'   End If
   '
End Sub

Private Sub Label33_Change()
   If XVALO(Label33) > 0 Then
      NC% = XVALO(Label33)
      Label3 = RASOCLI(NC%)
      Label32 = ""
      Label34 = ""
     Else
      Label3 = ""
   End If
End Sub

Private Sub Label37_DblClick()
   Call Command14_Click
End Sub


Private Sub Text1_LostFocus()
 Text1 = TRIM(FORMATNUM$(XVALO(Text1), "F13.0"))
End Sub

Private Sub Text10_LostFocus()
   Text10 = TRIM(FORMATNUM$(XVALO(Text10), "F13.0"))
End Sub

Private Sub label34_Change()
   '
   If XVALO(Label34) > 0 Then
      CPROV% = XVALO(Label34)
      Label4 = RASOCLI(-CPROV%)
      Label32 = ""
      Label33 = ""
     Else
      Label4 = ""
   End If
End Sub

Private Sub Text3_Change()
' 'Para Lectura de Codigo de Barras
'   If Len(Text3) > 16 Then
'      Codigo$ = UCase(TRIM$(Mid$(Text3, 1, 16)))
'      If CODINT(Codigo$) > 0 Then
'         LOTE$ = Mid$(Text3, 17, 12)
'         Text7 = Codigo$
'         Label36.Caption = TRIM$(LOTE$)
'         CANTI# = XVALO(Mid$(Text3, 29, 8))
'         If CANTI# > 0 Then
'            Text5 = FORMATNUM$(CANTI#, "F12." & CNTDC)
'         End If
''         Call Command10_Click
''         Text3 = ""
'      Else
'          COMUNI ("Código Ingresado No Válido")
'          Text3 = ""
'          Text7 = ""
'          Label36 = ""
'      End If
'   End If
'   '
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    '
    If KeyAscii <> 13 Or TRIM$(Text3) = "" Then
      Exit Sub
    End If
    '
    Call FRACLECBAR(Text3, Text7, Label36, Label11)
    '
    CIXI% = CODINT(Text7) 'CODINT%(TRIM$(Mid$(TENTRA$, 1, 16)))
    If CIXI% < 1 Then
        COMUNI ("Código Ingresado No Válido")
        Text3 = ""
        Text7 = ""
        Label36 = ""
        Exit Sub
    End If
    LOTE$ = TRIM(Label36) 'TRIM$(Mid$(TENTRA$, 17, 12))
    '
    If CantRegistros&("LOTINGRE", "COD_INTE = " & CIXI% & " AND IDENLOTE = '" & LOTE$ & "'") < 1 Then
        COMUNI ("Lote Ingresado No Válido")
        Text3 = ""
        Text7 = ""
        Label36 = ""
        Exit Sub
    End If
    '
    CANTI# = XVALO(Label11) 'XVALO(Mid$(Text3, 29, 8))
    CONDI$ = "COD_INTE = " & CIXI% & " AND IDENLOTE = '" & LOTE$ & "'"
    If CANTI# = 0 Then
        CANTI# = XVALO(VALOBADA("LOTINGRE", "(CanAlta - CanConsu)", CONDI$, "NOMESS"))
    End If
    '
    If CIXI% > 0 Then
        Text7 = CODEXT$(CIXI%)
        If CantRegistros&("LOTINGRE", CONDI$) > 0 Then
           Label36 = LOTE$
           Text5 = FORMATNUM$(CANTI#, "F13." & CNTDC)
           Call Command10_Click
        End If
    End If
    '
End Sub

Private Sub Text5_LostFocus()
    Text5 = FORMATNUM$(XVALO(Text5), "F13." & CNTDC$)
End Sub

Private Sub Label32_Change()
   
   If XVALO(Label32) > 0 Then
      Label20 = ECOARCH("TADEPOSI", Chr(TRIM(Label32)))
      Label33 = ""
      Label34 = ""
     Else
      Label20 = ""
   End If
End Sub

Private Sub Label32_DblClick()
   Call Command13_Click
End Sub

Private Sub Text7_Change()
   '
   If Label36 <> "" Then
      Call LIMPCONMAT
   End If
   
   CI% = CODINT%(Text7)
   Call ARMALILOT(CI%, XVALO(Label37), , , 1)
   '
   CONDI$ = "COD_INTE = " & Str(CI%)
   'Label30 = VALOBADA("MASTER", "Descripcion", "Codint = " & Str(CI%))
   'SE AGREGA PARA QUE LA DESCRIPCION SEA VACIA EN CASO DE QUE EL CODIGO INTERNA NO EXISTA
   Label30 = DESCRIT0$(CI%)
   If UCase(Label30) = "MATERIAL NO CODIFICADO" Then Label30 = ""
   '
End Sub

Private Sub Text7_DblClick()
   Call Command5_Click
End Sub

Sub LIMPCONMAT()
    '
    Label36 = ""
    Label24 = ""
    Label11 = ""
    Label30 = ""
    Text5 = ""
    '
End Sub

Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 113 Then
       Call Command5_Click
    End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Command8.SetFocus
   End If
End Sub

Private Sub Text7_LostFocus()
    If Text7 <> "" Then
       CI% = CODINT(Text7)
       Text7 = CODEXT(CI%)
    End If
End Sub

Private Sub Label37_Change()
   If XVALO(Label37) < 1 Then
      Label26 = ""
      Exit Sub
   End If
   '
   If XVALO(Label37) > 0 Then
      Label26 = ECOARCH("TADEPOSI", Chr(TRIM(Label37)))
      Text7 = ""
      Label30 = ""
      Call LIMPCONMAT
   End If
End Sub
'
Function ProNroComp(TM$) 'Tipo de movimiento (CodiMovi) en movisto

    Call ABRESTOCKS
    '
    SQLX$ = "SELECT max(cast(substring(DoSecCor, 4,6)as int))as Maximo FROM MOVISTO"
    SQLX$ = SQLX$ + " WHERE CodiMovi = '" & TM$ & "'"
    
    UNUM& = 0
    Dim MIRS As ADODB.Recordset
    On Error Resume Next
    Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
    If Not IsNull(MIRS!maximo) Then
        UNUM& = MIRS!maximo
    End If
    MIRS.Close
    On Error GoTo 0
    ProNroComp = 1 + UNUM&
    Set MIRS = Nothing
End Function
Sub limpiar()
    '
    Label20 = ""
    Label32 = ""
    Label33 = ""
    Label3 = ""
    Label34 = ""
    Label4 = ""
    Label37 = ""
    Label26 = ""
    Label30 = ""
    Label24 = ""
    Label36 = ""
    Label11 = ""
    List1.Clear
    Text11 = ""
    Text10 = ""
    Text8 = ""
    Text12 = ""
    Text7 = ""
    Text5 = ""
    Text1 = ""
    '
End Sub
Sub AsignaValor()
       Text7 = Codigo$
       Label36.Caption = TRIM$(LOTE$)
       Call Command10_Click
       
'      Text7 = ""
'      Label30 = ""
'      Text5 = ""
End Sub

Sub ENVIMAILODES(NumODLar$)
   '
   MAILADMIN$ = CONTROL("BOLRECEP", "REMIMAIL")
   '
   ASUNTO$ = "Orden de Despacho " & NumODLar$
   '
   CONDI$ = "DoSecLar = '" & NumODLar$ & "'"
   CONCEP$ = VALOBADA("ENCAREMI", "CONCEPTO", CONDI$, "NOMESS")
   PRODUC$ = VALOBADA("ENCAREMI", "PRODUCTO", CONDI$, "NOMESS")
   DEPO% = XVALO(VALOBADA("ENCAREMI", "Deposito", CONDI$, "NOMESS"))
   NUCL% = XVALO(VALOBADA("ENCAREMI", "Nume_Clie", CONDI$, "NOMESS"))
   NUPR% = XVALO(VALOBADA("ENCAREMI", "Nume_Prov", CONDI$, "NOMESS"))
   If DEPO% > 0 Then DETIN$ = ECOARCH$("TADEPOSI", Chr$(DEPO%))
   If NUCL% > 0 Then DETIN$ = RASOCLI(NUCL%)
   If NUPR% > 0 Then DETIN$ = RASOCLI(NUPR% * -1)
   TEXTOMAI$ = "Se ha generado orden de despacho " + NumODLar$ & "para " & DETIN$ & CRLF$
   TEXTOMAI$ = TEXTOMAI$ & "Concepto: " & CONCEP$ + CRLF$
   TEXTOMAI$ = TEXTOMAI$ & "Producto: " & PRODUC$ & CRLF$
   '
   MAI_DESTI$ = TRIM$(MAILADMIN$)
   MAI_DIREL$ = MAI_DESTI$
   MAI_ASUNT$ = ASUNTO$
   MAI_TEXTO$ = TEXTOMAI$
   '
   Call GENERAMAIL
   '
End Sub

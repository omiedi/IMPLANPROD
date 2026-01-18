VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMPRESUA14 
   BackColor       =   &H00FFFFFF&
   Caption         =   "COTIZACIÓN AHP"
   ClientHeight    =   9420
   ClientLeft      =   0
   ClientTop       =   -555
   ClientWidth     =   12855
   ForeColor       =   &H00000000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   12855
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton CARGA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Sel.Man."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   10320
      TabIndex        =   58
      ToolTipText     =   " Modo de Carga: Selección Manual de Lotes"
      Top             =   2600
      Width           =   1095
   End
   Begin VB.OptionButton CARGA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "V.Stock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   9180
      TabIndex        =   60
      ToolTipText     =   "Modo de Carga: Valida Stock de Maera General"
      Top             =   2600
      Width           =   1095
   End
   Begin VB.OptionButton CARGA 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Rápida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   8040
      TabIndex        =   61
      ToolTipText     =   "Modo de Carga: Directamente sin Validar Stock "
      Top             =   2600
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Actualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   11460
      TabIndex        =   62
      Top             =   2570
      Width           =   1365
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Graba e Imprime"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   63
      Top             =   8880
      Width           =   1995
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Height          =   1575
      Left            =   15
      TabIndex        =   44
      Top             =   0
      Width           =   12840
      Begin VB.Frame Frame3 
         BackColor       =   &H00404040&
         Height          =   1120
         Left            =   120
         TabIndex        =   55
         Top             =   210
         Width           =   1400
         Begin VB.CommandButton Command19 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1100
            TabIndex        =   57
            ToolTipText     =   "Seleccion de Cliente"
            Top             =   480
            Width           =   250
         End
         Begin VB.TextBox Text1 
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
            Height          =   300
            Left            =   120
            TabIndex        =   56
            Top             =   460
            Width           =   975
         End
         Begin VB.Label Label26 
            BackColor       =   &H00EEEEEE&
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
            Height          =   300
            Left            =   100
            TabIndex        =   59
            Top             =   765
            Width           =   1215
         End
      End
      Begin VB.TextBox Text2 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   220
         Width           =   4230
      End
      Begin VB.TextBox Text3 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   505
         Width           =   4230
      End
      Begin VB.TextBox Text4 
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   795
         Width           =   4230
      End
      Begin VB.TextBox Text18 
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
         Left            =   8055
         Locked          =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   790
         Width           =   4680
      End
      Begin VB.TextBox Text5 
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
         Left            =   8055
         Locked          =   -1  'True
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   505
         Width           =   4680
      End
      Begin VB.TextBox Text8 
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
         Left            =   8055
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   220
         Width           =   2055
      End
      Begin VB.TextBox Text17 
         Alignment       =   1  'Right Justify
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
         Left            =   11125
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1100
         Width           =   1605
      End
      Begin VB.TextBox Text20 
         Alignment       =   1  'Right Justify
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
         Left            =   12300
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1100
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.TextBox Text35 
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
         Height          =   285
         Left            =   2640
         TabIndex        =   46
         Top             =   1080
         Width           =   4230
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6975
         TabIndex        =   45
         ToolTipText     =   "Seleccion de Cliente"
         Top             =   1080
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   11340
         TabIndex        =   68
         Top             =   195
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   16777215
         Format          =   129630209
         CurrentDate     =   40810
      End
   End
   Begin VB.TextBox TEXT13 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   5070
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   43
      Top             =   7050
      Width           =   4500
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   15
      TabIndex        =   28
      Top             =   1560
      Width           =   12840
      Begin VB.TextBox TXTDESCRIPCION 
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
         Height          =   380
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   540
         Width           =   3855
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
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
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Text            =   "P.Venta"
         Top             =   270
         Width           =   1350
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
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
         Left            =   6300
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "Cantidad"
         Top             =   270
         Width           =   930
      End
      Begin VB.TextBox Text11 
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
         Left            =   2460
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "Descripción"
         Top             =   270
         Width           =   3855
      End
      Begin VB.TextBox Text12 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   38
         TabStop         =   0   'False
         Text            =   "Código"
         Top             =   270
         Width           =   2355
      End
      Begin VB.TextBox txtcodigo 
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
         Height          =   380
         Left            =   120
         TabIndex        =   37
         Top             =   540
         Width           =   2355
      End
      Begin VB.TextBox TXTPREUNIT 
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
         Height          =   380
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   36
         ToolTipText     =   "DOBLE CLICK PARA MODO EDICIÓN"
         Top             =   540
         Width           =   1350
      End
      Begin VB.TextBox TXTCANTIDAD 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   380
         Left            =   6300
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "   "
         ToolTipText     =   "('+')  PARA AGREGAR ('-') PARA BORRAR"
         Top             =   540
         Width           =   930
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
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
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "Importe"
         Top             =   270
         Width           =   1305
      End
      Begin VB.TextBox TXTIMPORTE 
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
         Height          =   380
         Left            =   11400
         Locked          =   -1  'True
         TabIndex        =   33
         ToolTipText     =   "('+')  PARA AGREGAR ('-') PARA BORRAR"
         Top             =   540
         Width           =   1305
      End
      Begin VB.TextBox Text16 
         Alignment       =   1  'Right Justify
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
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "Costo"
         Top             =   270
         Width           =   1350
      End
      Begin VB.TextBox Text15 
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
         Height          =   380
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   31
         ToolTipText     =   "DOBLE CLICK PARA MODO EDICIÓN"
         Top             =   540
         Width           =   1350
      End
      Begin VB.TextBox Text38 
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
         Height          =   380
         Left            =   8790
         Locked          =   -1  'True
         TabIndex        =   30
         ToolTipText     =   "DOBLE CLICK PARA MODO EDICIÓN"
         Top             =   540
         Width           =   1300
      End
      Begin VB.TextBox Text40 
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
         Left            =   8790
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "Vía de Entr."
         Top             =   270
         Width           =   1300
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "."
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
      Left            =   1920
      TabIndex        =   27
      Top             =   6660
      Width           =   255
   End
   Begin VB.TextBox Text23 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   1320
      TabIndex        =   26
      Top             =   6660
      Width           =   585
   End
   Begin VB.TextBox Text22 
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
      Left            =   2175
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6660
      Width           =   2850
   End
   Begin VB.TextBox Text39 
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
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "L. Precios"
      Top             =   6660
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Selecc.Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   15
      TabIndex        =   23
      ToolTipText     =   "Seleccion de Artículo (F2)"
      Top             =   2595
      Width           =   2490
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
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
      Left            =   10095
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Text            =   "Total:"
      Top             =   8055
      Width           =   975
   End
   Begin VB.TextBox Text14 
      Alignment       =   1  'Right Justify
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
      Left            =   11175
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   8055
      Width           =   1635
   End
   Begin VB.TextBox Text21 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   2600
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "MODIFICACIÓN DE PRESUPUESTO"
      ToolTipText     =   "Doble Click Para Salir de Modo Edición"
      Top             =   2570
      Visible         =   0   'False
      Width           =   5340
   End
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   5295
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   19
      Top             =   1560
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.TextBox TICAMBIS1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
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
      Left            =   8400
      MaxLength       =   14
      TabIndex        =   18
      Top             =   6720
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox Text24 
      Alignment       =   2  'Center
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
      Height          =   380
      Left            =   15
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "[F2] Presenta Lista de Códigos -  [F4] Edición de Precio de Venta -  [F5] Graba e Imprime Presupuesto Activo."
      Top             =   8400
      Width           =   12800
   End
   Begin VB.TextBox Total 
      Alignment       =   1  'Right Justify
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
      Left            =   11175
      TabIndex        =   16
      Text            =   " "
      Top             =   6720
      Width           =   1620
   End
   Begin VB.TextBox PORDESCU 
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
      Left            =   9735
      MaxLength       =   14
      TabIndex        =   15
      Top             =   7200
      Width           =   1065
   End
   Begin VB.TextBox Total 
      Alignment       =   1  'Right Justify
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
      Left            =   11175
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   " "
      Top             =   7080
      Width           =   1620
   End
   Begin VB.TextBox Text25 
      Alignment       =   2  'Center
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
      Left            =   9975
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "Sub-Total:"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.TextBox Text19 
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
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   7080
      Width           =   2850
   End
   Begin VB.TextBox Text26 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   1335
      TabIndex        =   11
      Top             =   7080
      Width           =   585
   End
   Begin VB.CommandButton Command2 
      Caption         =   "."
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
      Left            =   1935
      TabIndex        =   10
      Top             =   7080
      Width           =   255
   End
   Begin VB.TextBox Text27 
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
      Left            =   15
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "Cond. Pago"
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox Text28 
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
      Left            =   2190
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7560
      Width           =   2850
   End
   Begin VB.TextBox Text29 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   1335
      TabIndex        =   7
      Top             =   7560
      Width           =   585
   End
   Begin VB.CommandButton Command3 
      Caption         =   "."
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
      Left            =   1935
      TabIndex        =   6
      Top             =   7560
      Width           =   255
   End
   Begin VB.TextBox Text30 
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
      Left            =   15
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "Validez Hasta"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.TextBox Text31 
      Alignment       =   1  'Right Justify
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
      Left            =   11175
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1635
   End
   Begin VB.TextBox Text32 
      Alignment       =   2  'Center
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
      Left            =   10095
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "I.V.A.:"
      Top             =   7800
      Width           =   975
   End
   Begin VB.TextBox Text33 
      Alignment       =   1  'Right Justify
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
      Left            =   11175
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7560
      Width           =   1635
   End
   Begin VB.TextBox Text34 
      Alignment       =   2  'Center
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
      Left            =   10095
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Sub-Total:"
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   8565
      TabIndex        =   0
      ToolTipText     =   "Seleccion de Cliente"
      Top             =   1810
      Width           =   255
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   3675
      Left            =   0
      TabIndex        =   64
      Top             =   2880
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   6482
      _Version        =   393216
      BackColorFixed  =   16501405
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
   Begin VB.Label TICAMBIS 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pesos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5400
      TabIndex        =   65
      ToolTipText     =   "Doble Click para Cambiar Cotizacion"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   1935
      TabIndex        =   66
      Top             =   720
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label MONEMIS 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Pesos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6615
      TabIndex        =   67
      ToolTipText     =   "Doble-Click para 'Dólares'"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "T-C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5055
      TabIndex        =   69
      Top             =   6615
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "%:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10890
      TabIndex        =   70
      Top             =   7155
      Width           =   225
   End
   Begin VB.Menu mnu 
      Caption         =   "Menú"
      Begin VB.Menu MNUEDITPRESUP 
         Caption         =   "Editar Presupuesto"
      End
   End
   Begin VB.Menu pxcopia 
      Caption         =   "Pres.X.Copia"
      Begin VB.Menu mnupcopia 
         Caption         =   "Pres.x.Copia"
      End
   End
   Begin VB.Menu mnuvis 
      Caption         =   "Ver"
      Begin VB.Menu mnuvisualiza 
         Caption         =   "Visualizar Presupuesto"
      End
   End
   Begin VB.Menu MnuConfig 
      Caption         =   "Tablas de Configuración"
      Begin VB.Menu MnuTabEntr 
         Caption         =   "Tabla de Vias de Entrega"
      End
      Begin VB.Menu MnuArt_Prov 
         Caption         =   "Asociar Art. a Proveedor"
      End
      Begin VB.Menu MnuProv_Art 
         Caption         =   "Asociar Prov. a Artículo"
      End
      Begin VB.Menu mnuDemora 
         Caption         =   "Tabla de Demoras"
      End
      Begin VB.Menu MnuLimvias 
         Caption         =   "Límiites Segun Vias"
      End
      Begin VB.Menu mnutcoef 
         Caption         =   "Tipo de Coeficientes"
      End
      Begin VB.Menu mnuasigcoef 
         Caption         =   "Asignación de Coeficientes"
      End
   End
   Begin VB.Menu mnuSetup 
      Caption         =   "Setup"
      Begin VB.Menu mnuSetup1 
         Caption         =   "Setup"
      End
   End
End
Attribute VB_Name = "FRMPRESUA14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'HAY QUE DESPLEGAR DEL BOTON DE SELECCION DE CODIGO QUE MUESTRE LOS ITEMS DE LA LISTA DE PRECIOS
'ACTIVA , DEBERIA MOSTRAR LOS ITEMS CON PRECIO.
'VER TAMBIEN EL CASO EN QUE SE UTILICE LISTA 0 PARA QUE APAREZCAN TODOS LOS ITEMS EN EL BUSCADOR SIN PRECIO
'CARGA DE DETALLE POR TEXTO LIBRE

Function CANTI_ACUMU_LOTES()
   'SUMA LA CANTIDAD SELECCIONADA ACUMULADA DE LOTES
   '
   CANTI_ACUMU_LOTES = 0
   With INFOSIAHP
      For i& = 1 To .MSF5.Rows - 1
        LOTE$ = TRIM$(.MSF5.TextMatrix(i&, 1))
        CANTI = XVALO(.MSF5.TextMatrix(i&, 5))
        If LOTE$ <> "" Then
          If CANTI > 0 Then
            CANTI_ACUMU_LOTES = CANTI_ACUMU_LOTES + CANTI
          End If
        End If
      Next
   End With
   '
End Function

Function CARGALOTES$()
   'CARGA LOS LOTES SELECCIONADOS PARA EL ITEM - SOLO MODO DE CARGA MANUAL
   With INFOSIAHP
      For i& = 1 To .MSF5.Rows - 1
        LOTE$ = TRIM$(.MSF5.TextMatrix(i&, 1))
        CANTIX$ = SAFETEXT$(XVALO(.MSF5.TextMatrix(i&, 5)))
'        If LOTE$ <> "" Then
          If XVALO(CANTIX$) > 0 Then
            CARGALOTES$ = CARGALOTES$ & LOTE$ & ";" & CANTIX$ & ";"
          End If
'        End If
      Next
   End With
   '

   If InStrRev(CARGALOTES$, ";") = Len(CARGALOTES$) And CARGALOTES$ <> "" Then
      CARGALOTES$ = Left$(CARGALOTES$, Len(CARGALOTES$) - 1)
   End If
   '
   Debug.Print CARGALOTES$
   '
End Function

Sub CARGALOTES_ENGRILLA(FILA&)
   If FILA& < 1 Then Exit Sub
   '
   INFOSIAHP.MSF5.Rows = 1
   CADENA_LOTES$ = TRIM$(MSF1.TextMatrix(FILA&, 11))
   If CADENA_LOTES$ <> "" Then
      Call TEXTOLOTES(CADENA_LOTES$, ";")
      CI1% = CODINT%(MSF1.TextMatrix(FILA&, 1))
      NECESIDAD = XVALO(MSF1.TextMatrix(FILA&, 4))
     For i% = 1 To UBound(CADENATEX$) Step 2
      LOTE$ = CADENATEX$(i%)
      CANTI = CADENATEX$(i% + 1)
      ACUMUCANTI = ACUMUCANTI + XVALO(CANTI)
      STOCK1 = SaldoLote(CI1%, LOTE$)
      'OBTENER ESTOS DATOS Y CALCULAR EL PRECIO DE VENTA
      COSTOLOTE = 0
      VIADEENTREGA = 0
      PRECIOVENTA = 0
      GASTODENACION = 0
      '
      'SI NO ESTA SELECCIONADO AGREGO A NUEVA LISTA
      With INFOSIAHP
        .MSF5.Rows = .MSF5.Rows + 1
        .MSF5.TextMatrix(.MSF5.Rows - 1, 1) = LOTE$
        .MSF5.TextMatrix(.MSF5.Rows - 1, 2) = 100
        .MSF5.TextMatrix(.MSF5.Rows - 1, 3) = 110
        .MSF5.TextMatrix(.MSF5.Rows - 1, 4) = 120
        .MSF5.TextMatrix(.MSF5.Rows - 1, 5) = CANTI
      End With
     Next i%
'     INFOSIAHP.Text55 = NECESIDAD
     INFOSIAHP.Text5(2) = ACUMUCANTI
   End If
   
End Sub

Function COLORSTATUS$(CANTI, ST, PI, PC)
    'VERDE HAY STOCK
    ' AMARILLO: LA CANTIDAD PROYECTADA  +  EL STOCK ES - LO PEDIDO >= A LA DEMANDA
    'ROJO
    'LA CANTIDAD PROYECTADA  +  EL STOCK ES - LO PEDIDO <= A LA DEMANDA
    '
    CANTI1 = XVALO(CANTI)
    If CANTI > 0 Then
      ST1 = XVALO(ST)
      PI1 = XVALO(PI)
      PC1 = XVALO(PC)
      '
      If CANTI1 <= ST1 Then
         COLORSTATUS$ = "V"
      ElseIf (ST1 + PI1) - PC1 >= CANTI1 Then
         COLORSTATUS$ = "A"
      Else
         COLORSTATUS$ = "R"
      End If
    End If
End Function

Sub GENERA_TABLA_RELACION_ART_PROV()
   Call CREACAMPO("", "RELACODPROV", "Num_Prov", 4, 0, 1)
   Call CREACAMPO("", "RELACODPROV", "Cod_inte", 4, 0, 1)
   Call CREACAMPO("", "RELACODPROV", "Precio", 7, 0)
   Call CREACAMPO("", "RELACODPROV", "Moneda", 3, 2)
   Call CREACAMPO("", "RELACODPROV", "Fecha", 8, 0)
   Call CREACAMPO("", "RELACODPROV", "MarBorra", 3, 0)
   Call CREACAMPO("", "RELACODPROV", "Via", 3, 0)
End Sub

'Sub PRESENTA_INFO_ARTICULO(CIXI%, Optional FILA&)
'   '
'   Screen.MousePointer = 11
'   Call GENERA_TABLA_RELACION_ART_PROV
'   '
'   INFOSIAHP.txtcodigo = CODEXT$(CIXI%)
'   INFOSIAHP.TXTDESCRIPCION = DESCRIT0$(CIXI%)
'   INFOSIAHP.TXTCODICLIENTE = txtcodigo
''   INFOSIAHP.Text55 = XVALO(TXTCANTIDAD)
'
'   F$ = "F11." & CNTDC$(CI1%)  'FORMATO
'   ENCA$ = "Proveedor/A24;Stock/" & F$ & ";En Recep./" & F$ & ";Comp.Pend./" & F$ & ";Proy.Ing./" & F$
'   ENCA$ = ENCA$ + ";Ped.Cliente/" & F$ & ";NPROV/F0.0"
'   Call CARGA_ENCABEZADO(INFOSIAHP.MSF1, ENCA$, INFOSIAHP)
'   '
'   PESOX = PESUNI(CIXI%)
'
'   CI1% = CIXI%
'   'SI NO ESTA EN LA TABLA DE RELACIONES DE CODIGO PROVEEDOR PRESENTA LOS DATOS DEL MAESTRO Y EL STOCK GENERAL.
'   CONDI$ = "COD_INTE = " & CI1%
'   If CantRegistros&("RELACODPROV", CONDI$, "NOMESS") < 1 Then
'       j& = 1
'       INFOSIAHP.MSF1.Rows = 2
'       ST = FORMATNUM$(STOCKACT(CI1%), "F8.0")  ' STOCK POR PROVEEDOR
'       CANPED = FORMATNUM$(CANT_PEDIDA_PENDIENTE(CI1%), "F8.0")      ' CANTIDAD PEDIDA PENDIENTE DE ENTREGA ' NO CONSIDERA PROVEEDOR
'       COMPRAX = FORMATNUM$(COMPRA_PEND_DE_ENTR(0, CI1%), "F8.0")  ' CANTIDAD COMPRADA PENDIENTE DE ENTREGA
'       CANRECEP = FORMATNUM$(CANT_RECEP_PENDIENTE_APROB(0, CI1%), "F8.0")
'       '
'       'CARGO LA GRILLA
'       INFOSIAHP.MSF1.TextMatrix(j&, 1) = "Todos"
'       INFOSIAHP.MSF1.TextMatrix(j&, 2) = ST
'       INFOSIAHP.MSF1.TextMatrix(j&, 3) = CANRECEP
'       INFOSIAHP.MSF1.TextMatrix(j&, 4) = COMPRAX
'       INFOSIAHP.MSF1.TextMatrix(j&, 5) = XVALO(COMPRAX) + XVALO(CANRECEP)
'       '
'       INFOSIAHP.MSF1.TextMatrix(j&, 6) = CANPED
'       INFOSIAHP.MSF1.TextMatrix(j&, 7) = FORMATNUM$(PESOX, "F10.3")
'''       INFOSIAHP.MSF1.TextMatrix(j&, 10) = FECHATEX$(FECHA_COSTO_POR_PROV%(NPROV&, CI1%))
'       INFOSIAHP.MSF1.TextMatrix(j&, 7) = ""
''       GoTo 80
'
'   End If
'
'
'   SQLX$ = "SELECT DISTINCT NUM_PROV "
'   SQLX$ = SQLX$ + " FROM RELACODPROV WITH(NOLOCK) "
'   SQLX$ = SQLX$ + " Where COD_INTE = " & CI1%
'   '
'   INFOSIAHP.MSF1.Rows = 1
'   TOTST = 0: TOTCANRECEP = 0: TOTCOMPRAX = 0: TOTPROYING = 0: TOTCANPED = 0
'   j& = 0
'   Set RS = READSET(SQLX$)
'   With RS
'     Do While Not .EOF
'       j& = j& + 1
'       INFOSIAHP.MSF1.Rows = j& + 1
'       '
'       NPROV& = XVALO(!Num_Prov)
'       '
'       ST = FORMATNUM$(STOCK_PROV(NPROV&, CI1%), "F8.0")  ' STOCK POR PROVEEDOR
'       CANPED = FORMATNUM$(CANT_PEDIDA_PENDIENTE(CI1%), "F8.0")      ' CANTIDAD PEDIDA PENDIENTE DE ENTREGA ' NO CONSIDERA PROVEEDOR
'       COMPRAX = FORMATNUM$(COMPRA_PEND_DE_ENTR(NPROV&, CI1%), "F8.0")  ' CANTIDAD COMPRADA PENDIENTE DE ENTREGA
'       CANRECEP = FORMATNUM$(CANT_RECEP_PENDIENTE_APROB(NPROV&, CI1%), "F8.0")
'       '
'       'CARGO LA GRILLA
'       INFOSIAHP.MSF1.TextMatrix(j&, 1) = SAFETEXT$(TEL2CLI$(-1 * NPROV&))
'       INFOSIAHP.MSF1.TextMatrix(j&, 2) = ST: TOTST = TOTST + XVALO(ST)
'       INFOSIAHP.MSF1.TextMatrix(j&, 3) = CANRECEP: TOTCANRECEP = TOTCANRECEP + XVALO(CANRECEP)
'       INFOSIAHP.MSF1.TextMatrix(j&, 4) = COMPRAX: TOTCOMPRAX = TOTCOMPRAX + XVALO(COMPRAX)
'       PROYING = XVALO(COMPRAX) + XVALO(CANRECEP)
'       INFOSIAHP.MSF1.TextMatrix(j&, 5) = PROYING: TOTPROYING = TOTPROYING + XVALO(PROYING)
'
'       If j& = 1 Then 'SI ES LA PRIMERA VEZ PONE LA CANTIDAD PEDIDA YA QUE LA MISMA ES A MODO GENERAL
'          INFOSIAHP.MSF1.TextMatrix(j&, 6) = CANPED: TOTCANPED = TOTCANPED + XVALO(CANPED)
'          Call COLOREAR_GRILLA_SOLACELDA(INFOSIAHP.MSF1, vbRed, j&, 6, 1)
'       Else
'          INFOSIAHP.MSF1.TextMatrix(j&, 6) = ""
'       End If
'       INFOSIAHP.MSF1.TextMatrix(j&, 7) = SAFETEXT$(NPROV&)
'
'       .MoveNext
'     Loop
'   End With
'   '
'   RS.Close
'   Set RS = Nothing
'   'TOTALIZO
'   '
'   INFOSIAHP.MSF1.Rows = j& + 2
'   INFOSIAHP.MSF1.TextMatrix(j& + 1, 1) = "TODOS X FIFO"
'   Call COLOREAR_GRILLA_SOLACELDA(INFOSIAHP.MSF1, &H80FFFF, j& + 1, 1)
'    INFOSIAHP.MSF1.Row = j& + 1
'    INFOSIAHP.MSF1.COL = 1
'    INFOSIAHP.MSF1.CellFontUnderline = True 'subrayado
'
'   INFOSIAHP.MSF1.Rows = j& + 3
'   INFOSIAHP.MSF1.TextMatrix(j& + 2, 1) = "TOTALES:"
'   INFOSIAHP.MSF1.TextMatrix(j& + 2, 2) = TOTST
'   INFOSIAHP.MSF1.TextMatrix(j& + 2, 3) = TOTCANRECEP
'   INFOSIAHP.MSF1.TextMatrix(j& + 2, 4) = TOTCOMPRAX
'   INFOSIAHP.MSF1.TextMatrix(j& + 2, 5) = TOTPROYING
'   INFOSIAHP.MSF1.TextMatrix(j& + 2, 6) = TOTCANPED
'   'COLOREO EL TEXTO
'   Call COLOREAR_GRILLA_SOLAFILA(INFOSIAHP.MSF1, &HE0E0E0, j& + 2)
'   Call COLOREAR_GRILLA_SOLACELDA(INFOSIAHP.MSF1, &H800000, j& + 2, 1, 1)
'   Call COLOREAR_GRILLA_SOLACELDA(INFOSIAHP.MSF1, &H800000, j& + 2, 2, 1)
'   Call COLOREAR_GRILLA_SOLACELDA(INFOSIAHP.MSF1, &H800000, j& + 2, 3, 1)
'   Call COLOREAR_GRILLA_SOLACELDA(INFOSIAHP.MSF1, &H800000, j& + 2, 4, 1)
'   Call COLOREAR_GRILLA_SOLACELDA(INFOSIAHP.MSF1, &H800000, j& + 2, 5, 1)
'   Call COLOREAR_GRILLA_SOLACELDA(INFOSIAHP.MSF1, vbRed, j& + 2, 6, 1)
'   '
'      INFOSIAHP.txtcodigo = txtcodigo
'   INFOSIAHP.TXTDESCRIPCION = TXTDESCRIPCION
'
'   'MUESTRO LISTA DE LOTES SELECCIONADOS
''80 If FILA& > 0 Then
''     Call CARGALOTES_ENGRILLA(FILA&)
''   End If
'   '
'   Call SELECCIONARFILA(INFOSIAHP.MSF1, 1)
'
'   Screen.MousePointer = 1
'   INFOSIAHP.FILAX& = FILA&
'   INFOSIAHP.Show 1
'End Sub
Sub PRINTPRES()

    Call PRIPRESU(1)

End Sub
Sub PRIPRESU(VIPREVIA%)
    '
    FORIPRE$ = "PRSP-AQM"
    '
    If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") < 1 Then
       TX$ = "FALTA FORMULARIO DE IMPRESION DE REMITOS"
       TX$ = TX$ + "EL PRESUPUESTO SE HA GUARDADO, NO SALDRÁ IMPRESO"
       MsgBox TX$
       Exit Sub
    End If
    '
    Call LIMPIATETAB
    '
    TIPODOC$ = "Presupuesto" ' guardar el documento immpreso
'    NUMERI$ = "00000000" + TRIM$(Str$(Text17)): NUMERI$ = Right(NUMERI$, 8)
'    NUMERI$ = "PR-" & NUMERI$ & "-" & TRIM$(Text20)
    NUMERI$ = "00000000" + TRIM$(Str$(XVALO(Text17))): NUMERI$ = Right(NUMERI$, 8)
    VERSIOX$ = "00" & SAFETEXT$(XVALO(Text20)): VERSIOX$ = Right(VERSIOX$, 2)
    NUMERI$ = "PR-" & NUMERI$ + "-" & VERSIOX$

    NUMEDOC$ = NUMERI$
    
    'If VIPREVIA% = 1 Then
    '   CANCELPRINT% = 1
    'End If
    '
    'DATOS DE ENCABEZADO
    NC& = XVALO(Text1)
    
    CODPARAM$(1) = "FECH-EMI": DOCPARAM$(1) = FECHATEX$(CVINT(DTPicker1.Value))   'FECHA
    CODPARAM$(2) = "RASO-CLI": DOCPARAM$(2) = TRIM$(Text2)
    CODPARAM$(3) = "DOMI-CLI": DOCPARAM$(3) = TRIM$(Text3)
    CODPARAM$(4) = "LOCA-CLI": DOCPARAM$(4) = TRIM$(Text4)
    CODPARAM$(5) = "TELE-CLI": DOCPARAM$(5) = TRIM$(Text8)
    CODPARAM$(6) = "REF-MAT3":  DOCPARAM$(6) = TRIM$(Text18) 'MAIL
    CODPARAM$(7) = "REF-MAT4":  DOCPARAM$(7) = TRIM$(Text5) 'CELLUALR
    CODPARAM$(8) = "IMP-NETO":  DOCPARAM$(8) = TRIM$(Text14) 'SUBTOTAL
    CODPARAM$(9) = "IMP-TOTA":  DOCPARAM$(9) = TRIM$(Text14.TEXT) 'NONTO TOTAL
    CODPARAM$(10) = "NUME-COM":  DOCPARAM$(10) = NUMERI$ ' Text17 + 'NUMERO DE DOCUMENTO
        CODIPROVINCIA% = XVALO(VALOBADA("DEUDOR12", "Prov_Cli", CONDI$, "NOMESS"))
'    CODPARAM$(11) = "PCIA-CLI":  DOCPARAM$(11) = ECOPROVINCIA$(CODIPROVINCIA%) 'PROVINCIA
    CODPARAM$(12) = "OBSERVA": DOCPARAM$(12) = Text13
    '
    CAPARDOC% = 12
    '
    CACOLTAB1% = 1
    CODTABU1$(1) = "CANTIDAD"
    CODTABU1$(2) = "COD-MPRI"
    CODTABU1$(3) = "DES-MAT"
    CODTABU1$(4) = "UNIMED"
    CODTABU1$(5) = "PRE-NETO"
    CODTABU1$(6) = "NET-ITEM"
    CODTABU1$(7) = "REF-MAT5"
    CODTABU1$(8) = "REF-MAT6"
    CACOLTAB1% = 8
    '
    CAITAB1% = 0
    TIOPCI% = 0 ' BANDERA PARA IMPRIMIR TITULO DE OPCIONALES
    For U& = 1 To MSF1.Rows - 1
        CODE$ = TRIM$(MSF1.TextMatrix(U&, 1))
        CIIX% = CODINT%(CODE$)
        DESCRIX$ = MSF1.TextMatrix(U&, 2)
        CAN = XVALO(MSF1.TextMatrix(U&, 3))
        PREUNI# = XVALO(MSF1.TextMatrix(U&, 4))
        IMPORTX# = XVALO(MSF1.TextMatrix(U&, 5))
        OPCIONAL% = XVALO(MSF1.TextMatrix(U&, 6))
        PRESENTA$ = SAFETEXT$(VALOBADA("MASTER", "PiezasXCaja", "CODINT=" & CIIX%))
        If PRESENTA$ = "0" Then PRESENTA$ = ""
        
        ' IMPRIME TITULO DE PRODUCTOS OPCIONALES
        If OPCIONAL% = 1 And TIOPCI% = 0 Then
            CAITAB1% = CAITAB1% + 1
            If U& > 1 Then
                CAITAB1% = CAITAB1% + 1
                For UU% = 1 To CACOLTAB1%: TETABU1$(UU%, CAITAB1% - 1) = "": Next UU%
            End If
            TETABU1$(8, CAITAB1%) = "Productos opcionales:"
            CAITAB1% = CAITAB1% + 1
            For UU% = 1 To CACOLTAB1%: TETABU1$(UU%, CAITAB1%) = "": Next UU%
            TIOPCI% = 1
        End If
        
        ' IMPRIME DETALLE, DISTINGUIENDO SI EL PRODUCTO ES OPCIONAL O NO
        CAITAB1% = CAITAB1% + 1
        If OPCIONAL% = 1 Then
            COD$ = AJUSTI(CODEXT$(CIIX%), 16)
            DES$ = AJUSTI(DESCRIT0$(CIIX%), 22)
            UME$ = AJUSTI(UNIMED$(CIIX%), 3)
            PST$ = AJUSTI(PRESENTA$, 8)
            QTD$ = FORMATNUM$(CAN, "F8.1")
            PRE$ = FORMATNUM$(PREUNI#, "F8.2")
            TOT$ = FORMATNUM$(CAN * PREUNI#, "F9.2")
            TETABU1$(8, CAITAB1%) = COD$ & " " & QTD$ & " " & UME$ & " " & PST$ & " " & DES$ & " " & PRE$ & " " & TOT$
        Else
            TETABU1$(1, CAITAB1%) = TRIM$(FORMATNUM$(CAN, "F8.1"))
            TETABU1$(2, CAITAB1%) = CODEXT$(CIIX%)
            TETABU1$(3, CAITAB1%) = DESCRIT0$(CIIX%)
            TETABU1$(4, CAITAB1%) = TRIM$(UNIMED$(CIIX%))
            TETABU1$(5, CAITAB1%) = FORMATNUM$(PREUNI#, "F12.2")
            TETABU1$(6, CAITAB1%) = FORMATNUM$(IMPORTX#, "F12.2")
            TETABU1$(7, CAITAB1%) = PRESENTA$
            TETABU1$(8, CAITAB1%) = ""
        End If
        '
        ' AGREGA LOS CODIGOS QUE COMPONEN LA FORMULA DEL PRODUCTO
        If CantRegistros("FORMULAS", "CODICONJ=" & CIIX%, "NOMESS") > 0 Then
            SQLX$ = "SELECT * FROM FORMULAS WITH(NOLOCK) WHERE CODICONJ=" & CIIX%
            Set RST = READSET(SQLX$)
            With RST
                Do While Not .EOF
                    CI0% = XVALO(!CODIELEM)
                    QUNI# = XVALO(!USOBRUTO)
                    CAITAB1% = CAITAB1% + 1
                    TETABU1$(1, CAITAB1%) = ""
                    TETABU1$(2, CAITAB1%) = ""
                    TETABU1$(3, CAITAB1%) = ""
                    TETABU1$(4, CAITAB1%) = ""
                    TETABU1$(5, CAITAB1%) = ""
                    TETABU1$(6, CAITAB1%) = ""
                    TETABU1$(7, CAITAB1%) = ""
                    TETABU1$(8, CAITAB1%) = ""
                    
                    If OPCIONAL% = 1 Then
                        DESXX$ = "... " & AJUSTI(DESCRIT0$(CI0%), 25) & " " & FORMATNUM$(QUNI#, "F10.2")
                        TETABU1$(8, CAITAB1%) = DESXX$
                    Else
                        DESXX$ = "... " & AJUSTI(DESCRIT0$(CI0%), 25) & " " & FORMATNUM$(QUNI#, "F10.2")
                        TETABU1$(3, CAITAB1%) = DESXX$
                    End If
                    .MoveNext
                Loop
            End With
        End If
        '
        Call CARLINSEP         'Carga linea de separacion entre codigo y codigo
        '
    Next U&
    '
    ' VERIFICA SI EL CLIENTE TIENE MAIL. EN CASO AFIRMATIVO CONSULTA AL USUARIO SI QUIERE ENVIARLO POR ESTA VIA
    SENDBYMAIL$ = ""
    EMAIL$ = TRIM$(SAFETEXT$(VALOBADA("DEUDOR12", "Mail_Cli", "NUME_CLI=" & NC&)))
    If EMAIL$ <> "" Then
        RSP% = MsgBox("¿Desea enviar el presupuesto por mail?", vbQuestion + vbYesNo, "Envio por mail")
        If RSP% = 6 Then SENDBYMAIL$ = EMAIL$
    End If
    
    Call PRINTDOC("@" + FORIPRE$)

    SENDBYMAIL$ = ""
    
    'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
    DOCUNU& = DOCID&("PRESUPUESTO", NUMEDOC$, "")
5   Call MUESTRADOC(DOCUNU&)
    CANCELPRINT% = 0

99  Exit Sub
    '
End Sub


Private Sub Check1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        KeyAscii = 0
        Call LIMPIAR_TEXT_DETALLE
    End If
    
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        '
        'VALIDO EL CODIGO
        CI1% = CODINT%(txtcodigo)
        If CI1% < 1 Then
           MsgBox "CÓDIGO INEXISTENTE O NO VÁLIDO"
           Exit Sub
        End If

        'VALIDO QUE TENGA INGRESADA CANTIDAD
        If XVALO(TXTCANTIDAD.TEXT) < 0.05 Then
           MsgBox "FALTA CANTIDAD"
           TXTCANTIDAD.SetFocus
           Exit Sub
        End If
        'VALIDO QUE TENGA PRECIO
        If XVALO(TXTPREUNIT.TEXT) < 0.05 Then
           KeyAscii = 0
           MsgBox "FALTA PRECIO UNITARIO"
           TXTPREUNIT.SetFocus
           Exit Sub
        End If
        '
'        If VALIDA_COD_GRID(txtcodigo, FRMPRESU.MSF1, 1) > 0 Then
'         MsgBox "MODIFIQUE DATOS SOBRE EL MISMO"
'         GoTo 50
'        End If
        
        TXTPREUNIT.BackColor = &HE0E0E0 'CAMBIA A GRIS
        TXTPREUNIT.Locked = True 'DESHABILITO
        cargar_datos
        txtcodigo.SetFocus
50      Call LIMPIAR_TEXT_DETALLE

    End If
End Sub


Sub Command1_Click()
   Command1.Enabled = False
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     txtcodigo = ""
     txtcodigo = COD$
   End If
   Command1.Enabled = True
   

End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
'   Call CARGA_INDICE_LPRECIO
'   If RESPUESTA_BUSQUEDA$ <> "" Then Text23 = RESPUESTA_BUSQUEDA$
'   Command11.Enabled = True

End Sub

Private Sub Command19_Click()
    If XVALO(Text20) > 0 Then Exit Sub 'si es modificacion
    Command19.Enabled = False
    Call SELECHO("DEUDOR1")
    NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
    If NCLIE% > 0 Then
      Text1.TEXT = TRIM$(Str$(NCLIE%))
      'Call CARDACLI
      On Error Resume Next
      txtcodigo.SetFocus
      On Error GoTo 0
    End If

    Screen.MousePointer = 1
    Command19.Enabled = True
    
    
 
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   If DERACCE%("ALTACLI", "Alta de Nuevo Cliente") < 2 Then
      Exit Sub
   End If
   '
   Call APERBASDAT("CLIEN")

   Dim Tranclie As ADODB.Recordset
   '
   Screen.MousePointer = 11
   PRONUCLI& = XVALO(CONTROL$("", "PRONUCLI"))
   'If PRONUCLI& > 9999 Then PRONUCLI& = 0
   NROPRO& = PRONUCLI&
   If NROPRO& < 0 Then NROPRO& = 0
3  NROPRO& = NROPRO& + 1
   '
   'Set Tranclie = Badaclie.OpenRecordset("SELECT Nume_Cli FROM Deudor12 WHERE Nume_Cli = " & PRONUCLI&")
   'sql
   Set Tranclie = FLEXCONN.Execute(FILTSQL$("SELECT Nume_Cli FROM Deudor12 WHERE Nume_Cli = " & NROPRO&))
   If Not (Tranclie.EOF And Tranclie.BOF) Then 'si la encuentra (si el record no esta vacío)
      GoTo 3
   End If
'   '
   NCLIE = NROPRO&
   VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, MKI$(Int(NCLIE)), 13, 2)
   Call REPLA(VDEF$, Space$(62), 15, 62)
   Call REPLA(VDEF$, Chr$(1), 77, 1)
   Call REPLA(VDEF$, Space$(15), 78, 15)
   '
10 Screen.MousePointer = 1
   OBX$ = OBJPLA$("ALTACLIENACC", VDEF$)
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   '
   NCLINUE% = CVI(Mid$(OBX$, 13, 2))
   If NCLINUE% < 1 Or NCLINUE% > 29999 Then
      Call MENSERR(24, "Número de Cuenta no Válido.")
      GoTo 10
   End If
   '
   COCLINU$ = TRIM$(Left$(OBX$, 12))
   If COCLINU$ = "" Then
     COCLINU$ = "CLI-" + TRIM$(Str$(NCLINUE%))
     Call REPLA(VDEF$, COCLINU$, 1, 12)
     'Call MENSERR(24, "Codigo de Cuenta no Válido")
     'GoTo 10
   End If
   '
  
   Set Tranclie = FLEXCONN.Execute(FILTSQL$("SELECT Codi_Cli FROM Deudor12 where Codi_cli='" & COCLINU$ & "'"))
   If Not (Tranclie.EOF And Tranclie.BOF) Then
      Call MENSERR(24, "Ingreso no Válido.\Código de Cuenta Pre-existente.")
      GoTo 10
   End If
   '
   Set Tranclie = FLEXCONN.Execute(FILTSQL$("SELECT Nume_Cli FROM Deudor12 where Nume_cli=" & NCLINUE))
   If Not (Tranclie.EOF And Tranclie.BOF) Then
      Call MENSERR(24, "Ingreso no Válido.\Número de Cuenta Pre-existente.")
      GoTo 10
   End If
   '
   DCLINUE$ = TRIM$(Mid$(OBX$, 15, 62))
   If Len(DCLINUE$) < 1 Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Razón Social del Cliente")
      GoTo 10
   End If
   '
   PIVCLI% = Asc(Mid$(OBX$, 77, 1))
   If PIVCLI% < 0 Or PIVCLI% > 7 Then
      Call MENSERR(24, "Ingreso no Válido.\Posición IVA no Figura en Tablas.")
      GoTo 10
   End If
   '
   CUTT$ = TRIM$(Mid$(OBX$, 78, 15))
   If VALICUIT%(CUTT$) <> 1 Then
     If PIVCLI% <> 0 Then
       If PIVCLI% <> 5 Then
         Call MENSERR(24, "Ingreso no Válido.\Número de C.U.I.T. Incorrecto.")
         GoTo 10
       End If
     End If
   End If
   '
   If PIVCLI% = 0 And Len(CUTT$) > 12 Then
        Call MENSERR(24, "El Límite Máximo Para Número de Documento es de 12 Dígitos")
        GoTo 10
   End If
   '
   If CUTT$ <> "" Then
     'sql
     Set Tranclie = FLEXCONN.Execute(FILTSQL$("SELECT Nume_Cli FROM Deudor12 WHERE Cuit_Cli = '" & CUTT$ & "'"))
     If Not (Tranclie.EOF And Tranclie.BOF) Then
         If Tranclie!NUME_CLI > 0 Then
            NCLIX% = Tranclie!NUME_CLI
            If COMALTER%("Existe Registrada una Cuenta con el Mismo C.U.I.T.\  \'" + TRIM$(rasocli$(NCLIX%)) + "' (" + TRIM$(Str$(NCLIX%)) + ")\\Cancelar\Registrar Igual") <> 2 Then
                GoTo 10
            End If
         End If
     End If
   End If
   '
   Screen.MousePointer = 11
   '
   Dim Rdeudor12 As ADODB.Recordset
   Set Rdeudor12 = New ADODB.Recordset
   Rdeudor12.CursorType = adOpenKeyset
   Rdeudor12.LockType = adLockPessimistic
   Rdeudor12.Open FILTSQL$("Select * from Deudor12 where nume_cli=" & NCLINUE%), FLEXCONN, , , adCmdText
   With Rdeudor12
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
      !NUME_CLI = TRIM$(Str$(NCLINUE%))
      !Raso_Cli = DCLINUE$
      !Codi_Cli = COCLINU$
      !Cuit_Cli = CUTT$
      !Piva_Cli = PIVCLI%
      !Stat_Cli = 1
      '
      On Error Resume Next
      .Update
      If Err Then
         Call MENSERR(24, "Alta no Procesada Correctamente")
      End If
      On Error GoTo 0
   End With
   '
   Rdeudor12.Close
   Set Rdeudor12 = Nothing
   UPDATCLI% = 1
   
   Call CARGALISEL("DEUDOR1", "Deudor12", "Nume_Cli/F6.0>;Raso_Cli/A48", "Stat_Cli >= 0 and Nume_Cli>0 ORDER BY Raso_Cli")
   '
   Text1.TEXT = TRIM$(Str$(NCLINUE%))
   
99 Screen.MousePointer = 1
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Text1.Locked = True
   Command19.Enabled = False
   CUENTA& = XVALO(Text1)
   If CUENTA& < 1 Then
      MsgBox "FALTA NÚMERO DE CLIENTE"
      GoTo 99
   End If
   '

   'VALIDACION DE DETALLE
   If MSF1.Rows <= 1 Then
     TTXX$ = "FALTA INGRESAR DETALLE DE PRESUPUESTO"
     MsgBox TTXX$
     GoTo 99
   End If
   '
   Call Grabar_Presupuesto
   '
   Call PRINTPRES
'   Call LIMPIARPANTALLA(Text1)
   Text1 = ""
   Text17 = 1 + MAXNUMCOD&("PRESUVEN", "NroPres")
   'Text20 = "-00"

99 Command5.Enabled = True
   Text1.Locked = False 'TEXT DE CUENTA DE CLIENTE
   Command19.Enabled = True 'BOTON DE LISTA DE SELECCIN
   If Text21.Visible = True Then Text21.Visible = False 'CARTEL DE MODIFICACION DE PRESUPUESTO
   '
End Sub

Sub Grabar_Presupuesto()
    '
   OKX% = 0
   Call BLOQARCH("PRESUPU")
   '
   If Text21.Visible = True Then MODIFICA% = 1
   If MODIFICA% < 1 Then
    NPP& = NUPRESUP&
    If NPP& > XVALO(Text17.TEXT) Then
       Call CIERRARCH("PRESUPU")
       Call COMUNI("Número Ocupado")
       Text11.TEXT = TRIM$(Str$(NPP&))
       Exit Sub
    End If
   End If
   '
   Call ABREPRESVEN
   '\\\OT-06-0168-WAR-FIN///
   '
   If VALICLIE% < 1 Then
      MERRO$ = "Número de Cliente no Válido"
      GoTo 98
   End If
   NCLIE% = XVALO(Text1.TEXT)
   '
   If XVALO(Total(0).TEXT) < 1 Then
      MERRO$ = "Falta(n) Precio(s) de Lista"
      GoTo 98
   End If
   '
   If XVALO(Text14) < 1 Then
      MERRO$ = "Importe no Válido"
      GoTo 98
   End If
   '
   HOII% = HOY(HOS$)
   FEMI% = CVINT(DTPicker1.Value)
   If FEMI% < 1 Or FEMI% > HOII% Then
      MERRO$ = "Fecha de Emisión no Válida"
      GoTo 98
   End If
   '
   '
   NUOIT% = 0
   Call ABREPRESVEN
   CONDI$ = "NroPres= " & XVALO(Text17.TEXT)
   STRSQL = "Select * from Presuven"
   STRSQL = STRSQL + " where " & CONDI$
   Dim PresuVen As ADODB.Recordset
   Set PresuVen = New ADODB.Recordset
   PresuVen.Open FILTSQL$(STRSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   On Error Resume Next
   With PresuVen
   If MODIFICA% < 1 Then
    If Not (.EOF And .BOF) Then
       Call COMUNI("Número Ocupado")
       Text11.TEXT = TRIM$(Str$(NPP&))
       Call CIERRARCH("PRESUPU")
       Exit Sub
    End If
   End If
   STRSQL.Close
   Set STRSQL = Nothing
   '
   Call ACTUREGISTRO("PRESUVEN", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
   
   DESC$ = PORDESCU
   COEFIDESCU = COEFDESCU(DESC$)
   For U& = 1 To MSF1.Rows - 1
         CIIX% = CODINT%(MSF1.TextMatrix(U&, 1))
         CAN = XVALO(MSF1.TextMatrix(U&, 4))
         INETO = XVALO(MSF1.TextMatrix(U&, 8)) * (100 - XVALO(COEFIDESCU)) / (100)

         If CAN > 0.05 Then
           .AddNew   ' YA CONVERTIDO
           On Error Resume Next
           !CodiEmpr = CodiEmp%
           On Error GoTo 0
             !NroPres = XVALO(Text17.TEXT)
             '\\\OT-06-0168-WAR-09/05/06///
             !NumeCRubro = 0
             !Cod_Rubro = ""
             !Cod_Anio = ""
             '\\\OT-06-0168-WAR-FIN///
             !FechEmis = CVDAT(FEMI%)
             !NroClie = NCLIE%
             !RasoClie = Left$(Text2, 64)
             !Referen = "" 'Left$(Text10, 16)
             !AteSen = "" 'Left$(Text12, 48)
             !LisPre = XVALO(Text23)
             !ConPag = XVALO(Text26)
             !VENDED = 0
             !MonedEmi = MONEMI%
               MON% = MONEMI% '
               If MONEMI% = 1 Then MON% = 2 'SI ES PESOS GUARDA LA COTIZACION EN DOLARES, SI NO LA QUE CORRESPONDE
             !COTIZAMONEDA = FORMATNUM$(TICAMONE(MON%), "F12.4")
             !NuOrdIt = NUOIT%
               NUOIT% = NUOIT% + 1
             !CodiInt = CIIX%
             !CodiExt = Left$(CODEXT$(CIIX%), 24)
             !Descrip = Left$(DESCRIT0$(CIIX%), 64)
             !Especif = DESCRIT0$(CIIX%)
             !PreList = INETO
             !TDESCUEN = Left$(PORDESCU, 32)
                DESC$ = PORDESCU
                TTT1 = COEFDESCU(DESC$)
             !PorcDesc = TTT1
             !ImpoDesc = INETO * TTT1 / 100
             !PreUnit = INETO * (1 - TTT1 / 100)
             !Cantit = CAN
             !ImToNe = CAN * INETO * (1 - TTT1 / 100)
             !FechVent = CVDAT(0)
             !NroPedi = 0
             !NroFac_Cor = ""
             !NroFac_Lar = ""
             !Codi_Cli = CODICL$
             !Vers_Presu = 0
             !Num_Prov = XVALO(MSF1.TextMatrix(U&, 5))
             !COSTO_UNIT = Val(MSF1.TextMatrix(U&, 7))
             !Via_Entrega = XVALO(Left$(MSF1.TextMatrix(U&, 10), 2))
             .Update
'            TTXYZ$ = Space$(76)
'            Call REPLA(TTXYZ$, "- " + CODEXT$(CIIX%), 11, 18)
'            Call REPLA(TTXYZ$, DESCRIT0$(CIIX%), 29, 38)
'            Call REPLA(TTXYZ$, CSTRING$(MKS$(CAN), 3, 10, 1, 2), 57, 10)
'            Call AGREGA_ANOTAPROSPECT(CODICL$, TTXYZ$) ' AGREGA EL DETALLE EN EL ANOTADOR DEL PROSPECT
         End If
       Next U&
''''     End If
   End With
   PresuVen.Close
   Set PresuVen = Nothing
   '
   CONDI$ = CONDI$ & " AND MARBORRA = 1"
   Call BORRAREGISTROS("PRESUVEN", CONDI$, REGAF&, "NOMESS")
   MODIFICA% = 0
   Call OPLIPRE07.ACTINDIPRES(XVALO(Text11))
   Call CIERRARCH("PRESUPU")
   '
   '
   OKX% = 1
   Exit Sub
   '
98 Call CIERRARCH("PRESUPU")
   Call CIERRARCH("INDIPRES")
   Call MENSERR(24, MERRO$)
   '
End Sub


Function VALICLIE%()
   VCC% = 0
   If XVALO(Text1.TEXT) > 0 Then
     If XVALO(Text1.TEXT) <= 32767 Then
       NCLIE% = XVALO(Text1.TEXT)
       If REGICLI&(NCLIE%) > 0 Then
         VCC% = 1
       End If
     End If
   End If
   If VCC% = 0 Then
      CODIXX$ = TRIM$(Label26)
      CODICL$ = SAFETEXT$(VALOBADA("BSEGCOM", "Codi_Cli", "Codi_Cli ='" & CODIXX$ & "'"))
      If CODICL$ <> "" Then VCC% = 1
   End If
   '
99 VALICLIE% = VCC%
   '
End Function


Private Sub Command6_Click()
  Command6.Enabled = False
  Screen.MousePointer = 11

  For i& = 1 To MSF1.Rows - 1
       Call COLOREAR_GRILLA_SOLAFILA(FRMPRESUAHP.MSF1, vbBlue, i&)
       CI1% = CODINT%(MSF1.TextMatrix(i&, 1))
       NPROV& = XVALO(MSF1.TextMatrix(i&, 5))
       CANTI = XVALO(MSF1.TextMatrix(i&, 4))
       If CARGA(2).Value <> True Then
            ST = FORMATNUM$(STOCKACT(CI1%), "F8.0")  ' STOCK POR PROVEEDOR
            CANPED = FORMATNUM$(CANT_PEDIDA_PENDIENTE(CI1%), "F8.0")      ' CANTIDAD PEDIDA PENDIENTE DE ENTREGA ' NO CONSIDERA PROVEEDOR
            COMPRAX = FORMATNUM$(COMPRA_PEND_DE_ENTR(NPROV&, CI1%), "F8.0")  ' CANTIDAD COMPRADA PENDIENTE DE ENTREGA
            CANRECEP1 = FORMATNUM$(CANT_RECEP_PENDIENTE_APROB(NPROV&, CI1%), "F8.0")
            COLO$ = COLORSTATUS$(CANTI, ST, COMPRAX + CANRECEP1, CANPED)
       Else
            CANTILOTES = XVALO(MSF1.TextMatrix(i&, 12))
            CANTISOLICI = XVALO(MSF1.TextMatrix(i&, 4))
            If CANTILOTES <= CANTISOLICI Then COLO$ = "A"
            If CANTILOTES = 0 Then COLO$ = "R"
            If CANTILOTES >= CANTISOLICI Then COLO$ = "V"
            MSF1.TextMatrix(i&, 0) = COLO$
       End If
       
       'VUELVE EL COLOR A BLANCO
       Call COLOREAR_GRILLA_SOLAFILA(FRMPRESUAHP.MSF1, &HFEE7D6, i&)
       MSF1.TextMatrix(i&, 0) = COLO$
       If COLO$ = "V" Then Call COLOREAR_GRILLA_SOLACELDA(FRMPRESUAHP.MSF1, vbGreen, i&, 0)
       If COLO$ = "A" Then Call COLOREAR_GRILLA_SOLACELDA(FRMPRESUAHP.MSF1, vbYellow, i&, 0)
       If COLO$ = "R" Then Call COLOREAR_GRILLA_SOLACELDA(FRMPRESUAHP.MSF1, vbRed, i&, 0)
  Next i&
99     Screen.MousePointer = 1
       Command6.Enabled = True

End Sub

Private Sub Command7_Click()
   Call txtcodigo_Change
'  PROVEEDOR , VIA, DEMORA

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode 'ESTO ES PARA QUE DESP
        Case 113 'pulso  f2
        Call Command1_Click
        Case 116 'pulso  f5
        Command5_Click

    End Select
    
'   Call CREACAMPO("VENTAS", "RESE_COTI", "NroPres", 6, 4)
'   Call CREACAMPO("VENTAS", "RESE_COTI", "Status", 3, 0)
'   Call CREACAMPO("VENTAS", "RESE_COTI", "CodiInt", 4, 0)
'
'   Call CREACAMPO("VENTAS", "RESE_COTI", "CantBultos", 12, 0)
'   Call CREACAMPO("VENTAS", "RESE_COTI", "PesoNeto", 10, 128)
'   Call CREACAMPO("VENTAS", "RESE_COTI", "PesoTotal", 10, 128)
'   Call CREACAMPO("VENTAS", "RESE_COTI", "NCM", 12, 0)
'   Call CREACAMPO("VENTAS", "RESE_COTI", "CondVta", 10, 128)
'   Call CREACAMPO("VENTAS", "RESE_COTI", "Vers_Presu", 3, 0, 1)
'   Call CREACAMPO("VENTAS", "RESE_COTI", "Vers_Presu", 3, 0, 1)
'   Call CREACAMPO("VENTAS", "RESE_COTI", "Num_Prov", 4, 0, 1)
'   Call CREACAMPO("VENTAS", "RESE_COTI", "COSTO_UNIT", 6, 4)
'   Call CREACAMPO("VENTAS", "RESE_COTI", "Via_Entrega", 3, 0, 1)
'   Call CREACAMPO("VENTAS", "RESE_COTI", "MARBORRA", 3, 0, 1)
'   Call CREACAMPO("VENTAS", "RESE_COTI", "COTIZAMONEDA", 7, 0)

End Sub

Private Sub Form_Load()

   'Call APLICAR_SKIN(Me, App.Path & "\SKINS\dogmax.skn")
   'Call iniciar_variables

    Call CENTRAFORM(Me)
'    Show
    Refresh
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    '
15  EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
       EMPRE_SIN_USUARIO$ = REPLACAR$(EPAC$, USERNAME$, "")
       Caption = Caption + "  -  " + EMPRE_SIN_USUARIO$
    End If
    '
'    HOII% = HOY(HOS$)
'    Text9.TEXT = HOS$
    Text17.TEXT = TRIM$(Str$(NUPRESUP&))
    'Call NUCORUB
    '
    ARCHIPRE$ = "MASTER"
    '
    'DOLAR POR DEFECTO
    MONEMI% = 2  ' DOLAR
    MONEMIS = DEMONECO$(2)
    MONEMIS.Visible = True
    TICAMBIS = FORMATNUM$(TICAMONE(2), "F9.4")
    MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
    'CONVERPRECIO% = 1
    TICAMBIS.Visible = True
    Label11.Visible = True
    '
    If MONEMI% > 1 Then
      MONEMIS.ForeColor = RGB(255, 0, 0)
    End If
    '
    For U& = 1 To 2
       X$ = REGLEIDO$("GRUPIVA", U&)
       PRIVAX = CVS(Mid$(X$, 13, 4))
       If U& = 1 Then
           PORIVA0 = TRIM$(FORMATNUM$(PRIVAX, "F8.2"))
         Else
           PORIVA1 = TRIM$(FORMATNUM$(PRIVAX, "F8.2"))
       End If
    Next U&
    Call CIERRARCH("GRUPIVA")
    '
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    '
'
    DTPicker1.Value = Date
    
    Text17 = NUPRESUP&
    'Text20 = "-00"
    '
    Call CARGA_ENCABEZADO(FRMPRESUAHP.MSF1, "Código/A24;Cod.Cte/A24;Descripción/A24;Q/F4.A;Prov/AO;R.Social/A12;Costo/F10.2;Pr.Venta/F8.2;Importe/F10.2;Via Apr./A10;LOTES/A60;CANTI/F10", FRMPRESUAHP)
    '
    Call TEXTOLOTES$("A1;AA2;A3;", ";")
    Debug.Print TEX$
    Debug.Print CADENATEX$(1)
    Debug.Print CADENATEX$(2)
    Debug.Print CADENATEX$(3)
    Debug.Print UBound(CADENATEX$)
    'Debug.Print CADENATEX$(4)
    
End Sub
Function NUPRESUP&()
   '
   NPP& = 1
   Call ABREPRESVEN
   '
   Dim MIRS As ADODB.Recordset
   STRSQL = ("SELECT max(NroPres) as maximo FROM PresuVen")
   Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL))
   If Not IsNull(MIRS!Maximo) Then
      NPP& = 1 + MIRS!Maximo
   End If
   NUPRESUP& = NPP&
   MIRS.Close
   Set MIRS = Nothing
   '
End Function


Private Sub mnuedit_Click()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If XVALO(Text20) > 0 Then
      'Call APLICAR_SKIN(Me, App.Path & "\SKINS\dogmax.skn")
      Text1.Locked = False 'TEXT DE CUENTA DE CLIENTE
      Command19.Enabled = True 'BOTON DE LISTA DE SELECCIN
      If Text21.Visible = True Then Text21.Visible = False 'CARTEL DE MODIFICACION DE PRESUPUESTO
      Text1 = ""
   Else
      Close
      End
   End If
End Sub

Sub CARGA_PRESUPUESTOS()
    Screen.MousePointer = 11
    RESPUESTA_BUSQUEDA$ = ""

    'VALIDO POR SI NO HAY NINGUN DATO CARGADO
    CONDI$ = "NroPres > 0 "
    CNRG& = CantRegistros&("PRESUVENTA", CONDI$)
    If CNRG& < 1 Then
      MsgBox "No Hay Presupuestos Cargados"
      Exit Sub
    End If

    '
    LISTASELEC.List1.Clear
    LISTASELEC.List2.Visible = False
    LISTASELEC.List1.Visible = True
    '
    'ASIGNO EL ANCHO QUE VA A TENER CADA CAMPO A ESTAS 2 VARIABLES PUBLICAS
    CODIGO_LONG% = 14
    DESCRIP_LONG% = 48
    
    SQLX$ = " SELECT * "
    SQLX$ = SQLX$ + " FROM PRESUVENTA "
    SQLX$ = SQLX$ + " ORDER BY NroPres "

    Set TABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

    With TABTMP
        Do While (.EOF = False)
          NPRX$ = SAFETEXT$(!NroPres_Version)
          NCLIE& = XVALO(!NroClie)
          RAZONCLI$ = SAFETEXT$(!rasocli)
          FEX = CVINT(!FechEmis)
          '
          YY2$ = " - " + AJUSTI$(RAZONCLI$, 20)
          YY2$ = YY2$ + " - " + AJUSTD$(FECHATEX$(FEX), 8)
          COCU$ = TRIM$(NPRX$)
          
          Call REPLA(TTXX$, COCU$, 1, CODIGO_LONG%)
          Call REPLA(TTXX$, YY2$, CODIGO_LONG% + 2, DESCRIP_LONG%)
          JJ& = JJ& + 1
          LISTASELEC.List1.AddItem TTXX$
          '
          If NPX& <> NPXA& Then
            JJ& = JJ& + 1
          End If
          NPXA& = NPX&
          '
        .MoveNext
        Loop
      
    End With
    '
    LISTASELEC.Caption = "INDICE DE PRESUPUESTOS"
    LISTASELEC.List1.ListIndex = LISTASELEC.List1.ListCount - 1
    Screen.MousePointer = 1
60  LISTASELEC.Show 1

End Sub


Private Sub mnuabmcte_Click()
    ABMCLIENTES.Show 1
End Sub

Private Sub MnuArt_Prov_Click()
      '
         Call GENERA_TABLA_RELACION_ART_PROV
5        Call SELECHO("MASTER") ' DESPLIEGA LA LISTADO DE ARTICULOS
         '
         COD$ = TRIM$(VALACT1$("MASTER"))
         If COD$ = "" Then Exit Sub
         '
         CI% = CODINT%(COD$)
         ' BUSCO EN LA TABLA PROVEDCOD TODOS LOS REGISTROS QUE COINCIDAN CON EL CI%
         SQLX$ = "SELECT * FROM RELACODPROV WITH (NOLOCK)"
         SQLX$ = SQLX$ + "WHERE COD_INTE = " & CI% & " AND MARBORRA = 0 "
         Dim RS As ADODB.Recordset
         Set RS = FLEXCONN.Execute(FILTSQL$(SQLX$))
         '
         Call SETULTREG("!PROVIADE", 0)
         JJ& = 0
         With RS
            Do While Not .EOF ' RECORRO LA TABLA Y GRABO EN EL ARCHIVO PARA MOSTRAR LA VENTANA
                X$ = REGBLAN$("PROVIADE")
                Call REPLA(X$, MKI(XVALO(!Num_Prov)), 1, 2)
                Call REPLA(X$, MKD(XVALO(!Precio)), 3, 8)
                Call REPLA(X$, MKI(XVALO(!Moneda)), 11, 2)
                Call REPLA(X$, MKI(CVINT(!FECHA)), 13, 2)
                JJ& = JJ& + 1
                Call GRAREG("!PROVIADE", X$, JJ&)
            .MoveNext
            Loop
         End With
         RS.Close
         Set RS = Nothing
         
10       XYX% = VENTABU%("PROVIADE") ' MUESTRO LA VENTANA
         If XYX% < 0 Then GoTo 5     ' SI CIERRA VUELVE AL SELECHO
         '
         Screen.MousePointer = 11
         For U& = 1 To ULTREG&("!PROVIADE")
            X$ = REGLEIDO$("!PROVIADE", U&)
            COCO% = Asc(Left$(X$, 1))
            If COCO% < 1 Then
               Call MENSERR(24, "Código debe ser Numérico.")
               GoTo 10
            End If
            '
            DECO$ = TRIM$(Mid$(X$, 2, 32))
            If DECO$ = "" Then
               Call MENSERR(24, "Falta Descripción Para el Código Número '" + TRIM$(Str$(COCO%)) + "'.")
               GoTo 10
            End If
         '
         Next U&
      
         CONDI$ = "COD_INTE=" & CI%
         Call ACTUREGISTRO("RELACODPROV", "MARBORRA", CONDI$, 1, REG&) ' HACE UN MARBORRA

         SQLX$ = "SELECT * FROM RELACODPROV "
         SQLX$ = SQLX$ + " WHERE COD_INTE = " & CI%
         Set RS2 = New ADODB.Recordset
         RS2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          With RS2
            For i& = 1 To ULTREG&("!PROVIADE") ' RECORRO EL ARCHIVO Y GRABO
              XX$ = REGLEIDO$("!PROVIADE", i&) ' LA INFORMACION DE LA VENTANA
              NPRO% = CVI(Mid$(XX$, 1, 2))
              PRECIOX# = CVD(Mid$(XX$, 3, 8))
              MON% = CVI(Mid$(XX$, 11, 2))
              FE% = CVI(Mid$(XX$, 13, 2))
              
              .AddNew   ' YA CONVERTIDO
              
              On Error Resume Next
              !CodiEmpr = CodiEmp%
              On Error GoTo 0
              !Num_Prov = NPRO%
              !COD_INTE = CI%
              !VIA = 0
              If ROUND(PRECIOX#, 2) <> ROUND(XVALO(!Precio), 2) Then
                !Precio = PRECIOX#
                !Moneda = MON%
                FE% = HOY(HO$)
              End If
              '
              If FE% < 1 Then FE% = HOY(HO$)
              !FECHA = CVDAT(FE%)
              !MARBORRA = 0
              .Update
            Next i&
          End With
         RS2.Close
         Set RS2 = Nothing
         CONDI$ = "MARBORRA = 1"
         Call BORRAREGISTROS("RELACODPROV", CONDI$, REG&) ' BORRO LOS MARCADOS CON 1

         Call TRACENTRAL("PROVIADE", "")
         Call FRESHECHO("PROVIADE")
         Screen.MousePointer = 1
         Call COMUNI("Grabación Finalizada")
End Sub

Private Sub mnuasigcoef_Click()
      Call TRALOCAL("PORCCOEF", "")
      '
10    Screen.MousePointer = 1
      VTB% = VENTABU%("PORCCOEF")
      If VTB% < 0 Then Exit Sub
      '
      Screen.MousePointer = 11
      For U& = 1 To ULTREG&("!PORCCOEF")
         X$ = REGLEIDO$("!PORCCOEF", U&)
         COCO% = Asc(Left$(X$, 1))
         If COCO% < 1 Then
            Call MENSERR(24, "Falta Código de Via de Aprovisionamiento - Fila: " & U&)
            GoTo 10
         End If
         '
         COCO2% = Asc(Mid$(X$, 3, 2))
         If COCO2% < 1 Then
            Call MENSERR(24, "Falta Código de Tipo de Coeficiente - Fila: " & U&)
            GoTo 10
         End If
         '
         PORC = Asc(Mid$(X$, 3, 2))
         If PORC < 0.005 Then
            Call MENSERR(24, "Falta Coeficiente - Fila: " & U&)
            GoTo 10
         End If

      Next U&
      '
      Call TRACENTRAL("PORCCOEF", "")
      Call FRESHECHO("PORCCOEF")
       Screen.MousePointer = 1
      Call COMUNI("Grabación Finalizada")


End Sub

Private Sub mnuDemora_Click()
      Call TRALOCAL("DEMORAS", "")
      '
10    Screen.MousePointer = 1

      VTB% = VENTABU%("DEMORAS/TITUPAN=Demoras Según Proveedor y Via de Entrega")
      If VTB% < 0 Then Exit Sub
      '
      
      Screen.MousePointer = 11
      For U& = 1 To ULTREG&("!DEMORAS")
         X$ = REGLEIDO$("!DEMORAS", U&)
         COCO% = Asc(Left$(X$, 2))
         If COCO% < 1 Then
            Call MENSERR(24, "Código debe Vía de Entrega Debe ser Numérico.")
            GoTo 10
         End If
         '
      Next U&
      '
      'VALIDO QUE NO SE REPITA LA COMBINACION DE CODIGO PROVEEDOR Y VIA DE ENTREGA
      For U& = 1 To ULTREG&("!DEMORAS")
         X$ = REGLEIDO$("!DEMORAS", U&)
         COCO% = Asc(Left$(X$, 2))
         CODEMORA% = Asc(Mid$(X$, 3, 2))
         For J& = 1 To ULTREG&("!DEMORAS")
            If J& <> U& Then
              X$ = REGLEIDO$("!DEMORAS", J&)
              COCO2% = Asc(Left$(X$, 2))
              CODEMORA2% = Asc(Mid$(X$, 3, 2))
              If COCO% = COCO2% Then
                If CODEMORA% = CODEMORA2% Then
                   Call MENSERR(24, "Código Proveedor (" & CODEMORA% & ") y Código Demora  (" & COCO2% & ") Repetido.")
                   GoTo 10
                End If
              End If
            End If
          Next J&
      Next U&
      '
      Call TRACENTRAL("DEMORAS", "")
      Call FRESHECHO("DEMORAS")
      '
      Screen.MousePointer = 1
      Call COMUNI("Grabación Finalizada")
End Sub


 


Private Sub MNUEDITPRESUP_Click()
    '
    CONDI$ = "NroPres > 0 "
10  Call CARGALISEL("INDICEPR", "PRESUVEN", "NROPRES/F6.0;NroClie/F6.0;RasoClie/A24", CONDI$, "DISTINCT")
    Call SELECHO("INDICEPR/INDICE DE PRESUPUESTOS")
    NPR& = XVALO(VALACT1("INDICEPR"))
    If NPR& < 1 Then Exit Sub
    
    '
    Text17 = TRIM(Str$(NPR&))
    Text1 = VALACT1("INDICEPR")
    Text20 = "" 'para que siempre dispare el evento change
    Text1.Locked = True
    Command19.Enabled = False
    '
    Call PREPORCOP(NPR&)
    CONDI$ = "NroPres = " & NPR&
    NROPE& = XVALO(VALOBADA("PRESUVEN", "NroPedi", CONDI$, "NOMESS"))
    If NROPE& < 1 Then
       Text21.TEXT = "MODIFICACIÓN DE PRESUPUESTO "
    Else
       Text21.TEXT = "MODIFICACIÓN DE PRESUPUESTO - PEDIDO: " & NROPE&
    End If
    
    Text21.Visible = True

    '
End Sub
Function SIG_VERSION%(PRESUPUESTOX$)
       'PR-00000063-00
       NUMEPRE& = XVALO(Mid$(PRESUPUESTOX$, 4, 8))
       CONDI$ = "NroPres = " & NUMEPRE&
       SIG_VERSION% = 1 + VALOBADA("PRESUVENTA", "MAX(RIGHT(NroPres_Version,2))", CONDI$, "NOMESS")
End Function

Private Sub MnuLimvias_Click()
      Call TRALOCAL("KILOVIAS", "")
      '
10    Screen.MousePointer = 1

      VTB% = VENTABU%("KILOVIAS")
      If VTB% < 0 Then Exit Sub
      '
      Screen.MousePointer = 11
      For U& = 1 To ULTREG&("!KILOVIAS")
         X$ = REGLEIDO$("!KILOVIAS", U&)
         COCO% = Asc(Left$(X$, 2))
         If COCO% < 1 Then
            Call MENSERR(24, "Falta Código de Vía de Entrega")
            GoTo 10
         End If
         '
      Next U&
      '
      'VALIDO QUE NO SE REPITA LA COMBINACION DE CODIGO PROVEEDOR Y VIA DE ENTREGA
      For U& = 1 To ULTREG&("!KILOVIAS")
         X$ = REGLEIDO$("!KILOVIAS", U&)
         COCO% = Asc(Left$(X$, 2))
         For J& = 1 To ULTREG&("!KILOVIAS")
            If J& <> U& Then
              X$ = REGLEIDO$("!KILOVIAS", J&)
              COCO2% = Asc(Left$(X$, 2))
              If COCO% = COCO2% Then
                 Call MENSERR(24, "Código Via Repetido (" & COCO2% & ") ")
                 GoTo 10
              End If
            End If
          Next J&
      Next U&
      '
      Call TRACENTRAL("KILOVIAS", "")
      Call FRESHECHO("KILOVIAS")
      '
      Screen.MousePointer = 1
      Call COMUNI("Grabación Finalizada")
      '
End Sub

Private Sub mnupcopia_Click()

'   If XVALO(Text20) > 0 Then Exit Sub ' modo modificacion
    CONDI$ = "NroPres > 0 "
10  Call CARGALISEL("INDICEPR", "PRESUVEN", "NROPRES/F6.0;NroClie/F6.0;RasoClie/A24", CONDI$, "DISTINCT")
    Call SELECHO("INDICEPR/INDICE DE PRESUPUESTOS")
    NPR& = XVALO(VALACT1("INDICEPR"))
    If NPR& < 1 Then Exit Sub
    
    Call SELECHO("DEUDOR1")
    NCLIE% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
    If NCLIE% > 0 Then
      Text1 = NCLIE%
      Call PREPORCOP(XVALO(NPR&))
    End If
   '
End Sub
Sub PREPORCOP(NPRESU&)
   '
   NUCLI% = NCLIE% 'CVI(Mid$(XX$, 7, 2))'CLIENTE QUE SE ELIGIO
   Call ABREPRESVEN
   '
   SQLX$ = " SELECT * "
   SQLX$ = SQLX$ + " FROM PRESUVEN WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NroPres= " & NPRESU&
   SQLX$ = SQLX$ + " ORDER BY NuOrdIt "
   '
   Dim PRESUTEMP As ADODB.Recordset
   Set PRESUTEMP = READSET(SQLX$)
    '
   MSF1.Rows = 1
   Screen.MousePointer = 11
   With PRESUTEMP
      Do While Not .EOF
        i& = i& + 1
        CI1% = XVALO(!CodiInt)
        CODE$ = SAFETEXT$(!CodiExt)
        DESCRIX$ = SAFETEXT$(!Descrip)
        CANTI = XVALO(!Cantit)
        INETO = XVALO(!PreUnit)
        PREUNI# = INETO * (1 + TTT1 / 100)
        IMPORTX# = CANTI * PREUNI#
        MSF1.Rows = MSF1.Rows + 1
        MSF1.TextMatrix(i&, 1) = CODE$
        MSF1.TextMatrix(i&, 2) = "CODIGO CLIENTE"
        MSF1.TextMatrix(i&, 3) = DESCRIX$
        MSF1.TextMatrix(i&, 4) = TRIM$(FORMATNUM$(CANTI, "F8.1"))
        MSF1.TextMatrix(i&, 5) = TRIM$(XVALO(!Num_Prov))
        MSF1.TextMatrix(i&, 6) = TRIM$(rasocli(-1 * XVALO(!Num_Prov)))
        MSF1.TextMatrix(i&, 7) = TRIM$(FORMATNUM$(XVALO(!COSTO_UNIT), "F12.2"))
        MSF1.TextMatrix(i&, 8) = TRIM$(FORMATNUM$(PREUNI#, "F12.2"))
        MSF1.TextMatrix(i&, 9) = TRIM$(FORMATNUM$(IMPORTX#, "F12.2"))
        MSF1.TextMatrix(i&, 10) = SAFETEXT$(!Via_Entrega)
        

        PORDESCU = SAFETEXT$(!PorcDesc)
        

      .MoveNext
      Loop
   End With
   PRESUTEMP.Close
   Set PRESUTEMP = Nothing
   '
   Screen.MousePointer = 1
   Call ACTU_TOTAL
   ALPORCOPIA% = 1 'MARCA PARA SABER SI ESTA EN ALTA POR COPIA
End Sub

Private Sub MnuProv_Art_Click()
         '
         Call GENERA_TABLA_RELACION_ART_PROV
5        Call SELECHO("PROVED1")
         '
         NPRO% = XVALO(VALACT1("PROVED1"))
         If NPRO% < 1 Then Exit Sub
         '
         SQLX$ = "SELECT * FROM RELACODPROV  WITH(NOLOCK)"
         SQLX$ = SQLX$ + " WHERE NUM_PROV = " & NPRO% & " AND MARBORRA = 0 "
         Set RS = READSET(SQLX$)
         '
         Call SETULTREG("!ARTPROVD", 0)
         With RS
            On Error Resume Next
            .MoveFirst
            If Err Then
                On Error GoTo 0
                GoTo 10
            End If
            JJ& = 0
             Do While (.EOF = False)
                X$ = REGBLAN$("ARTPROVD")
                Q = !VIA
                Call REPLA(X$, MKI(XVALO(!COD_INTE)), 1, 2)
                Call REPLA(X$, MKD(XVALO(!Precio)), 3, 8)
                Call REPLA(X$, MKI(XVALO(!Moneda)), 11, 2)
                Call REPLA(X$, MKI(CVINT(!FECHA)), 13, 2)
                JJ& = JJ& + 1
                Call GRAREG("!ARTPROVD", X$, JJ&)
            .MoveNext
            Loop
         End With
         RS.Close
         Set RS = Nothing
         '
10       XYX% = VENTABU%("ARTPROVD")
         If XYX% < 0 Then GoTo 5
         '
         CONDI$ = "NUM_PROV =" & NPRO%
         Call ACTUREGISTRO("RELACODPROV", "MARBORRA", CONDI$, 1, REG&)
    
         
         SQLX$ = "SELECT * FROM RELACODPROV "
         SQLX$ = SQLX$ + " WHERE NUM_PROV = " & NPRO%
         Set RS1 = New ADODB.Recordset
         RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           With RS1
             For i& = 1 To ULTREG&("!ARTPROVD")
               XX$ = REGLEIDO$("!ARTPROVD", i&)
               
               CI% = CVI(Mid$(XX$, 1, 2))
               PRECIOX# = CVD(Mid$(XX$, 3, 8))
               MON% = CVI(Mid$(XX$, 11, 2))
               FE% = CVI(Mid$(XX$, 13, 2))
               CI% = CVI(Mid$(XX$, 1, 2))
               .AddNew   ' YA CONVERTIDO
               On Error Resume Next
               !CodiEmpr = CodiEmp%
               On Error GoTo 0
               !Num_Prov = NPRO%
               !COD_INTE = CI%
               !Precio = PRECIOX#
               !Moneda = MON%
               If FE% < 1 Then FE% = FE% = HOY(HO$)
               If ROUND(PRECIOX#, 2) <> ROUND(XVALO(!Precio)) Then FE% = HOY(HO$)
               !FECHA = CVDAT(FE%)
               !VIA = 0
               !MARBORRA = 0
               .Update
             Next i&
           End With
         RS1.Close
         Set RS1 = Nothing
         CONDI$ = CONDI + "AND  MARBORRA = 1  "
         Call BORRAREGISTROS("RELACODPROV", CONDI$, REG&)
         
         Call TRACENTRAL("ARTPROVD", "")
         Call FRESHECHO("ARTPROVD")
         Screen.MousePointer = 1
         Call COMUNI("Grabación Finalizada")

End Sub

Private Sub mnuSetup1_Click()
   OPPRES07.Show 1
End Sub

Private Sub MnuTabEntr_Click()
      Call TRALOCAL("VIASENTR", "")
      '
10    Screen.MousePointer = 1
      VTB% = VENTABU%("VIASENTR")
      If VTB% < 0 Then Exit Sub
      '
      Screen.MousePointer = 11
      For U& = 1 To ULTREG&("!VIASENTR")
         X$ = REGLEIDO$("!VIASENTR", U&)
         COCO% = Asc(Left$(X$, 1))
         If COCO% < 1 Then
            Call MENSERR(24, "Código debe ser Numérico.")
            GoTo 10
         End If
         '
         DECO$ = TRIM$(Mid$(X$, 2, 32))
         If DECO$ = "" Then
            Call MENSERR(24, "Falta Descripción Para el Código Número '" + TRIM$(Str$(COCO%)) + "'.")
            GoTo 10
         End If
         '
      Next U&
      '
      Call TRACENTRAL("VIASENTR", "")
      Call FRESHECHO("VIASENTR")
       Screen.MousePointer = 1
      Call COMUNI("Grabación Finalizada")
      'POSICIONES
      'DE PORCENTAJE 35,8
     '
End Sub

Private Sub mnutcoef_Click()
      Call TRALOCAL("TIPOCOEF", "")
      '
10    Screen.MousePointer = 1
      VTB% = VENTABU%("TIPOCOEF")
      If VTB% < 0 Then Exit Sub
      '
      Screen.MousePointer = 11
      For U& = 1 To ULTREG&("!TIPOCOEF")
         X$ = REGLEIDO$("!TIPOCOEF", U&)
         COCO% = Asc(Left$(X$, 1))
         If COCO% < 1 Then
            Call MENSERR(24, "Código debe ser Numérico.")
            GoTo 10
         End If
         '
         DECO$ = TRIM$(Mid$(X$, 2, 32))
         If DECO$ = "" Then
            Call MENSERR(24, "Falta Descripción Para el Código Número '" + TRIM$(Str$(COCO%)) + "'.")
            GoTo 10
         End If
         '
      Next U&
      '
      Call TRACENTRAL("TIPOCOEF", "")
      Call FRESHECHO("TIPOCOEF")
       Screen.MousePointer = 1
      Call COMUNI("Grabación Finalizada")

End Sub

Private Sub mnuvisualiza_Click()

10  Call CARGA_PRESUPUESTOS
    If RESPUESTA_BUSQUEDA$ = "" Then Exit Sub
   

    TIDOCUM$ = "Presupuesto"
    NUDOCU$ = RESPUESTA_BUSQUEDA$
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        PPXX% = InStr(XX$, "Nro.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
        POS1% = InStr(XX$, "Presupuesto")
        If POS1% > 0 Then
           NUMECOMPA$ = TRIM$(Mid$(XX$, POS1% + 12))
           If NUMECOMPA$ = TRIM$(NUDOCU$) Then
                DOCUNU& = CVS(Left$(XX$, 4))
                GoTo 5
           End If
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    MsgBox "DOCMENTO NO ENCONTRADO"
    GoTo 10
   '
5   Call MUESTRADOC(DOCUNU&)
    GoTo 10
End Sub


Private Sub MONEMIS_DblClick()
   '
   'VALIDO QUE NO HAYA NINGUN DATO EN LOS CASILLEROS DE ARRIBA ANTES DE MODIFICAR LA MONEDA X Q ESTA YA ESTA CONVERTIDA A LA MONEDA ACTIVA.
   If TRIM$(txtcodigo) <> "" Then
      Call COMUNI("Imposible Cambiar de Moneda Cuando Hay Información en Casilleros de Ingreso\Borre la Información y Vuelva a Intentar")
      Exit Sub
   End If
   
   Static CAMON%, DMONE$(), TCMONE()
   If CAMON% < 1 Then
     CAMON% = ULTREG&("TAMONED")
     If CAMON% < 1 Then
         CAMON% = 1
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         DMONE$(1) = "Pesos"
         TCMONE(1) = 1
       Else
         ReDim DMONE$(CAMON%), TCMONE(CAMON%)
         For U& = 1 To CAMON%
           XX$ = REGLEIDO$("TAMONED", U&)
           DMX$ = TRIM$(Left$(XX$, 8)) + " "
           PPXX% = InStr(DMX$, " ")
           DMX$ = Left$(DMX$, PPXX% - 1)
           '
           VHX = CVS(Mid$(XX$, 29, 4)): If VHX < 0.005 Then VHX = 1
           DMONE$(U&) = DMX$
           TCMONE(U) = VHX
         Next U&
     End If
   End If
   '
   MONEMIA% = MONEMI%
15 MONEMI% = MONEMI% + 1
   If MONEMI% > CAMON% Then MONEMI% = 1
   If DMONE$(MONEMI%) = "" Then GoTo 15
   '
   MONEMIS = DMONE$(MONEMI%)
   MONEMIS.ForeColor = RGB(0, 0, 255)
   If MONEMI% > 1 Then
      MONEMIS.ForeColor = RGB(255, 0, 0)
   End If
   '
   MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
   TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
   '
   If MONEMI% > 1 Then
     TICAMBIS.Visible = True
     'TICAMBIS.Enabled = False
   End If
   '
   Total(0) = FORMATNUM$(XVALO(Total(0)) * TICAMONE(MONEMIA%) / TICAMONE(MONEMI%), "F12.2")
   '
      COEFCON = TICAMONE(MONEMIA%) / TICAMONE(MONEMI%)
      For U& = 1 To MSF1.Rows - 1

        CAN = XVALO(MSF1.TextMatrix(U&, 3))
        'PORDESCUIN = CVS(Mid$(X$, 69, 4))
        COSTOX# = COEFCON * XVALO(MSF1.TextMatrix(U&, 7))
        PRECVENTA# = COEFCON * XVALO(MSF1.TextMatrix(U&, 8))
        IMPORTEX# = COEFCON * XVALO(MSF1.TextMatrix(U&, 9))
        MSF1.TextMatrix(U&, 7) = FORMATNUM$(COSTOX#, "F12.3")
        MSF1.TextMatrix(U&, 8) = FORMATNUM$(PRECVENTA#, "F12.3")
        MSF1.TextMatrix(U&, 9) = FORMATNUM$(IMPORTEX#, "F12.2")
        
      Next U&
      '
      Call ACTU_TOTAL
'      Call CARLISPRES
'      Call CIERRARCH("!CARDETPR")
'      Call CALTOPRES
'   End If
   '
End Sub




Private Sub MSF1_DblClick()
    i& = MSF1.MouseRow
    J& = MSF1.MouseCol
    If i& = 0 Or i& > MSF1.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF1.Row = i&
    MSF1.COL = J&
    On Error Resume Next

    REG& = i&
    If REG& < 1 Then
        Exit Sub
    End If
    '
    '
    OPX% = COMALTER%("Realmente Desea Quitar el Artículo: " & MSF1.TextMatrix(REG&, 3) & " Del Detalle de este Presupuesto\\Cancelar\Quitar")
    If OPX% = 2 Then
       REGI% = REG&
       Call borrafila(REGI%, FRMPRESUAHP, MSF1)
    End If
    '
    Call ACTU_TOTAL
    '
    On Error Resume Next
    VISIROWS = (MSF1.Height / MSF1.RowHeight(J&)) - 4
    If Err > 0 Then
      VISIROWS = 1
    End If
    On Error GoTo 0
    If VISIROWS > 0 Then
       If MSF1.TopRow < MSF1.Row - VISIROWS Then
          MSF1.TopRow = MSF1.Row - VISIROWS
       End If
    End If


End Sub

Private Sub MSFlexGrid2_Click()

End Sub

Private Sub MSF1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'     If Button = 2 Then
'        REG& = MSF1.MouseRow
'        If REG& < 1 Then Exit Sub
'        CI1% = CODINT%(MSF1.TextMatrix(REG&, 1))
'        Call PRESENTA_INFO_ARTICULO(CI1%, REG&)
'        Call cargar_datos
'     End If
     
End Sub

Private Sub PORDESCU_Change()
   Call ACTU_TOTAL
End Sub

Private Sub Text1_Change()
   
   
   Call LIMPIAR_TEXT_DETALLE
   Call LIMPIARPRESU
   MSF1.Rows = 1
   
   If TRIM$(Text1.TEXT) <> "" Then
      Call CARGADATOS_CLIENTE
   End If
   NC% = XVALO(Text1)
   'VALIDO SI LA LISTA ACTIVA ES VALIDA
   If XVALO(Text1) > 0 Then
    NLIST% = XVALO(Text23)
'    If NLIST% < 1 Or TRIM$(Text22) = "" Then
'      TX$ = "El CLIENTE ACTIVO NO POSEE LISTA DE PRECIOS ASIGNADA" + vbCrLf
'      TX$ = TX$ + "POR DEFECTO SE ASIGNA LISTA 1"
'      MsgBox TX$
'      NLIST% = 1
'      Text23 = NLIST%
'    End If
   End If
   Label26 = CODICLI$(NC%)
   'NUMERO DE PRESUPUESTO
'   If Text21.Visible = False Then
'    Text17 = 1 + MAXNUMCOD&("PRESUVENTA", "NroPres")
'    If XVALO(Text20) < 1 Then Text20 = "-00"
'    Text21.Visible = False
'
'   End If
   '
   'FECHA
   DTPicker1.Value = Date
   
'   If Command1.Enabled = True Then
'       Call FRMPRESUAHP.Command1_Click
'   End If
End Sub

Sub LIMPIARPRESU() 'SE LE PUEDE PASAR UN TEXTBOX QUE ESTE NO SE LIMPIE
    
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text8 = ""
    Text18 = ""
    Text13 = ""
    Text14 = ""
    Text19 = ""
    Text23 = ""
    Text20 = ""
    '
    MONEMI% = 2
    MONEMIS.Caption = Left$(DEMONECO$(MONEMI%), 8)
    TICAMBIS = FORMATNUM$(TICAMONE(MONEMI%), "F9.4")
    MONEMIS.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MONEMI%), "F10.4")) + " - Doble-Click para Cambiar Moneda"
    MONEMIS.ForeColor = RGB(0, 0, 255)
'    TICAMBIS.Visible = False

End Sub
Sub CARGADATOS_CLIENTE()
   CUENTA& = XVALO(Text1)
   If CUENTA& < 1 Then: Call LIMPIARPRESU: Exit Sub
   '
   CONDI$ = "Nume_Cli = " & CUENTA&
   '
   'VALIDO SI EXISTE
   CNRG& = CantRegistros&("DEUDOR12", CONDI$)
   If CNRG& < 1 Then Call LIMPIARPRESU: Exit Sub
   '
   'CARGO DATOS DE CLIENTE
   Text2 = VALOBADA("DEUDOR12", "Raso_Cli", CONDI$, "NOMESS")
   Text3 = VALOBADA("DEUDOR12", "Domi_Cli", CONDI$, "NOMESS")
   Text4 = VALOBADA("DEUDOR12", "LOCA_CLI", CONDI$, "NOMESS")
   Text8 = VALOBADA("DEUDOR12", "Tele_Cli", CONDI$, "NOMESS")
   Text5 = VALOBADA("DEUDOR12", "CELU_CLI", CONDI$, "NOMESS")
   Text18 = VALOBADA("DEUDOR12", "Mail_Cli", CONDI$, "NOMESS")
   Text19 = VALOBADA("DEUDOR12", "Nota_Cli", CONDI$, "NOMESS")
   Text23 = CEROIGUALVACIO(XVALO(VALOBADA("DEUDOR12", "Lpre_Cli", CONDI$, "NOMESS")))
   '

End Sub

Private Sub Text20_Change()
  If XVALO(Text20) > 0 Then
     Text21.Visible = True
     'Call APLICAR_SKIN(Me, App.Path & "\SKINS\GOLDEN.skn")
  Else
     'Call APLICAR_SKIN(Me, App.Path & "\SKINS\dogmax.skn")
     Text21.Visible = False
  End If
End Sub


Private Sub Text21_DblClick()
   Text1 = ""
   Text17 = 1 + MAXNUMCOD&("PRESUVEN", "NroPres")
   'Text20 = "-00"

   Text1.Locked = False 'TEXT DE CUENTA DE CLIENTE
   Command19.Enabled = True 'BOTON DE LISTA DE SELECCIN
   If Text21.Visible = True Then Text21.Visible = False 'CARTEL DE MODIFICACION DE PRESUPUESTO
   MSF1.Rows = 1
   On Error Resume Next
   Unload INFOSIAHP
   On Error GoTo 0
   Call ACTU_TOTAL
   PORDESCU = ""
   
   
End Sub


Private Sub Text23_Change()
'   CODI% = XVALO(Text23)
'   Text22 = ECOLPRECIOS$(CODI%)

End Sub

Private Sub TICAMBIS_DblClick()
   If DERACCE%("COTILOCA", "Ajustar Tipo de Cambio Local") >= 2 Then
      ENTRATICA.Frame4.Caption = "Cotización Local para " + MONEMIS.Caption
      ENTRATICA.Text4 = TICAMBIS.Caption
      ENTRATICA.Text4 = FORMATNUM(XVALO(ENTRATICA.Text4), "F10.4")
      ENTRATICA.Show 1
      If TRIM$(ENTRATICA.Text4) <> "" Then
        TICAMBIS = TRIM$(ENTRATICA.Text4)
      End If
   End If
   '

End Sub


Private Sub TICAMBIS_Change()
   If TICAMBIS.Enabled = True Then
     TICACH(MONEMI%) = XVALO(TICAMBIS)
   End If
   TICAMBX# = 1
   If MONEMI% > 1 Then
     TICAMBX# = XVALO(TICAMBIS)
   End If
End Sub

Private Sub TICAMBIS_LostFocus()
   TICAMBIS = FORMATNUM(XVALO(TICAMBIS), "F9.4")
End Sub

Private Sub TXTCANTIDAD_Change()

   Resultado# = XVALO(TXTCANTIDAD) * XVALO(TXTPREUNIT)
   TXTIMPORTE = FORMATNUM$(Resultado#, "F14.2")
   '
End Sub

Private Sub TXTCANTIDAD_KeyPress(KeyAscii As Integer)
   
    If KeyAscii = 27 Then
        KeyAscii = 0
        Call LIMPIAR_TEXT_DETALLE
    End If
   
'    If KeyAscii = 13 Then TXTPREUNIT.SetFocus
   
' *********** CODIGO ANTERIOR ************

    If KeyAscii = 27 Then
        KeyAscii = 0
        Call LIMPIAR_TEXT_DETALLE
    End If


    If KeyAscii = 13 Then
        KeyAscii = 0
        '
        'VALIDO EL CODIGO
        CI1% = CODINT%(txtcodigo)
        If CI1% < 1 Then
           MsgBox "CÓDIGO INEXISTENTE O NO VÁLIDO"
           Exit Sub
        End If

        'VALIDO QUE TENGA INGRESADA CANTIDAD
        If XVALO(TXTCANTIDAD.TEXT) < 0.05 Then
           MsgBox "FALTA CANTIDAD"
           TXTCANTIDAD.SetFocus
           Exit Sub
        End If
        '
'        If CARGA(2) = True Then 'OPCION DE CARGA CON SELECCION MANUAL DE LOTES
'          Call PRESENTA_INFO_ARTICULO(CI1%)
'        End If
        '
'        'VALIDO QUE TENGA PRECIO
'        If XVALO(TXTPREUNIT.TEXT) < 0.005 Then
'           KeyAscii = 0
'           MsgBox "FALTA PRECIO UNITARIO"
'           TXTPREUNIT.SetFocus
'           Exit Sub
'        End If
        '
'        If VALIDA_COD_GRID(txtcodigo, FRMPRESU.MSF1, 1) > 0 Then
'         MsgBox "MODIFIQUE DATOS SOBRE EL MISMO"
'         GoTo 50
'        End If

        TXTPREUNIT.BackColor = &HE0E0E0 'CAMBIA A GRIS
        TXTPREUNIT.Locked = True 'DESHABILITO
        cargar_datos
        txtcodigo = ""
        txtcodigo.SetFocus
50      Call LIMPIAR_TEXT_DETALLE
        '
        'CIERRO EL FORMULARIO DE DETALLE DE STOCKS
        Unload INFOSIAHP
        '
        'PARA QUE VUELVA A PRESENTAR LA LISTA DE CODIGOS
        If Command1.Enabled = True Then
           Call FRMPRESUAHP.Command1_Click
        End If
    End If


    

End Sub
Private Sub cargar_datos()
    '
'    VER QUE NO TRAE LA CANTIDAD
'    MSF1.Cols = 13
    If INFOSIAHP.FILAX& > 0 Then
       MSF1.Row = INFOSIAHP.FILAX&
       REG& = INFOSIAHP.FILAX&
'       TXTCANTIDAD = XVALO(INFOSIAHP.Text55)
    Else
       MSF1.Rows = MSF1.Rows + 1
       REG& = MSF1.Rows - 1
    End If
    
    CI1% = CODINT%(txtcodigo.TEXT)
    '
    MSF1.TextMatrix(REG&, 1) = txtcodigo.TEXT
    MSF1.TextMatrix(REG&, 2) = INFOSIAHP.TXTCODICLIENTE.TEXT
    MSF1.TextMatrix(REG&, 3) = DESCRIT0$(CI1%)
    MSF1.TextMatrix(REG&, 4) = TRIM$(FORMATNUM$(TXTCANTIDAD.TEXT, "F9.1"))
    MSF1.TextMatrix(REG&, 5) = INFOSIAHP.TXTCOSTO
    MSF1.TextMatrix(REG&, 5) = INFOSIAHP.TXTNUMEPROVE
'    MSF1.TextMatrix(REG&, 6) = INFOSIAHP.TXTPROVEEDOR
    
    
'    If CARGA(2).Value = True Then 'CASO QUE ES CARGA DE SELECCION MANUAL DE LOTES TRAE LA INFO DE INFOSIAHP
'        MSF1.TextMatrix(REG&, 7) = FORMATNUM$(INFOSIAHP.TXTCOSTO.TEXT, "f12.2")
'        MSF1.TextMatrix(REG&, 8) = FORMATNUM$(INFOSIAHP.TXTPRECIOVENTA, "f12.2")
'        MSF1.TextMatrix(REG&, 9) = FORMATNUM$(XVALO(INFOSIAHP.TXTPRECIOVENTA) * XVALO(TXTCANTIDAD), "f14.2")
'        MSF1.TextMatrix(REG&, 11) = CARGALOTES$ ' FUNCION
'        MSF1.TextMatrix(REG&, 12) = CANTI_ACUMU_LOTES ' FUNCION
'        'Call CANTI_SEGUN_LOTES_SELECCIONADOS(CARGALOTES$) 'ACA HABRIA QUE DECODIFICAR LA CANTIDAD DE LOTES SELECCIONADOS
'        ' **** Y QUE HACER EL COLORSTATUS PARA VER QUE COLOR DEVUELVE, EN ESTE CASO NO SE DEBERIA CONTEMPLAR LA CANTIDAD EN TRANSITO.
'        'VALIDA EL COLOR
'        CANTILOTES = XVALO(MSF1.TextMatrix(REG&, 12))
'        CANTISOLICI = XVALO(MSF1.TextMatrix(REG&, 4))
'        If CANTILOTES <= CANTISOLICI Then COLO$ = "A"
'        If CANTILOTES = 0 Then COLO$ = "R"
'        If CANTILOTES >= CANTISOLICI Then COLO$ = "V"
'        MSF1.TextMatrix(REG&, 0) = COLO$
'        '
'    Else
'        MSF1.TextMatrix(REG&, 7) = FORMATNUM$(Text15.TEXT, "f12.2")
'        MSF1.TextMatrix(REG&, 8) = FORMATNUM$(TXTPREUNIT, "f12.2")
'        MSF1.TextMatrix(REG&, 9) = FORMATNUM$(XVALO(TXTPREUNIT) * XVALO(TXTCANTIDAD), "f14.2")
'        COLO$ = COLORSTATUS$(TXTCANTIDAD, INFOSIAHP.Text33(0), INFOSIAHP.Text33(1), INFOSIAHP.Text33(2))
'        MSF1.TextMatrix(REG&, 0) = COLO$
'    End If
'    MSF1.TextMatrix(REG&, 10) = Text38
    '***ESTE BLOQUE DE 5 LINEAS HABRIA QUE MOVERLO LUEGO DEL ELSE ****
    If COLO$ = "V" Then Call COLOREAR_GRILLA_SOLACELDA(FRMPRESUAHP.MSF1, vbGreen, REG&, 0)
    If COLO$ = "A" Then Call COLOREAR_GRILLA_SOLACELDA(FRMPRESUAHP.MSF1, vbYellow, REG&, 0)
    If COLO$ = "R" Then Call COLOREAR_GRILLA_SOLACELDA(FRMPRESUAHP.MSF1, vbRed, REG&, 0)
    
'    MSF1.TextMatrix(REG&, 6) = "0" ' INDICADOR DE CODIGO OPCIONAL
'    If Check1.Value = 1 Then
'        MSF1.TextMatrix(REG&, 6) = "1"
'        MSF1.TextMatrix(REG&, 5) = ""
'    End If
'    MSF1.ColWidth(6) = 0 ' OCULTA LA COLUMNA
'    Check1.Value = 0 ' SETEA NO OPCIONAL POR DEFAULT
    '
    ' ORDENA GRILLA POR COLUMNA 6 Y EN FORMA ASCENDENTE (VALOR 7 DE SORT DATOS DE TIPO STRING)
'    MSF1.Col = 6
'    MSF1.Sort = 7
    '
    'SUMO EL IMPORTE
    Call ACTU_TOTAL
    '
    On Error Resume Next
    VISIROWS = (MSF1.Height / MSF1.RowHeight(J&)) - 4
    If Err > 0 Then
      VISIROWS = 1
    End If
    On Error GoTo 0
    If VISIROWS > 0 Then
       If MSF1.TopRow < MSF1.Row - VISIROWS Then
          MSF1.TopRow = MSF1.Row - VISIROWS
       End If
    End If
    '
    'desplazo verticalmente a la ultima columna
    'MSF1.LeftCol = MSF1.Cols - 1
    
    INFOSIAHP.FILAX& = 0
End Sub
Sub ACTU_TOTAL()


   '
   Static CLIA$, RNI%, APLIVA%
   If Text1.TEXT <> CLIA$ Then
     CLIA$ = Text1.TEXT: NCLIE% = XVALO(CLIA$)
     RNI% = 0: APLIVA% = 1
     If NCLIE% > 0 Then
       PVCLI% = PIVACLI%(NCLIE%)
       If PVCLI% = 2 Then RNI% = 1
       If PVCLI% = 5 Or PVCLI% = 7 Then APLIVA% = 0
     End If
   End If
   '
   SUTOTPED# = 0: SUIVA# = 0: SUIVARNI# = 0: TOIVA# = 0: TOTOTA# = 0
   For U& = 1 To MSF1.Rows - 1
     CIXXI% = CODINT%(MSF1.TextMatrix(U&, 1))
     CAN = XVALO(MSF1.TextMatrix(U&, 4))
'     PORDESCUIN = CVS(Mid$(X$, 69, 4))
'     If PORDESCUEIN < 0 Or PORDESCUEIN > 99 Then PORDESCUEIN = 0
     If CAN > 0.05 Then
       PRELI = XVALO(MSF1.TextMatrix(U&, 8))
       'If PRELI < 0.01 Then
       '  PRELI = PRELISTA(ARCHIPRE$, CIXXI%)
       'End If
       PREUNI = PRELI * (100 - PORDESCUIN) / 100
       TOT# = CDbl(CAN * PREUNI)
       SUTOTPED# = SUTOTPED# + TOT#
       SUIVA# = SUIVA# + (Int(TOT# * PORCEIVA(CIXXI%) + 0.5)) / 100
'       Call REPLA(X$, MKS$(PREUNI) + MKD$(TOT#), 51, 12)
'       Call GRAREG("!CARDETPR", X$, U&)
     End If
   Next U&
   '
   Total(0).TEXT = FORMATNUM$(SUTOTPED#, "F12.2")
   Total(0).Refresh
   '
   DESC$ = TRIM$(PORDESCU.TEXT)
   PODES = COEFDESCU(DESC$)
   TTTX# = SUNEFA2# * (100 - PODES) / 100
3   TTT1 = XVALO(Total(0).TEXT) * PODES / 100
   Total(1).TEXT = CSTRING$(MKS$(TTT1), 3, 11, 2, 2)
   '
   TTT3 = XVALO(Total(0).TEXT) - XVALO(Total(1).TEXT)
   '
    ACUMU_IMPORTE# = 0
'    For I& = 1 To MSF1.Rows - 1
'        IMPORTE# = XVALO(MSF1.TextMatrix(I&, 5))
'        ACUMU_IMPORTE# = ACUMU_IMPORTE# + IMPORTE#
'    Next I&
    TTTX# = SUIVA# * (100 - PODES) / 100
    Text33 = FORMATNUM$(XVALO(Total(0)) - XVALO(Total(1)), "F12.2")
    Text31 = FORMATNUM$(TTTX#, "f12.2")
    Text14 = FORMATNUM$(XVALO(Text33) + XVALO(Text31), "F12.2")

End Sub

Private Sub txtcodigo_Change()

    TXTDESCRIPCION = ""
    TXTCANTIDAD = ""
    TXTPREUNIT = ""
    TXTIMPORTE = ""
    'DESHABILITO EL TEXT DE PRECIO UNITARIO
    TXTPREUNIT.BackColor = &HE0E0E0 'CAMBIA A GRIS
    TXTPREUNIT.Locked = True 'DESHABILITO
    'MSF3.Rows = 1
    TXTPREUNIT.ToolTipText = ""

    If TRIM$(txtcodigo) = "" Then Exit Sub
    'VALIDO SI LA LISTA ACTIVA ES VALIDA
'    NLIST% = XVALO(Text23)
'    If NLIST% < 1 Or TRIM$(Text22) = "" Then
'      NLIST% = 1
'    End If
    '
    'BUSCO LOS DATOS
    CI1% = CODINT%(txtcodigo)
    If CI1% < 1 Then
       Exit Sub
    Else
       TXTDESCRIPCION = DESCRIT0$(CI1%)
    End If
    '
    On Error Resume Next
    Call PINTATEXT(TXTCANTIDAD)
    TXTCANTIDAD.SetFocus
    On Error GoTo 0
    
'    CONDI$ = " CODINT = " & CI1%
  
    
    
End Sub



Private Sub txtcodigo_GotFocus()
'   If TRIM$(txtcodigo.TEXT) <> "" Then
'       Command1_Click
'   End If
End Sub


'Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
'   If KeyAscii = 13 Then
'     CI1% = CODINT(txtcodigo)
'     If CI1% < 1 Then
'        Call COMUNI("Código Inexistente o No válido")
'        Exit Sub
'     End If
'
'     If KeyAscii = 13 Then
'        TXTDESCRIPCION = DESCRIT0(CI1%)
'        TXTCANTIDAD = ""
'        'TXTPREUNIT = FORMATNUM$(PRELISTA(TRIM$(Str$(NLIST%)), CI1%), "F12.2")
'        TXTIMPORTE = ""
'        '
'        Call PRESENTA_INFO_ARTICULO(CI1%)
'        '
'        ' SI SE SELECCIONO DE LA LISTA PASA EL FOCO A LA CANTIDAD
'
'        If Command1.Enabled = False Then
'           TXTCANTIDAD.SetFocus
'        End If
'     End If
'   End If
'End Sub

Private Sub TXTPREUNIT_Change()

   Resultado# = XVALO(TXTCANTIDAD) * XVALO(TXTPREUNIT)
   TXTIMPORTE = FORMATNUM$(Resultado#, "F14.2")

End Sub
Sub LIMPIAR_TEXT_DETALLE()
    TXTDESCRIPCION = ""
    TXTCANTIDAD = ""
    TXTPREUNIT = ""
    TXTIMPORTE = ""
    txtcodigo = ""
    'DESHABILITO EL TEXT DE PRECIO UNITARIO
    TXTPREUNIT.Locked = False
    TXTPREUNIT.ToolTipText = ""
    Text35 = ""
    Text15 = ""
    Text38 = ""
    
End Sub

Sub TXTPREUNIT_DblClick()

    If TXTPREUNIT.Locked = False Then
       TXTPREUNIT.BackColor = &HE0E0E0
       TXTPREUNIT.Locked = True
    Else
       TXTPREUNIT.Locked = False
       TXTPREUNIT.BackColor = &H80000005
       Call PINTATEXT(TXTPREUNIT)
    End If
    
End Sub

Private Sub TXTPREUNIT_KeyPress(KeyAscii As Integer)


    If TXTPREUNIT.Locked = True Then Exit Sub
    '
    
    If KeyAscii = 27 Then
        KeyAscii = 0
        Call LIMPIAR_TEXT_DETALLE
    End If
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If XVALO(TXTCANTIDAD.TEXT) < 0.05 Then
           KeyAscii = 0
           MsgBox "FALTA CANTIDAD"
           TXTCANTIDAD.SetFocus
           Exit Sub
        End If
        
        If XVALO(TXTPREUNIT.TEXT) < 0.05 Then
           KeyAscii = 0
           MsgBox "FALTA PRECIO UNITARIO"
           TXTPREUNIT.SetFocus
           Exit Sub
        End If

'        If VALIDA_COD_GRID(txtcodigo, FRMPRESU.MSF1, 1) > 0 Then
'         MsgBox "MODIFIQUE DATOS SOBRE EL MISMO"
'         GoTo 50
'        End If
        TXTPREUNIT.BackColor = &HE0E0E0 'CAMBIA A GRIS
        TXTPREUNIT.Locked = True 'DESHABILITO

        cargar_datos
        txtcodigo = ""
        txtcodigo.SetFocus
50      Call LIMPIAR_TEXT_DETALLE
        If Command1.Enabled = True Then
           Call FRMPRESUAHP.Command1_Click
        End If

     End If
    
End Sub
   
Private Sub TXTPREUNIT_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode 'ESTO ES PARA QUE DESP
        Case 115 'pulso  f4
        Call TXTPREUNIT_DblClick

    End Select

End Sub
Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
     CI1% = CODINT(txtcodigo)
     If CI1% < 1 Then
        CICLI& = CODICLIENTE&(txtcodigo)
        If CICLI& > 0 Then
          CI1% = CICLI&
        Else
          OPX% = COMALTER("Código Inexistente o No Válido\\Dar de Alta\Cancelar")
          If OPX% = 1 Then
            'ALTA DE CODIGO Y/O RELACION
            Call ALTACODIRELA
          End If
          Exit Sub
        End If
     End If
     
     TXTDESCRIPCION = DESCRIT0(CI1%)
     TXTCANTIDAD = ""
     'TXTPREUNIT = FORMATNUM$(PRELISTA(TRIM$(Str$(NLIST%)), CI1%), "F12.2")
     TXTIMPORTE = ""
     '
     '
     ' SI SE SELECCIONO DE LA LISTA PASA EL FOCO A LA CANTIDAD
     If Command1.Enabled = False Then
       TXTCANTIDAD.SetFocus
     End If
   End If
End Sub
'
Sub ALTACODIRELA()
    VDEF$ = String$(101, 0)
    CODI$ = TRIM$(txtcodigo)
    Call REPLA(VDEF$, CODI$, 77, 24)
    Call REPLA(VDEF$, MKI$(1 + NUMAS%), 1, 2)
10  OBX$ = OBJPLA$("RELACODI", VDEF$)
    If OBX$ = "" Then Exit Sub
    
    Call ALTARELACODIGO(OBX$, OK%)
    If OK% = 1 Then
      Call COMUNI("Alta Realizada Correctamente")
    ElseIf OK% = 2 Then
      Call COMUNI("Código Relacionado Correctamente")
    Else
      VDEF$ = OBX$
      GoTo 10
    End If
    '
    UPDATMASTER% = 1
    Call GENECOMAS
   
    
End Sub

VERSION 5.00
Begin VB.Form FAPECOBRA 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Facturas Pendientes de Cobranza "
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9630
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command10 
      Caption         =   "Deuda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8610
      Picture         =   "FAPECOBRA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Composición de Deuda a Fecha"
      Top             =   2310
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   8610
      Sorted          =   -1  'True
      TabIndex        =   39
      Top             =   3465
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      DownPicture     =   "FAPECOBRA.frx":014A
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Left            =   8610
      Picture         =   "FAPECOBRA.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Re-calcula y Refresca Opciones"
      Top             =   5300
      Width           =   855
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
      Height          =   810
      Left            =   4410
      TabIndex        =   37
      Top             =   5185
      Width           =   4005
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   315
         ScaleHeight     =   75
         ScaleWidth      =   3315
         TabIndex        =   40
         Top             =   420
         Width           =   3375
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   41
            Top             =   0
            Width           =   12000
         End
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Rango de Cuentas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   210
      TabIndex        =   28
      Top             =   3780
      Width           =   4005
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2625
         MaxLength       =   5
         TabIndex        =   34
         Top             =   420
         Width           =   855
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
         Height          =   360
         Left            =   3465
         TabIndex        =   33
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text9 
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
         Left            =   315
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   3330
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2625
         MaxLength       =   5
         TabIndex        =   31
         Text            =   " "
         Top             =   1260
         Width           =   855
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
         Height          =   360
         Left            =   3465
         TabIndex        =   30
         Top             =   1260
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
         Left            =   315
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1680
         Width           =   3330
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
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
         TabIndex        =   36
         Top             =   525
         Width           =   2325
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta :"
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
         TabIndex        =   35
         Top             =   1365
         Width           =   2325
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   4410
      TabIndex        =   25
      Top             =   4200
      Width           =   4005
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Pantalla"
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
         Index           =   4
         Left            =   525
         TabIndex        =   27
         Top             =   420
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Impresión"
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
         Index           =   8
         Left            =   2310
         TabIndex        =   26
         Top             =   420
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Formato de Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4410
      TabIndex        =   21
      Top             =   2310
      Width           =   4005
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Listado Tabulado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   420
         TabIndex        =   24
         Top             =   420
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Plan Financiero de Cobranza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   420
         TabIndex        =   23
         Top             =   840
         Width           =   3270
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Gráfico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   420
         TabIndex        =   22
         Top             =   1260
         Width           =   1380
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      Left            =   210
      TabIndex        =   8
      Top             =   315
      Width           =   4005
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Rango de Fechas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   17
         Top             =   2205
         Width           =   3750
         Begin VB.CommandButton Command11 
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
            Height          =   360
            Left            =   1650
            TabIndex        =   45
            Top             =   340
            Width           =   175
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
            Height          =   360
            Left            =   700
            TabIndex        =   43
            Top             =   340
            Width           =   1000
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
            Height          =   360
            Left            =   3465
            TabIndex        =   19
            Top             =   340
            Width           =   175
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2460
            TabIndex        =   18
            Top             =   340
            Width           =   1000
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta"
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
            Left            =   1770
            TabIndex        =   44
            Top             =   480
            Width           =   645
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde"
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
            Left            =   0
            TabIndex        =   20
            Top             =   480
            Width           =   645
         End
      End
      Begin VB.TextBox Text4 
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
         Left            =   315
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Todos - Sin Clasificación"
         Top             =   1680
         Width           =   3330
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
         Height          =   360
         Left            =   3465
         TabIndex        =   14
         Top             =   1260
         Width           =   175
      End
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
         Height          =   360
         Left            =   3045
         TabIndex        =   13
         Text            =   " "
         Top             =   1260
         Width           =   435
      End
      Begin VB.TextBox Text2 
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
         Left            =   315
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   " Todos - Sin Clasificación"
         Top             =   840
         Width           =   3330
      End
      Begin VB.CommandButton Command3 
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
         Height          =   360
         Left            =   3465
         TabIndex        =   10
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3045
         TabIndex        =   9
         Text            =   " "
         Top             =   420
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
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
         Left            =   630
         TabIndex        =   16
         Top             =   1365
         Width           =   2325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo de Cliente:"
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
         Left            =   630
         TabIndex        =   12
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ordenamiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   4410
      TabIndex        =   3
      Top             =   315
      Width           =   4005
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Por Fecha de Vencimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   420
         TabIndex        =   7
         Top             =   1365
         Width           =   2745
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Alfabético"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   2205
         TabIndex        =   6
         Top             =   840
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Numérico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   735
         TabIndex        =   5
         Top             =   840
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Por Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   420
         TabIndex        =   4
         Top             =   420
         Width           =   1380
      End
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
      Height          =   855
      Left            =   8610
      Picture         =   "FAPECOBRA.frx":03DE
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1365
      Width           =   855
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
      Height          =   700
      Left            =   8610
      Picture         =   "FAPECOBRA.frx":06E8
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Emite Reporte Solicitado"
      Top             =   4305
      Width           =   855
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
      Height          =   855
      Left            =   8610
      Picture         =   "FAPECOBRA.frx":09F2
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Abandona la Aplicación"
      Top             =   420
      Width           =   855
   End
End
Attribute VB_Name = "FAPECOBRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim INIRECON%
'
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   Command2.Enabled = False
Call COMPODEUNUE
   Call COMPODEU
   Command2.Enabled = True
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text5.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub PLANFINA()
  '\\\OT-06-0174-WAR-12/06/06///
  Dim PLAFIN#(32767, 5) 'MATRIZ QUE CONTIENE EL PLAN FINANCIERO
  Dim FECOSE(5) As Integer 'FECHA COMIENZO DE SEMANA
  '
  FLIM1% = HOY(HO$)
  NUMSEM% = SEMANU%(FLIM1%, DELX, ALX)
  FECOSE%(1) = DELX 'FECHA COMIENZO SEMANA UNO
  FECOSE%(2) = DIPOST(ALX) 'FECHA COMIENZO SEMANA DOS
  NUSEM% = SEMANU%(FECOSE%(2), DELX, ALX)
  FECOSE%(3) = DIPOST(ALX) 'FECHA COMIENZO SEMANA TRES
  NUSEM% = SEMANU%(FECOSE%(3), DELX, ALX)
  FECOSE%(4) = DIPOST(ALX) 'FECHA COMIENZO SEMANA CUATRO
  NUSEM% = SEMANU%(FECOSE%(4), DELX, ALX)
  FECOSE%(5) = DIPOST(ALX) 'FECHA COMIENZO SEMANA CUATRO
  FECOSE%(0) = (-32767)
  '
'1.- GENERAR FACTURAS PENDIENTES DE COBRANZA (FABECOBR)
  'Call FAPECOB2
'2.- DISTRIBUIR PENDIENTES EN LA MATRIZ
  FIN& = ULTREG&("FAPECOBR")
  NCMAX% = 0
  For JX& = 1 To FIN&
    '
    Call TRACE(24, JX&, FIN&)
    '
    X$ = REGLEIDO$("FAPECOBR", JX&)
    '
    FV1% = CVI(Mid$(X$, 69, 2))
    NPRO1% = CVI(Mid$(X$, 1, 2))
    IMPEN1# = CVD(Mid$(X$, 71, 8))
    If NPRO1% > 0 Then
      If NPRO1% > NCMAX% Then NCMAX% = NPRO1%
    End If
    '
    For KKI% = 0 To 4
      If FV1% >= FECOSE%(KKI%) Then  'VENCIDAS ANTERIORES
        If FV1% < FECOSE%(KKI% + 1) Then
          PLAFIN#(NPRO1%, KKI%) = PLAFIN#(NPRO1%, KKI%) + IMPEN1#
          GoTo 19
        End If
      End If
    Next KKI%
    PLAFIN#(NPRO1%, 5) = PLAFIN#(NPRO1%, 5) + IMPEN1#
19 Next JX&
   '
'3.- GENERAR ARCHIVO  PLAN FINANCIERO A PARTIR DE LA MATRIZ
    '
  REBLAZ$ = REGBLAN$("PLANFIN")
  JJ& = 0
  FIN& = NCMAX%
  For JX1% = 1 To FIN&
    '
    REGACT1& = JX1%
    Call TRACE(24, REGACT1&, FIN&)
    '
    TOTAPRO1# = 0
    For XB% = 0 To 5
     TOTAPRO1# = TOTAPRO1# + PLAFIN#(JX1%, XB%)
    Next XB%
    '
    If Abs(TOTAPRO1#) > 0.005 Then 'SI TIENE PENDIENTES
      Y$ = REBLAZ$
      Call REPLA(Y$, MKI$(JX1%), 1, 2)
      Call REPLA(Y$, MKD$(PLAFIN#(JX1%, 0)), 3, 8) 'VEN
      Call REPLA(Y$, MKD$(PLAFIN#(JX1%, 1)), 11, 8) '1SEM
      Call REPLA(Y$, MKD$(PLAFIN#(JX1%, 2)), 19, 8) '2SEM
      Call REPLA(Y$, MKD$(PLAFIN#(JX1%, 3)), 27, 8) '3SEM
      Call REPLA(Y$, MKD$(PLAFIN#(JX1%, 4)), 35, 8) '4SEM
      Call REPLA(Y$, MKD$(PLAFIN#(JX1%, 5)), 43, 8) 'POST
      Call REPLA(Y$, MKD$(TOTAPRO1#), 51, 8)        'TOTAL
      JJ& = JJ& + 1
      Call GRAREG("PLANFIN", Y$, JJ&)
    End If
    '
  Next JX1%
  Call SETULTREG("PLANFIN", JJ&)
'5.- LISTAR
  Call FINAL("*PLANFIN")
  '\\\OT-06-0174-WAR-FIN///
End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    Command10.Enabled = False
    '
    NCMIN% = Val(Text10.TEXT)
    NCMAX% = Val(Text8.TEXT)
    FINF = FECHANUM(Text5.TEXT)
    FLIM = FECHANUM(TRIM$(Text6.TEXT)) ': If FLIM < HOY(HOS$) Then FLIM = HOY(HOS$)
    '
    If NCMIN% < 1 Then
       Call MENSERR(24, "Número de Cuenta no Válido")
       Text10.SetFocus
       GoTo 99
    End If
    '
    If NCMAX% < 1 Then
       Call MENSERR(24, "Número de Cuenta no Válido")
       Text8.SetFocus
       GoTo 99
    End If
    '
    If ECOARCH$("DEUDOR1", MKI$(NCMIN%)) = "" Then
       Call MENSERR(24, "Cuenta Inexistente")
       Text10.SetFocus
       GoTo 99
    End If
    '
    If ECOARCH$("DEUDOR1", MKI$(NCMAX%)) = "" Then
       Call MENSERR(24, "Cuenta Inexistente")
       Text8.SetFocus
       GoTo 99
    End If
    '
    If FLIM < FINF Then
      Call MENSERR(24, "Rango de Fechas No Procesable")
      GoTo 99
    End If
    '
    Call FAPECOB2
    AMU$ = "": EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DOSIVA") > 0 Then AMU$ = "D"
    If Option1(4).Value = True Then  ' pantalla
         If Option1(3).Value = False Then
            Call FINAL("*COFAPE1" + AMU$)
            GoTo 99
          Else
            Call FINAL("*COFAPE2" + AMU$)
            GoTo 99
         End If
       Else                          ' impresora
         '\\\OT-06-00174-WAR-12/06/06///
         If Option1(6).Value = True Then
           Call PLANFINA
           GoTo 99
         End If
         '\\\OT-06-00174-WAR-FIN///
         If Option1(3).Value = False Then
            Call FINAL("*LIFAPE1" + AMU$)
            GoTo 99
          Else
            Call FINAL("*LIFAPE2" + AMU$)
            GoTo 99
         End If
    End If
    '
99  Command10.Enabled = True
    Command2.Enabled = True
    '
End Sub

Private Sub Command2_GotFocus()
    '
    If CARFAPEN% = 0 Then
        Refresh
        Screen.MousePointer = 11
        Call FAPECOB1
        CARFAPEN% = 1
    End If
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("GRUCOCLI")
    Text1.TEXT = VALACT1$("GRUCOCLI")
    If TRIM$(Text1.TEXT) = "" Then
        Text2.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Command4_Click()
    Call SELECHO("VENDEDOR")
    Text3.TEXT = VALACT1$("VENDEDOR")
    If TRIM$(Text3.TEXT) = "" Then
        Text4.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Command5_Click()
    Call FAPECOB1
End Sub

Private Sub Command6_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Command8_Click()
    Call SELECHO("DEUDOR1")
    NCMAX% = Val(VALACT1$("DEUDOR1"))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = RASOCLI$(NCMAX%)
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("DEUDOR1")
    NCMIN% = Val(VALACT1$("DEUDOR1"))
    If NCMIN% > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = RASOCLI$(NCMIN%)
    End If
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        Option1(0) = False
        Option1(1) = True
    End If
End Sub

Private Sub Text1_DBLCLICK()
    Call SELECHO("GRUCOCLI")
    Text1.TEXT = VALACT1$("GRUCOCLI")
    If TRIM$(Text1.TEXT) = "" Then
        Text2.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Text10_Change()
    If Abs(Val(Text10.TEXT)) < 32767 Then
        NCMIN% = Val(Text10.TEXT)
      Else
        Text10.TEXT = ""
    End If
End Sub

Private Sub Text10_LostFocus()
    NCMIN% = Val(Text10.TEXT)
    If NCMIN% < 1 Then
       Text9.TEXT = ""
       Exit Sub
    End If
    If RASOCLI$(NCMIN%) = "" Then
       Text9.TEXT = ""
       Exit Sub
    End If
    Text9.TEXT = RASOCLI$(NCMIN%)
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text3_DblClick()
    Call SELECHO("VENDEDOR")
    Text3.TEXT = VALACT1$("VENDEDOR")
    If TRIM$(Text3.TEXT) = "" Then
        Text4.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Text1_LostFocus()
    If Val(Text1.TEXT) < 256 Then
        GRUCLI% = Val(Text1.TEXT)
        GRUX$ = ECOARCH$("GRUCOCLI", MKI$(GRUCLI%))
        If GRUX$ <> "" Then
            Text2.TEXT = GRUX$
            Exit Sub
        End If
    End If
    Text1.TEXT = ""
    Text2.TEXT = "Todos - Sin Clasificación"
End Sub

Private Sub Text3_LostFocus()
    If Val(Text3.TEXT) < 256 Then
        VENDE% = Val(Text3.TEXT)
        VENDX$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        If VENDX <> "" Then
            Text4.TEXT = VENDX$
            Exit Sub
        End If
    End If
    Text3.TEXT = ""
    Text4.TEXT = "Todos - Sin Clasificación"
End Sub

Private Sub Text4_CLICK()
    Text3.SetFocus
End Sub

Private Sub Text1_Change()
    If Val(Text1.TEXT) < 256 Then
        GRUCLI% = Val(Text1.TEXT)
        GRUX$ = ECOARCH$("GRUCOCLI", MKI$(GRUCLI%))
        If GRUX$ <> "" Then
            Text2.TEXT = GRUX$
            Exit Sub
        End If
    End If
End Sub

Private Sub Text3_Change()
    If Val(Text3.TEXT) < 256 Then
        VENDE% = Val(Text3.TEXT)
        VENDX$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        If VENDX <> "" Then
            Text4.TEXT = VENDX$
            Exit Sub
        End If
    End If
End Sub

Private Sub Text5_DblClick()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text5.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text8_Change()
    If Abs(Val(Text8.TEXT)) < 32767 Then
        NCMAX% = Val(Text8.TEXT)
      Else
        Text8.TEXT = ""
    End If
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("DEUDOR1")
    NCMIN% = Val(VALACT1$("DEUDOR1"))
    If NCMIN% > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = RASOCLI$(NCMIN%)
    End If
End Sub

Private Sub Text6_DblClick()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text7_CLICK()
    Text8.SetFocus
End Sub

Private Sub Text8_DblClick()
    Call SELECHO("DEUDOR1")
    NCMAX% = Val(VALACT1$("DEUDOR1"))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = RASOCLI$(NCMAX%)
    End If
End Sub

Private Sub Text8_LostFocus()
    NCMAX% = Val(Text8.TEXT)
    If NCMAX% < 1 Then
       Text7.TEXT = ""
       Exit Sub
    End If
    If RASOCLI$(NCMAX%) = "" Then
       Text7.TEXT = ""
       Exit Sub
    End If
    Text7.TEXT = RASOCLI$(NCMAX%)
End Sub

Private Sub Text9_Click()
    Text10.SetFocus
End Sub

Sub COMPODEU()
    '
    
    Dim NROPO&(8196), CLIPO%(8196)
    Dim SARRACLI#(32767), SUDEUCLI#(32767)
    Dim TIS$(10)
    TDAT$ = "AS,VC,FC,ND,RB,DC,NC,AS"
    For I% = 0 To 7
        TIS$(I%) = Mid$(TDAT$, 3 * I% + 1, 2)
    Next I%
    '
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) < 1 Then
       Exit Sub
    End If
    '
    FLIM% = FECHANUM(VALACT1$("SELFECHA"))
    '
    FIN& = ULTREG&("FACTUR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("FACTUR", U&)
      NCLIE% = CVI(Left$(X$, 2))
      FEDOC% = CVI(Mid$(X$, 7, 2))
      If FEDOC% <= FLIM% Then
        TI% = Asc(Mid$(X$, 13, 1))
        If TI% < 0 Or TI% > 7 Then TI% = 2
        SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
        IMPO# = CVD(Mid$(X$, 15, 8))
        IMPO# = IMPO# + CVD(Mid$(X$, 23, 8))
        IMPO# = IMPO# + CVD(Mid$(X$, 31, 8))
        IMPO# = IMPO# + CVD(Mid$(X$, 49, 8))
        IMPO# = IMPO# + CVD(Mid$(X$, 57, 8))
        IMPO# = SG% * IMPO#
        SARRACLI#(NCLIE%) = SARRACLI#(NCLIE%) + IMPO#
        SARRATO# = SARRATO# + IMPO#
      End If
    Next U&
    '
    FIN& = ULTREG&("COBRAN")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("COBRAN", U&)
      NCLIE% = CVI(Left$(X$, 2))
      FEDOC% = CVI(Mid$(X$, 7, 2))
      If FEDOC% <= FLIM% Then
      IMPO# = (-1) * CVD(Mid$(X$, 15, 8))
        SARRACLI#(NCLIE%) = SARRACLI#(NCLIE%) + IMPO#
        SARRATO# = SARRATO# + IMPO#
      End If
    Next U&
    '
    ' LEER SVD202 - DESCARTAR DOCUMENTOS POSTERIORES A FLIM%
    UJ& = 0: SUDEUTO# = 0
    FIN& = ULTREG&("SVD202")
    For UI& = 1 To FIN&
       Call TRACE(20, UI&, FIN&)
       X$ = REGLEIDO$("SVD202", UI&)
       FF% = CVI(Mid$(X$, 19, 2))
       If FF% <= FLIM% Then
         NC% = CVI(Left$(X$, 2))
         If NC% > 0 Then
           If Abs(SARRACLI#(NC%)) >= 0.005 Then
             TI% = Asc(Mid$(X$, 3, 1))
             If TI% < 0 Or TI% > 10 Then TI% = 2
             TIX$ = TIS$(TI%): SIG% = 1
             If TI% > 3 Then SIG% = (-1)
             '
             NF& = CVS(Mid$(X$, 5, 4))
             While NF& > 999999
                NF& = NF& - 100000
             Wend
             NFS$ = TRIM$(Str$(NF&))
             While Len(NFS$) < 8
                NFS$ = " " + NFS$
             Wend
             NFS$ = TIX$ + NFS$
             '
15           FF% = CVI(Mid$(X$, 19, 2))
             FVI% = CVI(Mid$(X$, 9, 2))
             SALDEU# = SIG% * Abs(CVD(Mid$(X$, 11, 8)))
             If TI% = 0 Or TI% = 7 Then
                If Sgn(CVD(Mid$(X$, 21, 8))) < 0 Then
                   SALDEU# = (-1) * Abs(CVD(Mid$(X$, 11, 8)))
                End If
             End If
             If Abs(SALDEU#) < 0.005 Then GoTo 99
             '
             If FF% = 0 Then
                Screen.MousePointer = 1
                Call MENSERR(24, "Imposible Continuar Proceso de Cálculo.\Fecha de Factura no Válida en Cuenta '" + TRIM$(Str$(NC%)) + "'.\Revise la Composición de Deuda de Esta Cuenta.")
                Hide
                Exit Sub
             End If
             '
             IMORIG# = SIG% * CVD(Mid$(X$, 21, 8))
             ES$ = "    "
             If SALDEU# < 0 Then
                   ES$ = " CR "
                 Else
                   'If FVI% <= HOO% + 8 Then ES$ = " ** "
                   'If FVI% <= HOO% Then ES$ = "!!!!"
                   If FVI% < HOO% Then
                     FVIL& = FVI%: HOOL& = HOO%
                     ES$ = FORMATNUM$(DIENTRE%(FVIL&, HOOL&), "F4.0")
                     If Left$(ES$, 1) = " " Then Mid$(ES$, 1, 1) = "."
                   End If
             End If
             '
             Z$ = REGBLAN$("FAPECOB")
             Call REPLA(Z$, MKI$(NC%), 1, 2)
             Call REPLA(Z$, NFS$, 3, 10)
             Call REPLA(Z$, MKI$(FF%), 13, 2)
             Call REPLA(Z$, MKI$(FVI%), 15, 2)
             Call REPLA(Z$, MKD$(SALDEU#), 17, 8)
             Call REPLA(Z$, MKD$(IMORIG#), 25, 8)
             '
             UJ& = UJ& + 1
             Call GRAREG("FAPECOB", Z$, UJ&)
             SUDEUCLI#(NC%) = SUDEUCLI#(NC%) + SALDEU#
             SUDEUTO# = SUDEUTO# + SALDEU#
             '
           End If
         End If
       End If
       '
99   Next UI&
     '
     Call SETULTREG("FAPECOB", UJ&)
     Call CIERRARCH("FAPECOB")
     '
     Screen.MousePointer = 1
     '
    ' LEER FACTURAS - CARGAR VECTOR CON MOVIMIENTOS POSTERIORES A FLIM%
    CAMOPO% = 0
    FIN& = ULTREG&("FACTUR")
    For UI& = 1 To FIN&
      X$ = REGLEIDO$("FACTUR", UI&)
      TIPO% = Asc(Mid$(X$, 13, 1))
      If TIPO% > 0 Then
        If TIPO% < 4 Then
          FFAC% = CVI(Mid$(X$, 7, 2))
          If FFAC% > FLIM% Then
            CAMOPO% = CAMOPO% + 1
            If CAMOPO% > 8196 Then
              Call MENSERR(24, "Imposible Re-Calcular Composición de Deuda\en la Fecha Elegida.")
              Call FINAL("")
            End If
            NROPO&(CAMOPO%) = CVS(Mid$(X$, 3, 4))
            CLIPO%(CAMOPO%) = CVI(Left$(X$, 2))
          End If
        End If
      End If
    Next UI&
    '
    ' LEER APLICOB - DESCARTAR MOVIMIENTOS ANTERIORES A FLIM%
    '              - DESCARTAR MOVIMIENTOS QUE FIGURAN EN MOVIMIENTOS POSTERIORES
    '
    FIN& = ULTREG&("APLICOB")
    For UI& = 1 To FIN&
      X$ = REGLEIDO$("APLICOB", UI&)
      FAPLI% = CVI(Left$(X$, 2))
      If FAPLI% > FLIM% Then
        CLIAPLI% = CVI(Mid$(X$, 3, 2))
        If CLIAPLI% > 0 Then
          If Abs(SARRACLI#(CLIAPLI%) - SUDEUCLI#(CLIAPLI%)) >= 0.005 Then
            DOCUORI$ = Mid$(X$, 55, 18)
            If Mid$(DOCUORI$, 3, 1) <> "-" Or Mid$(DOCUORI$, 8, 1) <> "-" Or Mid$(DOCUORI$, 17, 1) <> "-" Then
   AAA = BBB
            End If
            NUDOAP& = Val(Mid$(DOCUORI$, 9, 8))
            For KKI% = 1 To CAMOPO%
              If NROPO&(KKI%) = NUDOAP& Then
                If CLIPO%(KKI%) = CLIAPLI% Then GoTo 29
              End If
            Next KKI%
            '
            Z$ = REGBLAN$("FAPECOB")
            Call REPLA(Z$, MKI$(CLIAPLI%), 1, 2)
             NF& = Val(Mid$(DOCUORI$, 10, 7))
             While NF& > 999999: NF& = NF& - 100000: Wend
             NFS$ = TRIM$(Str$(NF&))
             While Len(NFS$) < 8: NFS$ = " " + NFS$: Wend
             NFS$ = Left$(DOCUORI$, 2) + NFS$
            Call REPLA(Z$, NFS$, 3, 10)
            TDCC$ = UCase$(Left$(NFS$, 2))
            TIXA% = 2: If TDCC$ = "ND" Then TIXA% = 3
                       If TDCC$ = "NC" Then TIXA% = 6
                       If TDCC$ = "AJ" Then TIXA% = 0
            FF% = FLIM%
            RFFA$ = RECFILT$("FACTUR", 1, MKI$(CLIAPLI%))
            For KKLL& = 1 To Len(RFFA$) Step 4
              RELEL& = CVS(Mid$(RFFA$, KKLL&, 4))
              ZFF$ = REGLEIDO$("FACTUR", RELEL&)
              If CVS(Mid$(ZFF$, 3, 4)) = NF& Then
                If Asc(Mid$(ZFF$, 13, 1)) = TIXA% Then
                  FF% = CVI(Mid$(ZFF$, 7, 2))
                  GoTo 24
                End If
              End If
            Next KKLL&
            '
24          Call REPLA(Z$, MKI$(FF%), 13, 2)
            FVI% = FAPLI%
            Call REPLA(Z$, MKI$(FVI%), 15, 2)
            SALDEU# = CVD(Mid$(X$, 23, 8))
            Call REPLA(Z$, MKD$(SALDEU#), 17, 8)
            IMORIG# = CVD(Mid$(X$, 73, 8))
            Call REPLA(Z$, MKD$(IMORIG#), 25, 8)
            '
            UJ& = UJ& + 1
            Call GRAREG("FAPECOB", Z$, UJ&)
            SUDEUCLI#(CLIAPLI%) = SUDEUCLI#(CLIAPLI%) + SALDEU#
            SUDEUTO# = SUDEUTO# + SALDEU#
            '
          End If
        End If
      End If
29  Next UI&
    '
    FIN& = ULTREG&("APLICOB")
    For UI& = 1 To FIN&
      X$ = REGLEIDO$("APLICOB", UI&)
      FAPLI% = CVI(Left$(X$, 2))
      If FAPLI% <= FLIM% Then
        CLIAPLI% = CVI(Mid$(X$, 3, 2))
        DIFESAL# = SARRACLI#(CLIAPLI%) - SUDEUCLI#(CLIAPLI%)
        If Abs(DIFESAL#) >= 0.005 Then
          DOCUORI$ = Mid$(X$, 55, 18)
          If Mid$(DOCUORI$, 3, 1) <> "-" Or Mid$(DOCUORI$, 8, 1) <> "-" Or Mid$(DOCUORI$, 17, 1) <> "-" Then
AAA = BBB
          End If
          NUDOAP& = Val(Mid$(DOCUORI$, 9, 8))
          DOCUAP$ = Mid$(X$, 5, 3) + Mid$(X$, 16, 7)
          'For KKI% = 1 To CAMOPO%
          '  If NROPO&(KKI%) = NUDOAP& Then
          '    If CLIPO%(KKI%) = CLIAPLI% Then GoTo 39
          '  End If
          'Next KKI%
          '
          Z$ = REGBLAN$("FAPECOB")
          Call REPLA(Z$, MKI$(CLIAPLI%), 1, 2)
             NF& = Val(Mid$(DOCUAP$, 4, 7))
             While NF& > 999999: NF& = NF& - 100000: Wend
             NFS$ = TRIM$(Str$(NF&))
             While Len(NFS$) < 8: NFS$ = " " + NFS$: Wend
             NFS$ = Left$(DOCUAP$, 2) + NFS$
          Call REPLA(Z$, NFS$, 3, 10)
          FF% = FLIM%
          Call REPLA(Z$, MKI$(FF%), 13, 2)
          FVI% = FAPLI%
          Call REPLA(Z$, MKI$(FVI%), 15, 2)
          SALDEU# = CVD(Mid$(X$, 23, 8))
          Call REPLA(Z$, MKD$((-1) * SALDEU#), 17, 8)
          IMORIG# = CVD(Mid$(X$, 73, 8))
          Call REPLA(Z$, MKD$(IMORIG#), 25, 8)
          '
          UJ& = UJ& + 1
          Call GRAREG("FAPECOB", Z$, UJ&)
          SUDEUCLI#(CLIAPLI%) = SUDEUCLI#(CLIAPLI%) - SALDEU#
          SUDEUTO# = SUDEUTO# - SALDEU#
          '
        End If
      End If
39  Next UI&
    '
    FIN& = 32767
    For UI& = 1 To FIN&
      CLIAPLI% = UI&
      DIFESAL# = SARRACLI#(CLIAPLI%) - SUDEUCLI#(CLIAPLI%)
      If Abs(DIFESAL#) >= 0.005 Then
          '
          Z$ = REGBLAN$("FAPECOB")
          Call REPLA(Z$, MKI$(CLIAPLI%), 1, 2)
          NFS$ = "Dif.Saldo"
          Call REPLA(Z$, NFS$, 3, 10)
          FF% = FLIM%
          Call REPLA(Z$, MKI$(FF%), 13, 2)
          FVI% = FLIM%
          Call REPLA(Z$, MKI$(FVI%), 15, 2)
          SALDEU# = DIFESAL#
          Call REPLA(Z$, MKD$(SALDEU#), 17, 8)
          IMORIG# = CVD(Mid$(X$, 73, 8))
          Call REPLA(Z$, MKD$(DIFESAL#), 25, 8)
          '
          UJ& = UJ& + 1
          Call GRAREG("FAPECOB", Z$, UJ&)
          SUDEUCLI#(CLIAPLI%) = SUDEUCLI#(CLIAPLI%) + SALDEU#
          '
      End If
49  Next UI&
    '
    Call SETULTREG("FAPECOB", UJ&)
    Call CIERRARCH("FAPECOB")
    '              - AGREGAR A LISTA DE SVD202
    ' ORDENAR POR NUMERO O ALFABETICO Y POR FECHA EMISION
    '
    GRUCLI% = Val(Text1.TEXT)
    GRUX$ = TRIM$(Text2.TEXT)
    VENDE% = Val(Text3.TEXT)
    VENDX$ = TRIM$(Text4.TEXT)
    NCMIN% = Val(Text10.TEXT)
    NCMAX% = Val(Text8.TEXT)
    '
    MODORDE% = 1
    If Option1(2) = True Then MODORDE% = 2
    If Option1(3) = True Then MODORDE% = 3
    '
    FIN& = ULTREG&("FAPECOB")
    UJ& = 0
    List1.Clear
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("FAPECOB", U&)
        NC% = CVI(Left$(X$, 2))
        '
        If NC% >= NCMIN% Then
          If NC% <= NCMAX% Then
            FVI% = CVI(Mid$(X$, 15, 2))
            If GRUCLI% = 0 Or GRUCOCLI%(NC%) = GRUCLI% Then
              If VENDE% = 0 Or VENDCLI%(NC%) = VENDE% Then
                '
                TLISTA$ = Space$(36)
                If MODORDE% = 1 Then
                     Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 1, 6)
                  ElseIf MODORDE% = 2 Then
                     Call REPLA(TLISTA$, RASOCLI$(NC%), 1, 30)
                  ElseIf MODORDE% = 3 Then
                     Call REPLA(TLISTA$, CSTRING$(MKI$(FVI%), 1, 6, 0, 2), 1, 6)
                End If
                Call REPLA(TLISTA$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 31, 6)
                List1.AddItem TLISTA$
                '
              End If
            End If
          End If
        End If
        '
    Next U&
    '
    FIN& = List1.ListCount
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        List1.ListIndex = U& - 1
        REG& = Val(Mid$(List1.TEXT, 31, 8))
        X$ = REGLEIDO$("FAPECOB", REG&)
        '
        NC% = CVI(Left$(X$, 2))
        NFS$ = Mid$(X$, 3, 10)
        FF% = CVI(Mid$(X$, 13, 2))
        FVI% = CVI(Mid$(X$, 15, 2))
        SALDEU# = CVD(Mid$(X$, 17, 8))
        IMORIG# = CVD(Mid$(X$, 25, 8))
        ES$ = ""   ': If FVI% < HOY(HOS$) Then ES$ = "!!!!"
        If FVI% < HOY(HOS$) Then
           FVIL& = FVI%: HOOL& = HOY(HOS$)
           ES$ = FORMATNUM$(DIENTRE%(FVIL&, HOOL&), "F4.0")
           If Left$(ES$, 1) = " " Then Mid$(ES$, 1, 1) = "."
        End If        '
        Y$ = REGBLAN$("FAPECOBR")
        Call REPLA(Y$, MKI$(NC%), 1, 2)
        Call REPLA(Y$, RASOCLI$(NC%), 3, 30)
        Call REPLA(Y$, TEL1CLI$(NC%), 33, 16)
        Call REPLA(Y$, NFS$, 49, 10)
        Call REPLA(Y$, MKI$(FF%), 59, 2)
        Call REPLA(Y$, MKD$(IMORIG#), 61, 8)
        Call REPLA(Y$, MKI$(FVI%), 69, 2)
        Call REPLA(Y$, MKD$(SALDEU#), 71, 8)
        Call REPLA(Y$, ES$, 87, 10)
        '
        UJ& = UJ& + 1
        Call GRAREG("FAPECOBR", Y$, UJ&)
        '          '
    Next U&
    '
    Call SETULTREG("FAPECOBR", UJ&)
    Call CIERRARCH("FAPECOBR")
    '
    FEX = FLIM%
    Call HEADERS("FAPECOBR", "")
    Call HEADERS("FAPECOBR", "Composición de Deuda al: " + FECHATEX$(FEX))
    '
    Call FINAL("*LIFAPE1")
    '
    Screen.MousePointer = 1

End Sub

Sub FAPECOB1()
    '
10  Screen.MousePointer = 11
    OKEY% = 0
    '
    NCMIN% = 32767: NCMAX% = 0
    FLIM = HOY(HO$)
    FIN& = ULTREG&("SVD202")
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("SVD202", U&)
      NCX% = CVI(Left$(X$, 2))
      If NCX% > 0 Then
        If NCX% > NCMAX% Then NCMAX% = NCX%
        If NCX% < NCMIN% Then NCMIN% = NCX%
        '
        FEVE = CVI(Mid$(X$, 9, 2))
        If FEVE > FLIM Then FLIM = FEVE
      End If
    Next U&
    '
    Text10.TEXT = TRIM$(Str$(NCMIN%))
    Text9.TEXT = RASOCLI$(NCMIN%)
    Text8.TEXT = TRIM$(Str$(NCMAX%))
    Text7.TEXT = RASOCLI$(NCMAX%)
    Text6.TEXT = FECHATEX$(FLIM)
    Text5.TEXT = "01/01/80"
    '
    Dim TIS$(10)
    TDAT$ = "AS,VC,FC,ND,RB,DC,NC,AS"
    For I% = 0 To 7
        TIS$(I%) = Mid$(TDAT$, 3 * I% + 1, 2)
    Next I%
    '
    UJ& = 0
    For UI& = 1 To FIN&
       Call TRACE(20, UI&, FIN&)
       X$ = REGLEIDO$("SVD202", UI&)
       NC% = CVI(Left$(X$, 2))
       If NC% > 0 Then
           TI% = Asc(Mid$(X$, 3, 1))
           If TI% < 0 Or TI% > 10 Then TI% = 2
           TIX$ = TIS$(TI%): SIG% = 1
           If TI% > 3 Then SIG% = (-1)
           '
           NF = CVS(Mid$(X$, 5, 4))
           While NF > 999999
              NF = NF - 100000
           Wend
           NFS$ = TRIM$(Str$(NF))
           While Len(NFS$) < 8
              NFS$ = " " + NFS$
           Wend
           NFS$ = TIX$ + NFS$
           '
15         FF% = CVI(Mid$(X$, 19, 2))
           FVI% = CVI(Mid$(X$, 9, 2))
           SALDEU# = SIG% * Abs(CVD(Mid$(X$, 11, 8)))
           If TI% = 0 Or TI% = 7 Then
              If Sgn(CVD(Mid$(X$, 21, 8))) < 0 Then
                 SALDEU# = (-1) * Abs(CVD(Mid$(X$, 11, 8)))
              End If
           End If
           If Abs(SALDEU#) < 0.005 Then GoTo 99
           '
           If FF% = 0 Then
              Screen.MousePointer = 1
              Call MENSERR(24, "Imposible Continuar Proceso de Cálculo.\Fecha de Factura no Válida en cuenta '" + TRIM$(Str$(NC%)) + "'.\Revise la Composición de Deuda de Esta Cuenta.")
              Hide
              Exit Sub
           End If
           '
           IMORIG# = SIG% * CVD(Mid$(X$, 21, 8))
           ES$ = "    "
           MONEMX% = Asc(Mid$(X$, 31, 1))
           '
           If SALDEU# < 0 Then
                 ES$ = " CR "
               Else
                 'If FVI% <= HOO% + 8 Then ES$ = " ** "
                 'If FVI% <= HOO% Then ES$ = "!!!!"
                 If FVI% < HOO% Then
                   FVIL& = FVI%: HOOL& = HOO%
                   ES$ = FORMATNUM$(DIENTRE%(FVIL&, HOOL&), "F4.0")
                   If Left$(ES$, 1) = " " Then Mid$(ES$, 1, 1) = "."
                 End If
           End If
           '
           Z$ = REGBLAN$("FAPECOB")
           Call REPLA(Z$, MKI$(NC%), 1, 2)
           Call REPLA(Z$, NFS$, 3, 10)
           Call REPLA(Z$, MKI$(FF%), 13, 2)
           Call REPLA(Z$, MKI$(FVI%), 15, 2)
           Call REPLA(Z$, MKD$(SALDEU#), 17, 8)
           Call REPLA(Z$, MKD$(IMORIG#), 25, 8)
           Call REPLA(Z$, Chr$(MONEMX%), 33, 1)
           Call REPLA(Z$, Chr$(TI%), 34, 1)
           '
           UJ& = UJ& + 1
           Call GRAREG("FAPECOB", Z$, UJ&)
           '
       End If
       '
99   Next UI&
     '
     Call SETULTREG("FAPECOB", UJ&)
     Screen.MousePointer = 1
     '
End Sub
'

Sub FAPECOB2()
    '
    Screen.MousePointer = 11
    GRUCLI% = Val(Text1.TEXT)
    GRUX$ = TRIM$(Text2.TEXT)
    VENDE% = Val(Text3.TEXT)
    VENDX$ = TRIM$(Text4.TEXT)
    FINF = FECHANUM(Text5.TEXT)
    FLIM = FECHANUM(TRIM$(Text6.TEXT)) ': If FLIM < HOY(HOS$) Then FLIM = HOY(HOS$)
    NCMIN% = Val(Text10.TEXT)
    NCMAX% = Val(Text8.TEXT)
    '
    MODORDE% = 1
    If Option1(2) = True Then MODORDE% = 2
    If Option1(3) = True Then MODORDE% = 3
    '
    FIN& = ULTREG&("FAPECOB")
    UJ& = 0
    List1.Clear
    Dim SALDEPE#(32767)
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("FAPECOB", U&)
        NC% = CVI(Left$(X$, 2))
        SALDEU# = CVD(Mid$(X$, 17, 8))
        If NC% >= 0 Then
          SALDEPE#(NC%) = SALDEPE#(NC%) + SALDEU#
        End If
        '
       If NC% >= NCMIN% Then
          If NC% <= NCMAX% Then
            FVI% = CVI(Mid$(X$, 15, 2))
            If FVI% >= FINF And FVI% <= FLIM Then
              If GRUCLI% = 0 Or GRUCOCLI%(NC%) = GRUCLI% Then
                If VENDE% = 0 Or VENDCLI%(NC%) = VENDE% Then
                  '
                  TLISTA$ = Space$(36)
                  If MODORDE% = 1 Then
                       Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 1, 6)
                    ElseIf MODORDE% = 2 Then
                       Call REPLA(TLISTA$, RASOCLI$(NC%), 1, 30)
                    ElseIf MODORDE% = 3 Then
                       Call REPLA(TLISTA$, CSTRING$(MKI$(FVI%), 1, 6, 0, 2), 1, 6)
                  End If
                  Call REPLA(TLISTA$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 31, 6)
                  List1.AddItem TLISTA$
                  '
                End If
              End If
            End If
          End If
        End If
        '
    Next U&
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "SABO") > 0 Then
      ATXX% = 0
      FIN& = ULTREG&("DEUDOR1")
      For UI& = 1 To FIN&
        Call TRACE(20, UI&, FIN&)
        XD1$ = REGLEIDO$("DEUDOR1", UI&)
        NCI% = CVI(Left$(XD1$, 2))
        If NCI% > 0 Then
          If Abs(SALDEPE#(NCI%) - SALDCLI#(NCI%)) >= 0.005 Then
' AKX1$ = RASOCLI$(NCI%)
' ABC = SALDCLI#(NCI%)
            RESCUDEU.Text1 = ""
            RESCUDEU.Text1 = TRIM$(Str$(NCI%))
            If Abs(Val(RESCUDEU.Label5) - SALDCLI#(NCI%)) > 0.005 Then
              RESCUDEU.Show 1
              ATXX% = 1
            End If
          End If
        End If
      Next UI&
      If ATXX% > 0 Then
        '
        Screen.MousePointer = 11
        On Error Resume Next
        HOII% = HOY(HOS$)
        DRESPA$ = "C:\FLEXOFT\RP" + Left$(HOS$, 2) + Mid$(HOS$, 4, 2) + Mid$(HOS$, 7, 2)
        On Error Resume Next
        MkDir DRESPA$
        ABAK$ = "SVD202": FRESPA$ = LOADFILE$(LUDAT$ + ABAK$ + ".DAT"): Call SAVEFILE(DRESPA$ + "\" + ABAK$ + ".DAT", FRESPA$)
        ABAK$ = "FAPECOB": FRESPA$ = LOADFILE$(LUDAT$ + ABAK$ + ".DAT"): Call SAVEFILE(DRESPA$ + "\" + ABAK$ + ".DAT", FRESPA$)
        ABAK$ = "FACTUR": FRESPA$ = LOADFILE$(LUDAT$ + ABAK$ + ".DAT"): Call SAVEFILE(DRESPA$ + "\" + ABAK$ + ".DAT", FRESPA$)
        ABAK$ = "COBRAN": FRESPA$ = LOADFILE$(LUDAT$ + ABAK$ + ".DAT"): Call SAVEFILE(DRESPA$ + "\" + ABAK$ + ".DAT", FRESPA$)
        ABAK$ = "DEUDOR1": FRESPA$ = LOADFILE$(LUDAT$ + ABAK$ + ".DAT"): Call SAVEFILE(DRESPA$ + "\" + ABAK$ + ".DAT", FRESPA$)
        ABAK$ = "DEUDOR2": FRESPA$ = LOADFILE$(LUDAT$ + ABAK$ + ".DAT"): Call SAVEFILE(DRESPA$ + "\" + ABAK$ + ".DAT", FRESPA$)
        On Error GoTo 0
        Screen.MousePointer = 1
        '
        Call MENSERR(24, "Imposible Emitir Listado de Facturas\Pendientes de Cobranza en Contexto Actual.\  \Repita su Petición")
        Call FINAL("")
      End If
    End If
    '
    FIN& = List1.ListCount
    '
    UNIFICA% = 0
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        List1.ListIndex = U& - 1
        REG& = Val(Mid$(List1.TEXT, 31, 8))
        X$ = REGLEIDO$("FAPECOB", REG&)
        '
        NC% = CVI(Left$(X$, 2))
        NFS$ = Mid$(X$, 3, 10)
        NROL& = Val(Mid$(NFS$, 4, 7))
        FF% = CVI(Mid$(X$, 13, 2))
        FVI% = CVI(Mid$(X$, 15, 2))
        SALDEU# = CVD(Mid$(X$, 17, 8))
        IMORIG# = CVD(Mid$(X$, 25, 8))
If NC% = 3614 Then
AAA = BBB
End If
        ES$ = ""   ': If FVI% < HOY(HOS$) Then ES$ = "!!!!"
        If FVI% < HOY(HOS$) Then
           FVIL& = FVI%: HOOL& = HOY(HOS$)
           ES$ = FORMATNUM$(DIENTRE%(FVIL&, HOOL&), "F4.0")
           If Left$(ES$, 1) = " " Then Mid$(ES$, 1, 1) = "."
        End If
        MONEMX% = Asc(Mid$(X$, 33, 1))
        TIA% = Asc(Mid$(X$, 34, 1))
        '
        Y$ = REGBLAN$("FAPECOBD")
        Z$ = REGBLAN$("FAPECOBR")
        '
        Call REPLA(Y$, MKI$(NC%), 1, 2)
        Call REPLA(Y$, RASOCLI$(NC%), 3, 30)
        Call REPLA(Y$, TEL1CLI$(NC%), 33, 16)
        Call REPLA(Y$, NFS$, 49, 10)
        Call REPLA(Y$, MKI$(FF%), 59, 2)
        Call REPLA(Y$, MKD$(IMORIG#), 65, 8)
        Call REPLA(Y$, MKI$(FVI%), 73, 2)
        Call REPLA(Y$, MKD$(SALDEU#), 75, 8)
        '
        Call REPLA(Z$, Left$(Y$, 60), 1, 60)
        Call REPLA(Z$, MKD$(IMORIG#), 61, 8)
        Call REPLA(Z$, MKI$(FVI%), 69, 2)
        Call REPLA(Z$, MKD$(SALDEU#), 71, 8)
        '
        If MONEMX% = 2 Then
          '
          If UNIFICA% = 0 Then
            UNIFICA% = (-1)
            If COMALTER%("Puede Mantener Separados\los Saldos en Pesos y Dolares\o Unificarlos en Pesos al T.C. " + TRIM$(FORMATNUM$(TICAMONE(MONEMX%), "F10.4")) + "\\Mantener Separados\Unificar en Pesos") = 2 Then
              UNIFICA% = 1
            End If
          End If
          '
          If Abs(Abs(IMORIG#) - Abs(SALDEU#)) < 0.005 Then
            RFG$ = RECFILT$("CARBIMON", 2, MKI$(NC%))
            For KUL& = Len(RFG$) - 3 To 1 Step -4
              RECABIN& = CVS(Mid$(RFG$, KUL&, 4))
              XXC$ = REGLEIDO$("CARBIMON", RECABIN&)
              If CVS(Mid$(XXC$, 39, 4)) = NROL& Then
                If Asc(Mid$(XXC$, 37, 1)) = TIA% Then
                  SALDOPES# = CVD(Mid$(XXC$, 65, 8)) - CVD(Mid$(XXC$, 73, 8))
                  If Abs(SALDOPES# - SALDEU#) > 0.005 Then
                      TICAMOR# = CVD(Mid$(XXC$, 89, 8))
                      If TICAMOR# < 1 Then TICAMOR# = 1
                      SALDOL# = SALDEU# / TICAMOR#
                    Else
                      SALDOL# = CVD(Mid$(XXC$, 97, 8)) - CVD(Mid$(XXC$, 105, 8))
'AA2 = CVD(Mid$(XXC$, 73, 8))
                  End If
                  Call REPLA(Y$, MKD$(0), 75, 8)
                  If UNIFICA% = 1 Then Call REPLA(Y$, MKD$(SALDOL# * TICAMONE(MONEMX%)), 75, 8)
                  Call REPLA(Y$, MKD$(SALDOL#), 91, 8)
                  GoTo 20
                End If
              End If
            Next KUL&
          End If
        End If
        '
20      If MONEMX% = 2 Then
           Call REPLA(Y$, "U$S", 61, 4)
        End If
        Call REPLA(Y$, ES$, 107, 10)
        Call REPLA(Z$, ES$, 87, 10)
        '
        UJ& = UJ& + 1
        Call GRAREG("FAPECOBR", Z$, UJ&)
        Call GRAREG("FAPECOBD", Y$, UJ&)
        '          '
    Next U&
    '
    Call SETULTREG("FAPECOBR", UJ&)
    Call CIERRARCH("FAPECOBR")
    Call SETULTREG("FAPECOBD", UJ&)
    Call CIERRARCH("FAPECOBD")
    '
    Call HEADERS("FAPECOBR", "")
    If GRUCLI% > 0 Then Call HEADERS("FAPECOBR", "Grupo: " + GRUX$)
    If VENDE% > 0 Then Call HEADERS("FAPECOBR", "Vendedor: " + VENDX$)
    Call HEADERS("FAPECOBR", "Vencimiento: " + FECHATEX$(FINF) + " - " + FECHATEX$(FLIM))
    Call HEADERS("FAPECOBR", "Rango de Cuentas: " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)))
    '
    Call HEADERS("FAPECOBD", "")
    If GRUCLI% > 0 Then Call HEADERS("FAPECOBD", "Grupo: " + GRUX$)
    If VENDE% > 0 Then Call HEADERS("FAPECOBD", "Vendedor: " + VENDX$)
    Call HEADERS("FAPECOBD", "Vencimiento: " + FECHATEX$(FINF) + " - " + FECHATEX$(FLIM) + "   -   R.Cuentas: " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)))
    'Call HEADERS("FAPECOBD", "Rango de Cuentas: " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)))
    If UNIFICA% = 1 Then Call HEADERS("FAPECOBD", "Observaciones: Saldos en Dolares Unificados a Pesos al T.C. " + TRIM$(FORMATNUM$(TICAMONE(2), "F10.4")))
    '
    Screen.MousePointer = 1
    '
End Sub
'

Sub COMPODEUNUE()
    '
    Dim NROPO&(8196), CLIPO%(8196)
    Dim SARRACLI#(32767), SUDEUCLI#(32767)
    Dim TIS$(10)
    TDAT$ = "AS,VC,FC,ND,RB,DC,NC,AS"
    For I% = 0 To 7
        TIS$(I%) = Mid$(TDAT$, 3 * I% + 1, 2)
    Next I%
    '
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) < 1 Then
       Exit Sub
    End If
    '
    FLIM% = FECHANUM(VALACT1$("SELFECHA"))
    '
    NCLIMAX& = 0
    FIN& = ULTREG&("FACTUR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("FACTUR", U&)
      NCLIE% = CVI(Left$(X$, 2))
      FEDOC% = CVI(Mid$(X$, 7, 2))
      If FEDOC% <= FLIM% Then
        TI% = Asc(Mid$(X$, 13, 1))
        If TI% < 0 Or TI% > 7 Then TI% = 2
        SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
        IMPO# = CVD(Mid$(X$, 15, 8))
'If NCLIE% = 32767 Then
'AAA = BBB
'End If
        IMPO# = IMPO# + CVD(Mid$(X$, 23, 8))
        IMPO# = IMPO# + CVD(Mid$(X$, 31, 8))
        IMPO# = IMPO# + CVD(Mid$(X$, 49, 8))
        IMPO# = IMPO# + CVD(Mid$(X$, 57, 8))
        IMPO# = SG% * IMPO#
        SARRACLI#(NCLIE%) = SARRACLI#(NCLIE%) + IMPO#
        SARRATO# = SARRATO# + IMPO#
        '
        If Abs(IMPO#) >= 0.005 Then
          If NCLIE% > NCLIMAX& Then NCLIMAX& = NCLIE%
        End If
        '
      End If
    Next U&
    '
    FIN& = ULTREG&("COBRAN")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("COBRAN", U&)
      NCLIE% = CVI(Left$(X$, 2))
      FEDOC% = CVI(Mid$(X$, 7, 2))
      If FEDOC% <= FLIM% Then
        IMPO# = (-1) * CVD(Mid$(X$, 15, 8))
        SARRACLI#(NCLIE%) = SARRACLI#(NCLIE%) + IMPO#
        SARRATO# = SARRATO# + IMPO#
        '
        If Abs(IMPO#) >= 0.005 Then
          If NCLIE% > NCLIMAX& Then NCLIMAX& = NCLIE%
        End If
        '
      End If
    Next U&
    '
    If NCLIMAX& > 32766 Then NCLIMAX& = 32766 ' PARA EVITAR OVERFLOW
    INIRECON% = 0
    For NCLIE% = 0 To NCLIMAX&
      NCX& = NCLIE%
      Call TRACE(20, NCX&, NCLIMAX&)
      If Abs(SARRACLI#(NCLIE%)) >= 0.005 Then
        SUMARRAS# = SARRACLI#(NCLIE%)
        Call RECONDEU(NCLIE%, FLIM%, SUMARRAS#)
      End If
    Next NCLIE%
    '
    INIRECON% = 2   ' PARA PONER MARCA DE FIN DE ARCHIVO
    Call RECONDEU(NCLIE%, FLIM%, SUMARRAS#)
    
    Call CIERRARCH("FAPECOB")
    '
    '              - AGREGAR A LISTA DE SVD202
    ' ORDENAR POR NUMERO O ALFABETICO Y POR FECHA EMISION
    '
    GRUCLI% = Val(Text1.TEXT)
    GRUX$ = TRIM$(Text2.TEXT)
    VENDE% = Val(Text3.TEXT)
    VENDX$ = TRIM$(Text4.TEXT)
    NCMIN% = Val(Text10.TEXT)
    NCMAX% = Val(Text8.TEXT)
    '
    MODORDE% = 1
    If Option1(2) = True Then MODORDE% = 2
    If Option1(3) = True Then MODORDE% = 3
    '
    FIN& = ULTREG&("FAPECOB")
    UJ& = 0
    List1.Clear
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("FAPECOB", U&)
        NC% = CVI(Left$(X$, 2))
        '
        If NC% >= NCMIN% Then
          If NC% <= NCMAX% Then
            FVI% = CVI(Mid$(X$, 15, 2))
            If GRUCLI% = 0 Or GRUCOCLI%(NC%) = GRUCLI% Then
              If VENDE% = 0 Or VENDCLI%(NC%) = VENDE% Then
                '
                TLISTA$ = Space$(36)
                If MODORDE% = 1 Then
                     Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 1, 6)
                  ElseIf MODORDE% = 2 Then
                     Call REPLA(TLISTA$, RASOCLI$(NC%), 1, 30)
                  ElseIf MODORDE% = 3 Then
                     Call REPLA(TLISTA$, CSTRING$(MKI$(FVI%), 1, 6, 0, 2), 1, 6)
                End If
                Call REPLA(TLISTA$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 31, 6)
                List1.AddItem TLISTA$
                '
              End If
            End If
          End If
        End If
        '
    Next U&
    '
    FIN& = List1.ListCount
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        List1.ListIndex = U& - 1
        REG& = Val(Mid$(List1.TEXT, 31, 8))
        X$ = REGLEIDO$("FAPECOB", REG&)
        '
        NC% = CVI(Left$(X$, 2))
        NFS$ = Mid$(X$, 3, 10)
        FF% = CVI(Mid$(X$, 13, 2))
        FVI% = CVI(Mid$(X$, 15, 2))
        SALDEU# = CVD(Mid$(X$, 17, 8))
        IMORIG# = CVD(Mid$(X$, 25, 8))
        ES$ = ""   ': If FVI% < HOY(HOS$) Then ES$ = "!!!!"
        If FVI% < HOY(HOS$) Then
           FVIL& = FVI%: HOOL& = HOY(HOS$)
           ES$ = FORMATNUM$(DIENTRE%(FVIL&, HOOL&), "F4.0")
           If Left$(ES$, 1) = " " Then Mid$(ES$, 1, 1) = "."
        End If        '
        Y$ = REGBLAN$("FAPECOBR")
        Call REPLA(Y$, MKI$(NC%), 1, 2)
        Call REPLA(Y$, RASOCLI$(NC%), 3, 30)
        Call REPLA(Y$, TEL1CLI$(NC%), 33, 16)
        Call REPLA(Y$, NFS$, 49, 10)
        Call REPLA(Y$, MKI$(FF%), 59, 2)
        Call REPLA(Y$, MKD$(IMORIG#), 61, 8)
        Call REPLA(Y$, MKI$(FVI%), 69, 2)
        Call REPLA(Y$, MKD$(SALDEU#), 71, 8)
        Call REPLA(Y$, ES$, 87, 10)
        '
        UJ& = UJ& + 1
        Call GRAREG("FAPECOBR", Y$, UJ&)
        '          '
    Next U&
    '
    Call SETULTREG("FAPECOBR", UJ&)
    Call CIERRARCH("FAPECOBR")
    '
    FEX = FLIM%
    HOII% = HOY(HOS$)
    Call HEADERS("FAPECOBR", "")
    Call HEADERS("FAPECOBR", "Composición de Deuda al: " + FECHATEX$(FEX) + " (Reconstruida al " + HOS$ + ")")
    '
    Call FINAL("*LIFAPE1")
    '
    Screen.MousePointer = 1

End Sub

Sub RECONDEU(NC%, FLIM%, SUMARRAS#)
    '
    Static UJ&
    Dim DOCU$(8192), IMPO#(8192), REGDO&(8192), FFDOC%(8192), IORI#(8192)
    '
    If INIRECON% = 0 Then
      INIRECON% = 1
      UJ& = 0
    End If
    '
    If INIRECON% = 2 Then
      Call SETULTREG("FAPECOB", UJ&)
      Exit Sub
    End If
    '
    ATENTI% = 1
    '
100 Screen.MousePointer = 11
    '
    SUMADEUP# = SUMARRAS#
    '
    MAXDIAC& = ULTREG&("COBRAN")
    MAXE& = ULTREG&("FACTUR")
    '
    RFF$ = RECFILT$("FACTUR", 1, MKI$(NC%))
    RFC$ = RECFILT$("COBRAN", 1, MKI$(NC%))
    '
    NUDOCU% = 0
    FIN& = Len(RFF$)
    For U& = 1 To Len(RFF$) Step 4
      'Call TRACE(20, U&, FIN&)
      REG& = CVS(Mid$(RFF$, U&, 4))
      X$ = REGLEIDO$("FACTUR", REG&)
      If CVI(Mid$(X$, 7, 2)) <= FLIM% Then
110     NUDO& = CVS(Mid$(X$, 3, 4))
        TI% = Asc(Mid$(X$, 13, 1))
        If NUDO& < 1 Then
           If TI% = 0 Or TI% = 7 Then
              Call REPLA(X$, MKS$(REG&), 3, 4)
              Call GRAREG("FACTUR", X$, REG&)
              GoTo 110
           End If
        End If
        VA% = Asc(Mid$(X$, 14, 1))
        SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
        NUDOCU% = NUDOCU% + 1
        If NUDOCU% > 8192 Then
           Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
           Call CIERRARCH("*.*")
           Call FINAL("")
        End If
115     IMPO#(NUDOCU%) = CVD(Mid$(X$, 15, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 23, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 31, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 49, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 57, 8))
        IMPO#(NUDOCU%) = SG% * (Int(100 * IMPO#(NUDOCU%) + 0.5)) / 100
        NUSUCUVE% = CVI(Mid$(X$, 47, 2))
        DOCU$(NUDOCU%) = Left$(FORMDOC$(NUDO&, TI%, VA%, NUSUCUVE%), 2) + "-" + TRIM$(Str$(NUDO&))
        REGDO&(NUDOCU%) = REG&
        FFDOC%(NUDOCU%) = CVI(Mid$(X$, 7, 2))
        '
        If TI% > 0 Then
          If TI% < 7 Then
            If CVD(Mid$(X$, 15, 8)) < (-0.005) Then
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 15, 8)))), 15, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 23, 8)))), 23, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 31, 8)))), 31, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 49, 8)))), 49, 8)
              Call REPLA(X$, MKD$(Abs(CVD(Mid$(X$, 57, 8)))), 57, 8)
              Call GRAREG("FACTUR", X$, REG&)
              GoTo 115
            End If
          End If
        End If
      End If
    Next U&
    '
    FIN& = Len(RFC$)
    For U& = 1 To Len(RFC$) Step 4
      'Call TRACE(20, U&, FIN&)
      REG& = CVS(Mid$(RFC$, U&, 4))
      X$ = REGLEIDO$("COBRAN", REG&)
      If CVI(Mid$(X$, 7, 2)) <= FLIM% Then
        NUDO& = CVS(Mid$(X$, 3, 4))
        TI% = 4
        VA% = 1
        SG% = (-1)
        NUDOCU% = NUDOCU% + 1
        If NUDOCU% > 8192 Then
           Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
           Call CIERRARCH("*.*")
           Call FINAL("")
        End If
        IMPO#(NUDOCU%) = CVD(Mid$(X$, 15, 8))
        IMPO#(NUDOCU%) = SG% * (Int(100 * IMPO#(NUDOCU%) + 0.5)) / 100
        DOCU$(NUDOCU%) = "RB-" + TRIM$(Str$(NUDO&))
        REGDO&(NUDOCU%) = (-1) * REG&
        FFDOC%(NUDOCU%) = CVI(Mid$(X$, 7, 2))
      End If
    Next U&
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       'Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
    SUMAPEN# = 0
    For KKI% = 1 To NUDOCU%
       SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
       If Abs(SUMAPEN#) < 0.005 Then
          For KKJ% = 1 To KKI%
             IMPO#(KKJ%) = 0
          Next KKJ%
       End If
    Next KKI%
    '
    If NUDOCU% >= 1 Then
      If Left$(DOCU$(1), 2) = "AS" Then
        FIN& = ULTREG&("SALINID")
        For U& = 1 To FIN&
          X$ = REGLEIDO$("SALINID", U&)
             If CVI(Left$(X$, 2)) = NC% Then
               TI% = Asc(Mid$(X$, 3, 1))
               VA% = Asc(Mid$(X$, 4, 1))
               NUSUCUVE% = CVI(Mid$(X$, 29, 2))
               SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
               IPEN# = SG% * CVD(Mid$(X$, 11, 8))
               If Sgn(IPEN#) = Sgn(IMPO#(1)) Then
                 If Abs(IPEN#) <= Abs(IMPO#(1)) Then
                   IMPO#(1) = IMPO#(1) - IPEN#
                   NUDO& = CVS(Mid$(X$, 5, 4))
                   NUDOCU% = NUDOCU% + 1
                   If NUDOCU% > 8192 Then
                      Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
                      Call CIERRARCH("*.*")
                      Call FINAL("")
                   End If
                   IMPO#(NUDOCU%) = IPEN#
                   DOCU$(NUDOCU%) = Left$(FORMDOC$(NUDO&, TI%, VA%, NUSUCUVE%), 2) + "-" + TRIM$(Str$(NUDO&))
                   REGDO&(NUDOCU%) = 900000 + U&
                   FFDOC%(NUDOCU%) = CVI(Mid$(X$, 19, 2))
                 End If
               End If
             End If
          Next U&
       End If
    End If
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       'Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
' HASTA AQUI O.K
    KKI% = 0: SUMAPEX# = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          SUMAPEX# = SUMAPEX# + IMPO#(KKI%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    FIN& = ULTREG&("APLICOB")
    For U& = 1 To FIN&
       'Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("APLICOB", U&)
       If CVI(Mid$(X$, 3, 2)) = NC% Then
          RECI$ = Mid$(X$, 5, 18)
          RECI$ = Left$(RECI$, 2) + "-" + TRIM$(Val(Mid$(RECI$, 11)))
          FACU$ = Mid$(X$, 55, 18)
          PINI% = 11: If Mid$(FACU$, 17, 1) = "-" Then PINI% = 9
          FACU$ = Left$(FACU$, 2) + "-" + TRIM$(Val(Mid$(FACU$, PINI%, 8)))
          '
210       IAPLI# = CVD(Mid$(X$, 23, 8))
          IRECIMA# = 0: IFACMAX# = 0
          For KKI% = 1 To NUDOCU%
             If DOCU$(KKI%) = RECI$ Or (RECI$ = "AS-0" And Left$(DOCU$(KKI%), 2) = "AS") Then
               If IMPO#(KKI%) < 0 Then
                 IRECIMA# = IRECIMA# - IMPO#(KKI%)
               End If
             End If
             If DOCU$(KKI%) = FACU$ Then
               If IMPO#(KKI%) > 0 Then
                 IFACMAX# = IFACMAX# + IMPO#(KKI%)
               End If
             End If
          Next KKI%
          '
          If IAPLI# > IRECIMA# Then IAPLI# = IRECIMA#
          If IAPLI# > IFACMAX# Then IAPLI# = IFACMAX#
          '
          IRECIMA# = IAPLI#
          IFACMAX# = IAPLI#
          '
          If IAPLI# > 0 Then
             For KKI% = 1 To NUDOCU%
                If DOCU$(KKI%) = RECI$ Or (RECI$ = "AS-0" And Left$(DOCU$(KKI%), 2) = "AS") Then
                   If IMPO#(KKI%) < 0 Then
                      IAP1# = IRECIMA#
                      If IAP1# > Abs(IMPO#(KKI%)) Then IAP1# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) + IAP1#
                      IRECIMA# = IRECIMA# - IAP1#
                   End If
                End If
                If DOCU$(KKI%) = FACU$ Then
                   If IMPO#(KKI%) > 0 Then
                      IAP2# = IFACMAX#
                      If IAP2# > Abs(IMPO#(KKI%)) Then IAP2# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) - IAP2#
                      IFACMAX# = IFACMAX# - IAP2#
                   End If
                End If
             Next KKI%
          End If
          '
       End If
    Next U&
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       'Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Or (FFDOC%(KKI%) = FFDOC%(KKJ%) And Sgn(REGDO&(KKI%)) = Sgn(REGDO&(KKJ%)) > 0 And Abs(REGDO&(KKI%)) > Abs(REGDO&(KKJ%))) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
    SUMAPEN# = 0
    For KKI% = 1 To NUDOCU%
       SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
       If Abs(SUMAPEN#) < 0.005 Then
          For KKJ% = 1 To KKI%
             IMPO#(KKJ%) = 0
          Next KKJ%
       End If
    Next KKI%
    '
    KKI% = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    For KKI% = 1 To NUDOCU%
       For KKJ% = KKI% + 1 To NUDOCU%
          If Abs(IMPO#(KKI%) + IMPO#(KKJ%)) < 0.005 Then
             IMPO#(KKI%) = 0
             IMPO#(KKJ%) = 0
          End If
       Next KKJ%
    Next KKI%
    '
    KKI% = 0: SUMAPEX# = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          SUMAPEX# = SUMAPEX# + IMPO#(KKJ%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    Screen.MousePointer = 1
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       'Call TRACE(20, U&, FIN&)
       If Abs(IMPO#(KKI%)) >= 0.005 Then
          IORIG# = IMPO#(KKI%)
          If REGDO&(KKI%) > 0 Then
             If REGDO&(KKI%) > 900000 Then
                X$ = REGLEIDO$("SALINID", REGDO&(KKI%) - 900000)
                NRO = CVS(Mid$(X$, 5, 4))
                TIPO% = Asc(Mid$(X$, 3, 1))
                VA% = Asc(Mid$(X$, 4, 1))
                FFAC% = CVI(Mid$(X$, 9, 2))
                COPAGI% = 0
                GoTo 290
             End If
             X$ = REGLEIDO$("FACTUR", REGDO&(KKI%))
             NRO = CVS(Mid$(X$, 3, 4))
             TIPO% = Asc(Mid$(X$, 13, 1))
             VA% = Asc(Mid$(X$, 14, 1))
             FFAC% = CVI(Mid$(X$, 7, 2))
             COPAGI% = CPAGCLI%(NC%)
             If TIPO% < 1 Or TIPO% > 3 Then COPAGI% = 0
             SG% = 1: If TIPO% >= 4 Then If TIPO% <= 6 Then SG% = -1
             IORIG# = CVD(Mid$(X$, 15, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 23, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 31, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 49, 8))
             IORIG# = IORIG# + CVD(Mid$(X$, 57, 8))
             IORIG# = SG% * IORIG#
           Else
             X$ = REGLEIDO$("COBRAN", (-1) * REGDO&(KKI%))
             NRO = CVS(Mid$(X$, 3, 4))
             TIPO% = 4
             VA% = 0
             FFAC% = CVI(Mid$(X$, 7, 2))
             COPAGI% = 0
          End If
290       Z$ = REGBLAN$("FAPECOB")
          Call REPLA(Z$, MKI$(NC%), 1, 2)
          Call REPLA(Z$, DOCU$(KKI%), 3, 10)
          Call REPLA(Z$, MKI$(FFDOC%(KKI%)), 13, 2)
          Call REPLA(Z$, MKI$(FFDOC%(KKI%)), 15, 2)
          Call REPLA(Z$, MKD$(IMPO#(KKI%)), 17, 8)
          Call REPLA(Z$, MKD$(IORIG#), 25, 8)
          '
          UJ& = UJ& + 1
          Call GRAREG("FAPECOB", Z$, UJ&)
       End If
    Next KKI%
    '
    If Abs(SUMARRAS# - SUMAPEX#) >= 0.005 Then
      Call MENSERR(24, "Diferencia de Saldo Compensada en Composición de Deuda\de la Cuenta '" + TRIM$(RASOCLI$(NC%)) + "'.")
      '
      Z$ = REGBLAN$("FAPECOB")
      Call REPLA(Z$, MKI$(NC%), 1, 2)
      Call REPLA(Z$, "Dif.Saldo", 3, 10)
      Call REPLA(Z$, MKI$(FLIM%), 13, 2)
      Call REPLA(Z$, MKI$(FLIM%), 15, 2)
      Call REPLA(Z$, MKD$(SUMARRAS# - SUMAPEX#), 17, 8)
      Call REPLA(Z$, MKD$(SUMARRAS# - SUMAPEX#), 25, 8)
      '
      UJ& = UJ& + 1
      Call GRAREG("FAPECOB", Z$, UJ&)
    End If
    '
999 'Call CIERRARCH(ARCHIDEU$)
    'Call CIERRARCH("*.*")
    'Call BAJALDISCO
    'Screen.MousePointer = 1
    'ATENTI% = 0
End Sub




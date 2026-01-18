VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPLISCO07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ORDENES DE COMPRA - PROVEEDORES EXTERNOS"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   39
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   8700
      ScaleHeight     =   4185
      ScaleWidth      =   1035
      TabIndex        =   32
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   42
         Top             =   3720
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   43
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton AGenRep 
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
         Left            =   105
         Picture         =   "OPLISCO07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   2190
         Width           =   650
      End
      Begin VB.CommandButton Command18 
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
         Left            =   105
         Picture         =   "OPLISCO07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Configuración de Procesos y Formularios"
         Top             =   1485
         Width           =   650
      End
      Begin VB.CommandButton Command8 
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
         Left            =   105
         Picture         =   "OPLISCO07.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Acceso Directo a Otros Procesos"
         Top             =   2800
         Width           =   650
      End
      Begin VB.CommandButton Command29 
         Caption         =   "Help"
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
         Left            =   105
         Picture         =   "OPLISCO07.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Ayuda Flexoft en Línea"
         Top             =   780
         Width           =   650
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
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
         Left            =   105
         Picture         =   "OPLISCO07.frx":0D60
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   80
         Width           =   650
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00CDDADA&
      Caption         =   "CONSULTAS VARIAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   4410
      TabIndex        =   3
      Top             =   210
      Width           =   4005
      Begin VB.Frame Frame4 
         BackColor       =   &H00CDDADA&
         Caption         =   "SEGUIMIENTO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   210
         TabIndex        =   25
         Top             =   240
         Width           =   3585
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   840
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   29
            Top             =   840
            Width           =   2325
            Begin VB.CommandButton Command2 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   -20
               Picture         =   "OPLISCO07.frx":0EAA
               Style           =   1  'Graphical
               TabIndex        =   30
               ToolTipText     =   "Aprobar Ordenes de Compra"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Aprobar Ordenes de Compra"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   720
               TabIndex        =   31
               Top             =   75
               Width           =   1485
            End
         End
         Begin VB.PictureBox Picture14 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   525
            ScaleHeight     =   615
            ScaleWidth      =   2400
            TabIndex        =   26
            Top             =   420
            Width           =   2430
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
               Height          =   630
               Left            =   1850
               Picture         =   "OPLISCO07.frx":11B4
               Style           =   1  'Graphical
               TabIndex        =   41
               ToolTipText     =   "Ver y Reimprimir Con Versión"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command22 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "OPLISCO07.frx":14BE
               Style           =   1  'Graphical
               TabIndex        =   27
               ToolTipText     =   "Ver por Pantalla y Re-Imprimir"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Ver y Reimprimir"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   645
               TabIndex        =   28
               Top             =   75
               Width           =   1230
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00CDDADA&
         Caption         =   "CONTROL DE ULTIMAS COMPRAS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1560
         Left            =   210
         TabIndex        =   4
         Top             =   2100
         Width           =   3585
         Begin VB.PictureBox Picture10 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   840
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   5
            Top             =   735
            Width           =   2325
            Begin VB.CommandButton Command15 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   1750
               Picture         =   "OPLISCO07.frx":17C8
               Style           =   1  'Graphical
               TabIndex        =   38
               ToolTipText     =   "Consultas de Entrega por Proveedor"
               Top             =   0
               Width           =   560
            End
            Begin VB.CommandButton Command12 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "OPLISCO07.frx":1AD2
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Ultimas Compras por Proveedor"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Proveedor - Entregas"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   6
               Top             =   80
               Width           =   1065
            End
         End
         Begin VB.PictureBox Picture12 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   480
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   7
            Top             =   360
            Width           =   2325
            Begin VB.CommandButton Command20 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   -20
               Picture         =   "OPLISCO07.frx":1DDC
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Ultimas Compras por Material"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Por Artículo"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   750
               TabIndex        =   8
               Top             =   80
               Width           =   1005
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CDDADA&
      Caption         =   "CONTROL DE COMPRAS PENDIENTES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   240
      TabIndex        =   2
      Top             =   210
      Width           =   4005
      Begin VB.Frame Frame6 
         BackColor       =   &H00CDDADA&
         Caption         =   "Salida"
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
         Left            =   210
         TabIndex        =   22
         Top             =   3060
         Width           =   3585
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Pantalla"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   420
            TabIndex        =   24
            Top             =   250
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H00CDDADA&
            Caption         =   "Impresión"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   1995
            TabIndex        =   23
            Top             =   250
            Width           =   1275
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00CDDADA&
         Caption         =   "OPCIONES DE CONSULTA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   210
         TabIndex        =   9
         Top             =   360
         Width           =   3585
         Begin VB.PictureBox Picture11 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   945
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   10
            Top             =   1680
            Width           =   2325
            Begin VB.CommandButton Command9 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   0
               Picture         =   "OPLISCO07.frx":1F26
               Style           =   1  'Graphical
               TabIndex        =   11
               ToolTipText     =   "Material en Consignación"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label12 
               BackStyle       =   0  'Transparent
               Caption         =   "Material en Consignación"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   750
               TabIndex        =   12
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   735
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   13
            Top             =   1260
            Width           =   2325
            Begin VB.CommandButton Command5 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   -20
               Picture         =   "OPLISCO07.frx":2230
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Pendientes por Material"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label6 
               BackStyle       =   0  'Transparent
               Caption         =   "Por Material"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   15
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   525
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   16
            Top             =   840
            Width           =   2325
            Begin VB.CommandButton Command4 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   630
               Left            =   -20
               Picture         =   "OPLISCO07.frx":253A
               Style           =   1  'Graphical
               TabIndex        =   17
               ToolTipText     =   "Pendientes por Proveedor"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Por Proveedor"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   735
               TabIndex        =   18
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   645
            Left            =   315
            ScaleHeight     =   615
            ScaleWidth      =   2295
            TabIndex        =   19
            Top             =   420
            Width           =   2325
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
               Height          =   630
               Left            =   0
               Picture         =   "OPLISCO07.frx":297C
               Style           =   1  'Graphical
               TabIndex        =   20
               ToolTipText     =   "Ordenes de Compra por Número"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "O/C's por Número"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Left            =   630
               TabIndex        =   21
               Top             =   105
               Width           =   1590
            End
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "OPLISCO07.frx":2C86
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "OPLISCO07.frx":2EBA
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "OPLISCO07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim HAYMATERIALIB%
Dim OCOMP0TEMP As ADODB.Recordset

Function CALCULA_TOTAL_PENDIENTE#()
    '
    CALCULA_TOTAL_PENDIENTE# = 0
    '
    CONDI$ = "  Stat_Ocom < 8 "
    CONDI$ = CONDI$ + " AND Stat_Ocom <> -1 "
    CONDI$ = CONDI$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
    '
    'VERIFICO SI EXISTE ALGUN PROVEEDOR EXCLUIDO
    FIN& = ULTREG&("PROVEXC")
    If FIN& > 0 Then
        CONDI1$ = ""
        For II& = 1 To FIN&
            XX$ = REGLEIDO$("PROVEXC", II&)
            NPRO% = CVI(Mid$(XX$, 1, 2))
            CONDI1$ = CONDI1$ & " AND NPRO_OCOM<>" & NPRO% & " "
        Next II&
    End If
    '
    'PERMITE EL INGRESO DE UN LIMITE DE FECHA 'HASTA'
    Call SELECHO("SELFECHA/Compras Pendientes al ...")
    VDEV$ = TRIM$(VALACT1$("SELFECHA"))
    If VDEV$ = "" Then Exit Function
    FF% = FECHANUM(VDEV$): VDEV$ = FETEXCOR1$(FF%)
    CONDI2$ = " AND FENTRE_SOL<='" & VDEV$ & "'"
    '
    FIN& = CantRegistros&("OCOMDETA", CONDI$ & CONDI1$ & CONDI2$, "NOMESS")
    If FIN& < 1 Then Exit Function
    '
    SQLX$ = "Select * from OCOMDETA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " Where " & CONDI$ & CONDI1$ & CONDI2$
    Set RS = READSET(SQLX$)
    '
    With RS
     Do While Not .EOF
        J& = J& + 1
        Call TRACE(20, J&, FIN&)
        CANTI = XVALO(!Cant_Ocom) - XVALO(!Cant_Rec)
        
        COEFUNI = XVALO(!Coef_Unids): If COEFUNI < 0.0001 Then COEFUNI = 1
        MONEMISION% = XVALO(!MONE_COS)
        TICAMBIO = 1: If MONEMISION% >= 2 Then TICAMBIO = TICAMONE(MONEMISION%)
        If TICAMBIO <= 0 Then TICAMBIO = 1
        COSTOUNI# = XVALO(!prun_neto) * TICAMBIO
        CALCULA_TOTAL_PENDIENTE# = CALCULA_TOTAL_PENDIENTE# + (CANTI * COSTOUNI#)
     .MoveNext
     Loop
    End With
    
End Function

Sub COMPRAPENDIENTE()
    '
    'BORRO LOS REGISTROS
    '*****************************************************
    CONDI$ = "Cod_Inte >= 0"
    Call BORRAREGISTROS("DETAPENDIENTE", CONDI$, REGAF&, "NOMESS")
    '---------------------------------------------------------------------------------------

    'RECORRO LOS ITEMS DE LA ORDEN DE COMPRA SELECCIONADA
    '*****************************************************
    CONDI$ = "  Stat_Ocom < 8 "
    CONDI$ = CONDI$ + " AND Stat_Ocom <> -1 "
    CONDI$ = CONDI$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
    '
    'VERIFICO SI EXISTE ALGUN PROVEEDOR EXCLUIDO
    FIN& = ULTREG&("PROVEXC")
    If FIN& > 0 Then
        CONDI1$ = ""
        For II& = 1 To FIN&
            XX$ = REGLEIDO$("PROVEXC", II&)
            NPRO% = CVI(Mid$(XX$, 1, 2))
            CONDI1$ = CONDI1$ & " AND NPRO_OCOM<>" & NPRO% & " "
        Next II&
    End If
    '
    'PERMITE EL INGRESO DE UN LIMITE DE FECHA 'HASTA'
    Call SELECHO("SELFECHA/Compras Pendientes al ...")
    VDEV$ = TRIM$(VALACT1$("SELFECHA"))
    If VDEV$ = "" Then Exit Sub
    FF% = FECHANUM(VDEV$): VDEV$ = FETEXCOR1$(FF%)
    CONDI2$ = " AND FENTRE_SOL<='" & VDEV$ & "'"
    '
    Call APERBASDAT("COMPRAS")
    FIN& = CantRegistros&("OCOMDETA", CONDI$ & CONDI1$ & CONDI2$, "NOMESS")
    If FIN& < 1 Then
      Call COMUNI("No Existen Ordenes de Compra Pendientes de Entrega")
      Exit Sub
    End If
    '
    SQLX$ = "Select * from OCOMDETA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " Where " & CONDI$ & CONDI1$ & CONDI2$
    SQLX$ = SQLX$ + " ORDER BY FENTRE_SOL"
    Set RS = READSET(SQLX$)
    '
    With RS
     Do While Not .EOF
        J& = J& + 1
        Call TRACE(20, J&, FIN&)
        Ci% = XVALO(!cod_inte)
        NPRO% = XVALO(!Npro_Ocom)
        CANTI = XVALO(!Cant_Ocom) - XVALO(!Cant_Rec)
        
        COEFUNI = XVALO(!Coef_Unids): If COEFUNI < 0.0001 Then COEFUNI = 1
        MONEMISION% = XVALO(!MONE_COS)
        TICAMBIO = 1: If MONEMISION% >= 2 Then TICAMBIO = TICAMONE(MONEMISION%)
        If TICAMBIO <= 0 Then TICAMBIO = 1
        
        COSTOUNI# = XVALO(!prun_neto) * TICAMBIO
        fech% = CVINT(!Fentre_Sol)
        Mesx% = XVALO(Mid$(FECHATEX$(fech%), 4, 2))
        '
        costo = costo + COSTOUNI# * CANTI
          SQLX$ = "SELECT * FROM DETAPENDIENTE "
          SQLX$ = SQLX$ + " WHERE COD_INTE < 0 "
          Set RSTMP = New ADODB.Recordset
25        On Error Resume Next
          RSTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 25
          End If
          On Error GoTo 0
          '
          With RSTMP
            .AddNew
            !cod_inte = Ci%
            !Cantidad = CANTI
            !Npro_Ocom = NPRO%
            !Fentre_Sol = CVDAT(fech%)
            !COSUNI = COSTOUNI#
            !MESENT = Mesx%
            .Update
          End With
      .MoveNext
      Loop
     End With
     RSTMP.Close
     Set RSTMP = Nothing
    '---------------------------------------------------------------------------------------
    'PRESENTO LISTADO
    Call FINAL("?DETACOMPEN")
End Sub

Sub MATPENDCONS()
   '08/05/07 (OT-07-0182)
   Call APERBASDAT("STOCKS")
   SQLX$ = "Select * from MACONSIG WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE Cant_Asig > Cant_Entreg "
   '
   'Set GECONPENTEMP = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
'   Dim GECONPENTEMP As ADODB.Recordset
'   Set GECONPENTEMP = New ADODB.Recordset
'   GECONPENTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   Set GECONPENTEMP = READSET(SQLX$)
   With GECONPENTEMP
   On Error Resume Next
'       .MoveLast
'       'Valida que la tabla no este vacia
'       If Err Then
'         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
'         On Error GoTo 0
'         GoTo 90
'       End If
'       On Error Resume Next
'     FIN& = .RecordCount
'     .MoveFirst
     JJ& = 0: HUBORECU% = 0: U2& = 0
     Do While (.EOF = False)
         U2& = U2& + 1
         Call TRACE(20, U2&, FIN&)
         '
         NUMEOC& = XVALO(Mid$(!nro_ocom, 4, 6))
         NORDI% = XVALO(Mid$(!nro_ocom, 11, 2))
         '
         If OCOPENDIENTE%(NUMEOC&, NORDI%) < 1 Then
            GECONPENTEMP!cant_asig = GECONPENTEMP!cant_entreg
            GECONPENTEMP.Update
            HUBORECU% = 1
            GoTo 19
         End If
         '
         XXY$ = REGBLAN$("MACONSIG")
         Call REPLA(XXY$, MKI$(!Npro_Ocom), 3, 2)
         Call REPLA(XXY$, !nro_ocom, 21, 12)
         Call REPLA(XXY$, MKI$(!cod_inte), 33, 2)
         Call REPLA(XXY$, !Unidad, 51, 4)
         Call REPLA(XXY$, MKD$(!cant_asig), 55, 8)
         Call REPLA(XXY$, MKD$(!cant_entreg), 63, 8)
         SALPE# = 0
         SALPE# = !cant_asig - !cant_entreg
         Call REPLA(XXY$, MKD$(SALPE#), 119, 8)
         JJ& = JJ& + 1
         Call GRAREG("MACONSIG", XXY$, JJ&)
19   .MoveNext
     Loop
  End With
  GECONPENTEMP.Close
  Set GECONPENTEMP = Nothing
  '
  If JJ& < 1 Then
    Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
    GoTo 90
  End If
  '
  Call SETULTREG("MACONSIG", JJ&)
  Call CIERRARCH("MACONSIG")
  '
  If HUBORECU% > 0 Then
     Call COMUNI("Se han Suprimido Asignaciones\Correspondientes a Ordenes de Compras\Cumplidas o Anuladas.")
  End If
  '
  Call FILESORT("MACONSIG", "MACONSIG", 8, 8, "ASC")
  Call FILESORT("MACONSIG", "MACONSIG", 7, 7, "ASC")
  Call FINAL("*MAACONS")
  '
90 Screen.MousePointer = 1
  '08/05/07 (FIN)
End Sub

Private Sub OrdenesComprasPendientes()

   Screen.MousePointer = 11
   
   Call APERBASDAT("COMPRAS")
   
   Dim oReporte As cGenrep
   Dim Rst As ADODB.Recordset
    
   Set oReporte = New cGenrep
   Set Rst = oReporte.GeneraReporteConTablaTemporal("COCOPEN", "ORDENES DE COMPRAS PENDIENTES", _
             "Orden(T=O-Compra)/I10; FechaE(T=F.Emision)/D9; Prov(ECO=PROVED1-A32)/I10; Codigo/A24; Descripcion/A36; Cantidad/F10.1;" _
             & "UMC/A5; Mon/A5; PrecioN(T=P.Neto)/F12.4; Importe/F12.2; FES(T=F.E.S.)/D8; Recibido/F10.1;" _
             & "Pendiente/F10.1; PrecioPesos(T=P.Neto [$])/F12.4; ImportePesos(T=Importe [$]|TT=SI)/F12.2; Destino/A24; Sector/A24")
   
   
   SQLX$ = "Select * from OCOMDETA WITH(NOLOCK) "
   SQLX$ = SQLX$ + " Where Stat_Ocom < 3 "
   SQLX$ = SQLX$ + " AND Stat_Ocom <> -1 "
   SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
   SQLX$ = SQLX$ + " ORDER BY Nume_Ocom, Cod_Exte, FEntre_Sol "
   
   Set OCOMTEMP = READSET(SQLX$)

   With OCOMTEMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        Call MENSERR(24, "No se han encontrado ordenes de compras pendientes.")
        On Error GoTo 0
        GoTo 90
     End If
     On Error GoTo 0
     
     JJ& = 0: FIN& = OCOMTEMP.RecordCount
     Do While (.EOF = False)
         Call TRACE(24, JJ&, FIN&)
         
         MONEMI% = !Mone_Emis
         If MONEMI% <= 1 Then MONEX$ = "   "
         If MONEMI% = 2 Then MONEX$ = "U$S "
         If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
         TICAMBIO = 1
         If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
         '
         Call REPLA(XXY$, MONEX$, 63, 4)
         CAPEN = XVALO(!Cant_Ocom) - XVALO(!Cant_Rec)
         If CAPEN < 0.05 Then CAPEN = 0
         '
         CCO% = XVALO(!IContab)
         DESTINO$ = "": If CCO% > 0 Then DESTINO$ = CODCUE$(CCO%)
         
         Rst.AddNew
         Rst.Fields(0) = !nume_ocom
         Rst.Fields(1) = !Femi_Ocom
         Rst.Fields(2) = !Npro_Ocom
         Rst.Fields(3) = !Cod_Exte
         Rst.Fields(4) = Left(SAFETEXT$(!DESCRITEM), 36)
            If !cod_inte <= 0 Then Rst.Fields(4) = Left(SAFETEXT$(!DELIBRE), 36)
         Rst.Fields(5) = !Cant_Ocom
         Rst.Fields(6) = !Uni_Com
         Rst.Fields(7) = MONEX$
         Rst.Fields(8) = !prun_neto
         Rst.Fields(9) = XVALO(!Cant_Ocom) * XVALO(!prun_neto)
         Rst.Fields(10) = !Fentre_Sol
         Rst.Fields(11) = !Cant_Rec
         Rst.Fields(12) = CAPEN
            PreUnid# = XVALO(!prun_neto) * TICAMBIO
         Rst.Fields(13) = PreUnid#
         Rst.Fields(14) = XVALO(PreUnid#) * CAPEN
         Rst.Fields(15) = Left(DESTINO$, 24)
         Rst.Fields(16) = Left(SAFETEXT$(!Sect_Ocom), 24)
         Rst.Update
            
         JJ& = JJ& + 1
         .MoveNext
    Loop
   End With
   '
   oReporte.MostrarReporte
   '
90 Screen.MousePointer = 1

End Sub

Private Sub OrdenesComprasPorProveedor()

   Screen.MousePointer = 11
   
   Call APERBASDAT("COMPRAS")
   
   Dim oReporte As cGenrep
   Dim Rst As ADODB.Recordset
   Dim OCOMTEMP  As ADODB.Recordset
   
   Set oReporte = New cGenrep
   Set Rst = oReporte.GeneraReporteConTablaTemporal("STATOCOM", "ORDENES DE COMPRAS POR PROVEEDOR", _
             "Orden(T=O-Compra)/I10; Proveedor(FILTRO=Proveedor-PROVED1)/I10; FechaE(T=F.Emision)/D9; Codigo/A24; Descripcion/A36; " _
             & "UMC/A5; Cantidad/F10.1; Recibido/F10.1; FES(T=F.E.S.)/D8; Mon/A5; PrecioN(T=P.Neto)/F12.4;")
   
   
   SQLX$ = "SELECT * FROM OCOMDETA WITH(NOLOCK) "
   SQLX$ = SQLX$ + " ORDER BY Nume_Ocom, Cod_Exte, FEntre_Sol "
   Set OCOMTEMP = READSET(SQLX$)

   With OCOMTEMP
     FIN& = OCOMTEMP.RecordCount
     If FIN& < 1 Then
        Call MENSERR(24, "No se han encontrado ordenes de compras pendientes.")
        GoTo 90
     End If
     
     JJ& = 0:
     Do While (.EOF = False)
         Call TRACE(24, JJ&, FIN&)
         
         MONEMI% = !Mone_Emis
         If MONEMI% <= 1 Then MONEX$ = "   "
         If MONEMI% = 2 Then MONEX$ = "U$S "
         If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
         TICAMBIO = 1
         If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
         '
         Rst.AddNew
         Rst.Fields(0) = !nume_ocom
         Rst.Fields(1) = !Npro_Ocom
         Rst.Fields(2) = !Femi_Ocom
         Rst.Fields(3) = !Cod_Exte
         Rst.Fields(4) = Left(SAFETEXT$(!DESCRITEM), 36)
            If !cod_inte <= 0 Then Rst.Fields(4) = Left(SAFETEXT$(!DELIBRE), 36)
         Rst.Fields(5) = !Uni_Com
         Rst.Fields(6) = !Cant_Ocom
         Rst.Fields(7) = !Cant_Rec
         Rst.Fields(8) = !Fentre_Sol
         Rst.Fields(9) = MONEX$
         Rst.Fields(10) = !Pre_Unit
         Rst.Update
            
         JJ& = JJ& + 1
         .MoveNext
    Loop
   End With
   '
   On Error Resume Next
   OCOMTEMP.Close
   Set OCOMTEMP = Nothing
   On Error GoTo 0
   '
   oReporte.MostrarReporte
   '
90 Screen.MousePointer = 1

End Sub

Private Sub OrdenesPendientesPorProveedor()
   EBOL$ = OBJPLA$("SELPROVE", "")
   NPRO% = CVI(EBOL$)

   Screen.MousePointer = 11
   
   Call APERBASDAT("COMPRAS")
   
   Dim oReporte As cGenrep
   Dim Rst As ADODB.Recordset
    
   Set oReporte = New cGenrep
   Set Rst = oReporte.GeneraReporteConTablaTemporal("CORCPRO", "ORDENES DE COMPRAS PENDIENTES POR PROVEEDOR", _
             "Numero/I10; FechaE(T=F.Emision)/D9; Proveedor(ECO=PROVED1-A24)/I10; Codigo/A24; Descripcion/A36; Cantidad/F10.1;" _
             & "UMC/A5; Mon/A5; PrecioN(T=P.Neto)/F12.4; Importe/F12.2; Entrega/D8; Recibido/F10.1;" _
             & "Pendiente/F10.1; PrecioPesos(T=P.Neto [$])/F12.4; ImportePesos(T=Importe [$]|CC=3)/F14.2")
   
   
   SQLX$ = "Select * from OCOMDETA WITH(NOLOCK) "
   SQLX$ = SQLX$ + " Where Stat_Ocom < 3 "
   SQLX$ = SQLX$ + " AND Stat_Ocom <> -1 "
   SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
   If NPRO% > 0 Then SQLX$ = SQLX$ + " AND Npro_Ocom = " & NPRO%
   SQLX$ = SQLX$ + " ORDER BY Npro_Ocom, Nume_Ocom, Cod_Exte, FEntre_Sol "
   
   Set OCOMTEMP = READSET(SQLX$)

   With OCOMTEMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        Call MENSERR(24, "No se han encontrado ordenes de compras pendientes.")
        On Error GoTo 0
        GoTo 90
     End If
     On Error GoTo 0
     
     JJ& = 0: FIN& = OCOMTEMP.RecordCount
     Do While (.EOF = False)
         Call TRACE(24, JJ&, FIN&)
         
         MONEMI% = !Mone_Emis
         If MONEMI% <= 1 Then MONEX$ = "   "
         If MONEMI% = 2 Then MONEX$ = "U$S "
         If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
         TICAMBIO = 1
         If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
         '
         Call REPLA(XXY$, MONEX$, 63, 4)
         CAPEN = XVALO(!Cant_Ocom) - XVALO(!Cant_Rec)
         If CAPEN < 0.05 Then CAPEN = 0
         '
         CCO% = XVALO(!IContab)
         DESTINO$ = "": If CCO% > 0 Then DESTINO$ = CODCUE$(CCO%)
         
         Rst.AddNew
         Rst.Fields(0) = !nume_ocom
         Rst.Fields(1) = !Femi_Ocom
         Rst.Fields(2) = !Npro_Ocom
         Rst.Fields(3) = !Cod_Exte
         Rst.Fields(4) = Left(SAFETEXT$(!DESCRITEM), 36)
            If !cod_inte <= 0 Then Rst.Fields(4) = Left(SAFETEXT$(!DELIBRE), 36)
         Rst.Fields(5) = !Cant_Ocom
         Rst.Fields(6) = !Uni_Com
         Rst.Fields(7) = MONEX$
         Rst.Fields(8) = !prun_neto
         Rst.Fields(9) = XVALO(!Cant_Ocom) * XVALO(!prun_neto)
         Rst.Fields(10) = !Fentre_Sol
         Rst.Fields(11) = !Cant_Rec
         Rst.Fields(12) = CAPEN
            PreUnid# = XVALO(!prun_neto) * TICAMBIO
         Rst.Fields(13) = PreUnid#
         Rst.Fields(14) = XVALO(PreUnid#) * CAPEN
         Rst.Update
            
         JJ& = JJ& + 1
         .MoveNext
    Loop
   End With
   '
   oReporte.MostrarReporte
   '
90 Screen.MousePointer = 1

End Sub

Private Sub OrdenesPendientesPorMaterial()

   Screen.MousePointer = 11
   
   Call APERBASDAT("COMPRAS")
   
   Dim oReporte As cGenrep
   Dim Rst As ADODB.Recordset
   Dim OCOMTEMP As ADODB.Recordset
    
   Set oReporte = New cGenrep
   Set Rst = oReporte.GeneraReporteConTablaTemporal("CORCPRO", "ORDENES DE COMPRAS PENDIENTES POR PROVEEDOR", _
             "Numero/I10; FechaE(T=F.Emision)/D9; Proveedor(ECO=PROVED1-A24)/I10; Codigo/A24; Descripcion/A36; Cantidad/F10.1;" _
             & "UMC/A5; Mon/A5; PrecioN(T=P.Neto)/F12.4; Importe/F12.2; Entrega/D8; Recibido/F10.1;" _
             & "Pendiente/F10.1; PrecioPesos(T=P.Neto [$])/F12.4; ImportePesos(T=Importe [$]|CC=4)/F12.2")
   
   
   SQLX$ = "Select * from OCOMDETA WITH(NOLOCK) "
   SQLX$ = SQLX$ + " Where Stat_Ocom < 3 "
   SQLX$ = SQLX$ + " AND Stat_Ocom <> -1 "
   SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
   SQLX$ = SQLX$ + " ORDER BY Cod_Exte, Nume_Ocom, FEntre_Sol "
   
   Set OCOMTEMP = READSET(SQLX$)

   With OCOMTEMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        Call MENSERR(24, "No se han encontrado ordenes de compras pendientes.")
        On Error GoTo 0
        GoTo 90
     End If
     On Error GoTo 0
     
     JJ& = 0: FIN& = OCOMTEMP.RecordCount
     Do While (.EOF = False)
         Call TRACE(24, JJ&, FIN&)
         
         MONEMI% = !Mone_Emis
         If MONEMI% <= 1 Then MONEX$ = "   "
         If MONEMI% = 2 Then MONEX$ = "U$S "
         If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
         TICAMBIO = 1
         If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
         '
         Call REPLA(XXY$, MONEX$, 63, 4)
         CAPEN = XVALO(!Cant_Ocom) - XVALO(!Cant_Rec)
         If CAPEN < 0.05 Then CAPEN = 0
         '
         CCO% = XVALO(!IContab)
         DESTINO$ = "": If CCO% > 0 Then DESTINO$ = CODCUE$(CCO%)
         
         Rst.AddNew
         Rst.Fields(0) = !nume_ocom
         Rst.Fields(1) = !Femi_Ocom
         Rst.Fields(2) = !Npro_Ocom
         Rst.Fields(3) = !Cod_Exte
         Rst.Fields(4) = Left(SAFETEXT$(!DESCRITEM), 36)
            If !cod_inte <= 0 Then Rst.Fields(4) = Left(SAFETEXT$(!DELIBRE), 36)
         Rst.Fields(5) = !Cant_Ocom
         Rst.Fields(6) = !Uni_Com
         Rst.Fields(7) = MONEX$
         Rst.Fields(8) = !prun_neto
         Rst.Fields(9) = XVALO(!Cant_Ocom) * XVALO(!prun_neto)
         Rst.Fields(10) = !Fentre_Sol
         Rst.Fields(11) = !Cant_Rec
         Rst.Fields(12) = CAPEN
            PreUnid# = XVALO(!prun_neto) * TICAMBIO
         Rst.Fields(13) = PreUnid#
         Rst.Fields(14) = XVALO(PreUnid#) * CAPEN
         Rst.Update
            
         JJ& = JJ& + 1
         .MoveNext
    Loop
   End With
   '
   On Error Resume Next
   OCOMTEMP.Close
   Set OCOMTEMP = Nothing
   On Error GoTo 0
   '
   oReporte.MostrarReporte
   '
90 Screen.MousePointer = 1

End Sub
Sub REPORTE_COBERF()
    '
    FIN& = ULTREG&("COBERFA")
    If FIN& < 1 Then
        Call CIERRARCH("*.*")
        Exit Sub
    End If
    '
    OBX$ = OBJPLA$("DESHASCOD", VDEF$)
    If OBX$ <> "" Then
        Screen.MousePointer = 11
        CI1% = CVI(Left$(OBX$, 2))
        COD1$ = CODEXT$(CI1%)
        CI2% = CVI(Mid$(OBX$, 3, 2))
        COD2$ = CODEXT$(CI2%)
        '
        J& = 0
        For U& = 1 To FIN&
            Call TRACE(20, U&, FIN&)
            YY$ = REGLEIDO$("COBERFA", U&)
            CIII% = CVI(Mid$(YY$, 3, 2))
            CODX$ = CODEXT$(CIII%)
            If CODX$ >= COD1$ Then
                If CODX$ <= COD2$ Then
                  J& = J& + 1
                  Call GRAREG("COBERF1", YY$, J&)
                End If
            End If
        Next U&
        '
        Call SETULTREG("COBERF1", J&)
        Call CIERRARCH("COBERF1")
        Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        Call FINAL("*LICOB05")
        '
    End If
    '
End Sub

Private Sub UltimasComprasPorMaterial()
    '
    Call APERBASDAT("COMPRAS")
   
    Dim oReporte As cGenrep
    Dim Rst As ADODB.Recordset
    Dim GECONPENTEMP As ADODB.Recordset
    
    Set oReporte = New cGenrep
    Set Rst = oReporte.GeneraReporteConTablaTemporal("LICOMAR", "ULTIMAS COMPRAS POR MATERIAL", _
              "Codigo/A24; Descripcion/A36; FechaEm(T=F.Emision)/D9; Prov(ECO=PROVED1-A24)/I10; Operacion/A20; OCompra/A20;" _
              & "UM/A4; Cantidad(CC=1)/F10.1; Mon/A4; PrecioUnitario(T=P.Unit)/F12.4; Importe/F12.2; PrecioUnitPesos(T=P.Unit [$])/F12.4;" _
              & "ImportePesos(T=Importe [$]|CC=1|TT=SI)/F12.2")
   
    Has% = HOY(HOS$)
    FEX = Has%
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
10  OKKX% = 0
    Call DESHASFECHA(Des%, Has%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    '
    OKKX% = 1
    Screen.MousePointer = 11
    
    Desde = FECHATEXLAR$(Des%)
    Hasta = FECHATEXLAR$(Has%)
    REBLA$ = REGBLAN$("COPRECOM")
    XXZ$ = REGBLAN("COPRECOL")
    Call HEADERS("COPRECOM", "")
    Call HEADERS("COPRECOM", "Periodo: " + PERIO$)
    '
    Call APERBASDAT("COMPRAS")
    SQLX$ = "SELECT * FROM OCOMDETA WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE Stat_Ocom < 9 "
    SQLX$ = SQLX$ + " AND Femi_Ocom >= '" & Desde & "' "
    SQLX$ = SQLX$ + " AND Femi_Ocom <= '" & Hasta & "' "
    SQLX$ = SQLX$ + " AND Nume_Ocom > 0 "
    SQLX$ = SQLX$ + "ORDER BY Cod_Exte ASC, Femi_Ocom ASC"
    '
    Dim OCOMPTEMP As ADODB.Recordset
    Set OCOMPTEMP = READSET(SQLX$)
    With OCOMPTEMP
        FIN& = .RecordCount
        Do While (.EOF = False)
            KU& = KU& + 1
            Call TRACE(20, KU&, FIN&)
            
            FECHEMI% = CVINT(!Femi_Ocom)
            STAT% = XVALO(!Stat_Ocom)
            NPRO% = XVALO(!Npro_Ocom)
            CIXI% = XVALO(!cod_inte)
            DOPE$ = ""
            '
            NOPE% = !Cod_Oper
            If NOPE% > 0 Then
              DOPE$ = DESCRIOP$(CIXI%, NOPE%)
            End If
            NUOC& = XVALO(!nume_ocom)
            '
            UMEDI$ = !Uni_Com
            If TRIM$(UMEDI$) = "" Then
              UMEDI$ = Unimed$(CIXI%)
            End If
            CACOM = !Cant_Ocom
            '
            If STAT% = 8 Then                 ' SI SE DIO POR CUMPLIDA,
              CACOM = !Cant_Rec
            End If
            '
            PREUNI# = !Pre_Unit
            '
            MONEMI% = !Mone_Emis
            If MONEMI% <= 1 Then MONEX$ = "    "
            If MONEMI% = 2 Then MONEX$ = "U$S "
            If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
            TICAMBIO = 1
            If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
            '
            
            Rst.AddNew
            Rst.Fields(0) = !Cod_Exte
            Rst.Fields(1) = Left(!DESCRITEM, 36)
            If !cod_inte <= 0 Then Rst.Fields(1) = Left(!DELIBRE, 36)
            Rst.Fields(2) = CVDAT(FECHEMI%)
            Rst.Fields(3) = NPRO%
            Rst.Fields(4) = Left(DOPE$, 20)
            Rst.Fields(5) = NUOC&
            Rst.Fields(6) = Left(UMEDI$, 4)
            Rst.Fields(7) = CACOM
            Rst.Fields(8) = Left(MONEX$, 4)
            Rst.Fields(9) = PREUNI#
                ITOCOM# = PREUNI# * CACOM
            Rst.Fields(10) = ITOCOM#
            Rst.Fields(11) = PREUNI# * TICAMBIO
            Rst.Fields(12) = ITOCOM# * TICAMBIO
            
            If TRIM$(DEST1$) = "" Then
              CCO% = !IContab
              DEST1$ = CODCUE$(CCO%)
            End If
            '
            If IsNull(!Sect_Ocom) Then
                SECT1$ = ""
            Else
                SECT1$ = !Sect_Ocom
            End If
            Rst.Update
         
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    OCOMPTEMP.Close
    Set OCOMPTEMP = Nothing
    On Error GoTo 0
    '
    Screen.MousePointer = 1
    '
    oReporte.MostrarReporte
    '
End Sub

Private Sub UltimasComprasPorProveedor()
    '
    Call APERBASDAT("COMPRAS")
   
    Dim oReporte As cGenrep
    Dim Rst As ADODB.Recordset
    Dim GECONPENTEMP As ADODB.Recordset
    
    Set oReporte = New cGenrep
    Set Rst = oReporte.GeneraReporteConTablaTemporal("LICOMPR", "ULTIMAS COMPRAS POR PROVEEDOR", _
              "Prov(ECO=PROVED1-A24)/I10; FechaEm(T=F.Emision)/D9; Codigo/A24; Descripcion/A36; Operacion/A20; OCompra/A20;" _
              & "UM/A4; Cantidad(CC=1)/F10.1; Mon/A4; PrecioUnitario(T=P.Unit)/F12.4; Importe/F12.2; PrecioUnitPesos(T=P.Unit [$])/F12.4;" _
              & "ImportePesos(T=Importe [$]|CC=1|TT=SI)/F12.2")
   
    Has% = HOY(HOS$)
    FEX = Has%
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    Des% = FECHANUM(DESS$)
    VDEF$ = MKI$(Des%) + MKI$(Has%)
    '
10  OKKX% = 0
    Call DESHASFECHA(Des%, Has%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    '
    OKKX% = 1
    Screen.MousePointer = 11
    
    Desde = FECHATEXLAR$(Des%)
    Hasta = FECHATEXLAR$(Has%)
    REBLA$ = REGBLAN$("COPRECOM")
    XXZ$ = REGBLAN("COPRECOL")
    Call HEADERS("COPRECOM", "")
    Call HEADERS("COPRECOM", "Periodo: " + PERIO$)
    '
    Call APERBASDAT("COMPRAS")
    SQLX$ = "SELECT * FROM OCOMDETA WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE Stat_Ocom < 9 "
    SQLX$ = SQLX$ + " AND Femi_Ocom >= '" & Desde & "' "
    SQLX$ = SQLX$ + " AND Femi_Ocom <= '" & Hasta & "' "
    SQLX$ = SQLX$ + " AND Nume_Ocom > 0 "
    SQLX$ = SQLX$ + " ORDER BY Npro_Ocom ASC, Femi_Ocom ASC "
    '
    Dim OCOMPTEMP As ADODB.Recordset
    Set OCOMPTEMP = READSET(SQLX$)
    With OCOMPTEMP
        FIN& = .RecordCount
        Do While (.EOF = False)
            KU& = KU& + 1
            Call TRACE(20, KU&, FIN&)
            
            FECHEMI% = CVINT(!Femi_Ocom)
            STAT% = XVALO(!Stat_Ocom)
            NPRO% = XVALO(!Npro_Ocom)
            CIXI% = XVALO(!cod_inte)
            DOPE$ = ""
            '
            NOPE% = !Cod_Oper
            If NOPE% > 0 Then
              DOPE$ = DESCRIOP$(CIXI%, NOPE%)
            End If
            NUOC& = XVALO(!nume_ocom)
            '
            UMEDI$ = !Uni_Com
            If TRIM$(UMEDI$) = "" Then
              UMEDI$ = Unimed$(CIXI%)
            End If
            CACOM = !Cant_Ocom
            '
            If STAT% = 8 Then                 ' SI SE DIO POR CUMPLIDA,
              CACOM = !Cant_Rec
            End If
            '
            PREUNI# = !Pre_Unit
            '
            MONEMI% = !Mone_Emis
            If MONEMI% <= 1 Then MONEX$ = "    "
            If MONEMI% = 2 Then MONEX$ = "U$S "
            If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
            TICAMBIO = 1
            If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
            '
            
            Rst.AddNew
            Rst.Fields(0) = NPRO%
            Rst.Fields(1) = CVDAT(FECHEMI%)
            Rst.Fields(2) = !Cod_Exte
            Rst.Fields(3) = Left(!DESCRITEM, 36)
            If !cod_inte <= 0 Then Rst.Fields(3) = Left(!DELIBRE, 36)
            Rst.Fields(4) = Left(DOPE$, 20)
            Rst.Fields(5) = NUOC&
            Rst.Fields(6) = Left(UMEDI$, 4)
            Rst.Fields(7) = CACOM
            Rst.Fields(8) = Left(MONEX$, 4)
            Rst.Fields(9) = PREUNI#
                ITOCOM# = PREUNI# * CACOM
            Rst.Fields(10) = ITOCOM#
            Rst.Fields(11) = PREUNI# * TICAMBIO
            Rst.Fields(12) = ITOCOM# * TICAMBIO
            
            If TRIM$(DEST1$) = "" Then
              CCO% = !IContab
              DEST1$ = CODCUE$(CCO%)
            End If
            '
            If IsNull(!Sect_Ocom) Then
                SECT1$ = ""
            Else
                SECT1$ = !Sect_Ocom
            End If
            Rst.Update
         
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    OCOMPTEMP.Close
    Set OCOMPTEMP = Nothing
    On Error GoTo 0
    '
    Screen.MousePointer = 1
    '
    oReporte.MostrarReporte
    '
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

'*****/////COMANDO OCULTO/////*****
'Se oculto en el frame8 ("Opciones de Consulta")
'el Command9("Listado de Material en Consignación)"
'*****/////**************/////*****

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub

 Sub Command12_Click()
    Command12.Enabled = False
    Call UltimasComprasPorProveedor
'    MODO% = 0
'    Call SELCOMRAFE(OKKX%, MODO%)
'    '
'    If HAYMATERIALIB% = 1 Then
'        If OKKX% = 1 Then Call CONECRUN("*LICOMPL", "Listado de Compras por Proveedor")
'      Else
'        If OKKX% = 1 Then Call CONECRUN("*LICOMPR", "Listado de Compras por Proveedor")
'    End If
'    '
'    Screen.MousePointer = 1
99  Command12.Enabled = True
End Sub

Private Sub Command17_Click()
    '
    Command17.Enabled = False
    Screen.MousePointer = 11
    J& = 0: k& = 0: OCOA& = 0
    '
    For i& = 1 To ULTREG&("OCOMPRA")
         XX$ = REGLEIDO$("OCOMPRA", i&)
         '
         If CVS(Left(XX$, 4)) >= 0.5 Then    ' depura O/C's
             k& = k& + 1                     ' sin numero
             Call GRAREG("OCOMPRA", XX$, k&)
         End If
         '
         ST% = Asc(Mid$(XX$, 62, 1))
         '
         If ST% < 1 Then     '    excluye cerradas
            If CVS(Left$(XX$, 4)) <> OCOA& Then
               NP% = CVI(Mid$(XX$, 7, 2))
               If NP% < 1 Or NP% = 9999 Then
                  J& = J& + 1
                  Call GRAREG("!OCOMPRA", XX$, J&)
                  OCOA& = CVS(Left$(XX$, 4))
               End If
            End If
         End If
    Next i&
    '
    Call SETULTREG("OCOMPRA", k&)
    Call SETULTREG("!OCOMPRA", J&)
    Call CIERRARCH("!OCOMPRA")
    '
    If J& < 1 Then
          Call MENSERR(24, "No hay Pedidos de Suministro\Pendientes que Puedan Asignarse.")
          Screen.MousePointer = 1
          GoTo 99
    End If
    '
    INDI% = 1
    Screen.MousePointer = 1
    '
    XX$ = REGSEL$("!OCOMPRA")
    If Len(XX$) <= 4 Then GoTo 99
    '
    NUOCO& = CVS(Left$(XX$, 4))
    '
    YY$ = REGSEL$("PROVED1")
    If Len(YY$) <= 4 Then GoTo 99
    NP1% = CVI(Left$(YY$, 2))
    PROVEX$ = Mid$(YY$, 3, 30)
    '
    Screen.MousePointer = 11
    JJ& = 0
    FIN& = ULTREG&("OCOMPRA")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("OCOMPRA", U&)
       '
       If CVS(Left$(XX$, 4)) = NUOCO& Then
          CI0% = CVI(Mid$(XX$, 11, 2))
          CAOR = CVS(Mid$(XX$, 13, 4))
          FENTRE% = CVI(Mid$(XX$, 9, 2))
          YY$ = REGBLAN$("CARDETOC")
          Call REPLA(YY$, MKI$(CI0%), 1, 2)
          Call REPLA(YY$, Unimed$(CI0%), 3, 4)
          Call REPLA(YY$, MKS$(CAOR), 47, 4)
          Call REPLA(YY$, MKI$(FENTRE%), 67, 2)
          Call REPLA(YY$, MKS$(U&), 71, 4)
          JJ& = JJ& + 1
          Call GRAREG("!CARDETOC", YY$, JJ&)
       End If
    Next U&
    '
    Call SETULTREG("!CARDETOC", JJ&)
    Call CIERRARCH("!CARDETOC")
    Screen.MousePointer = 1
    '
    VTB% = VENTABU%("CAROCOM2/TITUPAN=P.S.Nro " + TRIM$(Str$(NUOCO&)) + " - " + TRIM$(rasocli$((-1) * NP1%)))
    If VTB% < 0 Then GoTo 99
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("OCOMPRA")
    For U& = 1 To ULTREG&("!CARDETOC")
       YY$ = REGLEIDO$("!CARDETOC", U&)
       CAOR = CVS(Mid$(YY$, 47, 4))
       REOCO& = CVS(Mid$(YY$, 71, 4))
       If REOCO& > 0 Then
          If REOCO& <= FIN& Then
             XX$ = REGLEIDO$("OCOMPRA", REOCO&)
             If CVS(Left$(XX$, 4)) = NUOCO& Then
                Call REPLA(XX$, MKI$(NP1%), 7, 2)
                Call REPLA(XX$, MKS$(CAOR), 13, 4)
                Call GRAREG("OCOMPRA", XX$, REOCO&)
             End If
          End If
       End If
    Next U&
    '
    FIN& = ULTREG&("MACONSIG")
    For U& = 1 To FIN&
       XX$ = REGLEIDO$("MACONSIG", U&)
       If XVALO(Mid$(XX$, 24, 6)) = NUOCO& Then
          Call REPLA(XX$, MKI$(NP1%), 3, 2)
          Call GRAREG("MACONSIG", XX$, U&)
       End If
    Next U&
    '
    Call CIERRARCH("*.*")
    Call FRESHECHO("!OCOMPRA")
    Screen.MousePointer = 1
    '
99  Command17.Enabled = True
    '
End Sub

Private Sub Command15_Click()
   Command15.Enabled = False
   '
10 Call DESHASFECHA(Des%, Has%, PERIO$)
   If PERIO$ = "" Then GoTo 99
   '
   Screen.MousePointer = 11
   Desde = FETEXCOR1$(Des%)
   Hasta = FETEXCOR1$(Has%)
   
   Call ABRECOMPRAS
   'Crear Tabla con Crea campo
   Call CREACAMPO("COMPRAS", "OCOMENTRE", "Nume_Ocom", 4, 4)
   Call CREACAMPO("COMPRAS", "OCOMENTRE", "Npro_Ocom", 3, 0)
   Call CREACAMPO("COMPRAS", "OCOMENTRE", "Cod_Inte", 3, 0)
   Call CREACAMPO("COMPRAS", "OCOMENTRE", "Cant_Ocom", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOMENTRE", "Cant_Rec", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOMENTRE", "Fentre_Sol", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOMENTRE", "Fech_Ing", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOMENTRE", "Porc_Ent_Item", 6, 0)
   Call CREACAMPO("COMPRAS", "OCOMENTRE", "Porc_Ent_Prov", 6, 0)
   Call CREACAMPO("COMPRAS", "OCOMENTRE", "NuOrden", 4, 4)
   '
   'BORRO REGISTROS
   Call BORRAREGISTROS("OCOMENTRE", "COD_INTE > 0 ", REGAF&, "NOMESS")
   '
   CIANT% = 0
   PROVANT% = 0
   NUOR& = 0
   '
   CONDI$ = " Stat_Ocom < 9 "
   CONDI$ = CONDI$ + " AND Femi_Ocom >= '" & Desde & "' "
   CONDI$ = CONDI$ + " AND Femi_Ocom <= '" & Hasta & "' "
   CONDI$ = CONDI$ + " AND Nume_Ocom > 0 "
   FIN& = CantRegistros("OCOMDETA", CONDI$, "NOMESS")
   
   SQLX$ = "SELECT * FROM OCOMDETA"
   SQLX$ = SQLX$ + " WHERE " + CONDI$
   SQLX$ = SQLX$ + " ORDER BY Npro_Ocom,Cod_Inte ASC "
   
   'RECORRO EL DETALLE DE LA ORDEN DE COMPRA
   U& = 0
   Dim OCCUMTEMP As ADODB.Recordset
   Set OCCUMTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With OCCUMTEMP
      Do While Not .EOF
         U& = U& + 1
         Call TRACE(20, U&, FIN&)
         'LEO Y ASIGNO VALORES A LAS VARIABLES
         Ci% = XVALO(!cod_inte)
         NPRO% = XVALO(!Npro_Ocom)
         NORCO& = XVALO(!nume_ocom)
         CANTREC = XVALO(!Cant_Rec)
         CANTOCOM = XVALO(!Cant_Ocom)
         FECHENTSOL% = CVINT(!Fentre_Sol)
         FECHOCOMP% = CVINT(!Femi_Ocom)
             CONDI$ = "NumOC = " & NORCO& & " AND Cod_Inte = " & Ci%
         FECHING% = CVINT(VALOBADA("BOLRECEP", "FechRecep", CONDI$))
         
         NUOR& = NUOR& + 1
         '
         'CREO EL REGISTRO EN TABLA OCOMENTRE
'100      RNDX& = RANDSTRING$(-6)
         '
'         If CantRegistros("OCOMENTRE", "Nume_Ocom = " & RNDX&, "NOMESS") > 0 Then GoTo 100
'         Call CreaRegistro("OCOMENTRE", "Nume_Ocom", RNDX&, "NOMESS")
'         '
'         'GRABO LOS DATOS
'         CONDI$ = "Nume_Ocom = " & RNDX&
'         Call ACTUREGISTRO("OCOMENTRE", "Npro_Ocom", CONDI$, NPRO%, REGAF&, "NOMESS")
'         Call ACTUREGISTRO("OCOMENTRE", "Cod_Inte", CONDI$, CI%, REGAF&, "NOMESS")
'         Call ACTUREGISTRO("OCOMENTRE", "Cant_Ocom", CONDI$, CANTOCOM, REGAF&, "NOMESS")
'         Call ACTUREGISTRO("OCOMENTRE", "Cant_Rec", CONDI$, CANTREC, REGAF&, "NOMESS")
'         Call ACTUREGISTRO("OCOMENTRE", "Fentre_Sol", CONDI$, FETEXCOR1$(FECHENTSOL%), REGAF&, "NOMESS")
'         Call ACTUREGISTRO("OCOMENTRE", "Fech_Ing", CONDI$, FECHING%, REGAF&, "NOMESS")
'         Call ACTUREGISTRO("OCOMENTRE", "NuOrden", CONDI$, NUOR&, REGAF&, "NOMESS")
'         '
'         Call ACTUREGISTRO("OCOMENTRE", "Nume_Ocom", CONDI$, NORCO&, REGAF&, "NOMESS")
         '
          SQLX$ = "SELECT * FROM OCOMENTRE "
          SQLX$ = SQLX$ + " WHERE COD_INTE < 0 "
          Set RSTMP = New ADODB.Recordset
25        On Error Resume Next
          RSTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 25
          End If
          On Error GoTo 0
          '
          With RSTMP
            ' AQUÍ DEVOLUCION DIRECTA DEL MATERIAL RECIBIDO
            .AddNew
            !nume_ocom = NORCO&
            !Npro_Ocom = NPRO%
            !cod_inte = Ci%
            !Cant_Ocom = CANTOCOM
            !Cant_Rec = CANTREC
            !Fentre_Sol = FETEXCOR1$(FECHENTSOL%)
            !Fech_Ing = FETEXCOR1$(FECHING%)
            !NUORDEN = NUOR&
            .Update
         End With
         RSTMP.Close
         Set RS = Nothing
             
          
         'VALIDO SI CAMBIO EL ITEMS EN ESE CASO GRABO EL PORCENTAJE
         If CIANT% > 0 And Ci% <> CIANT% Then
            PORC = 0
            If CANTIDOCOM > 0 Then
              PORC = (CANTIRECITEM / CANTIDOCOM) * 100
            End If
            CONDI1$ = "NuOrden =" & NUOR& - 1
            Call ACTUREGISTRO("OCOMENTRE", "Porc_Ent_Item", CONDI1$, PORC, REGAF&, "NOMESS")
            CANTIDOCOM = 0
            CANTIRECITEM = 0
         End If
         '
         'VALIDO SI CAMBIO EL PROVEEDOR EN ESE CASO GRABO EL PORCENTAJE
         If PROVANT% > 0 And NPRO% <> PROVANT% Then
            PORC = 0
            If CANTIDOCOM1 > 0 Then
              PORC = (CANTIRECITEM1 / CANTIDOCOM1) * 100
            End If
            CONDI2$ = "NuOrden =" & NUOR& - 1
            Call ACTUREGISTRO("OCOMENTRE", "Porc_Ent_Prov", CONDI2$, PORC, REGAF&, "NOMESS")
            CANTIDOCOM1 = 0
            CANTIRECITEM1 = 0
         End If
         '

         'ACUMULO LA CANTIDAD RECIBIDA Y DE LA ORDEN DE COMPRA PARA CALCULO DE PORCENTAJE DE ITEM
         CANTIDOCOM = CANTIDOCOM + CANTOCOM
         CANTIRECITEM = CANTIRECITEM + CANTREC
         '
         'ACUMULO LA CANTIDAD RECIBIDA Y DE LA ORDEN DE COMPRA PARA CALCULO DE PORCENTAJE DE PROVEEDOR
         CANTIDOCOM1 = CANTIDOCOM1 + CANTOCOM
         CANTIRECITEM1 = CANTIRECITEM1 + CANTREC

         'ACTUALIZO EL CODIGO Y PROVEEDOR DE COMPARACION
         CIANT% = Ci%
         PROVANT% = NPRO%
        .MoveNext
      Loop
   End With
   OCCUMTEMP.Close
   Set OCCUMTEMP = Nothing
   '
   'GRABO AL FINAL EN EL ULITMO REGISTRO
   PORC = (CANTIRECITEM / CANTIDOCOM) * 100
   CONDI$ = "NuOrden =" & NUOR&
   Call ACTUREGISTRO("OCOMENTRE", "Porc_Ent_Item", CONDI$, PORC, REGAF&, "NOMESS")
            '
   PORC = (CANTIRECITEM1 / CANTIDOCOM1) * 100
   Call ACTUREGISTRO("OCOMENTRE", "Porc_Ent_Prov", CONDI$, PORC, REGAF&, "NOMESS")
   '
   Screen.MousePointer = 1
   '
   If NUOR& > 0 Then
30   OPX% = COMALTER("Opciones de Listados\\Filtrado por Proveedor\Completo (Todos)")
     If OPX% < 1 Or OPX% > 2 Then GoTo 99
     If OPX% = 1 Then
        Call SELECHO("PROVED1")
        NPRO% = XVALO(TRIM$(VALACT1$("PROVED1")))
        If NPRO% < 1 Then GoTo 30
        Call FINAL("?OCOMENTR1(" & NPRO% & ")")
     Else
       Call FINAL("?OCOMENTRE")
     End If
   Else
     Call COMUNI("No se Encontraron Ordenes de Compras Pendientes en el Periodo Seleccionado")
     GoTo 10
   End If
   '
 
99 Command15.Enabled = True
End Sub




Private Sub Command18_Click()
    '
    Command18.Enabled = False
    'SETUPCOM.Show 1
    Call CONECRUN("FLSETUP/SETUPCOM", "Configuración de Funcionamiento")
    Command18.Enabled = True
    '
End Sub


Private Sub Command2_Click()
    Command2.Enabled = False
    FOAPROC07.Show 1
    Command2.Enabled = True
End Sub

 Sub Command20_Click()
    Command20.Enabled = False
    
    Call UltimasComprasPorMaterial
'    MODO% = 1
'    Call SELCOMRAFE(OKKX%, MODO%)
''    If OKKX% = 1 Then Call CONECRUN("*LICOMAR", "Listado de Compras por Articulo")
'    'HAY QUE DERIVAR PARA CADDA LISTADO SEGUN SI ES MAT. TEXTO LIBRE
'    If HAYMATERIALIB = 1 Then
'        If OKKX% = 1 Then Call CONECRUN("*LICOMAL", "Listado de Compras por Articulo")
'    Else
'        If OKKX% = 1 Then Call CONECRUN("*LICOMAR", "Listado de Compras por Articulo")
'    End If
'    Screen.MousePointer = 1
    Command20.Enabled = True
End Sub

Private Sub Command21_Click()
   Command21.Enabled = False
   Screen.MousePointer = 11
   JJ& = 0
   FIN& = ULTREG&("OCOMPRA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("OCOMPRA", U&)
      '
      ST% = Asc(Mid$(XX$, 62, 1))
      If ST% < 1 Then              '    excluye cerradas
         NP% = CVI(Mid$(XX$, 7, 2))
         If NP% < 1 Or NP% = 9999 Then
            CAPEN = CVS(Mid$(XX$, 13, 4)) - CVS(Mid$(XX$, 58, 4))
            If CAPEN > 0 Then
               JJ& = JJ& + 1
               Call GRAREG("!OCOMPRA", XX$, JJ&)
            End If
         End If
      End If
   Next U&
   '
   Call SETULTREG("!OCOMPRA", JJ&)
   Call CIERRARCH("!OCOMPRA")
   '
   Screen.MousePointer = 1
   If JJ& < 1 Then
      Call MENSERR(24, "No Hay Pedidos de Suministro Pendientes.")
      GoTo 99
   End If
   '
   VTB% = VENTABU%("PEDSUPEN")
   '
99 Command21.Enabled = True
   '
End Sub

 Sub Command22_Click()
    '
    Command22.Enabled = False
    'Cargo la lista con las O/C
    Call APERBASDAT("COMPRAS")
    '
3   Call SELECHO("INDIOCOM/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
    NPX& = XVALO(VALACT1$("INDIOCOM")) ' ACA LEO DEL INDIOCOM
    If NPX& < 1 Then GoTo 99
    '
    Screen.MousePointer = 11
    TIDOCUM$ = "Nota de Pedido"
    EPAC$ = UCase$(EMPREAC$)
    If InStr(EPAC$, "AHP") > 0 Then TIDOCUM$ = "ORDEN DE COMPRA"
    
    If InStr(EPAC$, "MINEPARTS") > 0 Then
       VERSIONX% = XVALO(VALOBADA("OCOMENCA", "OcomRevis", "NUME_OCOM = " & NPX&))
       If VERSIONX% < 1 Then
          Call COMUNI("Atención Debe Refrescar Orden de Compra Desde la Modificación\Para Generar la Imagen Correctamente")
          GoTo 99
       End If
    End If
    
    NUDOCU$ = TRIM$(Str$(NPX&))
    DOCUNU& = 0
    '
    NoPuedeImprimir% = 0
    MENSAJEDOC$ = ""
    CONDI$ = "NUME_OCOM = " & NPX&
    ST% = XVALO(VALOBADA("OCOMENCA", "STAT_OCOM", CONDI$, "NOMESS"))
    If ST% < 0 Then
        NoPuedeImprimir% = 1
        MENSAJEDOC$ = "Imposible Imprimir Esta Orden de Compra, La Misma Está Pendiente de Aprobación"
    End If
    Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$, , , MENSAJEDOC$, NoPuedeImprimir%)
    NoPuedeImprimir% = 0
    MENSAJEDOC$ = ""

    GoTo 3


'''    While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
'''    NUDOCU1$ = " " + NUDOCU$ + " "
'''    NUDOCU2$ = " " + NUDOCU$ + "-"
'''    Screen.MousePointer = 11
'''    '
'''    For UI& = ULTREG&("PDOCLST") To 1 Step -1
'''      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'''      If InStr(XX$, TIDOCUM$) > 4 Then
'''        PPXX% = InStr(XX$, "Nro.")
'''        If PPXX% > 0 Then
'''          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
'''        End If
'''        If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
'''          DOCUNU& = CVS(Left$(XX$, 4))
'''          GoTo 5
'''        End If
'''      End If
'''    Next UI&
'''    Screen.MousePointer = 1
'''    Call COMUNI("Documento no Encontrado")
'''    GoTo 99
'''    '
'''5   Call MUESTRADOC(DOCUNU&)
'''    GoTo 3
    '
99  Command22.Enabled = True
    '
90 End Sub
Sub Command3_Click()
    OPXX% = ALTERNA%("Base de Datos Completa\Ordenes Pendientes")
    If Option1.Value = True Then 'PANTALLA
        Select Case OPXX%
          Case 1
            'Call GENINDOC
            Call COCOMP0 ' SUB DE BdDATOS COMPLET.
            '
'            If HAYMATERIALIB = 1 Then
'                Call COMUNI("La Consulta por Pantalla\Permite Visualizar Solamente\Materiales CODIFICADOS.")
'            End If
'            '
'            Call FINAL("*COCOMP0")
          Case 2
            Call OrdenesComprasPendientes
            'ORD% = 0
            'Call CoCoPen(ORD%)        ' ORDENES PENDIENTES
            '
'            If HAYMATERIALIB = 1 Then
'                Call COMUNI("La Consulta por Pantalla\Permite Visualizar Solamente\Materiales CODIFICADOS.")
'            End If
'            '
'            Call FINAL("*COCOPEN")
          Case Else
            Exit Sub
        End Select
        '
      ElseIf Option2.Value = True Then 'IMPRESORA
        Select Case OPXX%
          Case 1
            Call COCOMP0
            If HAYMATERIALIB% = 1 Then 'SI HAY MAT. TEXTO LIBRE
                Call FINAL("*LICOMALI")
              Else
                HAYOCX% = 0: If EXISTE%(LUDAT$ + "OCOMDETX.RFS") > 0 Then HAYOCX% = 1
                If HAYOCX% < 1 Then
                    Call FINAL("*LICOMP0")
                  Else
                    Call FINAL("*LICOMPX")
                End If
            End If
          Case 2
            Call OrdenesComprasPendientes
'            ORD% = 0
            Call CoCoPen(ORD%)        ' ORDENES PENDIENTES
'            If HAYMATERIALIB% = 1 Then
'                Call FINAL("*LICOPEL")
'            Else
'                Call FINAL("*LICOPEN")
'            End If
          Case Else
            Exit Sub
        End Select
    End If
End Sub

Sub Command4_Click()
    Command4.Enabled = False
    OPX% = ALTERNA%("Listado de Ordenes de Compra\Pendientes por Proveedor")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    If OPX% = 1 Then
      Call OrdenesComprasPorProveedor
      'Call FINAL("?STATOCOM") ' LISTADO COMPLETO CON FILTRO DE FECHAS Y PROVEEDOR
      GoTo 99
    End If

'    ORD% = 1
'    Call CoCoPen(ORD%)    ' SELECCIONAR SOLO PENDIENTES
    If Option1.Value = True Then
        Call OrdenesPendientesPorProveedor
'        '
'        If HAYMATERIALIB = 1 Then
'            Call COMUNI("La Consulta por Pantalla\Permite Visualizar Solamente\Materiales CODIFICADOS.")
'        End If
'        '
'        Call FINAL("*CORCPRO")
      ElseIf Option2.Value = True Then
        Call OrdenesPendientesPorProveedor
'        If HAYMATERIALIB% = 1 Then
'            Call FINAL("*LOCPRLI")
'          Else
'            Call FINAL("*LORCPRO")
'        End If
    End If
99  Command4.Enabled = True
End Sub


Sub Command5_Click()
    Command5.Enabled = False
    'ORD% = 2
    Call CoCoPen(ORD%)    ' SELECCIONAR SOLO PENDIENTES
    If Option1.Value = True Then
        '
        Call OrdenesPendientesPorMaterial
'        If HAYMATERIALIB = 1 Then
'                Call COMUNI("La Consulta por Pantalla\Permite Visualizar Solamente\Materiales CODIFICADOS.")
'        End If
'        '
'        Call FINAL("*CORCMAT")
      ElseIf Option2.Value = True Then
        '
        Call OrdenesPendientesPorMaterial
'        If HAYMATERIALIB = 1 Then
'            Call FINAL("*LOCMATL")
'          Else
'            Call FINAL("*LORCMAT")
'        End If
    End If
    Command5.Enabled = True
End Sub



Private Sub Command50_Click()
    Call CIERRARCH("*.*")
    ACONEC$ = "brecep07/forecep07"
    Call CONECRUN(ACONEC$, "Recpecion de Proveedores")
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    'MUESTRA LAS IMAGNES DE LAS OC SEGUN VERSION
    
    'Cargo la lista con las O/C
    Call APERBASDAT("COMPRAS")
    '
3   Call SELECHO("INDIOCOM/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
    NPX& = XVALO(VALACT1$("INDIOCOM")) ' ACA LEO DEL INDIOCOM
    If NPX& < 1 Then GoTo 99
    '
    NUDOCU$ = TRIM$(Str$(NPX&))

    VERS% = XVALO(InputBox("", "Ingrese Número de Versión", VERS%))
    NUMECOMP$ = ""
    NOCOM$ = SAFETEXT$(NPX&)
    While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
    ITNU$ = TRIM$(Str$(VERS%))
    NOCOM$ = Left$(NOCOM$, 6) + "-" + ITNU$
    NUMECOMP$ = NOCOM$
      
    Screen.MousePointer = 11
    TIDOCUM$ = "Nota de Pedido"
    '
    Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$, NUMECOMP$)
    Screen.MousePointer = 1
    GoTo 3
99  Command6.Enabled = True
End Sub

'Private Sub Command6_Click()
'    Command6.Enabled = False
'    'Cargo la lista con las O/C
'    Call APERBASDAT("COMPRAS")
'    '
'    CONDI$ = " DOPRICOR  like 'RC-%'  GROUP BY DOPRICOR, FECHMOV order by CAST(SUBSTRING(DOPRICOR,4,10) AS INT)"
'    Call CARGALISEL("REMICOSI", "MOVISTO", "DOPRICOR/A12;FECHMOV/D8", CONDI$, "NOMESS")
'    '
'3   Call SELECHO("REMICOSI/Remitos de Consignación")
'    NUBUX$ = TRIM$(VALACT1$("REMICOSI"))
'    NPX& = XVALO(Mid$(VALACT1$("REMICOSI"), 4)) '
'    If NPX& < 1 Then GoTo 99
'    '
'    Screen.MousePointer = 11
'    TIDOCUM$ = "Remito de Traslado"
'    NUDOCU$ = TRIM$(Str$(NPX&))
'    '
'    Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$, NUBUX$)
'    GoTo 3
'99  Command6.Enabled = True
'
'End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
10 OPX% = ALTERNA%("Control de Recepciones\Plan de Compras (MRP)\Detalle Proyecciones de Compras\Detalle de Compras Pendientes\Otros Accesos Directos")
   If OPX% = 1 Then
       Call CONECRUN("BRECEP07/CONRECEP", "Recepcion de Proveedores"): GoTo 10
     ElseIf OPX% = 2 Then
       Call REPORTE_PLACOMP: GoTo 10
     ElseIf OPX% = 3 Then
       Call REPORTE_COBERF: GoTo 10
     ElseIf OPX% = 4 Then
       Call REPORTE_COMPRAPENDI: GoTo 10
     ElseIf OPX% = 5 Then
       Call ACCDIR("OCOMPRA")
   End If
   Command8.Enabled = True
End Sub

Sub REPORTE_COMPRAPENDI()
    Call CREARCAMPOS_ORCOMP_PENDIENTE
    OPXX% = COMALTER%("Listado de Compras Pendientes\\Visualizar Reporte\Actualizar Reporte")
    If OPXX% = 1 Then Call FINAL("?DETACOMPEN")
    If OPXX% = 2 Then Call COMPRAPENDIENTE
End Sub

Sub REPORTE_PLACOMP()
    '
    Call CREARCAMPOS_PLACOMP
    '
10  OPXX% = COMALTER%("Opciones de Consulta\\Visualizar Plan de Compras\Visualizar Solo Reporte\Actualizar Plan de Compras\Cambio de Cantidades\Proveedores Excluidos")
    If OPXX% = 1 Then Call GENERA_INFORMA(2): GoTo 10
    If OPXX% = 2 Then Call FINAL("?PLANCOMPRA"): GoTo 10
    If OPXX% = 3 Then Call GENERA_INFORMA(1): GoTo 10
    If OPXX% = 4 Then
        Call CARGALISEL("PLACOMP", "PLACOMP", "COD_INTE/C16+MASTER/A48;CANTIDAD/F12.2", "COD_INTE>0")
        Call SELECHO("PLACOMP/Necesidades de Compras")
        '
        XX$ = TRIM$(VALACT1$("PLACOMP"))
        If XX$ <> "" Then
            CIXI% = CODINT(XX$)
            CANTI1 = XVALO(VALOBADA("PLACOMP", "CANTIDAD", "COD_INTE=" & CIXI%))
            CANTIT$ = TRIM$(FORMATNUM$(CANTI1, "F12.2"))
            CANTI2 = XVALO(InputBox("", "Cambio de Cantidades", CANTIT$))
            OPX% = COMALTER%("¿Esta Seguro de Cambiar la Cantidad de '" & CANTI1 & "' Por '" & CANTI2 & "'?\\No, Continuar\Si, Cambiar")
            If OPX% = 2 Then
                Call ACTUREGISTRO("PLACOMP", "CANTIDAD", "COD_INTE=" & CIXI%, CANTI2, REGAF&)
                Call COMUNI("Cambio de Cantidades Grabado")
            End If
        End If
        GoTo 10
        '
    End If
    If OPXX% = 5 Then
        Call TRALOCAL("PROVEXC", "")
        '
        VTB% = VENTABU%("PROVEXC/Proveedores Excluidos")
        If VTB% < 0 Then Exit Sub
        '
        Call TRACENTRAL("PROVEXC", "")
        GoTo 10
    End If
    '
End Sub
Sub GENERA_INFORMA(MODO%)
    '
    'VALIDACION QUE HAYA DATOS
    If CantRegistros("OCOEMI", "COD_INTE  > 0", "NOMESS") < 1 Then
       Call COMUNI("Plan de Compras Sin Información")
       Exit Sub
    End If
    '
    If MODO% = 2 Then GoTo 80
    '
    Screen.MousePointer = 11
    'BORRO LA INFO DE PLACOMP
    Call BORRAREGISTROS("PLACOMP", "COD_INTE > 0", REGAF&, "NOMESS")
    '
    Dim RSTMP As ADODB.Recordset
    TOTALX# = 0
    'VERIFICO SI EXISTE ALGUN PROVEEDOR EXCLUIDO
    FIN& = ULTREG&("PROVEXC")
    If FIN& > 0 Then
        CONDI$ = ""
        For II& = 1 To FIN&
            XX$ = REGLEIDO$("PROVEXC", II&)
            NPRO% = CVI(Mid$(XX$, 1, 2))
            CONDI$ = CONDI$ & " AND NPRO_OCOM<>" & NPRO% & " "
        Next II&
    End If
    
    'LEO OCOEMI Y ESCRIBO PLACOMP
    SQLX$ = "SELECT COD_INTE, SUM(CANTIDAD) AS CANTI, "
    SQLX$ = SQLX$ + " NPRO_OCOM, StokAct, COSUNI, RaSo_Prov "
    SQLX$ = SQLX$ + " FROM OCOEMI WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE COD_INTE >  0 "
    SQLX$ = SQLX$ + CONDI$
    SQLX$ = SQLX$ + " GROUP BY COD_INTE, Npro_Ocom,StokAct,COSUNI, RaSo_Prov ORDER BY RaSo_Prov"
    'EL ORDENAMIENTO POR PROVEEDOR SE APLICA PARA QUE DESPUES EL NUMERO DE ORDEN ORDENADO EN PLACOMP
    'REFIERA AL ORDEN POR PROVEEDOR
    Set RS = READSET(SQLX$)
    '
    NUORD% = 0
    With RS
        Do While Not .EOF
          Ci% = XVALO(!cod_inte)
          CANTI = XVALO(!CANTI)
          NPRO% = XVALO(!Npro_Ocom)
          RASOPRO$ = SAFETEXT$(!Raso_Prov)
          TMT% = TIMATE%(Ci%)
          REFETIMA$ = ECOARCH$("TATIMAT", Chr$(TMT%))
          FLIA$ = CODFAMIL$(Ci%)
          REFEFAMI$ = ECOARCH$("TAFAMIL", FLIA$)
          stock = XVALO(!StokAct)
          CONSU = ROUND#(CONSUMOPROMEDIO#(Ci%), 2)
          AUTOPROY = AUTONOMIASM$(Ci%)
          FE% = CVINT(Fentre_Sol)
          COSUNITARIO = ROUND#(XVALO(!COSUNI), 4)
          NUORD% = NUORD% + 1
          '
          SQLX$ = "SELECT * FROM PLACOMP "
          SQLX$ = SQLX$ + " WHERE COD_INTE < 0 "
          Set RSTMP = New ADODB.Recordset
25        On Error Resume Next
          RSTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 25
          End If
          On Error GoTo 0
          '
          With RSTMP
            ' AQUÍ DEVOLUCION DIRECTA DEL MATERIAL RECIBIDO
            .AddNew
            !NUORDEN = NUORD%
            !cod_inte = Ci%
            !Cantidad = CANTI
            !Npro_Ocom = NPRO%
            !RasoProve = Left$(RASOPRO$, 64)
            !TIMATE = TMT%
            !RefeTimate = Left$(REFETIMA$, 32)
            !familia = FLIA$
            !RefeFamilia = Left$(REFEFAMI$, 32)
            !StokAct = stock
            !ConsuProm = CONSU
            !Autono_Proy = AUTONOMIASM$(Ci%)
                FF% = CVINT(VALOBADA("OCOEMI", "FENTRE_SOL", "COD_INTE = " & Ci% & " ORDER BY FENTRE_SOL ASC", "NOMESS"))
                FF% = DIASPOST(FF%, LEAD_TIME%(Ci%) + 5)
            !Fentre_Sol = CVDAT(FF%)
            !COSUNI = COSUNITARIO
            TOTALX# = CANTI * COSUNITARIO
            ACUMU_TOTALX# = ACUMU_TOTALX# + TOTALX#
            ASDF = CODEXT$(Ci%)
            .Update
          End With
          '
        .MoveNext
        Loop
    End With
    On Error Resume Next
    RSTMP.Close
    Set RSTMP = Nothing
    On Error GoTo 0
    '
    'ORDENAR SEGUN FECHA PARA AGREGAR LA PRIORIDAD
    SQLX$ = "SELECT * FROM PLACOMP WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE COD_INTE > 0 "
    SQLX$ = SQLX$ + " ORDER BY FENTRE_SOL ASC"
    Set ORDTMP = READSET(SQLX$)
    With ORDTMP
      Do While Not .EOF
        J& = J& + 1
        Ci% = XVALO(!cod_inte)
        'GRABO CON ACTUREGISTRO POR QUE NO PUEDO ACTUALIZAR
        'UN RECORDSET QUE LA CONSULTA TENGA ORDER BY
        Call ACTUREGISTRO("PLACOMP", "PRIORIDAD", "COD_INTE = " & Ci%, J&, REGAF&, "NOMESS/NOAUDIT")
        .MoveNext
      Loop
    End With
    ORDTMP.Close
    Set ORDTMP = Nothing
    '
    Screen.MousePointer = 1
    '
80  If MODO% = 2 Then
        ACUMU_TOTALX# = 0
        SQLX$ = "SELECT * FROM PLACOMP WITH(NOLOCK) "
        SQLX$ = SQLX$ + " WHERE COD_INTE > 0 "
        Set RSTMP = New ADODB.Recordset
        Set RSTMP = READSET(SQLX$)
        With RSTMP
         Do While Not .EOF
            CANTI = XVALO(!Cantidad)
            COSUNITARIO = XVALO(!COSUNI)
            TOTALX# = CANTI * COSUNITARIO
            ACUMU_TOTALX# = ACUMU_TOTALX# + TOTALX#
           .MoveNext
         Loop
        End With
        RSTMP.Close
        Set RSTMP = Nothing
    End If
    '
    FECHOR$ = Left$(Control$("CANECOM", "FECHOCAL"), 8)
    FECHA% = FECHANUM(FECHOR$)
    FF1% = DIASPOST(FECHA%, 30) ' SE SUMA 30 DIAS PARA CAER EN EL MES SIGUIENTE, PORQUE LA PROGRAMACION ES PARA EL MES SIGUIENTE AL CALCULO
    Mesx$ = METEXTO$(FF1%)
    MESX_NUM$ = Mid$(FECHOR$, 4, 2)
    ANOX$ = Right$(FECHATEXLAR$(FECHA%), 4)
    MESX_ANOX$ = Mesx$
    '
    CONDI$ = "  Stat_Ocom < 8 "
    CONDI$ = CONDI$ + " AND Stat_Ocom <> -1 "
    CONDI$ = CONDI$ + " AND Cant_Rec < Cant_Ocom - 0.005 "
    TOTAL_PENDIENTE# = CALCULA_TOTAL_PENDIENTE#
    '
    FILTX$ = TRIM$(MESX_ANOX$) & ";" & TRIM$(FECHATEX$(FECHA%)) & ";" & TRIM$(MESX_NUM$ & " de " & MESX_ANOX$) & ";" & TRIM$(Mesx$) & ";" & TRIM$(FORMATNUM$(TOTAL_PENDIENTE#, "F15.2")) & ";" & TRIM$(FORMATNUM$(ACUMU_TOTALX#, "F15.2"))
    Call STDFORM("PLCO-GBL", FILTX$)
    '
End Sub
Sub CREARCAMPOS_PLACOMP()
    '
    Call ABRECONEXION
    Call CREACAMPO("COMPRAS", "PLACOMP", "NuOrden", 3, 0, 1)
    Call CREACAMPO("COMPRAS", "PLACOMP", "Cod_Inte", 3, 0, 1)
    Call CREACAMPO("COMPRAS", "PLACOMP", "Cantidad", 6, 4)
    Call CREACAMPO("COMPRAS", "PLACOMP", "Npro_Ocom", 3, 2)
    Call CREACAMPO("COMPRAS", "PLACOMP", "RasoProve", 10, 64)
    Call CREACAMPO("COMPRAS", "PLACOMP", "Timate", 3, 0, 1)
    Call CREACAMPO("COMPRAS", "PLACOMP", "RefeTimate", 10, 32)
    Call CREACAMPO("COMPRAS", "PLACOMP", "Familia", 10, 4)
    Call CREACAMPO("COMPRAS", "PLACOMP", "RefeFamilia", 10, 32)
    Call CREACAMPO("COMPRAS", "PLACOMP", "StokAct", 6, 0)
    Call CREACAMPO("COMPRAS", "PLACOMP", "ConsuProm", 6, 0)
    Call CREACAMPO("COMPRAS", "PLACOMP", "Autono_Proy", 10, 10)
    Call CREACAMPO("COMPRAS", "PLACOMP", "Fentre_Sol", 8, 0)
    Call CREACAMPO("COMPRAS", "PLACOMP", "Prioridad", 3, 0, 1)
    Call CREACAMPO("COMPRAS", "PLACOMP", "Cosuni", 6, 0)
    '
End Sub
Sub CREARCAMPOS_ORCOMP_PENDIENTE()
    '
    Call ABRECONEXION
    Call CREACAMPO("COMPRAS", "DETAPENDIENTE", "Cod_Inte", 3, 0, 1)
    Call CREACAMPO("COMPRAS", "DETAPENDIENTE", "Npro_Ocom", 3, 2)
    Call CREACAMPO("COMPRAS", "DETAPENDIENTE", "Cantidad", 6, 4)
    Call CREACAMPO("COMPRAS", "DETAPENDIENTE", "Fentre_Sol", 8, 0)
    Call CREACAMPO("COMPRAS", "DETAPENDIENTE", "Cosuni", 6, 0)
    Call CREACAMPO("COMPRAS", "DETAPENDIENTE", "MESENT", 3, 0, 1)
    '
End Sub

Sub Command9_Click()
    Command9.Enabled = False
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    TXT1$ = "Opciones de Emisión de Listado:\\Secuencial Completo\Pendientes Por Proveedor\Pendientes Por Artículo\Reservados a Consignar\Cancelar": NN% = 4
    If InStr(EPAC$, "MEDITEA") > 0 Then
        TXT1$ = "Opciones de Emisión de Listado:\\Secuencial Completo\Pendientes Por Proveedor\Pendientes Por Artículo\Reservados a Consignar\Seguimiento de Consignación\Cancelar": NN% = 5
        NN% = 5
    End If
    MODOEMI% = COMALTER%(TXT1$)
    If MODOEMI% < 1 Or MODOEMI% > NN% Then GoTo 99
    If MODOEMI% = 1 Then
         Call LITMAOC(MODOEMI%)
         Call FINAL("*LIMATCO")
       ElseIf MODOEMI% = 2 Then
         Call ConsignacionPendientesPorProveedor
       ElseIf MODOEMI% = 3 Then
         Call ConsignacionPendientesPorArticulo
       ElseIf MODOEMI% = 4 Then 'LISTA AQUELLOS ARTICULOS A CONSIGNAR DONDE
         Call MaterialesPendientesDeConsignar       'LA CANTIDAD ASIGNADA ES MAYOR A LA CANTIDAD ENTREGADA
       ElseIf MODOEMI% = 5 Then
         If InStr(EPAC$, "MEDITEA") > 0 Then
            Call SEGUICONSI
         Else
            Call SaldoConsignadoAFechaValorizado
         End If
    End If
99  Screen.MousePointer = 1
   
   
   
'   'ACA LISTADO DE MAT EN CONSIGNACION
'       EPAC$ = TRIM$(UCase$(EMPREAC$))
'
'       If Option1.Value = True Then
'            TXT$ = "Opciones de Consulta:\\Por Proveedor\Por Artículo\Cancelar": NN% = 2
'            If InStr(EPAC$, "MEDITEA") > 0 Then
'              TXT = "Opciones de Consulta:\\Por Proveedor\Por Artículo\Seguimiento de Consignación\Cancelar"
'              NN% = 3
'            End If
'            MODOEMI% = COMALTER%(TXT$)
'            If MODOEMI% < 1 Or MODOEMI% > NN% Then GoTo 99
'            If MODOEMI% <> 3 Then
'              Screen.MousePointer = 11
'              Call GENEMATCON ' PARA EVITAR LA DEMORA QUE PROVOCA ESTA SUB
'            End If
'            '
'            If MODOEMI% = 1 Then
'                 Call FINAL("*COMATCO")
'               ElseIf MODOEMI% = 2 Then
'                 Call FINAL("*COMACON")
'               ElseIf MODOEMI% = 3 Then
'                 Call SEGUICONSI
'            End If
'            Screen.MousePointer = 1
'            Command9.Enabled = True
'            GoTo 99
'          ElseIf Option2.Value = True Then
'            'TXT1$ = "Opciones de Emisión de Listado:\\Secuencial Completo\Pendientes Por Proveedor\Pendientes Por Artículo\Reservados a Consignar\Consignaciones Valorizadas\Cancelar": NN% = 5
'            TXT1$ = "Opciones de Emisión de Listado:\\Secuencial Completo\Pendientes Por Proveedor\Pendientes Por Artículo\Reservados a Consignar\Cancelar": NN% = 4
'            If InStr(EPAC$, "MEDITEA") > 0 Then
'                TXT1$ = "Opciones de Emisión de Listado:\\Secuencial Completo\Pendientes Por Proveedor\Pendientes Por Artículo\Reservados a Consignar\Seguimiento de Consignación\Cancelar": NN% = 5
'                NN% = 5
'            End If
'            Screen.MousePointer = 11
'            MODOEMI% = COMALTER%(TXT1$)
'            If MODOEMI% < 1 Or MODOEMI% > NN% Then GoTo 99
'            If MODOEMI% = 1 Then
'                 Call LITMAOC(MODOEMI%)
'                 Call FINAL("*LIMATCO")
'               ElseIf MODOEMI% = 2 Or MODOEMI% = 3 Then
'                 'Call LITMAOC(MODOEMI%)
'                 Call GENCONPEN1(MODOEMI%)
'               ElseIf MODOEMI% = 4 Then 'LISTA AQUELLOS ARTICULOS A CONSIGNAR DONDE
'                 Call MATPENDCONS       'LA CANTIDAD ASIGNADA ES MAYOR A LA CANTIDAD ENTREGADA
'               ElseIf MODOEMI% = 5 Then
'                 If InStr(EPAC$, "MEDITEA") > 0 Then
'                    Call SEGUICONSI
'                 Else
'                    Call SaldoConsignadoAFechaValorizado
'                 End If
'            End If
'        End If
'99  Screen.MousePointer = 1
    Command9.Enabled = True
End Sub

Sub SaldoConsignadoAFechaValorizado()

10      Call SELECHO("SELFECHA/CONSIGNADO HASTA EL ...")
        VDEV$ = TRIM$(VALACT1$("SELFECHA"))
        If VDEV$ = "" Then Exit Sub
        FechaCorte$ = FETEXCOR1$(FECHANUM(VDEV$))
        '
        opcionValorizacion% = COMALTER%("Opciones de Valorización:\\Costo de Suministro\Costo Calculado (Análisis de Costos)")
        If opcionValorizacion% < 1 Or opciones% > 2 Then Exit Sub
        
        If opcionValorizacion% = 2 Then
101         Call SELECHO("SELMESA/CALCULO COSTOS")
            PERIODO$ = TRIM$(VALACT1$("SELMESA"))
            If PERIODO$ <> "" Then
                If CantRegistros("HISTOCOSTOS", "PERIODO='" & PERIODO$ & "'") = 0 Then
                    Call COMUNI("No Se Encuentra Información de Costos\para el Período Elegido.")
                    GoTo 101
                End If
                NIVELCO% = COMALTER%("Nivel de Valorización:\\Costo de Suministros\Costo Directo\Costo Standard")
                If NIVELCO% < 1 Or NIVELCO% > 3 Then GoTo 101
                If NIVELCO% = 1 Then
                    NIVELCOSTO% = 0
                ElseIf NIVELCO% = 2 Then
                    NIVELCOSTO% = 1
                ElseIf NIVELCO% = 3 Then
                    NIVELCOSTO% = 2
                End If
            Else
                GoTo 10
            End If
        End If
        
1112    Call CIERRARCH("*.*")
        Call COPYSTRU("BOLREDET", "CONSIVAL")
        Call BLANARCH("CONSIVAL")
        '
        ' *** OBTIENE LAS CONSIGNACIONES REALIZADAS AL PROVEEDOR ELEGIDO Y HASTA LA FECHA DE CORTE
        SQLX$ = "SELECT NPRO_OCOM, COD_INTE, NRO_OCOM, SUM(CANT_ENTREG) AS ENTREGADO, SUM(CANT_RENDID) AS RENDIDO FROM MACONSIG "
        CONDI$ = " Npro_Ocom > 0 "
        If NC% > 0 Then CONDI$ = CONDI$ & " AND Npro_Ocom = " & CStr(NC%)
        If FechaCorte$ <> "" Then CONDI$ = CONDI$ & " AND FEMI_OCOM<='" & FechaCorte$ & "'"
        SQLX$ = SQLX$ & " WHERE " & CONDI$
        SQLX$ = SQLX$ & " GROUP BY COD_INTE, NRO_OCOM, NPRO_OCOM ORDER BY NPRO_OCOM "
        '
        FIN& = CantRegistros("MACONSIG", CONDI$)
        II& = 0
        '
        Dim MACONSITMP As ADODB.Recordset
        Set MACONSITMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        With MACONSITMP
         On Error Resume Next
         .MoveFirst
         If Err Then
            On Error GoTo 0
            GoTo 555
         Else
            On Error GoTo 0
            Screen.MousePointer = 11
            II& = 0
            Do While .EOF = False
              II& = II& + 1
              Call TRACE(24, II&, FIN&)
              '
              Ci% = XVALO(!cod_inte)
              NumeroProveedor% = XVALO(!Npro_Ocom)
              ORDENCOMPRA$ = Mid$(!nro_ocom, 1, 9)
              CondicionSql$ = "CODI_ELEM=" & CStr(Ci%) & " AND NRO_OCOM LIKE '" & ORDENCOMPRA$ & "%'" & " And NU_PROVE = " & CStr(NumeroProveedor%)
              CondicionSql$ = CondicionSql$ & " AND (TIPOMOVIM = 'DV' OR TIPOMOVIM = 'CC' OR TIPOMOVIM = 'QR')"
              If FechaCorte$ <> "" Then CondicionSql$ = CondicionSql$ & " AND FE_RECEP<='" & FechaCorte$ & "'"
              
              cantidadRendida# = XVALO(VALOBADA("RENCONSIG", "SUM(CANT_RENDID)", CondicionSql$))
              CAPENE# = XVALO(!ENTREGADO) - cantidadRendida#
              NUOC& = XVALO(Mid$(!nro_ocom, 4, 6))
              '
              If CAPENE# > 0.05 Then
                Y$ = REGBLAN$("BOLREDET")
                Call REPLA(Y$, MKI$(Ci%), 83, 2)
                Call REPLA(Y$, Unimed$(Ci%), 85, 4)
                Call REPLA(Y$, MKD$(CAPENE#), 89, 8)
                Call REPLA(Y$, MKI$(NumeroProveedor%), 105, 2)
                Call REPLA(Y$, MKS$(NUOC&), 151, 4)
                    If opcionValorizacion% = 1 Then
                        COU = COSUNI(Ci%) * TICAMONE(MONECOSTO%(Ci%))
                    Else
                        COU = COSCALCUHISTO(Ci%, MODOCO%, PERIODO$, NIVELCOSTO%)
                    End If
                Call REPLA(Y$, MKS$(CDbl(COU)), 159, 4)
                Call REPLA(Y$, MKD$(CAPENE# * COU), 171, 8)
                HAYOC% = 1
                Call REGAPP("CONSIVAL", Y$)
              End If
            .MoveNext
            Loop
          End If
        End With
        MACONSITMP.Close
        Set MACONSITMP = Nothing
        '
555     If HAYOC% = 0 Then
            Call COMUNI("No hay Consignaciones Pendientes\para este Proveedor.")
        Else
            Call FILESORT("CONSIVAL", "", 5, 5)
            Call HEADERS("CONSIVAL", "")
            valorizado$ = "Costo Suministro"
            If opcionValorizacion% = 2 Then
                If NIVELCOSTO% = 0 Then valorizado$ = "Costo Calculado - Suministros"
                If NIVELCOSTO% = 1 Then valorizado$ = "Costo Calculado - Directo"
                If NIVELCOSTO% = 2 Then valorizado$ = "Costo Calculado - Standard"
            End If
            Call HEADERS("CONSIVAL", "Valorizacion: " + valorizado$)
            Call FINAL("*CONSIVAL")
        End If
        '
        Screen.MousePointer = 1
        
End Sub

Sub SEGUICONSI()

10  Call SELECHO("INDIOCOM/Indice General de Ordenes de Compra (Pendientes-Cumplidas-Anuladas)")
    NRO& = XVALO(VALACT1$("INDIOCOM")) ' ACA LEO DEL INDIOCOM
    If NRO& < 1 Then Exit Sub
    '
    NOCOM$ = TRIM$(Str$(NRO&))
    While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
    NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-"
    '
    CONDI$ = "Nro_Ocom Like '" & NOCOM$ & "%'"
    If CantRegistros&("MACONSIG", CONDI$, "NOMESS") < 1 Then
       Call COMUNI("No Existen Consignaciones Relacionadas con Esta Orden de Compra")
       GoTo 10
    End If
    '
    Call FORCOMP07.GENERALISTAPICK(NOCOM$)
    '
End Sub

Private Sub ConsignacionPendientesPorProveedor()

   Screen.MousePointer = 11
   
   Call APERBASDAT("COMPRAS")
   
   Dim oReporte As cGenrep
   Dim Rst As ADODB.Recordset
   Dim GECONPENTEMP As ADODB.Recordset
   
   Set oReporte = New cGenrep
   Set Rst = oReporte.GeneraReporteConTablaTemporal("LCONPE2", "CONSIGNACIONES PENDIENTES POR PROVEEDOR", _
             "Nro(ECO=PROVED1-A48)/I10; OCompra(T=O-Compra)/A12; Codigo(ECO=MASTER-A48)/C24; UM/I4; Asignada/F10.1; " _
             & "Entregada/F10.1; FechaE(T=Fecha)/D8; Rendido/F10.1; Saldo/F10.1")
   
   
   Call APERBASDAT("STOCKS")
   SQLX$ = "Select * from MACONSIG WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE ( Cant_Entreg - Cant_Rendid )>0 "
   SQLX$ = SQLX$ + " ORDER BY Npro_Ocom "
   '
   Set GECONPENTEMP = READSET(SQLX$)
   With GECONPENTEMP
        On Error Resume Next
        JJ& = 0: FIN& = GECONPENTEMP.RecordCount
        Do While (.EOF = False)
              JJ& = JJ& + 1
              Call TRACE(24, JJ&, FIN&)
              
              Rst.AddNew
              Rst.Fields(0) = !Npro_Ocom
              Rst.Fields(1) = !nro_ocom
              Rst.Fields(2) = !cod_inte
              Rst.Fields(3) = !Unidad
              Rst.Fields(4) = !cant_asig
              Rst.Fields(5) = !cant_entreg
              Rst.Fields(6) = !FECHA_ENTREG
              Rst.Fields(7) = !CANT_RENDID
                  SALPE# = !cant_entreg - !CANT_RENDID
                  If SALPE# < 0 Then SALPE# = 0
              Rst.Fields(8) = SALPE#
              Rst.Update
              
              .MoveNext
        Loop
   End With
   '
   On Error Resume Next
   GECONPENTEMP.Close
   Set GECONPENTEMP = Nothing
   On Error GoTo 0
   '
   oReporte.MostrarReporte
   '
90 Screen.MousePointer = 1

End Sub

Private Sub ConsignacionPendientesPorArticulo()

   Screen.MousePointer = 11
   
   Call APERBASDAT("COMPRAS")
   
   Dim oReporte As cGenrep
   Dim Rst As ADODB.Recordset
   Dim GECONPENTEMP As ADODB.Recordset
   
   Set oReporte = New cGenrep
   Set Rst = oReporte.GeneraReporteConTablaTemporal("LCONPE3", "CONSIGNACIONES PENDIENTES POR ARTICULO", _
             "Nro(ECO=PROVED1-A48)/I10; OCompra(T=O-Compra)/A12; Codigo(ECO=MASTER-A48)/C24; UM/A4; Asignada/F10.1; " _
             & "Entregada/F10.1; FechaE(T=Fecha)/D8; Rendido/F10.1; Saldo/F10.1")
   
   
   Call APERBASDAT("STOCKS")
   SQLX$ = "Select * from MACONSIG WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE ( Cant_Entreg - Cant_Rendid )>0 "
   SQLX$ = SQLX$ + " ORDER BY Cod_Inte "
   '
   Set GECONPENTEMP = READSET(SQLX$)
   With GECONPENTEMP
        On Error Resume Next
        JJ& = 0: FIN& = GECONPENTEMP.RecordCount
        Do While (.EOF = False)
              JJ& = JJ& + 1
              Call TRACE(24, JJ&, FIN&)
              
              Rst.AddNew
              Rst.Fields(0) = !Npro_Ocom
              Rst.Fields(1) = !nro_ocom
              Rst.Fields(2) = !cod_inte
              Rst.Fields(3) = !Unidad
              Rst.Fields(4) = !cant_asig
              Rst.Fields(5) = !cant_entreg
              Rst.Fields(6) = !FECHA_ENTREG
              Rst.Fields(7) = !CANT_RENDID
                  SALPE# = !cant_entreg - !CANT_RENDID
                  If SALPE# < 0 Then SALPE# = 0
              Rst.Fields(8) = SALPE#
              Rst.Update
              
              .MoveNext
        Loop
   End With
   '
   On Error Resume Next
   GECONPENTEMP.Close
   Set GECONPENTEMP = Nothing
   On Error GoTo 0
   '
   oReporte.MostrarReporte
   '
90 Screen.MousePointer = 1

End Sub

Private Sub MaterialesPendientesDeConsignar()

    Screen.MousePointer = 11
   
    Call APERBASDAT("COMPRAS")
     
    Dim oReporte As cGenrep
    Dim Rst As ADODB.Recordset
    Dim GECONPENTEMP As ADODB.Recordset
    
    Set oReporte = New cGenrep
    Set Rst = oReporte.GeneraReporteConTablaTemporal("MAACONS", "MATERIALES PENDIENTES DE CONSIGNAR", _
              "Nro(ECO=PROVED1-A48)/I10; OCompra(T=O-Compra)/A12; Codigo(ECO=MASTER-A48)/C24; UM/A4; Asignada/F10.1; " _
              & "Entregada/F10.1; Saldo/F10.1")
    
    
    Call APERBASDAT("STOCKS")
    SQLX$ = "Select * from MACONSIG WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE Cant_Asig > Cant_Entreg "
    SQLX$ = SQLX$ + " ORDER BY NRO_OCOM"
    Set GECONPENTEMP = READSET(SQLX$)
    With GECONPENTEMP
       JJ& = 0: FIN& = GECONPENTEMP.RecordCount
       Do While (.EOF = False)
           U2& = U2& + 1
           Call TRACE(20, U2&, FIN&)
           
           Rst.AddNew
           Rst.Fields(0) = !Npro_Ocom
           Rst.Fields(1) = !nro_ocom
           Rst.Fields(2) = !cod_inte
           Rst.Fields(3) = !Unidad
           Rst.Fields(4) = !cant_asig
           Rst.Fields(5) = !cant_entreg
                SALPE# = !cant_asig - !cant_entreg
                If SALPE# < 0 Then SALPE# = 0
           Rst.Fields(6) = SALPE#
           Rst.Update

19         .MoveNext
       Loop
   End With

   On Error Resume Next
   GECONPENTEMP.Close
   Set GECONPENTEMP = Nothing
   On Error GoTo 0
   '
   oReporte.MostrarReporte
   '
90 Screen.MousePointer = 1

End Sub


Sub GENCONPEN1(MODOEMI%)
    Screen.MousePointer = 11
   Call APERBASDAT("STOCKS")
   SQLX$ = "Select * from MACONSIG WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE ( Cant_Entreg - Cant_Rendid )>0 "
   'SQLX$ = SQLX$ + " WHERE (Cant_Asig - Cant_Entreg )>0 "
   '
   Select Case MODOEMI%
    Case 2
        SQLX$ = SQLX$ + " ORDER BY Npro_Ocom "
    Case 3
        SQLX$ = SQLX$ + " ORDER BY Cod_Inte "
   End Select
   '
   'Set GECONPENTEMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
'   Dim GECONPENTEMP As ADODB.Recordset
'   Set GECONPENTEMP = New ADODB.Recordset
'   GECONPENTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   Set GECONPENTEMP = READSET(SQLX$)
   With GECONPENTEMP
   On Error Resume Next
'       .MoveLast
'       'Valida que la tabla no este vacia
'       If Err Then
'         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
'         On Error GoTo 0
'         MODOEMI% = 0
'         GoTo 90
'       End If
'       On Error Resume Next
'   FIN& = .RecordCount
'   .MoveFirst
   JJ& = 0
   Do While (.EOF = False)
         XXY$ = REGBLAN$("LCONPEN")
         Call REPLA(XXY$, MKI$(CVINT(!Femi_Ocom)), 1, 2)
         Call REPLA(XXY$, MKI$(!Npro_Ocom), 3, 2)
         Call REPLA(XXY$, MKD$(!Secc_Ejec), 5, 4)
         If IsNull(!NRO_ORFA) Then
            Call REPLA(XXY$, "", 9, 12)
         Else
            Call REPLA(XXY$, !NRO_ORFA, 9, 12)
         End If
         Call REPLA(XXY$, !nro_ocom, 21, 12)
         Call REPLA(XXY$, MKI$(!cod_inte), 33, 2)
         If IsNull(!LOTE_OCOM) Then
            Call REPLA(XXY$, "", 35, 16)
         Else
            Call REPLA(XXY$, !LOTE_OCOM, 35, 16)
         End If
         Call REPLA(XXY$, !Unidad, 51, 4)
         Call REPLA(XXY$, MKD$(!cant_asig), 55, 8)
         Call REPLA(XXY$, MKD$(!cant_entreg), 63, 8)
         Call REPLA(XXY$, MKI$(CVINT(!FECHA_ENTREG)), 71, 2)
         If IsNull(!NREMI_CONSIG) Then
            Call REPLA(XXY$, "", 73, 6)
         Else
            Call REPLA(XXY$, !NREMI_CONSIG, 73, 6)
         End If
         Call REPLA(XXY$, MKD(!Valo_Unit), 93, 8)
         Call REPLA(XXY$, MKI$(!Moneda), 101, 2)
         Call REPLA(XXY$, MKD$(!Valo_Tot), 103, 8)
         Call REPLA(XXY$, MKD$(!CANT_RENDID), 111, 8)
           SALPE# = !cant_entreg - !CANT_RENDID
           If SALPE# < 0 Then SALPE# = 0
         Call REPLA(XXY$, MKD$(SALPE#), 119, 8)
         If IsNull(!Observa) Then
            Call REPLA(XXY$, "", 129, 16)
         Else
            Call REPLA(XXY$, !Observa, 129, 16)
         End If
         JJ& = JJ& + 1
         Call GRAREG("LCONPEN", XXY$, JJ&)
     .MoveNext
   Loop
  End With
  '
  If JJ& < 1 Then
     Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
     MODOEMI% = 0
     GoTo 90
  End If
  '
  Call SETULTREG("LCONPEN", JJ&)
  Call CIERRARCH("LCONPEN")
  '
  Select Case MODOEMI%
    Case 2
      Call FINAL("*LCONPE2")
    Case 3
      Call FILESORT("LCONPEN", "LCONPEN", 8, 8, "ASC")
      Call FINAL("*LCONPE3")
  End Select
  '
90 Screen.MousePointer = 1
   End Sub

Sub LITMAOC(MODO%)
Screen.MousePointer = 11
   'NUEVO
   Call APERBASDAT("STOCKS")
   SQLX$ = "Select * from MACONSIG with(NOLOCK)"
   '
   Select Case MODO%
    Case 1
        SQLX$ = SQLX$
    Case 2
        SQLX$ = SQLX$ + " ORDER BY Npro_Ocom "
    Case 3
        SQLX$ = SQLX$ + " ORDER BY Cod_Inte "
   End Select
   '
   'Set LITMAOCTEMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
'    Dim LITMAOCTEMP As ADODB.Recordset
'    Set LITMAOCTEMP = New ADODB.Recordset
'    LITMAOCTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText

   Set LITMAOCTEMP = READSET(SQLX$)
   With LITMAOCTEMP
'     On Error Resume Next
'     .MoveLast
'     'Valida que la tabla no este vacia
'     If Err Then
'       Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
'       On Error GoTo 0
'       MODO% = 0
'       GoTo 90
'     End If
'     On Error GoTo 0
'     FIN& = .RecordCount
'     .MoveFirst
     JJ& = 0
     '
     Do While (.EOF = False)
        Select Case MODO%
          Case 1
            XXY$ = REGBLAN$("MACONSIG")
          Case Else
            XXY$ = REGBLAN$("LCONPEN")
        End Select
        '
        If !cod_inte > 0 Then
          Call REPLA(XXY$, MKS$(CVINT(!Femi_Ocom)), 1, 2)
          Call REPLA(XXY$, MKI$(!Npro_Ocom), 3, 2)
              SEJEC# = 0
              On Error Resume Next
              SEJEC# = XVALO(!Secc_Ejec)
              On Error GoTo 0
          Call REPLA(XXY$, MKD$(SEJEC#), 5, 4)
          Call REPLA(XXY$, SAFETEXT$(!NRO_ORFA), 9, 12)
          Call REPLA(XXY$, SAFETEXT$(!nro_ocom), 21, 12)
          Call REPLA(XXY$, MKI$(!cod_inte), 33, 2)
          Call REPLA(XXY$, SAFETEXT$(!LOTE_OCOM), 35, 16)
          Call REPLA(XXY$, SAFETEXT$(!Unidad), 51, 4)
          Call REPLA(XXY$, MKD$(XVALO(!cant_asig)), 55, 8)
          Call REPLA(XXY$, MKD$(XVALO(!cant_entreg)), 63, 8)
          Call REPLA(XXY$, MKI$(CVINT(!FECHA_ENTREG)), 71, 2)
            If IsNull(!NREMI_CONSIG) = False Then Call REPLA(XXY$, !NREMI_CONSIG, 73, 6)
              VUNIT# = 0
              On Error Resume Next
              VUNIT# = XVALO(!Valo_Unit)
              On Error GoTo 0
          Call REPLA(XXY$, MKD(VUNIT#), 93, 8)
              MONEII% = 0
              On Error Resume Next
              MONEII% = XVALO(!Moneda)
              On Error GoTo 0
          Call REPLA(XXY$, MKI$(MONEII%), 101, 2)
              VATOTA# = 0
              On Error Resume Next
              VATOTA# = XVALO(!Valo_Tot)
              On Error GoTo 0
          Call REPLA(XXY$, MKD$(VATOTA#), 103, 8)
              CARENDI# = 0
              On Error Resume Next
              CARENDI# = XVALO(!CANT_RENDID)
              On Error GoTo 0
          Call REPLA(XXY$, MKD$(CARENDI#), 111, 8)
              SALDOPE# = 0
              On Error Resume Next
              SALDOPE# = XVALO(!cant_entreg) - CARENDI#
              If SALDOPE# < 0 Then SALDOPE# = 0
              On Error GoTo 0
          Call REPLA(XXY$, MKD$(SALDOPE#), 119, 8)
          Call REPLA(XXY$, SAFETEXT$(!Observa), 129, 16)
          JJ& = JJ& + 1
          Select Case MODO%
            Case 1
               Call GRAREG("MACONSIG", XXY$, JJ&)
            Case Else
               Call GRAREG("LCONPEN", XXY$, JJ&)
          End Select
        End If
     .MoveNext
     Loop
   End With
   Set LITMAOCTEMP = Nothing
   '
   If JJ& < 1 Then
     Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
     MODO% = 0
     GoTo 90
   End If
   
   Select Case MODO%
     Case 1
        Call SETULTREG("MACONSIG", JJ&)
        Call CIERRARCH("MACONSIG")
     Case Else
        Call SETULTREG("LCONPEN", JJ&)
        Call CIERRARCH("LCONPEN")
   End Select
   '
90 Screen.MousePointer = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Openforms = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call APLICACIONSKINS(Me, Picture2)
    Call GRATITFOR(Caption)
    
    '
End Sub

Sub PEDISUMI()             ' NUEVO PEDIDO DE SUMINISTRO
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
100 Call CIERRARCH("*.*")
    REGXX& = ULTREG&("OCOMPRA")
    NRO& = 1
    If REGXX& > 0 Then
        NRO& = 1 + CVS(Left$(REGLEIDO$("OCOMPRA", REGXX&), 4))
    End If
    If NRO& <= REGXX& Then
        Screen.MousePointer = 11
        For U& = 1 To REGXX&
            X0$ = REGLEIDO$("OCOMPRA", U&)
            If CVS(Left$(X0$, 4)) >= NRO& Then
                NRO& = 1 + CVS(Left$(X0$, 4))
            End If
        Next U&
        Screen.MousePointer = 1
    End If
    '
    X0$ = REGBLAN$("OCOMPRA")
    Call REPLA(X0$, MKS$(NRO&), 1, 4)
    Call REPLA(X0$, MKI$(HO%), 5, 2)
    Call REPLA(X0$, MKI$(HO%), 9, 2)
    Call REPLA(X0$, "S", 57, 1)
    '
101 OBJE$ = "PEDISUMI"
    '
    X1$ = OBJPLA$(OBJE$, X0$)
    If X1$ = "" Then GoTo 299
    '
    X0$ = X1$
    NRO& = CVS(Left$(X1$, 4))
    If NRO& < 1 Then GoTo 101
    '
    FEMI% = CVI(Mid$(X1$, 5, 2))
    If FEMI% < 1 Then
        Call MENSERR(24, "Falta Fecha de Emisión")
        GoTo 101
    End If
    '
    NPX$ = MKI$(0)
    NPRO% = CVI(NPX$)
    '
    Call BLANARCH("!CARDETOC")
    '
102 TITU$ = TRIM$(Str$(NRO&)) + " - Uso Interno"
    VTB% = VENTABU%("CARPESU1/TITUPAN=Pedido de Suministro Nro. " + TITU$)
    If VTB% < 0 Then GoTo 101
    '
    CAICO% = 0
    For J& = 1 To ULTREG&("!CARDETOC")
       X$ = REGLEIDO$("!CARDETOC", J&)
       CI0% = CVI(Left$(X$, 2))
       If CODEXT$(CI0%) = "" Then
          Call MENSERR(24, "Código de Artículo no Válido\en Item Número " + TRIM$(Str$(J&)) + ".")
          GoTo 102
       End If
       CANTU = CVS(Mid$(X$, 47, 4))
       If CANTU < 0.05 Then
          Call MENSERR(24, "Falta Indicar Cantidad en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          GoTo 102
       End If
       CAICO% = J&
       '
       Call REPLA(X$, Unimed$(CI0%), 3, 4)
       If CVI(Mid$(X$, 67, 2)) < HO% Then
          Call REPLA(X$, MKI$(HO%), 67, 2)
       End If
       Call REPLA(X$, "  ", 69, 2)
       '
       Call LEEEXPLO(CODEXT$(CI0%), 1)
       If CAIT% > 0 Then
          Call REPLA(X$, " *", 69, 2)
       End If
       '
       Call GRAREG("!CARDETOC", X$, J&)
       '
    Next J&
    '
    If CAICO% < 1 Then
       Call MENSERR(24, "Falta Ingresar Detalle Ordenado")
       GoTo 102
    End If
    '
103 VTB% = VENTABU%("CARPESU2/TITUPAN=Pedido de Suministro Nro. " + TITU$)
    If VTB% < 0 Then GoTo 102
    '
    TOTANE# = 0
    For J& = 1 To ULTREG&("!CARDETOC")
       X$ = REGLEIDO$("!CARDETOC", J&)
       CI0% = CVI(Left$(X$, 2))
       CANTU = CVS(Mid$(X$, 47, 4))
       FENTRE% = CVI(Mid$(X$, 67, 2))
       '
       If FENTRE% < FEMI% Then
          Call MENSERR(24, "Fecha de Entrega Incoherente\en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          Call REPLA(X0$, MKI$(FEMI%), 9, 2)
          GoTo 103
       End If
       '
    Next J&
    '
    Call CIERRARCH("!CARDETOC")
    '
    Call BLOQARCH("OCOMPRA")
    Call BLOQARCH("MACONSIG")
    '
    For J& = 1 To ULTREG&("!CARDETOC")
       X$ = REGLEIDO$("!CARDETOC", J&)
       CI0% = CVI(Left$(X$, 2))
       CANTU = CVS(Mid$(X$, 47, 4))
       PreUnid# = CVD(Mid$(X$, 51, 8))
       'PORDE = CVS(Mid$(X$, 55, 4))
       FENTRE% = CVI(Mid$(X$, 67, 2))
       CONSIMAT% = 0: If TRIM$(Mid$(X$, 69, 2)) <> "" Then CONSIMAT% = 1
       '
       X2$ = X1$
       Call REPLA(X2$, MKI$(FENTRE%), 9, 2)
       Call REPLA(X2$, MKI$(CI0%), 11, 2)
       Call REPLA(X2$, MKS$(CANTU), 13, 4)
       'Call REPLA(X2$, MKS$(0), 58, 4)
       'Call REPLA(X2$, Chr$(1 + CONSIMAT%), 57, 1) ' normal
       Call REPLA(X1$, Chr$(0), 62, 1)          ' abierta
       '
       Call REGAPP("OCOMPRA", X2$)
       '
       If CONSIMAT% = 1 Then
          '
          NOCOM$ = TRIM$(Str$(NRO&))
          While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
          ITNU$ = TRIM$(Str$(J&))
          While Len(ITNU$) < 2: ITNU$ = "0" + ITNU$: Wend
          NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-" + ITNU$
          '
          Call LEEEXPLO(CODEXT$(CI0%), CANTU)
          For i% = 1 To CAIT%
             YY$ = REGBLAN$("MACONSIG")
             Call REPLA(YY$, MKI$(FEMI%), 1, 2)
             Call REPLA(YY$, NPX$, 3, 2)
             Call REPLA(YY$, NOCOM$, 21, 12)
             '
             CII% = CIJ%(i%)
             CANTI# = CDbl(CANTU * USOI(i%))
             '
             If TRIM$(CODEXT$(CII%)) <> "" Then
               If CANTI# > 0.05 Then
                 Call REPLA(YY$, MKI$(CII%), 33, 2)
                 Call REPLA(YY$, Unimed$(CII%), 51, 4)
                 Call REPLA(YY$, MKD$(CANTI#), 55, 8)
                 Call REGAPP("MACONSIG", YY$)
               End If
             End If
          Next i%
       End If
       '
    Next J&
    '
    Call CIERRARCH("OCOMPRA")
    Call CIERRARCH("MACONSIG")
    '
' preguntar si imprime si/no
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      FECHDOC$ = FETEXTO$(FEMI%)
      NUMEDOC$ = TRIM$(Str$(NRO&))
      While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Pedido de Suministro"
      '
      CODPARAM$(1) = "COND-PAG"
      CODPARAM$(2) = "OBSERVA"
      CODPARAM$(3) = "IMP-NETO"
      CODPARAM$(4) = "POR-IVA"
      CODPARAM$(5) = "IMP-IVA"
      CODPARAM$(6) = "IMP-TOTA"
      CAPARDOC% = 6
      '
      DOCPARAM$(1) = FORPAG$
      DOCPARAM$(2) = OBSE$
      DOCPARAM$(3) = Str$(TOTANE#)
      DOCPARAM$(4) = Str$(ALIVA)
      DOCPARAM$(5) = Str$(TOTIVA#)
      DOCPARAM$(6) = Str$(TOTATO#)
      
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "UNIMED"
      CODTABU1$(4) = "CANTIDAD"
      CODTABU1$(5) = "FECH-VEN"
      CODTABU1$(6) = "REF-MAT1"
      CODTABU1$(7) = "PRE-LIST"
      CODTABU1$(8) = "PORDESCU"
      CODTABU1$(9) = "PRE-NETO"
      CODTABU1$(10) = "NET-ITEM"
      CACOLTAB1% = 10
      '
      DESTIDOC% = NC1%
      CAITAB1% = 0
      CAITCRU% = 0
      '
      For J& = 1 To ULTREG&("!CARDETOC")
         X$ = REGLEIDO$("!CARDETOC", J&)
         CI0% = CVI(Left$(X$, 2))
         CANTU = CVS(Mid$(X$, 47, 4))
         PreUnid# = CVD(Mid$(X$, 51, 8))
         'PORDE = CVS(Mid$(X$, 55, 4))
         FEX = CVI(Mid$(X$, 67, 2))
         CONSIMAT% = 0: If TRIM$(Mid$(X$, 69, 2)) <> "" Then CONSIMAT% = 1
         '
         CAITAB1% = CAITAB1% + 1
         TETABU1$(1, CAITAB1%) = CODEXT$(CI0%)
         TETABU1$(2, CAITAB1%) = DESCRIT0$(CI0%)
         TETABU1$(3, CAITAB1%) = Mid$(X$, 3, 4)
         TETABU1$(4, CAITAB1%) = TRIM$(Str$(CANTU))
         TETABU1$(5, CAITAB1%) = FECHATEX$(FEX)
         TETABU1$(6, CAITAB1%) = Mid$(X$, 69, 2)
         TETABU1$(7, CAITAB1%) = Str$(PreUnid#)
         TETABU1$(8, CAITAB1%) = Str$(PORDE)
         TETABU1$(9, CAITAB1%) = Str$(PreUnid# * (1 - PORDE / 100))
         TETABU1$(10, CAITAB1%) = Str$(CANTU * PreUnid# * (1 - PORDE / 100))
        '
109   Next J&
      '
      Call PRINTDOC("FORMOCOM")   ' Orden de Compra
      '
      For KK1% = 1 To CACOLTAB1%
        For KK2% = 1 To CAITAB1%
          TETABU1$(KK1%, KK2%) = ""
        Next KK2%
      Next KK1%
      '
    End If
    
'consignacion inmediata si/no

    GoTo 100
    '
299 Call CIERRARCH("*.*")
    '
End Sub
'
Sub CoCoPen(ORDEN%)
   '
   If RECLEN("OPENDETA") < 160 Then
     Call COMUNI("ATENCION: Controlador 'OPENDETA.RFS' Desactualizado.\Hay Campos que no Serán Listados.\  \Solicite Actualización S/Cargo a Soporte de Sistemas.")
   End If
   '
   Screen.MousePointer = 11
   'NUEVO 11/03/2005
   'MARCA PARA SABER SI HAY MATERIALES DE TEXTO LIBRE
   HAYMATERIALIB% = 0
   '
   Call APERBASDAT("COMPRAS")
   'SQLX$ = "Select * from OCOMDETA WITH(NOLOCK)"
   SQLX$ = "Select * from OCOMDETA WITH(NOLOCK) "
   SQLX$ = SQLX$ + " Where Stat_Ocom < 3 "
   SQLX$ = SQLX$ + " AND Stat_Ocom <> -1 "
   SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.001 "
   '\\\OT-5-0397-FG-13/06/05///
   Select Case ORDEN%
    Case 0
        SQLX$ = SQLX$ + " ORDER BY Nume_Ocom, Cod_Exte, FEntre_Sol "
    Case 1
        SQLX$ = SQLX$ + " ORDER BY Npro_Ocom, Nume_Ocom, Cod_Exte, FEntre_Sol "
    Case 2
        SQLX$ = SQLX$ + " ORDER BY Cod_Exte, Nume_Ocom, FEntre_Sol "
    End Select
   '\\\OT-5-0397-FG-FIN///
'   Set OCOMTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   Set OCOMTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
'   Dim OCOMTEMP As ADODB.Recordset
'   Set OCOMTEMP = New ADODB.Recordset
'   OCOMTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockOptimistic, adCmdText

   With OCOMTEMP
   On Error Resume Next
       '.MoveLast
       'Valida que la tabla no este vacia
       .MoveFirst
       If Err Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         On Error GoTo 0
         GoTo 90
       End If
   On Error GoTo 0
   .MoveFirst
  JJ& = 0
  TC& = 0
  Do While (.EOF = False)

     XXX$ = REGBLAN$("OPENDETA")
         XXY$ = XXX$ + String$(32, 0)
         MONEMI% = !Mone_Emis
         If MONEMI% <= 1 Then MONEX$ = "   "
         If MONEMI% = 2 Then MONEX$ = "U$S "
         If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
         TICAMBIO = 1
         If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
         '
         Call REPLA(XXY$, MONEX$, 63, 4)
         CAPEN = XVALO(!Cant_Ocom) - XVALO(!Cant_Rec)
         If CAPEN < 0.05 Then CAPEN = 0
         '
'         If InStr(1, EMPREAC$, "GABAL") > 0 Then
'            CARECI0 = XVALO(!CANT_REC)
'            CARECI1 = XVALO(VALOBADA("BOLRECEP", "SUM(CANTOAPRO)", "COD_INTE=" & CStr(XVALO(!cod_inte)) & " AND NUMOC=" & CStr(XVALO(!NUME_OCOM))))
'            If Abs(CARECI0 - CARECI1) > 0.05 Then
'               OPX% = COMALTER%("Se Ha Detectado Una Diferencia en el Material Recibido en OC " & SAFETEXT$(!NUME_OCOM) & "\BR= " & TRIM$(FORMATNUM$(CARECI0, "F12.1")) & _
'               "\OC=" & TRIM$(FORMATNUM$(CARECI1, "F12.1")) & " ¿Desea Corregir?\\No, Continuar\Si, Actualizar")
'               If OPX% = 2 Then
'                   CAPEN = CARECI
'                   !CANT_REC = CARECI
'
'               End If
'            End If
'         End If
         '
         PreUnid# = XVALO(!prun_neto) * TICAMBIO
         Call REPLA(XXY$, MKS$(!nume_ocom), 1, 4) 'AGREGO EL Nº DE ORDEN DE COMPRA
         Call REPLA(XXY$, MKI$(CVINT%(!Femi_Ocom)), 5, 2)
         Call REPLA(XXY$, MKI$(!Npro_Ocom), 7, 2)
         Call REPLA(XXY$, MKI$(!cod_inte), 9, 2)
         Call REPLA(XXY$, MKS$(!Cant_Ocom), 55, 4)
         Call REPLA(XXY$, SAFETEXT$(!Uni_Com), 59, 4)
         Call REPLA(XXY$, MKD$(!prun_neto), 79, 8)
         Call REPLA(XXY$, MKD$(XVALO(!Cant_Ocom) * XVALO(!prun_neto)), 87, 8)
         Call REPLA(XXY$, MKI$(CVINT%(!Fentre_Sol)), 95, 2)
         Call REPLA(XXY$, !CONSIMAT, 97, 2)
         Call REPLA(XXY$, MKS$(XVALO(!Cant_Rec)), 99, 4)
         Call REPLA(XXY$, MKI$(MONEMI%), 121, 2)
         Call REPLA(XXY$, MKS$(CAPEN), 129, 4)
           PreUnid# = XVALO(!prun_neto) * TICAMONE(MONEMI%)
         Call REPLA(XXY$, MKD$(XVALO(PreUnid#)), 133, 8)
         Call REPLA(XXY$, MKD$(XVALO(PreUnid#) * CAPEN), 141, 8)
         '
         If XVALO(!cod_inte) <= 0 Then HAYMATERIALIB% = 1 'HAY MATERIAL TEXTO LIBRE
         '
         JJ& = JJ& + 1
         If !cod_inte > 0 Then
            TC& = TC& + 1
            Call GRAREG("OPENDETA", XXY$, TC&)
         End If
         '
         XXZ$ = REGBLAN("OPENDETL")
         XXZ$ = Left$(XXY$, 10) + Space$(118) + Mid$(XXY$, 11)
         If !cod_inte <= 0 Then
             Call REPLA(XXZ$, MKI$(XVALO(!cod_inte)), 9, 2)
             Call REPLA(XXZ$, SAFETEXT$(!Cod_Exte), 11, 16)
             '\\\OT-05-0690-WAR-21/10/05///
             'Call REPLA(XXZ$, !Descritem, 27, 64)
             Call REPLA(XXZ$, SAFETEXT$(!DELIBRE), 27, 64)
             '\\\OT-05-0690-WAR-FIN///
           Else
             CIXI% = XVALO(!cod_inte)
             Call REPLA(XXZ$, CODEXT$(CIXI%), 11, 16)
             Call REPLA(XXZ$, DESCRIT0$(CIXI%), 27, 64)
         End If
         '
         NOTR$ = SAFETEXT$(!Nro_Otrab)
         NGR$ = SAFETEXT$(!Nro_Grup)
         DEST1$ = NOTR$
         If DEST1$ <> "" Then
           If NGR$ <> "" Then
             DEST1$ = DEST1$ + " / "
           End If
         End If
         DEST1$ = DEST1$ + NGR$
         If TRIM$(DEST1$) = "" Then
           CCO% = XVALO(!IContab)
           DEST1$ = CODCUE$(CCO%)
         End If
         Call REPLA(XXZ$, DEST1$, 91, 24)
         Call REPLA(XXZ$, SAFETEXT$(!Sect_Ocom), 115, 16)
         '
         Call GRAREG("OPENDETL", XXZ$, JJ&)
         '
     '.Update
     .MoveNext
    Loop
   End With
   '
   Call SETULTREG("OPENDETA", TC&)
   Call CIERRARCH("OPENDETA")
   '
   Call SETULTREG("OPENDETL", JJ&)
   Call CIERRARCH("OPENDETL")
   '
90 Screen.MousePointer = 1
End Sub

Sub COCOMP0()
    '
    Screen.MousePointer = 11
    'MARCA PARA SABER SI HAY MATERIALES DE TEXTO LIBRE
    HAYMATERIALIB% = 0
    HAYOCX% = 0: If EXISTE%(LUDAT$ + "OCOMDETX.RFS") > 0 Then HAYOCX% = 1
    'NUEVa subrutina para base de datos completa de OC
    Call APERBASDAT("COMPRAS")
    SQLX$ = "Select * from OCOMDETA WITH(NOLOCK) where Stat_Ocom < 9 AND Nume_Ocom > 0"

    Dim MIRS As ADODB.Recordset
    STRSQL = "Select count(Nume_Ocom) as CdadReg from OCOMDETA WITH(NOLOCK) "
    STRSQL = STRSQL + " where Stat_Ocom < 9"
    Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL))
    If Not IsNull(MIRS!cdadreg) Then
      FIN& = MIRS!cdadreg 'cantidad de registros
    End If
    MIRS.Close
    Set MIRS = Nothing
    '
    Set OCOMP0TEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With OCOMP0TEMP
      On Error Resume Next
      .MoveFirst
      If Err Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         On Error GoTo 0
         JJ& = 0: TC& = 0
         GoTo 90
      End If
      On Error GoTo 0
      '
      Screen.MousePointer = 11
      JJ& = 0: TC& = 0: KU& = 0
      REBLA$ = REGBLAN$("OCOMDETA") + Space$(32)
      Do While (.EOF = False)
        KU& = KU& + 1
        Call TRACE(20, KU&, FIN&)
        XXX$ = REBLA$
        XXY$ = XXX$
        '
        MONEMI% = !Mone_Emis
        If MONEMI% <= 1 Then MONEX$ = "   "
        If MONEMI% = 2 Then MONEX$ = "U$S "
        If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
        TICAMBIO = 1
        If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
        '
        Call REPLA(XXY$, MONEX$, 121, 2)
        CAPEN = XVALO(!Cant_Ocom) - XVALO(!Cant_Rec)
        If CAPEN < 0.05 Then CAPEN = 0
        PreUnid# = XVALO(!prun_neto) * TICAMBIO

        Call REPLA(XXY$, MKS$(XVALO(!nume_ocom)), 1, 4) 'AGREGO EL Nº DE ORDEN DE COMPRA
        Call REPLA(XXY$, MKI$(CVINT%(!Femi_Ocom)), 5, 2)
        Call REPLA(XXY$, MKI$(XVALO(!Npro_Ocom)), 7, 2)
        Call REPLA(XXY$, MKI$(XVALO(!cod_inte)), 9, 2)
        '
        If XVALO(!cod_inte) <= 0 Then HAYMATERIALIB% = 1 'HAY MATERIAL TEXTO LIBRE
        '
        Call REPLA(XXY$, MKS$(XVALO(!Cant_Ocom)), 55, 4)
        If IsNull(!Uni_Com) Then
          UNICOMOC$ = "  "
        Else
          UNICOMOC$ = !Uni_Com
        End If
        Call REPLA(XXY$, UNICOMOC$, 59, 4)
        Call REPLA(XXY$, MKD$(XVALO(!prun_neto)), 79, 8)
        Call REPLA(XXY$, MKD$(XVALO(!Cant_Ocom) * XVALO(!prun_neto)), 87, 8)
        Call REPLA(XXY$, MKI$(CVINT%(!Fentre_Sol)), 95, 2)
        If IsNull(!CONSIMAT) Then
          CONSIMATOC$ = "  "
        Else
          CONSIMATOC$ = !CONSIMAT
        End If
        Call REPLA(XXY$, CONSIMATOC$, 97, 2)
        Call REPLA(XXY$, MKS$(XVALO(!Cant_Rec)), 99, 4)
        Call REPLA(XXY$, MKI$(MONEMI%), 121, 2)
        Call REPLA(XXY$, MKI$(XVALO(!Stat_Ocom)), 124, 1)
        '
        If !cod_inte > 0 Then
           TC& = TC& + 1
           Call GRAREG("OCOMDETA", XXY$, TC&)
           If HAYOCX% > 0 Then
               CUEII% = XVALO(!IContab)
             Call REPLA(XXY$, CODCUE$(CUEII%), 129, 16)
               SECCP$ = SAFETEXT$(!Sect_Ocom)
             Call REPLA(XXY$, SECCP$, 145, 16)
             Call GRAREG("OCOMDETX", XXY$, TC&)
           End If
        End If
        '
        XXZ$ = Left$(XXY$, 10) + Space$(118) + Mid$(XXY$, 11)
        If !cod_inte <= 0 Then
            Call REPLA(XXY$, MKI$(XVALO(!cod_inte)), 9, 2)
            Call REPLA(XXZ$, SAFETEXT$(!Cod_Exte), 11, 16)
            Call REPLA(XXZ$, SAFETEXT$(!DELIBRE), 27, 64)
          Else
            CIXI% = !cod_inte
            Call REPLA(XXZ$, CODEXT$(CIXI%), 11, 16)
            Call REPLA(XXZ$, DESCRIT0$(CIXI%), 27, 64)
        End If
        '
        If IsNull(!Nro_Otrab) Then
           NOTR$ = ""
        Else
           NOTR$ = !Nro_Otrab
        End If
        If IsNull(!Nro_Grup) Then
           NGR$ = ""
        Else
           NGR$ = !Nro_Grup
        End If
        DEST1$ = NOTR$
        If DEST1$ <> "" Then
          If NGR$ <> "" Then
            DEST1$ = DEST1$ + " / "
          End If
        End If
        DEST1$ = DEST1$ + NGR$
        '
        '\\\OT-05-0690-WAR-21/10/05///
         If TRIM$(DEST1$) = "" Then
           CCO% = XVALO(!IContab)
           DEST1$ = CODCUE$(CCO%)
         End If
         If IsNull(!Sect_Ocom) Then
                SECT1$ = ""
           Else
                SECT1$ = !Sect_Ocom
         End If
         Call REPLA(XXZ$, SECT1$, 115, 14)
        '\\\OT-05-0690-WAR-FIN///
        '
        Call REPLA(XXZ$, DEST1$, 91, 24)
        JJ& = JJ& + 1
        Call GRAREG("OCOMDETL", XXZ$, JJ&)
        '
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("OCOMDETA", TC&)
    If HAYOCX% > 0 Then Call SETULTREG("OCOMDETX", TC&)
    '\\\OT-05-0693-WAR-26/10/05///
    Call FILESORT("OCOMDETL", "OCOMDETL", 1, 1, "ASC")
    Call FILESORT("OCOMDETA", "OCOMDETA", 1, 1, "ASC")
    '\\\OT-05-0693-WAR-FIN///
    Call SETULTREG("OCOMDETL", JJ&)
    Call CIERRARCH("OCOMDETA")
    Call CIERRARCH("OCOMDETL")
    '
90  Screen.MousePointer = 1


'    Dim oReporte As cGenrep
'    Dim Rst As ADODB.Recordset
'
'    Screen.MousePointer = 11
'
'    Call APERBASDAT("COMPRAS")
'
'    Set oReporte = New cGenrep
'    Set Rst = oReporte.GeneraReporteConTablaTemporal("COCOMP0", "BASE DE DATOS COMPLETA DE ORDENES DE COMPRA", _
'              "Orden/I10; Fecha/D8; /I10; Codigo/A24; Descripcion/A36; Cantidad/F10.1;" _
'              & " UMCompra(T=UM-Com)/A9; PrecioN(T=Precio Neto)/F12.4; Importe/F12.2; FES(T=F.E.S.)/D8; Recibido/F10.1;" _
'              & " Moneda/I6; Status/I6; Imputacion/A20; Sector/A24")
'
'    SQLX$ = "Select * from OCOMDETA WITH(NOLOCK) where Stat_Ocom < 9 AND Nume_Ocom > 0"
'    '
'    Set OCOMP0TEMP = READSET(SQLX$)
'    With OCOMP0TEMP
'      On Error Resume Next
'      .MoveFirst
'      If Err Then
'         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
'         On Error GoTo 0
'         JJ& = 0: TC& = 0
'         GoTo 90
'      End If
'      On Error GoTo 0
'      '
'      Screen.MousePointer = 11
'      JJ& = 0: TC& = 0: KU& = 0
'      FIN& = OCOMP0TEMP.RecordCount
'      Do While (.EOF = False)
'        KU& = KU& + 1
'        Rst.AddNew
'        Call TRACE(20, KU&, FIN&)
'        '
'        Rst.Fields(0) = !NUME_OCOM
'        Rst.Fields(1) = !FEMI_OCOM
'        Rst.Fields(2) = !NPRO_OCOM
'        Rst.Fields(3) = !Cod_Inte
'        Rst.Fields(4) = !Cant_Ocom
'
'        Rst.Fields(7) = !PrUn_Neto
'        Rst.Fields(8) = XVALO(!Cant_Ocom) * XVALO(!PrUn_Neto)
'        Rst.Fields(9) = !Fentre_Sol
'
''        If IsNull(!CONSIMAT) Then
''          CONSIMATOC$ = "  "
''        Else
''          CONSIMATOC$ = !CONSIMAT
''        End If
'
'        Rst.Fields(10) = !Cant_Rec
'        Rst.Fields(11) = MONEMI%
'        Rst.Fields(12) = !Stat_Ocom
'        '
'        CUEII% = XVALO(!IContab)
'        SECCP$ = SAFETEXT$(!Sect_Ocom)
'        Rst.Fields(13) = CODCUE$(CUEII%)
'        Rst.Fields(14) = SECCP$
'
'        If !Cod_Inte <= 0 Then
'            Rst.Fields(4) = Mid$(!Delibre, 1, 36)
'        End If
'
'        JJ& = JJ& + 1
'        Rst.Update
'        '
'      .MoveNext
'      Loop
'    End With
'    '
'    oReporte.MostrarReporte
'    '
'    '
'90  Screen.MousePointer = 1

End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Sub SELCOMRAFE(OKKX%, MODO%)
   '
   'CODIGO NUEVO AL 11/3/2005 WAR
   'MARCA PARA SABER SI HAY MAT. TEXTO LIBRE
   HAYMATERIALIB% = 0
   '
   Has% = HOY(HOS$)
   FEX = Has%
   HASS$ = FECHATEX$(FEX)
   DESS$ = "01" + Mid$(HASS$, 3)
   Des% = FECHANUM(DESS$)
   VDEF$ = MKI$(Des%) + MKI$(Has%)
   '
10 OKKX% = 0
   Call DESHASFECHA(Des%, Has%, PERIO$)
   If PERIO$ = "" Then Exit Sub
   '
   OKKX% = 1
   Screen.MousePointer = 11
   'Nuevo
   Desde = FECHATEXLAR$(Des%)
   Hasta = FECHATEXLAR$(Has%)
   REBLA$ = REGBLAN$("COPRECOM")
   XXZ$ = REGBLAN("COPRECOL")
   Call HEADERS("COPRECOM", "")
   Call HEADERS("COPRECOM", "Periodo: " + PERIO$)
   '
   Call APERBASDAT("COMPRAS")
   SQLX$ = "SELECT * FROM OCOMDETA"
   SQLX$ = SQLX$ + " WHERE Stat_Ocom < 9 "
   'SQLX$ = SQLX$ + " AND CDBL(Femi_Ocom) >= " & DES1 & " "
   'SQLX$ = SQLX$ + " AND CDBL(Femi_Ocom) <=   " & HAS1 & "  "
   SQLX$ = SQLX$ + " AND Femi_Ocom >= '" & Desde & "' "
   SQLX$ = SQLX$ + " AND Femi_Ocom <= '" & Hasta & "' "
   SQLX$ = SQLX$ + " AND Nume_Ocom > 0 "
   
   If MODO% = 1 Then
        SQLX$ = SQLX$ + "ORDER BY Cod_Exte ASC, Femi_Ocom ASC"
      ElseIf MODO% = 0 Then
        SQLX$ = SQLX$ + "ORDER BY Npro_Ocom ASC, Femi_Ocom ASC "
   End If
   '
    Dim OCOMPTEMP As ADODB.Recordset
    Set OCOMPTEMP = New ADODB.Recordset
    OCOMPTEMP.CursorType = adOpenKeyset
    OCOMPTEMP.LockType = adLockReadOnly
    OCOMPTEMP.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
    With OCOMPTEMP
        JJ& = 0
        TC& = 0
        On Error Resume Next
        .MoveLast
        If Err Then
           On Error GoTo 0
           Call MENSERR(24, "Ningún Registro Seleccionado.")
           GoTo 80
        End If
        On Error GoTo 0
        KU& = 0
        FIN& = .RecordCount
        .MoveFirst
        Do While (.EOF = False)
            KU& = KU& + 1
            Call TRACE(20, KU&, FIN&)
            FECHEMI% = CVINT(!Femi_Ocom)
            STAT% = XVALO(!Stat_Ocom)
            NPRO% = XVALO(!Npro_Ocom)
            CIXI% = XVALO(!cod_inte)
            DOPE$ = ""
            '
            NOPE% = !Cod_Oper
            If NOPE% > 0 Then
              DOPE$ = DESCRIOP$(CIXI%, NOPE%)
            End If
            NUOC& = XVALO(!nume_ocom)
            '
            UMEDI$ = !Uni_Com
            If TRIM$(UMEDI$) = "" Then
              UMEDI$ = Unimed$(CIXI%)
            End If
            CACOM = !Cant_Ocom
            '
            If STAT% = 8 Then                 ' SI SE DIO POR CUMPLIDA,
              CACOM = !Cant_Rec
            End If
            '
            PREUNI# = !Pre_Unit
            '
            MONEMI% = !Mone_Emis
            If MONEMI% <= 1 Then MONEX$ = "    "
            If MONEMI% = 2 Then MONEX$ = "U$S "
            If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
            TICAMBIO = 1
            If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
            '
            YYY$ = REBLA$
            Call REPLA(YYY$, MKI$(CIXI%), 1, 2)
            Call REPLA(YYY$, MKI$(FECHEMI%), 3, 2)
            Call REPLA(YYY$, MKI$(NPRO%), 5, 2)
            Call REPLA(YYY$, MKI$(CIXI%), 7, 2)
            Call REPLA(YYY$, MKI$(FECHEMI%), 9, 2)
            Call REPLA(YYY$, DOPE$, 11, 22)
            Call REPLA(YYY$, MKS$(NUOC&), 33, 4)
            Call REPLA(YYY$, UMEDI$, 37, 4)
            Call REPLA(YYY$, MKS$(CACOM), 41, 4)
            Call REPLA(YYY$, MONEX$, 49, 4)
            Call REPLA(YYY$, MKD$(PREUNI#), 53, 8)
            ITOCOM# = PREUNI# * CACOM
            Call REPLA(YYY$, MKD$(ITOCOM#), 61, 8)
            Call REPLA(YYY$, MKD$(PREUNI# * TICAMBIO), 69, 8)
            Call REPLA(YYY$, MKD$(ITOCOM# * TICAMBIO), 77, 8)
            '
            If !cod_inte <= 0 Then HAYMATERIALIB% = 1
            '
            JJ& = JJ& + 1
            If !cod_inte > 0 Then
                 TC& = TC& + 1
                 Call GRAREG("COPRECOM", YYY$, TC&)
            End If
            '
             XXZ$ = Left$(YYY$, 6) + Space$(122) + Mid$(YYY$, 7)
             If !cod_inte <= 0 Then
                 Call REPLA(XXY$, MKI$(!cod_inte), 1, 2)
                 Call REPLA(XXZ$, !Cod_Exte, 7, 16)
                 '\\\OT-05-0690-WAR-24/10/05///
                 'Call REPLA(XXZ$, !Descritem, 23, 64)
                 Call REPLA(XXZ$, !DELIBRE, 23, 64)
                 '\\\OT-05-0690-WAR-FIN///
               Else
                 CIXI% = !cod_inte
                 Call REPLA(XXZ$, CODEXT$(CIXI%), 7, 16)
                 Call REPLA(XXZ$, DESCRIT0$(CIXI%), 23, 64)
             End If
             '
             If IsNull(!Nro_Otrab) Then
                    NOTR$ = ""
               Else
                    NOTR$ = !Nro_Otrab
             End If
             If IsNull(!Nro_Grup) Then
                    NGR$ = ""
               Else
                    NGR$ = !Nro_Grup
             End If
             DEST1$ = NOTR$
             If DEST1$ <> "" Then
               If NGR$ <> "" Then
                 DEST1$ = DEST1$ + " / "
               End If
             End If
             DEST1$ = DEST1$ + NGR$
             '\\\OT-05-0690-WAR-24/10/05///
             If TRIM$(DEST1$) = "" Then
               CCO% = !IContab
               DEST1$ = CODCUE$(CCO%)
             End If
             '
             If IsNull(!Sect_Ocom) Then
                    SECT1$ = ""
               Else
                    SECT1$ = !Sect_Ocom
             End If
             Call REPLA(XXZ$, SECT1$, 111, 18)
             '\\\OT-05-0690-WAR-FIN///
             Call REPLA(XXZ$, DEST1$, 87, 24)
             '
'                Call REPLA(XXZ$, DOPE$, 133, 22)
'                Call REPLA(XXZ$, MKS$(NUOC&), 155, 4)
'                Call REPLA(XXZ$, UMEDI$, 159, 4)
'                Call REPLA(XXZ$, MKS$(CACOM), 163, 4)
'                Call REPLA(XXZ$, MONEX$, 171, 4)
'                Call REPLA(XXZ$, MKD$(PREUNI#), 175, 8)
'                ITOCOM# = PREUNI# * CACOM
'                Call REPLA(XXZ$, MKD$(ITOCOM#), 183, 8)
'                Call REPLA(XXZ$, MKD$(PREUNI# * TICAMBIO), 191, 8)
'                Call REPLA(XXZ$, MKD$(ITOCOM# * TICAMBIO), 199, 8)
             '
             Call GRAREG("COPRECOL", XXZ$, JJ&)
         .MoveNext
      Loop
    End With
   '
80 Call SETULTREG("COPRECOM", TC&)
   Call CIERRARCH("COPRECOM")
   'Call FINAL("*COPRECOM")
   '
   Call SETULTREG("COPRECOL", JJ&)
   Call CIERRARCH("COPRECOL")
   Screen.MousePointer = 1
   '
'   'CODIGO ANTIGUO
'   HAS% = HOY(HOS$)
'   FEX = HAS%
'   HASS$ = FECHATEX$(FEX)
'   DESS$ = "01" + Mid$(HASS$, 3)
'   DES% = FECHANUM(DESS$)
'   VDEF$ = MKI$(DES%) + MKI$(HAS%)
'   '
'10 OKKX% = 0
'   OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
'   If OBX$ = "" Then
'      Exit Sub
'   End If
'   OKKX% = 1
'   DES% = CVI(Left$(OBX$, 2)): DES1 = CDbl(CVDAT(DES%))
'   HAS% = CVI(Mid$(OBX$, 3, 2)): HAS1 = CDbl(CVDAT(HAS%))
'   If HAS% < DES% Then GoTo 10
'   FEX = DES%: PERIO$ = "Periodo: " + FECHATEX$(FEX)
'   FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
'   '
'   Screen.MousePointer = 11
'   'Nuevo
'   REBLA$ = REGBLAN$("COPRECOM")
'   Call HEADERS("COPRECOM", "")
'   Call HEADERS("COPRECOM", "Periodo: " + PERIO$)
'
'   CALL APERBASDAT("COMPRAS")
'   SQLX$ = "SELECT * FROM OCOMDETA"
'   SQLX$ = SQLX$ + " WHERE Stat_Ocom <> 9 "
'   SQLX$ = SQLX$ + " AND CDBL(Femi_Ocom) >= " & DES1 & " "
'   SQLX$ = SQLX$ + " AND CDBL(Femi_Ocom) <=   " & HAS1 & "  "
'   If MODO% = 1 Then
'        SQLX$ = SQLX$ + " ORDER BY  Cod_Exte "
'      ElseIf MODO% = 0 Then
'        SQLX$ = SQLX$ + " ORDER BY Npro_Ocom  "
'   End If
'
'
'    Set OCOMPTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
'    On Error GoTo 90
'    With OCOMPTEMP
'        .MoveFirst
'        JJ& = 0
'   '
'        Do While (.EOF = False)
'          FECHEMI% = CVINT(!Femi_Ocom)
'            STAT% = !Stat_Ocom
'               NPRO% = !Npro_Ocom
'               CIXI% = !COD_INTE
'               DOPE$ = ""
'               '
'               NOPE% = !Cod_Oper
'               If NOPE% > 0 Then
'                 DOPE$ = DESCRIOP$(CIXI%, NOPE%)
'               End If
'               NUOC& = !Nume_Ocom
'               '
'               UMEDI$ = !Uni_Com
'               If TRIM$(UMEDI$) = "" Then
'                 UMEDI$ = UNIMED$(CIXI%)
'               End If
'               CACOM = !Cant_Ocom
'               '
'               If STAT% = 8 Then                 ' SI SE DIO POR CUMPLIDA,
'                 CACOM = !Cant_Rec
'               End If
'               '
'               PREUNI# = !Pre_Unit
'               '
'               MONEMI% = !Mone_Emis
'               If MONEMI% <= 1 Then MONEX$ = "    "
'               If MONEMI% = 2 Then MONEX$ = "U$S "
'               If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
'               TICAMBIO = 1
'               If MONEMI% >= 2 Then TICAMBIO = TICAMONE(MONEMI%)
'               '
'               YYY$ = REBLA$
'               Call REPLA(YYY$, MKI$(CIXI%), 1, 2)
'               Call REPLA(YYY$, MKI$(FECHEMI%), 3, 2)
'               Call REPLA(YYY$, MKI$(NPRO%), 5, 2)
'               Call REPLA(YYY$, MKI$(CIXI%), 7, 2)
'               Call REPLA(YYY$, MKI$(FECHEMI%), 9, 2)
'               Call REPLA(YYY$, DOPE$, 11, 22)
'               Call REPLA(YYY$, MKS$(NUOC&), 33, 4)
'               Call REPLA(YYY$, UMEDI$, 37, 4)
'               Call REPLA(YYY$, MKS$(CACOM), 41, 4)
'               Call REPLA(YYY$, MONEX$, 49, 4)
'               Call REPLA(YYY$, MKD$(PREUNI#), 53, 8)
'               ITOCOM# = PREUNI# * CACOM
'               Call REPLA(YYY$, MKD$(ITOCOM#), 61, 8)
'               Call REPLA(YYY$, MKD$(PREUNI# * TICAMBIO), 69, 8)
'               Call REPLA(YYY$, MKD$(ITOCOM# * TICAMBIO), 77, 8)
'               JJ& = JJ& + 1
'               Call GRAREG("COPRECOM", YYY$, JJ&)
'         .MoveNext
'      Loop
'    End With
'   '
'   Call SETULTREG("COPRECOM", JJ&)
'   Call CIERRARCH("COPRECOM")
'   Screen.MousePointer = 1
'   '
90 End Sub
'
Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
   Call BAJALDISCO
End Sub
'
Private Sub GENINDOC()
   '
   Call APERBASDAT("COMPRAS")
   SQLX$ = "Select Nume_Ocom,Refe_Ocom  from OCOMENCA WITH(NOLOCK)"
   SQLX$ = SQLX$ & " WHERE Nume_Ocom > 0"
   'Set OCOMP0TEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   Set OCOMP0TEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With OCOMP0TEMP
       On Error Resume Next
       .MoveFirst
       'Valida que la tabla no este vacia
       If Err Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         On Error GoTo 0
         GoTo 90
       End If
       On Error GoTo 0
       JJ& = 0
       Do While (.EOF = False)
         XXX$ = REGBLAN$("OCOMENCA")
         Call REPLA(XXX$, MKS$(!nume_ocom), 1, 4)
         Call REPLA(XXX$, !Refe_OcoM, 5, 44)
         JJ& = JJ& + 1
         Call GRAREG("OCOMENCA", XXX$, JJ&)
         .MoveNext
        Loop
        Call SETULTREG("OCOMENCA", JJ&)
        Call CIERRARCH("OCOMENCA")
    End With
    OCOMP0TEMP.Close
    Set OCOMP0TEMP = Nothing
    
90 End Sub
Sub GENEMATCON()
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = "SELECT * FROM MACONSIG WITH(NOLOCK)"
    SQLX$ = SQLX$ + " ORDER BY Femi_Ocom "
    'Set LIMACONTMP = Stocks.OpenRecordset(SQLX$, 4)
'    Dim LIMACONTMP As ADODB.Recordset
'    Set LIMACONTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    Set LIMACONTMP = READSET(SQLX$)
    With LIMACONTMP
'     On Error Resume Next
'       .MoveFirst
'       'Valida que la tabla no este vacia
'       If Err Then
'         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
'         On Error GoTo 0
'         GoTo 91
'       End If
'       On Error GoTo 0
       JJ& = 0
       Do While (.EOF = False)
         XXX$ = REGBLAN$("MACONSIG")
         On Error Resume Next
         Call REPLA(XXX$, MKS$(CVINT(!Femi_Ocom)), 1, 2)
         Call REPLA(XXX$, MKI$(!Npro_Ocom), 3, 2)
         Call REPLA(XXX$, MKD$(!Secc_Ejec), 5, 4)
         If IsNull(!NRO_ORFA) Then
            Call REPLA(XXX$, "", 9, 12)
         Else
            Call REPLA(XXX$, !NRO_ORFA, 9, 12)
         End If
         Call REPLA(XXX$, !nro_ocom, 21, 12)
         Call REPLA(XXX$, MKI$(!cod_inte), 33, 2)
         If IsNull(!LOTE_OCOM) Then
            Call REPLA(XXX$, "", 35, 16)
         Else
            Call REPLA(XXX$, !LOTE_OCOM, 35, 16)
         End If
         Call REPLA(XXX$, !Unidad, 51, 4)
         Call REPLA(XXX$, MKD$(!cant_asig), 55, 8)
         Call REPLA(XXX$, MKD$(!cant_entreg), 63, 8)
         Call REPLA(XXX$, MKI$(CVINT(!FECHA_ENTREG)), 71, 2)
         If IsNull(!NREMI_CONSIG) Then
            Call REPLA(XXX$, "", 73, 6)
         Else
            Call REPLA(XXX$, !NREMI_CONSIG, 73, 6)
         End If
         Call REPLA(XXX$, MKD(!Valo_Unit), 93, 8)
         Call REPLA(XXX$, MKI$(!Moneda), 101, 2)
         Call REPLA(XXX$, MKD$(!Valo_Tot), 103, 8)
         Call REPLA(XXX$, MKD$(!CANT_RENDID), 111, 8)
           SALPE# = !cant_entreg - !CANT_RENDID
           If SALPE# < 0 Then SALPE# = 0
         Call REPLA(XXX$, MKD$(SALPE#), 119, 8)
         If IsNull(!Observa) Then
            Call REPLA(XXX$, "", 129, 16)
         Else
            Call REPLA(XXX$, !Observa, 129, 16)
         End If
         On Error GoTo 0
         JJ& = JJ& + 1
         Call GRAREG("MACONSIG", XXX$, JJ&)
         .MoveNext
        Loop
        Call SETULTREG("MACONSIG", JJ&)
        Call CIERRARCH("MACONSIG")
    End With
    
    If JJ& < 1 Then
         Call MENSERR(24, "Ningún Registro Seleccionado según Condición SQL")
         GoTo 91
    End If
    '
91 End Sub


VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form SEGUIPRO_TB 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seguimiento y Estadísticas de Producción"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   12375
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   14208
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      BackColor       =   14865344
      TabCaption(0)   =   "CONFIGURACIONES"
      TabPicture(0)   =   "SEGUIPRO_TB.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "CONTROL DE LA PRODUCCIÓN"
      TabPicture(1)   =   "SEGUIPRO_TB.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label9"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label12"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label13"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "GRID2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command9"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command10"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Command13"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "UNIDADES EN FABRICA"
      TabPicture(2)   =   "SEGUIPRO_TB.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command3"
      Tab(2).Control(1)=   "Command11"
      Tab(2).Control(2)=   "GRID3"
      Tab(2).Control(3)=   "Label16"
      Tab(2).Control(4)=   "Label15"
      Tab(2).Control(5)=   "Label14"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "ESTADÍSTICAS"
      TabPicture(3)   =   "SEGUIPRO_TB.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label17"
      Tab(3).Control(1)=   "Label18"
      Tab(3).Control(2)=   "Label19"
      Tab(3).Control(3)=   "Label20"
      Tab(3).Control(4)=   "GRID4"
      Tab(3).Control(5)=   "Command12"
      Tab(3).Control(6)=   "Combo2"
      Tab(3).Control(7)=   "Command2"
      Tab(3).ControlCount=   8
      Begin VB.CommandButton Command13 
         Caption         =   "Todas las Secciones"
         Height          =   495
         Left            =   10380
         TabIndex        =   47
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Todas las Secciones"
         Height          =   495
         Left            =   -64620
         TabIndex        =   46
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Todas las Secciones"
         Height          =   495
         Left            =   -64620
         TabIndex        =   45
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00E0E0E0&
         Height          =   345
         Left            =   -68040
         TabIndex        =   43
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command12 
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
         Left            =   -72360
         TabIndex        =   38
         Top             =   720
         Width           =   175
      End
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
         Left            =   -72360
         TabIndex        =   33
         Top             =   720
         Width           =   175
      End
      Begin VB.CommandButton Command10 
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
         Left            =   9120
         TabIndex        =   26
         Top             =   720
         Width           =   175
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
         Left            =   2640
         TabIndex        =   25
         Top             =   720
         Width           =   175
      End
      Begin VB.Frame Frame2 
         Caption         =   "CANTIDAD DE EMPLEADOS / DIAS LABORABLES"
         Height          =   3255
         Left            =   -74760
         TabIndex        =   16
         Top             =   4560
         Width           =   6255
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1920
            TabIndex        =   21
            Top             =   2640
            Width           =   855
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2880
            TabIndex        =   20
            Top             =   2640
            Width           =   855
         End
         Begin VB.CommandButton Command7 
            Caption         =   "Grabar"
            Height          =   375
            Left            =   4920
            TabIndex        =   19
            Top             =   2640
            Width           =   975
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Actualizar"
            Height          =   375
            Left            =   3840
            TabIndex        =   18
            Top             =   2640
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00E0E0E0&
            Height          =   345
            Left            =   960
            TabIndex        =   17
            Top             =   360
            Width           =   1215
         End
         Begin MSFlexGridLib.MSFlexGrid GRID1 
            Height          =   1605
            Left            =   240
            TabIndex        =   22
            Top             =   780
            Width           =   5715
            _ExtentX        =   10081
            _ExtentY        =   2831
            _Version        =   393216
            Rows            =   1
            Cols            =   3
            FixedRows       =   0
            FixedCols       =   0
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
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Año"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -120
            TabIndex        =   24
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   240
            TabIndex        =   23
            Top             =   2640
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "DATOS DE AUSENCIAS"
         Height          =   3735
         Left            =   -74760
         TabIndex        =   1
         Top             =   600
         Width           =   11655
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
            Height          =   360
            Left            =   2400
            TabIndex        =   8
            Top             =   480
            Width           =   175
         End
         Begin VB.TextBox Text3 
            Appearance      =   0  'Flat
            Height          =   540
            Left            =   3120
            MaxLength       =   64
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   2820
            Width           =   3135
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Height          =   540
            Left            =   6360
            MaxLength       =   64
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   2820
            Width           =   3495
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1680
            TabIndex        =   5
            Top             =   2820
            Width           =   1335
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Actualizar"
            Height          =   255
            Left            =   9960
            TabIndex        =   4
            Top             =   2760
            Width           =   1335
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
            Height          =   360
            Left            =   8640
            TabIndex        =   3
            Top             =   480
            Width           =   175
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Grabar"
            Height          =   375
            Left            =   9960
            TabIndex        =   2
            Top             =   3060
            Width           =   1335
         End
         Begin MSFlexGridLib.MSFlexGrid GRID 
            Height          =   1605
            Left            =   240
            TabIndex        =   9
            Top             =   1020
            Width           =   11115
            _ExtentX        =   19606
            _ExtentY        =   2831
            _Version        =   393216
            Rows            =   1
            Cols            =   4
            FixedRows       =   0
            FixedCols       =   0
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
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1200
            TabIndex        =   15
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Sección"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   555
            Width           =   855
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   240
            TabIndex        =   13
            Top             =   2820
            Width           =   1335
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2760
            TabIndex        =   12
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   6840
            TabIndex        =   11
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2D3C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Mes y Año"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5400
            TabIndex        =   10
            Top             =   555
            Width           =   1215
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GRID2 
         Height          =   6405
         Left            =   360
         TabIndex        =   32
         Top             =   1320
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   11298
         _Version        =   393216
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
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
      Begin MSFlexGridLib.MSFlexGrid GRID3 
         Height          =   6405
         Left            =   -74640
         TabIndex        =   37
         Top             =   1320
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   11298
         _Version        =   393216
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
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
      Begin MSFlexGridLib.MSFlexGrid GRID4 
         Height          =   6405
         Left            =   -74640
         TabIndex        =   42
         Top             =   1320
         Width           =   11355
         _ExtentX        =   20029
         _ExtentY        =   11298
         _Version        =   393216
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
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
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Año"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -69000
         TabIndex        =   44
         Top             =   795
         Width           =   855
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   -72000
         TabIndex        =   41
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Sección"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   40
         Top             =   795
         Width           =   855
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   -73560
         TabIndex        =   39
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   -73560
         TabIndex        =   36
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Sección"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74520
         TabIndex        =   35
         Top             =   795
         Width           =   855
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   -72000
         TabIndex        =   34
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2D3C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Mes y Año"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   31
         Top             =   795
         Width           =   1215
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7320
         TabIndex        =   30
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3000
         TabIndex        =   29
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Sección"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Top             =   795
         Width           =   855
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1440
         TabIndex        =   27
         Top             =   720
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu Inicio 
      Caption         =   "Inicio"
      Begin VB.Menu Cerrar 
         Caption         =   "Cerrar"
      End
   End
   Begin VB.Menu ConsultasListados 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu ExportaExcel 
         Caption         =   "Exportar a Excel"
      End
      Begin VB.Menu Separador1 
         Caption         =   "-"
      End
      Begin VB.Menu ListadosVarios 
         Caption         =   "Listados Varios"
      End
   End
End
Attribute VB_Name = "SEGUIPRO_TB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DC_Text2 As String
Dim LoadFormulario As Integer

Sub CompletaGrillaUnidadesEnFabrica()

    GRID3.Rows = 1
    
    Seccion$ = Label16
        
    Color1& = &H8000000F: Color2& = &HEEEEDD
    ColorActivo& = &HC0E0DD
    
    On Error GoTo ErrorGrilla
    
    CantidadSecciones& = 1
    Call EncabezadoGrillaUnidades(GRID3.Rows - 1)
    
    If Seccion$ = "" Then
        CantidadSecciones& = ULTREG&("SECPROD")
        If CantidadSecciones& > 1 Then
            GRID3.Rows = 0: GRID3.Rows = 1
        End If
    End If
    
    FilaDetalle% = 0
    For IdSeccion& = 1 To CantidadSecciones&
        
        If CantidadSecciones& > 1 Then
            Seccion$ = TRIM$(FORMATNUM$(CVS(REGLEIDO$("SECPROD", IdSeccion&)), "F10.4"))
            DenominacionSeccion$ = TRIM$(Mid$(REGLEIDO$("SECPROD", IdSeccion&), 5))
            GRID3.TextMatrix(GRID3.Rows - 1, 0) = UCase(DenominacionSeccion$)
            
            RangoFila$ = CStr(GRID3.Rows - 1) & ":" & CStr(GRID3.Rows - 1)
            Call DisenioDeCelda(GRID3, RangoFila$, "0:0", "", 0, True, False, True, 1, vbBlack, Color1&)
            
            GRID3.Rows = GRID3.Rows + 2
            Call EncabezadoGrillaUnidades(GRID3.Rows - 1)
            FilaDetalle% = GRID3.Rows - 1
        End If
        
        ' OBTIENE LOS DATOS IDENTIFICATORIOS DE LAS O-FABRICACION DE LA SECCION ELEGIDA,
        ' QUE TIENEN FECHA DE INICIO DE PROCESO PERO NO DE CIERRE DE OF
        FechaNula% = FECHANUM("01/01/80")
        SQLX$ = "SELECT IdSubOr, NumeOrfa, NumeroChasis, TipoChasis, Destino, LineaInterno FROM ORDEFABR WITH(NOLOCK) WHERE SecciPro='" & Seccion$ & "' AND inicio_informado > '" & FETEXCOR1$(FechaNula%) & "' AND Status_Proceso = 1 AND StatuOrf < 3"
        Dim ORDEFA As ADODB.Recordset
        Set ORDEFA = READSET(SQLX$)
        
        With ORDEFA
            Do While Not .EOF
                ' COMPLETA LA GRILLA
                Orden$ = SAFETEXT$(!IdSubOr)
                Chasis$ = TRIM$(SAFETEXT$(!NumeroChasis))
                Tipo$ = TRIM$(SAFETEXT$(!TipoChasis))
                DestinoFabricacion$ = TRIM$(SAFETEXT$(!DESTINO))
                Linea$ = ECOARCH$("DESTFAB", DestinoFabricacion$)
                
                Interno$ = TRIM$(SAFETEXT$(!LineaInterno))
            
                GRID3.Rows = GRID3.Rows + 1
                GRID3.TextMatrix(GRID3.Rows - 1, 0) = Orden$
                GRID3.TextMatrix(GRID3.Rows - 1, 1) = Chasis$
                GRID3.TextMatrix(GRID3.Rows - 1, 2) = Tipo$
                GRID3.TextMatrix(GRID3.Rows - 1, 3) = Linea$
                GRID3.TextMatrix(GRID3.Rows - 1, 4) = Interno$
                    
                ' COLOREA LA FILA
                If ColorActivo& = Color1& Then
                    ColorActivo& = Color2&
                Else
                    ColorActivo& = Color1&
                End If
                RangoFila$ = CStr(GRID3.Rows - 1) & ":" & CStr(GRID3.Rows - 1)
                Call DisenioDeCelda(GRID3, RangoFila$, "0:4", "", 0, False, False, False, 1, vbBlack, ColorActivo&)
                
                .MoveNext
            Loop
        End With
    
        ORDEFA.Close
        Set ORDEFA = Nothing
        
        If CantidadSecciones& > 1 Then
            GRID3.Rows = GRID3.Rows + 2
        End If
        
    Next IdSeccion&
    
    On Error GoTo 0
    Exit Sub
    
ErrorGrilla:
    Call MENSERR(24, "Se Ha Detectado un Error al Escribar la Grilla de Ausencias\ \Consulte a su Soporte de Sistemas\ \" & Err.Description)
    On Error GoTo 0
    
End Sub

Function ExcelConvierteColumnaNumeroALetra(NumeroColumna As Integer) As String

    ' CONVIERTE EL NUMERO DE COLUMNA EN SU LETRA RESPECTIVA
    ' PARA ESO UTILIZA LA TABLA ASCII QUE DESDE LA POSICION 65 COMIENZA LAS LETRAS MAYUSCULAS
    Letra$ = Chr$(65 + (NumeroColumna - 1))
    ExcelConvierteColumnaNumeroALetra = Letra$
    
End Function

Private Sub Cerrar_Click()

    Unload Me
    
End Sub

Private Sub Combo1_Click()
    
    ' SI CAMBIA EL AÑO ACTUALIZA LA GRILLA DE EMPLEADOS
    If Combo1.TEXT <> "" Then Call CompletaGrillaEmpleados

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)

    ' NO PERMITE MODIFICAR EL COMBO DE AÑOS. INHIBE CUALQUIER DATO INTRODUCIDO
    KeyAscii = 0
    
End Sub

Private Sub Combo2_Click()

    ' SI CAMBIA EL AÑO ACTUALIZA LA GRILLA DE ESTADISTICAS
    Seccion$ = Label17
    
    If Combo2.TEXT <> "" And LoadFormulario = 0 And Seccion$ <> "" Then Call CompletaGrillaEstadisticas

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)

    ' NO PERMITE MODIFICAR EL COMBO DE AÑOS. INHIBE CUALQUIER DATO INTRODUCIDO
    KeyAscii = 0
    
End Sub

Private Sub Command1_Click()
    
    ' SELECCIONA SECTOR PRODUCTIVO
    Call SELECHO("SECPROD")
    SECTOR$ = TRIM$(VALACT1$("SECPROD"))
        
    Label2.Caption = SECTOR$
    Label4.Caption = TRIM$(VALACT2$("SECPROD"))
    
End Sub

Private Sub Command10_Click()
   
    ' SELECCIONA PERIODO MES + AÑO
    Call SELECHO("SELMESA")
    PERIO$ = TRIM$(VALACT1$("SELMESA"))
      
    If PERIO$ <> "" Then Label12 = PERIO$
    
End Sub

Private Sub Command11_Click()

    ' SELECCIONA SECTOR PRODUCTIVO
    Call SELECHO("SECPROD")
    SECTOR$ = TRIM$(VALACT1$("SECPROD"))
    
    Label16.Caption = SECTOR$
    Label14.Caption = TRIM$(VALACT2$("SECPROD"))

End Sub

Private Sub Command12_Click()

    ' SELECCIONA SECTOR PRODUCTIVO
    Call SELECHO("SECPROD")
    Seccion$ = TRIM$(VALACT1$("SECPROD"))
    
    If Seccion$ <> "" Then
        Label17.Caption = Seccion$
        Label19.Caption = TRIM$(VALACT2$("SECPROD"))
    End If
    
End Sub

Private Sub Command13_Click()

    PERIODO$ = Label12.Caption
    If PERIODO$ <> "" Then
        Label9.Caption = "": Label11.Caption = ""
        CompletaGrillaControlProduccion
    Else
        Call COMUNI("Falta Elegir el Período a Consultar")
    End If
    
End Sub

Private Sub Command2_Click()
    
    Anio$ = Combo2.TEXT
    If Anio$ <> "" Then
        Label17.Caption = "": Label19.Caption = ""
        CompletaGrillaEstadisticas
    Else
        Call COMUNI("Falta Elegir el Año a Consultar")
    End If
    
End Sub

Private Sub Command3_Click()

    Label16.Caption = "": Label14.Caption = ""
    Call CompletaGrillaUnidadesEnFabrica
    
End Sub


Private Sub Command4_Click()

    ' ACTUALIZA LA GRILLA DE DATOS DE AUSENCIA
    If GRID.Row > 0 Then
    
        HorasAus$ = TRIM$(Text2)
        Motivo$ = TRIM$(Text3)
        Observacion$ = TRIM$(Text4)
        
        FilaGrid% = GRID.Row
        GRID.TextMatrix(FilaGrid%, 1) = HorasAus$
        GRID.TextMatrix(FilaGrid%, 2) = Motivo$
        GRID.TextMatrix(FilaGrid%, 3) = Observacion$
        
    End If
    
End Sub

Private Sub Command5_Click()
   
    ' SELECCIONA PERIODO: MES + AÑO
    Call SELECHO("SELMESA")
    PERIODO$ = TRIM$(VALACT1$("SELMESA"))
        
    Label5 = PERIODO$

End Sub

Private Sub Command6_Click()

    ' GRABA LOS DATOS DE AUSENCIA EN LA TABLA SECCIFAB
    If GRID.Rows > 1 Then
        On Error GoTo ErrorGrilla
        
        Seccion$ = Label2.Caption
        PERIODO$ = Label5.Caption
        FechaNula% = FECHANUM("01/01/80")
        ' MARCA LOS REGISTROS EXISTENTE PARA LUEGO SER ELIMINADOS
        CONDI$ = "seccion='" & Seccion$ & "' AND periodo='" & PERIODO$ & "' AND fechaseteo <> '" & FETEXCOR1$(FechaNula%) & "'"
        Call ACTUREGISTRO("SECCIFAB", "MARBORRA", CONDI$, 1, REGAF&)
        '
        ' GRABA LOS NUEVOS REGISTROS
        Dim SECCIFAB As ADODB.Recordset
        Set SECCIFAB = New ADODB.Recordset
        SQLX$ = "SELECT * FROM SECCIFAB WHERE fechaseteo='" & FETEXCOR1$(FechaNula%) & "'"
        SECCIFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With SECCIFAB
            For FilaGrid% = 1 To GRID.Rows - 1
                FechaSet$ = GRID.TextMatrix(FilaGrid%, 0)
                HorasAus$ = GRID.TextMatrix(FilaGrid%, 1)
                Motivo$ = GRID.TextMatrix(FilaGrid%, 2)
                Observacion$ = GRID.TextMatrix(FilaGrid%, 3)
                
                .AddNew
                
                !Seccion = XVALO(Seccion$)
                !FechaSeteo = CVDAT(FECHANUM(FechaSet$))
                !PERIODO = PERIODO$
                !HorasAusencia = XVALO(HorasAus$)
                !MotivoAusencia = Motivo$
                !Observaciones = Observacion$
                !MarBorra = 0
            
                .Update
            Next FilaGrid%
        End With
        SECCIFAB.Close
        Set SECCIFAB = Nothing
              
        ' ELIMINA LOS REGISTROS MARCADO AL COMIENZO
        CONDI$ = "seccion='" & Seccion$ & "' AND periodo='" & PERIODO$ & "' AND marborra = 1"
        Call BORRAREGISTROS("SECCIFAB", CONDI$, REGAF&)
        
        On Error GoTo 0
        
        Call COMUNI("Grabación Finalizada")
    Else
        Call COMUNI("Imposible Grabar.\No Existe Una Sección Activa.")
    End If
    
    Exit Sub
    
ErrorGrilla:
    Call MENSERR(24, "Se Ha Detectado un Error al Grabar la Grilla de Datos de Ausencias\ \Consulte a su Soporte de Sistemas\ \" & Err.Description)
    On Error GoTo 0
    On Error Resume Next
    SECCIFAB.Close
    Set SECCIFAB = Nothing
    On Error GoTo 0
    
End Sub

Private Sub Command7_Click()

    ' GRABA LOS DATOS DE CANT DE EMPLEADOS Y DIAS LABORABLES EN LA TABLA SECCIFAB
    ' LA FECHASETEO CON VALOR '01/01/80' INDICA QUE ES UN REGISTRO QUE CONTIENE ESTOS DATOS
    If GRID1.Rows > 1 And Label2 <> "" Then
        On Error GoTo ErrorGrilla
        
        Seccion$ = Label2.Caption
        PERIODO$ = Mid$(TRIM$(Combo1.TEXT), 3, 2)
        FechaNula% = FECHANUM("01/01/80")
        ' MARCA LOS REGISTROS EXISTENTE PARA LUEGO SER ELIMINADOS
        CONDI$ = "seccion='" & Seccion$ & "' AND periodo LIKE '%" & PERIODO$ & "%' AND fechaseteo = '" & FETEXCOR1$(FechaNula%) & "'"
        Call ACTUREGISTRO("SECCIFAB", "MARBORRA", CONDI$, 1, REGAF&)
        '
        ' GRABA LOS NUEVOS REGISTROS
        Dim SECCIFAB As ADODB.Recordset
        Set SECCIFAB = New ADODB.Recordset
        SQLX$ = "SELECT * FROM SECCIFAB WHERE fechaseteo='" & FETEXCOR1$(FechaNula%) & "'"
        SECCIFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With SECCIFAB
            For FilaGrid% = 1 To GRID1.Rows - 1
                Meses$ = TRIM$(GRID1.TextMatrix(FilaGrid%, 0))
                CantEmpleados% = XVALO(GRID1.TextMatrix(FilaGrid%, 1))
                DiasLabora% = XVALO(GRID1.TextMatrix(FilaGrid%, 2))
                
                .AddNew
                
                !Seccion = XVALO(Seccion$)
                !FechaSeteo = CVDAT(FechaNula%)
                !PERIODO = Meses$
                !HorasAusencia = 0
                !MotivoAusencia = ""
                !Observaciones = ""
                !CantidadEmpleados = CantEmpleados%
                !DiasLaborables = DiasLabora%
                !NumOrden = FilaGrid%
                !MarBorra = 0
            
                .Update
            Next FilaGrid%
        End With
        SECCIFAB.Close
        Set SECCIFAB = Nothing
        
        ' ELIMINA LOS REGISTROS MARCADO AL COMIENZO
        CONDI$ = "seccion='" & Seccion$ & "' AND periodo LIKE '%" & PERIODO$ & "%' AND marborra = 1"
        Call BORRAREGISTROS("SECCIFAB", CONDI$, REGAF&)
        
        On Error GoTo 0
        
        Call COMUNI("Grabación Finalizada")
    Else
        Call COMUNI("Imposible Grabar.\Falta Sección Activa y/o Año.")
    End If

    Exit Sub
    
ErrorGrilla:
    Call MENSERR(24, "Se Ha Detectado un Error al Grabar la Grilla de Cantidad de Empleados\ \Consulte a su Soporte de Sistemas\ \" & Err.Description)
    On Error GoTo 0
    On Error Resume Next
    SECCIFAB.Close
    Set SECCIFAB = Nothing
    On Error GoTo 0
    
End Sub

Private Sub Command8_Click()
    
    ' ACTUALIZA LA GRILLA DE DATOS DE CANT EMPLEADOS
    If GRID1.Row > 0 Then
    
        CantEmpleados$ = TRIM$(Text1)
        DiasLabora$ = TRIM$(Text5)
        
        FilaGrid% = GRID1.Row
        GRID1.TextMatrix(FilaGrid%, 1) = CantEmpleados$
        GRID1.TextMatrix(FilaGrid%, 2) = DiasLabora$
        
    End If
    
End Sub

Private Sub Command9_Click()

    ' SELECCIONA SECTOR PRODUCTIVO
    Call SELECHO("SECPROD")
    SECTOR$ = TRIM$(VALACT1$("SECPROD"))
    
    If SECTOR$ = "" Then Exit Sub
    
    Label9.Caption = SECTOR$
    Label11.Caption = TRIM$(VALACT2$("SECPROD"))
    
End Sub

Private Sub ExportaExcel_Click()

    ' EXPORTA CONSULTAS A EXCEL
    Dim NombreDelArchivo As String
    Dim Grilla As MSFlexGrid

    Solapa% = SSTab1.Tab
    TituloExcel$ = ""
    Select Case Solapa%
         Case 0: Set Grilla = GRID
         Case 1: Seccion$ = UCase(Label11)
                 If Seccion$ = "" Then Seccion$ = "Todas las Secciones"
                 TituloExcel$ = "Control de Producción | " & Seccion$: Set Grilla = GRID2
         
         Case 2: Seccion$ = UCase(Label14)
                 If Seccion$ = "" Then Seccion$ = "Todas las Secciones"
                 TituloExcel$ = "Unidades en Fábrica | " & Seccion$: Set Grilla = GRID3
         
         Case 3: Seccion$ = UCase(Label19)
                 If Seccion$ = "" Then Seccion$ = "Todas las Secciones"
                 TituloExcel$ = "Estadísticas | " & Seccion$: Set Grilla = GRID4
    End Select
    
    If Solapa% > 0 And Grilla.Rows > 1 Then
        CommonDialog1.CancelError = True
        CommonDialog1.DialogTitle = "Exportación de Informe"
        CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
        CommonDialog1.FilterIndex = 1
        On Error Resume Next
        CommonDialog1.ShowOpen
        On Error GoTo 0
        '
        Openforms = DoEvents
        NombreDelArchivo = UCase$(CommonDialog1.FileName)
        
        If TRIM$(NombreDelArchivo) <> "" Then
            If UCase(Mid$(NombreDelArchivo, Len(TRIM$(NombreDelArchivo)) - 3)) <> ".XLS" Then
                NombreDelArchivo = NombreDelArchivo & ".xls"
            End If
            Call GeneraExcelDeGrilla(NombreDelArchivo, Grilla, TituloExcel$)
        End If
    End If

End Sub

Private Sub Form_Load()

    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    Call ABRECONEXION
    '
    Call ABREORFAB
    '
    LoadFormulario = 1
    '
    ' DIBUJA TODAS LAS GRILLAS
    Call EncabezadoDeGrillas
    '
    ' COMPLETA EL CAMBO DE AÑOS EN CONFIGURACION Y ACTIVA EL AÑO EN CURSO
    Combo1.AddItem "2015"
    Combo1.AddItem "2016"
    Combo1.AddItem "2017"
    Combo1.AddItem "2018"
    Combo1.AddItem "2019"
    Combo1.AddItem "2020"
    Combo1.AddItem "2021"
    Combo1.AddItem "2022"
    Combo1.AddItem "2023"
    Combo1.AddItem "2024"
    Combo1.AddItem "2025"
    
    AnioHoy$ = XVALO(Mid$(FECHATEXLAR$(HOY(HO$)), 7, 4))
    
    For Item% = 0 To Combo1.ListCount - 1
        If AnioHoy$ = Combo1.LIST(Item%) Then
            Combo1.ListIndex = Item%
            Combo1.TEXT = "": Combo1.TEXT = AnioHoy$
            Exit For
        End If
    Next Item%
    
    ' COMPLETA EL CAMBO DE AÑOS EN ESTADISTICAS Y ACTIVA EL AÑO EN CURSO
    Combo2.AddItem "2015"
    Combo2.AddItem "2016"
    Combo2.AddItem "2017"
    Combo2.AddItem "2018"
    Combo2.AddItem "2019"
    Combo2.AddItem "2020"
    Combo2.AddItem "2021"
    Combo2.AddItem "2022"
    Combo2.AddItem "2023"
    Combo2.AddItem "2024"
    Combo2.AddItem "2025"
    
    AnioHoy$ = XVALO(Mid$(FECHATEXLAR$(HOY(HO$)), 7, 4))
    
    For Item% = 0 To Combo1.ListCount - 1
        If AnioHoy$ = Combo1.LIST(Item%) Then
            Combo2.ListIndex = Item%
            Combo2.TEXT = "": Combo2.TEXT = AnioHoy$
            Exit For
        End If
    Next Item%
    
    Call CARSELMESA
    
    LoadFormulario = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Me
    
End Sub

Private Sub GRID_Click()

    ' CARGA LOS DATOS DEL DIA ELEGIDO EN LA GRILLA DE AUSENCIAS,
    ' EN LOS CASILLEROS PARA SU MODIFICACION
    FilaGrid% = GRID.MouseRow
    
    If FilaGrid% > 0 Then
        FechaSet$ = GRID.TextMatrix(FilaGrid%, 0)
        HorasAus$ = GRID.TextMatrix(FilaGrid%, 1)
        Motivo$ = GRID.TextMatrix(FilaGrid%, 2)
        Observacion$ = GRID.TextMatrix(FilaGrid%, 3)
        
        Label3.Caption = FechaSet$
        Text2 = HorasAus$
        Text3 = Motivo$
        Text4 = Observacion$
    End If
    
End Sub

Private Sub GRID1_Click()

    ' CARGA LOS DATOS DEL DIA ELEGIDO EN LA GRILLA DE CANT EMPLEADOS,
    ' EN LOS CASILLEROS PARA SU MODIFICACION
    FilaGrid% = GRID1.MouseRow
    
    If FilaGrid% > 0 Then
        Meses$ = GRID1.TextMatrix(FilaGrid%, 0)
        CantEmpleados$ = GRID1.TextMatrix(FilaGrid%, 1)
        DiasLabora$ = GRID1.TextMatrix(FilaGrid%, 2)
       
        Label8 = Meses$
        Text1 = CantEmpleados$
        Text5 = DiasLabora$
    End If
    
End Sub

Private Sub Label12_Change()

    ' ELIGE FECHA EN SOLAPA DE CONTROL DE PRODUCCION
    Seccion$ = Label9.Caption
    PERIODO$ = Label12.Caption
    
    If Seccion$ <> "" And PERIODO$ <> "" And LoadFormulario = 0 Then Call CompletaGrillaControlProduccion
    
End Sub

Private Sub Label16_Change()

    ' ELIGE SECCION EN SOLAPA UNIDADES EN FABRICA
    Seccion$ = Label16
    
    If Seccion$ <> "" And LoadFormulario = 0 Then Call CompletaGrillaUnidadesEnFabrica
    
End Sub

Private Sub Label17_Change()

    ' ELIGE SECCION EN SOLAPA ESTADISTICAS
    Seccion$ = Label17
    
    If Seccion$ <> "" Then
        Call CompletaGrillaEstadisticas
    Else
        GRID4.Rows = 1
    End If
    
End Sub

Private Sub Label2_Change()

    ' ELIGE SECCION EN SOLAPA CONFIGURACION
    Seccion$ = Label2
    PERIODO$ = Label5.Caption
    
    If PERIODO$ = "" And Seccion$ = "" Then
        GRID.Rows = 1
    Else
        Call CompletaGrillaAusencias
    End If
    
    ' CARGA GRILLA DE CANT EMPLEADOS
    Call CompletaGrillaEmpleados
    
End Sub


Sub CalendarioEnBlancoGrillaAusencias()

    ' COMPLETA GRILLA DE AUSENCIAS CUANDO TODAVIA NO HAY INFORMACION EN LA TABLA SQL
    GRID.Rows = 1
    
    PERIO$ = Label5
    
    Color1& = &H8000000F: Color2& = &HEEEEDD
    ColorActivo& = &HC0E0DD
    
    ' OBTIENE DESDE Y HASTA DEL MES ELEGIDO
    Call MESANO$(PERIO$, DES%, HAS%)
    If DES% = 0 And HAS% = 0 Then Exit Sub
    
    For NumeroFecha% = DES% To HAS%
        
        GRID.Rows = GRID.Rows + 1
        GRID.TextMatrix(GRID.Rows - 1, 0) = FECHATEX$(NumeroFecha%)
        GRID.TextMatrix(GRID.Rows - 1, 1) = ""
        GRID.TextMatrix(GRID.Rows - 1, 2) = ""
        GRID.TextMatrix(GRID.Rows - 1, 3) = ""
        
        If ColorActivo& = Color1& Then
            ColorActivo& = Color2&
        Else
            ColorActivo& = Color1&
        End If
        RangoFila$ = CStr(GRID.Rows - 1) & ":" & CStr(GRID.Rows - 1)
        Call DisenioDeCelda(GRID, RangoFila$, "0:3", "", 0, False, False, False, 1, vbBlack, ColorActivo&)
                
    Next NumeroFecha%
        
End Sub

Private Sub Label5_Change()

    ' ELIGE FECHA EN SOLAPA CONFIGURACION
    Seccion$ = Label2
    PERIODO$ = Label5
    
    If Seccion$ = "" Or PERIODO$ = "" Then
        GRID.Rows = 1: GRID.Rows = 2
        Exit Sub
    End If
    
    Call CompletaGrillaAusencias
    
End Sub

Sub CompletaGrillaAusencias()

    Label3 = "": Text2 = "": Text3 = "": Text4 = ""

    FechaNula% = FECHANUM("01/01/80")
    Seccion$ = Label2
    PERIODO$ = Label5
    
    Color1& = &H8000000F: Color2& = &HEEEEDD
    ColorActivo& = &HC0E0DD
    
    ' OBTIENE LOS DATOS DE LA TABLA AUXILIAR DE SECCIONES (SECCIFAB), CORRESPONDIENTES A LA SECCION Y PERIODO ELEGIDOS
    ' A DIFERENCIA DE LO NECESARIO EN LA GRILLA DE CANT EMPLEADOS, LA FECHA DE SETEO ES <> '01/01/80', YA QUE CONTIENE
    ' LA FECHA EN LA CUAL SE INDICA LAS HORAS DE AUSENCIAS, MOTIVO, ETC.
    On Error GoTo ErrorGrilla
    SQLX$ = "SELECT * FROM SECCIFAB WITH(NOLOCK) WHERE"
    SQLX$ = SQLX$ & " fechaseteo >'" & FETEXCOR1$(FechaNula%) & "' AND seccion='" & Seccion$ & "' AND periodo='" & PERIODO$ & "'"
    SQLX$ = SQLX$ & " ORDER BY fechaseteo ASC"
    
    If CantRegistros("SECCIFAB", CONDI$) > 0 Then
        Dim SECCIFAB As ADODB.Recordset
        Set SECCIFAB = READSET(SQLX$)
        
        GRID.Rows = 2: FilaGrid% = 0
        With SECCIFAB
            Do While Not .EOF
                ' COMPLETA LA GRILLA
                FechaSet$ = FECHATEX$(CVINT(!FechaSeteo))
                HorasAus# = XVALO(!HorasAusencia)
                Motivo$ = SAFETEXT$(!MotivoAusencia)
                Observacion$ = SAFETEXT$(!Observaciones)
                
                FilaGrid% = FilaGrid% + 1
                GRID.TextMatrix(FilaGrid%, 0) = FechaSet$
                GRID.TextMatrix(FilaGrid%, 1) = FORMATNUM$(HorasAus#, "F4.2")
                GRID.TextMatrix(FilaGrid%, 2) = Motivo$
                GRID.TextMatrix(FilaGrid%, 3) = Observacion$
                GRID.Rows = GRID.Rows + 1
                
                ' COLOREA LA FILA
                If ColorActivo& = Color1& Then
                    ColorActivo& = Color2&
                Else
                    ColorActivo& = Color1&
                End If
                RangoFila$ = CStr(GRID.Rows - 2) & ":" & CStr(GRID.Rows - 2)
                Call DisenioDeCelda(GRID, RangoFila$, "0:3", "", 0, False, False, False, 1, vbBlack, ColorActivo&)

                .MoveNext
            Loop
        End With
        GRID.Rows = GRID.Rows - 1
        GRID.RowSel = 1
        SECCIFAB.Close
        Set SECCIFAB = Nothing
    Else
        Call CalendarioEnBlancoGrillaAusencias
    End If
    
    On Error GoTo 0
    Exit Sub
    
ErrorGrilla:
    Call MENSERR(24, "Se Ha Detectado un Error al Escribar la Grilla de Ausencias\ \Consulte a su Soporte de Sistemas\ \" & Err.Description)
    On Error GoTo 0
    
End Sub

Private Sub Label9_Change()
    
    ' ELIGE LA SECCION EN LA SOLAPA CONTROL DE PRODUCCION
    PERIODO$ = Label12.Caption
    
    If PERIODO$ <> "" And LoadFormulario = 0 Then
        Call CompletaGrillaControlProduccion
    Else
        GRID2.Rows = 1
    End If
    
End Sub

Private Sub ListadosVarios_Click()

    Resultado% = COMALTER%("Opciones de Trabajo:\\Ver Lista de Reportes\Agregar Nuevos Reportes")
    
    If Resultado% < 1 Or Resultado% > 2 Then Exit Sub
    
    If Resultado% = 1 Then
        Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    Else
        Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
    End If
    
End Sub

Private Sub Text1_Change()

    ' NO PERMITE INGRESAR DATO SI LA GRILLA NO CONTIENE DATOS
    If GRID1.Rows < 2 Then
        Text1 = ""
        Exit Sub
    End If
    
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    ' VALIDA QUE EN EL CASILLERO SE INGRESEN CARACTERES QUE SEAN SOLO NUMEROS
    If GRID1.Rows < 3 Then Exit Sub
    
    Caracter$ = Chr$(KeyAscii)
    If IsNumeric(Caracter$) = False And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub Text2_Change()

    ' NO PERMITE INGRESAR DATO SI LA GRILLA NO CONTIENE DATOS
    If GRID.Rows < 2 Then
        Text2 = ""
        Exit Sub
    End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

    ' VALIDA QUE EN EL CASILLERO SE INGRESEN CARACTERES QUE SEAN SOLO NUMEROS
    If GRID.Rows < 3 Then Exit Sub
    
    Caracter$ = Chr$(KeyAscii)
    If IsNumeric(Caracter$) = False And KeyAscii <> 8 And KeyAscii <> 46 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub Text3_Change()

    ' NO PERMITE INGRESAR DATO SI LA GRILLA NO CONTIENE DATOS
    If GRID.Rows < 2 Then
        Text3 = ""
        Exit Sub
    End If
    
End Sub

Private Sub Text4_Change()

    ' NO PERMITE INGRESAR DATO SI LA GRILLA NO CONTIENE DATOS
    If GRID.Rows < 2 Then
        Text4 = ""
        Exit Sub
    End If

End Sub

Sub CompletaGrillaEmpleados()

    GRID1.Rows = 1
    
    If Label2 = "" Then Exit Sub
    
    Label8 = "": Text1 = "": Text5 = ""
    
    Color1& = &H8000000F: Color2& = &HEEEEDD
    ColorActivo& = &HC0E0DD
    
    FechaNula% = FECHANUM("01/01/80")
    Seccion$ = Label2
    
    On Error GoTo ErrorGrilla
    
    ' OBTIENE LOS DATOS DE LA TABLA AUXILIAR DE SECCIONES (SECCIFAB), CORRESPONDIENTES A LA SECCION ELEGIDA
    ' A DIFERENCIA DE LO NECESARIO EN LA GRILLA DE AUSENCIAS, LA FECHA DE SETEO ES = '01/01/80' FILTRANDO DE ESA FORMA
    ' A LOS REGISTROS REQUERIDOS PARA ESTA GRILLA
    PERIODO$ = Mid$(TRIM$(Combo1.TEXT), 3, 2)
    CONDI$ = " fechaseteo ='" & FETEXCOR1$(FechaNula%) & "' AND seccion='" & Seccion$ & "' AND periodo LIKE '%" & PERIODO$ & "%'"
    SQLX$ = "SELECT * FROM SECCIFAB WITH(NOLOCK) WHERE " & CONDI$ & " ORDER BY numorden ASC"
    
    If CantRegistros("SECCIFAB", CONDI$) > 0 Then
        Dim SECCIFAB As ADODB.Recordset
        Set SECCIFAB = READSET(SQLX$)
        
        With SECCIFAB
            Do While Not .EOF
                ' COMPLETA LA GRILLA
                Meses$ = SAFETEXT$(!PERIODO)
                CantEmpleados% = XVALO(!CantidadEmpleados)
                DiasLabora% = XVALO(!DiasLaborables)
                
                GRID1.Rows = GRID1.Rows + 1
                GRID1.TextMatrix(GRID1.Rows - 1, 0) = Meses$
                GRID1.TextMatrix(GRID1.Rows - 1, 1) = FORMATNUM$(CantEmpleados%, "I3")
                GRID1.TextMatrix(GRID1.Rows - 1, 2) = FORMATNUM$(DiasLabora%, "I3")

                ' COLOREA LA FILA
                If ColorActivo& = Color1& Then
                    ColorActivo& = Color2&
                Else
                    ColorActivo& = Color1&
                End If
                RangoFila$ = CStr(GRID1.Rows - 1) & ":" & CStr(GRID1.Rows - 1)
                Call DisenioDeCelda(GRID1, RangoFila$, "0:2", "", 0, False, False, False, 1, vbBlack, ColorActivo&)
                
                .MoveNext
            Loop
        End With
        SECCIFAB.Close
        Set SECCIFAB = Nothing
        GRID1.RowSel = 1
    Else
        Call CalendarioEnBlancoGrillaEmpleados
    End If
    
    On Error GoTo 0
    Exit Sub
    
ErrorGrilla:
    Call MENSERR(24, "Se Ha Detectado un Error al Escribar la Grilla de Empleados\ \Consulte a su Soporte de Sistemas\ \" & Err.Description)
    On Error GoTo 0

End Sub

Sub CalendarioEnBlancoGrillaEmpleados()

    ' INICILIZA LA GRILLA DE EMPLEADOS CUANDO TODAVIA NO EXISTEN REGISTROS EN LA TABLA SQL
    Anio$ = Mid$(TRIM$(Combo1.TEXT), 3, 2)

    GRID1.Rows = 1: GRID1.Rows = GRID1.Rows + 1
    GRID1.TextMatrix(1, 0) = "ENERO " & Anio$: GRID1.Rows = GRID1.Rows + 1
    GRID1.TextMatrix(2, 0) = "FEBRERO " & Anio$: GRID1.Rows = GRID1.Rows + 1
    GRID1.TextMatrix(3, 0) = "MARZO " & Anio$: GRID1.Rows = GRID1.Rows + 1
    GRID1.TextMatrix(4, 0) = "ABRIL " & Anio$: GRID1.Rows = GRID1.Rows + 1
    GRID1.TextMatrix(5, 0) = "MAYO " & Anio$: GRID1.Rows = GRID1.Rows + 1
    GRID1.TextMatrix(6, 0) = "JUNIO " & Anio$: GRID1.Rows = GRID1.Rows + 1
    GRID1.TextMatrix(7, 0) = "JULIO " & Anio$: GRID1.Rows = GRID1.Rows + 1
    GRID1.TextMatrix(8, 0) = "AGOSTO " & Anio$: GRID1.Rows = GRID1.Rows + 1
    GRID1.TextMatrix(9, 0) = "SETIEMBRE " & Anio$: GRID1.Rows = GRID1.Rows + 1
    GRID1.TextMatrix(10, 0) = "OCTUBRE " & Anio$: GRID1.Rows = GRID1.Rows + 1
    GRID1.TextMatrix(11, 0) = "NOVIEMBRE " & Anio$: GRID1.Rows = GRID1.Rows + 1
    GRID1.TextMatrix(12, 0) = "DICIEMBRE " & Anio$
    
    Color1& = &H8000000F: Color2& = &HEEEEDD
    ColorActivo& = &HC0E0DD
    
    For Meses% = 1 To 12
    
        If ColorActivo& = Color1& Then
            ColorActivo& = Color2&
        Else
            ColorActivo& = Color1&
        End If
        RangoFila$ = CStr(Meses%) & ":" & CStr(Meses%)
        Call DisenioDeCelda(GRID1, RangoFila$, "0:2", "", 0, False, False, False, 1, vbBlack, ColorActivo&)
        
    Next Meses%
    

End Sub

Private Sub Text5_Change()

    ' NO PERMITE INGRESAR DATO SI LA GRILLA NO CONTIENE DATOS
    If GRID1.Rows < 2 Then
        Text5 = ""
        Exit Sub
    End If
    
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)

    ' VALIDA QUE EN EL CASILLERO SE INGRESEN CARACTERES QUE SEAN SOLO NUMEROS
    If GRID1.Rows < 3 Then Exit Sub
    
    Caracter$ = Chr$(KeyAscii)
    If IsNumeric(Caracter$) = False And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    
End Sub


Sub CompletaGrillaControlProduccion()

    GRID2.Rows = 1
    
    Seccion$ = Label9
    PERIODO$ = Label12
    Call MESANO(PERIODO$, DES%, HAS%)
    ExisteInfo% = 0
        
    Color1& = &H8000000F: Color2& = &HEEEEDD
    ColorActivo& = &HC0E0DD
    
    On Error GoTo ErrorGrilla
    
    CantidadSecciones& = 1
    Call EncabezadoGrillaControl(GRID2.Rows - 1)
    
    If Seccion$ = "" Then
        CantidadSecciones& = ULTREG&("SECPROD")
        If CantidadSecciones& > 1 Then
            GRID2.Rows = 0: GRID2.Rows = 1
        End If
    End If
    
    FilaDetalle% = 0
    For IdSeccion& = 1 To CantidadSecciones&
    
        If CantidadSecciones& > 1 Then
            Seccion$ = TRIM$(FORMATNUM$(CVS(REGLEIDO$("SECPROD", IdSeccion&)), "F10.4"))
            DenominacionSeccion$ = TRIM$(Mid$(REGLEIDO$("SECPROD", IdSeccion&), 5))
            GRID2.TextMatrix(GRID2.Rows - 1, 0) = UCase(DenominacionSeccion$)
            
            RangoFila$ = CStr(GRID2.Rows - 1) & ":" & CStr(GRID2.Rows - 1)
            Call DisenioDeCelda(GRID2, RangoFila$, "0:0", "", 0, True, False, True, 1, vbBlack, Color1&)
            
            GRID2.Rows = GRID2.Rows + 2
            Call EncabezadoGrillaControl(GRID2.Rows - 1)
            FilaDetalle% = GRID2.Rows - 1
        End If
           
        For Dia% = DES% To HAS%
            
            FechaFilaComienzo% = GRID2.Rows - 1
            ExisteInfo% = 0
            
            FechaReferencia$ = FECHATEX$(Dia%)
                CONDI$ = "seccion='" & Seccion$ & "' AND fechaseteo='" & FETEXCOR1$(Dia%) & "'"
            MotivoAus$ = SAFETEXT$(VALOBADA("SECCIFAB", "MotivoAusencia", CONDI$))
            Observacion$ = SAFETEXT$(VALOBADA("SECCIFAB", "Observaciones", CONDI$))
            HsAusencia# = XVALO(VALOBADA("SECCIFAB", "HorasAusencia", CONDI$))
            
            ' ESTE INFORME REQUIERE A LAS O/F QUE TIENEN INFORME DE INICIO DE PROCESO DENTRO DEL PERIODO ELEGIDO PERO NO CIERRE DE OF (COLUMNA OT ENTRADA)
            Dim ORDEFA As ADODB.Recordset
            SQLX$ = "SELECT * FROM ORDEFABR WITH(NOLOCK) WHERE seccipro='" & Seccion$ & "' AND (inicio_informado>='" & FETEXCOR1$(Dia%) & "' AND inicio_informado<'" & FETEXCOR1$(Dia% + 1) & "') ORDER BY numeorfa DESC" ' AND Status_Proceso = 1 AND StatuOrf < 3 ORDER BY numeorfa DESC"
            Set ORDEFA = READSET(SQLX$)
        
            With ORDEFA
                If Not .EOF Then
                    Do While Not .EOF
                        ' COMPLETA LA GRILLA
                        GRID2.Rows = GRID2.Rows + 1
                        GRID2.TextMatrix(GRID2.Rows - 1, 0) = FechaReferencia$
                        GRID2.TextMatrix(GRID2.Rows - 1, 1) = SAFETEXT$(!NumeOrfa)
                        
                        GRID2.TextMatrix(GRID2.Rows - 1, 3) = Observacion$
                        GRID2.TextMatrix(GRID2.Rows - 1, 4) = TRIM$(FORMATNUM$(HsAusencia#, "F5.1"))
                        GRID2.TextMatrix(GRID2.Rows - 1, 5) = MotivoAus$
                
                        ' COLOREA LA FILA
                        If ExisteInfo% = 0 Then
                            If ColorActivo& = Color1& Then
                                ColorActivo& = Color2&
                            Else
                                ColorActivo& = Color1&
                            End If
                            ExisteInfo% = 1
                        End If
                        RangoFila$ = CStr(GRID2.Rows - 1) & ":" & CStr(GRID2.Rows - 1)
                        Call DisenioDeCelda(GRID2, RangoFila$, "0:5", "", 0, False, False, False, 0, vbBlack, ColorActivo&)
                        
                        .MoveNext
                    Loop
                End If
            End With
            ORDEFA.Close
            Set ORDEFA = Nothing
            
            ' ESTE INFORME REQUIERE ADEMAS A LAS O/F QUE TIENEN FECHA DE CIERRE DENTRO DEL PERIODO ELEGIDO (COLUMNA OT SALIDA)
            SQLX$ = "SELECT * FROM ORDEFABR WITH(NOLOCK) WHERE seccipro='" & Seccion$ & "' AND (FinDeProceso >='" & FETEXCOR1$(Dia%) & "' AND FinDeProceso<'" & FETEXCOR1$(Dia% + 1) & "') ORDER BY numeorfa DESC"
            Set ORDEFA = READSET(SQLX$)
        
            With ORDEFA
                Do While Not .EOF
                    FechaFilaComienzo% = FechaFilaComienzo% + 1
                    If FechaFilaComienzo% > GRID2.Rows - 1 Then GRID2.Rows = GRID2.Rows + 1
                    GRID2.TextMatrix(FechaFilaComienzo%, 0) = FechaReferencia$
                    GRID2.TextMatrix(FechaFilaComienzo%, 2) = SAFETEXT$(!NumeOrfa)
                    
                    If ExisteInfo% = 0 Then
                        If ColorActivo& = Color1& Then
                            ColorActivo& = Color2&
                        Else
                            ColorActivo& = Color1&
                        End If
                        GRID2.TextMatrix(GRID2.Rows - 1, 3) = Observacion$
                        GRID2.TextMatrix(GRID2.Rows - 1, 4) = TRIM$(FORMATNUM$(HsAusencia#, "F5.1"))
                        GRID2.TextMatrix(GRID2.Rows - 1, 5) = MotivoAus$
                    End If
                    RangoFila$ = CStr(FechaFilaComienzo%) & ":" & CStr(FechaFilaComienzo%)
                    Call DisenioDeCelda(GRID2, RangoFila$, "0:5", "", 0, False, False, False, 0, vbBlack, ColorActivo&)
                    ExisteInfo% = 1
                    
                    .MoveNext
                Loop
            End With
            ORDEFA.Close
            Set ORDEFA = Nothing
            
            If ExisteInfo% = 0 Then
                FechaFilaComienzo% = FechaFilaComienzo% + 1
                If FechaFilaComienzo% > GRID2.Rows - 1 Then GRID2.Rows = GRID2.Rows + 1
                GRID2.TextMatrix(FechaFilaComienzo%, 0) = FechaReferencia$
                GRID2.TextMatrix(GRID2.Rows - 1, 3) = Observacion$
                GRID2.TextMatrix(GRID2.Rows - 1, 4) = TRIM$(FORMATNUM$(HsAusencia#, "F5.1"))
                GRID2.TextMatrix(GRID2.Rows - 1, 5) = MotivoAus$
                
                RangoFila$ = CStr(FechaFilaComienzo%) & ":" & CStr(FechaFilaComienzo%)
                If ColorActivo& = Color1& Then
                    ColorActivo& = Color2&
                Else
                    ColorActivo& = Color1&
                End If
                Call DisenioDeCelda(GRID2, RangoFila$, "0:5", "", 0, False, False, False, 0, vbBlack, ColorActivo&)
            End If
            
        Next Dia%
    
        If CantidadSecciones& > 1 Then
            GRID2.Rows = GRID2.Rows + 2
        End If
        
    Next IdSeccion&
    
    On Error GoTo 0
    Exit Sub
    
ErrorGrilla:
    Call MENSERR(24, "Se Ha Detectado un Error al Escribar la Grilla de Control\ \Consulte a su Soporte de Sistemas\ \" & Err.Description)
    On Error GoTo 0

End Sub

Sub CompletaGrillaEstadisticas()

    GRID4.Rows = 1
    
    Seccion$ = Label17
    Anio$ = Mid$(Combo2.TEXT, 3, 2)
    FechaNula% = FECHANUM("01/01/80")
        
    Color1& = &H8000000F: Color2& = &HEEEEDD
    ColorActivo& = &HC0E0DD
    
    On Error GoTo ErrorGrilla
        
    CantidadSecciones& = 1
    Call EncabezadoGrillaEstadisticas(GRID4.Rows - 1)
    
    If Seccion$ = "" Then
        CantidadSecciones& = ULTREG&("SECPROD")
        If CantidadSecciones& > 1 Then
            GRID4.Rows = 0: GRID4.Rows = 1
        End If
    End If
    
    FilaDetalle% = 0
    For IdSeccion& = 1 To CantidadSecciones&
        If CantidadSecciones& > 1 Then
            Seccion$ = TRIM$(FORMATNUM$(CVS(REGLEIDO$("SECPROD", IdSeccion&)), "F10.4"))
            DenominacionSeccion$ = TRIM$(Mid$(REGLEIDO$("SECPROD", IdSeccion&), 5))
            GRID4.TextMatrix(GRID4.Rows - 1, 0) = UCase(DenominacionSeccion$)
            
            RangoFila$ = CStr(GRID4.Rows - 1) & ":" & CStr(GRID4.Rows - 1)
            Call DisenioDeCelda(GRID4, RangoFila$, "0:0", "", 0, True, False, True, 1, vbBlack, Color1&)
            
            GRID4.Rows = GRID4.Rows + 2
            Call EncabezadoGrillaEstadisticas(GRID4.Rows - 1)
            FilaDetalle% = GRID4.Rows - 1
        End If
        
        For Meses% = 1 To 12
            ' RECORRE LOS MESES DEL AÑO ELEGIDO Y OBTIENE PARA CADA CASO LO REQUERIDO EN CADA COLUMNA
            ' LOS DATOS SON TOMADOS DE LAS TABLAS ORDEFABR Y SECCIFAB
            MesActivo$ = ""
            Select Case Meses%
                Case 1: MesActivo$ = "ENERO "
                Case 2: MesActivo$ = "FEBRERO "
                Case 3: MesActivo$ = "MARZO "
                Case 4: MesActivo$ = "ABRIL "
                Case 5: MesActivo$ = "MAYO "
                Case 6: MesActivo$ = "JUNIO "
                Case 7: MesActivo$ = "JULIO "
                Case 8: MesActivo$ = "AGOSTO "
                Case 9: MesActivo$ = "SETIEMBRE "
                Case 10: MesActivo$ = "OCTUBRE "
                Case 11: MesActivo$ = "NOVIEMBRE "
                Case 12: MesActivo$ = "DICIEMBRE "
            End Select
            '
            PERIODO$ = MesActivo$ & Anio$
            Call MESANO(PERIODO$, MesDesde%, MesHasta%)
            '
            CONDI$ = " fechaseteo ='" & FETEXCOR1$(FechaNula%) & "' AND seccion='" & Seccion$ & "' AND periodo='" & PERIODO$ & "'"
            DiasLaborables% = XVALO(VALOBADA("SECCIFAB", "DiasLaborables", CONDI$))
            Personal% = XVALO(VALOBADA("SECCIFAB", "CantidadEmpleados", CONDI$))
            CONDI$ = "FinDeProceso >= '" & FETEXCOR1$(MesDesde%) & "' AND FinDeProceso <= '" & FETEXCOR1$(MesHasta%) & "' AND SecciPro='" & Seccion$ & "'"
            Unidades% = XVALO(VALOBADA("ORDEFABR", "Sum(Canti_fin_proceso)", CONDI$))
            '
            HorasTotales# = (Personal% * 8.5) * DiasLaborables%
            '
            CONDI$ = "seccion='" & Seccion$ & "' AND periodo='" & PERIODO$ & "'"
            HorasAusencia# = XVALO(VALOBADA("SECCIFAB", "Sum(HorasAusencia)", CONDI$))
            '
            HorasProductivas# = HorasTotales# - HorasAusencia#
            '
            PersonalProductivo# = 0
            If HorasTotales# > 0 Then
                PersonalProductivo# = (HorasProductivas# * Personal%) / HorasTotales
            End If
            '
            UnidadesPorPersonal% = 0
            If PersonalProductivo# > 0 Then
                UnidadesPorPersonal% = Unidades% / PersonalProductivo#
            End If
            '
            UnidadesPorDia# = 0
            If DiasLaborables% > 0 Then
                UnidadesPorDia# = Unidades% / DiasLaborables%
            End If
    
            GRID4.Rows = GRID4.Rows + 1
            GRID4.TextMatrix(FilaDetalle% + Meses%, 0) = PERIODO$
            GRID4.TextMatrix(FilaDetalle% + Meses%, 1) = CStr(DiasLaborables%)
            GRID4.TextMatrix(FilaDetalle% + Meses%, 2) = CStr(Unidades%)
            GRID4.TextMatrix(FilaDetalle% + Meses%, 3) = CStr(Personal%)
            GRID4.TextMatrix(FilaDetalle% + Meses%, 4) = TRIM$(FORMATNUM$(HorasTotales#, "F10.1"))
            GRID4.TextMatrix(FilaDetalle% + Meses%, 5) = TRIM$(FORMATNUM$(HorasAusencia#, "F10.1"))
            GRID4.TextMatrix(FilaDetalle% + Meses%, 6) = TRIM$(FORMATNUM$(HorasProductivas#, "F10.1"))
            GRID4.TextMatrix(FilaDetalle% + Meses%, 7) = TRIM$(FORMATNUM$(PersonalProductivo#, "F10.2"))
            GRID4.TextMatrix(FilaDetalle% + Meses%, 8) = TRIM$(FORMATNUM$(UnidadesPorPersonal%, "F10.2"))
            GRID4.TextMatrix(FilaDetalle% + Meses%, 9) = TRIM$(FORMATNUM$(UnidadesPorDia#, "F10.2"))
            
            If ColorActivo& = Color1& Then
                ColorActivo& = Color2&
            Else
                ColorActivo& = Color1&
            End If
            RangoFila$ = CStr(GRID4.Rows - 1) & ":" & CStr(GRID4.Rows - 1)
            Call DisenioDeCelda(GRID4, RangoFila$, "0:9", "", 0, False, False, False, 1, vbBlack, ColorActivo&)
                    
        Next Meses%
        
        If CantidadSecciones& > 1 Then
            GRID4.Rows = GRID4.Rows + 2
        End If
        
    Next IdSeccion&
    
    On Error GoTo 0
    Exit Sub
    
ErrorGrilla:
    Call MENSERR(24, "Se Ha Detectado un Error al Escribar la Grilla de Estadísticas\ \Consulte a su Soporte de Sistemas\ \" & Err.Description)
    On Error GoTo 0

End Sub

Sub EncabezadoDeGrillas()

    Call EncabezadoGrillaConfiguracion
    Call EncabezadoGrillaControl(0)
    Call EncabezadoGrillaUnidades(0)
    Call EncabezadoGrillaEstadisticas(0)

End Sub
Sub EncabezadoGrillaConfiguracion()

    GRID.TextMatrix(0, 0) = "Fecha"
    GRID.TextMatrix(0, 1) = "Horas Ausencia": GRID.ColWidth(1) = GRID.ColWidth(1) + 500
    GRID.TextMatrix(0, 2) = "Motivo": GRID.ColWidth(2) = GRID.ColWidth(2) + 3000
    GRID.TextMatrix(0, 3) = "Observaciones": GRID.ColWidth(3) = GRID.ColWidth(3) + 3000
    Call DisenioDeCelda(GRID, "0:0", "0:3", "", 0, True, False, False, 0, vbWhite, RGB(153, 184, 208))
    '
    GRID1.TextMatrix(0, 0) = "Periodo": GRID1.ColWidth(0) = GRID1.ColWidth(0) + 300
    GRID1.TextMatrix(0, 1) = "Cant. Empleados": GRID1.ColWidth(1) = GRID1.ColWidth(1) + 500
    GRID1.TextMatrix(0, 2) = "Dias Laborables": GRID1.ColWidth(2) = GRID1.ColWidth(2) + 500
    Call DisenioDeCelda(GRID1, "0:0", "0:2", "", 0, True, False, False, 0, vbWhite, RGB(153, 184, 208))

End Sub

Sub EncabezadoGrillaControl(NumeroDeFila As Integer)

    If NumeroDeFila = 0 Then
        GRID2.Rows = 1
        GRID2.Cols = 1: GRID2.Cols = 6
        GRID2.MergeCells = flexMergeRestrictColumns: GRID2.MergeCol(0) = True
        GRID2.ColWidth(3) = GRID2.ColWidth(3) + 1500
        GRID2.ColWidth(4) = GRID2.ColWidth(4) + 400
        GRID2.ColWidth(5) = GRID2.ColWidth(5) + 1000
    End If
    
    GRID2.TextMatrix(NumeroDeFila, 0) = "Fecha"
    GRID2.TextMatrix(NumeroDeFila, 1) = "OT Entrada"
    GRID2.TextMatrix(NumeroDeFila, 2) = "OT Salida"
    GRID2.TextMatrix(NumeroDeFila, 3) = "Observaciones"
    GRID2.TextMatrix(NumeroDeFila, 4) = "Hs Ausencias"
    GRID2.TextMatrix(NumeroDeFila, 5) = "Motivo del faltante"
    RangoFila$ = CStr(NumeroDeFila) & ":" & CStr(NumeroDeFila)
    Call DisenioDeCelda(GRID2, RangoFila$, "0:5", "", 0, True, False, False, 0, vbWhite, RGB(153, 184, 208))

End Sub

Sub EncabezadoGrillaUnidades(NumeroDeFila As Integer)

    If NumeroDeFila = 0 Then
        GRID3.Rows = 1
        GRID3.Cols = 0: GRID3.Cols = 5
        GRID3.ColWidth(0) = GRID3.ColWidth(0) + 200
        GRID3.ColWidth(1) = GRID3.ColWidth(1) + 1200
        GRID3.ColWidth(2) = GRID3.ColWidth(2) + 1700
        GRID3.ColWidth(3) = GRID3.ColWidth(3) + 2000
    End If
    
    GRID3.TextMatrix(NumeroDeFila, 0) = "OT"
    GRID3.TextMatrix(NumeroDeFila, 1) = "Número de Chasis"
    GRID3.TextMatrix(NumeroDeFila, 2) = "Tipo de Chasis"
    GRID3.TextMatrix(NumeroDeFila, 3) = "Línea"
    GRID3.TextMatrix(NumeroDeFila, 4) = "Interno"
    RangoFila$ = CStr(NumeroDeFila) & ":" & CStr(NumeroDeFila)
    Call DisenioDeCelda(GRID3, RangoFila$, "0:4", "", 0, True, False, False, 0, vbWhite, RGB(153, 184, 208))

End Sub

Sub EncabezadoGrillaEstadisticas(NumeroDeFila As Integer)

    If NumeroDeFila = 0 Then
        GRID4.Rows = 1
        GRID4.Cols = 0: GRID4.Cols = 10
        GRID4.ColWidth(0) = GRID4.ColWidth(0) + 300
        GRID4.ColWidth(1) = GRID4.ColWidth(1) + 800
        GRID4.ColWidth(2) = GRID4.ColWidth(2) + 300
        GRID4.ColWidth(3) = GRID4.ColWidth(3) + 300
        GRID4.ColWidth(4) = GRID4.ColWidth(4) + 500
        GRID4.ColWidth(5) = GRID4.ColWidth(5) + 500
        GRID4.ColWidth(6) = GRID4.ColWidth(6) + 500
        GRID4.ColWidth(7) = GRID4.ColWidth(7) + 500
        GRID4.ColWidth(8) = GRID4.ColWidth(8) + 700
        GRID4.ColWidth(9) = GRID4.ColWidth(9) + 500
    End If
    
    GRID4.TextMatrix(NumeroDeFila, 0) = "Meses"
    GRID4.TextMatrix(NumeroDeFila, 1) = "Días Laborables"
    GRID4.TextMatrix(NumeroDeFila, 2) = "Unidades"
    GRID4.TextMatrix(NumeroDeFila, 3) = "Personal"
    GRID4.TextMatrix(NumeroDeFila, 4) = "Horas Totales"
    GRID4.TextMatrix(NumeroDeFila, 5) = "Horas Perdida"
    GRID4.TextMatrix(NumeroDeFila, 6) = "Hs. Produsctivas"
    GRID4.TextMatrix(NumeroDeFila, 7) = "Pers. Productivo"
    GRID4.TextMatrix(NumeroDeFila, 8) = "Unidades por Pers."
    GRID4.TextMatrix(NumeroDeFila, 9) = "Unidades por Día"
    RangoFila$ = CStr(NumeroDeFila) & ":" & CStr(NumeroDeFila)
    Call DisenioDeCelda(GRID4, RangoFila$, "0:9", "", 0, True, False, False, 0, vbWhite, RGB(153, 184, 208))

End Sub


VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Begin VB.Form SOLNOTADEB 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9150
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   18315
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   18315
   StartUpPosition =   2  'CenterScreen
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
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "Nota de Débito "
      Top             =   240
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.TextBox TXTNOTADEBITO 
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
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   5880
      ScaleHeight     =   75
      ScaleWidth      =   2955
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   3015
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "T.Cambio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7560
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      OleObjectBlob   =   "SOLNOTADEB.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "SOLNOTADEB.frx":0234
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8895
      Left            =   0
      TabIndex        =   9
      Top             =   240
      Width           =   18320
      _ExtentX        =   32306
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Lista de Notas de Débito a Generar"
      TabPicture(0)   =   "SOLNOTADEB.frx":0292
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "MSF1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "MSF0"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TXTCOTIZACION"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Text2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TXTCLIENTE"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TXTRASOCLI"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Text4"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Picture1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "N.Débitos Generadas"
      TabPicture(1)   =   "SOLNOTADEB.frx":02AE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSFGEN1"
      Tab(1).Control(1)=   "MSFGEN0"
      Tab(1).Control(2)=   "Command6"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Informes y Tabla de Monedas"
      TabPicture(2)   =   "SOLNOTADEB.frx":02CA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command7"
      Tab(2).Control(1)=   "Command5"
      Tab(2).Control(2)=   "MSFNOC"
      Tab(2).Control(3)=   "MSFMON"
      Tab(2).Control(4)=   "Label2"
      Tab(2).Control(5)=   "Label1"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Historial - Configuraciones"
      TabPicture(3)   =   "SOLNOTADEB.frx":02E6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame6"
      Tab(3).Control(1)=   "Frame2"
      Tab(3).Control(2)=   "Frame1"
      Tab(3).Control(3)=   "Frame5"
      Tab(3).ControlCount=   4
      Begin VB.Frame Frame5 
         Height          =   8160
         Left            =   -63720
         TabIndex        =   59
         Top             =   620
         Width           =   3100
         Begin VB.CommandButton Command4 
            Caption         =   "Grabar Configuraciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   64
            Top             =   1920
            Width           =   2895
         End
         Begin VB.Frame Frame10 
            Caption         =   "Plantilla de Hoja de Trabajo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1320
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   2850
            Begin VB.TextBox Text7 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
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
               TabIndex        =   69
               Top             =   480
               Width           =   1215
            End
            Begin VB.CommandButton Command13 
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
               Height          =   315
               Left            =   2565
               TabIndex        =   61
               Top             =   480
               Width           =   175
            End
            Begin VB.Label Label188 
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
               Height          =   255
               Left            =   60
               TabIndex        =   63
               Top             =   840
               Width           =   2715
            End
            Begin VB.Label Label14 
               BackStyle       =   0  'Transparent
               Caption         =   "Formulario"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   240
               TabIndex        =   62
               Top             =   525
               Width           =   1005
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Anotaciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4450
         Left            =   -74760
         TabIndex        =   49
         Top             =   2040
         Width           =   10860
         Begin VB.Frame Frame14 
            Caption         =   "Frame3"
            Height          =   3795
            Left            =   10290
            TabIndex        =   68
            Top             =   630
            Width           =   15
         End
         Begin VB.Frame Frame12 
            Caption         =   "Frame3"
            Height          =   4110
            Left            =   8085
            TabIndex        =   66
            Top             =   630
            Width           =   15
            Begin VB.Frame Frame13 
               Caption         =   "Frame3"
               Height          =   4110
               Left            =   0
               TabIndex        =   67
               Top             =   0
               Width           =   855
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Frame3"
            Height          =   4110
            Left            =   1155
            TabIndex        =   65
            Top             =   720
            Width           =   15
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00C0C0C0&
            Height          =   435
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   10515
            TabIndex        =   54
            Top             =   315
            Width           =   10580
            Begin VB.Label Label22 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Fecha     Detalle                                                   Usuario"
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   0
               TabIndex        =   55
               Top             =   0
               Width           =   10520
            End
         End
         Begin VB.ListBox List1 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3660
            Left            =   120
            TabIndex        =   53
            Top             =   735
            Width           =   10580
         End
         Begin VB.Frame Frame3 
            Caption         =   "Frame3"
            Height          =   4110
            Left            =   1155
            TabIndex        =   52
            Top             =   630
            Width           =   15
         End
         Begin VB.Frame Frame4 
            Caption         =   "Frame4"
            Height          =   4005
            Left            =   8085
            TabIndex        =   51
            Top             =   630
            Width           =   15
         End
         Begin VB.Frame Frame7 
            Caption         =   "Frame7"
            Height          =   3795
            Left            =   10290
            TabIndex        =   50
            Top             =   630
            Width           =   15
         End
         Begin VB.Line Line1 
            X1              =   10580
            X2              =   10580
            Y1              =   320
            Y2              =   6225
         End
         Begin VB.Line Line4 
            X1              =   120
            X2              =   10725
            Y1              =   4410
            Y2              =   4410
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Novedades"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -74760
         TabIndex        =   37
         Top             =   6720
         Width           =   10860
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00C0C0C0&
            Height          =   425
            Left            =   120
            ScaleHeight     =   360
            ScaleWidth      =   10515
            TabIndex        =   40
            Top             =   315
            Width           =   10580
            Begin VB.Label Label9 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Fecha     Detalle                                                   Usuario"
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Left            =   0
               TabIndex        =   42
               Top             =   0
               Width           =   10560
            End
            Begin VB.Label Label6 
               Caption         =   "Label5"
               Height          =   15
               Left            =   0
               TabIndex        =   41
               Top             =   420
               Width           =   1380
            End
            Begin VB.Line Line2 
               X1              =   0
               X2              =   1155
               Y1              =   420
               Y2              =   420
            End
         End
         Begin VB.CommandButton Command9 
            Caption         =   "O.K"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   745
            Left            =   8475
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   1140
            Width           =   750
         End
         Begin VB.TextBox Text6 
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1175
            Left            =   1275
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   38
            Top             =   735
            Width           =   7155
         End
         Begin VB.Label Label21 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   48
            Top             =   735
            Width           =   1170
         End
         Begin VB.Line Line5 
            X1              =   8295
            X2              =   8295
            Y1              =   735
            Y2              =   1995
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Registrar Anotación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   9465
            TabIndex        =   47
            Top             =   1305
            Width           =   960
         End
         Begin VB.Line Line7 
            X1              =   10695
            X2              =   10695
            Y1              =   315
            Y2              =   1900
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
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
            Left            =   8340
            TabIndex        =   46
            Top             =   840
            Width           =   2220
         End
         Begin VB.Line Line3 
            X1              =   0
            X2              =   10665
            Y1              =   1900
            Y2              =   1900
         End
         Begin VB.Label Label10 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label10"
            Height          =   75
            Left            =   8490
            TabIndex        =   45
            Top             =   1080
            Width           =   2205
         End
         Begin VB.Label Label57 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   90
            Index           =   0
            Left            =   135
            TabIndex        =   44
            Top             =   1080
            Width           =   1155
         End
         Begin VB.Label Label19 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Index           =   1
            Left            =   120
            TabIndex        =   43
            Top             =   1095
            Width           =   1170
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Historial de Diferencas de Cambio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1320
         Left            =   -74760
         TabIndex        =   32
         Top             =   620
         Width           =   10890
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
            Height          =   300
            Left            =   3070
            TabIndex        =   58
            Top             =   495
            Width           =   180
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
            Height          =   300
            Left            =   2060
            TabIndex        =   56
            Top             =   840
            Width           =   180
         End
         Begin VB.Label Label23 
            BackColor       =   &H80000016&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   2280
            TabIndex        =   57
            Top             =   840
            Width           =   8340
         End
         Begin VB.Label Ltota 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Recibo"
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
            Index           =   0
            Left            =   -840
            TabIndex        =   36
            Top             =   525
            Width           =   1905
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000016&
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
            ForeColor       =   &H00000000&
            Height          =   300
            Left            =   1155
            TabIndex        =   35
            Top             =   480
            Width           =   1905
         End
         Begin VB.Label Label11 
            BackColor       =   &H80000016&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1155
            TabIndex        =   34
            Top             =   840
            Width           =   900
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente:"
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
            Left            =   315
            TabIndex        =   33
            Top             =   885
            Width           =   750
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   8160
         ScaleHeight     =   240
         ScaleWidth      =   375
         TabIndex        =   31
         ToolTipText     =   "Caso Cuando se Genera una N. Débito y el mismo contiene un Anticipo No Aplicado"
         Top             =   780
         Width           =   405
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   29
         Text            =   "XXXXX"
         Top             =   1020
         Width           =   405
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Listar "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -59235
         TabIndex        =   25
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Dif.Cambio C / N.Débito Generada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   23
         Top             =   1020
         Width           =   4815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Listar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69960
         TabIndex        =   22
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Cancel          =   -1  'True
         Caption         =   "Listar Pendientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   1220
         Picture         =   "SOLNOTADEB.frx":0302
         TabIndex        =   15
         Top             =   780
         Width           =   180
      End
      Begin VB.TextBox TXTRASOCLI 
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1400
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1020
         Width           =   3885
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
         Left            =   1400
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "R.Social"
         Top             =   780
         Width           =   3885
      End
      Begin VB.TextBox TXTCLIENTE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   30
         TabIndex        =   0
         Top             =   1020
         Width           =   1180
      End
      Begin VB.TextBox Text2 
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
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Cliente"
         Top             =   780
         Width           =   1180
      End
      Begin VB.TextBox TXTCOTIZACION 
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   12990
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1020
         Width           =   1335
      End
      Begin VB.TextBox Text1 
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   11520
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "T.Cambio (HOY)"
         Top             =   1020
         Width           =   1485
      End
      Begin MSFlexGridLib.MSFlexGrid MSF0 
         Height          =   2565
         Left            =   0
         TabIndex        =   17
         ToolTipText     =   "Doble Click Para Generar Nota de Débito"
         Top             =   1380
         Width           =   18250
         _ExtentX        =   32200
         _ExtentY        =   4524
         _Version        =   393216
         Rows            =   3
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         GridLines       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
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
      Begin MSFlexGridLib.MSFlexGrid MSF1 
         Height          =   4725
         Left            =   0
         TabIndex        =   18
         ToolTipText     =   "Filas En Rojo Falta Cotización de Dolar En Factura o a Fecha de Acreditación"
         Top             =   4020
         Width           =   18250
         _ExtentX        =   32200
         _ExtentY        =   8334
         _Version        =   393216
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         GridLines       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
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
      Begin MSFlexGridLib.MSFlexGrid MSFGEN0 
         Height          =   2565
         Left            =   -74880
         TabIndex        =   19
         ToolTipText     =   "Doble Click Para Generar Nota de Débito  -  Boton Derecho Para Quitar de la Lista"
         Top             =   1500
         Width           =   17715
         _ExtentX        =   31247
         _ExtentY        =   4524
         _Version        =   393216
         Rows            =   3
         Cols            =   4
         FixedRows       =   2
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         GridLines       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
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
      Begin MSFlexGridLib.MSFlexGrid MSFGEN1 
         Height          =   4605
         Left            =   -74880
         TabIndex        =   20
         ToolTipText     =   "Filas En Rojo Falta Cotización de Dolar En Factura o a Fecha de Acreditación"
         Top             =   4140
         Width           =   17835
         _ExtentX        =   31459
         _ExtentY        =   8123
         _Version        =   393216
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         GridLines       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
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
      Begin MSFlexGridLib.MSFlexGrid MSFNOC 
         Height          =   7365
         Left            =   -74880
         TabIndex        =   21
         ToolTipText     =   "Doble Click Para Generar Nota de Débito"
         Top             =   1380
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   12991
         _Version        =   393216
         Rows            =   3
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         GridLines       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
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
      Begin MSFlexGridLib.MSFlexGrid MSFMON 
         Height          =   7365
         Left            =   -64080
         TabIndex        =   26
         ToolTipText     =   "Doble Click Para Generar Nota de Débito"
         Top             =   1380
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   12991
         _Version        =   393216
         Rows            =   3
         Cols            =   4
         BackColor       =   16777215
         BackColorFixed  =   16501405
         BackColorBkg    =   16777215
         FillStyle       =   1
         GridLines       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
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
      Begin VB.Label Label4 
         Caption         =   "Sin Cotización de Moneda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8880
         TabIndex        =   30
         Top             =   1065
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Recibo con Notas de Débito Generada-s"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8640
         TabIndex        =   28
         Top             =   825
         Width           =   2655
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Tabla de Cotización de Moneda por Día"
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
         Left            =   -63240
         TabIndex        =   27
         Top             =   1020
         Width           =   4095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Lista de Recibos a no Considerar para Cálculo de Diferencia de Cambio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   24
         Top             =   900
         Width           =   4095
      End
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu MNUARCHIVO 
      Caption         =   "ARCHIVO"
      Visible         =   0   'False
      Begin VB.Menu mnucopiar 
         Caption         =   " Copiar Celda Seleccionada"
         Visible         =   0   'False
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu mnudesap 
         Caption         =   "Des-Aplicar Dif.Cambio de Nota de Débito/Crédito"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MNUARCH2 
      Caption         =   "ARCH2"
      Visible         =   0   'False
      Begin VB.Menu mnuquitarlistapendientes 
         Caption         =   "Quitar de Lista de Pendienes"
         Visible         =   0   'False
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu mnucopiar2 
         Caption         =   "Copiar Celda Seleccionada"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MNUARCH3 
      Caption         =   "ARCH3"
      Visible         =   0   'False
      Begin VB.Menu mnucopiar3 
         Caption         =   "Copiar Celda Seleccionada"
         Visible         =   0   'False
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuvolveraincluir 
         Caption         =   "Volver a Incluir en Recibos a Calcular."
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuaccesodirecto 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuctacte 
         Caption         =   "Cta.Corriente"
      End
      Begin VB.Menu MNUTRAZABRECIBO 
         Caption         =   "Consulta de Aplicaciones(Por Recibo o N/C)"
      End
   End
End
Attribute VB_Name = "SOLNOTADEB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GLOBALFRM_strContenidoCelda$
Dim GLOBALFRM_FILA&
Sub CONTROLCONTADO(NCLIE)
   'ESTE CONTROL POR SI LA FECHA DE ACREDITACION DE LA COBRANZA ES LA MISMA FECHA DE LA FACTURA
   'el dia 03/03/23 Fernando dijo que se quite este control
   'el caso fue que habia cobrado al dia con cheque pero a las 48 hs la cotizacion ya habia cambiado y por este control _
   'se tomamba la fecha de acreditacion del cheque no se consideraba las 48 hs.
   Exit Sub
   
   For i& = 2 To MSF1.Rows - 1
      FECHACREDFACTURA$ = MSF1.TextMatrix(i&, 9)
      FACTURAX$ = MSF1.TextMatrix(i&, 8)
      For J& = 1 To MSF1.Rows - 1
         FECHACREDCHEQUE$ = MSF1.TextMatrix(J&, 4)
         IMPOPESOS# = XVALO(MSF1.TextMatrix(J&, 5))
         If FECHANUM(FECHACREDCHEQUE$) = FECHANUM(FECHACREDFACTURA$) Then
           CONDI$ = "TIPOMOVIM = 'FVEN' AND CODICOM_LAR = '" & FACTURAX$ & "' AND NUCLIEN = " & NCLIE
           COTIFAC# = XVALO(VALOBADA("CAJABANC", "COTIZA_DOLAR", CONDI$, "NOMESS"))
           If COTIFAC# > 0 Then
              MSF1.TextMatrix(J&, 6) = FORMATNUM$(COTIFAC#, "F12.2")
              MSF1.TextMatrix(J&, 7) = FORMATNUM$(IMPOPESOS# / COTIFAC#, "F12.2")
              Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF1, vbBlack, J&) ' POR SI ESTABA EN ROJO
           End If
         End If
      Next J&
   Next i&
End Sub

Sub DESHABILITAR_CONTROLES()
    Command2.Enabled = False
    TXTCLIENTE.Enabled = False
    MSF0.Enabled = False
    MSF1.Enabled = False
    Command6.Enabled = False
    MSFGEN0.Enabled = False
    MSFGEN1.Enabled = False
    Command5.Enabled = False
    Command7.Enabled = False
End Sub
Sub HABILITAR_CONTROLES()
    Command2.Enabled = True
    TXTCLIENTE.Enabled = True
    MSF0.Enabled = True
    MSF1.Enabled = True
    Command6.Enabled = True
    MSFGEN0.Enabled = True
    MSFGEN1.Enabled = True
    Command5.Enabled = True
    Command7.Enabled = True
End Sub

Sub GRABAHOJADETRABAJO()
    NDEB$ = TRIM$(Me.TXTNOTADEBITO)
    If NDEB$ = "" Then Exit Sub
    NC = XVALO(MSF0.TextMatrix(MSF0.Row, 1))
    RECIBOX$ = TRIM$(MSF0.TextMatrix(MSF0.Row, 3))
    SQLX$ = " SELECT * FROM HojaTrabajoDifCambio "
    SQLX$ = SQLX$ + " WHERE NOTADEBITOGENERADA = '" & NDEB$ & "'"
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    JJ& = 0
    rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    For i& = 2 To MSF1.Rows - 1
        If MSF1.TextMatrix(i&, 0) = "||" Then Exit For
        With rs
          .AddNew
          !CODIEMPR = CODIEMPR%
          !NCLIEN = NC
          !recibo = RECIBOX$
            FEARECI% = CVINT(MSF1.TextMatrix(i&, 2))
          If FEARECI% > 0 Then
             !FechaRecibo = CVDAT(FEARECI%)
          End If
          !MedioPago = MSF1.TextMatrix(i&, 3)
            FEACREDPAG% = CVINT(MSF1.TextMatrix(i&, 4))
          If FEACREDPAG% > 0 Then
             !FechaAcreditPago = CVDAT(FEACREDPAG%)
          End If
          !ImpoMedioPagoPesos = XVALO(MSF1.TextMatrix(i&, 5))
          !TCambioFechaAcredit = XVALO(MSF1.TextMatrix(i&, 6))
          !DolaresCobrados = XVALO(MSF1.TextMatrix(i&, 7))
          !FACTURA_APL = TRIM$(MSF1.TextMatrix(i&, 8))
             FEFACAP% = CVINT(MSF1.TextMatrix(i&, 9))
          If FEFACAP% > 0 Then
            !FECHA_FACTURA_APL = CVDAT(FEFACAP%)
          End If
          !ImpoDolSobreFactura = XVALO(MSF1.TextMatrix(i&, 10))
          !Status = 0
          !MARBORRA = 0
          !NOTADEBITOGENERADA = NDEB$
             NCOM& = XVALO(Mid$(NDEB$, 9, 10))
          !NumeComp = NCOM&
          !DifCambioEnPesos = XVALO(FORFASQL.Total(0)) ' sin iva
          !DifCambioPesosBruto = XVALO(FORFASQL.Total(6)) 'importe total bruto
          JJ& = JJ& + 1
          !ORDEN = JJ&
          !ImpoPesosAplicados = XVALO(MSF1.TextMatrix(i&, 10))
          .Update
        End With
    Next i&
     '
    TXTX$ = "Generación De Nota de Débito: " & NDEB$ & vbCrLf
    TXTX$ = TXTX$ + " DifCambioEnPesos: $ " & FORFASQL.Total(0)
    
    Call ANOTADIFCAMBIO(NC, RECIBOX$, NCOM&, TXTX$)
   
   
End Sub

Private Sub Command1_Click()
   Call DESHABILITAR_CONTROLES
   command1.Enabled = False
   Screen.MousePointer = 11
   MSF0.Rows = 1
   TXTCLIENTE = ""
   Call GENERACION_NOTAS_X_DIF_CAMBIO("", "PENDIENTES")
   MSF1.Redraw = False: MSF0.Redraw = False
   Call COLOREAR_GRILLASQL(MSF0, &HE0E0E0)
   Call COLOREAR_GRILLASQL(MSF1, &HE0E0E0)
   MSF1.Redraw = True: MSF0.Redraw = True
   command1.Enabled = True
   Call HABILITAR_CONTROLES
   Screen.MousePointer = 1
End Sub
Private Sub Command10_Click()
   Command10.Enabled = False
10 Call SELECHO("DEUDOR12")
   NCLI = XVALO(VALACT1$("DEUDOR12"))
   If NCLI > 0 Then
      Label11 = NCLI
   End If
   Command10.Enabled = True
End Sub


Private Sub Command12_Click(Index As Integer)
    

End Sub


Private Sub Command13_Click()
    FORSELEC.ABUSQUE = Text1
    Call SELECHO("INDISQR")
    CODFOR$ = VALACT1$("INDISQR")
    If TRIM$(CODFOR$) = "" Then
        Text7.SetFocus
      Else
        Text7.TEXT = CODFOR$
    End If

End Sub


Private Sub Command2_Click()
    Command2.Enabled = False
    
    Call SELECHO("DEUDOR12")
    NCLI = XVALO(TRIM$(VALACT1$("DEUDOR12")))
    If NCLI > 0 Then
       Me.TXTCLIENTE = NCLI
    End If
    Command2.Enabled = True
End Sub

Sub GENERACION_NOTAS_X_DIF_CAMBIO(MODO$, MODO2$)
   'RECORRO EL CAJABANC ORDENADO POR DOCUMENTO Y FECHA  SOLO RECIBOS
   'TENER EN CUENTA TRANSFERENCIAS, RETENCIONES
   If MODO$ = "POR CLIENTE" Then
     If XVALO(TXTCLIENTE) < 1 Then
       Call COMUNI("Falta Seleccionar Cuenta de Cliente")
       Exit Sub
     End If
   End If
   '
   If XVALO(Me.TXTCOTIZACION) < 0.005 Then
      Call COMUNI("Falta Importe de Cotización a Fecha")
      Exit Sub
   End If
   '
   Call DESHASFECHA(Des%, Has%, PERIO$)
10 If Des% < 1 Then
      OPX% = COMALTER("No se Ha Seleccionado Rango de Fechas\\Cancelar\Ver Todos\Seleccionar Rango de Fechas")
      If OPX% <= 1 Then GoTo 99
      If OPX% = 3 Then GoTo 10
   End If
   
   DES1$ = FECHATEXLAR(Des%)
   HAS1$ = FECHATEXLAR(Has%)
   
   Call PoneEnCeroNull("CAJABANC", "Genera_Debito", REGAF&, "NOMESS")
   HOYY% = HOY(HO$)
   CONDI$ = "TIPOMOVIM = 'CHRECI'  AND NUCLIEN > 0"
   If Des% > 0 Then CONDI$ = CONDI$ + " AND FECHEMISI >= '" & DES1$ & "' AND FECHEMISI <= '" & HAS1$ & "'"
   
   If MODO$ = "POR CLIENTE" Then CONDI$ = CONDI$ + " AND NUCLIEN = " & XVALO(TXTCLIENTE)
   FIN& = CantRegistros&("CAJABANC", CONDI$, "NOMESS")
   SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " & CONDI$
   SQLX$ = SQLX$ + " and Genera_Debito >= 0"
   SQLX$ = SQLX$ + " ORDER BY CODICOM_LAR ASC ,FECHACRED DESC"
   Set rs = READSET(SQLX$)
   With rs
     If .EOF Then
       Call COMUNI("No Se Detectaron Recibos Segun Filtro ")
       GoTo 50
     Else
        Do While Not .EOF
          JJ& = JJ& + 1
          Call TRACE(20, JJ&, FIN&)
          If SAFETEXT$(!OpciMovim) = "REDON" Then GoTo 20
          RECIX$ = SAFETEXT$(!codicom_lar)
            If RECIX$ = RECIX_ANT$ Then GoTo 20 'LEO SOLAMENTE EL PRIMERO QUE ES EL PAGO MAS LEJANO EN EL TIEMPO
               If CVINT(!fechacred) <= HOYY% Then 'SI ES EL PAGO MAS LEJANO ES <= HOY LLAMO A LA FUNCION PARA GENERAR LA ND.
                  NC = XVALO(!NUCLIEN)
                  
                  If MODO2$ = "PENDIENTES" Then
                    'VALIDO SI YA SE GENERO LA N DEB O SI TODAVIA TIENE SALDO POR IMPUTAR
                    CONDI$ = "CODICOM_LAR = '" & RECIX$ & "' AND TIPOMOVIM= 'RCOB' AND NUCLIEN = " & NC
                    ImpoOrigRecibo# = XVALO(VALOBADA("CAJABANC", "IHABECOMP", CONDI$, "NOMESS"))
                    CONDI$ = "RECIBO = '" & RECIX$ & "' AND NCLIEN = " & NC
                    ImpoApliSegunNDAplicada# = XVALO(VALOBADA("HojaTrabajoDifCambio", "SUM(ImpoPesosAplicados)", CONDI$, "NOMESS"))
                    FaltaAplicar# = Abs(ImpoOrigRecibo# - ImpoApliSegunNDAplicada#)
                    If CantRegistros&("HojaTrabajoDifCambio", CONDI$, "NOMESS") > 0 And FaltaAplicar# < 1 Then GoTo 20  ' SI YA ESTA GENEARADA LA ND, NEXT
                  End If
                  If MODO2$ = "GENERADAS" Then
                    CONDI$ = "RECIBO = '" & RECIX$ & "' AND NCLIEN = " & NC
                    If CantRegistros&("HojaTrabajoDifCambio", CONDI$, "NOMESS") < 1 Then GoTo 20  ' SI YA ESTA GENEARADA LA ND, NEXT
                  End If
                  Call AGREGAR_A_TABLA_HojaTrabajoDifCambio(RECIX$, NC, 1, MODO2$)
               End If
               
          RECIX_ANT$ = RECIX$
20        .MoveNext
        Loop
     End If
   End With
   
50 rs.Close
   Set rs = Nothing
   TPXI& = MSF0.Rows - 8: If TPXI& < 1 Then TPXI& = 1
   On Error Resume Next
   MSF0.TopRow = TPXI&
   On Error GoTo 0
99
End Sub

Sub AGREGAR_A_TABLA_HojaTrabajoDifCambio(recibo$, NCLIE, Optional IMPRIME_ENCA_RECI%, Optional MODO2$)
    
    
    'CRITERIO
    'RECALCULA EL IMPORTE DE CADA NOTA DE DEBITO A GENERAR SI ESTA SUPERA EL VALOR MINIMO CONFIGURADO
    'LO MUESTRA PARA GENEARAR NOTA DE DEBITO
    Call CREAR_TABLA_HojaTrabajoDifCambio
    DifAplicada# = 0
    '
    Dim MSF00 As msFlexGrid
    Dim MSF01 As msFlexGrid
      
    If MODO2$ = "PENDIENTES" Then
      Set MSF00 = MSF0
      Set MSF01 = MSF1
    ElseIf MODO2$ = "GENERADAS" Then
      Set MSF00 = MSFGEN0
      Set MSF01 = MSFGEN1
    End If
    MSF00.Redraw = False
    Static LEIDO$, ValorMinimoNDebDifCambio#
    If LEIDO$ = "" Then
      'EN SETUP DE RECIBOS SE CONFIGURA UN MINIMO PARA QUE SE GENEREN LAS N.DEBITOS SI NO SUPERA ESTO NO LAS MUESTRA.
      ValorMinimoNDebDifCambio# = XVALO(Control("RECISTA", "MINDIFCM"))
      LEIDO$ = "SI"
    End If
    CMC0% = CUECOINT%(Control$("", "CUECHECA"))
    CTAGAN% = CUECOINT%(Control("RETENCIO", "GANANCIA"))
    CTAINGBRU% = CUECOINT%(Control("RETENCIO", "INGRBRUT"))
    MSF01.Rows = 2
    HOII% = HOY(HO$)
    'VALIDO SI ES SABADO O DOMINGO
10  If DIASEM%(HOII%) >= 6 Then
      HOII% = DIASANTES(HOII%, 1)
      GoTo 10
    End If
    DIAMENOS2% = DIASANTES(HOII%, 2)
    'ARRANCA POR LAS ACREDITACIONES POSTERIORES A LA FECHA, EN CASO QUE SE QUIERA VER ANTERIORES SE VA A TENER QUE VER.
    If MODO2$ = "PENDIENTES" Then
      'VALIDO SI EXISTE EN LA TABLA HojaTrabajoDifCambio ES POR QUE TIENE UNA SEGUNDA APLICACION
      CONDI$ = "RECIBO = '" & recibo$ & "' AND NCLIEN = " & NCLIE & " AND STATUS = 0"
      If CantRegistros&("HojaTrabajoDifCambio", CONDI$, "NOMESS") Then
         ImpoApliSegunNDAplicada# = XVALO(VALOBADA("HojaTrabajoDifCambio", "SUM(ImpoPesosAplicados)", CONDI$, "NOMESS"))
         CONDI$ = "CODOCAP_LAR = '" & recibo$ & "'"
         CONDI$ = CONDI$ + " AND NUCLIEN = " & NCLIE
         ImpoAplicado# = XVALO(VALOBADA("APLICOBRA", "IHABECOMP", CONDI$, "NOMESS"))
         DifAplicada# = Abs(ImpoAplicado# - ImpoApliSegunNDAplicada#)
      End If
    End If
     'CONSULTO EL CAJABANC SEGUN EL RECIBO PARA VER SI NO HAY CHEEQUES POR EL TEMA DE LAS 48 HS
    'POR QUE TENGO QUE CONSIDERAR QUE ESTEN TODOS LOS VALORES COBRADOS
    CONDI$ = " CODICOM_LAR = '" & recibo$ & "'"
    CONDI$ = CONDI$ + " AND NUCLIEN = " & NCLIE
    CONDI$ = CONDI$ + " AND CUECONTAI = " & CMC0%
    CONDI$ = CONDI$ + " AND FECHACRED >= '" & FETEXCOR1$(DIAMENOS2%) & "'"
    CONDI$ = CONDI$ + " AND TIPOMOVIM= 'CHRECI'"
    If CantRegistros&("CAJABANC", CONDI$, "NOMESS") > 0 Then Exit Sub
    '
    '
    SQLX$ = " SELECT * FROM CAJABANC WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE CODICOM_LAR = '" & recibo$ & "'"
    SQLX$ = SQLX$ + " AND NUCLIEN = " & NCLIE
'    SQLX$ = SQLX$ + " AND CUECONTAI = " & CMC0%
    SQLX$ = SQLX$ + " AND TIPOMOVIM= 'CHRECI'"
    SQLX$ = SQLX$ + " ORDER BY FECHACRED ASC"
    '
    J& = 1
    ACUMUIMPOREC# = 0
    ACUMUDOLCOB# = 0
    ACUMUIMPOAPLI# = 0
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        J& = J& + 1
        MSF01.Rows = J& + 1
        MSF01.TextMatrix(J&, 1) = SAFETEXT$(!CodiCom_Cor)
        MSF01.TextMatrix(J&, 2) = FECHATEX$(CVINT(!FECHEMISI))
        MSF01.TextMatrix(J&, 3) = SAFETEXT$(!DESCRIMOV)
        MSF01.TextMatrix(J&, 4) = FECHATEX$(CVINT(!fechacred))
        MSF01.TextMatrix(J&, 5) = FORMATNUM$(XVALO(!idebecomp), "F12.2")
'           ACUMUIMPOREC# = ACUMUIMPOREC# + XVALO(!IDEBECOMP)
           'VALIDO SI ES CHEQUE LA FECHA DE ACREDITACION VA A SER 48 HS. DESPUES A PARTIR DEL 1 DIA HABIL
           CMC0REG% = XVALO(!CUECONTAI)
           FE% = CVINT(!fechacred)
           If CMC0REG% = CMC0% Then
'                 BB = DIASEM%(FE%)
'                 AAAAA = FECHATEX$(FE%)
'                 DIATEX1 = DSEM$(FE%)
                 Select Case DIASEM%(FE%)
                   Case 0
                     FE% = DIASPOST(FE%, 3) 'DOMINGO LE SUMA 3 SE ACREDITA EL MIERCOLES
'                     DIATEX = DSEM$(FE%)
                   Case 1, 2, 3
                     FE% = DIASPOST(FE%, 2) 'LUNES,MARTES O MIERCOLES LE SUMA 2 SE ACREDITA EL MIERCOLES,JUEVES O VIERNES
'                     DIATEX = DSEM$(FE%)
                    Case 4, 5, 6
                     FE% = DIASPOST(FE%, 4) 'JUEVES,VIERNES O SABADO LE SUMA 4 SE ACREDITA EL LUNES,MARTES O MIERCOLES
'                     DIATEX = DSEM$(FE%)
                 End Select
           End If
           COTIHISTO# = COTIFECHA_HISTODOL#(FE%, 2)
        If COTIHISTO# > 0 Then
          MSF01.TextMatrix(J&, 6) = FORMATNUM$(COTIHISTO#, "F12.4")
          MSF01.TextMatrix(J&, 7) = FORMATNUM$(XVALO(!idebecomp) / COTIHISTO#, "F12.2")
        Else
          MSF01.TextMatrix(J&, 6) = "XXXXXXXXX"
          MSF01.TextMatrix(J&, 7) = "XXXXXXXXX"
        End If
        ACUMUIMPOAPLI# = ACUMUIMPOAPLI# + ROUND(XVALO(!idebecomp), 2)
'        ACUMUDOLCOB# = ACUMUDOLCOB# + XVALO(MSF01.TextMatrix(j&, 7))
        If COTIHISTO# <= 0 Then
           Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF01, vbRed, J&)
        End If
        .MoveNext
      Loop
    End With
    SQLX$ = " SELECT * FROM APLICOBRA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE CODOCAP_LAR = '" & recibo$ & "'"
    SQLX$ = SQLX$ + " AND NUCLIEN = " & NCLIE
    SQLX$ = SQLX$ + "  order by codocor_cor asc"
    '
    J& = 1
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
        FACTU$ = SAFETEXT$(!codocor_lar)
        If DifAplicada# > 1 Then
          'SI TIENE QUE MOSTRAR UN RECIBO QUE YA SE GENERO NOTA DE DEBITO Y QUEDO UN SALDO QUE AHORA SE APLICO
          'VERIFICA COMPARANDO UNO A UNO SI LA FACTURA YA FUE CONSIDERADA CUANDO SE GENERO LA ND ANTERIOR
          'SI LA DETECTA PASA AL SIGUIENTE, DE ESTE MODO SOLO VA APRESENTAR LAS QUE NO HAYAN GENERADO ND
          CONDI$ = "RECIBO = '" & recibo$ & "' AND  NCLIEN = " & NCLIE
          CONDI$ = CONDI$ + " AND FACTURA_APL = '" & FACTU$ & "' AND STATUS = 0"
          If CantRegistros("HojaTrabajoDifCambio", CONDI$, "NOMESS") > 0 Then GoTo 75  'NEXT
        End If
        '' ESTAS DOS FILAS LAS CAMBIE
        IMPOENFAC# = XVALO(!IMPOAPLI): ' ESTAS DOS FILAS LAS CAMBIE
        If IMPOENFAC# < 0 Then
           Call AGREGA_RECIBO_EN_RECIBO(recibo$, FACTU$, NCLIE, MSF01.Rows - 1, MODO2$, IMPOENFAC#)
            J& = MSF01.Rows
'            J& = J& + 1
            If MSF01.Rows < J& + 1 Then
              MSF01.Rows = J& + 1
            End If
           GoTo 75
        End If
        ' Hasta ACA
        J& = J& + 1
        If MSF01.Rows < J& + 1 Then
          MSF01.Rows = J& + 1
        End If
        MSF01.TextMatrix(J&, 8) = FACTU$
          'SE 30-04-19 CASO CLIENTE 92 'AS-181130-00 (AGREGO TIMPOMOVIM = 'AJUD' POR QUE HUBO N RECIBO QUE TENIA  UNA RETENCION DE GANANACIA QUE FIGURA COMO AJUSTE Y LA MOSTRABA PERO NO TRAIA LA COTIIZACION.)
          
          CONDI$ = "(TIPOMOVIM = 'AJUD' or TIPOMOVIM = 'FVEN')  AND CODICOM_LAR = '" & FACTU$ & "' AND NUCLIEN = " & NCLIE
          FEFAC% = CVINT(VALOBADA("CAJABANC", "FECHEMISI", CONDI$, "NOMESS"))
        MSF01.TextMatrix(J&, 9) = FECHATEX$(FEFAC%)
          COTIFAC# = XVALO(VALOBADA("CAJABANC", "COTIZA_DOLAR", CONDI$, "NOMESS"))
          
          IMPOENFAC# = XVALO(!IMPOAPLI) ' ESTO ESTABA ASI
          IMPORTEFACTURADOLAR# = 0
          If COTIFAC# > 0 Then
             IMPORTEFACTURADOLAR# = IMPOENFAC# / COTIFAC#
             MSF01.TextMatrix(J&, 10) = FORMATNUM$(IMPORTEFACTURADOLAR#, "F12.2")
             ULTIMA_COTIZA# = COTIFAC# 'guardo la ultima cotizacion de la factura aplicada , para aplicar a los recibos con anticipos no aplicados.
          Else
             If Left$(FACTU$, 3) = "RB-" Then  'contempla si es un recibo toma el 't.cambio' de la ultima factura aplicada del recibo original solicitado por Fernando_
               IMPORTEFACTURADOLAR# = IMPOENFAC# / ULTIMA_COTIZA#  'con esto soluciona el tema cuando en un recibo hay aplicaciones a facturas y tambien un_
               MSF01.TextMatrix(J&, 10) = FORMATNUM$(IMPORTEFACTURADOLAR#, "F12.2") 'anticipo no aplicado de otro recibo aplicado en este.
             Else
               MSF01.TextMatrix(J&, 10) = "XXXXXXXXX"
             End If
          End If
          
'          ACUMUIMPORFAC# = ACUMUIMPORFAC# + IMPORTEFACTURADOLAR#
        
        MSF01.TextMatrix(J&, 11) = ""
        MSF01.TextMatrix(J&, 12) = ""
        MSF01.TextMatrix(J&, 13) = FORMATNUM$(IMPOENFAC#, "F12.2")
        
        If COTIFAC# <= 0 Then
          Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF01, vbRed, J&)
        End If
75      .MoveNext
      Loop
    End With
    If J& = 1 Then
      'SI I& EN ESTE MOMENTO ESTA EN 1 ES POR QUE NO ENTRO EN APLICOBRA (O SEA NO TIENE APLICACIONES) POR LO TANTO SE VA.
      MSF01.Rows = 2
      Exit Sub
    End If
    '
    Call CONTROLCONTADO(NCLIE)

    J& = MSF01.Rows
    MSF01.Rows = MSF01.Rows + 1
    MSF01.TextMatrix(J&, 0) = "||"
    'AGREGAR TOTALES
    'DE MSF.ROWS EN ADELANTE EN COLUMNAS 5, 7, 10
    J& = MSF01.Rows + 1
    MSF01.Rows = J& + 1
    MSF01.TextMatrix(J&, 5) = "____________"
    MSF01.TextMatrix(J&, 7) = "____________"
    MSF01.TextMatrix(J&, 10) = "____________"
    
    J& = J& + 1
    MSF01.Rows = J& + 1
    ACUMUIMPOREC# = SUMACOLUMNA#(MSF1, 5, 2)
    ACUMUDOLCOB# = SUMACOLUMNA#(MSF1, 7, 2)
    ACUMUIMPORFAC# = SUMACOLUMNA#(MSF1, 10, 2)
    
    MSF01.TextMatrix(J&, 5) = FORMATNUM$(ACUMUIMPOREC#, "F12.2") '''
    MSF01.TextMatrix(J&, 7) = FORMATNUM$(ACUMUDOLCOB#, "F12.2")
    MSF01.TextMatrix(J&, 10) = FORMATNUM$(ACUMUIMPORFAC#, "F12.2")
    DIFCAMBIOENDOLARES# = ACUMUIMPORFAC# - ACUMUDOLCOB#
    MSF01.TextMatrix(J&, 11) = FORMATNUM$(DIFCAMBIOENDOLARES#, "F12.2")
       COTIHISTO# = COTIFECHA_HISTODOL#(HOY(HO$), 2)
       If COTIHISTO# > 0 Then
          DifCambioEnPesos# = DIFCAMBIOENDOLARES# * COTIHISTO#
          MSF01.TextMatrix(J&, 12) = FORMATNUM$(DifCambioEnPesos#, "F12.2")
       Else
          MSF01.TextMatrix(J&, 12) = "XXXXXXXXX"
       End If
    
    '
    'SI HAY DIFERENCIA DE CAMBIO LO IMPRIME EN LA GRILLA SUPERIOR
    If Abs(DifCambioEnPesos#) > ValorMinimoNDebDifCambio# Then
      If IMPRIME_ENCA_RECI% > 0 Then
        MSF00.Rows = MSF00.Rows + 1
        MSF00.TextMatrix(MSF00.Rows - 1&, 1) = NCLIE
        MSF00.TextMatrix(MSF00.Rows - 1, 2) = rasocli$(NCLIE)
        MSF00.TextMatrix(MSF00.Rows - 1, 3) = recibo$
        MSF00.TextMatrix(MSF00.Rows - 1, 4) = MSF01.TextMatrix(2, 2)
        CONDI$ = "TIPOMOVIM = 'RCOB'  AND CODICOM_LAR = '" & recibo$ & "'"
        ImpoOrigRecibo# = XVALO(VALOBADA("CAJABANC", "IHABECOMP", CONDI$, "NOMESS"))
        MSF00.TextMatrix(MSF00.Rows - 1, 5) = FORMATNUM$(ImpoOrigRecibo#, "F12.2")
        NDEB$ = ""
           If MODO2$ <> "PENDIENTES" And DifAplicada# < 1 Then
             CONDI$ = "RECIBO = '" & recibo$ & "' AND NCLIEN = " & NCLIE
             NDEB$ = VALOBADA("HojaTrabajoDifCambio", "NOTADEBITOGENERADA", CONDI$, "NOMESS")
           End If
        MSF00.TextMatrix(MSF00.Rows - 1, 6) = FORMATNUM$(ACUMUIMPOAPLI#, "F12.2")
        MSF00.TextMatrix(MSF00.Rows - 1, 7) = NDEB$
        MSF00.TextMatrix(MSF00.Rows - 1, 8) = FORMATNUM$(DifCambioEnPesos#, "F12.2")
        If DifAplicada# > 0 Then Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF00, vbBlue, MSF00.Rows - 1) 'colorea en azul si el recibo tenia una nota de debito generada (caso anticipos no aplicados)
      End If
    Else
      MSF01.Rows = 2
    End If
'   '
    Call COLOREAR_GRILLASQL(MSF1, &HE0E0E0)
    If ACUMUDOLCOB# > ACUMUIMPORFAC# Then
       MSF1.Rows = MSF1.Rows + 1
       MSF1.TextMatrix(MSF1.Rows - 1, 3) = "Atención Los Dolares  Cobrados"
       MSF1.TextMatrix(MSF1.Rows - 1, 4) = "Es Mayor al "
       MSF1.TextMatrix(MSF1.Rows - 1, 5) = "Facturado"
       Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF1, vbWhite, MSF1.Rows - 1)
       Call COLOREAR_GRILLA_SOLAFILA(MSF1, vbBlue, MSF1.Rows - 1, 0, 0, 1)
    End If
    
    MSF00.Redraw = True

End Sub
'Genera_Debito
'
'el doble click para visualizar va a tener que ser en las generadas
'
'mnudesap
'
'PROBAR
'
'BOTON DERECHO SOBRE LISTAS DE PENDIENTES PARA QUITAR DE LA LISTA DE CALCULO
'
'LISTA DE RECIBOS A NO CONSIDERAR PARA CALCULO DE DIFERENCIA DE CAMBIO
'
'EN ULTIMA SOLAPA PROBAR QUE SE LISTE CORRECTAMENTE LOS RECIBOS A NO CONSIDERAR.
'
'REVERTIR EL PUNTO ANTERIOR
'
'Y LA GRILLA DE AL LADO MOSTRAR LAS MONEDAS. PROBAR
'
'ACTUALIZAR TABLA DE MONEDAS PROBAR
'
'DifAplicada#
'SE AGREGO QUE SI ES UNA SEGUNDA APLICACION
'MUESTRE EL RECIBO = CON TODOS LOS DATOS
'PERO SOLO LA NUEVA APLICACION PARA QUE GENERE LA NOTA DE DEBITO POR LA DIFERENCIA.
'PROBAR SI FUNCIONA BIEN
'
'
'VER LUEGO DE UNA SEGUNDA APLICACION SI PRESENTA TODA LA INFO.
'
'
'
'VER DE FILTRAR POR RECIBO LAS QUE POSEEN N DEB GENERADAS
'REVISAR LAS RUTINAS PARA GRABAR LAS TRANSACCIONES EN EL ANOTADOR

Sub AGREGA_RECIBO_EN_RECIBO(recibo$, FACTU$, NCLIE, Fila&, MODO2$, Impo_en_Fact_Aplicado)
    'Exit Sub
    'ESTO ES PARA QUE EN LOS CASOS QUE UN RECIBO CONTENGA OTRO RECIBO COMO MEDIO DE PAGO
    'CASO QUE GENERE UN RECIBO Y TOME DE LOS PEDIENTES OTRO RECIBO Y UNA FACTURA.
    'TAMBIEN SE CONSIDERO EL CASO DE LAS NOTAS DE CREDITO CON EL MISMO CRITERIO
    'EJEMPLO EN AHP RECIBO 3973
    Dim MSF00 As msFlexGrid
    Dim MSF01 As msFlexGrid
    If MODO2$ = "PENDIENTES" Then
      Set MSF00 = MSF0
      Set MSF01 = MSF1
    ElseIf MODO2$ = "GENERADAS" Then
      Set MSF00 = MSFGEN0
      Set MSF01 = MSFGEN1
    End If
    MSF00.Redraw = False
    Static LEIDO$, ValorMinimoNDebDifCambio#
    If LEIDO$ = "" Then
      'EN SETUP DE RECIBOS SE CONFIGURA UN MINIMO PARA QUE SE GENEREN LAS N.DEBITOS SI NO SUPERA ESTO NO LAS MUESTRA.
      ValorMinimoNDebDifCambio# = XVALO(Control("RECISTA", "MINDIFCM"))
      LEIDO$ = "SI"
    End If
    
    CMC0% = CUECOINT%(Control$("", "CUECHECA"))
    CTAGAN% = CUECOINT%(Control("RETENCIO", "GANANCIA"))
    CTAINGBRU% = CUECOINT%(Control("RETENCIO", "INGRBRUT"))
    
    SQLX$ = " SELECT * FROM APLICOBRA WITH(NOLOCK)"
    'SQLX$ = SQLX$ + " WHERE CODOCAP_LAR = '" & recibo$ & "'" ESTA MODIFIQUE 13/09/18 P'Q' CUANDO HABIA P.EJ. UN
    SQLX$ = SQLX$ + " WHERE CODOCOR_LAR = '" & FACTU$ & "'"   'RECIBO Y UNA NOTA DE CREDITO CADA VEZ QUE ENTRABA AGREGABA 2 VECES
    SQLX$ = SQLX$ + " AND NUCLIEN = " & NCLIE
    SQLX$ = SQLX$ + " AND (LEFT(CODOCOR_LAR,5) = 'RB-X-' OR LEFT(CODOCOR_LAR,3) = 'NC-')"
    SQLX$ = SQLX$ + "  order by CODOCAP_LAR asc"
    Set rs = READSET(SQLX$)
    With rs
    Do While Not .EOF
       REXIBO$ = SAFETEXT$(!codocor_lar)
       'LEER ACA TAMBIEN EL IMPORTE APLICADO. Y LO DEBERIA LLEVAR ABAJO
       GoSub CARGARDATOSENGRILLA
      .MoveNext
    Loop
    End With
    rs.Close: Set rs = Nothing
    Exit Sub
    
CARGARDATOSENGRILLA: 'VER QUE EL IMPORTE SEA EL SEA EL QUE ESTA EN EL APLICOBRA.

    SQLX$ = " SELECT * FROM CAJABANC WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE CODICOM_LAR = '" & REXIBO$ & "'"
    SQLX$ = SQLX$ + " AND NUCLIEN = " & NCLIE
'    SQLX$ = SQLX$ + " AND CUECONTAI = " & CMC0%
    SQLX$ = SQLX$ + " AND (TIPOMOVIM= 'CHRECI' OR TIPOMOVIM= 'FVEN')"
    SQLX$ = SQLX$ + " ORDER BY FECHACRED ASC"
    '
    J& = Fila& + 1
    ACUMUIMPOREC# = 0
    ACUMUDOLCOB# = 0
    ACUMUIMPOAPLI# = 0
    Set RS2 = READSET(SQLX$)
    With RS2
      Do While Not .EOF
        J& = J& + 1
        MSF01.Rows = J& + 1
        MSF01.TextMatrix(J&, 1) = SAFETEXT$(!CodiCom_Cor)
        MSF01.TextMatrix(J&, 2) = FECHATEX$(CVINT(!FECHEMISI))
        MSF01.TextMatrix(J&, 3) = SAFETEXT$(!DESCRIMOV)
        MSF01.TextMatrix(J&, 4) = FECHATEX$(CVINT(!fechacred))
        TMVM$ = SAFETEXT$(!TIPOMOVIM)
        'ESTE ENJUAGUE ES PARA PODER LEER CORRECTAMENTE LOS VALORES CUANDO SE INCLUYEN LAS NC O RB- DENRO DE OTROS RECIBOS
        If TMVM$ = "CHRECI" Or Left(MSF01.TextMatrix(J&, 1), 3) = "NC-" Or Left(MSF01.TextMatrix(J&, 1), 3) = "RB-" Then
           MSF01.TextMatrix(J&, 5) = FORMATNUM$(XVALO(!idebecomp), "F12.2") 'ESTE SERIA EL VALOR TOTAL DEL RECIBO
        End If
        If TMVM$ = "FVEN" And Left(MSF01.TextMatrix(J&, 1), 3) = "NC-" Then
           MSF01.TextMatrix(J&, 5) = FORMATNUM$(XVALO(!ihabecomp), "F12.2")
        End If
        If Abs(Impo_en_Fact_Aplicado) > 0 And Abs(Impo_en_Fact_Aplicado) < XVALO(!idebecomp) Then
           MSF01.TextMatrix(J&, 5) = FORMATNUM$(Abs(XVALO(Impo_en_Fact_Aplicado)), "F12.2") 'ESTE DEVUELVE EL IMPORTE APLICADO EN EL RECIBO PARA CUANDO HAY EN UN RECIBO DENTRO DE OTRO QUE POSEA VARIOS MEDIOS DE PAGO EJ. NC 570 RB-4127 DEL 06/09/18
        End If
'           ACUMUIMPOREC# = ACUMUIMPOREC# + XVALO(!IDEBECOMP)
           'VALIDO SI ES CHEQUE LA FECHA DE ACREDITACION VA A SER 48 HS. DESPUES A PARTIR DEL 1 DIA HABIL
           CMC0REG% = XVALO(!CUECONTAI)
           FE% = CVINT(!fechacred)
           If CMC0REG% = CMC0% Then
'                 BB = DIASEM%(FE%)
'                 AAAAA = FECHATEX$(FE%)
'                 DIATEX1 = DSEM$(FE%)
                 Select Case DIASEM%(FE%)
                   Case 0
                     FE% = DIASPOST(FE%, 3) 'DOMINGO LE SUMA 3 SE ACREDITA EL MIERCOLES
'                     DIATEX = DSEM$(FE%)
                   Case 1, 2, 3
                     FE% = DIASPOST(FE%, 2) 'LUNES,MARTES O MIERCOLES LE SUMA 2 SE ACREDITA EL MIERCOLES,JUEVES O VIERNES
'                     DIATEX = DSEM$(FE%)
                    Case 4, 5, 6
                     FE% = DIASPOST(FE%, 4) 'JUEVES,VIERNES O SABADO LE SUMA 4 SE ACREDITA EL LUNES,MARTES O MIERCOLES
'                     DIATEX = DSEM$(FE%)
                 End Select
           End If
           COTIHISTO# = COTIFECHA_HISTODOL#(FE%, 2)
        If COTIHISTO# > 0 Then
          MSF01.TextMatrix(J&, 6) = FORMATNUM$(COTIHISTO#, "F12.4")
          'MSF01.TextMatrix(J&, 7) = FORMATNUM$(XVALO(!IDEBECOMP) / COTIHISTO#, "F12.2")
          MSF01.TextMatrix(J&, 7) = FORMATNUM$(XVALO(MSF01.TextMatrix(J&, 5)) / COTIHISTO#, "F12.2") 'MODIFICADO ESTE SERIA EL VALOR TOTAL DEL RECIBO o NC
        Else
          MSF01.TextMatrix(J&, 6) = "XXXXXXXXX"
          MSF01.TextMatrix(J&, 7) = "XXXXXXXXX"
        End If
        'ACUMUIMPOAPLI# = ACUMUIMPOAPLI# + ROUND(XVALO(!IDEBECOMP), 2) 'MODIFICADO ESTE SERIA EL VALOR TOTAL DEL RECIBO o NC
        ACUMUIMPOAPLI# = ACUMUIMPOAPLI# + ROUND(XVALO(MSF01.TextMatrix(J&, 5)), 2)
'        ACUMUDOLCOB# = ACUMUDOLCOB# + XVALO(MSF01.TextMatrix(j&, 7))
        If COTIHISTO# <= 0 Then
           Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF01, vbRed, J&)
        End If
        .MoveNext
      Loop
    End With
    RS2.Close: Set RS2 = Nothing
    '
    
'
    MSF00.Redraw = True
   Return
End Sub

'50              If DIASEM%(FE%) < 1 Or DIASEM%(FE%) >= 6 Then
'                   FE% = DIASPOST(FE%, 1)
'                   GoTo 50
'                End If
'
'                FE% = DIASPOST(FE%, 2)
'60              If DIASEM%(FE%) < 1 Or DIASEM%(FE%) >= 6 Then
'                   FE% = DIASPOST(FE%, 1)
'                   DIATEX = DSEM$(FE%)
'                   GoTo 60
'                End If

'    End Sub
Sub ANOTADIFCAMBIO(NCLIE, recibo$, NumeCom&, TXTX$, Optional FF%, Optional MODO$)
    '
    If recibo$ <> "" Then
       '
       '
       HOII% = HOY(HOS$)
       fech$ = HOS$
       If FF% > 0 Then
          fech$ = FECHATEX(FF%) ' SI TRAE FECHA EL PARAMETRO TOMO ESTA
       End If
       '
       Call FRATEXTO(TXTX$, 56) ' FRACCIONO EL TEXTO
       '
       For KKI% = 1 To CARETEX% ' RECORRO DESDE 1 A LA CANTIDAD DE FRACCIONES
           TTYY$ = Space$(96)
           If KKI% = 1 Then
              Call REPLA(TTYY$, fech$, 1, 8) ' SI ES LA PRIMERA VEZ GRABA LA FECHA Y USUARIO
              Call REPLA(TTYY$, USERTER$, 69, 24)
           End If
           Call REPLA(TTYY$, RETEX$(KKI%), 11, 56) ' SIEMPRE QUE PASA GUARDA EL CENTRO DEL TEXTO
           ANOTANTE$ = ANOTANTE$ + TTYY$ + vbCrLf
       Next KKI%
       '
       ANOTANTE$ = ANOTANTE$ & String$(88, "=") & vbCrLf
       '
       SQLX$ = "SELECT * FROM ANOTADIFCAMBIO"
       SQLX$ = SQLX$ + " WHERE NUCLIEN = " & NCLIE
       'If NUMECOM& > 0 Then SQLX$ = SQLX$ + " AND NUMECOMP = " & NUMECOM&
       SQLX$ = SQLX$ + " AND RECIBO = '" & recibo$ & "'"
       
       Dim rs As New ADODB.Recordset
       Set rs = New ADODB.Recordset
25     On Error Resume Next
       rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
       End If
       On Error GoTo 0
       '
       With rs
         If Not (.BOF And .EOF) Then
            ANOTRIOR$ = SAFETEXT$(!Historial)
         Else
            .AddNew
            !CODIEMPR = CodiEmp%
            !NUCLIEN = NCLIE
            !NumeComp = NumeCom&
            !recibo = recibo$
         End If
         !Historial = ANOTRIOR$ + ANOTANTE$
         .Update
       End With
       rs.Close
       Set rs = Nothing
       '
    End If
    '
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   PARAM_COTIZA# = XVALO(TXTCOTIZACION)
   Call ACTUALIZAR_COTIZACION(2, HOY(HO$), PARAM_COTIZA#)
   Command3.Enabled = True
End Sub
Sub ACTUALIZAR_COTIZACION(Optional IDMON%, Optional FE%, Optional PARAM_COTIZA#)
   
    If IDMON% < 1 Then IDMON% = 2
    If FE% < 1 Then
       'SI NO TIENE PARAMETRO LA FECHA SE PUEDE EDITAR
       ENTICAFE.Text9.Locked = False
       ENTICAFE.Command9.Enabled = True
       FE% = HOY(HO$)
    Else
       ' EN CASO CONTRARIO LA FECHA NO SE EDITA
       ENTICAFE.Text9.Locked = True
       ENTICAFE.Command9.Enabled = False
    End If
    'SI SE LE PASA IMPORTE MUESTRA EL PARAMETRO CASO CONTRARIO EL TICAMONE DE HOY
    If PARAM_COTIZA# > 0 Then
       ENTICAFE.Text4 = PARAM_COTIZA#
    Else
       ENTICAFE.Text4 = TRIM$(FORMATNUM$(TICAMONE(IDMON%), "F12.4"))
    End If
    ENTICAFE.Label6.Caption = UCase(DEMONECO$(IDMON%))
    ENTICAFE.Text9 = FECHATEX$(FE%)
    ENTICAFE.Show 1
   '
    If ENTICAFE.Label1 = "OK" Then
        Call CREAR_TABLA_HISTOCAMBIO
        COTIZADOL# = XVALO(ENTICAFE.Text4)
        FE% = FECHANUM(ENTICAFE.Text9)
        CONDI$ = "  FECHA = '" & FETEXCOR1$(FE%) & "' AND ID = " & IDMON%
        Dim rs As ADODB.Recordset
        Set rs = New ADODB.Recordset
        SQLX$ = "SELECT * FROM HISTOCAMBIO"
        SQLX$ = SQLX$ + " WHERE " & CONDI$
        rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
         With rs
           If .EOF Then
             .AddNew
             !FECHA = CVDAT(FE%)
             !ID = IDMON%
             !Descripcion = Left$(TRIM$(UCase(DEMONECO$(IDMON%))), 24)
             !CODIEMPR = CodiEmp%
           End If
           !COTIZA_DOLAR = COTIZADOL#
           .Update
         End With
        If FE% = HOY(HO$) Then Me.TXTCOTIZACION = FORMATNUM$(COTIZADOL#, "F12.4")
   End If
End Sub

Private Sub Command4_Click()
  Call GRACONTROL("DIFECAMB", "HOJATRAB", TRIM$(Text7))
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Screen.MousePointer = 11
   Call DESHABILITAR_CONTROLES
   MSFNOC.Rows = 1
   i& = 0
   CONDI$ = "TIPOMOVIM = 'RCOB'"
   If MODO$ = "POR CLIENTE" Then CONDI$ = CONDI$ + " AND NUCLIEN = " & XVALO(TXTCLIENTE)
   SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " & CONDI$
   SQLX$ = SQLX$ + " and Genera_Debito < 0"
   SQLX$ = SQLX$ + " ORDER BY CODICOM_LAR ASC "
   Set rs = READSET(SQLX$)
   With rs
    Do While Not .EOF
      i& = i& + 1
      MSFNOC.Rows = i& + 1
      NC = XVALO(!NUCLIEN)
      MSFNOC.TextMatrix(i&, 1) = NC
      MSFNOC.TextMatrix(i&, 2) = rasocli$(NC)
      MSFNOC.TextMatrix(i&, 3) = SAFETEXT$(!codicom_lar)
      .MoveNext
    Loop
   End With
   rs.Close
   Set rs = Nothing
   Screen.MousePointer = 1
   Call HABILITAR_CONTROLES
   Command5.Enabled = True

End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call DESHABILITAR_CONTROLES
    MSFGEN0.Rows = 1
    SQLX$ = " SELECT RECIBO,NCLIEN,NOTADEBITOGENERADA,ORDEN FROM HojaTrabajoDifCambio WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE ORDEN = 1"
    SQLX$ = SQLX$ + " GROUP BY RECIBO,NCLIEN,NOTADEBITOGENERADA,ORDEN  ORDER BY NOTADEBITOGENERADA DESC"
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
          NC = XVALO(!NCLIEN)
          RASO$ = rasocli$(NC)
          RECIBOX$ = SAFETEXT$(!recibo)
'          FERECI% = CVINT(!FechaRecibo)
          CONDI$ = "TIPOMOVIM = 'RCOB'  AND CODICOM_LAR = '" & RECIBOX$ & "'"
          ImpoOrigRecibo# = XVALO(VALOBADA("CAJABANC", "IHABECOMP", CONDI$, "NOMESS"))
          NOTDEBGEN$ = SAFETEXT$(!NOTADEBITOGENERADA)
'          DIFCMBPESOS# = XVALO(!DifCambioEnPesos)
          J& = J& + 1
          MSFGEN0.Rows = J& + 1
          MSFGEN0.TextMatrix(J&, 1) = NC
          MSFGEN0.TextMatrix(J&, 2) = RASO$
          MSFGEN0.TextMatrix(J&, 3) = RECIBOX$
          MSFGEN0.TextMatrix(J&, 4) = FORMATNUM$(ImpoOrigRecibo#, "F12.2")
          MSFGEN0.TextMatrix(J&, 5) = NOTDEBGEN$
          .MoveNext
      Loop
    End With
    Call HABILITAR_CONTROLES
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
     Command7.Enabled = False
     Call DESHABILITAR_CONTROLES
     Call DESHASFECHA(Des%, Has%, PERIO$)
     If Des% < 1 Then GoTo 99
     
     DES1$ = FECHATEXLAR(Des%)
     HAS1$ = FECHATEXLAR(Has%)
     FE% = Des%
     For i% = Des% To Has%
       J& = J& + 1
       MSFMON.Rows = J& + 1
       If DIASEM%(FE%) > 0 And DIASEM%(FE%) < 6 Then ' NO SE INCLUYE SABADO Y DOMINGO
            MSFMON.TextMatrix(J&, 1) = FECHATEX$(FE%)
            MSFMON.TextMatrix(J&, 2) = 2
            MSFMON.TextMatrix(J&, 3) = DEMONECO$(2)
            MSFMON.TextMatrix(J&, 4) = FORMATNUM$(COTIFECHA_HISTODOL#(FE%, 2), "F12.4")
       ElseIf DIASEM%(FE%) = 6 Then
            MSFMON.TextMatrix(J&, 1) = "Sábado"
            Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSFMON, &HE0E0E0, J&)
            Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSFMON, vbBlue, J&)
       ElseIf DIASEM%(FE%) = 0 Then
            MSFMON.TextMatrix(J&, 1) = "Domingo"
            Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSFMON, &HE0E0E0, J&)
            Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSFMON, vbRed, J&)
       End If
       
       FE% = DIASPOST(FE%, 1)
     Next i%
   
99   Command7.Enabled = True
     Call HABILITAR_CONTROLES
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
    Call APERBASDAT("CABAN")
    'SQLX$ = "SELECT * FROM CAJABANC "
    'SQLX$ = SQLX$ + "where INT(CDBL(FechConta)) >= " & DESDAT & " "
    'SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) <= " & HASDAT & " "
    NC = XVALO(Label11)
    CONDI$ = " TipoMovim = 'RCOB' "
    If NC > 0 Then CONDI$ = CONDI$ + " AND NUCLIEN = " & NC
    CONDI$ = CONDI$ + "Order by FechConta, NUMECOMP, nuclien ASC "
    Call CARGALISEL("LISRECI", "CAJABANC", "CODICOM_LAR/A18;FECHEMISI/D8;IHABECOMP/F12.2", CONDI$, "NOMESS")
    Call SELECHO("LISRECI")
    NRECI$ = TRIM$(VALACT1$("LISRECI"))
    If NRECI$ = "" Then GoTo 99
    '
    Label7.Caption = NRECI$

   
99 Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   RECIBOX$ = TRIM$(Label7)
   If RECIBOX$ <> "" Then
     If TRIM$(Text6) <> "" Then
       Call FRATEXTO(Text6, 56)
       
       For KKI% = 1 To CARETEX%
         TTX1$ = Space$(76)
         Call REPLA(TTX1$, RETEX$(KKI%), 11, 56)
         If KKI% = 1 Then
           Call REPLA(TTX1$, HOS$, 1, 8)
           Call REPLA(TTX1$, USERTER$, 69, 24)
         End If
         TTXX$ = TTXX$ + TTX1$ + Chr$(10) + Chr$(13)
       Next KKI%
       '
       '\\\OT-08-0098-12/03/08-///
       SQLX$ = "SELECT NUCLIEN,HISTORIAL FROM ANOTADIFCAMBIO"
       SQLX$ = SQLX$ + " WHERE RECIBO = '" & RECIBOX$ & "'"
       Dim rs As New ADODB.Recordset
       Set rs = New ADODB.Recordset
25     On Error Resume Next
       rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
       End If
       On Error GoTo 0
       '
       With rs
          If Not (.BOF And .EOF) Then
             ANOTA$ = !Historial
             ANOTA$ = ANOTA$ + TTXX$
             ANOTA$ = ANOTA$ + String$(88, "=") + Chr$(10) + Chr$(13)
             !Historial = ANOTA$
             .Update
           End If
       End With
       rs.Close
       Set rs = Nothing
       '
       Text6 = ""
       '
       List1.Clear
       Label7 = "": Label7 = RECIBOX$
       '
     End If
   End If
   Command9.Enabled = True
End Sub

Private Sub Form_Load()
    Call APLICACIONSKINS(Me)
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    Call ABRECONEXION
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If

    Call CREAR_TABLA_HojaTrabajoDifCambio
    Call CREAR_TABLA_ANOTADIFCAMBIO
    
    TXCAMPOS$ = "Cta/f10;R.Social/A48;N.Recibo/A24;Fecha/d8;Imp.Recibo $/f12.2;Imp.Aplic.$/F12.2;Nota Débito/A18;Dif.Camb.$/F12.2"
    Call CARGA_ENCABEZADO(MSF0, TXCAMPOS$, Me)
    TXCAMPOS$ = "N.Recibo/A12;Fecha/d8;Ch./A28;Vencimiento/d10;Imp. $/F10.2;Fecha de Ch./F11.2; Cobrados/F10.2;Factura:/a18;Fecha:/D8;Imp Dol./f10.2;Dif.Camb.U$S/F10.2;$ a Fecha/F10.2;APLICADO EN PESOS/F0"
    Call CARGA_ENCABEZADO(MSF1, TXCAMPOS$, Me)
    MSF1.Rows = 3
    '
    TXCAMPOS$ = "Cta/f10;R.Social/A48;N.Recibo/A24;Fecha/d8;$ a Fecha/f12.2;Nota Débito/A18;Dif.Camb.$/F12.2"
'    Call CARGA_ENCABEZADO(MSFPEN0, TXCAMPOS$, Me)
'    TXCAMPOS$ = "N.Recibo/A12;Fecha/d8;Ch./A20;Acredit/d8;Imp. $/F10.2;Fecha de Ch./F11.2; Cobrados/F10.2;Factura:/a18;Fecha:/D8;Imp Dol./f10.2;Dif.Camb.U$S/F10.2;$ a Fecha/F10.2"
'    Call CARGA_ENCABEZADO(MSFPEN1, TXCAMPOS$, Me)
'    MSFPEN1.Rows = 3
    '
    TXCAMPOS$ = "Cta/f10;R.Social/A48;N.Recibo/A24;Imp.Recibo/F12.2;Nota Débito/A18"
    Call CARGA_ENCABEZADO(MSFGEN0, TXCAMPOS$, Me)
    TXCAMPOS$ = "N.Recibo/A18;Fecha/d8;Ch./A20;Acredit/d8;Imp. $/F10.2;Fecha de Ch./F11.2; Cobrados/F10.2;Factura:/a18;Fecha:/D8;Imp Dol./f10.2;Dif.Camb.U$S/F10.2;$ a Fecha/F10.2"
    Call CARGA_ENCABEZADO(MSFGEN1, TXCAMPOS$, Me)
    MSFGEN1.Rows = 3

    'PASO EL ENCABEZADO A LA SEGUNDA FILA Y AGREGO OTRO TEXTO A LA PRIMERA PARA COMBINAR CELDAS DE LA FILA 0
    For i& = 1 To MSF1.Cols - 1
       MSF1.TextMatrix(1, i&) = MSF1.TextMatrix(0, i&)
       Select Case i&
           Case 1, 2
              MSF1.TextMatrix(0, i&) = "RECIBO"
           Case 3, 4, 5
              MSF1.TextMatrix(0, i&) = "VALOR ENTREGADO"
           Case 6
              MSF1.TextMatrix(0, i&) = "T.de Cambio a"
              Case 7
              MSF1.TextMatrix(0, i&) = "Dolares"
           Case 8, 9, 10
              MSF1.TextMatrix(0, i&) = "APLICADO A:"
           Case 11, 12
              MSF1.TextMatrix(0, i&) = "DIFERENCIA DE CAMBIO"
           Case Else
             MSF1.TextMatrix(0, i&) = " "
       End Select
    Next i&
    TXCAMPOS$ = "Cta/f10;R.Social/A28;N.Recibo/A24"
    Call CARGA_ENCABEZADO(MSFNOC, TXCAMPOS$, Me)
    
    TXCAMPOS$ = "Fecha/d8;Moneda/f7;Denomincacion/A20;Cotización/F12"
    Call CARGA_ENCABEZADO(MSFMON, TXCAMPOS$, Me)
   
        'PASO EL ENCABEZADO A LA SEGUNDA FILA Y AGREGO OTRO TEXTO A LA PRIMERA PARA COMBINAR CELDAS DE LA FILA 0
'    For I& = 1 To MSFPEN1.Cols - 1
'       MSFPEN1.TextMatrix(1, I&) = MSFPEN1.TextMatrix(0, I&)
'       Select Case I&
'           Case 1, 2
'              MSFPEN1.TextMatrix(0, I&) = "RECIBO"
'           Case 3, 4, 5
'              MSFPEN1.TextMatrix(0, I&) = "VALOR ENTREGADO"
'           Case 6
'              MSFPEN1.TextMatrix(0, I&) = "T.de Cambio a"
'              Case 7
'              MSFPEN1.TextMatrix(0, I&) = "Dolares"
'           Case 8, 9, 10
'              MSFPEN1.TextMatrix(0, I&) = "APLICADO A:"
'           Case 11, 12
'              MSFPEN1.TextMatrix(0, I&) = "DIFERENCIA DE CAMBIO"
'           Case Else
'             MSFPEN1.TextMatrix(0, I&) = " "
'       End Select
'    Next I&

'        'PASO EL ENCABEZADO A LA SEGUNDA FILA Y AGREGO OTRO TEXTO A LA PRIMERA PARA COMBINAR CELDAS DE LA FILA 0
    For i& = 1 To MSFGEN1.Cols - 1
       MSFGEN1.TextMatrix(1, i&) = MSFGEN1.TextMatrix(0, i&)
       Select Case i&
           Case 1, 2
              MSFGEN1.TextMatrix(0, i&) = "RECIBO"
           Case 3, 4, 5
              MSFGEN1.TextMatrix(0, i&) = "VALOR ENTREGADO"
           Case 6
              MSFGEN1.TextMatrix(0, i&) = "T.de Cambio a"
              Case 7
              MSFGEN1.TextMatrix(0, i&) = "Dolares"
           Case 8, 9, 10
              MSFGEN1.TextMatrix(0, i&) = "APLICADO A:"
           Case 11, 12
              MSFGEN1.TextMatrix(0, i&) = "DIFERENCIA DE CAMBIO"
           Case Else
             MSFGEN1.TextMatrix(0, i&) = " "
       End Select
    Next i&


    MSF1.MergeRow(0) = True
    MSF1.MergeCells = flexMergeFree
    MSF1.FixedRows = 2
    
    MSFGEN1.MergeRow(0) = True
    MSFGEN1.MergeCells = flexMergeFree
    MSFGEN1.FixedRows = 2
    '
    Me.TXTCOTIZACION = FORMATNUM$(COTIFECHA_HISTODOL#(HOY(HO$), 2), "F12.4")
    Text7 = Control("DIFECAMB", "HOJATRAB")
'    For II% = 0 To Label8.UBound
'       Label8(II%).Caption = CONTROL("DIFECAMB", "NOCONSI" & SAFETEXT$(II%))
'       Label5(II%) = ECOARCH$("ECUECON", Label8(II%))
'    Next II%

End Sub
Sub MENU_CELDA3(MSF As msFlexGrid, REG&, COL%)
        If MSF.Rows <= 1 Then Exit Sub
        
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL3
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        GLOBALFRM_FILA& = REG&
        PopupMenu MNUARCH3
        Call INVICTRL3
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA
End Sub

Sub INVICTRL()
  '
  mnuCopiar.Visible = Not (mnuCopiar.Visible)
  mnudesap.Visible = Not (mnudesap.Visible)
  '
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Call FINAL("")
End Sub

Private Sub Label11_Change()
    List1.Clear
    Label23 = rasocli$(XVALO(Label11))
End Sub

Private Sub Label7_Change()
   List1.Clear
   RECIBOX$ = TRIM$(Label7)
   If RECIBOX$ = "" Then Exit Sub
   '
   NC = XVALO(Label11)
      
20 Call APERBASDAT("PEDCLI")
   SQLX$ = "SELECT HISTORIAL,NUCLIEN FROM ANOTADIFCAMBIO WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE RECIBO = '" & RECIBOX$ & "'"
   If NC > 0 Then SQLX$ = SQLX$ + " AND NUCLIEN = " & NC
   
   Set rs = READSET(SQLX$)
   '
   With rs
    On Error Resume Next
    .MoveFirst
    If Err < 1 Then
        On Error GoTo 0
        
        XX$ = SAFETEXT$(!Historial): LONGI& = Len(XX$)
        Label11 = XVALO(!NUCLIEN)
        '
        Inicio& = 1
        Do While Inicio& < LONGI&
            FIN& = InStr(Inicio&, XX$, Chr$(10))
            If FIN& > 0 Then
               valor = 0
               If Inicio > 1 Then valor = 1
               TTXX$ = Mid$(XX$, Inicio& - valor, FIN& - Inicio&)
               List1.AddItem TTXX$
               Inicio& = FIN& + 2
            Else
               TTXX$ = Mid$(XX$, Inicio&)
               List1.AddItem TTXX$
               Exit Do
            End If
        Loop
    End If
    On Error GoTo 0
   End With
   '
   ' ESTO ES PARA POSICIONARSE AL FINAL DE LA LISTA
   TPXI& = List1.ListCount - 15
   If TPXI& < 0 Then TPXI& = 0
   List1.TopIndex = TPXI&
   '
   List1.Refresh
   Openforms = DoEvents
   rs.Close
   Set rs = Nothing
   '
   ' 21/01/2006: al consultar un pedido que recupere
   '             la consistencia de preparacion y despacho
   If REPARX% <> 1 Then ' NO ES UNA REPARACIóN
     Screen.MousePointer = 11
     HUBOREC% = 0
     Call VERIDESPRE(NPX&, 0, HUBORECU%)
     If HUBORECU% > 0 Then HUBOREC% = 1
     Screen.MousePointer = 1
     If HUBOREC% > 0 Then
       Call COMUNI("Se Han Recuperado Acumulados\de Preparación y Despacho.")
     End If
   End If
   '

End Sub



Private Sub mnucopiar_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub

Private Sub mnuctacte_Click()
    RECUEC07.Text1 = Me.TXTCLIENTE
    
    RECUEC07.Show 1
End Sub

Private Sub mnudesap_Click()
      REG& = GLOBALFRM_FILA&
      If REG& < 1 Then Exit Sub
      RECIBOX$ = TRIM$(MSFGEN0.TextMatrix(REG&, 3))
      NOTADEBGENERADAX$ = TRIM$(MSFGEN0.TextMatrix(REG&, 5))
30    OPX% = COMALTER%("Confirma Des-Aplicar Nota de Débito Generada a Recibo: " & RECIBOX$ & " \Esta Acción No Anula La Nota de Débito\Permitirá Recalcular la Diferencia de Cambio y Generar Otra Nota de Débito\\Cancelar\Des-Aplicar")
      If OPX% <= 1 Then Exit Sub
      If OPX% = 2 Then
            NDB$ = TRIM$(MSFGEN0.TextMatrix(REG&, 5))
            NC = XVALO(MSFGEN0.TextMatrix(REG&, 1))

            CONDI$ = "NCLIEN = " & NC
            CONDI$ = CONDI$ + " AND RECIBO = '" & RECIBOX$ & "'"
            CONDI$ = CONDI$ + " AND NOTADEBITOGENERADA =   '" & NOTADEBGENERADAX$ & "'"
            Call BORRAREGISTROS("HojaTrabajoDifCambio", CONDI$, REGAF&, "NOMESS")
            If REG& > 1 Then
              Call MSFGEN0.RemoveItem(REG&)
            Else
              MSFGEN0.Rows = 1
            End If
            '
            TXTX$ = "Se Des-Aplicó la Nota de Débito: " & NDEB$ & vbCrLf
            TXTX$ = TXTX$ + " Generada a Recibo: $ " & RECIBOX$
    
            Call ANOTADIFCAMBIO(NC, RECIBOX$, 0, TXTX$)

      End If
  
End Sub

Private Sub MNUTRAZABRECIBO_Click()
    RECICO07.Command11_Click
End Sub

Private Sub MSF0_Click()
   i& = MSF0.MouseRow
   J& = MSF0.MouseCol
   '
   If i& < 1 Then
    Static MODO  As Boolean
    If (MSF0.MouseRow = 0) Then
        ' Ordena en forma ascendente
        If MODO Then
            MSF0.COL = MSF0.MouseCol
            MSF0.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSF0.COL = MSF0.MouseCol
            MSF0.Sort = 1
            MODO = True
        End If
     End If
     Exit Sub
   End If
   '
   If J& < 1 Then Exit Sub
   On Error Resume Next
   RECIX$ = TRIM$(MSF0.TextMatrix(i&, 3))
   If RECIX$ = "" Then On Error GoTo 0: Exit Sub
   On Error GoTo 0
   NC = XVALO(MSF0.TextMatrix(i&, 1))
   Me.TXTNOTADEBITO = TRIM$(MSF0.TextMatrix(i&, 7))
   Call DESHABILITAR_CONTROLES
   Call AGREGAR_A_TABLA_HojaTrabajoDifCambio(RECIX$, NC, 0, "PENDIENTES")
   Call HABILITAR_CONTROLES
'   Call COLOREAR_GRILLASQL(MSF1, &HE0E0E0)

End Sub

Private Sub MSF1_DblClick()
    REG& = MSF1.MouseRow
    COL& = MSF1.MouseCol
    If REG& < 1 Then Exit Sub
    RECIBOX$ = MSF1.TextMatrix(REG&, 1)
    NUDOCU$ = MSF1.TextMatrix(REG&, 8)
    If COL& = 8 Then
       NC = XVALO(Me.TXTCLIENTE)
       If NC < 1 Then
          NC = VALOBADA("CAJABANC", "NUCLIEN", "CODICOM_COR = '" & RECIBOX$ & "'", "NOMESS")
       End If
       If Left$(NUDOCU$, 3) = "RB-" Then
          'CASO DE RECIBO CON UN RECIBO APLICADO EN EL MISMO (ANTICIPOS NO APLICADOS.)
          NUDOCU1$ = SAFETEXT$(XVALO(Mid(NUDOCU$, 11))): GoTo 30
       Else
         CONSIDERAR_MOVIVIE% = 1
         Call MUESTRASTSCREEN(NC, NUDOCU$)
         CONSIDERAR_MOVIVIE% = 0
       End If
    ElseIf COL& = 1 Then
       NUDOCU1$ = SAFETEXT$(XVALO(Mid(RECIBOX$, 4)))
30     FERECI$ = MSF1.TextMatrix(REG&, 2)
       DOCUNU& = DOCID&("Recibo", NUDOCU1$, NUDOCU1$, FERECI$)
       Call MUESTRADOC(DOCUNU&, "", "Recibo", NUDOCU1$, NUDOCU1$)
    End If
End Sub


Private Sub MSFGEN0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
     REG& = MSFGEN0.MouseRow
     COL% = MSFGEN0.MouseCol
     Call MENU_CELDA(MSFGEN0, REG&, COL%)
  End If
End Sub
Private Sub MSF0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
     REG& = MSF0.MouseRow
     COL% = MSF0.MouseCol
     Call MENU_CELDA2(MSF0, REG&, COL%)
  End If
End Sub

Private Sub MSF0_DblClick()
        Me.TXTCLIENTE.Enabled = False
        Unload FORFASQL 'SI NO DESCARGO EL FRM NO FUNCIONA LOS DOBLE CLICK HARDCODEADOS DEL LABEL7
        VUELTA% = 0
        i& = MSF0.MouseRow
        If i& < 1 Then Exit Sub

        NDB$ = TRIM$(MSF0.TextMatrix(i&, 7))
        If NDB$ <> "" Then
           FORIMPRESION$ = TRIM$(Control("DIFECAMB", "HOJATRAB"))
           If FORIMPRESION$ <> "" Then
              Call COMUNI("Falta Configurar Formulario de Hoja de Trabajo")
              SSTab1.Tab = 3
              GoTo 99
           End If
           
            Screen.MousePointer = 11
            RECIX$ = TRIM$(MSF0.TextMatrix(i&, 3))
            NC = XVALO(MSF0.TextMatrix(i&, 1))
            ARCSQR$ = FORIMPRESION$
            FILTX$ = NDB$ & ";" & TRIM$(Str$(NC)) & ";" & RECIX$
            Screen.MousePointer = 1
            Call STDFORM(ARCSQR$, FILTX$)
            
            GoTo 99
        End If
        Screen.MousePointer = 11
        NC = XVALO(MSF0.TextMatrix(i&, 1))
        ImporteRecibo = XVALO(MSF0.TextMatrix(i&, 5))
        ImpoAplicado = XVALO(MSF0.TextMatrix(i&, 6))
        IMPOR = XVALO(MSF0.TextMatrix(i&, 8))

        'VALIDO SI ES DIFERENTE EL IMPORTE APLICADO QUE EL TOTAL DEL RECIBO
        If Abs(ImporteRecibo - ImpoAplicado) > 0.9 Then
           OPX% = COMALTER%("Atención!! Importe Total del Recibo Es Superior al Importe Aplicado\\Cancelar\Continuar\Abrir Cta.Corriente")
           If OPX% <= 1 Then Exit Sub
           If OPX% = 3 Then RECUEC07.Text1 = NC: RECUEC07.Show 1: GoTo 99
        End If
        FORFASQL.Text1 = NC
        FORFASQL.CARDACLI
'        FORFASQL.Text1_KeyPress (13)
'        FORFASQL.Command6_Click
        Call T_Espera(0.3)
        RECIBOX$ = TRIM$(MSF0.TextMatrix(i&, 3))
        If Mid$(RECIBOX$, 4, 1) = "X" Then
          'DE FORMATO LARGO 'RB-X-0000-00001840' A FORMATO CORTO  RB-001840-00
          RECIBOCORX$ = Mid$(RECIBOX$, 13):
          RECIBOCORX$ = "RB-" & RECIBOCORX$ & "-00"
        End If
        If IMPOR > 0 Then
          'NOTA DEBITO
40        Call FORFASQL.Label7_Click: T_Espera (0.5)
          If TRIM$(FORFASQL.Label7) <> "N.Débito:" Then
            VUELTA% = VUELTA% + 1
            If VUELTA% < 10 Then GoTo 40 Else Me.TXTCLIENTE.Enabled = True: Exit Sub
          End If
          DETALLEX$ = "Dif Cambio s/ Rec: " + RECIBOCORX$
        Else
50        Call FORFASQL.Label7_Click: Call T_Espera(0.3)
          If TRIM$(FORFASQL.Label7) <> "N.Crédito:" Then
            VUELTA% = VUELTA% + 1
            If VUELTA% < 10 Then GoTo 50 Else Me.TXTCLIENTE.Enabled = True: Exit Sub
          End If
          Call FORFASQL.Label6_Click
          DETALLEX$ = "Dif Cambio s/ Rec: " + RECIBOCORX$
'          DETALLEX$ = "Nota de Crédito x Dif. Cambio (Recibo: " + RECIBOX$ & ")"
        End If
        '
        'LE QUITO EL IVA
        IMPOR = Abs(IMPOR): IMPOR = IMPOR / 1.21
        TOT# = Abs(IMPOR) ' PARA REPETIR EL IMPORTE
        '
        TX$ = REGBLAN$("!CARDEFAC")
        '
        Call SETULTREG("!CARDEFAC", 0)
        Call REPLA(TX$, MKS$(1), 1, 4)
        Call REPLA(TX$, DETALLEX$, 7, 64)
        Call REPLA(TX$, MKS(IMPOR), 429, 4)
        Call REPLA(TX$, MKD(TOT#), 433, 8)
        Call REGAPP("!CARDEFAC", TX$)
        Call REPLA(TX$, MKS$(PORCEIVA(CI%)), 441, 4)
        Call SETULTREG("!CARDEFAC", 1)
        Call FORFASQL.CARDETFAC
'         FORFASQL.Command6_Click
        FORFASQL.Show 1
        '
        OKX% = 1
'        ' Advertencia por que es posible que el total de la factura difiera segun tasa de iva
'        Call COMUNI("Verifique Que la Tasa de Iva \Aplicada a Esta Nota de Débito Sea la Correcta  ")
99      Screen.MousePointer = 1
        Me.TXTCLIENTE.Enabled = True
End Sub


Private Sub MSF0_SelChange()
   i& = MSF0.MouseRow
   If i& < 1 Then Exit Sub
   '
   If J& <> 1 Then Exit Sub
   
   RECIX$ = TRIM$(MSF0.TextMatrix(i&, 3))
   NC = XVALO(MSF0.TextMatrix(i&, 1))
   Call AGREGAR_A_TABLA_HojaTrabajoDifCambio(RECIX$, NC)

End Sub


Private Sub Text4_DblClick()
   Call Command3_Click
End Sub

Private Sub Text4_LostFocus()
    Text4 = FORMATNUM$(XVALO(Text4), "F12.2")
End Sub


Private Sub MSFGEN0_Click()
   i& = MSFGEN0.MouseRow
   J& = MSFGEN0.MouseCol
   
   If i& < 1 Then Exit Sub
   '
   If J& < 1 Then Exit Sub
   
   MSFGEN1.Rows = 2
   On Error Resume Next
   RECIX$ = TRIM$(MSFGEN0.TextMatrix(i&, 3))
   If Err > 0 Then
     On Error GoTo 0
     Exit Sub
   End If
   NC = XVALO(MSFGEN0.TextMatrix(i&, 1))
   
    Me.TXTNOTADEBITO.TEXT = TRIM$(MSFGEN0.TextMatrix(i&, 5))
'   Call AGREGAR_A_TABLA_HojaTrabajoDifCambio(RECIX$, NC, 0, "GENERADAS")
    NDEB$ = TRIM$(MSFGEN0.TextMatrix(i&, 5))
    
    ACUM_IMPOMEPAGPES# = 0
    ACU_DOLCOBRA# = 0
    ACU_IMPODOLFACTU# = 0
    i& = 1
    
    SQLX$ = " SELECT * FROM HojaTrabajoDifCambio WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NOTADEBITOGENERADA = '" & NDEB$ & "'"
    SQLX$ = SQLX$ + " AND NCLIEN = " & NC
    SQLX$ = SQLX$ + " AND RECIBO = '" & RECIX$ & "'"
    SQLX$ = SQLX$ + " ORDER BY ORDEN"
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
          
            RECIBOX$ = SAFETEXT$(!recibo)
            FERECI% = CVINT(!FechaRecibo)
            MEDPAGO$ = SAFETEXT$(!MedioPago)
            FEACREDPAG% = CVINT(!FechaAcreditPago)
            IMPOMEPAGPES# = XVALO(!ImpoMedioPagoPesos): ACUM_IMPOMEPAGPES# = ACUM_IMPOMEPAGPES# + IMPOMEPAGPES#
            TCAMBIOFEACRED# = XVALO(!TCambioFechaAcredit)
            DOLCOBRA# = XVALO(!DolaresCobrados): ACU_DOLCOBRA# = ACU_DOLCOBRA# + DOLCOBRA#
            FACTAPL$ = SAFETEXT$(!FACTURA_APL)
            FEFACTURA% = CVINT(!FECHA_FACTURA_APL)
            DIFCAMBIOPESOS# = XVALO(!DifCambioEnPesos)
            IMPODOLFACTU# = XVALO(!ImpoDolSobreFactura): ACU_IMPODOLFACTU# = ACU_IMPODOLFACTU# + IMPODOLFACTU#
            i& = i& + 1
            MSFGEN1.Rows = i& + 1
            MSFGEN1.TextMatrix(i&, 1) = RECIBOX$
            MSFGEN1.TextMatrix(i&, 2) = FECHATEX$(FERECI%)
            MSFGEN1.TextMatrix(i&, 3) = MEDPAGO$
            MSFGEN1.TextMatrix(i&, 4) = FECHATEX$(FEACREDPAG%)
            MSFGEN1.TextMatrix(i&, 5) = FORMATNUM$(IMPOMEPAGPES#, "F12.2")
            MSFGEN1.TextMatrix(i&, 6) = FORMATNUM$(TCAMBIOFEACRED#, "F12.2")
            MSFGEN1.TextMatrix(i&, 7) = FORMATNUM$(DOLCOBRA#, "F12.2")
            MSFGEN1.TextMatrix(i&, 8) = FACTAPL$
            MSFGEN1.TextMatrix(i&, 9) = FECHATEX$(FEFACTURA%)
            MSFGEN1.TextMatrix(i&, 10) = FORMATNUM$(IMPODOLFACTU#, "F12.2")
            MSFGEN1.TextMatrix(i&, 11) = ""
           .MoveNext
           Loop
      End With
      i& = i& + 1
      MSFGEN1.Rows = i& + 1
      MSFGEN1.TextMatrix(i&, 0) = "||"
      MSFGEN1.TextMatrix(i&, 5) = "___________"
      MSFGEN1.TextMatrix(i&, 7) = "___________"
      MSFGEN1.TextMatrix(i&, 10) = "___________"
      
      i& = i& + 1
      MSFGEN1.Rows = i& + 1
      MSFGEN1.TextMatrix(i&, 5) = FORMATNUM$(ACU_IMPOMEPAGPES#, "F12.2")
      MSFGEN1.TextMatrix(i&, 7) = FORMATNUM$(ACU_DOLCOBRA#, "F12.2")
      MSFGEN1.TextMatrix(i&, 10) = FORMATNUM$(ACU_IMPODOLFACTU#, "F12.2")
      MSFGEN1.TextMatrix(i&, 11) = FORMATNUM$(ACU_IMPODOLFACTU# - ACU_DOLCOBRA#, "F12.2")
      MSFGEN1.TextMatrix(i&, 12) = FORMATNUM$(DIFCAMBIOPESOS#, "F12.2")

   
   
   
   
   
   
   
   
   Call COLOREAR_GRILLASQL(MSFGEN1, &HE0E0E0)

End Sub

Private Sub MSFPEN0_Click()
   i& = MSFPEN0.Row
   J& = MSFPEN0.COL
   '
   If i& < 1 Then Exit Sub
   '
   If J& <> 1 Then Exit Sub
   
   RECIX$ = TRIM$(MSFPEN0.TextMatrix(i&, 3))
   NC = XVALO(MSFPEN0.TextMatrix(i&, 1))
   Me.TXTNOTADEBITO = TRIM$(MSFPEN0.TextMatrix(i&, 7))
   Call AGREGAR_A_TABLA_HojaTrabajoDifCambio(RECIX$, NC)
'   Call COLOREAR_GRILLASQL(MSFPEN1, &HE0E0E0)

End Sub

Private Sub MSFGEN0_DblClick()
        MSFGEN0.Enabled = False
        MSFGEN1.Enabled = False
        Me.TXTCLIENTE.Enabled = False
        
        i& = MSFGEN0.MouseRow
        If i& < 1 Then Exit Sub
        On Error Resume Next
        NDB$ = TRIM$(MSFGEN0.TextMatrix(i&, 5))
        If Err > 0 Then
          On Error GoTo 0
          Exit Sub
        End If
        If NDB$ <> "" Then
           FORIMPRESION$ = TRIM$(Control("DIFECAMB", "HOJATRAB"))
           If FORIMPRESION$ = "" Then
              Call COMUNI("Falta Configurar Formulario de Hoja de Trabajo")
              SSTab1.Tab = 3
              GoTo 99
           End If

            RECIX$ = TRIM$(MSFGEN0.TextMatrix(i&, 3))
            NC = XVALO(MSFGEN0.TextMatrix(i&, 1))
            ARCSQR$ = FORIMPRESION$
            FILTX$ = NDB$ & ";" & TRIM$(Str$(NC)) & ";" & RECIX$
            Call STDFORM(ARCSQR$, FILTX$)

            GoTo 99
        End If

'        NC = XVALO(MSFGEN0.TextMatrix(I&, 1))
'        ImporteRecibo = XVALO(MSFGEN0.TextMatrix(I&, 5))
'        ImpoAplicado = XVALO(MSFGEN0.TextMatrix(I&, 6))
'        'VALIDO SI ES DIFERENTE EL IMPORTE APLICADO QUE EL TOTAL DEL RECIBO
'        If Abs(ImporteRecibo - ImpoAplicado) > 0.9 Then
'           OPX% = COMALTER%("Atención!! Importe Total del Recibo Es Superior al Importe Aplicado\\Cancelar\Continuar\Abrir Cta.Corriente")
'           If OPX% <= 1 Then Exit Sub
'           If OPX% = 3 Then RECUEC07.Text1 = NC: RECUEC07.Show 1: GoTo 99
'        End If
'
'        FORFASQL.Text1 = NC
'        FORFASQL.Label7_Click
'        RECIBOX$ = TRIM$(MSFGEN0.TextMatrix(I&, 3))
'
'        IMPOR = XVALO(MSFGEN0.TextMatrix(I&, 5))
'        TOT# = IMPOR ' PARA REPETIR EL IMPORTE
'        '
'        TX$ = REGBLAN$("!CARDEFAC")
'        DETALLEX$ = "Nota de Débito x Dif. Cambio (Recibo: " + RECIBOX$ & ")"
'        '
'        Call SETULTREG("!CARDEFAC", 0)
'        Call REPLA(TX$, MKS$(1), 1, 4)
'        Call REPLA(TX$, DETALLEX$, 7, 64)
'        Call REPLA(TX$, MKS(IMPOR), 429, 4)
'        Call REPLA(TX$, MKD(TOT#), 433, 8)
'        Call REGAPP("!CARDEFAC", TX$)
'        Call REPLA(TX$, MKS$(PORCEIVA(CI%)), 441, 4)
'        Call SETULTREG("!CARDEFAC", 1)
'        Call FORFASQL.CARDETFAC
''         FORFASQL.Command6_Click
'        FORFASQL.Show 1
'        '
'        OKX% = 1
'        ' Advertencia por que es posible que el total de la factura difiera segun tasa de iva
'        Call COMUNI("Verifique Que la Tasa de Iva \Aplicada a Esta Nota de Débito Sea la Correcta  ")
99      MSFGEN0.Enabled = True
        MSFGEN1.Enabled = True
        Me.TXTCLIENTE.Enabled = True

End Sub

Private Sub MSFMON_DblClick()
        REG& = MSFMON.MouseRow
        COL& = MSFMON.MouseCol
        If REG& < 1 Then Exit Sub
        IDMON% = XVALO(MSFMON.TextMatrix(REG&, 2))
        If IDMON% < 1 Then Exit Sub
        FE% = FECHANUM(MSFMON.TextMatrix(REG&, 1))
        PARAM_COTIZA# = XVALO(MSFMON.TextMatrix(REG&, 4))
        
        Call ACTUALIZAR_COTIZACION(2, FE%, PARAM_COTIZA#)
        CONDI$ = "  FECHA = '" & FETEXCOR1$(FE%) & "' AND ID = " & IDMON%
        COTIZA# = XVALO(VALOBADA("HISTOCAMBIO", "COTIZA_DOLAR", CONDI$, "NOMESS"))
        MSFMON.TextMatrix(REG&, 4) = FORMATNUM$(COTIZA#, "F12.4")
        
End Sub


Private Sub Text7_Change()
    CODFOR$ = TRIM$(Text7)
    Label188 = ECOARCH$("INDISQR", CODFOR$)

End Sub

Private Sub TXTCLIENTE_Change()
  
  Me.TXTRASOCLI = rasocli$(XVALO(Me.TXTCLIENTE))
  Call T_Espera(0.3)
  Static NCLIE
  If NCLIE <> XVALO(TXTCLIENTE) And XVALO(TXTCLIENTE) > 0 Then
    Call TXTCLIENTE_KeyPress(13)
  End If
  NCLIE = XVALO(TXTCLIENTE)
End Sub

Private Sub TXTCLIENTE_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      Call DESHABILITAR_CONTROLES
      MSF0.Rows = 1
      MSF1.Rows = 2
      Screen.MousePointer = 11
      Call GENERACION_NOTAS_X_DIF_CAMBIO("POR CLIENTE", "PENDIENTES")
      Call COLOREAR_GRILLASQL(MSF0, &HE0E0E0)
      Call COLOREAR_GRILLASQL(MSF1, &HE0E0E0)
      Screen.MousePointer = 1
      Call HABILITAR_CONTROLES
   End If
End Sub
Sub MENU_CELDA2(MSF As msFlexGrid, REG&, COL%)
        If MSF.Rows <= 1 Then Exit Sub
        
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL2
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        GLOBALFRM_FILA& = REG&
        PopupMenu MNUARCH2
        Call INVICTRL2
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA
End Sub

Sub INVICTRL2()
  '
  mnuquitarlistapendientes.Visible = Not (mnuquitarlistapendientes.Visible)
  mnucopiar2.Visible = Not (mnucopiar2.Visible)
  '
End Sub

Private Sub mnucopiar2_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub
Private Sub mnuquitarlistapendientes_Click()
      REG& = GLOBALFRM_FILA&
      If REG& < 1 Then Exit Sub
      RECIBOX$ = TRIM$(MSF0.TextMatrix(REG&, 3))
30    OPX% = COMALTER%("Confirma Suprimir de Manera Permanente el Cálculo\ Por Diferencia de Cambio del Recibo: " & RECIBOX$ & " \\Cancelar\Suprimir del Cálculo")
      If OPX% <= 1 Then Exit Sub
      If OPX% = 2 Then
            NC = XVALO(MSF0.TextMatrix(REG&, 1))
            CONDI$ = "NUCLIEN = " & NC
            CONDI$ = CONDI$ + " AND CODICOM_LAR = '" & RECIBOX$ & "'"
            Call ACTUREGISTRO("CAJABANC", "Genera_Debito", CONDI$, -1, REGAF&, "NOMESS")
            Call MSF0.RemoveItem(REG&)
            TXTX$ = "Se Suprime del Cálculo por Diferencia de Cambio"
            Call ANOTADIFCAMBIO(NC, RECIBOX$, 0, TXTX$)

      End If
  
End Sub

'================================================================================================
Private Sub MSFNOC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
     REG& = MSFNOC.MouseRow
     COL% = MSFNOC.MouseCol
     Call MENU_CELDA3(MSFNOC, REG&, COL%)
  End If
End Sub

Sub MENU_CELDA(MSF As msFlexGrid, REG&, COL%)
        If MSF.Rows <= 1 Then Exit Sub
        
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, REG&, Int(COL%))   ' LA PINTA DE AMARILLA
        Call INVICTRL
        GLOBALFRM_strContenidoCelda$ = TRIM$(MSF.TextMatrix(REG&, COL%))  'CELDA SELECCIONADA
        GLOBALFRM_FILA& = REG&
        PopupMenu mnuArchivo
        Call INVICTRL
        On Error Resume Next
        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbWhite, REG&, COL%)  ' LA PINTA DE AMARILLA
        On Error GoTo 0
End Sub

Sub INVICTRL3()
  '
  mnuvolveraincluir.Visible = Not (mnuvolveraincluir.Visible)
  mnucopiar3.Visible = Not (mnucopiar3.Visible)
  '
End Sub

Private Sub mnucopiar3_Click()
       Clipboard.Clear
       Clipboard.SetText GLOBALFRM_strContenidoCelda$
End Sub
Private Sub mnuvolveraincluir_Click()
      REG& = GLOBALFRM_FILA&
      If REG& < 1 Then Exit Sub
      RECIBOX$ = TRIM$(MSFNOC.TextMatrix(REG&, 3))
30    OPX% = COMALTER%("Confirma Volver a Incluir para el Cálculo\ Por Diferencia de Cambio del Recibo: " & RECIBOX$ & " \\Cancelar\Incluir en el Cálculo")
      If OPX% <= 1 Then Exit Sub
      If OPX% = 2 Then
            NC = XVALO(MSFNOC.TextMatrix(REG&, 1))
            CONDI$ = "NUCLIEN = " & NC
            CONDI$ = CONDI$ + " AND CODICOM_LAR = '" & RECIBOX$ & "'"
            Call ACTUREGISTRO("CAJABANC", "Genera_Debito", CONDI$, 0, REGAF&, "NOMESS")
            If REG& > 1 Then
               Call MSFNOC.RemoveItem(REG&)
            Else
               MSFNOC.Rows = 1
            End If
            TXTX$ = "Se Vuelve a Incluir en el Cálculo por Diferencia de Cambio"
            Call ANOTADIFCAMBIO(NC, RECIBOX$, 0, TXTX$)
      End If
End Sub


Private Sub TXTCOTIZACION_DblClick()
   Call Command3_Click
End Sub



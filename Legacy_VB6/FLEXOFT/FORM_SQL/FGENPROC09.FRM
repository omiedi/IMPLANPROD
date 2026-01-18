VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FGENPROC09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generador de Procesos FLEXOFT"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7365
      Left            =   105
      TabIndex        =   4
      Top             =   840
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   12991
      _Version        =   327681
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "FGENPROC09.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Encabezado"
      TabPicture(1)   =   "FGENPROC09.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Grilla"
      TabPicture(2)   =   "FGENPROC09.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   6420
         Left            =   105
         TabIndex        =   5
         Top             =   630
         Width           =   11450
         Begin VB.TextBox TXTCANREGI 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   4830
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   42
            Top             =   5145
            Width           =   6250
         End
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   9185
            ScaleHeight     =   75
            ScaleWidth      =   900
            TabIndex        =   40
            Top             =   3360
            Width           =   960
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   41
               Top             =   105
               Width           =   12000
            End
         End
         Begin VB.ListBox List2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1485
            IntegralHeight  =   0   'False
            ItemData        =   "FGENPROC09.frx":0054
            Left            =   275
            List            =   "FGENPROC09.frx":0056
            TabIndex        =   39
            Top             =   4140
            Width           =   4425
         End
         Begin VB.Frame Frame2 
            Caption         =   "Presentacion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3060
            Left            =   8190
            TabIndex        =   24
            Top             =   315
            Width           =   3060
            Begin VB.TextBox Text6 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Left            =   1470
               TabIndex        =   34
               Top             =   1935
               Width           =   630
            End
            Begin VB.TextBox Text5 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Left            =   210
               TabIndex        =   33
               Top             =   1935
               Width           =   1170
            End
            Begin VB.Frame Frame4 
               Caption         =   "Pos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   640
               Left            =   105
               TabIndex        =   28
               Top             =   315
               Width           =   2850
               Begin VB.TextBox Text4 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   250
                  Left            =   1290
                  TabIndex        =   30
                  Top             =   240
                  Width           =   600
               End
               Begin VB.TextBox Text3 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   250
                  Left            =   340
                  TabIndex        =   29
                  Top             =   260
                  Width           =   600
               End
               Begin VB.Label Label11 
                  Caption         =   "V:"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   1050
                  TabIndex        =   32
                  Top             =   270
                  Width           =   795
               End
               Begin VB.Label Label10 
                  Caption         =   "H:"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   105
                  TabIndex        =   31
                  Top             =   280
                  Width           =   795
               End
            End
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
               Height          =   435
               Left            =   1470
               Picture         =   "FGENPROC09.frx":0058
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   2520
               Width           =   540
            End
            Begin VB.TextBox Text2 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Left            =   210
               TabIndex        =   26
               Top             =   1365
               Width           =   1905
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   280
               Left            =   1260
               TabIndex        =   25
               Top             =   1005
               Width           =   855
            End
            Begin VB.Label Label12 
               Caption         =   "Form:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1365
               TabIndex        =   38
               Top             =   1710
               Width           =   645
            End
            Begin VB.Label Label5 
               Caption         =   "Eco:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   105
               TabIndex        =   37
               Top             =   1710
               Width           =   645
            End
            Begin VB.Label Label7 
               Caption         =   "Label:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   105
               TabIndex        =   36
               Top             =   1140
               Width           =   645
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "Form:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   420
               TabIndex        =   35
               Top             =   1050
               Width           =   795
            End
         End
         Begin VB.CommandButton Command5 
            Height          =   540
            Left            =   7665
            Picture         =   "FGENPROC09.frx":0362
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   2310
            Width           =   330
         End
         Begin VB.CommandButton Command4 
            Height          =   540
            Left            =   7665
            Picture         =   "FGENPROC09.frx":066C
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   1155
            Width           =   330
         End
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2120
            IntegralHeight  =   0   'False
            ItemData        =   "FGENPROC09.frx":0976
            Left            =   5670
            List            =   "FGENPROC09.frx":0978
            TabIndex        =   21
            Top             =   950
            Width           =   1905
         End
         Begin VB.CommandButton CMDBORRAR 
            Caption         =   "Borrar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Left            =   10185
            TabIndex        =   20
            Top             =   3765
            Width           =   870
         End
         Begin VB.CommandButton EJECUTAR 
            Caption         =   "Ejecutar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   370
            Left            =   9135
            TabIndex        =   19
            Top             =   3645
            Width           =   960
         End
         Begin VB.CommandButton GUARDACON 
            Caption         =   "Guardar Consulta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   8190
            TabIndex        =   18
            Top             =   3255
            Width           =   855
         End
         Begin VB.CommandButton CMDPEGAR 
            Caption         =   "Pegar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Left            =   10185
            TabIndex        =   17
            Top             =   3510
            Width           =   870
         End
         Begin VB.CommandButton CMDCOPIAR 
            Caption         =   "Copiar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   250
            Left            =   10185
            TabIndex        =   16
            Top             =   3255
            Width           =   870
         End
         Begin VB.CommandButton Command1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   7505
            Picture         =   "FGENPROC09.frx":097A
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Agregar Condición"
            Top             =   3255
            Width           =   540
         End
         Begin VB.TextBox TXTCONSULTA 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   960
            Left            =   4830
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   4140
            Width           =   6250
         End
         Begin VB.ListBox LSTMOSTARCMP 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2120
            IntegralHeight  =   0   'False
            ItemData        =   "FGENPROC09.frx":0C84
            Left            =   3255
            List            =   "FGENPROC09.frx":0C86
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   13
            Top             =   945
            Width           =   1905
         End
         Begin VB.ListBox LSTTABLAS 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2120
            IntegralHeight  =   0   'False
            Left            =   420
            Sorted          =   -1  'True
            TabIndex        =   12
            Top             =   945
            Width           =   2325
         End
         Begin VB.Frame Frame5 
            Caption         =   "Condición"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   315
            TabIndex        =   6
            Top             =   3150
            Width           =   7750
            Begin VB.CommandButton Command3 
               Caption         =   "?"
               Height          =   300
               Left            =   6825
               TabIndex        =   11
               ToolTipText     =   "Caracter Comodin"
               Top             =   315
               Width           =   260
            End
            Begin VB.ComboBox CMBCAMPOS 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   945
               Sorted          =   -1  'True
               TabIndex        =   10
               Text            =   "CMBCAMPOS"
               Top             =   315
               Width           =   1710
            End
            Begin VB.ComboBox CMBOPERADOR 
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
               ItemData        =   "FGENPROC09.frx":0C88
               Left            =   2835
               List            =   "FGENPROC09.frx":0C8A
               TabIndex        =   9
               TabStop         =   0   'False
               Text            =   "CMBOPERADOR"
               Top             =   315
               Width           =   1080
            End
            Begin VB.CommandButton Command2 
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
               Height          =   300
               Left            =   6435
               TabIndex        =   8
               ToolTipText     =   "Lista de Valores Posibles"
               Top             =   315
               Width           =   330
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   4095
               TabIndex        =   7
               Text            =   "Combo1"
               Top             =   315
               Width           =   2235
            End
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Orden:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4620
            TabIndex        =   45
            Top             =   705
            Width           =   1275
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Campos:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2415
            TabIndex        =   44
            Top             =   705
            Width           =   1275
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Tabla:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   100
            TabIndex        =   43
            Top             =   705
            Width           =   600
         End
      End
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      TabIndex        =   3
      ToolTipText     =   "Doble Click para Nueva Consulta"
      Top             =   105
      Width           =   1695
   End
   Begin VB.Label Label1 
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
      Left            =   945
      TabIndex        =   2
      Top             =   360
      Width           =   2325
   End
   Begin VB.Label Label9 
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
      Height          =   550
      Left            =   3465
      TabIndex        =   1
      Top             =   105
      Width           =   4635
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   -210
      TabIndex        =   0
      ToolTipText     =   "Doble Click para Nueva Consulta"
      Top             =   420
      Width           =   1065
   End
End
Attribute VB_Name = "FGENPROC09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim KLIK%
Dim MODIFENCA%, ILIS1&
'
Private Sub CMDBORRAR_Click()
TXTCONSULTA.TEXT = ""
End Sub


Private Sub label1_Change()
    Dim FISELE$(128)
    List2.Clear
    
    CANFI% = 0
    STRSQL = "SELECT * FROM SQL_PROCEDS where COD_PROCEX = '" & Label1.Caption & "'"
    Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenSnapshot)
    If Not (MIRS.EOF And MIRS.BOF) Then
       MIRS.MoveFirst
       Label9 = MIRS!DES_PROCEX
       TTXX$ = MIRS!PAR_PROCEX
       '
       While Len(TTXX$) > 0
          PPXX% = InStr(TTXX$, Chr$(124))
          If PPXX% < 1 Then PPXX% = 1 + Len(TTXX$)
          COMAI$ = Left$(TTXX$, PPXX% - 1)
          TTXX$ = Mid$(TTXX$, PPXX% + 1)
          '
          If Left$(COMAI$, 10) = "DATABASE =" Then
              Call CMBBASE_CLICK
              CMBBASE = UCase$(TRIM$(Mid$(COMAI$, 11)))
              GoTo 19
            ElseIf Left$(COMAI$, 10) = "SELTABLE =" Then
              TASELE$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
              For KKI% = 1 To LSTTABLAS.ListCount
                 If UCase$(LSTTABLAS.List(KKI% - 1)) = TASELE Then
                   LSTTABLAS.Selected(KKI% - 1) = True
                   Call LSTTABLAS_Click
                   GoTo 19
                 End If
              Next KKI%
            ElseIf Left$(COMAI$, 10) = "SELFIELD =" Then
              CANFI% = CANFI% + 1
              FISELE$(CANFI%) = TRIM$(Mid$(COMAI$, 11))
              CASELE$ = UCase$(TRIM$(Mid$(COMAI$, 11, 64)))
              For KKI% = 1 To LSTMOSTARCMP.ListCount
                 If UCase$(TRIM$(Left$(LSTMOSTARCMP.List(KKI% - 1), 64))) = CASELE$ Then
                   LSTMOSTARCMP.RemoveItem KKI% - 1
                   LSTMOSTARCMP.AddItem TRIM$(Mid$(COMAI$, 11, 64)), KKI% - 1
                   LSTMOSTARCMP.Selected(KKI% - 1) = True
                   GoTo 19
                 End If
              Next KKI%
            ElseIf Left$(COMAI$, 10) = "CONDITIO =" Then
              List2.AddItem TRIM$(Mid$(COMAI$, 11))
            Else
          End If
19     Wend
       '
       List1.Clear
       For KKI% = 1 To CANFI%
          List1.AddItem FISELE$(KKI%)
       Next KKI%
       '
       Call ARMA_CONSULTX
       Exit Sub
       '
    End If
    '
    Call LSTTABLAS_Click
    '
End Sub

Private Sub Command1_Click()
   If CMBCAMPOS.TEXT = "" Or CMBOPERADOR = "" Then
      MsgBox "FALTAN DATOS PARA LA CONDICION"
      Exit Sub
   End If
   If Len(TRIM$(CMBCAMPOS.TEXT)) > 24 Then
      MsgBox "Nombre de Campo Demasiado Largo"
      Exit Sub
   End If
   SQLY$ = AJUSTI$(CMBCAMPOS.TEXT, 24) + " " + AJUSTI$(CMBOPERADOR.TEXT, 8)
   VALOCO$ = TRIM$(Combo1.TEXT)
   TPC = TIPODECAMPO(LSTTABLAS.TEXT, CMBCAMPOS.TEXT)
   If (TPC = 52 Or TPC = 62) And VALOCO$ = "" Then
      MsgBox "FALTAN DATOS PARA LA CONDICION"
      Exit Sub
   End If
   '
   If TPC <> 52 And TPC <> 62 Then VALOCO$ = "'" + VALOCO$ + "'"
   SQLY$ = SQLY$ + " " + VALOCO$
   List2.AddItem SQLY$
   Call ARMA_CONSULTX
   Exit Sub
   
      If KLIK% > 0 Then GoTo 25
      SQLX$ = TXTCONSULTA.TEXT + vbCrLf
      SQLX$ = SQLX$ + " Where " & CMBCAMPOS.TEXT
      SQLX$ = SQLX$ + " " + CMBOPERADOR.TEXT
        VALOCO$ = Combo1.TEXT
        TPC = TIPODECAMPO(LSTTABLAS.TEXT, CMBCAMPOS.TEXT)
        If TPC <> 52 And TPC <> 62 Then
           VALOCO$ = "'" + VALOCO$ + "'"
        End If
      SQLX$ = SQLX$ + " " + VALOCO$
      KLIK% = KLIK% + 1
      TXTCONSULTA.TEXT = SQLX$
      Exit Sub
      
25    SQLX$ = TXTCONSULTA.TEXT + vbCrLf
      SQLX$ = SQLX$ + " AND " & CMBCAMPOS.TEXT
      SQLX$ = SQLX$ + " " + CMBOPERADOR.TEXT
        VALOCO$ = Combo1.TEXT
        TPC = TIPODECAMPO(LSTTABLAS.TEXT, CMBCAMPOS.TEXT)
        If TPC <> 52 And TPC <> 62 Then
           VALOCO$ = "'" + VALOCO$ + "'"
        End If
      SQLX$ = SQLX$ + " " + VALOCO$
      TXTCONSULTA.TEXT = SQLX$
      
End Sub

Function TIPODECAMPO(NTABLA$, NCAMPO$)
    Dim MIRS As ADODB.Recordset
    Set MIRS = New ADODB.Recordset
    STRSQL = "Select systypes.name, systypes.xtype from sysobjects inner join syscolumns on "
    STRSQL = STRSQL & "sysobjects.id = syscolumns.id inner join systypes "
    STRSQL = STRSQL & "on syscolumns.xtype = systypes.xtype "
    STRSQL = STRSQL & "where sysobjects.name = '" & NTABLA$ & "' "
    STRSQL = STRSQL & "and syscolumns.name ='" & NCAMPO$ & "'"
    Set MIRS =  FLEXCONN.EXECUTE (FILTSQL$ (STRSQL))
    If Not (MIRS.EOF And MIRS.BOF) Then
        TIPODECAMPO = MIRS!xtype
    End If

End Function

Private Sub CMDCOPIAR_Click()
   Clipboard.SetText Mid$(TXTCONSULTA.TEXT, TXTCONSULTA.SelStart + 1, Text1.SelLength)

End Sub

Private Sub CORTAR_Click()
   COMIE = TXTCONSULTA.SelStart
   LOSEL = TXTCONSULTA.SelLength
   TXTCONSULTA.TEXT = Left$(TXTCONSULTA.TEXT, COMIE) + Mid$(TXTCONSULTA.TEXT, COMIE + LOSEL + 1)

End Sub


Private Sub Command2_Click()
   If CMBCAMPOS.TEXT <> "" Then
      Screen.MousePointer = 11
      SQLX$ = "SELECT DISTINCT " & CMBCAMPOS.TEXT & " AS VALOR FROM " & LSTTABLAS.TEXT & " WITH(NOLOCK)"
      SQLX$ = SQLX$ + " ORDER BY " & CMBCAMPOS.TEXT & " ASC"
      '
      Combo1.Clear
      Dim VAPOSI As ADODB.Recordset
      Set VAPOSI = New ADODB.Recordset
      On Error Resume Next
      Set VAPOSI =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
      If Err Then
         On Error GoTo 0
         Call MENSERR(24, "Imposible Ejecutar Consulta - Guarde y Re-Intente")
         Exit Sub
      End If
      On Error GoTo 0
      With VAPOSI
         Do While Not .EOF
            On Error Resume Next
            Combo1.AddItem !VALOR
            On Error GoTo 0
         .MoveNext
         Loop
      End With
      Screen.MousePointer = 1
      '
   End If
   
End Sub

Private Sub Command3_Click()
10  COMO$ = InputBox$("Leyenda (Prompt)", "Configuración Consultas SQL", CMBCAMPOS.TEXT)
    If COMO$ = "" Then GoTo 10
    Combo1.Clear
    Combo1.AddItem ("¿" + COMO$ + "?")
    Combo1.TEXT = "¿" + COMO$ + "?"
End Sub

Private Sub Command4_Click()
   For I& = 1 To List1.ListCount
     If List1.Selected(I& - 1) = True Then
       If I& > 1 Then
         CAMPOSE$ = List1.List(I& - 1)
         List1.RemoveItem (I& - 1)
         List1.AddItem CAMPOSE$, I& - 2
         List1.Selected(I& - 2) = True
         Call ARMA_CONSULTX
         Exit Sub
       End If
     End If
   Next I&
End Sub

Private Sub Command5_Click()
   For I& = 1 To List1.ListCount
     If List1.Selected(I& - 1) = True Then
       If I& < List1.ListCount Then
         CAMPOSE$ = List1.List(I& - 1)
         List1.RemoveItem (I& - 1)
         List1.AddItem CAMPOSE$, I&
         List1.Selected(I&) = True
         Call ARMA_CONSULTX
         Exit Sub
       End If
     End If
   Next I&
   Call ARMA_CONSULTX
End Sub

Private Sub Command6_Click()
    MODIFENCA% = 0
    I& = List1.ListIndex
    If I& >= 0 Then
       TOTALI$ = "0": If Check1.Value = 1 Then TOTALI$ = "1"
       AGRUPA$ = "0": If Check2.Value = 1 Then AGRUPA$ = "1"
       MODOSOR$ = "    "
       If Option2.Value = True Then MODOSOR$ = "ASC "
       If Option3.Value = True Then MODOSOR$ = "DESC"
       CAMPOSE$ = AJUSTI$(List1.List(I&), 64) + AJUSTI$(Text1, 10) + AGRUPA$ + TOTALI$ + MODOSOR$ + AJUSTI$(Text2, 64) + AJUSTI$(Text5, 16) + AJUSTI$(Text6, 8)
       List1.RemoveItem (I&)
       List1.AddItem CAMPOSE$, I&
       List1.Selected(I&) = True
    End If
    For I& = 1 To LSTMOSTARCMP.ListCount
       If TRIM$(Left$(LSTMOSTARCMP.List(I& - 1), 64)) = TRIM$(Left$(CAMPOSE$, 64)) Then
          LSTMOSTARCMP.RemoveItem (I& - 1)
          LSTMOSTARCMP.AddItem CAMPOSE$, I& - 1
          LSTMOSTARCMP.Selected(I& - 1) = True
       End If
    Next I&
    Call ARMA_CONSULTX
End Sub

Private Sub Command8_Click()
   'ADAT$ = Combo2.TEXT
   ADAT$ = Label1.Caption
   Call CARGADEBASE(ADAT$, TEBUSCA$, OKXX%)
   
End Sub

Private Sub EJECUTAR_Click()
    SHOWREP07.Label2 = ""
    If TRIM$(Label9) <> "" Then
          SHOWREP07.Label3 = Label9        ' Descripcion de la consulta
       Else
          SHOWREP07.Label3 = "Consulta en Desarrollo"
    End If
    TEXMO$ = ARMA_CONSULTB
    Call DECOD_CONSULTA(TEXMO$)
    'If TRIM$(Combo2.TEXT) <> "" Then
    If TRIM$(Label1.Caption) <> "" Then
         SHOWREP07.Label2 = ""
         'SHOWREP07.Label2 = Combo2.TEXT
         SHOWREP07.Label2 = Label1.Caption
      Else
         SHOWREP07.Label2 = "_TEXT_"
    End If
    SHOWREP07.Show 1
    MODICONSU% = 0
End Sub

Private Sub GUARDACON_Click()
    '
    'FGUAPRO09.Text1 = Combo2.TEXT
    FGUAPRO09.Text1 = Label1.Caption
    'CODMO$ = Combo2.TEXT
    CODMO$ = Label1.Caption
    FGUAPRO09.Text2 = Label9
    FGUAPRO09.Text3 = ""
    FGUAPRO09.Label11 = ""
    FGUAPRO09.Text4 = ""
    FGUAPRO09.Label18 = ""
    FGUAPRO09.Show 1
    If FGUAPRO09.Label1 = "OK" Then
10     CODMO$ = TRIM$(FGUAPRO09.Text1)
       DESMO$ = TRIM$(FGUAPRO09.Text2)
       TEXMO$ = ARMA_CONSULTB
       '
       Call ABRETACON
       STRSQL = "SELECT * FROM SQL_PROCEDS WHERE COD_PROCEX = '" & CODMO$ & "'"
       Set CONSUSQL = BaDaConsul.OpenRecordset(STRSQL, dbOpenDynaset)
       With CONSUSQL
          If (.BOF And .EOF) Then
            .AddNew   ' YA CONVERTIDO
            ON ERROR RESUME NEXT
            !CodiEmpr = Codiemp%
            ON ERROR GOTO 0
            GoTo 15
          End If
          If COMALTER%("Ya Existe una Consulta Registrada\Bajo Ese Código.\\Cancela Guardar\Sobre-Escribir") <> 2 Then GoTo 20
          .Edit
15        !COD_PROCEX = CODMO$
          !DES_PROCEX = DESMO$
          !TIP_PROCEX = 0        ' REPORTE TABULADO
          !PAR_PROCEX = TEXMO$
          .Update
       End With
       '
20     On Error Resume Next
       CONSUSQL.Close
       On Error GoTo 0
       Set CONSUSQL = Nothing
       '
       If EXISTE%("F:\FLEXOFT\NOVEDAD\SQL_PROCEX\.") > 0 Then
           ASALI$ = "F:\FLEXOFT\NOVEDAD\SQL_PROCEX\" + CODMO$ + ".SPF"
           Call SAVEFILE(ASALI$, TEXMO$)
       End If
       '
       If FGUAPRO09.Text3 <> "" Then
          If FGUAPRO09.Label11 <> "" Then
             ASALI$ = FGUAPRO09.Label11 + "\" + FGUAPRO09.Text3 + ".SPF"
             Call SAVEFILE(ASALI$, TEXMO$)
          End If
       End If
       '
    End If
    '
    Label1.Caption = CODMO$
    '
End Sub

Private Sub Label8_DblClick()
    Unload Me
End Sub

Private Sub List1_Click()
   If MODIFENCA% = 2 Then Exit Sub
   If MODIFENCA% = 1 Then
      If COMALTER%("¿ Descarta Cambios ?\\No, Conservar\Si, Descartar") <> 2 Then
         MODIFENCA% = 2
         On Error Resume Next
         List1.ListIndex = ILIS1&
         On Error GoTo 0
         MODIFENCA% = 1
         Exit Sub
      End If
   End If
   '
   CAMPOSE$ = List1.TEXT
   Text1 = TRIM$(Mid$(CAMPOSE$, 65, 10))
   Check1.Value = 0: If Mid$(CAMPOSE$, 76, 1) = "1" Then Check1.Value = 1
   Check2.Value = 0: If Mid$(CAMPOSE$, 75, 1) = "1" Then Check2.Value = 1
   Text2 = TRIM$(Mid$(CAMPOSE$, 81, 64))
   Text5 = TRIM$(Mid$(CAMPOSE$, 145, 16))
   Text6 = TRIM$(Mid$(CAMPOSE$, 161, 8))
   '
   ORDEX$ = TRIM$(UCase$(Mid$(CAMPOSE$, 77, 4)))
   Option1.Value = True
   If ORDEX$ = "ASC" Then Option2.Value = True
   If ORDEX$ = "DESC" Then Option3.Value = True
   MODIFENCA% = 0
   ILIS1& = List1.ListIndex
End Sub

Private Sub List2_DblClick()
   TTXX$ = List2.TEXT
   I& = List2.ListIndex
   If I& >= 0 Then
      CMBCAMPOS.TEXT = TRIM$(Left$(TTXX$, 24))
      CMBOPERADOR.TEXT = TRIM$(Mid$(TTXX$, 26, 8))
        VALORX$ = TRIM$(REPLACAR$(Mid$(TTXX$, 35), "'", ""))
      Combo1.TEXT = VALORX$
      List2.RemoveItem I&
   End If
   Call ARMA_CONSULTX
End Sub

Private Sub LSTMOSTARCMP_Click()
  Call ARMASELFIELD
End Sub

Sub ARMASELFIELD()
    '
    For I& = 0 To LSTMOSTARCMP.ListCount - 1
      ESTADO = LSTMOSTARCMP.Selected(I&)
      CAMPO$ = TRIM$(LSTMOSTARCMP.List(I&))
      If ESTADO = True Then
           For J& = 1 To List1.ListCount
              If TRIM$(Left$(List1.List(J& - 1), 64)) = TRIM$(Left$(CAMPO$, 64)) Then GoTo 20
           Next J&
           List1.AddItem CAMPO$
           List1.Selected(List1.ListCount - 1) = True
         ElseIf ESTADO = False Then
10         For J& = 1 To List1.ListCount
              If TRIM$(Left$(List1.List(J& - 1), 64)) = TRIM$(Left$(CAMPO$, 64)) Then
                 List1.RemoveItem (J& - 1)
                 If J& > 1 Then List1.Selected(J& - 2) = True
                 GoTo 10
              End If
           Next J&
      End If
20  Next I&
    Call ARMA_CONSULTX
    '
End Sub
'
Private Sub LSTTABLAS_Click()
    TXTCONSULTA.TEXT = ""
    CMBCAMPOS.Clear
    CMBOPERADOR.TEXT = ""
    Combo1.Clear
    LSTMOSTARCMP.Clear
    List1.Clear
    '
    If TRIM$(LSTTABLAS.TEXT) = "" Then Exit Sub
    '
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
    Set AUX =  FLEXCONN.EXECUTE (FILTSQL$ ("select * from " & LSTTABLAS.TEXT))
    For I& = 1 To AUX.Fields.Count
        CMBCAMPOS.AddItem AUX.Fields(I& - 1).Name
        LSTMOSTARCMP.AddItem AUX.Fields(I& - 1).Name
    Next I&
    Call ARMA_CONSULTX
End Sub

Private Sub CMBBASE_CLICK()
    LSTTABLAS.Clear
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
22  SQLX$ = "Select name from sysobjects WHERE LEFT(NAME,3) <> 'dt_' and LEFT(NAME,3) <> 'pk_' and LEFT(NAME,3) <> 'sys'"
    Set AUX =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
    Do While Not AUX.EOF
       If InStr(UCase$(AUX!Name), "DTPROPERT") > 0 Then GoTo 29
       LSTTABLAS.AddItem UCase$(AUX!Name)
29  AUX.MoveNext
    Loop
End Sub
'
Private Sub Form_Load()
  Call ABRECONEXION
  KLIK% = 0
  CMBOPERADOR.Clear
  
  CMBOPERADOR.AddItem " =  "
  CMBOPERADOR.AddItem " <  "
  CMBOPERADOR.AddItem " <= "
  CMBOPERADOR.AddItem " >  "
  CMBOPERADOR.AddItem " >= "
  CMBOPERADOR.AddItem " <> "
  CMBOPERADOR.AddItem "LIKE"
  '
  Call MUESTRATABLAS

End Sub

Sub MUESTRATABLAS()
    LSTTABLAS.Clear
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
22  SQLX$ = "Select name from sysobjects WHERE LEFT(NAME,3) <> 'dt_' and LEFT(NAME,3) <> 'pk_' and LEFT(NAME,3) <> 'sys'"
    Set AUX =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
    Do While Not AUX.EOF
       If InStr(UCase$(AUX!Name), "DTPROPERT") > 0 Then GoTo 29
       LSTTABLAS.AddItem UCase$(AUX!Name)
29  AUX.MoveNext
    Loop
End Sub
'
Sub ARMA_CONSULTX()
    KLIK% = 0
    SQLX$ = "SELECT ": SQLZ$ = "": SQLT$ = ""
    VUELTA& = 0: VUELTB& = 0: VUELTC& = 0: AGRUPA% = 0
    For I& = 0 To List1.ListCount - 1
       CAMPO$ = TRIM$(Left$(List1.List(I&), 64))
       ALIASCAMPO$ = CAMPO$
       If Mid$(List1.List(I&), 75, 1) = "1" Then
          ALIASCAMPO$ = "Sx" + CAMPO$
          CAMPO$ = "SUM(" + CAMPO$ + ") AS " + ALIASCAMPO$
          AGRUPA% = 1
       End If
       If VUELTA& = 0 Then
           SQLX$ = SQLX$ + CAMPO$
          Else
           SQLX$ = SQLX$ + ", " + CAMPO$
       End If
       VUELTA& = VUELTA& + 1
       '
       ORDEX$ = UCase$(TRIM$(Mid$(List1.List(I&), 77, 4)))
       If ORDEX$ <> "" Then
          If VUELTB& > O Then SQLZ$ = SQLZ$ + ", "
          SQLZ$ = SQLZ$ + ALIASCAMPO$ + " " + ORDEX$
          VUELTB& = VUELTB& + 1
       End If
       '
       AGRUX$ = Mid$(List1.List(I&), 75, 1)
       If AGRUX$ <> "1" Then
          If VUELTC& > O Then SQLT$ = SQLT$ + ", "
          SQLT$ = SQLT$ + CAMPO$ + " "
          VUELTC& = VUELTC& + 1
       End If

    Next I&
    '
    If VUELTA = 0 Then SQLX$ = "SELECT *"
    '
    SQLY$ = " FROM " + LSTTABLAS.TEXT
    '
    If List2.ListCount > 0 Then
       JJ& = 0
       For I& = 0 To List2.ListCount - 1
          SQLY$ = SQLY$ '+ vbCrLf
          If JJ& < 1 Then SQLY$ = SQLY$ + " WHERE " Else SQLY$ = SQLY$ + " AND "
          SQLY$ = SQLY$ + TRIM$(Left$(List2.List(I&), 24)) + " " + TRIM$(Mid$(List2.List(I&), 26, 8)) + " " + TRIM$(Mid$(List2.List(I&), 35))
          JJ& = JJ& + 1
       Next I&
    End If
    '
    TXTCONSULTA.TEXT = SQLX$ + SQLY$
       If AGRUPA% = 1 Then TXTCONSULTA.TEXT = TXTCONSULTA.TEXT + " GROUP BY " + SQLT$
       If SQLZ$ <> "" Then TXTCONSULTA.TEXT = TXTCONSULTA.TEXT + " ORDER BY " + SQLZ$
    TXTCANREGI.TEXT = "SELECT COUNT(*) AS CANREGIS" + SQLY$
    '
End Sub

Function ARMA_CONSULTB()
    '
    CONSUB$ = ""
    CONSUB$ = "CODQUERY = " + TRIM$(FGUAPRO09.Text1) + Chr$(124)
    CONSUB$ = CONSUB$ + "DESQUERY = " + TRIM$(FGUAPRO09.Text2) + Chr$(124)
    CONSUB$ = CONSUB$ + "TIPQUERY = " + "   0" + Chr$(124)
    CONSUB$ = CONSUB$ + "SELTABLE = " + LSTTABLAS.TEXT + Chr$(124)
    '
    For I& = 0 To List1.ListCount - 1
       CAMPO$ = List1.List(I&)
       CONSUB$ = CONSUB$ + "SELFIELD = " + CAMPO$ + Chr$(124)
    Next I&
    '
    For I& = 0 To List2.ListCount - 1
       CONDI$ = List2.List(I&)
       CONSUB$ = CONSUB$ + "CONDITIO = " + CONDI$ + Chr$(124)
    Next I&
    '
MsgBox CONSUB
    ARMA_CONSULTB = CONSUB$
End Function

Private Sub CMDPEGAR_Click()
   COMIE = TXTCONSULTA.SelStart
   LOSEL = TXTCONSULTA.SelLength
   TXTCONSULTA.TEXT = Left$(TXTCONSULTA.TEXT, COMIE) + Clipboard.GetText + Mid$(TXTCONSULTA.TEXT, COMIE + LOSEL + 1)

End Sub

'QUEDA PENIDIENTE SI ES UN TEXTO  EL QUE VA A COMPARAR VAYA ENTRE _
COMILLAS SIMPLES (HACERLO CON ISNUMERIC ) _
SI EL OPERADOR ES LIKE DEBERIA LLEVAR EL SIMBOLO DEL PORCENTAJE _
FALTA TAMBIEN EL TEMA DEL EDITOR DE TEXTO _
VALIDAR LOS CAMPOS DE CONDICION _
CAMBIAR LAS PROPIEDADES DE LOS COMBO Y TEXT DE LA CONDICION PAARA QUE _
NO SEAN EDITABLES _
Y EJECUTAR LA CONSULTA

Private Sub Option1_Click()
   MODIFENCA% = 1
End Sub

Private Sub Option2_Click()
   MODIFENCA% = 1
End Sub

Private Sub Option3_Click()
   MODIFENCA% = 1
End Sub

Private Sub Text1_Change()
   MODIFENCA% = 1
End Sub

Private Sub Text2_Change()
   MODIFENCA% = 1
End Sub

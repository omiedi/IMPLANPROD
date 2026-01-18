VERSION 5.00
Begin VB.Form IMPOROBSOL 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importación Tabla de Obsoletos"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7110
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Tabla"
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
      Left            =   6090
      Picture         =   "IMPOROBSOL.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Editar Tabla de Obsoletos"
      Top             =   1260
      Width           =   855
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Archivo de Transferencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5265
      Left            =   210
      TabIndex        =   1
      Top             =   210
      Width           =   5685
      Begin VB.TextBox Text1 
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
         TabIndex        =   19
         Top             =   3570
         Width           =   4110
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Separador de Campos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   2940
         TabIndex        =   15
         Top             =   3990
         Width           =   2430
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Punto y Coma (:)"
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
            Left            =   315
            TabIndex        =   17
            Top             =   315
            Value           =   -1  'True
            Width           =   1800
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Coma (,)"
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
            Left            =   315
            TabIndex        =   16
            Top             =   630
            Width           =   1800
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Archivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   315
         TabIndex        =   12
         Top             =   3990
         Width           =   2430
         Begin VB.OptionButton Option2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Sobre-Escribir"
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
            Left            =   315
            TabIndex        =   14
            Top             =   630
            Width           =   1800
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Agregar al Final"
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
            Left            =   315
            TabIndex        =   13
            Top             =   315
            Value           =   -1  'True
            Width           =   1800
         End
      End
      Begin VB.CommandButton Command3 
         Height          =   960
         Left            =   4725
         Picture         =   "IMPOROBSOL.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Iniciar Lectura de Archivo"
         Top             =   2460
         Width           =   645
      End
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   2940
         Pattern         =   "*.CSV"
         TabIndex        =   7
         Top             =   420
         Width           =   2430
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   315
         TabIndex        =   3
         Top             =   840
         Width           =   2430
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   1470
         TabIndex        =   2
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia:"
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
         TabIndex        =   18
         Top             =   3600
         Width           =   1065
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   4725
         TabIndex        =   11
         Top             =   3100
         Width           =   645
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad:"
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
         Left            =   315
         TabIndex        =   10
         Top             =   470
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   2940
         TabIndex        =   8
         Top             =   2460
         Width           =   1695
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
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
         Left            =   1785
         TabIndex        =   6
         Top             =   2520
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   315
         TabIndex        =   5
         Top             =   2895
         Width           =   4320
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Carpeta:"
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
         TabIndex        =   4
         Top             =   2625
         Width           =   1065
      End
   End
   Begin VB.CommandButton Command1 
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
      Left            =   6090
      Picture         =   "IMPOROBSOL.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   5535
      Picture         =   "IMPOROBSOL.frx":059E
      Top             =   4950
      Width           =   2010
   End
End
Attribute VB_Name = "IMPOROBSOL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub



Private Sub Command2_Click()
   Command2.Enabled = False
   Call TRALOCAL("TAOBSOLE", "")
   VTB% = VENTABU%("TAOBSOLE")
   If VTB% >= 0 Then
      Call TRACENTRAL("TAOBSOLE", "")
      Call FILESORT("TAOBSOLE", "", 1, 1, "ASC")
   End If
   Command2.Enabled = True
End Sub

Private Sub Command3_Click()
    '
    Command3.Enabled = False
    Call LEEARTRA
    Command3.Enabled = True
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
End Sub


Private Sub Dir1_Change()
   Label11.Caption = Dir1.Path
   File1.Path = Dir1.Path
   Label1 = ""
End Sub

Private Sub Drive1_Change()
   On Error GoTo 20
   Dir1.Path = Drive1.Drive
   GoTo 99
   '
20 Resume 21
21 On Error GoTo 0
   Call MENSERR(24, "Unidad no Disponible.")
   Drive1.Drive = Left$(CurDir$, 2)
   '
99 On Error GoTo 0
End Sub

Private Sub File1_Click()
  Label1 = File1.filename
End Sub

Private Sub Form_Load()
    '
    EPAC$ = EMPREAC$
    If EPAC$ <> "" Then
      Caption = Caption + " - " + EMPREAC$
    End If
    '
    Label11.Caption = Dir1.Path
    File1.Path = Dir1.Path
    '
    HOII% = HOY(HOS$)
    Text1 = "Importación del " + HOS$
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
End Sub

Sub LEEARTRA()
   '
   If TRIM$(Label1) = "" Then
     Call MENSERR(24, "Falta Nombre de Archivo")
     Exit Sub
   End If
   '
   ARIMPO$ = TRIM$(Label11)
   If Right$(ARIMPO$, 1) <> "\" Then
     ARIMPO$ = ARIMPO$ + "\"
   End If
   ARIM$ = TRIM$(Label1)
   ARIMPO$ = ARIMPO$ + ARIM$
   '
   If EXISTE%(ARIMPO$) < 1 Then
     Call MENSERR(24, "Archivo Inexistente")
     Exit Sub
   End If
   '
   Screen.MousePointer = 11
   N1% = FILEOPEN%(ARIMPO$, 1, 128)
   '
   JJ& = ULTREG&("TAOBSOLE")              ' AGREGAR
   If Option2.Value = True Then JJ& = 0   ' SOBREESCRIBIR
   '
   CARASEP% = Asc(";")
   If Option3.Value = True Then CARASEP% = Asc(",")
   '
   REFE$ = TRIM$(Text1)
   '
   While Not EOF(N1%)
     Line Input #N1%, AAX$
     PPXX% = InStr(AAX$, Chr$(CARASEP%))
     If PPXX% > 0 Then
       BBX$ = Mid$(AAX$, PPXX% + 1)
       CODD$ = TRIM$(Left$(AAX$, PPXX% - 1))
       CIXI% = CODINT%(CODD$)
       If CIXI% > 0 Then
         PPXX% = InStr(BBX$, Chr$(CARASEP%))
         If PPXX% > 0 Then
           CAXS$ = TRIM$(Mid$(BBX$, PPXX% + 1))
           If Option4.Value = True Then Call REPLACAR(CAXS$, " ", "")
           CANN# = Val(CAXS$)
           If CANN# >= 0.05 Then
             BB$ = MKI$(CIXI%) + MKD$(CANN#) + REFE$
             JJ& = JJ& + 1
             Call GRAREG("TAOBSOLE", BB$, JJ&)
           End If
         End If
       End If
     End If
   Wend
   '
   Call SETULTREG("TAOBSOLE", JJ&)
   Call CIERRARCH("TAOBSOLE")
   Close #N1%
   '
   Call FILESORT("TAOBSOLE", "", 1, 1, "ASC")
   Call TRALOCAL("TAOBSOLE", "")
   VTB% = VENTABU%("TAOBSOLE")
   If VTB% >= 0 Then
      Call TRACENTRAL("TAOBSOLE", "")
      Call FILESORT("TAOBSOLE", "", 1, 1, "ASC")
   End If
  '
End Sub


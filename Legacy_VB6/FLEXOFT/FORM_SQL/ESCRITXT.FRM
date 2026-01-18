VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ESCRITXT 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Genera Archivos de Mail"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   14760
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command5 
      Caption         =   "Restaurar"
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
      Left            =   12360
      TabIndex        =   30
      Top             =   3600
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Generar Archivo"
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
      Left            =   12600
      TabIndex        =   28
      Top             =   8880
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Revertir Selección"
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
      Left            =   240
      TabIndex        =   27
      Top             =   8880
      Width           =   1815
   End
   Begin VB.ListBox List4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   10680
      Sorted          =   -1  'True
      TabIndex        =   26
      ToolTipText     =   "Doble Click para Seleccionar Formulario yMostrar Menu"
      Top             =   1680
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Frame Frame3 
      Height          =   1850
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   5775
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   915
         TabIndex        =   22
         Top             =   120
         Width           =   2340
      End
      Begin VB.DirListBox Dir1 
         Appearance      =   0  'Flat
         Height          =   1215
         Left            =   90
         TabIndex        =   21
         Top             =   480
         Width           =   3165
      End
      Begin VB.Frame Frame2 
         Caption         =   "Filtros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   3300
         TabIndex        =   14
         Top             =   150
         Width           =   2400
         Begin VB.OptionButton Option1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "EXE"
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
            Left            =   90
            TabIndex        =   20
            Top             =   315
            Width           =   645
         End
         Begin VB.OptionButton Option2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "PRF"
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
            Left            =   795
            TabIndex        =   19
            Top             =   315
            Width           =   645
         End
         Begin VB.OptionButton Option3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "RFS"
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
            Left            =   1500
            TabIndex        =   18
            Top             =   315
            Width           =   645
         End
         Begin VB.OptionButton Option4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "FRM"
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
            Left            =   765
            TabIndex        =   17
            Top             =   600
            Width           =   645
         End
         Begin VB.OptionButton Option5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "VBP"
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
            Left            =   90
            TabIndex        =   16
            Top             =   600
            Value           =   -1  'True
            Width           =   645
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
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
            Left            =   1440
            TabIndex        =   15
            Top             =   570
            Width           =   825
         End
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
         Left            =   -240
         TabIndex        =   23
         Top             =   165
         Width           =   1065
      End
   End
   Begin VB.CommandButton Command6 
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
      Height          =   550
      Left            =   13920
      Picture         =   "ESCRITXT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2280
      Visible         =   0   'False
      Width           =   700
   End
   Begin VB.TextBox Text2 
      Height          =   435
      Left            =   420
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   6405
      Visible         =   0   'False
      Width           =   3165
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Archivo Seleccionado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   1700
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   14505
      Begin VB.CommandButton Command2 
         Caption         =   "Generar Archivo Separado por punto y coma"
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
         Left            =   9360
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   960
         Width           =   4185
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "Ver"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   13400
         TabIndex        =   25
         Top             =   400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1155
         Left            =   90
         TabIndex        =   6
         Top             =   400
         Width           =   9165
      End
      Begin VB.Label Label5 
         Caption         =   "Formularios de Proyecto Activo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   9360
         TabIndex        =   31
         Top             =   170
         Visible         =   0   'False
         Width           =   3855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Selección de Archivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   8625
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         IntegralHeight  =   0   'False
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "Doble Click para Seleccionar Proyectdo"
         Top             =   240
         Width           =   8325
      End
      Begin VB.ListBox List2 
         Height          =   2205
         Left            =   5670
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   420
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.FileListBox File1 
         Height          =   2040
         Left            =   3780
         TabIndex        =   3
         Top             =   420
         Visible         =   0   'False
         Width           =   1695
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
         Left            =   4305
         TabIndex        =   4
         Top             =   1995
         Visible         =   0   'False
         Width           =   4320
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
         Height          =   705
         Left            =   3600
         TabIndex        =   2
         Top             =   2160
         Width           =   5625
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
         Left            =   3255
         TabIndex        =   1
         Top             =   1920
         Width           =   1065
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   4845
      Left            =   240
      TabIndex        =   24
      Top             =   3840
      Width           =   14475
      _ExtentX        =   25532
      _ExtentY        =   8546
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
   Begin VB.Label Label4 
      Caption         =   "Pintar con un Click los Menú que NO van a estar Visibles"
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
      Left            =   240
      TabIndex        =   29
      Top             =   3600
      Width           =   6135
   End
   Begin VB.Label Label3 
      Height          =   435
      Left            =   4200
      TabIndex        =   12
      Top             =   6405
      Width           =   3690
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu mnuAyu1 
         Caption         =   "Como debe ser el archivo original"
      End
   End
End
Attribute VB_Name = "ESCRITXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Public NAMEFRM$


Private Sub Command1_Click()
    List4.Clear

    For U& = 1 To List1.ListCount
      Z$ = List1.LIST(U& - 1)
      NARCHI1% = FILEOPEN%(Z$, 1, 128)  'ABRE EL ARCHIVO DE LECTURA
      Do While Not EOF(NARCHI1%) ' RECORRE
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
       If Left$(XX$, 7) = "Form=.." Then
       'Print #NARCHI2%, XX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
          List4.AddItem XX$
       End If
     Loop
    Next U&
      
    Close #NARCHI1%
    
'    Call DuplicadosEnList(List4)
    Call DuplicadosList1(List4)
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   If EXISTE%(List1.LIST(0)) < 1 Then Call COMUNI("Archivo Original Inexistente"): GoTo 99
   
   
   POS1% = InStrRev(List1.LIST(0), "\")
   Ruta$ = Left(List1.LIST(0), POS1%)
   If EXISTE%(Ruta$ & "MAILDEP2.txt") > 0 Then
       On Error Resume Next
       Kill (Ruta$ & "MAILDEP2.txt")
       On Error GoTo 0
   End If
   NARCHI2% = FILEOPEN%(Ruta$ & "MAILDEP2.txt", 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
   NARCHI1% = FILEOPEN%(List1.LIST(0), 1, 128) 'ABRE EL ARCHIVO DE LECTURA
   J& = 0
   Do While Not EOF(NARCHI1%) ' RECORRE
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
       XX1$ = TRIM$(UCase$(XX$))
       If XX1$ <> "" Then
         TEXTOX$ = TEXTOX$ & XX1$ & " ; "
         J& = J& + 1
       End If
       
       If J& > 9 Then
          TEXTOX$ = TEXTOX$ + vbCrLf + vbCrLf
          Print #NARCHI2%, TEXTOX$
          TEXTOX$ = ""
          J& = 0
       End If
    Loop
    If J& > 0 Then
       Print #NARCHI2%, TEXTOX$
    End If
    
    Close #NARCHI1%

    Close #NARCHI2%
    
'    On Error Resume Next
'    Kill (List1.LIST(0))
'    On Error GoTo 0
    '
    'COPIO EL ARCHIVO ESCRITO SOBRE EL ORIGINAL
'    XY1% = COPYFILE(Ruta$ & "MAILDEP2.txt", List1.LIST(0))
'    On Error Resume Next
'    Kill (Ruta$ & "MAILDEP2.txt")
'    On Error GoTo 0
    MsgBox "Archivo Generado en: " & Ruta$ & "MAILDEP2.txt"
    
   LINEA$ = Ruta$ & "MAILDEP2.txt"
   Archi$ = TRIM$(LINEA$)
   If EXISTE%(Archi$) > 0 Then
        Call ShellExecute(Me.hwnd, "Open", Archi$, "", "", 1)
    End If

99  Command2.Enabled = True
   
End Sub

Private Sub Command3_Click()
   '
   Command3.Enabled = False
   Screen.MousePointer = 11
   FIN& = GRID.Rows - 1
   GRID.Redraw = False
   For i& = 1 To FIN&
       Call TRACE(20, i&, FIN&)
       Call COLOREAR_GRILLA_SOLAFILA(Me.GRID, vbYellow, i&)
   Next i&
   GRID.Redraw = True
   Screen.MousePointer = 1

   Command3.Enabled = True
   '
End Sub

Private Sub Command4_Click()
   '
   TEXTOAESCRIBIR$ = ""
   If TRIM$(NAMEFRM$) = "" Then Exit Sub
   NAMEFRM$ = UCase$(NAMEFRM$)
   For i& = 1 To GRID.Rows - 1
      If GRID.TextMatrix(i&, 0) = "*" Then
         TEXTOAESCRIBIR$ = TEXTOAESCRIBIR$ + GRID.TextMatrix(i&, 2) + vbCrLf
      End If
   Next i&
   Call EscribeCONFIGMENU(NAMEFRM$, TEXTOAESCRIBIR$)
   
End Sub

Sub EscribeCONFIGMENU(TITUFORM$, TEXTOAESCRIBIR$)

   CARPETOK$ = LUDAT$
   
   NARCHI2% = FILEOPEN%(CARPETOK$ + "CONFIGMENU2.TXT", 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
   If EXISTE%(CARPETOK$ + "CONFIGMENU.TXT") > 0 Then
      NARCHI1% = FILEOPEN%(CARPETOK$ + "CONFIGMENU.TXT", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
   Else
      GoTo 70
   End If
   Do While Not EOF(NARCHI1%) ' RECORRE
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
       XX1$ = TRIM$(UCase$(XX$))
       If XX1$ = "|INICIO " & TITUFORM$ Then NOESCRIBIR% = 1
        If NOESCRIBIR% < 1 Then
            Print #NARCHI2%, XX$
       End If
       If XX1$ = "|FINAL " & TITUFORM$ Then NOESCRIBIR% = 0
    Loop
    Close #NARCHI1%
    
70  Print #NARCHI2%, "|INICIO " & TITUFORM$
    Print #NARCHI2%, TEXTOAESCRIBIR$
    Print #NARCHI2%, "|FINAL " & TITUFORM$
    Close #NARCHI2%
    
    On Error Resume Next
    Kill (CARPETOK$ + "CONFIGMENU.TXT")
    On Error GoTo 0
    '
    'COPIO EL ARCHIVO ESCRITO SOBRE EL ORIGINAL
    XY1% = COPYFILE(CARPETOK$ + "CONFIGMENU2.TXT", CARPETOK$ + "CONFIGMENU.TXT")
    On Error Resume Next
    Kill (CARPETOK$ + "CONFIGMENU2.TXT")
    On Error GoTo 0
99  Call COMUNI("Actualizado")
   
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   If EXISTE%(LUDAT$ + "CONFIGMENU.TXT") < 1 Then GoTo 99
   '
   TITUFORM$ = UCase$(NAMEFRM$)
   J& = 0
   Dim NAMERUTTXT$()
   NARCHI1% = FILEOPEN%(LUDAT$ + "CONFIGMENU.TXT", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
   Do While Not EOF(NARCHI1%) ' RECORRE
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
       XX1$ = TRIM$(UCase$(XX$))
       If XX1$ = "|INICIO " & TITUFORM$ Then LEER1% = 1
       If XX1$ = "|FINAL " & TITUFORM$ Then LEER1% = 0
       If LEER1% > 0 Then
          J& = J& + 1
          ReDim Preserve NAMERUTTXT$(J&)
          NAMERUTTXT$(J&) = XX1$
       ElseIf J& > 0 Then 'SI J ES > 0 POR QUE YA ESTUVO ESCRIBIENDO EL VECTOR Y SI LLEGA A FINAL SALE
         Exit Do
       End If
    Loop
    Close #NARCHI1%
   If J& < 1 Then GoTo 99
   For i& = 1 To GRID.Rows - 1
       NAMERUT$ = TRIM$(UCase$(GRID.TextMatrix(i&, 2)))
       For L& = 1 To UBound(NAMERUTTXT$)
          If NAMERUT$ = NAMERUTTXT$(L&) Then Call COLOREAR_GRILLA_SOLAFILA(GRID, vbYellow, i&, 0, 0, 0)
       Next L&
   Next i&
   
   
99 Command5.Enabled = True
End Sub

Private Sub Command6_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Dir1_Change()
   Label11.Caption = Dir1.Path
   File1.Path = Dir1.Path
   Label1 = ""
   Call Text1_Change
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

Private Sub GRID_Click()
        Call COLOREAR_GRILLA_SOLAFILA(GRID, vbYellow, GRID.Row, 0, 0, 0)
End Sub

Private Sub List3_DblClick()
  If TRIM$(Left$(List3.TEXT, 32)) <> Label1.Caption Then
     GRID.Rows = 1
     List4.Clear
     List1.Clear
  End If
  Label1 = TRIM$(Left$(List3.TEXT, 32))

  Call CONECTADOC
  Call Command1_Click 'muestro los formularios del proyecto seleccionad
   
End Sub

Private Sub Form_Load()
    '
    Screen.MousePointer = 11
'    Call CENTRAFORM(Me)
'    Show
    Refresh
    '
'    Call VERJOBID(OKEY%)
'    If OKEY% < 1 Then Call FINAL("")
'
'    EPAC$ = EMPREAC$
'    If EPAC$ <> "" Then
'      Caption = Caption + " - " + EMPREAC$
'    End If
    '
    On Error Resume Next
    Drive1.Drive = "C:"
    If Err Then
      On Error GoTo 0
      Drive1.Drive = Left$(App.Path, 2)
    End If
    On Error GoTo 0
    'Dir1.Path = "\FLEXOFT\NOVEDAD"
    'Dir1.Path = "\Program Files (x86)\Microsoft Visual Studio\VB98\FLEXOFT\PROY_SQL"
    Label11.Caption = Dir1.Path
    File1.Path = Dir1.Path
    Text1 = "*.TXT"
    Screen.MousePointer = 1
    Me.Visible = True
    '
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   Call CIERRARCH("*.*")
'           Call FINAL("")
End Sub

Private Sub Form_Resize()
'   If WindowState = 1 Then
'     Call CIERRARCH("*.*")
'     Call BAJALDISCO
'   End If
End Sub

Private Sub Label11_Change()
  If Right$(Label11, 1) = "\" Then
    Label11 = Left$(Label11, Len(Label11) - 1)
  End If
End Sub

Private Sub List1_DblClick()
   '
   ILISTA& = List1.ListIndex
   List1.RemoveItem ILISTA&
   '
End Sub

Sub CONECTADOC()
   '
   
  If TRIM$(Label1) <> "" Then
     XX$ = TRIM$(Label11) + "\" + TRIM$(Label1)
     For U& = 1 To List1.ListCount
       If List1.LIST(U& - 1) = XX$ Then Exit Sub
     Next U&
     List1.AddItem XX$
     List1.Refresh
     DoEvents
   End If
   '
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
   If Button = 2 Then
     If Y > 0 Then
       YACT% = Int(5 * Y / List1.Height)
       On Error Resume Next
       List1.ListIndex = List1.TopIndex + YACT%
       TTXX$ = List1.TEXT + Chr$(13) + Chr$(10)
       TTXX$ = TTXX$ + "Actualizado: " + Format$(FileDateTime(List1.TEXT), "dd/mm/yyyy - hh:mm") + " hs" + Chr$(13) + Chr$(10)
       TTXX$ = TTXX$ + "Tamaño: " + TRIM$(Str$(FileLen(List1.TEXT))) + " bytes."
       If Err = 0 Then
         MsgBox TTXX$
       End If
       On Error GoTo 0
     End If
   End If
End Sub

Private Sub List4_Click()
    ILISTA& = List4.ListIndex
    RUTAFRM$ = Mid$(List4.LIST(ILISTA&), 8)
    POS1% = InStrRev(Dir1.Path, "\")
    Ruta$ = Left(Dir1.Path, POS1% - 1)
    RUTAF$ = Ruta$ + RUTAFRM$
    POS2% = InStrRev(RUTAFRM$, "\")
    NAMEFRM$ = Mid$(RUTAFRM$, POS2% + 1)
    If Right(UCase(NAMEFRM$), 4) = ".FRM" Then
      NAMEFRM$ = Mid$(NAMEFRM$, 1, Len(NAMEFRM$) - 4)
    End If

    Screen.MousePointer = 11
    '
    Call CARGA_ENCABEZADO(GRID, "Caption/A48;Nombre Rutina/A48", Me)
    NARCHI1% = FILEOPEN%(RUTAF$, 1, 128)  'ABRE EL ARCHIVO DE LECTURA
    Do While Not EOF(NARCHI1%) ' RECORRE
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
       If Left(XX$, 16) = "   Begin VB.Menu" Then
            J& = J& + 1
            GRID.Rows = J& + 1
            LEIDO$ = UCase$(TRIM$(XX$))
            POS1% = InStr(LEIDO$, "=")
            NAMERUTINA$ = TRIM$(Mid$(LEIDO$, 14))
            GRID.TextMatrix(J&, 2) = NAMERUTINA$
       End If

       If Left(XX$, 19) = "      Begin VB.Menu" Then
            J& = J& + 1
            GRID.Rows = J& + 1
            LEIDO$ = UCase$(TRIM$(XX$))
            POS1% = InStr(LEIDO$, "=")
            NAMERUTINA$ = TRIM$(Mid$(LEIDO$, 14))
            GRID.TextMatrix(J&, 2) = NAMERUTINA$
       End If
       If Left(XX$, 22) = "         Begin VB.Menu" Then
            J& = J& + 1
            GRID.Rows = J& + 1
            LEIDO$ = UCase$(TRIM$(XX$))
            POS1% = InStr(LEIDO$, "=")
            NAMERUTINA$ = TRIM$(Mid$(LEIDO$, 14))
            GRID.TextMatrix(J&, 2) = NAMERUTINA$
       End If
       If Left(XX$, 25) = "            Begin VB.Menu" Then
            J& = J& + 1
            GRID.Rows = J& + 1
            LEIDO$ = UCase$(TRIM$(XX$))
            POS1% = InStr(LEIDO$, "=")
            NAMERUTINA$ = TRIM$(Mid$(LEIDO$, 14))
            GRID.TextMatrix(J&, 2) = NAMERUTINA$
       End If
       '
       If J& < 1 Then GoTo 80 'SI NO ENTRO EN NINGUN Begin VB.Menu no es menu
       If Left(XX$, 13) = "      Caption" Then
            LEIDO$ = UCase$(TRIM$(XX$))
            POS1% = InStr(LEIDO$, "=")
            CAPTL$ = TRIM$(Mid$(LEIDO$, POS1% + 1))
            CAPTL$ = "---" & REPLACAR(CAPTL$, Chr(34), "")
            GRID.TextMatrix(J&, 1) = CAPTL$
            If CAPTL$ = "-" Then J& = J& - 1
       End If

       If Left(XX$, 16) = "         Caption" Then
            LEIDO$ = UCase$(TRIM$(XX$))
            POS1% = InStr(LEIDO$, "=")
            CAPTL$ = TRIM$(Mid$(LEIDO$, POS1% + 1))
            CAPTL$ = "------" & REPLACAR(CAPTL$, Chr(34), "")
            GRID.TextMatrix(J&, 1) = CAPTL$
            If CAPTL$ = "-" Then J& = J& - 1
       End If
      
       
       If Left(XX$, 19) = "            Caption" Then
            LEIDO$ = UCase$(TRIM$(XX$))
            POS1% = InStr(LEIDO$, "=")
            CAPTL$ = TRIM$(Mid$(LEIDO$, POS1% + 1))
            CAPTL$ = "---------" & REPLACAR(CAPTL$, Chr(34), "")
            GRID.TextMatrix(J&, 1) = CAPTL$
            If CAPTL$ = "-" Then J& = J& - 1 ' POR SI ES SEPARADOR
       End If
       If Left(XX$, 22) = "               Caption" Then
            LEIDO$ = UCase$(TRIM$(XX$))
            POS1% = InStr(LEIDO$, "=")
            CAPTL$ = TRIM$(Mid$(LEIDO$, POS1% + 1))
            CAPTL$ = "------------" & REPLACAR(CAPTL$, Chr(34), "")
            GRID.TextMatrix(J&, 1) = CAPTL$
            If CAPTL$ = "-" Then J& = J& - 1
       End If
       If Left(XX$, 25) = "                  Caption" Then
            LEIDO$ = UCase$(TRIM$(XX$))
            POS1% = InStr(LEIDO$, "=")
            CAPTL$ = TRIM$(Mid$(LEIDO$, POS1% + 1))
            CAPTL$ = "---------------" & REPLACAR(CAPTL$, Chr(34), "")
            GRID.TextMatrix(J&, 1) = CAPTL$
            If CAPTL$ = "-" Then J& = J& - 1
       End If
       


        If UCase$(Left$(TRIM$(XX$), 3)) = "END" Then
           If ULTIMOEND$ = "End" Then
             If GRID.Rows > 1 Then
                J& = J& + 1
                GRID.Rows = J& + 1
                Call COLOREAR_GRILLA_SOLO_UNA_AFILA(Me.GRID, &H808080, J&)
                GRID.TextMatrix(J&, 1) = "-------------"
                GRID.TextMatrix(J&, 2) = "-------------"
             End If
           End If
        End If
       ULTIMOEND$ = TRIM$(XX$)
       If XX$ = "END" Then
          Exit Do
       End If

80   Loop
    '
    Close #NARCHI1%
    Call Command5_Click ' PINTA LOS QUE ESTAN EN EL ARCHIVO
    Screen.MousePointer = 1
End Sub

Private Sub mnuAyu1_Click()
   TX1$ = "EL ARCHIVO ORIGINAL DEBE TENER LA LISTA DE MAIL ENCOLUMNADA DE A UNO" + vbCrLf
   TX1$ = TX1$ + "EJEMPLO:" + vbCrLf + vbCrLf
   TX1$ = TX1$ + "info@sysflom.com.ar" + vbCrLf
   TX1$ = TX1$ + "omar.demaio@sysflom.com.ar" + vbCrLf
   TX1$ = TX1$ + "flavio.grasso@sysflom.com.ar" + vbCrLf
   TX1$ = TX1$ + "...." + vbCrLf
   TX1$ = TX1$ + "...." + vbCrLf
   TX1$ = TX1$ + "...." + vbCrLf
   TX1$ = TX1$ + "-----------------------------------------" + vbCrLf + vbCrLf + vbCrLf
   TX1$ = TX1$ + "Paso 1: Se debera elegir el archivo original" + vbCrLf
   TX1$ = TX1$ + "Quedando este en list inferior" + vbCrLf
   TX1$ = TX1$ + "Paso 2: Click en boton: " + vbCrLf
   TX1$ = TX1$ + "Generar Archivo Separado por punto y coma"
   
   MsgBox TX1$, vbInformation
   
End Sub


Private Sub Option1_Click()
   Text1 = "*.EXE"
End Sub

Private Sub Option2_Click()
   Text1 = "*.PRF"
End Sub

Private Sub Option3_Click()
   Text1 = "*.RFS"
End Sub

Private Sub Option4_Click()
   Text1 = "*.FRM"
End Sub

Private Sub Option5_Click()
   Text1 = "*.VBP"
End Sub

Private Sub Text1_Change()
   File1.Pattern = Text1
   List2.Clear
   For KU& = 1 To File1.ListCount
     FIDAT$ = Format$(CDbl(FileDateTime(Label11 + "\" + File1.LIST(KU& - 1))), "00000#.#####   ")
     While Len(FIDAT$) < 15: FIDAT$ = FIDAT$ + " ": Wend
     List2.AddItem FIDAT$ + File1.LIST(KU& - 1)
   Next KU&
   '
   List3.Clear
   For KU& = List2.ListCount To 1 Step -1
     YY$ = List2.LIST(KU& - 1)
     IZQUI$ = REPLACAR$(Left$(YY$, 14), ",", ".")
     XX$ = Left$(Mid$(YY$, 16, 32) + Space$(32), 32) + Format$(CDate(Val(IZQUI$)), "dd/mm/yy  hh:mm:ss")
     List3.AddItem XX$
   Next KU&
End Sub


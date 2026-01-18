VERSION 5.00
Begin VB.Form APARFUEN2 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Apareo de fuentes"
   ClientHeight    =   10935
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   15315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   15315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command14 
      Caption         =   "Command14"
      Height          =   975
      Left            =   10440
      TabIndex        =   38
      Top             =   360
      Width           =   495
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   9120
      TabIndex        =   37
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Apli- ca SKIN"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11130
      TabIndex        =   36
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Con- ver- tir"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11880
      TabIndex        =   35
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Clear"
      Enabled         =   0   'False
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
      TabIndex        =   34
      Top             =   960
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Comment and Correct"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   12600
      TabIndex        =   33
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Busquedas"
      Enabled         =   0   'False
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
      Left            =   13680
      TabIndex        =   32
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Poner FILTSQL"
      Enabled         =   0   'False
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
      Left            =   13680
      TabIndex        =   31
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Quitar FILTSQL"
      Enabled         =   0   'False
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
      Left            =   13680
      TabIndex        =   30
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   8925
      TabIndex        =   29
      ToolTipText     =   "Comparacion Integral"
      Top             =   2520
      Width           =   350
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Incremental"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   13755
      TabIndex        =   28
      Top             =   3430
      Width           =   1485
   End
   Begin VB.CommandButton Command6 
      Caption         =   "<<<<<<   Actualiza LOCAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   7770
      TabIndex        =   27
      Top             =   3620
      Width           =   2745
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Actualiza REMOTO   >>>>>>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   4725
      TabIndex        =   26
      Top             =   3620
      Width           =   2850
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "APARFUEN2.frx":0000
      Top             =   9360
      Width           =   7500
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   7730
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "APARFUEN2.frx":0006
      Top             =   9360
      Width           =   7500
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Proxima Diferencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   11800
      TabIndex        =   17
      Top             =   3430
      Width           =   1940
   End
   Begin VB.ListBox List6 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   11805
      TabIndex        =   13
      Top             =   1785
      Width           =   3420
   End
   Begin VB.ListBox List5 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4920
      Left            =   7720
      MultiSelect     =   2  'Extended
      TabIndex        =   12
      Top             =   4305
      Width           =   7500
   End
   Begin VB.ListBox List4 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4920
      Left            =   105
      MultiSelect     =   2  'Extended
      TabIndex        =   11
      Top             =   4305
      Width           =   7500
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   9280
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1785
      Width           =   2400
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   6510
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1785
      Width           =   2400
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   2040
      TabIndex        =   5
      Top             =   300
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Height          =   1080
      Left            =   8280
      Picture         =   "APARFUEN2.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Inicia Exploraciónde la Red"
      Top             =   300
      Width           =   690
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   105
      TabIndex        =   2
      Top             =   1785
      Width           =   6300
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   5760
      TabIndex        =   1
      Top             =   300
      Width           =   2430
   End
   Begin VB.DirListBox Dir1 
      Height          =   540
      Left            =   120
      TabIndex        =   0
      Top             =   820
      Width           =   5430
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   105
      TabIndex        =   25
      Top             =   10320
      Width           =   7470
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   7725
      TabIndex        =   24
      Top             =   10320
      Width           =   7470
   End
   Begin VB.Label Label112 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remoto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   7770
      TabIndex        =   23
      Top             =   3980
      Width           =   990
   End
   Begin VB.Label Label111 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Local:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Left            =   105
      TabIndex        =   22
      ToolTipText     =   "Doble Click Para Editar"
      Top             =   3980
      Width           =   690
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   330
      Left            =   1680
      TabIndex        =   21
      Top             =   3980
      Width           =   5055
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   330
      Left            =   8925
      TabIndex        =   20
      Top             =   3980
      Width           =   5055
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C000&
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   6825
      TabIndex        =   16
      Top             =   3980
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   14490
      TabIndex        =   15
      Top             =   3980
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Rutinas y Funciones:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11800
      TabIndex        =   14
      Top             =   1515
      Width           =   2880
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Equipo Remoto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9280
      TabIndex        =   10
      Top             =   1515
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Equipo Local:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6510
      TabIndex        =   9
      Top             =   1515
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Carpetas a Explorar:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   105
      TabIndex        =   8
      Top             =   1515
      Width           =   3930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Equipo Remoto:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   -270
      TabIndex        =   4
      Top             =   380
      Width           =   1935
   End
End
Attribute VB_Name = "APARFUEN2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RENGLOX$(32767), CANREN&
Dim ESTARUT$, OTRARUT$
Dim LLAMASKIN% ' VARIABLES DE PROCESO SKIN
Dim CREAPIC%
Dim INCRUSTASKIN%
Dim REFERCOMPON%
Dim ADESTI$


Private Sub Command1_Click()
   '
   CARCHI% = 0
   '
   OTROEQ$ = TRIM$(Combo1.TEXT)
   If OTROEQ$ = "" Then
      MsgBox "Identifique Equipo Remoto !!!"
      Exit Sub
   End If
   '
   ESTEQ$ = "C:"
   DIRFSX$ = OTROEQ$ + "\ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB\FLEXOFT"
   If OTROEQ$ = "\\Notebook-bauer\C" Then
      DIRFSX$ = OTROEQ$ + "\Program Files\DEVSTUDIO\VB\FLEXOFT"
   ElseIf OTROEQ$ = "z:\" Then
      DIRFSX$ = OTROEQ$ + "FLEXOFT"
   End If
   DIEXPLO$ = DIRFSX$ + "\"
   '
   Screen.MousePointer = 11
   '
   ESTARUT$ = "D:\Archivos de Programa\Microsoft Visual Studio\DevStudio\VB\FLEXOFT\"
   'OTRARUT$ = OTROEQ$ + "\ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB\FLEXOFT\"
   OTRARUT$ = OTROEQ$ + "\Microsoft Visual Studio\VB98\FLEXOFT\"
   If OTROEQ$ = "\\Notebook-bauer\C" Then
      OTRARUT$ = OTROEQ$ + "\Program Files\DEVSTUDIO\VB\FLEXOFT\"
   ElseIf OTROEQ$ = "z:\" Then
      OTRARUT$ = OTROEQ$ + "FLEXOFT\"
   End If
   '
   ' EXPLORA LOS DIRECTORIOS LOCALES
   List1.Clear: List2.Clear
   List1.AddItem ESTARUT$
   ILIST1& = 0
   '
   Do While ILIST1& < List1.ListCount
      RUTAX$ = List1.List(ILIST1&)
      On Error Resume Next
      NOMBREX$ = Dir(RUTAX$, 16) ' Recupera la primera entrada.
      If Err Then
         On Error GoTo 0
         GoTo 6
      End If
      On Error GoTo 0
      Do While NOMBREX$ <> "" ' Inicia el bucle.
        ' Ignora el directorio actual y el que lo abarca.
        If NOMBREX$ <> "." And NOMBREX$ <> ".." Then
           ' Utiliza comparación a nivel de bits para asegurarse de que MiNombre es un directorio.
           If (GetAttr(RUTAX$ & NOMBREX$) And vbDirectory) = vbDirectory Then
              If UCase$(RUTAX$ + NOMBREX$) <> ULODAY$ Then ' EXCLUYE DIRECTORIO TRABAJO
                 List1.AddItem RUTAX$ + NOMBREX$ + "\"
                 List1.ListIndex = List1.ListCount - 1
              End If
           End If ' solamente si representa un directorio.
        End If
        NOMBREX$ = Dir ' Obtiene siguiente entrada.
        Openforms = DoEvents
      Loop
6     ILIST1& = ILIST1& + 1
   Loop
   '
   ' CARGA EN LIST2 LOS ARCHIVOS LOCALES
   For i& = 1 To List1.ListCount
      'Dir1.Path = List1.List(I& - 1)
      File1.Path = List1.List(i& - 1)
      File1.Refresh
      estaruta$ = File1.Path + "\"
      For j& = 1 To File1.ListCount
         FINAME$ = TRIM$(UCase$(File1.List(j& - 1)))
           If InStr(FINAME$, ".FRX") > 0 Then GoTo 17
           If InStr(FINAME$, ".VBW") > 0 Then GoTo 17
           If InStr(FINAME$, ".PDF") > 0 Then GoTo 17
           If InStr(FINAME$, ".PRF") > 0 Then GoTo 17
           If InStr(FINAME$, ".LOG") > 0 Then GoTo 17
           
         TTXX$ = AJUSTI$(FINAME$, 64)
         TTXX$ = TTXX$ + estaruta$ + FINAME$
         List2.AddItem TTXX$
         Openforms = DoEvents
17    Next j&
19 Next i&
   '
   ' EXPLORA DIRECTORIO DEL OTRO EQUIPO
   List1.Clear: List3.Clear
   List1.AddItem OTRARUT$
   ILIST1& = 0
   '
   Do While ILIST1& < List1.ListCount
      RUTAX$ = List1.List(ILIST1&)
      On Error Resume Next
      NOMBREX$ = Dir(RUTAX$, 16) ' Recupera la primera entrada.
      If Err Then
         On Error GoTo 0
         GoTo 6
      End If
      On Error GoTo 0
      Do While NOMBREX$ <> "" ' Inicia el bucle.
        ' Ignora el directorio actual y el que lo abarca.
        If NOMBREX$ <> "." And NOMBREX$ <> ".." Then
           ' Utiliza comparación a nivel de bits para asegurarse de que MiNombre es un directorio.
           If (GetAttr(RUTAX$ & NOMBREX$) And vbDirectory) = vbDirectory Then
              If UCase$(RUTAX$ + NOMBREX$) <> ULODAY$ Then ' EXCLUYE DIRECTORIO TRABAJO
                 List1.AddItem RUTAX$ + NOMBREX$ + "\"
                 List1.ListIndex = List1.ListCount - 1
              End If
           End If ' solamente si representa un directorio.
        End If
        NOMBREX$ = Dir ' Obtiene siguiente entrada.
        Openforms = DoEvents
      Loop
      ILIST1& = ILIST1& + 1
   Loop
   '
   ' CARGA EN LIST3 LOS ARCHIVOS REMOTOS
   For i& = 1 To List1.ListCount
      Dir1.Path = List1.List(i& - 1)
      Dir1.Refresh
      File1.Path = Dir1.Path
      File1.Refresh
      otraruta$ = File1.Path + "\"
      'File1.Refresh
      For j& = 1 To File1.ListCount
         FINAME$ = TRIM$(UCase$(File1.List(j& - 1)))
         If FINAME$ = "BADASTRU.MDB" Or InStr(FINAME$, ".LOG") > 0 Or InStr(FINAME$, ".PRF") > 0 Then
            ACOMPLE$ = OTRARUT$ + FINAME$
            On Error Resume Next
            Kill ACOMPLE$
            On Error GoTo 0
            GoTo 27
         End If
            '
            If InStr(FINAME$, ".FRX") > 0 Then GoTo 27
            If InStr(FINAME$, ".VBW") > 0 Then GoTo 27
            If InStr(FINAME$, ".PDF") > 0 Then GoTo 27
            If InStr(FINAME$, ".PRF") > 0 Then GoTo 27
            If InStr(FINAME$, ".LOG") > 0 Then GoTo 27
         TTXX$ = AJUSTI$(FINAME$, 64)
         TTXX$ = TTXX$ + otraruta$ + FINAME$
         List3.AddItem TTXX$
         Openforms = DoEvents
27    Next j&
29 Next i&
   '
   ' DEPURA LOS QUE COINCIDEN
   For ILISTA& = 1 To List2.ListCount
32    ESTEAR$ = TRIM$(Left$(List2.List(ILISTA& - 1), 64))
      estearco$ = TRIM$(Mid$(List2.List(ILISTA& - 1), 65))
      PPXX% = InStr(UCase$(estearco$), "ARCHIVOS DE PRO")
      If PPXX% < 1 Then PPXX% = 1
      ESTEAR1$ = Mid$(estearco$, PPXX%)
      '
      For j& = 1 To List3.ListCount
        OTROAR$ = TRIM$(Left$(List3.List(j& - 1), 64))
        '
        If OTROAR$ = ESTEAR$ Then
           OTROARCO$ = TRIM$(Mid$(List3.List(j& - 1), 65))
           'PPXX% = InStr(UCase$(OTROARCO$), "ARCHIVOS DE PROGRAMA")
           PPXX% = InStr(UCase$(OTROARCO$), "FLEXOFT")
           PPYY% = 20
           If PPXX% < 1 Then
              PPXX% = InStr(UCase$(OTROARCO$), "PROGRAM FILES")
              PPYY% = 13
           End If
           If PPXX% < 1 Then PPXX% = 1
           OTROAR1$ = Mid$(OTROARCO$, PPXX%)
           '
           'If Mid$(OTROAR1$, 1 + PPYY%) = Mid$(ESTEAR1$, 1 + 20) Then
           If Mid$(OTROAR1$, 1) = Mid$(ESTEAR1$, 59) Then
              FECHESTE# = 0: FECHOTRO# = 0
              On Error Resume Next
              FECHESTE# = FileDateTime(estearco$)
              FECHOTRO# = FileDateTime(OTROARCO$)
              On Error GoTo 0
              '
              If FECHESTE# = FECHOTRO# Then GoTo 35
              '
              ARCHESTE$ = LOADFILE$(estearco$)
              ARCHOTRO$ = LOADFILE$(OTROARCO$)
              If ARCHESTE$ <> ARCHOTRO$ Then GoTo 38
              '
35            aa2 = List2.List(ILISTA& - 1)
              aa3 = List3.List(j& - 1)
              List2.RemoveItem (ILISTA& - 1)
              List3.RemoveItem (j& - 1)
              Openforms = DoEvents
              GoTo 32
           End If
        End If
38    Next j&
39 Next ILISTA&
   '
   Call COMUNI("Finalizada la carga de diferencias encontradas")
   '
   For i& = 1 To List3.ListCount
      Debug.Print Mid$(List3.List(i& - 1), 64)
   Next i&
   Screen.MousePointer = 1
   '
End Sub

Private Sub Command10_Click()
   '
   Call ABRECONEXION
   Call ABRE_CORREC
   If TRIM$(List3.TEXT) = "" Then
      Call MENSERR(24, "Identifique Modulo")
      Exit Sub
   End If
   If TRIM$(List6.TEXT) = "" Then
      Call MENSERR(24, "Identifique Rutina")
      Exit Sub
   End If
   '
   COMMCORR.Label2 = List3.TEXT
      ARUT$ = TRIM$(List6.TEXT)
      PPXX% = InStr(ARUT$, "(")
      If PPXX% > 1 Then ARUT$ = TRIM$(Left$(ARUT$, PPXX% - 1))
      PPXX% = InStr(ARUT$, "'")
      If PPXX% > 1 Then ARUT$ = TRIM$(Left$(ARUT$, PPXX% - 1))
   COMMCORR.Label3 = ARUT$
   Select Case Combo1.TEXT
      Case "\\Ernesto\sistema (C)"
         COMMCORR.Label5 = "MC"
      Case "\\Desarrollo\Datos (D)"
         COMMCORR.Label5 = "OD"
      Case "\\Soporte\D"
         COMMCORR.Label5 = "MC"
   End Select
   Call COMMCORR.CARGACORREC
   COMMCORR.Show 1
End Sub

Private Sub Command11_Click()
   '
   Call ABRECONEXION
   Call ABRE_CORREC
   '
   Select Case Combo1.TEXT
      Case "\\Ernesto\sistema (C)"
         CORRCLEAR.Label5 = "MC"
      Case "\\Desarrollo\Datos (D)"
         CORRCLEAR.Label5 = "OD"
      Case "\\Soporte\D"
         CORRCLEAR.Label5 = "MC"
   End Select
   CORRCLEAR.Show 1
End Sub

Private Sub Command12_Click()
   '
   Command12.Enabled = False
   '
10 ABUSKE1$ = "FUNCTION "
   ABUSKE2$ = "SUB "
   '
   ESTEQ$ = "C:"
   Screen.MousePointer = 11
   ESTARUT$ = "C:\ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB\FLEXOFT\"
   '
   ' EXPLORA LOS DIRECTORIOS LOCALES
   List1.Clear: List2.Clear
   List1.AddItem ESTARUT$
   ILIST1& = 0
   '
   Do While ILIST1& < List1.ListCount
      RUTAX$ = List1.List(ILIST1&)
      On Error Resume Next
      NOMBREX$ = Dir(RUTAX$, 16) ' Recupera la primera entrada.
      If Err Then
         On Error GoTo 0
         GoTo 6
      End If
      On Error GoTo 0
      Do While NOMBREX$ <> "" ' Inicia el bucle.
        ' Ignora el directorio actual y el que lo abarca.
        If NOMBREX$ <> "." And NOMBREX$ <> ".." Then
           ' Utiliza comparación a nivel de bits para asegurarse de que MiNombre es un directorio.
           If (GetAttr(RUTAX$ & NOMBREX$) And vbDirectory) = vbDirectory Then
              If UCase$(RUTAX$ + NOMBREX$) <> ULODAY$ Then ' EXCLUYE DIRECTORIO TRABAJO
                 List1.AddItem RUTAX$ + NOMBREX$ + "\"
                 List1.ListIndex = List1.ListCount - 1
              End If
           End If ' solamente si representa un directorio.
        End If
        NOMBREX$ = Dir ' Obtiene siguiente entrada.
        Openforms = DoEvents
      Loop
6     ILIST1& = ILIST1& + 1
   Loop
   '
   ' CARGA EN LIST2 LOS ARCHIVOS LOCALES
   For i& = 1 To List1.ListCount
      File1.Path = List1.List(i& - 1)
      estaruta$ = File1.Path + "\"
      For j& = 1 To File1.ListCount
         FINAME$ = TRIM$(UCase$(File1.List(j& - 1)))
           If InStr(FINAME$, ".FRX") > 0 Then GoTo 17
           If InStr(FINAME$, ".VBW") > 0 Then GoTo 17
           If InStr(FINAME$, ".PDF") > 0 Then GoTo 17
         TTXX$ = AJUSTI$(FINAME$, 64)
         TTXX$ = TTXX$ + estaruta$ + FINAME$
         List2.AddItem TTXX$
         Openforms = DoEvents
17    Next j&
19 Next i&
   '
   For U& = 1 To List2.ListCount
      ACOMPLE$ = TRIM$(Mid$(List2.List(U& - 1), 65))
      ADESTI$ = "F:\CONVERTI" + Mid$(ACOMPLE$, 3)
      N1% = FILEOPEN%(ACOMPLE$, 1, 128)
      TXTX$ = ""
      Do While Not EOF(N1%)
         Line Input #N1%, AAA$
         If Left$(UCase$(AAA$), 9) = "FUNCTION " Or Left$(UCase$(AAA$), 4) = "SUB " Then
            Call CONVIERTE(AAA$)
         End If
         If TXTX$ <> "" Then TXTX$ = TXTX$ + vbCrLf
      Loop
      Close #N1%
      '
      Call SAVEFILE(ADESTI$, TXTX$)
      '
   Next U&
   '
99 Command12.Enabled = True
   '
End Sub

Sub CONVIERTE(XYZT$)
    ORI$ = XYZT$
    PPXX% = InStr(XYZT$, "(")
    If PPXX% < 1 Then GoTo 99
    '
    IZQUI$ = Left$(XYZT$, PPXX%)
    RESTO$ = Mid$(XYZT$, PPXX% + 1)
    If TRIM$(RESTO$) = "" Then GoTo 90
    '
10  PPZZ% = InStr(RESTO$, ",")
    If PPZZ% < 1 Then PPZZ% = InStr(RESTO$, ")")
    If PPZZ% < 1 Then PPZZ% = 1 + Len(RESTO$)
    '
    MEDIO$ = RTrim(Left$(RESTO$, PPZZ% - 1))
    If MEDIO$ = "" Then GoTo 80
    If InStr(UCase$(MEDIO$), " AS ") < 1 Then
       If Right$(MEDIO$, 1) = "%" Then
             MEDIO$ = Left$(MEDIO$, Len(MEDIO$) - 1) + " AS INTEGER"
         ElseIf Right$(MEDIO$, 1) = "$" Then
             MEDIO$ = Left$(MEDIO$, Len(MEDIO$) - 1) + " AS STRING"
         ElseIf Right$(MEDIO$, 1) = "&" Then
             MEDIO$ = Left$(MEDIO$, Len(MEDIO$) - 1) + " AS LONG"
         ElseIf Right$(MEDIO$, 1) = "#" Then
             MEDIO$ = Left$(MEDIO$, Len(MEDIO$) - 1) + " AS DOUBLE"
         Else
             MEDIO$ = MEDIO$ + " AS SINGLE"
       End If
    End If
    '
80  IZQUI$ = IZQUI$ + MEDIO$ + Mid$(RESTO$, PPZZ%, 1)
    RESTO$ = Mid$(RESTO$, PPZZ% + 1)
    If TRIM$(RESTO$) <> "" Then GoTo 10
    '
90  XYZT$ = IZQUI$
    If XYZT$ <> ORI$ Then
       AAA = BBB
    End If
    '
99 End Sub

Private Sub Command13_Click()
  '
   Command13.Enabled = False
   '
10 ABUSKE1$ = "FUNCTION "
   ABUSKE2$ = "SUB "
   '
   ESTEQ$ = "C:"
   Screen.MousePointer = 11
   ESTARUT$ = "D:\ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB\FLEXOFT\"
   '
   ' EXPLORA LOS DIRECTORIOS LOCALES
   List1.Clear: List2.Clear
   List1.AddItem ESTARUT$
   ILIST1& = 0
   '
   Do While ILIST1& < List1.ListCount
      RUTAX$ = List1.List(ILIST1&)
      On Error Resume Next
      NOMBREX$ = Dir(RUTAX$, 16) ' Recupera la primera entrada.
      If Err Then
         On Error GoTo 0
         GoTo 6
      End If
      On Error GoTo 0
      Do While NOMBREX$ <> "" ' Inicia el bucle.
        ' Ignora el directorio actual y el que lo abarca.
        If NOMBREX$ <> "." And NOMBREX$ <> ".." Then
           ' Utiliza comparación a nivel de bits para asegurarse de que MiNombre es un directorio.
           If (GetAttr(RUTAX$ & NOMBREX$) And vbDirectory) = vbDirectory Then
              If UCase$(RUTAX$ + NOMBREX$) <> ULODAY$ Then ' EXCLUYE DIRECTORIO TRABAJO
                Debug.Print NOMBREX$
                 List1.AddItem RUTAX$ + NOMBREX$ + "\"
                 List1.ListIndex = List1.ListCount - 1
              End If
           End If ' solamente si representa un directorio.
        End If
        NOMBREX$ = Dir ' Obtiene siguiente entrada.
        Openforms = DoEvents
      Loop
6     ILIST1& = ILIST1& + 1
   Loop
   '
   ' CARGA EN LIST2 LOS ARCHIVOS LOCALES
   For i& = 1 To List1.ListCount
      File1.Path = List1.List(i& - 1)
      estaruta$ = File1.Path + "\"
      For j& = 1 To File1.ListCount
         FINAME$ = TRIM$(UCase$(File1.List(j& - 1)))
'           If InStr(FINAME$, ".FRX") > 0 Then GoTo 17
'           If InStr(FINAME$, ".VBW") > 0 Then GoTo 17
'           If InStr(FINAME$, ".PDF") > 0 Then GoTo 17
         If InStr(FINAME$, ".FRM") < 1 Then GoTo 17
         TTXX$ = AJUSTI$(FINAME$, 64)
         TTXX$ = TTXX$ + estaruta$ + FINAME$
         List2.AddItem TTXX$
         Openforms = DoEvents
17    Next j&
19 Next i&
   '
   For U& = 1 To List2.ListCount
      ACOMPLE$ = TRIM$(Mid$(List2.List(U& - 1), 65))
      ADESTI$ = "F:\CONVERSK" + Mid$(ACOMPLE$, 3)
      N1% = FILEOPEN%(ACOMPLE$, 1, 128)
      TXTX$ = "": CANREN& = 0: LLAMASKIN% = 0: CREAPIC% = 0: INCRUSTASKIN% = 0: REFERCOMPON% = 0
      Do While Not EOF(N1%)
         CANREN& = CANREN& + 1
         Line Input #N1%, RENGLOX$(CANREN&)
         If TRIM$(UCase$(RENGLOX$(CANREN&))) = "BEGIN VB.PICTUREBOX PICTURE999" Then CREAPIC% = 1
         If TRIM$(UCase$(RENGLOX$(CANREN&))) = "BEGIN ACTIVESKINLIBCTL.SKIN" Then INCRUSTASKIN% = 1
         If TRIM$(UCase$(RENGLOX$(CANREN&))) = "CALL APLICACIONSKINS" Then LLAMASKIN% = 1
         If InStr(UCase$(RENGLOX$(CANREN&)), "ACTSKIN4.OCX") > 0 Then REFERCOMPON% = 1
      Loop
      Close #N1%
      '
      If LLAMASKIN% + CREAPIC% + INCRUSTASKIN% + REFERCOMPON% < 4 Then
         Call CONVERSKIN    ' CONVIERTE Y GRABA
      End If
      '
90 Next U&
   '
99 Command13.Enabled = True
End Sub

Sub CONVERSKIN()

     N3% = FILEOPEN%(ADESTI$, 2, 128)
     'RECORDAR DE CERRAR
     TXPICTURE$ = "Begin VB.PictureBox Picture999" + vbCrLf
     TXPICTURE$ = TXPICTURE$ + "Height = 495" + vbCrLf
     TXPICTURE$ = TXPICTURE$ + "Index = 0" + vbCrLf
     TXPICTURE$ = TXPICTURE$ + "Left = 6480" + vbCrLf
     TXPICTURE$ = TXPICTURE$ + "ScaleHeight = 435" + vbCrLf
     TXPICTURE$ = TXPICTURE$ + "ScaleWidth = 1155" + vbCrLf
     TXPICTURE$ = TXPICTURE$ + "Top = 0" + vbCrLf
     TXPICTURE$ = TXPICTURE$ + "Visible = 0" + vbCrLf
     TXPICTURE$ = TXPICTURE$ + "Width = 1215" + vbCrLf
     TXPICTURE$ = TXPICTURE$ + "End"
     '
     TXSKIN$ = "Begin ACTIVESKINLibCtl.Skin Skin1" + vbCrLf
     TXSKIN$ = TXSKIN$ + "Left = 5760" + vbCrLf
     TXSKIN$ = TXSKIN$ + "OleObjectBlob   =   " & Chr$(34) & "SITSTOKskin.frx" & Chr$(34) & ":014A" + vbCrLf
     TXSKIN$ = TXSKIN$ + "Top = 3120" + vbCrLf
     TXSKIN$ = TXSKIN$ + "End"
     '
     TXREFERCOMPON$ = "Object = " & Chr$(34) & "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0" & Chr$(34) & "; " & Chr$(34) & "actskin4.ocx" & Chr$(34) & ""
     '
     FIN& = CANREN&
     
     For i& = 0 To FIN&
'       If LLAMASKIN% + CREAPIC% + INCRUSTASKIN% + REFERCOMPON% = 4 Then Exit For
       FILA$ = RENGLOX$(i&)
       If REFERCOMPON% = 0 Then
          If InStr(UCase$(FILA$), "VERSION") > 0 Then
             FILA$ = FILA$ + vbCrLf + TXREFERCOMPON$
             REFERCOMPON% = 1
          End If
       End If
       '
       'AGREGA DESPUES DE LA LEYENDA PRIVATE SUB FORM_LOAD LA LLAMADA
       If LLAMASKIN% = 0 Then
        If InStr(UCase$(FILA$), "PRIVATE SUB FORM_LOAD") > 0 Then
            FILA$ = FILA$ + vbCrLf + "CALL APLICACIONSKINS"
            LLAMASKIN% = 1
         End If
       End If
       '
       'AGREGA EL OBJETO PICUTRE
       If CREAPIC% = 0 Then
         If InStr(UCase$(FILA$), "BEGIN VB.MENU") > 0 Then
            Print #N3%, TXPICTURE$
            CREAPIC% = 1
         End If
         '
         If InStr(UCase$(RENGLOX$(i& + 2)), "ATTRIBUTE VB_NAME") > 0 Then
            Print #N3%, TXPICTURE$
            CREAPIC% = 1
         End If
       End If
         
       'INCRUSTA EL SKIN
       If INCRUSTASKIN% = 0 Then
         If InStr(UCase$(FILA$), "BEGIN VB.MENU") > 0 Then
            Print #N3%, TXSKIN$
            INCRUSTASKIN% = 1
         End If
       '
         If InStr(UCase$(RENGLOX$(i& + 2)), "ATTRIBUTE VB_NAME") > 0 Then
            Print #N3%, TXSKIN$
            INCRUSTASKIN% = 1
         End If
       End If
       
       Print #N3%, FILA$
     Next i&
     Close #N3%
End Sub


Private Sub Command14_Click()
   CARCHI% = 0
   '
   OTROEQ$ = TRIM$(Combo1.TEXT)
   If OTROEQ$ = "" Then
      MsgBox "Identifique Equipo Remoto !!!"
      Exit Sub
   End If
   '
   ESTEQ$ = "C:"
   DIRFSX$ = OTROEQ$ + "\ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB\FLEXOFT"
   If OTROEQ$ = "\\Notebook-bauer\C" Then
      DIRFSX$ = OTROEQ$ + "\Program Files\DEVSTUDIO\VB\FLEXOFT"
   ElseIf OTROEQ$ = "z:\" Then
      DIRFSX$ = OTROEQ$ + "FLEXOFT"
   End If
   DIEXPLO$ = DIRFSX$ + "\"
   '
   Screen.MousePointer = 11
   '
   ESTARUT$ = "L:\Archivos de Programa\Microsoft Visual Studio\DevStudio\VB\FLEXOFT\"
   'OTRARUT$ = OTROEQ$ + "\ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB\FLEXOFT\"
   OTRARUT$ = OTROEQ$ + "\Microsoft Visual Studio\VB98\FLEXOFT\"
   If OTROEQ$ = "\\Notebook-bauer\C" Then
      OTRARUT$ = OTROEQ$ + "\Program Files\DEVSTUDIO\VB\FLEXOFT\"
   ElseIf OTROEQ$ = "z:\" Then
      OTRARUT$ = OTROEQ$ + "FLEXOFT\"
   End If
   '
   ' EXPLORA LOS DIRECTORIOS LOCALES
   List1.Clear: List2.Clear
   List1.AddItem ESTARUT$
   ILIST1& = 0
   '
   Do While ILIST1& < List1.ListCount
      RUTAX$ = List1.List(ILIST1&)
      On Error Resume Next
      NOMBREX$ = Dir(RUTAX$, 16) ' Recupera la primera entrada.
      If Err Then
         On Error GoTo 0
         GoTo 6
      End If
      On Error GoTo 0
      Do While NOMBREX$ <> "" ' Inicia el bucle.
        ' Ignora el directorio actual y el que lo abarca.
        If NOMBREX$ <> "." And NOMBREX$ <> ".." Then
           ' Utiliza comparación a nivel de bits para asegurarse de que MiNombre es un directorio.
           If (GetAttr(RUTAX$ & NOMBREX$) And vbDirectory) = vbDirectory Then
              If UCase$(RUTAX$ + NOMBREX$) <> ULODAY$ Then ' EXCLUYE DIRECTORIO TRABAJO
                 List1.AddItem RUTAX$ + NOMBREX$ + "\"
                 List1.ListIndex = List1.ListCount - 1
              End If
           End If ' solamente si representa un directorio.
        End If
        NOMBREX$ = Dir ' Obtiene siguiente entrada.
        Openforms = DoEvents
      Loop
6     ILIST1& = ILIST1& + 1
   Loop
   '
   ' CARGA EN LIST2 LOS ARCHIVOS LOCALES
   For i& = 1 To List1.ListCount
      'Dir1.Path = List1.List(I& - 1)
      File1.Path = List1.List(i& - 1)
      File1.Refresh
      estaruta$ = File1.Path + "\"
      For j& = 1 To File1.ListCount
         FINAME$ = TRIM$(UCase$(File1.List(j& - 1)))
           If InStr(FINAME$, ".FRX") > 0 Then GoTo 17
           If InStr(FINAME$, ".VBW") > 0 Then GoTo 17
           If InStr(FINAME$, ".PDF") > 0 Then GoTo 17
           If InStr(FINAME$, ".PRF") > 0 Then GoTo 17
           If InStr(FINAME$, ".LOG") > 0 Then GoTo 17
           
         TTXX$ = AJUSTI$(FINAME$, 64)
         TTXX$ = TTXX$ + estaruta$ + FINAME$
         List2.AddItem TTXX$
         Openforms = DoEvents
17    Next j&
19 Next i&
   '
   ' EXPLORA DIRECTORIO DEL OTRO EQUIPO
   List1.Clear: List3.Clear
   List1.AddItem OTRARUT$
   ILIST1& = 0
   '
   Do While ILIST1& < List1.ListCount
      RUTAX$ = List1.List(ILIST1&)
      On Error Resume Next
      NOMBREX$ = Dir(RUTAX$, 16) ' Recupera la primera entrada.
      If Err Then
         On Error GoTo 0
         GoTo 6
      End If
      On Error GoTo 0
      Do While NOMBREX$ <> "" ' Inicia el bucle.
        ' Ignora el directorio actual y el que lo abarca.
        If NOMBREX$ <> "." And NOMBREX$ <> ".." Then
           ' Utiliza comparación a nivel de bits para asegurarse de que MiNombre es un directorio.
           If (GetAttr(RUTAX$ & NOMBREX$) And vbDirectory) = vbDirectory Then
              If UCase$(RUTAX$ + NOMBREX$) <> ULODAY$ Then ' EXCLUYE DIRECTORIO TRABAJO
                 List1.AddItem RUTAX$ + NOMBREX$ + "\"
                 List1.ListIndex = List1.ListCount - 1
              End If
           End If ' solamente si representa un directorio.
        End If
        NOMBREX$ = Dir ' Obtiene siguiente entrada.
        Openforms = DoEvents
      Loop
      ILIST1& = ILIST1& + 1
   Loop
   '
   ' CARGA EN LIST3 LOS ARCHIVOS REMOTOS
   For i& = 1 To List1.ListCount
      Dir1.Path = List1.List(i& - 1)
      Dir1.Refresh
      File1.Path = Dir1.Path
      File1.Refresh
      otraruta$ = File1.Path + "\"
      'File1.Refresh
      For j& = 1 To File1.ListCount
         FINAME$ = TRIM$(UCase$(File1.List(j& - 1)))
         If FINAME$ = "BADASTRU.MDB" Or InStr(FINAME$, ".LOG") > 0 Or InStr(FINAME$, ".PRF") > 0 Then
            ACOMPLE$ = OTRARUT$ + FINAME$
            On Error Resume Next
            Kill ACOMPLE$
            On Error GoTo 0
            GoTo 27
         End If
            '
            If InStr(FINAME$, ".FRX") > 0 Then GoTo 27
            If InStr(FINAME$, ".VBW") > 0 Then GoTo 27
            If InStr(FINAME$, ".PDF") > 0 Then GoTo 27
            If InStr(FINAME$, ".PRF") > 0 Then GoTo 27
            If InStr(FINAME$, ".LOG") > 0 Then GoTo 27
         TTXX$ = AJUSTI$(FINAME$, 64)
         TTXX$ = TTXX$ + otraruta$ + FINAME$
         List3.AddItem TTXX$
         Openforms = DoEvents
27    Next j&
29 Next i&

   ' DEPURA LOS QUE COINCIDEN
   For ILISTA& = 1 To List2.ListCount
32    ESTEAR$ = TRIM$(Left$(List2.List(ILISTA& - 1), 64))
      estearco$ = TRIM$(Mid$(List2.List(ILISTA& - 1), 65))
      PPXX% = InStr(UCase$(estearco$), "\FLEXOFT")
      If PPXX% < 1 Then Exit Sub
      ESTEAR1$ = Combo1.TEXT & Mid$(estearco$, PPXX%)
      '
      FileName = Dir(ESTEAR1$)
      If FileName = "" Then
        Call FileCopy(ESTEAR$, ESTEAR1$)
      End If
39 Next ILISTA&


End Sub

Private Sub Command2_Click()
   List4.Selected(List4.ListIndex) = False
   List5.Selected(List5.ListIndex) = False
   List4.ListIndex = List4.ListIndex + 1
   List5.ListIndex = List5.ListIndex + 1
   List4.Selected(List4.ListIndex) = True
   List5.Selected(List5.ListIndex) = True
   List4.Refresh
   List5.Refresh
   Call Command4_Click
End Sub

Private Sub Command3_Click()
    '
    ALOCA$ = "": AREMO$ = ""
    ALO$ = UCase$(TRIM$(Left$(List2.TEXT, 64)))
    ALOCA$ = TRIM$(Mid$(List2.TEXT, 65))
    '
10  For KKI% = 1 To List3.ListCount
       If TRIM$(UCase$(Left$(List3.List(KKI% - 1), 64))) = ALO$ Then
          List3.ListIndex = KKI% - 1
          AREMO$ = TRIM$(Mid$(List3.List(KKI% - 1), 65))
          GoTo 80
       End If
    Next KKI%
    AREMO$ = TRIM$(Combo1.TEXT) + Mid$(ALOCA$, 3)
    If VUELTA% = 0 Then
       VUELTA% = 1
       TTXX$ = Left$(List2.TEXT, 64) + AREMO$
       List3.AddItem TTXX$
       GoTo 10
    End If
    '
80  If ALOCA$ <> "" Then
      If AREMO$ <> "" Then
        Call COMPATOTA(ALOCA$, AREMO$)
      End If
    End If
    '
End Sub

Private Sub Command4_Click()
   ini1% = List4.ListIndex
   If ini1% < 0 Then Exit Sub
   If List4.Selected(ini1%) = False Then Exit Sub
   '
   INI2% = List5.ListIndex
   If INI2% < 0 Then Exit Sub
   If List5.Selected(INI2%) = False Then Exit Sub
   '
   If UCase$(TRIM$(List4.TEXT)) <> UCase$(TRIM$(List5.TEXT)) Then
      Call MENSERR(24, "Debe Iniciar en Registros Coincidentes.")
      Exit Sub
   End If
   '
   For KI1% = ini1% To List4.ListCount - 1
     List4.Selected(KI1%) = False
   Next KI1%
   '
   For KI2% = INI2% To List5.ListCount - 1
     List5.Selected(KI2%) = False
   Next KI2%
   '
   JJ% = 0
5  JJ% = JJ% + 1
'   If UCase$(LTrim(List4.List(ini1% + JJ%))) <> UCase$(LTrim(List5.List(INI2% + JJ%))) Then
6  LX4$ = REPLACAR$(UCase$(LTrim(List4.List(ini1% + JJ%))), " ", "")
   LX5$ = REPLACAR$(UCase$(LTrim(List5.List(INI2% + JJ%))), " ", "")
   If InStr(LX4$, "RETAINED") > 0 Then ini1% = ini1% + 1: GoTo 6
   If InStr(LX5$, "RETAINED") > 0 Then INI2% = INI2% + 1: GoTo 6
   '
   If LX4$ <> LX5$ Then
      ABUXI$ = "OLEAUTOMATION": If InStr(LX4$, ABUXI$) > 0 And InStr(LX5$, ABUXI$) > 0 Then GoTo 9
      ABUXI$ = "MICROSOFTDAO3.6OBJECTLIBRARY": If InStr(LX4$, ABUXI$) > 0 And InStr(LX5$, ABUXI$) > 0 Then GoTo 9
      ABUXI$ = "MICROSOFTACTIVEXDATAOBJECTS2.6LIBRARY": If InStr(LX4$, ABUXI$) > 0 And InStr(LX5$, ABUXI$) > 0 Then GoTo 9
      ' MsgBox LX4$ + vbCrLf + LX5$
      On Error Resume Next
      List4.Selected(ini1% + JJ%) = True
      On Error GoTo 0
      If List4.TopIndex < ini1% + JJ% - 20 Then List4.TopIndex = ini1% + JJ% - 20
      Text1 = List4.TEXT
      On Error Resume Next
      List5.Selected(INI2% + JJ%) = True
      On Error GoTo 0
      If List5.TopIndex < INI2% + JJ% - 20 Then List5.TopIndex = INI2% + JJ% - 20
      Text2 = List5.TEXT
      Exit Sub
   End If
9  If InStr(UCase$(LTrim(List4.List(ini1% + JJ%))), "END SUB") > 0 Then GoTo 80
   If InStr(UCase$(LTrim(List4.List(ini1% + JJ%))), "END FUNCTION") > 0 Then GoTo 80
   If ini1% + JJ% >= List4.ListCount Then
     If INI2% + JJ% >= List5.ListCount Then
       GoTo 80
     End If
   End If
   GoTo 5
   '
80 MsgBox "No Más Diferencias"
   '
End Sub

Private Sub Command5_Click()
   ILIST1& = List2.ListIndex
   ILIST2& = List3.ListIndex
   If TRIM$(Left$(List2.TEXT, 64)) <> TRIM$(Left$(List3.TEXT, 64)) Then GoTo 90
   '
   If ILIST1& >= 0 Then
      If ILIST2& >= 0 Then
        ESTEAR$ = Label10
        ACOMPLE$ = Label11
        If Left$(ACOMPLE$, 40) = "\\Notebook-bauer\C\ARCHIVOS DE PROGRAMA\" Then
           ACOMPLE$ = "\\Notebook-bauer\C\PROGRAM FILES\" + Mid$(ACOMPLE$, 41)
        End If
        '
        ' RESPALDA LA VERSION ANTERIOR
        ARESPA$ = ACOMPLE$
        Call GUARDAVIE(ARESPA$)
        '
        On Error Resume Next
        FileCopy ESTEAR$, ACOMPLE$
        If Err Then
           On Error GoTo 0
           MsgBox "Imposible Copiar " + ESTEAR$ + " hacia " + ACOMPLE$
        End If
        On Error GoTo 0
        '
        PPXX% = InStr(ESTEAR$, ".FRM")
        PPYY% = InStr(ACOMPLE$, ".FRM")
        If PPXX% > 0 Then
          If PPYY% > 0 Then
            ESTEAR$ = Left$(ESTEAR$, PPXX% - 1) + ".FRX" + Mid$(ESTEAR$, PPXX% + 4)
            ACOMPLE$ = Left$(ACOMPLE$, PPYY% - 1) + ".FRX" + Mid$(ACOMPLE$, PPYY% + 4)
            On Error Resume Next               ' POR SI NO EXISTE EL .FRX O .VBW
            FileCopy ESTEAR$, ACOMPLE$
            On Error GoTo 0
          End If
        End If
        '
        PPXX% = InStr(ESTEAR$, ".VBP")
        PPYY% = InStr(ACOMPLE$, ".VBP")
        If PPXX% > 0 Then
          If PPYY% > 0 Then
            ESTEAR$ = Left$(ESTEAR$, PPXX% - 1) + ".VBW" + Mid$(ESTEAR$, PPXX% + 4)
            ACOMPLE$ = Left$(ACOMPLE$, PPYY% - 1) + ".VBW" + Mid$(ACOMPLE$, PPYY% + 4)
            On Error Resume Next
            FileCopy ESTEAR$, ACOMPLE$
            If Err Then
               On Error GoTo 0
               MsgBox "Imposible Copiar " + ESTEAR$ + " hacia " + ACOMPLE$
            End If
            On Error GoTo 0
          End If
        End If
        '
        List2.RemoveItem ILIST1&: List2.Refresh
        List3.RemoveItem ILIST2&: List3.Refresh
        List4.Clear
        List5.Clear
        List6.Clear
        Text1 = ""
        Text2 = ""
        Label10 = ""
        Label11 = ""
        '
        DoEvents
        Exit Sub
      End If
   End If
90 MsgBox "Imposible Copiar"
End Sub

Sub GUARDAVIE(RUTCH$)
   EXXI% = EXISTE%(RUTCH$)
   If EXXI% > 0 Then
     CARPEVI$ = "F:\FUENVIE\"
     PPXX% = InStr(RUTCH$, "VB\FLEXOFT\")
     If PPXX% > 0 Then
       RUTVI$ = Mid$(RUTCH$, PPXX%)
10     PPYY% = InStr(RUTVI$, "\")
       If PPYY% > 1 Then
         CARPEVI$ = CARPEVI$ + Left$(RUTVI$, PPYY% - 1)
         On Error Resume Next
         MkDir CARPEVI$
         On Error GoTo 0
         CARPEVI$ = CARPEVI$ + "\"
         RUTVI$ = Mid$(RUTVI$, PPYY% + 1)
         GoTo 10
       End If
       '
       On Error Resume Next
       FIDAT& = FileDateTime(RUTCH$)
       FIDAV& = FileDateTime(CARPEVI$ + RUTVI$)
       On Error GoTo 0
       '
       If FIDAT& = FIDAV& Or FIDAV& = 0 Then ' COPIA DIRECTAMENTE
          GoTo 30
       End If
       '
       ANAMCO$ = Left$(CARPEVI$ + RUTVI$, Len(CARPEVI$ + RUTVI$) - 1)
       '
       ' REVISAR SI YA HAY UNO IDENTICO RESPALDADO
       ARES$ = LOADFILE$(RUTCH$)
       For KI% = 0 To 9
         If EXISTE%(ANAMCO$ + TRIM$(Str$(KI%))) > 0 Then
           AVIE$ = LOADFILE$(ANAMCO$ + TRIM$(Str$(KI%)))
           If AVIE$ = ARES$ Then Exit Sub ' HAY UNO IDENTICO - NO SE RESPALDA
         End If
       Next KI%
       '
       On Error Resume Next
       Kill ANAMCO$ + "9"
       On Error GoTo 0
       If EXISTE%(ANAMCO$ + "9") > 0 Then
          MsgBox "Imposible Instalar. No Pudo Borrarse " + ANAMCO$ + "9"
          Stop
       End If
       '
       On Error Resume Next
       For KI% = 8 To 0 Step -1
          Name ANAMCO$ + TRIM$(Str$(KI%)) As ANAMCO$ + TRIM$(Str$(KI% + 1))
       Next KI%
       On Error GoTo 0
       If EXISTE%(DIREVIE$ + ANAMCO$ + "0") > 0 Then
          MsgBox "Imposible Instalar. No Pudo Renombrarse " + ANAMCO$ + "0 como " + ANAMCO$ + "1"
       End If
       '
       On Error Resume Next
       Name CARPEVI$ + RUTVI$ As ANAMCO$ + "0"
       On Error GoTo 0
       '
30     On Error Resume Next
       FileCopy RUTCH$, CARPEVI$ + RUTVI$
       On Error GoTo 0
       '
     End If
   End If
End Sub

Private Sub Command6_Click()
   ILIST1& = List2.ListIndex
   ILIST2& = List3.ListIndex
   '
   If ILIST1& >= 0 And ILIST2& >= 0 Then
     If TRIM$(Left$(List2.TEXT, 64)) <> TRIM$(Left$(List3.TEXT, 64)) Then GoTo 90
   End If
   '
   If ILIST2& >= 0 Then
     ESTEAR$ = TRIM$(Label10)
     ACOMPLE$ = TRIM$(Label11)
     If ACOMPLE$ <> "" Then
        If ESTEAR$ = "" Then
          PPXX% = InStr(UCase$(ACOMPLE$), "\ARCHIVOS DE PRO")
          If PPXX% < 1 Then GoTo 90
          ESTEAR$ = "C:" + Mid$(ACOMPLE$, PPXX%)
        End If
        '
        ARESPA$ = ESTEAR$
        Call GUARDAVIE(ARESPA$)
        '
        On Error Resume Next
        FileCopy ACOMPLE$, ESTEAR$
        If Err Then
           On Error GoTo 0
           GoTo 90
        End If
        On Error GoTo 0
        '
        PPXX% = InStr(ESTEAR$, ".FRM")
        PPYY% = InStr(ACOMPLE$, ".FRM")
        If PPXX% > 0 Then
          If PPYY% > 0 Then
            ESTEAR$ = Left$(ESTEAR$, PPXX% - 1) + ".FRX" + Mid$(ESTEAR$, PPXX% + 4)
            ACOMPLE$ = Left$(ACOMPLE$, PPYY% - 1) + ".FRX" + Mid$(ACOMPLE$, PPYY% + 4)
            On Error Resume Next
            FileCopy ACOMPLE$, ESTEAR$
            On Error GoTo 0
          End If
        End If
        '
        PPXX% = InStr(ESTEAR$, ".VBP")
        PPYY% = InStr(ACOMPLE$, ".VBP")
        If PPXX% > 0 Then
          If PPYY% > 0 Then
            ESTEAR$ = Left$(ESTEAR$, PPXX% - 1) + ".VBW" + Mid$(ESTEAR$, PPXX% + 4)
            ACOMPLE$ = Left$(ACOMPLE$, PPYY% - 1) + ".VBW" + Mid$(ACOMPLE$, PPYY% + 4)
            On Error Resume Next
            FileCopy ACOMPLE$, ESTEAR$
            On Error GoTo 0
          End If
        End If
        '
        If ILIST1& >= 0 Then List2.RemoveItem ILIST1&
        If ILIST2& >= 0 Then List3.RemoveItem ILIST2&
        List4.Clear
        List5.Clear
        List6.Clear
        Text1 = ""
        Text2 = ""
        Label10 = ""
        Label11 = ""
        '
        DoEvents
        Exit Sub
     End If
   End If
   '
90 MsgBox "Imposible Copiar"
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   '
   CARCHI% = 0
   ESTEQ$ = "C:"
   Screen.MousePointer = 11
   '
   ESTARUT$ = "C:\ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB\FLEXOFT\"
   '
   ' EXPLORA LOS DIRECTORIOS LOCALES
   List1.Clear: List2.Clear
   List1.AddItem ESTARUT$
   ILIST1& = 0
   '
   Do While ILIST1& < List1.ListCount
      RUTAX$ = List1.List(ILIST1&)
      On Error Resume Next
      NOMBREX$ = Dir(RUTAX$, 16) ' Recupera la primera entrada.
      If Err Then
         On Error GoTo 0
         GoTo 6
      End If
      On Error GoTo 0
      Do While NOMBREX$ <> "" ' Inicia el bucle.
        ' Ignora el directorio actual y el que lo abarca.
        If NOMBREX$ <> "." And NOMBREX$ <> ".." Then
           ' Utiliza comparación a nivel de bits para asegurarse de que MiNombre es un directorio.
           If (GetAttr(RUTAX$ & NOMBREX$) And vbDirectory) = vbDirectory Then
              If UCase$(RUTAX$ + NOMBREX$) <> ULODAY$ Then ' EXCLUYE DIRECTORIO TRABAJO
                 List1.AddItem RUTAX$ + NOMBREX$ + "\"
                 List1.ListIndex = List1.ListCount - 1
              End If
           End If ' solamente si representa un directorio.
        End If
        NOMBREX$ = Dir ' Obtiene siguiente entrada.
        Openforms = DoEvents
      Loop
6     ILIST1& = ILIST1& + 1
   Loop
   '
   ' CARGA EN LIST2 LOS ARCHIVOS LOCALES
   For i& = 1 To List1.ListCount
      File1.Path = List1.List(i& - 1)
      estaruta$ = File1.Path + "\"
      For j& = 1 To File1.ListCount
         FINAME$ = TRIM$(UCase$(File1.List(j& - 1)))
           If InStr(FINAME$, ".FRX") > 0 Then GoTo 17
           If InStr(FINAME$, ".VBW") > 0 Then GoTo 17
           If InStr(FINAME$, ".PDF") > 0 Then GoTo 17
         TTXX$ = AJUSTI$(FINAME$, 64)
         TTXX$ = TTXX$ + estaruta$ + FINAME$
         List2.AddItem TTXX$
         Openforms = DoEvents
17    Next j&
19 Next i&
   '
   For U& = 1 To List2.ListCount
      ACOMPLE$ = TRIM$(Mid$(List2.List(U& - 1), 65))
      N1% = FILEOPEN%(ACOMPLE$, 1, 128)
      N2% = FILEOPEN%("C:\ASALID.WKF", 2, 128)
      HUBOREP% = 0
      Do While Not EOF(N1%)
        Line Input #N1%, XX$
10      PPXX% = InStr(UCase$(XX$), "FILTSQL$(")
        If PPXX% > 0 Then
          If Mid$(XX$, PPXX% - 1, 1) <> Chr$(34) Then
            If Left$(TRIM$(XX$), 8) <> "Function" Then
              IZQUI$ = Left$(XX$, PPXX% - 1)
              DERE$ = Mid$(XX$, PPXX% + 9)
              PPYY% = InStr(DERE$, ")"): If PPYY% < 1 Then PPYY% = 1 + Len(DERE$)
              DERE$ = Left$(DERE$, PPYY% - 1) + Mid$(DERE$, PPYY% + 1)
              YY$ = IZQUI$ + DERE$
              ' MsgBox XX$ + VBCRFL + vbCrLf + YY$
              XX$ = YY$
              HUBOREP% = 1
              GoTo 10
            End If
          End If
        End If
        Print #N2%, XX$
      Loop
      Close #N1%
      Close #N2%
      '
      If HUBOREP% = 1 Then
        ZZ$ = LOADFILE$("C:\ASALID.WKF")
        Call SAVEFILE(ACOMPLE$, ZZ$)
      End If
   Next U&
   '
   Screen.MousePointer = 1
   Call COMUNI("Conversion Completa")
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   '
   CARCHI% = 0
   ESTEQ$ = "C:"
   Screen.MousePointer = 11
   '
   ESTARUT$ = "C:\ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB\FLEXOFT\"
   If EXISTE%("D:\ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB\FLEXOFT\.") > 0 Then
      ESTARUT$ = "D:\ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB\FLEXOFT\"
      ESTEQ$ = "D:"
   End If
   MsgBox ESTARUT$
   '
   ' EXPLORA LOS DIRECTORIOS LOCALES
   List1.Clear: List2.Clear
   List1.AddItem ESTARUT$
   ILIST1& = 0
   '
   Do While ILIST1& < List1.ListCount
      RUTAX$ = List1.List(ILIST1&)
      On Error Resume Next
      NOMBREX$ = Dir(RUTAX$, 16) ' Recupera la primera entrada.
      If Err Then
         On Error GoTo 0
         GoTo 6
      End If
      On Error GoTo 0
      Do While NOMBREX$ <> "" ' Inicia el bucle.
        ' Ignora el directorio actual y el que lo abarca.
        If NOMBREX$ <> "." And NOMBREX$ <> ".." Then
           ' Utiliza comparación a nivel de bits para asegurarse de que MiNombre es un directorio.
           If (GetAttr(RUTAX$ & NOMBREX$) And vbDirectory) = vbDirectory Then
              If UCase$(RUTAX$ + NOMBREX$) <> ULODAY$ Then ' EXCLUYE DIRECTORIO TRABAJO
                 List1.AddItem RUTAX$ + NOMBREX$ + "\"
                 List1.ListIndex = List1.ListCount - 1
              End If
           End If ' solamente si representa un directorio.
        End If
        NOMBREX$ = Dir ' Obtiene siguiente entrada.
        Openforms = DoEvents
      Loop
6     ILIST1& = ILIST1& + 1
   Loop
   '
   ' CARGA EN LIST2 LOS ARCHIVOS LOCALES
   For i& = 1 To List1.ListCount
      File1.Path = List1.List(i& - 1)
      estaruta$ = File1.Path + "\"
      For j& = 1 To File1.ListCount
         FINAME$ = TRIM$(UCase$(File1.List(j& - 1)))
           If InStr(FINAME$, ".FRX") > 0 Then GoTo 17
           If InStr(FINAME$, ".VBW") > 0 Then GoTo 17
           If InStr(FINAME$, ".PDF") > 0 Then GoTo 17
         TTXX$ = AJUSTI$(FINAME$, 64)
         TTXX$ = TTXX$ + estaruta$ + FINAME$
         List2.AddItem TTXX$
         Openforms = DoEvents
17    Next j&
19 Next i&
   '
   For U& = 1 To List2.ListCount
      ACOMPLE$ = TRIM$(Mid$(List2.List(U& - 1), 65))
      If InStr(ACOMPLE$, "04.") > 0 Then GoTo 95
      '
      N1% = FILEOPEN%(ACOMPLE$, 1, 128)
      N2% = FILEOPEN%("C:\ASALID.WKF", 2, 128)
      HUBOREP% = 0
      Do While Not EOF(N1%)
        Line Input #N1%, XX$
If Left$(TRIM$(XX$), 1) = "'" Then GoTo 80     ' COMENTADAS PASA DIRECTO
        PPXX% = InStr(UCase$(XX$), "FLEXCONN.EXECUTE")
        If PPXX% > 0 Then
           If InStr(UCase$(XX$), "FLEXCONN.EXECUTE (FILTSQL$") = PPXX% Then GoTo 48
           If Mid$(XX$, PPXX% - 1, 1) = Chr$(34) Then GoTo 48
           RESTO$ = TRIM$(Mid$(XX$, PPXX% + 16))
           PRICA$ = Left$(RESTO$, 1)
           If PRICA$ = Chr$(34) Then
              KKK% = 0
              For JJ% = 1 To Len(RESTO$)
                If Mid$(RESTO$, JJ%, 1) = "(" Then KKK% = KKK% + 1
                If Mid$(RESTO$, JJ%, 1) = ")" Then KKK% = KKK% - 1
                If KKK% = 0 Then
                   If Mid$(RESTO$, JJ%, 1) = "'" Then
                      RESTO$ = "(" + Left$(RESTO$, JJ%) + "))" + Mid$(RESTO$, JJ% + 1)
                      GoTo 25
                   End If
                End If
              Next JJ%
              RESTO$ = "(" + RESTO$ + "))"
              GoTo 25
            ElseIf PRICA$ = "(" Then
              KKK% = 0
              For JJ% = 1 To Len(RESTO$)
                If Mid$(RESTO$, JJ%, 1) = "(" Then KKK% = KKK% + 1
                If Mid$(RESTO$, JJ%, 1) = ")" Then KKK% = KKK% - 1
                If KKK% = 0 Then
                   If Mid$(RESTO$, JJ%, 1) = ")" Then
                      RESTO$ = Left$(RESTO$, JJ%) + ")" + Mid$(RESTO$, JJ% + 1)
                      GoTo 25
                   End If
                End If
              Next JJ%
              RESTO$ = RESTO$ + ")"
              GoTo 25
           End If
           GoTo 48
25         YY$ = Left$(XX$, PPXX% - 1) + " FLEXCONN.EXECUTE (FILTSQL$ " + RESTO$
'   MsgBox XX$ + vbCrLf + YY$
           XX$ = YY$
           HUBOREP% = 1
        End If
        '
        '
        '
48      PPXX% = InStr(UCase$(XX$), ".OPEN ")
        If PPXX% > 0 Then
           PPXX1% = InStr(UCase$(XX$), ".OPEN FILTSQL$")
           If PPXX1% = PPXX% Then GoTo 58
           RESTO$ = TRIM$(Mid$(XX$, PPXX% + 6))
           PRICA$ = Left$(RESTO$, 1)
           If PRICA$ <> "(" Then
              For JJ% = 2 To Len(RESTO$)
                If Mid$(RESTO$, JJ%, 1) = "," Or Mid$(RESTO$, JJ%, 1) = "'" Then
                   RESTO$ = "(" + Left$(RESTO$, JJ% - 1) + ")" + Mid$(RESTO$, JJ%)
                   GoTo 55
                End If
              Next JJ%
              RESTO$ = "(" + RESTO$ + ")"
              GoTo 55
            ElseIf PRICA$ = "(" Then
              GoTo 55
           End If
           GoTo 58
55         YY$ = Left$(XX$, PPXX% - 1) + ".OPEN FILTSQL$ " + RESTO$
'   MsgBox XX$ + vbCrLf + YY$
           XX$ = YY$
           HUBOREP% = 1
        End If
        '
        '
        '
58      PPXX% = InStr(UCase$(XX$), ".ADDNEW")
        If PPXX% > 0 Then
           If InStr(XX$, "' YA CONVERTIDO") > 0 Then GoTo 80
           '
           RSNA$ = Left$(XX$, PPXX% - 1)
           '
           YY$ = XX$ + "   ' YA CONVERTIDO" + vbCrLf + Space$(Len(RSNA$)) + "ON ERROR RESUME NEXT"
           YY$ = YY$ + vbCrLf + RSNA$ + "!CodiEmpr = Codiemp%"
           YY$ = YY$ + vbCrLf + Space$(Len(RSNA$)) + "ON ERROR GOTO 0"
'  MsgBox XX$ + vbCrLf + YY$
           XX$ = YY$
'MsgBox XX$
           HUBOREP% = 1
        End If
        '
        '
        '
80      Print #N2%, XX$
89    Loop
      Close #N1%
      Close #N2%
      '
      If HUBOREP% = 1 Then
        ZZ$ = LOADFILE$("C:\ASALID.WKF")
        Call SAVEFILE(ACOMPLE$, ZZ$)
      End If
95 Next U&
   '
   Screen.MousePointer = 1
   Call COMUNI("Conversion Completa")
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   '
   Command9.Enabled = False
   '
   OPX% = COMALTER%("Elija el Tipo de Busqueda:\\String en Programas Fuentes\Archivos en Red LAN")
   If OPX% = 1 Then
        GoTo 10
      ElseIf OPX% = 2 Then
        FORBUSKE.Show 1
        GoTo 99
      Else
        GoTo 99
   End If
   '
10 ABUSKE$ = InputBox$("Texto a Buscar", "Busqueda por Texto")
   If ABUSKE$ = "" Then
      Command9.Enabled = True
      Exit Sub
   End If
   '
   ESTEQ$ = "C:"
   BANDERA% = 0
   Screen.MousePointer = 11
   ESTARUT$ = "C:\ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB\FLEXOFT\"
   '
   ' EXPLORA LOS DIRECTORIOS LOCALES
20 List1.Clear: List2.Clear
   List1.AddItem ESTARUT$
   ILIST1& = 0
   '
   Do While ILIST1& < List1.ListCount
      RUTAX$ = List1.List(ILIST1&)
      On Error Resume Next
      NOMBREX$ = Dir(RUTAX$, 16) ' Recupera la primera entrada.
      If Err Then
         On Error GoTo 0
         GoTo 6
      End If
      On Error GoTo 0
      Do While NOMBREX$ <> "" ' Inicia el bucle.
        ' Ignora el directorio actual y el que lo abarca.
        If NOMBREX$ <> "." And NOMBREX$ <> ".." Then
           ' Utiliza comparación a nivel de bits para asegurarse de que MiNombre es un directorio.
           If (GetAttr(RUTAX$ & NOMBREX$) And vbDirectory) = vbDirectory Then
              If UCase$(RUTAX$ + NOMBREX$) <> ULODAY$ Then ' EXCLUYE DIRECTORIO TRABAJO
                 List1.AddItem RUTAX$ + NOMBREX$ + "\"
                 List1.ListIndex = List1.ListCount - 1
              End If
           End If ' solamente si representa un directorio.
        End If
        NOMBREX$ = Dir ' Obtiene siguiente entrada.
        Openforms = DoEvents
      Loop
6     ILIST1& = ILIST1& + 1
   Loop
   '
   ' CARGA EN LIST2 LOS ARCHIVOS LOCALES
   For i& = 1 To List1.ListCount
      File1.Path = List1.List(i& - 1)
      estaruta$ = File1.Path + "\"
      If File1.ListCount > 0 Then BANDERA% = 1
      For j& = 1 To File1.ListCount
         FINAME$ = TRIM$(UCase$(File1.List(j& - 1)))
           If InStr(FINAME$, ".FRX") > 0 Then GoTo 17
           If InStr(FINAME$, ".VBW") > 0 Then GoTo 17
           If InStr(FINAME$, ".PDF") > 0 Then GoTo 17
         TTXX$ = AJUSTI$(FINAME$, 64)
         TTXX$ = TTXX$ + estaruta$ + FINAME$
         List2.AddItem TTXX$
         Openforms = DoEvents
17    Next j&
19 Next i&
   '
   If BANDERA% = 0 Then
      If ESTARUT$ = "C:\ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB\FLEXOFT\" Then
         ESTARUT$ = "C:\PROGRAM FILES\DEVSTUDIO\VB\FLEXOFT\"
         GoTo 20
      End If
   End If
   '
   For U& = 1 To List2.ListCount
      ACOMPLE$ = TRIM$(Mid$(List2.List(U& - 1), 65))
      TTXX$ = LOADFILE$(ACOMPLE$)
      If InStr(UCase$(TTXX$), UCase$(ABUSKE$)) > 0 Then
         MsgBox ACOMPLE$
      End If
   Next U&
   '
   Screen.MousePointer = 1
   MsgBox "FIN"
   '
99 Command9.Enabled = True
   '
End Sub

Private Sub Form_Load()
   Combo1.Clear
   'Combo1.AddItem "\\NOTEBOOK\C_NOTEBOOK"
   'Combo1.AddItem "\\Notebook-bauer\C"    ' \Program Files\DevStudio\VB\FLEXOFT
   'Combo1.AddItem "\\Soporte\D"
   'Combo1.AddItem "\\Desarrollo\Datos (D)"
   'Combo1.AddItem "\\Ventas\DATOS (D)"
   'Combo1.AddItem "\\Ernesto\sistema (C)"
   Combo1.AddItem "z:\"
End Sub

Sub COMPARFUEN(ARLOC$, AREMO$)
    '
    If EXISTE%(ARLOC$) < 1 Then 'Exit Sub
      If EXISTE%(AREMO$) < 1 Then
        Exit Sub
      End If
    End If
    '
    Screen.MousePointer = 11
    FECHOTRO# = 0: FECHESTE# = 0
    On Error Resume Next
    FECHESTE# = FileDateTime(ARLOC$)
    Label6 = Format$(FECHESTE#, "DD/MM/yy  hh:mm:ss")
    FECHOTRO# = FileDateTime(AREMO$)
    Label7 = Format$(FECHOTRO#, "DD/MM/yy  hh:mm:ss")
    On Error GoTo 0
    '
    Label8.Visible = False
    Label9.Visible = False
    If FECHOTRO# > FECHESTE# Then
        Label9.Visible = True
      Else
        Label8.Visible = True
    End If
    '
    Label10 = ARLOC$: N1% = 0
    Label11 = AREMO$: N2% = 0
    If EXISTE%(ARLOC$) > 0 Then N1% = FILEOPEN%(ARLOC$, 1, 128)
    If EXISTE%(AREMO$) > 0 Then N2% = FILEOPEN%(AREMO$, 1, 128)
    PRIDIF& = -1
    CONTA& = 0
    RUTIAC$ = "": RUTIANT$ = ""
    '
    List4.Clear
    List5.Clear
    Text1 = ""
    Text2 = ""
    '
    Label112 = APARFUEN2.Combo1.TEXT + ":"
    Label6.Left = Label111.Left + Label111.Width + 60
    Label7.Left = Label112.Left + Label112.Width + 60
    List6.Clear
    '
    Dim RUTNA1$(512), BLOQ1$(512), BLOQ1R$(512), RUTNA2$(512), BLOQ2$(512), BLOQ2R$(512)
    Dim COMBLO2&(512), FINBLO2&(512)
    CARUT1% = 0
    If N1% < 1 Then GoTo 10
    Do While Not EOF(N1%)
      Line Input #N1%, TX1$
      TTXX$ = TRIM$(UCase$(TX1$))
      If Left$(TTXX$, 4) = "SUB " Or Left$(TTXX$, 12) = "PRIVATE SUB " Or Left$(TTXX$, 9) = "FUNCTION " Or Left$(TTXX$, 14) = "BEGIN VB.FORM " Or Left$(TTXX$, 8) = "TYPE=EXE" Then
        CARUT1% = CARUT1% + 1
        TX11$ = TX1$
        If Left$(UCase$(TX1$), 7) = "PRIVATE" Then TX11$ = Mid$(TX1$, 9)
        RUTNA1$(CARUT1%) = LTrim$(TX11$)
      End If
      BLOQ1$(CARUT1%) = BLOQ1$(CARUT1%) + TX1$ + CRLF$
      LX5$ = REPLACAR$(TRIM$(UCase$(TX1$)), " ", "")
      BLOQ1R$(CARUT1%) = BLOQ1R$(CARUT1%) + LX5$ + CRLF$
    Loop
    Close #N1%: N1% = 0
    '
10  CARUT2% = 0: POSPUN2& = 0
    If N2% < 1 Then GoTo 15
    Do While Not EOF(N2%)
      Line Input #N2%, TX2$
      TTXX$ = TRIM$(UCase$(TX2$))
      If Left$(TTXX$, 4) = "SUB " Or Left$(TTXX$, 12) = "PRIVATE SUB " Or Left$(TTXX$, 9) = "FUNCTION " Or Left$(TTXX$, 14) = "BEGIN VB.FORM " Or Left$(TTXX$, 8) = "TYPE=EXE" Then
        FINBLO2&(CARUT2%) = POSPUN2&
        CARUT2% = CARUT2% + 1
        COMBLO2(CARUT2%) = POSPUN2& + 1
        TX22$ = TX2$
        POSPUN2& = POSPUN2& + Len(TX2$) + 2
        If Left$(UCase$(TX2$), 7) = "PRIVATE" Then TX22$ = Mid$(TX2$, 9)
        RUTNA2$(CARUT2%) = LTrim$(TX22$)
      End If
      BLOQ2$(CARUT2%) = BLOQ2$(CARUT2%) + TX2$ + CRLF$
      LX5$ = REPLACAR$(TRIM$(UCase$(TX2$)), " ", "")
      BLOQ2R$(CARUT2%) = BLOQ2R$(CARUT2%) + LX5$ + CRLF$
    Loop
    Close #N2%: N2% = 0
    FINBLO2&(CARUT2%) = POSPUN2&
    '
15  For KI1% = 1 To CARUT1%
      For KI2% = 1 To CARUT2%
        If UCase$(BLOQ2R$(KI2%)) = UCase$(BLOQ1R$(KI1%)) Then
          BLOQ1$(KI1%) = ""
          BLOQ2$(KI2%) = ""
          GoTo 19
        End If
      Next KI2%
19  Next KI1%
    '
    For KI1% = 1 To CARUT1%
      If BLOQ1$(KI1%) <> "" Then
        List6.AddItem RUTNA1$(KI1%)
        TTXX$ = BLOQ1$(KI1%)
        Do While TTXX$ <> ""
          PPXX% = InStr(TTXX$, CRLF$)
          List4.AddItem Left$(TTXX$, PPXX% - 1)
          TTXX$ = Mid$(TTXX$, PPXX% + 2)
        Loop
        '
        For KI2% = 1 To CARUT2%
          If UCase$(RUTNA2$(KI2%)) = UCase$(RUTNA1$(KI1%)) Then
            TTXX$ = BLOQ2$(KI2%)
            Do While TTXX$ <> ""
              PPXX% = InStr(TTXX$, CRLF$)
              List5.AddItem Left$(TTXX$, PPXX% - 1)
              TTXX$ = Mid$(TTXX$, PPXX% + 2)
            Loop
            BLOQ2$(KI2%) = ""
            GoTo 29
          End If
        Next KI2%
        '
      End If
29  Next KI1%
    '
    For KI2% = 1 To CARUT2%
      If BLOQ2$(KI2%) <> "" Then
        List6.AddItem RUTNA2$(KI2%)
        TTXX$ = BLOQ2$(KI2%)
        Do While TTXX$ <> ""
          PPXX% = InStr(TTXX$, CRLF$)
          List5.AddItem Left$(TTXX$, PPXX% - 1)
          TTXX$ = Mid$(TTXX$, PPXX% + 2)
        Loop
      End If
    Next KI2%
    '
    If List4.ListCount < 1 Then
       If List5.ListCount < 1 Then
          Call COMPATOTA(ARLOC$, AREMO$)
       End If
    End If
    '
    Screen.MousePointer = 1
    DoEvents
    '
End Sub

Sub COMPATOTA(ARLOC$, AREMO$)
    '
    MsgBox "Comparacion Total"
    '
    If EXISTE%(ARLOC$) > 0 Then N1% = FILEOPEN%(ARLOC$, 1, 128)
    If EXISTE%(AREMO$) > 0 Then N2% = FILEOPEN%(AREMO$, 1, 128)
    PRIDIF& = -1
    CONTA& = 0
    RUTIAC$ = "": RUTIANT$ = ""
    '
    List4.Clear
    List5.Clear
    Text1 = ""
    Text2 = ""
    '
    Label112 = APARFUEN2.Combo1.TEXT + ":"
    Label6.Left = Label111.Left + Label111.Width + 60
    Label7.Left = Label112.Left + Label112.Width + 60
    List6.Clear
    '
    Dim RUTNA1$(512), BLOQ1$(512), RUTNA2$(512), BLOQ2$(512)
    CARUT1% = 1
    If N1% < 1 Then GoTo 10
    Do While Not EOF(N1%)
      Line Input #N1%, TX1$
      List4.AddItem TX1$
      If List4.ListCount = 1 Then
         List6.AddItem TX1$
      End If
    Loop
    Close #N1%: N1% = 0
    '
10  CARUT2% = 1
    If N2% < 1 Then GoTo 15
    Do While Not EOF(N2%)
      Line Input #N2%, TX2$
      List5.AddItem TX2$
      If List5.ListCount = 1 Then
         If List6.ListCount > 0 Then
            If List6.List(0) = TX2$ Then
               GoTo 12
            End If
         End If
         List6.AddItem TX2$
      End If
12  Loop
    Close #N2%: N2% = 0
    '
15  Screen.MousePointer = 1
    DoEvents
    '
End Sub
'
Private Sub Label111_DblClick()
   ALOCA$ = TRIM$(Mid$(List2.TEXT, 65))
   AAA = Shell("C:\Archivos de programa\DevStudio\VB\VB5 " + ALOCA$)
End Sub

Private Sub List2_DblClick()
    VUELTA% = 0
    ALO$ = UCase$(TRIM$(Left$(List2.TEXT, 64)))
    ALOCA$ = TRIM$(Mid$(List2.TEXT, 65))
    PPXX% = InStr(UCase$(ALOCA$), "DEVSTUDIO")
      If PPXX% < 1 Then PPXX% = 1
    ESTEAR1$ = Mid$(ALOCA$, PPXX%)
    '
10  For KKI% = 1 To List3.ListCount
       Debug.Print TRIM$(UCase$(Left$(List3.List(KKI% - 1), 64)))
       If TRIM$(UCase$(Left$(List3.List(KKI% - 1), 64))) = ALO$ Then
          AREMO$ = TRIM$(Mid$(List3.List(KKI% - 1), 65))
          PPXX0% = InStr(UCase$(AREMO$), "FLEXOFT")
          PPXX1% = InStr(UCase$(ALOCA$), "FLEXOFT")
            If PPXX0% < 1 Then PPXX0% = 1
          OTROAR1$ = Mid$(AREMO$, PPXX0%)
          ESTEAR1$ = Mid$(ALOCA$, PPXX1%)
          If OTROAR1$ = ESTEAR1$ Then
            List3.ListIndex = KKI% - 1
            AREMO$ = TRIM$(Mid$(List3.List(KKI% - 1), 65))
            GoTo 80
          End If
       End If
    Next KKI%
    AREMO$ = TRIM$(Combo1.TEXT) + Mid$(ALOCA$, 4)
    If VUELTA% = 0 Then
       VUELTA% = 1
       TTXX$ = Left$(List2.TEXT, 64) + AREMO$
       List3.AddItem TTXX$
       GoTo 10
    End If
    '
80  Call COMPARFUEN(ALOCA$, AREMO$)
    '
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
      If List2.ListIndex >= 0 Then
        ILIST1& = List2.ListIndex
        ALOCA$ = TRIM$(Mid$(List2.TEXT, 65))
        If MsgBox("Confirma Borrar " + ALOCA$, 259, "Borrado de Archivo") = vbYes Then
          Kill ALOCA$
          List2.RemoveItem ILIST1&
        End If
      End If
    End If
End Sub

Private Sub List3_DblClick()
    '
    List2.ListIndex = (-1)
    ARE$ = UCase$(TRIM$(Left$(List3.TEXT, 64)))
    AREMO$ = TRIM$(Mid$(List3.TEXT, 65))
    PPXX% = InStr(UCase$(AREMO$), "FLEXOFT")
      If PPXX% < 1 Then PPXX% = 1
    OTROAR1$ = Mid$(AREMO$, PPXX%)
    '
    For KKI% = 1 To List2.ListCount
        Debug.Print TRIM$(UCase$(Left$(List2.List(KKI% - 1), 64)))
       If TRIM$(UCase$(Left$(List2.List(KKI% - 1), 64))) = ARE$ Then
          ALOCA$ = TRIM$(Mid$(List2.List(KKI% - 1), 65))
          PPXX0% = InStr(UCase$(ALOCA$), "FLEXOFT")
          PPXX1% = InStr(UCase$(OTROAR1$), "FLEXOFT")
            If PPXX0% < 1 Then PPXX0% = 1
          ESTEAR1$ = Mid$(ALOCA$, PPXX0%)
          OTROAR1$ = Mid$(OTROAR1$, PPXX1%)
          If OTROAR1$ = ESTEAR1$ Then
             List2.ListIndex = KKI% - 1
             ALOCA$ = TRIM$(Mid$(List2.List(KKI% - 1), 65))
             GoTo 80
          End If
       End If
    Next KKI%
    ALOCA$ = ""
    '
80  Call COMPARFUEN(ALOCA$, AREMO$)
    '
End Sub

Private Sub List3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
      If List3.ListIndex >= 0 Then
        ILIST1& = List3.ListIndex
        AREMO$ = TRIM$(Mid$(List3.TEXT, 65))
        If MsgBox("Confirma Borrar " + AREMO$, 259, "Borrado de Archivo") = vbYes Then
          On Error Resume Next
          Kill AREMO$
          On Error GoTo 0
          List3.RemoveItem ILIST1&
        End If
      End If
    End If
End Sub

Private Sub List4_Click()
   Text1 = List4.TEXT
End Sub

Private Sub List4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      TTXX$ = ""
      For II& = 1 To List4.ListCount
         If List4.Selected(II& - 1) = True Then
            TTXX$ = TTXX$ + List4.List(II& - 1) + vbCrLf
         End If
      Next II&
      '
      On Error Resume Next
      Clipboard.Clear
      Clipboard.SetText TTXX$
      If Err Then
         On Error GoTo 0
         MsgBox "Imposible Abrir el Portapapeles - Reintente"
      End If
      On Error GoTo 0
      '
   End If
End Sub

Private Sub List5_Click()
   Text2 = List5.TEXT
End Sub

Private Sub List5_DblClick()
   EDIFUEN.Top = Top + List5.Top - 150
   EDIFUEN.Left = Left + List5.Left - 100
   AREMO$ = TRIM$(Mid$(List3.TEXT, 65))
   EDIFUEN.Text1 = LOADFILE$(AREMO$)
   AA1$ = LOADFILE$(AREMO$)
   X = Len(AA1$)
   Y = Len(EDIFUEN.Text1)
   EDIFUEN.Show 1
End Sub

Private Sub List5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      TTXX$ = ""
      For II& = 1 To List5.ListCount
         If List5.Selected(II& - 1) = True Then
            TTXX$ = TTXX$ + List5.List(II& - 1) + vbCrLf
         End If
      Next II&
      '
      On Error Resume Next
      Clipboard.Clear
      Clipboard.SetText TTXX$
      If Err Then
         On Error GoTo 0
         MsgBox "Imposible Abrir el Portapapeles - Reintente"
      End If
      On Error GoTo 0
      '
   End If
End Sub

Private Sub List6_Click()
   For KI1% = 1 To List4.ListCount
     List4.Selected(KI1% - 1) = False
   Next KI1%
   '
   For KI2% = 1 To List5.ListCount
     List5.Selected(KI2% - 1) = False
   Next KI2%
   '
   Text1 = "": Text2 = ""
   TBUS$ = List6.TEXT
   ini1% = 0: INI2% = 0
   For KI1% = 1 To List4.ListCount
     If LTrim$(List4.List(KI1% - 1)) = TBUS$ Or UCase$(LTrim$(List4.List(KI1% - 1))) = "PRIVATE " + UCase$(TBUS$) Then
       ini1% = KI1%
       GoTo 10
     End If
   Next KI1%
   '
10 For KI2% = 1 To List5.ListCount
     If LTrim$(List5.List(KI2% - 1)) = TBUS$ Or UCase$(LTrim$(List5.List(KI2% - 1))) = "PRIVATE " + UCase$(TBUS$) Then
       INI2% = KI2%
       GoTo 20
     End If
   Next KI2%
   '
20 KI1% = ini1%
   KI2% = INI2%
   '
   If ini1% > 0 And INI2% <= 0 Then
     List4.Selected(ini1% - 1) = True
     Text1 = List4.List(ini1% - 1)
     If List4.TopIndex < ini1% - 20 Then List4.TopIndex = ini1% - 20
     Exit Sub
   End If
   '
   If ini1% <= 0 And INI2% > 0 Then
     List5.Selected(INI2% - 1) = True
     If List5.TopIndex < INI2% - 20 Then List5.TopIndex = INI2% - 20
     Text2 = List5.List(INI2% - 1)
     Exit Sub
   End If
   '
' 25 If TRIM$(UCase$(List4.List(KI1% - 1))) <> TRIM$(UCase$(List5.List(KI2% - 1))) Then
25 LX4$ = REPLACAR$(TRIM$(UCase$(List4.List(KI1% - 1))), " ", "")
   LX5$ = REPLACAR$(TRIM$(UCase$(List5.List(KI2% - 1))), " ", "")
   If InStr(LX4$, "RETAINED") > 0 Then KI1% = KI1% + 1: GoTo 25
   If InStr(LX5$, "RETAINED") > 0 Then KI2% = KI2% + 1: GoTo 25
   If LX4$ <> LX5$ Then
      ABUXI$ = "OLEAUTOMATION": If InStr(LX4$, ABUXI$) > 0 And InStr(LX5$, ABUXI$) > 0 Then GoTo 9
      ABUXI$ = "MICROSOFTDAO3.6OBJECTLIBRARY": If InStr(LX4$, ABUXI$) > 0 And InStr(LX5$, ABUXI$) > 0 Then GoTo 9
      ABUXI$ = "MICROSOFTACTIVEXDATAOBJECTS2.6LIBRARY": If InStr(LX4$, ABUXI$) > 0 And InStr(LX5$, ABUXI$) > 0 Then GoTo 9
      List4.Selected(KI1% - 1) = True
      If List4.TopIndex < KI1% - 20 Then List4.TopIndex = KI1% - 20
      Text1 = List4.List(KI1% - 1)
      List5.Selected(KI2% - 1) = True
      If List5.TopIndex < KI2% - 20 Then List5.TopIndex = KI2% - 20
      Text2 = List5.List(KI2% - 1)
      Exit Sub
   End If
9  KI1% = KI1% + 1
   KI2% = KI2% + 1
   If KI1% <= List4.ListCount Then
     If KI2% <= List5.ListCount Then
       GoTo 25
     End If
   End If
   '
   List4.Selected(ini1% - 1) = True
   List4.TopIndex = ini1% - 1
   Text1 = List4.List(ini1% - 1)
   List5.Selected(INI2% - 1) = True
   List5.TopIndex = INI2% - 1
   Text2 = List5.List(INI2% - 1)
   '
   MsgBox "Diferencia no Encontada"
   If List6.ListCount <= 1 Then
      Call Command5_Click
   End If
   '
End Sub

Sub ABRE_CORREC()
   Call CREACAMPO("ADMIFLEX", "CORRPROG", "TecRespo", 10, 16, 1)
   Call CREACAMPO("ADMIFLEX", "CORRPROG", "Modulo", 10, 64)
   Call CREACAMPO("ADMIFLEX", "CORRPROG", "Rutina", 10, 64)
   Call CREACAMPO("ADMIFLEX", "CORRPROG", "Deta_Corr", 12, 0)
End Sub

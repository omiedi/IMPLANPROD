VERSION 5.00
Begin VB.Form SHOWPDF 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Mensajes y Consultas Telefónicas FLEXOFT"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pdf1 
      Height          =   7500
      Left            =   0
      ScaleHeight     =   7440
      ScaleWidth      =   10965
      TabIndex        =   13
      Top             =   0
      Width           =   11030
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   620
      Left            =   240
      ScaleHeight     =   615
      ScaleWidth      =   10695
      TabIndex        =   14
      Top             =   7560
      Width           =   10695
      Begin VB.CommandButton Command6 
         Caption         =   "ok"
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
         Left            =   10290
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   50
         Width           =   375
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
         Left            =   8040
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   300
         Width           =   2655
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   960
         TabIndex        =   16
         Top             =   0
         Width           =   5175
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Pre-Instalar Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   6240
         TabIndex        =   15
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Enviar a:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   18
         Top             =   75
         Width           =   2085
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivos Adjuntos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -120
         TabIndex        =   17
         Top             =   30
         Width           =   1005
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   9000
      Left            =   11030
      ScaleHeight     =   8940
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   0
      Width           =   1200
      Begin VB.CommandButton Command10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   140
         Picture         =   "SHOWPDF.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Enviar a Otro Usuario de la LAN"
         Top             =   6000
         Width           =   600
      End
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
         Height          =   600
         Left            =   140
         Picture         =   "SHOWPDF.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Guardar como Archivo"
         Top             =   5400
         Width           =   600
      End
      Begin VB.CommandButton Command8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   140
         Picture         =   "SHOWPDF.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Imprimir el Documento"
         Top             =   4800
         Width           =   600
      End
      Begin VB.CommandButton Command7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   140
         Picture         =   "SHOWPDF.frx":0C7E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Enviar Confirmacion de  Lectura"
         Top             =   4200
         Width           =   600
      End
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
         Height          =   600
         Left            =   140
         Picture         =   "SHOWPDF.frx":0F88
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Ir al Final del Documento"
         Top             =   2800
         Width           =   600
      End
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
         Height          =   600
         Left            =   140
         Picture         =   "SHOWPDF.frx":1292
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Página Siguiente"
         Top             =   2200
         Width           =   600
      End
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
         Height          =   600
         Left            =   140
         Picture         =   "SHOWPDF.frx":159C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Página Anterior"
         Top             =   1600
         Width           =   600
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
         Height          =   600
         Left            =   140
         Picture         =   "SHOWPDF.frx":18A6
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Ir al Comienzo del Documento"
         Top             =   1000
         Width           =   600
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   90
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   2
         Top             =   7350
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   12000
         End
      End
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
         Height          =   550
         Left            =   140
         Picture         =   "SHOWPDF.frx":1BB0
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Salir"
         Top             =   105
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "SHOWPDF.frx":1CFA
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   7710
         Width           =   1515
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   975
      Left            =   6000
      TabIndex        =   22
      Top             =   4440
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   615
      Left            =   6120
      TabIndex        =   21
      Top             =   3120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   20000
      Y1              =   7510
      Y2              =   7510
   End
   Begin VB.Label Label1 
      Height          =   330
      Left            =   105
      TabIndex        =   12
      Top             =   210
      Visible         =   0   'False
      Width           =   6315
   End
End
Attribute VB_Name = "SHOWPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'    Pdf1.gotoFirstPage
End Sub

Sub Command10_Click()
   Command10.Enabled = False
   MsgBox "No Disponible por Ahora", vbCritical
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
    WEBROWSR.Timer2.Enabled = False
    For KKK& = 1 To List1.ListCount
       ARNAME$ = UCase$(TRIM$(Left$(List1.List(KKK& - 1), 32)))
       ARFULLNAME$ = AJUSTI$(List1.List(KKK& - 1), 48)
       If InStr(ARNAME$, ".EXE") > 0 Then
          Call WEBROWSR.INSTALACON(ARFULLNAME$)
       End If
    Next KKK&
    List1.Clear
    WEBROWSR.Timer2.Enabled = True
End Sub

Private Sub Command2_Click()
'    Pdf1.gotoPreviousPage
    
End Sub

Private Sub Command3_Click()
'    Pdf1.gotoNextPage
End Sub

Private Sub Command4_Click()
'   Pdf1.gotoLastPage
End Sub

Private Sub Command5_Click()
   Unload Me
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   If TRIM$(Combo1.TEXT) <> "" Then
      ORIID$ = TRIM$(UCase$(USERID$))
      DESTID$ = TRIM$(Mid$(Combo1.TEXT, 65, 12))
      If DESTID$ <> "" Then
        DESTID$ = AJUSTI$(DESTID$, 12)
        IDME$ = Label5
        '
        Label8 = CARPEDOWN$
        '
        LISNOVE1$ = ""
        LISNOVE$ = LOADFILE$(CARPEDOWN$ + "NOPENPRO.TXT")
        For U& = 1 To Len(LISNOVE$) Step 128
           ARCHAI$ = Mid$(LISNOVE$, U&, 128)
           If TRIM$(Mid$(ARCHAI$, 103, 4)) = IDME$ Then
              ORIUS$ = TRIM$(UCase$(Mid$(ARCHAI$, 75, 12)))
              CONCO$ = TRIM$(UCase$(Mid$(ARCHAI$, 107, 12)))
              If ORIUS$ = ORIID$ Or CONCO$ = ORIID$ Or ORIUS$ = "" Then
                 Mid$(ARCHAI$, 75, 12) = DESTID$
                 Mid$(ARCHAI$, 107, 12) = Space$(12)
                 LISNOVE1$ = LISNOVE1$ + ARCHAI$
              End If
           End If
        Next U&
        '
        While Len(LISNOVE$) Mod 128 <> 0: LISNOVE$ = LISNOVE$ + " ": Wend
        LISNOVE$ = LISNOVE$ + LISNOVE1$
        While Len(LISNOVE$) Mod 128 <> 0: LISNOVE$ = LISNOVE$ + " ": Wend
        Call SAVEFILE(CARPEDOWN$ + "NOPENPRO.TXT", LISNOVE$)
        '
      End If
   End If
   '
   Command6.Enabled = True
End Sub

Sub Command7_Click()
   Command7.Enabled = False
   Call WEBROWSR.CONFIRMALEC(Label1)
   Command7.Enabled = True
End Sub

Sub Command8_Click()
'   Pdf1.printWithDialog
End Sub

Sub Command9_Click()
   PROCFILE.CommonDialog1.CancelError = True
   PROCFILE.CommonDialog1.DialogTitle = "Guardar como Archivo PDF"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   PROCFILE.CommonDialog1.FILTER = "Documento PDF(*.PDF|*.PDF|"
   PROCFILE.CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   PROCFILE.CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   AREXPO$ = UCase$(PROCFILE.CommonDialog1.FileName)
   If TRIM$(AREXPO$) <> "" Then
      ARORI$ = CARPEDOWN$ + TRIM$(Left$(Label1, 32))
      TTXX$ = LOADFILE$(ARORI$)
      If TTXX$ <> "" Then
          Call SAVEFILE(AREXPO$, TTXX$)
      End If
   End If
   '
99 End Sub

Private Sub Form_Load()
   '
   Combo1.Clear
   LFI$ = LOADFILE$("C:\FLEXOFT\FLEXRAM.DRV")
   USERNU% = CVI(Mid$(LFI$, 9, 2)) Mod 100
   If USERNU% < 1 Then USERNU% = 1
   '
   VUSER$ = LOADFILE$("USER.DAT")
   For KKI% = 1 To Len(LFI$) / 64
     If KKI% <> USERNU% Then
       F0$ = Mid$(VUSER$, 64 * KKI% - 63, 64)
       USID$ = Mid$(F0$, 31, 12)
          If Left$(USID$, 1) = "@" Then USID$ = DESENCRI$(Mid$(F0$, 32, 11)) + " "
       USNA$ = Left$(F0$, 30)
          If Left$(USNA$, 1) = "@" Then USNA$ = DESENCRI$(Mid$(F0$, 2, 29)) + " "
       If TRIM$(USID$) <> "" Then
         If TRIM$(USNA$) <> "" Then
           Combo1.AddItem AJUSTI$(USNA$, 64) + USID$
         End If
       End If
     End If
   Next KKI%
   '
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   '
   ARFULLNA$ = Label1
   ARNA$ = TRIM$(Left$(ARFULLNA$, 32))
   FIDATI$ = TRIM$(Mid$(ARFULLNA$, 33, 16))
   '
   SQLX$ = "SELECT * FROM CONTRO_DOWN"
   SQLX$ = SQLX$ + " WHERE File_Name = '" & ARNA$ & "' "
   SQLX$ = SQLX$ + " AND File_DateTime = '" & FIDATI$ & "'"
   SQLX$ = SQLX$ + " AND (FeHor_Lect = '' OR ISNULL(FeHor_Lect) = TRUE) "
   SQLX$ = SQLX$ + " AND User_Dest = '" & USERID$ & "'"
   '
   Call CREABACONLEC
   Set MIRS = BADACON.OpenRecordset(SQLX$, dbOpenDynaset)
   On Error Resume Next
   MIRS.MoveFirst
   If Err Then
      MIRS.Close
      BADACON.Close
      On Error GoTo 0
      Exit Sub                 ' no hay nada que bajar
   End If
   MIRS.Close
   BADACON.Close
   On Error GoTo 0
   '
   If MsgBox("¿ Enviar Confirmación de Lectura a FLEXOFT ?", vbYesNo + vbDefaultButton2 + vbQuestion, "Confirmación de Lectura") = 6 Then
      Call WEBROWSR.CONFIRMALEC(ARFULLNA$)
   End If
   '
End Sub

Private Sub Form_Resize()
'   Picture1.Left = Width - 1000
'   Picture1.Height = Height - 350
'   Pdf1.Width = Width - 1000
'   Pdf1.Height = Height - (860 + Height - ScaleHeight)
'   Line1.Y1 = Pdf1.Height + 10
'   Line1.Y2 = Pdf1.Height + 10
'   Picture2.Top = Line1.Y1 + 170
'   Picture2.Left = (Picture1.Left - Picture2.Width) / 2
'   Command11.Top = List1.Top + List1.Height - Command11.Height
End Sub

Private Sub List1_DblClick()
    WEBROWSR.Timer2.Enabled = False
    ARNAME$ = UCase$(TRIM$(Left$(List1.TEXT, 32)))
    ARFULLNAME$ = AJUSTI$(List1.TEXT, 48)
    If InStr(ARNAME$, ".EXE") > 0 Then
          Call WEBROWSR.INSTALACON(ARFULLNAME$)
          ' ELIMINA DE LA LISTA DE SELECCION
          ARFU$ = AJUSTI$(ARFULLNAME$, 48)
65        For KKIL& = 1 To WEBROWSR.List3.ListCount
            If Left$(List1.List(KKIL& - 1), 48) = ARFU$ Then
              List1.RemoveItem KKIL& - 1
              List1.Refresh
              GoTo 65
            End If
          Next KKIL&
       Else
          MsgBox "Imposible Instalar '" & ARNAME$ & "'."
    End If
    WEBROWSR.Timer2.Enabled = True
End Sub

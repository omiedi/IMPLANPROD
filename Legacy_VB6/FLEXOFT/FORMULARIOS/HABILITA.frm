VERSION 5.00
Begin VB.Form HABILITA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Habilitar Ejecutables"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4095
      TabIndex        =   8
      Top             =   2590
      Width           =   855
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   1890
      Pattern         =   "*.EXE"
      TabIndex        =   6
      Top             =   1575
      Width           =   2010
   End
   Begin VB.TextBox Text3 
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
      Left            =   1890
      TabIndex        =   5
      Text            =   "C:\FLEXOWIN"
      Top             =   1155
      Width           =   3060
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
      Height          =   300
      Left            =   1890
      TabIndex        =   3
      Text            =   "C:\FLEXOWIN "
      Top             =   735
      Width           =   3060
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
      Height          =   300
      Left            =   1890
      TabIndex        =   1
      Text            =   " "
      Top             =   315
      Width           =   1485
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Módulos:"
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
      Left            =   105
      TabIndex        =   7
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Carpeta Destino:"
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
      Left            =   105
      TabIndex        =   4
      Top             =   1260
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Unidad - Carpeta:"
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
      Left            =   105
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Serie del Volumen:"
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
      Left            =   105
      TabIndex        =   0
      Top             =   420
      Width           =   1695
   End
End
Attribute VB_Name = "HABILITA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   '
   Command1.Enabled = False
   '
   KKII% = 0
   SEVOL$ = Trim$(Text1.Text)
   If SEVOL$ = "" Then GoTo 99
   '
   While Len(SEVOL$) < 13
     KKII% = KKII% + 1
     SEVOL$ = SEVOL$ + Chr$(18 + Asc(Mid$(SEVOL$, KKII%, 1)))
   Wend
   SEVOL$ = ENCRIPTA$(Left$(SEVOL$, 13))
   '
   For U& = 1 To File1.ListCount
     '
     File1.ListIndex = U& - 1
     '
     FINA$ = File1.List(U& - 1)
     LOFI$ = LOADFILE$(Text2.Text + "\" + FINA$)
     PQXX& = InStr(LOFI$, "Microsoft Vísual Basic 5.O")
     '
     If PQXX& > 0 Then
       '
       VENUX$ = VENUM$(Text2.Text + "\" + FINA$)
       PPXX% = InStr(FINA$, ".")
       If PPXX% < 1 Then PPXX% = 1 + Len(FINA$)
       EXENA$ = ENCRIPTA$(UCase$(Left$(FINA$, PPXX% - 1)))
       If Len(EXENA$) < 6 Then GoTo 19
       '
       While Len(EXENA$) < 13
         CARAVA% = (Asc(Mid$(EXENA$, 2, 1)) + Asc(Mid$(EXENA$, 4, 1))) Mod 256
         EXENA$ = Chr$(CARAVA%) + EXENA$
         CARAVA% = (Asc(Mid$(EXENA$, Len(EXENA$) - 3, 1)) + Asc(Mid$(EXENA$, Len(EXENA$) - 1, 1))) Mod 256
         EXENA$ = EXENA$ + Chr$(CARAVA%)
       Wend
       EXENA$ = Left$(EXENA$, 13)
       '
       CSTRI$ = Chr$(64)
       For KKI% = 2 To 13
         CARAVA% = (KKI% - 1) * Asc(Mid$(EXENA$, KKI%, 1)) + (KKI% + 3) * Asc(Mid$(SEVOL$, KKI%, 1))
         While CARAVA% < 48: CARAVA% = CARAVA% + 43: Wend
         While CARAVA% > 90: CARAVA% = CARAVA% - 43: Wend
         CSTRI$ = CSTRI$ + Chr$(CARAVA%)
       Next KKI%
       '
       If Mid$(LOFI$, PQXX& + 27, 13) <> CSTRI$ Then
         Mid$(LOFI$, PQXX& + 27, 13) = CSTRI$
         If Mid$(LOFI$, PQXX& + 40, 10) = ".........." Then
           Mid$(LOFI$, PQXX& + 40, 10) = VENUX$
         End If
         DESTI$ = Text3.Text + "\" + FINA$
         Call SAVEFILE(DESTI$, LOFI$)
       End If
       '
     End If
     '
19 Next U&
   '
99 Command1.Enabled = True
   '
End Sub

Private Sub Form_Load()
   File1.Path = Text2.Text
   File1.Pattern = "*.EXE"
End Sub

Function LOADFILE$(FINA$)
   '
   LFX$ = ""
   SCMA = Screen.MousePointer
   Screen.MousePointer = 11
   '
   On Error GoTo 97
   LOPUCO& = FileLen(FINA$)
   NX% = FreeFile
   Open FINA$ For Binary Access Read Write As #NX% Len = 2048
   LFX$ = String$(LOPUCO&, 0)
   Get #NX%, 1, LFX$
   GoTo 99
   '
97 Resume 98
98 LFX$ = ""
   '
99 If NX% > 0 Then Close #NX%
   On Error GoTo 0
   Screen.MousePointer = SCMA
   LOADFILE$ = LFX$
   '
End Function

Sub SAVEFILE(FINA$, LFX$)
   '
   SCMA = Screen.MousePointer
   Screen.MousePointer = 11
   '
   On Error GoTo 97
   NX% = FreeFile
   Open FINA$ For Binary Access Read Write As #NX% Len = 2048
   Put #NX%, 1, LFX$
   GoTo 99
   '
97 Resume 99
   '
99 If NX% > 0 Then Close #NX%
   On Error GoTo 0
   Screen.MousePointer = SCMA
   '
End Sub
'
Function ENCRIPTA$(DE$)
    '
    DE1$ = ""
    For I% = 1 To Len(DE$)
        XL% = Asc(Mid$(DE$, I%))
        Y% = (XL% + 23 * I% + 17) Mod 256
        DE1$ = DE1$ + Chr$(Y%)
    Next I%
    ENCRIPTA$ = DE1$
    '
End Function
'

Function DESENCRI$(DE$)
    '
    DE1$ = ""
    For I% = 1 To Len(DE$)
        XL% = Asc(Mid$(DE$, I%))
        Y% = (XL% - 23 * I% - 17)
        While Y% < 0
            Y% = Y% + 256
        Wend
        DE1$ = DE1$ + Chr$(Y%)
    Next I%
    DESENCRI$ = DE1$
    '
End Function
'

Function VENUM$(FI$)
   '
   XX$ = Trim$(FileDateTime(FI$))
   For KKI% = 1 To Len(XX$)
     ZZ$ = Mid$(XX$, KKI%, 1)
     KKJ% = Asc(ZZ$)
     If KKJ% <= 32 Then GoTo 5
     If KKJ% > 90 Then GoTo 5
     YY$ = YY$ + ZZ$
   Next KKI%
   '
5  PPXX% = InStr(YY$, "200")
   If PPXX% > 0 Then
     YY$ = Left$(YY$, PPXX% - 1) + Mid$(YY$, PPXX% + 2)
   End If
   '
10 For KKI% = 1 To Len(YY$)
     If Mid$(YY$, KKI%, 1) = "/" Then
       Mid$(YY$, KKI%, 1) = "."
       GoTo 10
     End If
   Next KKI%
   '
   VENUM$ = "V." + YY$
   '
End Function


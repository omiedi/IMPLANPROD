VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form SELECARCH 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adjuntar Archivos"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   31
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   5200
      Left            =   6720
      ScaleHeight     =   5145
      ScaleWidth      =   2295
      TabIndex        =   26
      Top             =   0
      Width           =   2355
      Begin VB.PictureBox Picture16 
         Height          =   5150
         Left            =   0
         ScaleHeight     =   5085
         ScaleWidth      =   1020
         TabIndex        =   33
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         Begin VB.Image Image3 
            Height          =   5150
            Left            =   0
            Picture         =   "SELECARCH.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1095
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   29
         Top             =   4440
         Width           =   650
         Begin VB.Frame Frame18 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   30
            Top             =   105
            Width           =   12000
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
         Height          =   660
         Left            =   120
         Picture         =   "SELECARCH.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   90
         TabIndex        =   27
         Top             =   1520
         Width           =   645
      End
      Begin VB.Image Image6 
         Height          =   390
         Left            =   -315
         Picture         =   "SELECARCH.frx":24E2
         ToolTipText     =   "Doble-Click para (Re)-Habilitar Estación de Trabajo"
         Top             =   4740
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Selección Origen de Restauración"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   240
      TabIndex        =   18
      Top             =   3720
      Width           =   6345
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
         Height          =   1500
         IntegralHeight  =   0   'False
         Left            =   240
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   360
         Width           =   5805
      End
      Begin VB.Label Label7 
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
         Left            =   360
         TabIndex        =   23
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label6 
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
         Left            =   360
         TabIndex        =   22
         Top             =   1560
         Width           =   5265
      End
      Begin VB.Label Label5 
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
         Left            =   1680
         TabIndex        =   21
         Top             =   600
         Width           =   4320
      End
      Begin VB.Label Label4 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1950
         Width           =   5775
      End
   End
   Begin VB.ListBox List4 
      Height          =   1620
      Left            =   4080
      Sorted          =   -1  'True
      TabIndex        =   25
      Top             =   4080
      Width           =   1170
   End
   Begin VB.FileListBox File2 
      Height          =   1650
      Left            =   1800
      TabIndex        =   24
      Top             =   4080
      Width           =   1695
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
      Height          =   780
      Left            =   6720
      Picture         =   "SELECARCH.frx":4404
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Aprueba y Procesa el Reemplazo"
      Top             =   5280
      Width           =   760
   End
   Begin VB.TextBox Text2 
      Height          =   435
      Left            =   1920
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   4440
      Width           =   3165
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Selección Archivo a Reemplazar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3550
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6345
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
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
         Height          =   705
         Left            =   240
         TabIndex        =   9
         Top             =   2280
         Width           =   5805
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0FFFF&
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
            Height          =   225
            Left            =   2325
            TabIndex        =   14
            Top             =   315
            Width           =   645
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0FFFF&
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
            Height          =   225
            Left            =   1620
            TabIndex        =   13
            Top             =   315
            Width           =   645
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
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
            Height          =   225
            Left            =   915
            TabIndex        =   12
            Top             =   315
            Width           =   645
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
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
            Height          =   225
            Left            =   210
            TabIndex        =   11
            Top             =   315
            Value           =   -1  'True
            Width           =   645
         End
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
            Height          =   300
            Left            =   3120
            TabIndex        =   10
            Top             =   280
            Width           =   2505
         End
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         IntegralHeight  =   0   'False
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   5805
      End
      Begin VB.ListBox List2 
         Height          =   1620
         Left            =   3960
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   1170
      End
      Begin VB.FileListBox File1 
         Height          =   1650
         Left            =   1080
         TabIndex        =   3
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   3120
         Width           =   5775
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
         Left            =   1680
         TabIndex        =   4
         Top             =   600
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
         Height          =   300
         Left            =   360
         TabIndex        =   2
         Top             =   1560
         Width           =   5265
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
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1065
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "SELECARCH.frx":470E
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "SELECARCH.frx":4788
      Top             =   1200
   End
   Begin VB.Label Label3 
      Height          =   435
      Left            =   1320
      TabIndex        =   15
      Top             =   5040
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
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
      Begin VB.Menu Aprobar 
         Caption         =   "Aprbar"
      End
   End
End
Attribute VB_Name = "SELECARCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub aprobar_Click()
    Call Command2_Click
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   '
   ASELE$ = TRIM$(Left$(Label2, 32))
   If ASELE$ = "" Then
       Call MENSERR(24, "No se ha Elegido Archivo.")
       GoTo 99
   End If
   '
   If EXISTE%(App.Path + "\" + ASELE$) < 1 Then
       Call MENSERR(24, "Archivo a Reemplazar Inexistente.")
       GoTo 99
   End If
   '
   AREPLA$ = TRIM$(Left$(Label4, 32))
   If AREPLA$ = "" Then
       Call MENSERR(24, "No se ha Elegido Archivo a Restaurar.")
       GoTo 99
   End If
   '
   If EXISTE%(App.Path + "\VIEJOS\" + AREPLA$) < 1 Then
       Call MENSERR(24, "Archivo a Restaurar Inexistente.")
       GoTo 99
   End If
   '
   TTXX$ = "¿ Confirma Reemplazar Archivo\'" + ASELE$ + " del " + TRIM$(Mid$(Label2, 33)) + "'"
   TTXX$ = TTXX$ + "\por el Siguiente Archivo\'" + AREPLA$ + " del " + TRIM$(Mid$(Label4, 33)) + "' ?"
   TTXX$ = TTXX$ + "\   \IMPORTANTE: Esta Transaccion no Podrá Deshacerse !!!"
   If COMALTER%(TTXX$ + "\\No, Conservar Actual\Si, Reemplazar") <> 2 Then GoTo 99
   '
'1.- BORRAR ULTIMA VERSION DE ..\NOVEDAD
   DIRNU$ = PARENTFOLDER$(DIREPRO$) + "\NOVEDAD\"
   On Error Resume Next
   Kill DIRNU$ + ASELE$
   On Error GoTo 0
   '
   DIRNU$ = "C:\FLEXOFT\NOVEDAD\"
   On Error Resume Next
   Kill DIRNU$ + ASELE$
   On Error GoTo 0
   '
   DIRNU$ = "F:\FLEXOFT\NOVEDAD\"
   On Error Resume Next
   Kill DIRNU$ + ASELE$
   On Error GoTo 0
   '
   DIRNU$ = Left$(App.Path, 1) + ":\FLEXOFT\NOVEDAD\"
   On Error Resume Next
   Kill DIRNU$ + ASELE$
   On Error GoTo 0
   '
'2.- VERIFICAR BORRADOS EN ..\NOVEDAD
   EXINO% = 0
   If EXISTE%(PARENTFOLDER$(DIREPRO$) + "\NOVEDAD\" + ASELE$) > 0 Then EXINO% = 1
   If EXISTE%("C:\FLEXOFT\NOVEDAD\" + ASELE$) > 0 Then EXINO% = 1
   If EXISTE%("F:\FLEXOFT\NOVEDAD\" + ASELE$) > 0 Then EXINO% = 1
   If EXISTE%(Left$(App.Path, 1) + ":\FLEXOFT\NOVEDAD\" + ASELE$) > 0 Then EXINO% = 1
   '
   If EXINO% > 0 Then
      TTXX$ = "Imposible Completar Transaccion Solicitada.\  \"
      TTXX$ = TTXX$ + "No se Ha Eliminado " + ASELE$ + " de la Carpeta ../../NOVEDAD\"
      TTXX$ = TTXX$ + "Borre el Archivo Manualmente\y Re-Intente la Transaccion."
      Call MENSERR(24, TTXX$)
      GoTo 99
   End If
   '
'3.- ELIMINAR ASELE$ DE LA CARPETA ..\VIEJOS
   CARPEVIE$ = App.Path + "\VIEJOS\"
   On Error Resume Next
   Kill CARPEVIE$ + ASELE$
   On Error GoTo 0
   '
'4.- EN CARPETA ..\VIEJOS COPIAR AREPLA$ COMO ASELE$
   On Error Resume Next
   FileCopy CARPEVIE$ + AREPLA$, CARPEVIE$ + ASELE$
   If Err Then
      On Error GoTo 0
      CAPVI$ = REPLACAR$(CARPEVIE$, "\", "/")
      Call MENSERR(24, "Imposible Copiar '" + AREPLA$ + "' Como '" + ASELE$ + "'\en Carpeta " + CAPVI$ + ".")
      GoTo 99
   End If
   On Error GoTo 0
   '
'5.- GUARDAR EL ASELE ACTUAL EN LA CARPETA VIEJOS
   ANAMCO$ = Left$(ASELE$, Len(ASELE$) - 1)
   If EXISTE%(CARPEVIE$ + ANAMCO$ + "0") > 0 Then
      AVIE0$ = LOADFILE$(CARPEVIE$ + ANAMCO$ + "0")
      ASEL0$ = LOADFILE$(App.Path + "\" + ASELE$)
      If ASEL0$ <> AVIE0$ Then
         On Error Resume Next
         Kill CARPEVIE$ + ANAMCO$ + "9"
         On Error GoTo 0
         If EXISTE%(CARPEVIE$ + ANAMCO$ + "9") > 0 Then
            Call MENSERR(24, "Imposible Completar Transaccion Solicitada.\  \No Pudo Borrarse " + CARPEVIE$ + ANAMCO$ + "0.")
            GoTo 99
         End If
         On Error Resume Next
         For KI% = 8 To 0 Step -1
            Name CARPEVIE$ + ANAMCO$ + TRIM$(Str$(KI%)) As CARPEVIE$ + ANAMCO$ + TRIM$(Str$(KI% + 1))
         Next KI%
         On Error GoTo 0
         If EXISTE%(CARPEVIE$ + ANAMCO$ + "0") > 0 Then
            Call MENSERR(24, "Imposible Completar Transaccion Solicitada.\   \No Pudo Renombrarse " + CARPEVIE$ + ANAMCO$ + "0 como " + CARPEVIE$ + ANAMCO$ + "1")
            On Error GoTo 0
            GoTo 99
         End If
      End If
   End If
   '
   On Error Resume Next
   FileCopy App.Path + "\" + ASELE$, CARPEVIE$ + ANAMCO$ + "0"
   If Err Then
      On Error GoTo 0
      CAPVI$ = REPLACAR$(CARPEVIE$, "\", "/")
      Call MENSERR(24, "Imposible Copiar '" + ASELE$ + "' Como '" + ANAMCO$ + "'\a Carpeta " + CAPVI$ + ".")
      GoTo 99
   End If
   On Error GoTo 0
   '
'6.- BORRAR ASELE$ EN LA CARPETA ACTUAL
20 On Error Resume Next
   Kill App.Path + "\" + ASELE$
   If Err Then
      On Error GoTo 0
      TTXX$ = "Imposible Borrar '" + ASELE$ + "' en Carpeta de Trabajo Actual."
      TTXX$ = TTXX$ + "\   \Causas Posibles:"
      TTXX$ = TTXX$ + "\1.- Archivo en Uso por Otro Usuario"
      TTXX$ = TTXX$ + "\2.- Archivo Marcado como 'Solo Lectura'"
      TTXX$ = TTXX$ + "\3.- Usuario sin Derecho sobre Carpeta Destino"
      If COMALTER%(TTXX$ + "\   \\Resolver y Re-Intentar\Cancelar Restauracion\") = 1 Then GoTo 20
      GoTo 99
   End If
   On Error GoTo 0
   '
30 On Error Resume Next
   FileCopy CARPEVIE$ + ASELE$, App.Path + "\" + ASELE$
   If Err Then
      On Error GoTo 0
      TTXX$ = "Imposible Copiar '" + ASELE$ + "' Hacia Carpeta de Trabajo Actual."
      TTXX$ = TTXX$ + "\   \Causas Posibles:"
      TTXX$ = TTXX$ + "\1.- Archivo en Uso por Otro Usuario"
      TTXX$ = TTXX$ + "\2.- Archivo Marcado como 'Solo Lectura'"
      TTXX$ = TTXX$ + "\3.- Usuario sin Derecho sobre Carpeta Destino"
      If COMALTER%(TTXX$ + "\   \\Resolver y Re-Intentar\Cancelar Restaurar\") = 1 Then GoTo 30
      GoTo 99
   End If
   On Error GoTo 0
   '
   Call COMUNI("Restauracion Completada")
   Unload Me                         ' para forzar el refresco de las listas
   '
99 Command2.Enabled = True
   '
End Sub

Private Sub Command6_Click()
   Hide
End Sub

Private Sub Label2_Change()
   ASELE$ = TRIM$(Left$(Label2, 32))
   PPXX% = InStr(ASELE$, "."): If PPXX% < 1 Then PPXX% = 1 + Len(ASELE$)
   File2.Path = App.Path + "\VIEJOS"
   File2.Pattern = Left$(ASELE$, PPXX%) + "*"
   List4.Clear
   For KU& = 1 To File2.ListCount
     FIDAT$ = Format$(CDbl(FileDateTime(File2.Path + "\" + File2.List(KU& - 1))), "00000#.#####   ")
     While Len(FIDAT$) < 15: FIDAT$ = FIDAT$ + " ": Wend
     List4.AddItem FIDAT$ + File2.List(KU& - 1)
   Next KU&
   '
   List1.Clear
   For KU& = List4.ListCount To 1 Step -1
     YY$ = List4.List(KU& - 1)
     IZQUI$ = REPLACAR$(Left$(YY$, 14), ",", ".")
     XX$ = Left$(Mid$(YY$, 16, 32) + Space$(32), 32) + Format$(CDate(Val(IZQUI$)), "dd/mm/yy  hh:mm:ss")
     List1.AddItem UCase$(XX$)
   Next KU&
End Sub

Private Sub List1_Click()
  Label4 = List1.TEXT
End Sub

Private Sub LIST3_Click()
  Label1 = TRIM$(Left$(List3.TEXT, 32))
  Label2 = List3.TEXT
End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame18.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Option1.BackColor = &HFCD7B6
      Me.Option2.BackColor = &HFCD7B6
      Me.Option3.BackColor = &HFCD7B6
      Me.Option4.BackColor = &HFCD7B6
      
      ' &H8000000F
    End If
    '
    Call APLICACIONSKINS(Me)
    '
    Screen.MousePointer = 11
    EPAC$ = EMPREAC$
    If EPAC$ <> "" Then
      Caption = Caption + " - " + EMPREAC$
    End If
    '
    On Error Resume Next
    Label11.Caption = App.Path
    File1.Path = App.Path
    Text1 = "*.EXE"
    Screen.MousePointer = 1
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
   Unload Me
End Sub

Private Sub Label11_Change()
  If Right$(Label11, 1) = "\" Then
    Label11 = Left$(Label11, Len(Label11) - 1)
  End If
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

Private Sub salir_Click()
   Call Command6_Click
End Sub

Private Sub Text1_Change()
   File1.Pattern = Text1
   List2.Clear
   For KU& = 1 To File1.ListCount
     FIDAT$ = Format$(CDbl(FileDateTime(Label11 + "\" + File1.List(KU& - 1))), "00000#.#####   ")
     While Len(FIDAT$) < 15: FIDAT$ = FIDAT$ + " ": Wend
     List2.AddItem FIDAT$ + File1.List(KU& - 1)
   Next KU&
   '
   List3.Clear
   For KU& = List2.ListCount To 1 Step -1
     YY$ = List2.List(KU& - 1)
     IZQUI$ = REPLACAR$(Left$(YY$, 14), ",", ".")
     XX$ = Left$(Mid$(YY$, 16, 32) + Space$(32), 32) + Format$(CDate(Val(IZQUI$)), "dd/mm/yy  hh:mm:ss")
     List3.AddItem UCase$(XX$)
   Next KU&
End Sub

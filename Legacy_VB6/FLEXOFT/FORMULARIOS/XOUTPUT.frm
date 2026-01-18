VERSION 5.00
Begin VB.Form XOUTPUT 
   BackColor       =   &H00C0EEFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Exportar"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Archivo de Texto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   6930
      TabIndex        =   33
      Top             =   1995
      Width           =   1905
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Hoja de Cálculo (Planilla Excel)"
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
      Height          =   540
      Index           =   2
      Left            =   6930
      TabIndex        =   34
      Top             =   3780
      Width           =   1695
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Exportar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   6720
      TabIndex        =   35
      Top             =   1575
      Width           =   4950
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0EEFF&
         Caption         =   "Opciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   210
         TabIndex        =   40
         Top             =   840
         Width           =   2010
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0EEFF&
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
            Index           =   1
            Left            =   210
            TabIndex        =   43
            Top             =   525
            Width           =   1695
         End
         Begin VB.OptionButton Option3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0EEFF&
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
            Index           =   0
            Left            =   210
            TabIndex        =   42
            Top             =   315
            Width           =   1695
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00C0EEFF&
            Caption         =   "Mostrar Archivo"
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
            Left            =   210
            TabIndex        =   41
            Top             =   840
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command4 
         Height          =   315
         Left            =   4200
         Picture         =   "XOUTPUT.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   300
         Width           =   540
      End
      Begin VB.DirListBox Dir1 
         Height          =   2115
         Left            =   2520
         TabIndex        =   38
         Top             =   735
         Width           =   2220
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   2520
         TabIndex        =   37
         Top             =   315
         Width           =   1275
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   210
         TabIndex        =   36
         Text            =   " "
         Top             =   3255
         Width           =   4530
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   2310
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Line Line6 
         X1              =   2310
         X2              =   2310
         Y1              =   105
         Y2              =   2940
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   4935
         Y1              =   2940
         Y2              =   2940
      End
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Impresora:"
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
      Index           =   0
      Left            =   420
      TabIndex        =   32
      Top             =   2100
      Width           =   1275
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   7035
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   30
      Top             =   1050
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
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
      Left            =   9030
      Picture         =   "XOUTPUT.frx":047A
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Guarda Configuración del Listado"
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Archivo de Trabajo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   14
      Top             =   210
      Width           =   6315
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
         Height          =   330
         Left            =   4305
         TabIndex        =   16
         Text            =   " "
         Top             =   315
         Width           =   1590
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   315
         TabIndex        =   15
         Text            =   " "
         Top             =   735
         Width           =   5610
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo de Trabajo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1890
         TabIndex        =   17
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Setup"
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
      Left            =   8085
      Picture         =   "XOUTPUT.frx":0784
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Configura Impresora"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run"
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
      Left            =   9975
      Picture         =   "XOUTPUT.frx":0BC6
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Emite Listado"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton BOTEXIT 
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
      Left            =   10920
      Picture         =   "XOUTPUT.frx":0ED0
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Abandona el Listador"
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0EEFF&
      Caption         =   "Dispositivo de Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   210
      TabIndex        =   0
      Top             =   1575
      Width           =   6315
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1575
         TabIndex        =   25
         Text            =   " "
         Top             =   840
         Width           =   2955
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0EEFF&
         Caption         =   "Medidas en Milimetros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1170
         Left            =   1575
         TabIndex        =   18
         Top             =   2100
         Width           =   4320
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   3
            Left            =   3255
            TabIndex        =   27
            Text            =   " "
            Top             =   735
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   2
            Left            =   3255
            TabIndex        =   26
            Text            =   " "
            Top             =   420
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   23
            Text            =   " "
            Top             =   735
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   21
            Text            =   " "
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Margen Sup:"
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
            Index           =   3
            Left            =   1995
            TabIndex        =   29
            Top             =   420
            Width           =   1170
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Izquierdo:"
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
            Index           =   2
            Left            =   2310
            TabIndex        =   28
            Top             =   735
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Largo:"
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
            Index           =   1
            Left            =   -105
            TabIndex        =   22
            Top             =   735
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ancho:"
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
            Index           =   0
            Left            =   -105
            TabIndex        =   20
            Top             =   420
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0EEFF&
         Caption         =   "Tamaño"
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
         Left            =   4830
         TabIndex        =   12
         Top             =   315
         Width           =   1065
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   210
            TabIndex        =   13
            Text            =   " "
            Top             =   420
            Width           =   645
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0EEFF&
         Caption         =   "Copias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   4830
         TabIndex        =   10
         Top             =   1260
         Width           =   1065
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   315
            TabIndex        =   11
            Text            =   " "
            Top             =   315
            Width           =   435
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0EEFF&
         Caption         =   "Orientación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   1575
         TabIndex        =   7
         Top             =   1260
         Width           =   2955
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0EEFF&
            Caption         =   "Paisaje"
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
            Index           =   1
            Left            =   1785
            TabIndex        =   9
            Top             =   315
            Width           =   1065
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0EEFF&
            Caption         =   "Retrato"
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
            Index           =   0
            Left            =   210
            TabIndex        =   8
            Top             =   315
            Width           =   1065
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0EEFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   435
         Left            =   630
         TabIndex        =   6
         Top             =   4305
         Width           =   3690
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1575
         TabIndex        =   4
         Text            =   " "
         Top             =   420
         Width           =   2955
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Papel:"
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
         Left            =   630
         TabIndex        =   19
         Top             =   2205
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fuente:"
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
         Left            =   525
         TabIndex        =   5
         Top             =   945
         Width           =   855
      End
   End
End
Attribute VB_Name = "XOUTPUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IMPREDETER As String
'

Private Sub BOTEXIT_Click()
    'Call CIERRARCH("*.*")
    If COMMA$ <> "" Then
    '   Call BLOQEXE("*")
    End If
    'WLISTAB.Hide
End Sub

Private Sub Combo1_LostFocus()
    '
    Dim PRX As Printer
    For I% = 1 To Printers.Count
        If Printers(I% - 1).DeviceName = Combo1.Text Then
            ' La define como predeterminada del sistema.
            Set Printer = Printers(I% - 1)
            GoTo 10
        End If
    Next
    '
10  Combo1.Text = Printer.DeviceName
    '
End Sub

Private Sub Command1_Click()
    '
    If Trim$(IMPREDETER$) <> Trim$(Combo1.Text) Then
        MENSA$ = "Imposible Aplicar Impresora Seleccionada:\"
        MENSA$ = MENSA$ + "'" + Trim$(Combo1.Text) + "'\   \"
        MENSA$ = MENSA$ + "FLEXOFT ha Restaurado la Impresora Predeterminada:\"
        MENSA$ = MENSA$ + "'" + Trim$(IMPREDETER$) + "'\  \"
        MENSA$ = MENSA$ + "Para Utilizar la Elegida, Cambie la Impresora\"
        MENSA$ = MENSA$ + "Predeterminada para Windows."
        Combo1.Text = IMPREDETER$
        'Call COMUNI(MENSA$)
        Combo1.SetFocus
        Exit Sub
    End If
    '
    'Call GRADATIMPRE
    'Call CARDATIMPRE
    '
    SOBRESCRI% = 0
    If Option3(1) = True Then
        SOBRESCRI% = 1
    End If
    '
    PORIAN = Printer.Orientation
    If Option2(1).Value = True Then
          Printer.Orientation = 2
        Else
          Printer.Orientation = 1
    End If
    '
    'Call VERDATLIS
    '
    If DATLISTOK% <= 0 Then
         On Error Resume Next
         If Trim$(ARCHISAL$) = "" Then Printer.Orientation = PORIAN
         On Error GoTo 0
         Exit Sub
      ElseIf DATLISTOK% = 1 Then
         If SALEXCEL% = 1 Then
           'Call ARMAENCA
           'Call LISEXCEL
           'Call CIERRARCH("*.*")
           'WLISTAB.Hide
           Exit Sub
         End If
         'Call SETPRINTER
         'Call CARHEADERS
         'Call ARMAENCA
         If CACOPIAS% < 1 Then CACOPIAS% = 1
         CACOP% = CACOPIAS%
         If ARCHISAL$ <> "" Then CACOP% = 1
         For COP% = 1 To CACOP%
            'Call INITOTA
            If TNC% < 1 Then
                'Call ENCAHOJA
            End If
            'Call LISTABU
         Next COP%
    End If
    '
    'Call CIERRARCH("*.*")
    PRINTERSET% = 0
    Printer.Orientation = PORIAN
    If COMMA$ <> "" Then
     '  Call BLOQEXE("*")
    End If
    'WLISTAB.Hide
    '
    If ARCHISAL$ <> "" Then
     '   If WLISTAB.Check1.Value = 1 Then
            LCB$ = LUBAS$
       '     If EXISTE%("C:\FLEXOFT\WORDPAD.EXE") = 1 Then
                 Shell "C:\FLEXOFT\WORDPAD.EXE " + ARCHISAL$, vbMaximizedFocus
               Else
                 MENSA$ = "Imposible Mostrar Archivo Generado."
                 MENSA$ = MENSA$ + "\  \Falta Utilitario 'WORDPAD.EXE'"
                 MENSA$ = MENSA$ + "\en Directorio 'C:/FLEXOFT'."
                 MENSA$ = MENSA$ + "\   \Instale este Programa Utilitario\y Re-Genere el Listado."
        '         Call COMUNI(MENSA$)
         '   End If
      '  End If
    End If
    '
End Sub
'

Private Sub Command2_Click()
    'Call SETUPRINTER
End Sub

Private Sub Command3_Click()
    'Call GRADATIMPRE
    'Call CARDATIMPRE
End Sub

Private Sub Form_Load()
    '
    Refresh
    PORIAN = Printer.Orientation
    '
    Printer.TrackDefault = False
    '
    'Call CARDATLIS
    'Call CARDATIMPRE
    '
    ARCHISAL$ = ""
    Text2.Text = ARCHISAL$
    Option1(0).Value = True        ' IMPRESORA
    '
    SOBRESCRI% = 0
    FORMALI$ = FORMA$
    '
    Combo1.Clear
    Dim PRX As Printer
    For Each PRX In Printers
        Combo1.AddItem PRX.DeviceName
        ' La define como predeterminada del sistema.
        'Set Printer = prx
    Next
    Combo1.Text = Printer.DeviceName
    IMPREDETER$ = Printer.DeviceName
    '
    If Left$(UCase$(ATRIBU$), 4) = "LPT1" Then
        'WLISTAB.Refresh
        'Call SETPRINTER
        'Call CARHEADERS
        'Call ARMAENCA
        'Call INITOTA
        If TNC% < 1 Then
            'Call ENCAHOJA
        End If
        'Call LISTABU
        'Call CIERRARCH("*.*")
        PRINTERSET% = 0
        If COMMA$ <> "" Then
        '   Call BLOQEXE("*")
        End If
        'WLISTAB.Hide
    End If
    '
    Printer.Orientation = PORIAN
    'If EXISTE%("FLEXCEL.EXE") > 0 Then
    '  Option1(2).Enabled = True
    'End If
    Screen.MousePointer = 1
    '
End Sub
'

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        ARCHISAL$ = ""
        Text2.Text = ""
        Text2.Enabled = False
        'Text8.Text = ""
        'Text8.Enabled = False
        Option3(0).Value = False
        Option3(0).Enabled = False
        Option3(1).Value = False
        Option3(1).Enabled = False
        Check1.Enabled = False
        SALEXCEL% = 0
      ElseIf Index = 1 Then
        ARCHISAL$ = LUDAT$ + Trim$(ARCHITRA$) + ".TXT"
        Text2.Text = ARCHISAL$
        'Text8.Text = ""
        'Text8.Enabled = False
        Option3(0).Value = True
        Option3(0).Enabled = True
        Option3(1).Enabled = True
        Check1.Enabled = True
        SALEXCEL% = 0
        Text2.Enabled = True
        Text2.SetFocus
      ElseIf Index = 2 Then
        ARCHISAL$ = LUDAT$ + Trim$(ARCHITRA$) + ".XLS"
        Text8.Text = ARCHISAL$
        Text2.Text = ""
        Text2.Enabled = False
        Option3(0).Value = False
        Option3(0).Enabled = False
        Option3(1).Value = False
        Option3(1).Enabled = False
        Check1.Enabled = False
        SALEXCEL% = 1
        Text8.Enabled = True
        Text8.SetFocus
    End If
End Sub

Private Sub Text3_GotFocus()
    If Option1(0) = True Then
        Option1(0).SetFocus
      Else
        Text2.SetFocus
    End If
End Sub
'

Private Sub Text1_GotFocus()
    If Option1(0) = True Then
        Option1(0).SetFocus
      Else
        Text2.SetFocus
    End If
End Sub


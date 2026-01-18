VERSION 5.00
Begin VB.Form COMANOCO 
   BackColor       =   &H00EEEEDD&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listados de Ordenes de Trabajo"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4830
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4095
      ScaleHeight     =   75
      ScaleWidth      =   570
      TabIndex        =   13
      Top             =   1050
      Width           =   630
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   -225
         TabIndex        =   14
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EEEEDD&
      Caption         =   "Selección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   1995
      TabIndex        =   10
      Top             =   1050
      Width           =   1905
      Begin VB.OptionButton Option1 
         BackColor       =   &H00EEEEDD&
         Caption         =   "Todas"
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
         Left            =   105
         TabIndex        =   11
         Top             =   260
         Width           =   1170
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00EEEEDD&
         Caption         =   "Solo Pendientes"
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
         TabIndex        =   12
         Top             =   630
         Width           =   1750
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEEEDD&
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   105
      TabIndex        =   7
      Top             =   1050
      Width           =   1800
      Begin VB.CheckBox Check2 
         BackColor       =   &H00EEEEDD&
         Caption         =   "O/Reparación"
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
         Left            =   105
         TabIndex        =   9
         Top             =   535
         Width           =   1590
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00EEEEDD&
         Caption         =   "O/Trabajo"
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
         Left            =   105
         TabIndex        =   8
         Top             =   220
         Width           =   1380
      End
   End
   Begin VB.CommandButton Command19 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   4095
      Picture         =   "COMANOCO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Emisión Listado Impreso"
      Top             =   1300
      Width           =   630
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
      Height          =   690
      Left            =   4095
      Picture         =   "COMANOCO.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cerrar - Salir de la Aplicación"
      Top             =   210
      Width           =   630
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEDD&
      Caption         =   "Rango de Fechas"
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
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   3795
      Begin VB.Label Label2 
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
         Height          =   300
         Left            =   2625
         TabIndex        =   2
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label1 
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
         Height          =   300
         Left            =   870
         TabIndex        =   1
         Top             =   315
         Width           =   960
      End
      Begin VB.Label Label4 
         BackColor       =   &H00EEEEDD&
         Caption         =   "Hasta:"
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
         Left            =   1995
         TabIndex        =   4
         Top             =   420
         Width           =   645
      End
      Begin VB.Label Label3 
         BackColor       =   &H00EEEEDD&
         Caption         =   "Desde:"
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
         TabIndex        =   3
         Top             =   420
         Width           =   645
      End
   End
End
Attribute VB_Name = "COMANOCO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Hide
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    '
    DES% = FECHANUM(Label1)
    HAS% = FECHANUM(Label2)
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    If Check1.Value = 1 And Check2.Value = 1 Then
       If Option1.Value = True Then ' Todas
          Screen.MousePointer = 11
          JJ& = 0
          Call COPYSTRU("PARTEMOD", "PARTESEL")
          FIN& = ULTREG&("PARTEMOD")
          For U& = 1 To FIN&
            Call TRACE(20, U&, FIN&)
            XX$ = REGLEIDO$("PARTEMOD", U&)
            FEMII% = CVI(Mid$(XX$, 3, 2))
            If FEMII% >= DES% Then
              If FEMII% <= HAS% Then
                SSS$ = Mid$(XX$, 5, 2)
                TIEDEC = DIFHORA(CVI(Mid$(XX$, 53, 2)), CVI(Mid$(XX$, 55, 2)))
                Call REPLA(XX$, MKS$(TIEDEC), 65, 4)
                JJ& = JJ& + 1
                Call GRAREG("PARTESEL", XX$, JJ&)
              End If
            End If
          Next U&
          '
          Call SETULTREG("PARTESEL", JJ&)
          '
       Else ' Pendientes
          Screen.MousePointer = 11
          JJ& = 0
          Call COPYSTRU("PARTEMOD", "PARTESEL")
          FIN& = ULTREG&("PARTEMOD")
          For U& = 1 To FIN&
            Call TRACE(20, U&, FIN&)
            XX$ = REGLEIDO$("PARTEMOD", U&)
            FEMII% = CVI(Mid$(XX$, 3, 2))
            If FEMII% >= DES% Then
              If FEMII% <= HAS% Then
                ORTRA$ = Mid$(XX$, 5, 12)
                If Left$(ORTRA$, 2) = "OT" Then
                  ZZ$ = FILTERGETRECORD$("INDIOTRA", 1, ORTRA$)
                  If Mid$(ZZ$, 63, 2) = MKI$(0) Then
                    TIEDEC = DIFHORA(CVI(Mid$(XX$, 53, 2)), CVI(Mid$(XX$, 55, 2)))
                    Call REPLA(XX$, MKS$(TIEDEC), 65, 4)
                    JJ& = JJ& + 1
                    Call GRAREG("PARTESEL", XX$, JJ&)
                  End If
                Else
                  If Left$(ORTRA$, 2) = "OR" Then
                    NUMOR& = Val(Mid$(ORTRA$, 4))
                    ZZ$ = FILTERGETRECORD$("PRESREPA", 1, MKS$(NUMOR&))
                    If CVI(Mid$(ZZ$, 123, 2)) < 1 Then
                      TIEDEC = DIFHORA(CVI(Mid$(XX$, 53, 2)), CVI(Mid$(XX$, 55, 2)))
                      Call REPLA(XX$, MKS$(TIEDEC), 65, 4)
                      JJ& = JJ& + 1
                      Call GRAREG("PARTESEL", XX$, JJ&)
                    End If
                  End If
                End If
              End If
            End If
          Next U&
          '
          Call SETULTREG("PARTESEL", JJ&)
       End If
    Else
       If Check1.Value = 1 Then ' O/Trabajo
         If Option1.Value = True Then ' Todas
            Screen.MousePointer = 11
            JJ& = 0
            Call COPYSTRU("PARTEMOD", "PARTESEL")
            FIN& = ULTREG&("PARTEMOD")
            For U& = 1 To FIN&
              Call TRACE(20, U&, FIN&)
              XX$ = REGLEIDO$("PARTEMOD", U&)
              FEMII% = CVI(Mid$(XX$, 3, 2))
              If FEMII% >= DES% Then
                If FEMII% <= HAS% Then
                  SSS$ = Mid$(XX$, 5, 2)
                  If Mid$(XX$, 5, 2) = "OT" Then
                    TIEDEC = DIFHORA(CVI(Mid$(XX$, 53, 2)), CVI(Mid$(XX$, 55, 2)))
                    Call REPLA(XX$, MKS$(TIEDEC), 65, 4)
                    JJ& = JJ& + 1
                    Call GRAREG("PARTESEL", XX$, JJ&)
                  End If
                End If
              End If
            Next U&
            '
            Call SETULTREG("PARTESEL", JJ&)
            '
         Else ' Solo las pendientes
            '
            Screen.MousePointer = 11
            JJ& = 0
            Call COPYSTRU("PARTEMOD", "PARTESEL")
            FIN& = ULTREG&("PARTEMOD")
            For U& = 1 To FIN&
              Call TRACE(20, U&, FIN&)
              XX$ = REGLEIDO$("PARTEMOD", U&)
              FEMII% = CVI(Mid$(XX$, 3, 2))
              If FEMII% >= DES% Then
                If FEMII% <= HAS% Then
                  ORTRA$ = Mid$(XX$, 5, 12)
                  If Left$(ORTRA$, 2) = "OT" Then
                    ZZ$ = FILTERGETRECORD$("INDIOTRA", 1, ORTRA$)
                    If Mid$(ZZ$, 63, 2) = MKI$(0) Then
                      TIEDEC = DIFHORA(CVI(Mid$(XX$, 53, 2)), CVI(Mid$(XX$, 55, 2)))
                      Call REPLA(XX$, MKS$(TIEDEC), 65, 4)
                      JJ& = JJ& + 1
                      Call GRAREG("PARTESEL", XX$, JJ&)
                    End If
                  End If
                End If
              End If
            Next U&
            '
            Call SETULTREG("PARTESEL", JJ&)
            '
         End If
       End If
       '
       If Check2.Value = 1 Then ' O/Reparacion
         If Option1.Value = True Then ' Todas
            Screen.MousePointer = 11
            JJ& = 0
            Call COPYSTRU("PARTEMOD", "PARTESEL")
            FIN& = ULTREG&("PARTEMOD")
            For U& = 1 To FIN&
              Call TRACE(20, U&, FIN&)
              XX$ = REGLEIDO$("PARTEMOD", U&)
              FEMII% = CVI(Mid$(XX$, 3, 2))
              If FEMII% >= DES% Then
                If FEMII% <= HAS% Then
                  SSS$ = Mid$(XX$, 5, 2)
                  If Mid$(XX$, 5, 2) = "OR" Then
                    TIEDEC = DIFHORA(CVI(Mid$(XX$, 53, 2)), CVI(Mid$(XX$, 55, 2)))
                    Call REPLA(XX$, MKS$(TIEDEC), 65, 4)
                    JJ& = JJ& + 1
                    Call GRAREG("PARTESEL", XX$, JJ&)
                  End If
                End If
              End If
            Next U&
            '
            Call SETULTREG("PARTESEL", JJ&)
            '
         Else ' Solo las pendientes
            '
            Screen.MousePointer = 11
            JJ& = 0
            Call COPYSTRU("PARTEMOD", "PARTESEL")
            FIN& = ULTREG&("PARTEMOD")
            For U& = 1 To FIN&
              Call TRACE(20, U&, FIN&)
              XX$ = REGLEIDO$("PARTEMOD", U&)
              FEMII% = CVI(Mid$(XX$, 3, 2))
              If FEMII% >= DES% Then
                If FEMII% <= HAS% Then
                  ORTRA$ = Mid$(XX$, 5, 12)
                  If Left$(ORTRA$, 2) = "OR" Then
                    NUMOR& = Val(Mid$(ORTRA$, 4))
                    ZZ$ = FILTERGETRECORD$("PRESREPA", 1, MKS$(NUMOR&))
                    If CVI(Mid$(ZZ$, 123, 2)) < 1 Then
                      TIEDEC = DIFHORA(CVI(Mid$(XX$, 53, 2)), CVI(Mid$(XX$, 55, 2)))
                      Call REPLA(XX$, MKS$(TIEDEC), 65, 4)
                      JJ& = JJ& + 1
                      Call GRAREG("PARTESEL", XX$, JJ&)
                    End If
                  End If
                End If
              End If
            Next U&
            '
            Call SETULTREG("PARTESEL", JJ&)
            '
         End If
       End If
    End If
    '
    Call CIERRARCH("PARTESEL")
    '
    Call FILESORT("PARTESEL", "", 4, 5, "ASC")
    Call HEADERS("PARTESEL", "")
    Call HEADERS("PARTESEL", "Corresponde a: " + PERIO$)
    Call FINAL("*LIPAROT")
    Screen.MousePointer = 1
    '
    Command19.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Label1 = FECHATEX$(CInt(COMANO.Label5))
    Label2 = FECHATEX$(CInt(COMANO.Label8))
    Check1.Value = 1
    Check2.Value = 1
    Option1.Value = True
    '
End Sub

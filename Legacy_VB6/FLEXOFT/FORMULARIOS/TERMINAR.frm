VERSION 5.00
Begin VB.Form TERMINAR 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Terminar Sesión"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4530
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   3465
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
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
      Left            =   2415
      TabIndex        =   1
      Top             =   2100
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
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
      Left            =   420
      TabIndex        =   0
      Top             =   2100
      Width           =   1695
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Salir de FLEXOFT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   630
      TabIndex        =   4
      Top             =   1365
      Width           =   3585
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Copia de Respaldo (backup)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   630
      TabIndex        =   3
      Top             =   945
      Width           =   3690
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Reanudar Ejecución"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   630
      TabIndex        =   2
      Top             =   525
      Value           =   -1  'True
      Width           =   3585
   End
End
Attribute VB_Name = "TERMINAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Option3.Value = True Then
        If PROCACTI% <> 0 Then Exit Sub
        Call GENBACKUP
        Call GRACONTROL("", "FECHORA", "")
        Call FUERASYS
        '
      ElseIf Option2.Value = True Then
        If PROCACTI% <> 0 Then Exit Sub
        Call GENBACKUP
        '
'        LCB$ = LUBAS$
'        If EXISTE%("C:\FLEXOFT\ARJ.EXE") < 1 Then
'           BRXX$ = REPLACAR$("C:\FLEXOFT\ARJ.EXE", "\", "/")
'           FLEXOFTA.List1.Visible = False
'           Call COMUNI("No se Verificó la Existencia del Programa Utilitario\  \'" + BRXX$ + "'.\  \No Podrá Ejecutarse Normalmente\la Opción de Confección de Backup.\  \Consulte a su Soporte de Sistemas.")
'           Exit Sub
'        End If
'        '
'        MENSA$ = "Para la Confección de Copias de Respaldo el\Sistema FLEXOFT Invoca los Programas Utilitarios\Denominados PKZIP.EXE y ARJ.EXE,\los cuales NO Integran el Suministro del Sistema Aplicativo.\  \"
'        MENSA$ = MENSA$ + "Es Responsabilidad Exclusiva del Usuario, el Legitimar\ Frente al Correspondiente Derecho-habiente\las Copias no Comerciales Suministradas\Junto con el Programa de Aplicación.\  \"
'        MENSA$ = MENSA$ + "La Presente Constituye una Comunicación Legal."
'        If COMALTER%(MENSA$ + "\\Continuar Proceso\Cancelar") <> 1 Then
'            Exit Sub
'        End If
'        '
'        LU1$ = Left$(CurDir, 2) + LUDAT$
'        On Error Resume Next
'        LUDAXX$ = LUDAT$
'        Kill LUDAXX$ + "*.TXT"
'        On Error GoTo 0
'        ULOAC$ = Left$(CurDir, 2)
'        LUDATAC$ = TRIM$(LUDAT$)
'        If Len(LUDATAC$) > 1 Then
'           If Right$(LUDATAC$, 1) = "\" Then
'              LUDATAC$ = Left$(LUDATAC$, Len(LUDATAC$) - 1)
'           End If
'        End If
'        If Mid$(LUDATAC$, 2, 1) <> ":" Then
'           LUDATAC$ = ULOAC$ + LUDATAC$
'        End If
'        '
'        NX% = FILEOPEN%("C:\FLEXOFT\BORDATA.BAT", 2, 0)
'        Print #NX%, "DEL A:DATOS.*"
'        Close #NX%
'        '
'        NX% = FILEOPEN%("C:\FLEXOFT\FLEXBAT.BAT", 2, 0)
'        Print #NX%, "@ECHO OFF"
'        Print #NX%, "Cls"
'        Print #NX%, "ECHO Coloque el Primer Diskette de Respaldo en la Disquetera ..."
'        Print #NX%, "ECHO Siga las Instrucciones que Aparecen en Ingles ..."
'        Print #NX%, "ECHO Recuerde Legitimar su Copia del programa ARJ !!!"
'        Print #NX%, "ECHO Pulse RETURN para seguir ... CRTL-C para abortar ..."
'        Print #NX%, "PAUSE"
'        Print #NX%, "CLS"
'        Print #NX%, "DEL A:DATOS.*"
'
'        'Print #NX%, ULOAC$           ' nuevo esquema
'        Print #NX%, "CD " + LUDATAC$ ' nuevo esquema
'        LU1$ = ULOAC$                   ' nuevo esquema
'        '
'        COMARJ$ = "C:\FLEXOFT\ARJ a A:DATOS "
'        COMARJ$ = COMARJ$ + LU1$ + "*.DAT "
'        COMARJ$ = COMARJ$ + LU1$ + "*.OLD "
'        COMARJ$ = COMARJ$ + LU1$ + "*.X* "
'        COMARJ$ = COMARJ$ + LU1$ + "*.FRM "
'        COMARJ$ = COMARJ$ + LU1$ + "*.RFS "
'        COMARJ$ = COMARJ$ + "-vsC:\FLEXOFT\BORDATA.BAT -va"
'        Print #NX%, COMARJ$
'        Print #NX%, "EXIT"
'        Close #NX%
'        '
'        If EXISTE%("C:\FLEXOFT\FLEXPIF.PIF") = 1 Then
'               Shell "C:\FLEXOFT\FLEXPIF.PIF", vbNormalFocus
'             Else
'               TXT$ = REPLACAR$("C:\FLEXOFT\FLEXPIF.PIF", "\", "/")
'               Call MENSERR(24, "No Existe " + TXT$)
'        End If
'        Call GRACONTROL("", "FECHORA", HOS$ + Time$)
        Call FUERASYS
        '
      ElseIf Option1.Value = True Then
        TERMINAR.Hide
    End If
End Sub

Private Sub Command2_Click()
    TERMINAR.Hide
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub Option1_DblClick()
   TERMINAR.Hide
   'If PROCACTI% > 0 Then Exit Sub
   'USNBR% = 0
   'FLESCRIT.Frame2.Visible = False
   'FLESCRIT.Frame3.Visible = False
   'FLESCRIT.Frame4.Visible = False
   'FLESCRIT.Frame1.Visible = True
   'Hide
   'Call FRESHALL
   'FLESCRIT.Text1.SetFocus
End Sub

Private Sub Option3_DblClick()
    If PROCACTI% <> 0 Then Exit Sub
    Call GENBACKUP
    Call GRACONTROL("", "FECHORA", "")
    Call FUERASYS
End Sub

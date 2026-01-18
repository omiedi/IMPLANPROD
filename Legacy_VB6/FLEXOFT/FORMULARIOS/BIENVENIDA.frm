VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form BIENVENIDA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ingreso Al Sistema"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   1785
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   14
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   1155
      OleObjectBlob   =   "BIENVENIDA.frx":0000
      TabIndex        =   13
      Top             =   2940
      Visible         =   0   'False
      Width           =   1485
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "BIENVENIDA.frx":007A
      Top             =   2730
   End
   Begin VB.Frame Frame1 
      Height          =   2010
      Left            =   105
      TabIndex        =   1
      Top             =   735
      Width           =   4215
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Lan-Id:"
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
         Left            =   1365
         TabIndex        =   11
         Top             =   630
         Width           =   645
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa:"
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
         Left            =   1185
         TabIndex        =   10
         Top             =   315
         Width           =   855
      End
      Begin VB.Label Label9 
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
         Height          =   225
         Left            =   2100
         TabIndex        =   9
         Top             =   630
         Width           =   2070
      End
      Begin VB.Label Label8 
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
         Height          =   225
         Left            =   2070
         TabIndex        =   8
         Top             =   315
         Width           =   2115
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario: "
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
         Left            =   1260
         TabIndex        =   7
         Top             =   945
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Servidor:"
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
         Left            =   1185
         TabIndex        =   6
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label5 
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
         Height          =   225
         Left            =   2070
         TabIndex        =   5
         Top             =   1260
         Width           =   2115
      End
      Begin VB.Label Label3 
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
         Height          =   225
         Left            =   2070
         TabIndex        =   4
         Top             =   945
         Width           =   2115
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Base de Datos: "
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
         Left            =   630
         TabIndex        =   3
         Top             =   1575
         Width           =   1485
      End
      Begin VB.Label Label7 
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
         Height          =   225
         Left            =   2070
         TabIndex        =   2
         Top             =   1575
         Width           =   2115
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Continuar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2835
      TabIndex        =   0
      Top             =   2835
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Bienvenidos al Sistema Flexoft !!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   105
      TabIndex        =   12
      Top             =   210
      Width           =   4215
   End
End
Attribute VB_Name = "BIENVENIDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Hide
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Sub APLICACIONSKINS(FRM As Form)
    '
    If CAPTURAERRORSKIN%(FRM) > 0 Then
       Exit Sub
    End If

    'VALIDA SI EXISTE EL ARCHIVO DE CONFIGURACION
    RUTXX$ = LUBAS$
    If INTERPRE% > 0 Then RUTXX$ = LUDAT$
    If EXISTE%(RUTXX$ + "FLEXSKIN.CFG") < 1 Then
       Exit Sub
    End If
    '
    'LEO EL ARCHIVO DE CONFIGURACION
    ARCHIVOSKIN$ = Trim$(Left$(LOADFILE$(RUTXX$ + "FLEXSKIN.CFG"), 512))
    '
    ' VALIDO EL ARCHIVO SI EXISTE
    If EXISTE%(ARCHIVOSKIN$) < 1 Then
      Exit Sub
    End If
    '
    Call PROCESO_SKIN1(FRM, FRM.Picture999(0))
    Call APLICAR_SKIN(FRM, ARCHIVOSKIN$)

End Sub


Sub APLICAR_SKIN(ByVal FORMULARIO As Form, SKINNSS As String)

'    APLICO EL SKIN
    FORMULARIO.Skin1.LoadSkin SKINNSS
    FORMULARIO.Skin1.ApplySkin FORMULARIO.hWnd
    '
    'LE HAGO TOMAR AL PICTURE EL COLOR DEL SKINS
    For I% = 1 To FORMULARIO.Picture999.Count - 1
      FORMULARIO.Picture999(I%).BackColor = FORMULARIO.Skin1.PanelColor
    Next I%

    For Each CONTROLX In FORMULARIO.Controls
      If TypeOf CONTROLX Is PictureBox Then
        If CONTROLX.Name <> "MARCOBARRA" Then
         CONTROLX.BackColor = FORMULARIO.Skin1.PanelColor
       End If
      End If
    Next
End Sub

Sub PROCESO_SKIN1(FORMULARIO As Form, PICTUREX As PictureBox)
     For Each CONTROLX In FORMULARIO.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO

        If TypeOf CONTROLX Is Label Then   ' SI ES UN LABEL
           Set OBJLABEL = CONTROLX ' SETEO EL OBJETO
           With OBJLABEL
              On Error Resume Next
              If .Visible = False And OBJLABEL.Name <> "LEYEN" Then GoTo 60 ' si no es visible no aplica el picture
              If TypeOf .Container Is PictureBox And TypeOf CONTROLX Is CommandButton Then GoTo 60
              On Error GoTo 0
              '
              If .BackStyle = 0 Or .BorderStyle = 0 Then
                 J% = J% + 1
                 If InStr(UCase(OBJLABEL.Name), "LEYEN") < 1 Then
                    Load FORMULARIO.SkinLabel999(J%)
                    With FORMULARIO.SkinLabel999(J%)
                       Set .Container = OBJLABEL.Container
                       'LE PASO LA MISMA MEDIDA QUE EL LABEL
                       .Visible = True
                       .Width = OBJLABEL.Width
                       .Top = OBJLABEL.Top
                       .Height = OBJLABEL.Height
                       .Left = OBJLABEL.Left
                       .Caption = OBJLABEL.Caption
                       .Font = OBJLABEL.Font
                       .Alignment = OBJLABEL.Alignment
                       If OBJLABEL.Alignment = 1 Then .Alignment = 2
                       If OBJLABEL.Alignment = 2 Then .Alignment = 1
                       .Visible = True
                    End With
                    .Visible = False ' OCULTA EL LABEL ORIGINAL
                    GoTo 60
                 End If
 
              ElseIf .BackStyle = 1 Or .BorderStyle = 1 Then
                 AA1 = FORMULARIO.Name
               
                 'CREO UN PICTURE
                 I% = I% + 1
                Load FORMULARIO.Picture999(I%)
                With FORMULARIO.Picture999(I%)
                    ' HAY QUE SETEAR EL .CONTAINER POR QUE SI NO NO FUNCIONA
                    Set .Container = OBJLABEL.Container
                    'LE PASO LA MISMA MEDIDA QUE EL LABEL
                    .Visible = True
                    .Width = OBJLABEL.Width
                    .Top = OBJLABEL.Top
                    .Height = OBJLABEL.Height
                    .Left = OBJLABEL.Left
                    .BorderStyle = 0
                    Set OBJLABEL.Container = FORMULARIO.Picture999(I%)
                    OBJLABEL.Top = 0
                    OBJLABEL.Left = 0
                    If OBJLABEL.BackStyle = 0 Then
                      OBJLABEL.BorderStyle = 0
                    End If
                End With
              End If
          End With
        End If
60   Next

End Sub

Function CAPTURAERRORSKIN%(FRM As Form)
   
   RUTXX$ = LUBAS$
   If INTERPRE% > 0 Then RUTXX$ = LUDAT$

   ARCHIVOSKIN$ = Trim$(Left$(LOADFILE$(RUTXX$ + "FLEXSKIN.CFG"), 512))
   On Error Resume Next
     A = FRM.SkinLabel999(0)
     FRM.Skin1.LoadSkin ARCHIVOSKIN$
   If Err > 0 Then CAPTURAERRORSKIN% = 1
   On Error GoTo 0

End Function

Private Sub Form_Activate()
   Call APLICACIONSKINS(Me)
End Sub


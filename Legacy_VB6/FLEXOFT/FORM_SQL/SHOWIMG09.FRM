VERSION 5.00
Begin VB.Form SHOWIMG09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   645
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   900
      TabIndex        =   0
      ToolTipText     =   "Haga Doble-Click para Imprimir Documento"
      Top             =   0
      Visible         =   0   'False
      Width           =   960
   End
End
Attribute VB_Name = "SHOWIMG09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AGRAFIC$
Public EQUIP$
Private Sub Form_Unload(Cancel As Integer)
   HORIMAGEN = Left
   VERTIMAGEN = Top
End Sub

Private Sub Picture2_DblClick()
   '
   IGRA% = INDIMAGEN%
   HOII% = HOY(HOS$)
   DOCUORIG$ = "GEMANTIN-" + HOS$ + "-" + Left$(Time$, 5)
   Call PRIDOIMG(IGRA%)
   '
End Sub
Sub PRIDOIMG(IGRA%)
   '
   Screen.MousePointer = 11
   '
   PRIFORGEN$ = Mid$(TRIM$(CONTROL$("", "FORGRAFI")), 2)
   CAIMAGEN% = 0
   '
   If PRIFORGEN$ <> "" Then
       '
       FORMALTER.Image3.Picture = LoadPicture("")
       FORMALTER.Image3.Picture = LoadPicture(AGRAFIC$)
       '
       HIMAGEN = FORMALTER.Image3.Picture.Width
       VIMAGEN = FORMALTER.Image3.Picture.Height
       If HIMAGEN > VIMAGEN Then
            PRIFORM$ = "H" + PRIFORGEN$
          Else
            PRIFORM$ = "V" + PRIFORGEN$
       End If
       '
       If EXISTE%(LUDAT$ + PRIFORM$ + ".FRM") Then
             '
             CAPARDOC% = 0
             CACOLTAB1% = 0
             CAITAB1% = 0
             CACOLTAB2% = 0
             CAITAB2% = 0
             '
             CODPARAM$(1) = "DES-MAT": DOCPARAM$(1) = EQUIP$
             CAPARDOC% = 1
             '
             COPLANO$ = AGRAFIC$
             If COPLANO$ <> "" Then
               If EXISTE%(COPLANO$) > 0 Then
                 CAIMAGEN% = 1
                 CODIMAGEN$(1) = "GRAFIPIE"
                 ARCHIMAGEN$(1) = COPLANO$
               End If
             End If
       End If
       Call PRINTDOC("@" + PRIFORM$)
     End If
   Screen.MousePointer = 1
   '
End Sub

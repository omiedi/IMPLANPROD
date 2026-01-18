VERSION 5.00
Begin VB.Form LISITBAJ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Items de Stock Dados de Baja"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Haga un Doble Click para Re-Vivir el Item"
      Top             =   0
      Width           =   7785
   End
End
Attribute VB_Name = "LISITBAJ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub List1_DblClick()
   '
   Screen.MousePointer = 11
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Screen.MousePointer = 11
   CII% = CODINT%(TRIM$(Left$(List1.TEXT, 15)))
   If Abs(CII%) > 0 Then
     Master.Edit
     Master!STATUS = 1
     Master.Update
     UPDATMASTER% = 1
     Unload Me
   End If
   '
End Sub

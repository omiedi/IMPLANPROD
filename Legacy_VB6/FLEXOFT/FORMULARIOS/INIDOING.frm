VERSION 5.00
Begin VB.Form INIDOING 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inicia Proceso Documentos de Ingenieria"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "INIDOING"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Call VERJOBID
   DOCUINGE.Show
   Hide
End Sub

VERSION 5.00
Begin VB.Form GS_HNPRINCIPAL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestos Sumas - Honorarios"
   ClientHeight    =   3330
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu archimas 
      Caption         =   "Archivos Maestros"
      Begin VB.Menu masterclientes 
         Caption         =   "Maestro de clientes"
      End
      Begin VB.Menu impuestos 
         Caption         =   "Lista de impuestos"
      End
   End
   Begin VB.Menu honorarios 
      Caption         =   "Honorarios"
      Begin VB.Menu facturahonorarios 
         Caption         =   "Carga de factura de honorarios"
      End
      Begin VB.Menu cobrohonorarios 
         Caption         =   "Cobro de honorarios"
      End
      Begin VB.Menu cuentacorriente 
         Caption         =   "Cuenta corriente"
      End
   End
   Begin VB.Menu consultas 
      Caption         =   "Consultas y Listados"
   End
End
Attribute VB_Name = "GS_HNPRINCIPAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub facturahonorarios_Click()

    GS_HNFACTURA.Show 1
    
End Sub

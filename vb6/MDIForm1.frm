VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   8790
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15285
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuConsultar 
      Caption         =   "Consultar"
   End
   Begin VB.Menu mnuInserir 
      Caption         =   "Inserir"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuConsultar_Click()
    frmConsultar.Show
End Sub


Private Sub mnuInserir_Click()
    frmInserir.Show
End Sub

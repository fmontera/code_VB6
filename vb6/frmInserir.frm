VERSION 5.00
Begin VB.Form frmInserir 
   Caption         =   "Inserir"
   ClientHeight    =   4230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4230
   ScaleWidth      =   5490
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtCelular 
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtTelefone 
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtEndereco 
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox txtNome 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   3375
      End
      Begin VB.CommandButton cmdInserir 
         Caption         =   "Inserir"
         Height          =   375
         Left            =   3360
         TabIndex        =   1
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblCelular 
         Caption         =   "Celular"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblTelefone 
         Caption         =   "Telefone"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label lblEndereço 
         Caption         =   "Endereço"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblNome 
         Caption         =   "Nome"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmInserir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdInserir_Click()

    If (MsgBox("Deseja inserir o nome ?", vbYesNo) = vbYes) Then
        
         Dim http As MSXML2.XMLHTTP
         Dim result As String
         Dim url As String
         
         'chama API
         Set http = CreateObject("MSXML2.ServerXMLHTTP")
         
         url = "http://localhost:49817/Api/pessoa/CadastrarPessoas/{""Nome"":"""
         url = url & txtNome.Text & """"
         url = url & ",""Endereco"":"""
         url = url & txtEndereco.Text & """"
         url = url & ",""Telefone"":"""
         url = url & txtTelefone.Text & """"
         url = url & ",""Celular"":"""
         url = url & txtCelular.Text & """}"
         'Debug.Print url
         
         http.Open "POST", url, False
         http.send
        
         webrequest = http.responseText
        
    End If
    
End Sub

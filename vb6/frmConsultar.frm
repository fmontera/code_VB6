VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmConsultar 
   Caption         =   "Consultar"
   ClientHeight    =   5310
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   8265
   Begin VB.Frame Frame1 
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7815
      Begin VB.CommandButton Command1 
         Caption         =   "teste Conexão"
         Height          =   375
         Left            =   5760
         TabIndex        =   8
         Top             =   4320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Pesquisa por"
         Height          =   1455
         Left            =   360
         TabIndex        =   3
         Top             =   3240
         Width           =   4935
         Begin VB.TextBox txtTelefone 
            Height          =   375
            Left            =   1080
            TabIndex        =   6
            Top             =   840
            Width           =   3255
         End
         Begin VB.TextBox txtNome 
            Height          =   375
            Left            =   1080
            TabIndex        =   4
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label lblTelefone 
            Caption         =   "Telefone"
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   900
            Width           =   735
         End
         Begin VB.Label lblNome 
            Caption         =   "Nome"
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   300
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdPesquisar 
         Caption         =   "Pesquisar"
         Height          =   375
         Left            =   5760
         TabIndex        =   1
         Top             =   3360
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid grdDados 
         Height          =   2655
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4683
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
      End
   End
End
Attribute VB_Name = "frmConsultar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SeparaLinha() As String
Dim SeparaColuna() As String
Dim info As String


Private Sub Form_Load()

    grdDados.Clear
    grdDados.Row = 0
    grdDados.Col = 0
    grdDados.Text = "Nome"
    
    grdDados.Row = 0
    grdDados.Col = 1
    grdDados.Text = "Endereço"
    
    grdDados.Row = 0
    grdDados.Col = 2
    grdDados.Text = "Telefone"
    
    grdDados.Row = 0
    grdDados.Col = 3
    grdDados.Text = "Celular"
    
End Sub


Private Sub cmdPesquisar_Click()

    Dim http As MSXML2.XMLHTTP
    Dim result As String
    Dim url As String
    Dim i As Integer
    
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    
    If (txtNome.Text = "") And (txtTelefone.Text = "") Then
    
        url = "http://localhost:49817//Api/pessoa/ConsultarPessoas"
        http.Open "GET", url, False
        http.send
        
        SeparaLinha = Split(http.responseText, "{""Nome""")
        
        Call preencheGrid
    
    Else
    
        If (txtNome.Text <> "") Then
        
            url = "http://localhost:49817//Api/pessoa/ConsultarPessoasPorNome/" & txtNome.Text
            http.Open "GET", url, False
            http.send
                    
            webrequest = http.responseText
            SeparaLinha = Split(http.responseText, "{""Nome""")
            
            Call preencheGrid
        
        Else
        
            If (txtTelefone.Text <> "") Then
            
                url = "http://localhost:49817//Api/pessoa/ConsultarPessoasPorTelefone/" & txtTelefone.Text
                http.Open "GET", url, False
                http.send
                        
                webrequest = http.responseText
                SeparaLinha = Split(http.responseText, "{""Nome""")
                
                Call preencheGrid
            
            End If
            
        End If
        
    End If
    
    'teste
    'Dim objWepApi As New MSOSOAPLib30.SoapClient30
    'objWepApi.MSSoapInit ("http://localhost:49817/Api/pessoa/ConsultarPessoas")
    
End Sub

Private Sub preencheGrid()
    
    grdDados.Clear
    grdDados.Rows = 2
    
    grdDados.Clear
    grdDados.Row = 0
    grdDados.Col = 0
    grdDados.Text = "Nome"
    
    grdDados.Row = 0
    grdDados.Col = 1
    grdDados.Text = "Endereço"
    
    grdDados.Row = 0
    grdDados.Col = 2
    grdDados.Text = "Telefone"
    
    grdDados.Row = 0
    grdDados.Col = 3
    grdDados.Text = "Celular"
    
    For i = 1 To UBound(SeparaLinha)
            
        'retira caracteres do retorno
        info = Replace(Replace(Replace(Replace(SeparaLinha(i), """", ""), ":", ""), "}", ""), "]", "")
        
        'retira nome dos campos
        info = Replace(Replace(Replace(info, "Endereco", ""), "Telefone", ""), "Celular", "")
        
        SeparaColuna() = Split(info, ":")
        valor = Split(SeparaColuna(j), ",")
        
        grdDados.TextMatrix(i, 0) = valor(0)
        grdDados.TextMatrix(i, 1) = valor(1)
        grdDados.TextMatrix(i, 2) = valor(2)
        grdDados.TextMatrix(i, 3) = valor(3)
        grdDados.Rows = grdDados.Rows + 1
        
    Next i

End Sub



Private Sub Command1_Click()
    
    Dim Conn As New ADODB.Connection
    Dim RS As New ADODB.Recordset
        
    Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\projeto\agenda.mdb;Persist Security Info=False"
    
    Conn.Open
    
    Set RS = Conn.Execute("select Nome,Endereco,Telefone,Celular from pessoas where nome like '" & txtNome.Text & "%'")
      
    If Not RS.EOF Then
        MsgBox RS!celular
    End If
    
End Sub



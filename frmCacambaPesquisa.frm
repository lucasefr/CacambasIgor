VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCacambaGerenciamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Localiza Caçamba"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16320
   Icon            =   "frmCacambaPesquisa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   16320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid flxPesquisa 
      Height          =   7695
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   16335
      _ExtentX        =   28813
      _ExtentY        =   13573
      _Version        =   393216
      BackColorBkg    =   12648447
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   735
      Left            =   15120
      TabIndex        =   0
      Top             =   7800
      Width           =   1095
   End
End
Attribute VB_Name = "frmCacambaGerenciamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum colGridCliente
        idCacamba = 0
        NumeroCacamba = 1
        Disponivel = 2
        Cliente = 3
        NumeroPedidoLocacao = 4
        DataLocacao = 5
        Situacao = 6
        Endereco = 7
        Numero = 8
        Bairro = 9
        Cidade = 10
        Valor = 11
End Enum

Private Sub carregaGridCacambas()
        
        Dim strSQL                              As String
        Dim rsCacamba                           As ADODB.Recordset
        
        geraGridPesquisa
        
        strSQL = Empty
        strSQL = strSQL & " SELECT" & vbCrLf
        strSQL = strSQL & " TB1.idCacambas" & vbCrLf
        strSQL = strSQL & " , TB1.Numero" & vbCrLf
        strSQL = strSQL & " , (SELECT pedidolocacao.NumeroControle FROM pedidolocacao  WHERE TB1.idCacambas = pedidolocacao.idCacambas AND pedidoLocacao.dataRealRetirada is null ORDER BY idPedidoLocacao DESC LIMIT 1) as NumeroPedido" & vbCrLf
        strSQL = strSQL & " , (SELECT pedidolocacao.DataLocacao FROM pedidolocacao WHERE TB1.idCacambas = pedidolocacao.idCacambas AND pedidoLocacao.dataRealRetirada is null ORDER BY idPedidoLocacao DESC LIMIT 1) as DataLocacao" & vbCrLf
        strSQL = strSQL & " , (SELECT Clientes.Nome FROM pedidolocacao INNER JOIN Clientes ON PedidoLocacao.IdClientes = Clientes.IdClientes WHERE TB1.idCacambas = pedidolocacao.idCacambas AND pedidoLocacao.dataRealRetirada is null ORDER BY idPedidoLocacao DESC LIMIT 1) as NomeCliente" & vbCrLf
        strSQL = strSQL & " , (SELECT pedidolocacao.Endereco FROM pedidolocacao WHERE TB1.idCacambas = pedidolocacao.idCacambas AND  pedidoLocacao.dataRealRetirada is null ORDER BY idPedidoLocacao DESC LIMIT 1) as Endereco" & vbCrLf
        strSQL = strSQL & " , (SELECT pedidolocacao.EnderecoNumero FROM pedidolocacao WHERE TB1.idCacambas = pedidolocacao.idCacambas AND  pedidoLocacao.dataRealRetirada is null ORDER BY idPedidoLocacao DESC LIMIT 1) as EnderecoNumero" & vbCrLf
        strSQL = strSQL & " , (SELECT pedidolocacao.Bairro FROM pedidolocacao WHERE TB1.idCacambas = pedidolocacao.idCacambas AND pedidoLocacao.dataRealRetirada is null ORDER BY idPedidoLocacao DESC LIMIT 1)  as Bairro" & vbCrLf
        strSQL = strSQL & " , (SELECT pedidolocacao.Cidade FROM pedidolocacao WHERE TB1.idCacambas = pedidolocacao.idCacambas AND pedidoLocacao.dataRealRetirada is null ORDER BY idPedidoLocacao DESC LIMIT 1) as Cidade" & vbCrLf
        strSQL = strSQL & " , (SELECT pedidolocacao.ValorTotal FROM pedidolocacao WHERE TB1.idCacambas = pedidolocacao.idCacambas AND pedidoLocacao.dataRealRetirada is null ORDER BY idPedidoLocacao DESC LIMIT 1) as ValorTotal" & vbCrLf
        strSQL = strSQL & " , (SELECT pedidolocacao.Situacao FROM pedidolocacao WHERE TB1.idCacambas = pedidolocacao.idCacambas AND pedidoLocacao.dataRealRetirada is null ORDER BY idPedidoLocacao DESC LIMIT 1) as Situacao" & vbCrLf
        strSQL = strSQL & " FROM cacambas TB1"
        strSQL = strSQL & " ORDER BY TB1.idCacambas"
  
        Set rsCacamba = New ADODB.Recordset
        
        With flxPesquisa
                rsCacamba.Open strSQL, dbCacamba, adOpenStatic, adLockReadOnly
                While Not rsCacamba.EOF
                        .AddItem ""
                        .TextMatrix(.Rows - 1, colGridCliente.idCacamba) = NVL(rsCacamba!idCacambas, "")
                        .TextMatrix(.Rows - 1, colGridCliente.Cliente) = NVL(rsCacamba!NomeCliente, "")
                        .TextMatrix(.Rows - 1, colGridCliente.Disponivel) = IIf(NVL(rsCacamba!NumeroPedido, "") = "", "SIM", "NÃO")
                        .TextMatrix(.Rows - 1, colGridCliente.NumeroCacamba) = NVL(rsCacamba!Numero, "")
                        .TextMatrix(.Rows - 1, colGridCliente.NumeroPedidoLocacao) = NVL(rsCacamba!NumeroPedido, "")
                        .TextMatrix(.Rows - 1, colGridCliente.DataLocacao) = NVL(rsCacamba!DataLocacao, "")
                        .TextMatrix(.Rows - 1, colGridCliente.Valor) = Format(NVL(rsCacamba!ValorTotal, "0"), "###,##0.00")
                        .TextMatrix(.Rows - 1, colGridCliente.Situacao) = NVL(rsCacamba!Situacao, "")
                        .TextMatrix(.Rows - 1, colGridCliente.Endereco) = NVL(rsCacamba!Endereco, "")
                        .TextMatrix(.Rows - 1, colGridCliente.Numero) = NVL(rsCacamba!EnderecoNumero, "")
                        .TextMatrix(.Rows - 1, colGridCliente.Bairro) = NVL(rsCacamba!Bairro, "")
                        .TextMatrix(.Rows - 1, colGridCliente.Cidade) = NVL(rsCacamba!Cidade, "")
                        
                        rsCacamba.MoveNext
                Wend
        End With
        
        
End Sub


Private Sub geraGridPesquisa()
        
        With flxPesquisa
                
                .Rows = 1
                .Cols = 12
                .FixedCols = 0
                .TextMatrix(0, colGridCliente.idCacamba) = "Codigo"
                .TextMatrix(0, colGridCliente.NumeroCacamba) = "Nº Caçamba"
                .TextMatrix(0, colGridCliente.Cliente) = "Cliente"
                .TextMatrix(0, colGridCliente.Disponivel) = "Disponível"
                .TextMatrix(0, colGridCliente.NumeroPedidoLocacao) = "Nº Pedido"
                .TextMatrix(0, colGridCliente.DataLocacao) = "Data Locação"
                .TextMatrix(0, colGridCliente.Situacao) = "Situacao"
                .TextMatrix(0, colGridCliente.Endereco) = "Endereço"
                .TextMatrix(0, colGridCliente.Numero) = "Numero"
                .TextMatrix(0, colGridCliente.Bairro) = "Bairro"
                .TextMatrix(0, colGridCliente.Cidade) = "Cidade"
                .TextMatrix(0, colGridCliente.Valor) = "Total"
                
                .ColWidth(colGridCliente.idCacamba) = 800
                .ColWidth(colGridCliente.Disponivel) = 900
                .ColWidth(colGridCliente.Cliente) = 3000
                .ColWidth(colGridCliente.NumeroCacamba) = 1000
                .ColWidth(colGridCliente.NumeroPedidoLocacao) = 1000
                .ColWidth(colGridCliente.DataLocacao) = 1300
                .ColWidth(colGridCliente.Situacao) = 1000
                .ColWidth(colGridCliente.Endereco) = 3500
                .ColWidth(colGridCliente.Numero) = 800
                .ColWidth(colGridCliente.Bairro) = 2500
                .ColWidth(colGridCliente.Cidade) = 2500
                .ColWidth(colGridCliente.Valor) = 1500
                
                
                .SelectionMode = flexSelectionByRow
                .GridLines = flexGridInset
        End With
        
End Sub


Private Sub cmdSair_Click()
        Unload Me
End Sub


Private Sub Form_Load()
        carregaGridCacambas
End Sub

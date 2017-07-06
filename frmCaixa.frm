VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCaixa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caixa"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13905
   Icon            =   "frmCaixa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   13905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid flxPedidoLocacaoSemCaixa 
      Height          =   2535
      Left            =   240
      TabIndex        =   19
      Top             =   1680
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   4471
      _Version        =   393216
      BackColorBkg    =   12648447
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "Localizar"
      Height          =   735
      Left            =   5640
      TabIndex        =   15
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   735
      Left            =   6825
      TabIndex        =   14
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   735
      Left            =   9180
      TabIndex        =   13
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Height          =   735
      Left            =   10365
      TabIndex        =   12
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   735
      Left            =   11535
      TabIndex        =   11
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   735
      Left            =   12720
      TabIndex        =   10
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   735
      Left            =   7995
      TabIndex        =   9
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Frame fraPedidoLocacao 
      Height          =   5895
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   13695
      Begin MSFlexGridLib.MSFlexGrid flxPedidoLocacaoCaixa 
         Height          =   2295
         Left            =   120
         TabIndex        =   20
         Top             =   3480
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   4048
         _Version        =   393216
         BackColorBkg    =   4
      End
      Begin VB.CommandButton cmdExcluirPedido 
         Caption         =   "Excluir Pedido"
         Height          =   735
         Left            =   12840
         TabIndex        =   8
         Top             =   4200
         Width           =   735
      End
      Begin VB.CommandButton cmdIncluirPedido 
         Caption         =   "Incluir Pedido"
         Height          =   735
         Left            =   12840
         TabIndex        =   7
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Pedidos de Locação do Caixa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   12615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         Caption         =   "Pedidos de Locação Sem Caixa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   12615
      End
   End
   Begin VB.Frame fraCaixa 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13695
      Begin MSMask.MaskEdBox mskValorCaixa 
         Height          =   375
         Left            =   11760
         TabIndex        =   3
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDataCaixa 
         Height          =   375
         Left            =   2400
         TabIndex        =   2
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtNumeroCaixa 
         Height          =   375
         Left            =   120
         MaxLength       =   10
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Valor do Caixa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11880
         TabIndex        =   18
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Data do Caixa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblCNPJCPF 
         Caption         =   "Número do Caixa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vOperacao               As operacao
Private vIdCaixa                As Integer

Private Enum colGridPedidoLocacao
        idPedidoLocacao = 0
        NumeroPedidoLocacao = 1
        NumeroCacamba = 2
        nome = 3
        DataLocacao = 4
        dataRealRetirada = 5
        ValorServico = 6
        ValorDesconto = 7
        ValorTotal = 8
        Situacao = 9
        FormaPagamento = 10
End Enum


Private Sub cmdExcluirPedido_Click()
        excluiPedidoLocacao
End Sub

Private Sub cmdIncluirPedido_Click()
        inclirPedidoLocacao
End Sub

Private Sub txtNumeroCaixa_KeyPress(KeyAscii As Integer)
                        
        KeyAscii = SoNumeros(KeyAscii)
        
        If KeyAscii = 0 Then
                Exit Sub
        End If

End Sub

Private Sub cmdAlterar_Click()
        controlaBotoes Me, operacao.Inclusao
        fraCaixa.Enabled = True
        fraPedidoLocacao.Enabled = True
        vOperacao = Alteracao
        carregaGridPedidoLocacaoSemCaixa
End Sub

Private Sub cmdCancelar_Click()
        controlaBotoes Me, operacao.consulta
        fraCaixa.Enabled = False
        fraPedidoLocacao.Enabled = False
        vOperacao = consulta
        geraGridPedidoLocacaoSemCaixa
        carregaCaixa 0
End Sub

Private Sub cmdExcluir_Click()
        Dim strSQL                      As String
        
        If MsgBox("Confirma exlusão?", vbQuestion + vbYesNo + vbDefaultButton2, "Caçambas") = vbNo Then
                Exit Sub
        End If
        
        strSQL = Empty
        strSQL = strSQL & " UPDATE " & vbCrLf
        strSQL = strSQL & " PedidoLocacao " & vbCrLf
        strSQL = strSQL & " SET idCaixa = null"
        strSQL = strSQL & " WHERE idCaixa = " & vIdCaixa
        
        dbCacamba.Execute strSQL
        
        strSQL = Empty
        strSQL = strSQL & "DELETE FROM Caixa WHERE idCaixa = " & vIdCaixa
        
        dbCacamba.Execute strSQL
        
        controlaBotoes Me, operacao.consulta
        fraCaixa.Enabled = False
        fraPedidoLocacao.Enabled = False
        carregaCaixa 0
        vOperacao = consulta
        
End Sub

Private Sub cmdGravar_Click()
        Dim strSQL                      As String
        Dim rsCaixa                     As ADODB.Recordset
        
        If Not validaGravacao Then Exit Sub
        
        strSQL = Empty
        
        If vOperacao = Inclusao Then
        
                strSQL = strSQL & " INSERT INTO Caixa " & vbCrLf
                strSQL = strSQL & " (NumeroCaixa " & vbCrLf
                strSQL = strSQL & " ,DataCaixa " & vbCrLf
                strSQL = strSQL & ") VALUES ( "
                strSQL = strSQL & txtNumeroCaixa.Text & vbCrLf
                strSQL = strSQL & ",'" & Format(mskDataCaixa.FormattedText, "yyyymmdd") & "'" & vbCrLf
                strSQL = strSQL & ")" & vbCrLf
                
                dbCacamba.Execute strSQL
                
                strSQL = Empty
                strSQL = strSQL & " SELECT idCaixa FROM Caixa ORDER BY idCaixa DESC LIMIT 1"
                
                
                Set rsCaixa = New ADODB.Recordset
                
                With rsCaixa
                        .Open strSQL, dbCacamba, adOpenStatic, adLockReadOnly
                        If Not .EOF Then
                                gravaPedidoLocacaoCaixa !idCaixa
                        End If
                End With
        
        Else
                        
                strSQL = strSQL & " UPDATE Caixa " & vbCrLf
                strSQL = strSQL & " SET " & vbCrLf
                strSQL = strSQL & " NumeroCaixa = " & txtNumeroCaixa.Text & vbCrLf
                strSQL = strSQL & " ,DataCaixa = '" & Format(mskDataCaixa.FormattedText, "yyyymmdd") & "'" & vbCrLf
                strSQL = strSQL & " WHERE idCaixa = " & vIdCaixa
                
                dbCacamba.Execute strSQL
                
                gravaPedidoLocacaoCaixa vIdCaixa
        
        End If
        
        
        
        
        
        

        carregaCaixa 0
        controlaBotoes Me, operacao.consulta
        vOperacao = consulta
        fraCaixa.Enabled = False
        fraPedidoLocacao.Enabled = False
        geraGridPedidoLocacaoSemCaixa
End Sub

Private Sub cmdIncluir_Click()
        controlaBotoes Me, operacao.Inclusao
        fraCaixa.Enabled = True
        fraPedidoLocacao.Enabled = True
        limpaCampos
        vIdCaixa = 0
        vOperacao = Inclusao
        carregaGridPedidoLocacaoSemCaixa
        geraGridPedidoLocacaoComCaixa
End Sub

Private Sub cmdLocalizar_Click()
        Dim strSQL                      As String
        Dim vColunas                    As String
        
        strSQL = Empty
        strSQL = strSQL & " SELECT " & vbCrLf
        strSQL = strSQL & " Caixa.idCaixa " & vbCrLf
        strSQL = strSQL & " ,Caixa.NumeroCaixa " & vbCrLf
        strSQL = strSQL & " ,Caixa.DataCaixa" & vbCrLf
        strSQL = strSQL & " ,(SELECT SUM(ValorTotal) as Valor FROM PedidoLocacao WHERE PedidoLocacao.idCaixa = caixa.idCaixa )" & vbCrLf
        strSQL = strSQL & " FROM Caixa " & vbCrLf
        strSQL = strSQL & " GROUP BY Caixa.idCaixa, Caixa.NumeroCaixa, Caixa.DataCaixa " & vbCrLf
        strSQL = strSQL & " ORDER BY idCaixa DESC "
        
        vColunas = "Código,Número Caixa, Data Caixa, Valor Caixa"
        
        frmPesquisa.carregaGridPesquisa 4, strSQL, vColunas, 1
        frmPesquisa.Show vbModal
        
        If vgRetornoConsulta <> 0 Then
                vIdCaixa = vgRetornoConsulta
                carregaCaixa vIdCaixa
        End If
        
End Sub

Private Sub cmdSair_Click()
        Unload Me
End Sub

Private Sub Form_Load()
        controlaBotoes Me, operacao.consulta
        fraCaixa.Enabled = False
        fraPedidoLocacao.Enabled = False
        carregaCaixa
End Sub

Private Sub limpaCampos()
        
        txtNumeroCaixa.Text = ""
        mskDataCaixa.Text = "__/__/____"
        mskValorCaixa.Text = 0
        
End Sub

Private Function validaGravacao() As Boolean

        validaGravacao = False
        
        If Trim(txtNumeroCaixa.Text) = "" Then
                MsgBox "Informe o Numero!", vbInformation, "Caçambas"
                txtNumeroCaixa.SetFocus
                Exit Function
        End If
        
        If mskDataCaixa.Text = "__/__/____" Then
                MsgBox "Infomre a Data do Caixa!", vbInformation, "Caçambas"
                Exit Function
        End If
        
        If flxPedidoLocacaoCaixa.Rows = 1 Then
                MsgBox "Nenhum pedido de locação foi selecionado!", vbInformation, "Caçambas"
                Exit Function
        End If
        
        validaGravacao = True
        
End Function

Private Sub carregaCaixa(Optional pIdCaixa As Integer)

        Dim strSQL                      As String
        Dim rsCaixa                     As ADODB.Recordset
        
        strSQL = Empty
        If pIdCaixa = 0 Then
                strSQL = strSQL & " SELECT  " & vbCrLf
                strSQL = strSQL & " caixa.idCaixa " & vbCrLf
                strSQL = strSQL & " , caixa.NumeroCaixa " & vbCrLf
                strSQL = strSQL & " , caixa.dataCaixa " & vbCrLf
                strSQL = strSQL & " , (SELECT SUM(ValorTotal) FROM PedidoLocacao WHERE PedidoLocacao.idCaixa = Caixa.idCaixa) as ValorCaixa" & vbCrLf
                strSQL = strSQL & " FROM Caixa " & vbCrLf
                strSQL = strSQL & " ORDER BY idCaixa DESC LIMIT 1"
                
        Else
                strSQL = strSQL & " SELECT  " & vbCrLf
                strSQL = strSQL & " caixa.idCaixa " & vbCrLf
                strSQL = strSQL & " , caixa.NumeroCaixa " & vbCrLf
                strSQL = strSQL & " , caixa.dataCaixa " & vbCrLf
                strSQL = strSQL & " , (SELECT SUM(ValorTotal) FROM PedidoLocacao WHERE PedidoLocacao.idCaixa = Caixa.idCaixa) as ValorCaixa" & vbCrLf
                strSQL = strSQL & " FROM Caixa " & vbCrLf
                strSQL = strSQL & " WHERE idCaixa = " & pIdCaixa
        End If
        
        Set rsCaixa = New ADODB.Recordset
        
        With rsCaixa
                .Open strSQL, dbCacamba, adOpenStatic, adLockReadOnly
                If Not .EOF Then
                        vIdCaixa = !idCaixa
                        txtNumeroCaixa.Text = !NumeroCaixa
                        mskDataCaixa.Text = !dataCaixa
                        mskValorCaixa = NVL(!valorCaixa, 0)
                        carregaGridPedidoLocacaoComCaixa
                        geraGridPedidoLocacaoSemCaixa
                End If
        End With
        
        
End Sub


Private Sub geraGridPedidoLocacaoComCaixa()
        
        With flxPedidoLocacaoCaixa
                .Rows = 1
                .Cols = 11
                .FixedCols = 0
                
                .TextMatrix(0, colGridPedidoLocacao.idPedidoLocacao) = ""
                .TextMatrix(0, colGridPedidoLocacao.NumeroPedidoLocacao) = "Pedido"
                .TextMatrix(0, colGridPedidoLocacao.NumeroCacamba) = "Caçamba"
                .TextMatrix(0, colGridPedidoLocacao.nome) = "Cliente"
                .TextMatrix(0, colGridPedidoLocacao.DataLocacao) = "Data Locação"
                .TextMatrix(0, colGridPedidoLocacao.dataRealRetirada) = "Data Retirada"
                .TextMatrix(0, colGridPedidoLocacao.ValorServico) = "Valor"
                .TextMatrix(0, colGridPedidoLocacao.ValorDesconto) = "Desconto"
                .TextMatrix(0, colGridPedidoLocacao.ValorTotal) = "Total"
                .TextMatrix(0, colGridPedidoLocacao.Situacao) = "Situação"
                .TextMatrix(0, colGridPedidoLocacao.FormaPagamento) = "Forma Pagamento"
                
                .ColWidth(colGridPedidoLocacao.idPedidoLocacao) = 0
                .ColWidth(colGridPedidoLocacao.NumeroPedidoLocacao) = 850
                .ColWidth(colGridPedidoLocacao.NumeroCacamba) = 900
                .ColWidth(colGridPedidoLocacao.nome) = 2500
                .ColWidth(colGridPedidoLocacao.DataLocacao) = 1200
                .ColWidth(colGridPedidoLocacao.dataRealRetirada) = 1200
                .ColWidth(colGridPedidoLocacao.ValorServico) = 900
                .ColWidth(colGridPedidoLocacao.ValorDesconto) = 900
                .ColWidth(colGridPedidoLocacao.ValorTotal) = 700
                .ColWidth(colGridPedidoLocacao.Situacao) = 1600
                .ColWidth(colGridPedidoLocacao.FormaPagamento) = 1400
                
                
                
                .SelectionMode = flexSelectionByRow
                .GridLines = flexGridInset
        End With
        
        
End Sub

Private Sub geraGridPedidoLocacaoSemCaixa()


        With flxPedidoLocacaoSemCaixa
                .Rows = 1
                .Cols = 11
                .FixedCols = 0
                
                .TextMatrix(0, colGridPedidoLocacao.idPedidoLocacao) = ""
                .TextMatrix(0, colGridPedidoLocacao.NumeroPedidoLocacao) = "Pedido"
                .TextMatrix(0, colGridPedidoLocacao.NumeroCacamba) = "Caçamba"
                .TextMatrix(0, colGridPedidoLocacao.nome) = "Cliente"
                .TextMatrix(0, colGridPedidoLocacao.DataLocacao) = "Data Locação"
                .TextMatrix(0, colGridPedidoLocacao.dataRealRetirada) = "Data Retirada"
                .TextMatrix(0, colGridPedidoLocacao.ValorServico) = "Valor"
                .TextMatrix(0, colGridPedidoLocacao.ValorDesconto) = "Desconto"
                .TextMatrix(0, colGridPedidoLocacao.ValorTotal) = "Total"
                .TextMatrix(0, colGridPedidoLocacao.Situacao) = "Situação"
                .TextMatrix(0, colGridPedidoLocacao.FormaPagamento) = "Forma Pagamento"
                
                .ColWidth(colGridPedidoLocacao.idPedidoLocacao) = 0
                .ColWidth(colGridPedidoLocacao.NumeroPedidoLocacao) = 850
                .ColWidth(colGridPedidoLocacao.NumeroCacamba) = 900
                .ColWidth(colGridPedidoLocacao.nome) = 2500
                .ColWidth(colGridPedidoLocacao.DataLocacao) = 1200
                .ColWidth(colGridPedidoLocacao.dataRealRetirada) = 1200
                .ColWidth(colGridPedidoLocacao.ValorServico) = 900
                .ColWidth(colGridPedidoLocacao.ValorDesconto) = 900
                .ColWidth(colGridPedidoLocacao.ValorTotal) = 700
                .ColWidth(colGridPedidoLocacao.Situacao) = 1600
                .ColWidth(colGridPedidoLocacao.FormaPagamento) = 1400
                
                
                
                .SelectionMode = flexSelectionByRow
                .GridLines = flexGridInset
        End With
End Sub

Private Sub carregaGridPedidoLocacaoSemCaixa()

        Dim strSQL                              As String
        Dim rsPedidoLocacao                     As ADODB.Recordset
        
        geraGridPedidoLocacaoSemCaixa
        
        strSQL = Empty
        strSQL = strSQL & " SELECT " & vbCrLf
        strSQL = strSQL & " PedidoLocacao.idPedidoLocacao " & vbCrLf
        strSQL = strSQL & " ,Clientes.Nome " & vbCrLf
        strSQL = strSQL & " ,Cacambas.Numero as NumeroCacamba " & vbCrLf
        strSQL = strSQL & " ,NumeroControle " & vbCrLf
        strSQL = strSQL & " ,DataLocacao " & vbCrLf
        strSQL = strSQL & " ,DataRealRetirada " & vbCrLf
        strSQL = strSQL & " ,ValorServico " & vbCrLf
        strSQL = strSQL & " ,ValorDesconto " & vbCrLf
        strSQL = strSQL & " ,ValorTotal " & vbCrLf
        strSQL = strSQL & " ,Situacao " & vbCrLf
        strSQL = strSQL & " ,FormaPagamento " & vbCrLf
        strSQL = strSQL & " FROM PedidoLocacao " & vbCrLf
        strSQL = strSQL & " INNER JOIN Cacambas ON PedidoLocacao.idCacambas = Cacambas.idCacambas" & vbCrLf
        strSQL = strSQL & " INNER JOIN Clientes ON PedidoLocacao.idClientes = Clientes.idClientes" & vbCrLf
        strSQL = strSQL & " WHERE PedidoLocacao.Situacao IN ('Pago na Entrega', 'Pago na Retirada') " & vbCrLf
        strSQL = strSQL & " AND PedidoLocacao.idCaixa IS NULL"
        
        With flxPedidoLocacaoSemCaixa
                Set rsPedidoLocacao = New ADODB.Recordset
                rsPedidoLocacao.Open strSQL, dbCacamba, adOpenStatic, adLockReadOnly
                While Not rsPedidoLocacao.EOF
                        .AddItem ""
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.idPedidoLocacao) = NVL(rsPedidoLocacao!idPedidoLocacao, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.NumeroPedidoLocacao) = NVL(rsPedidoLocacao!numeroControle, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.NumeroCacamba) = NVL(rsPedidoLocacao!NumeroCacamba, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.nome) = NVL(rsPedidoLocacao!nome, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.DataLocacao) = NVL(rsPedidoLocacao!DataLocacao, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.dataRealRetirada) = NVL(rsPedidoLocacao!dataRealRetirada, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.ValorServico) = Format(NVL(rsPedidoLocacao!ValorServico, "0"), "###,##0.00")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.ValorDesconto) = Format(NVL(rsPedidoLocacao!ValorDesconto, "0"), "###,##0.00")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.ValorTotal) = Format(NVL(rsPedidoLocacao!ValorTotal, "0"), "###,##0.00")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.Situacao) = NVL(rsPedidoLocacao!Situacao, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.FormaPagamento) = NVL(rsPedidoLocacao!FormaPagamento, "")
                        
                        rsPedidoLocacao.MoveNext
                Wend
        End With
End Sub


Private Sub carregaGridPedidoLocacaoComCaixa()

        Dim strSQL                              As String
        Dim rsPedidoLocacao                     As ADODB.Recordset
        
        geraGridPedidoLocacaoComCaixa
        
        strSQL = Empty
        strSQL = strSQL & " SELECT " & vbCrLf
        strSQL = strSQL & " PedidoLocacao.idPedidoLocacao " & vbCrLf
        strSQL = strSQL & " ,Clientes.Nome " & vbCrLf
        strSQL = strSQL & " ,Cacambas.Numero as NumeroCacamba " & vbCrLf
        strSQL = strSQL & " ,NumeroControle " & vbCrLf
        strSQL = strSQL & " ,DataLocacao " & vbCrLf
        strSQL = strSQL & " ,DataRealRetirada " & vbCrLf
        strSQL = strSQL & " ,ValorServico " & vbCrLf
        strSQL = strSQL & " ,ValorDesconto " & vbCrLf
        strSQL = strSQL & " ,ValorTotal " & vbCrLf
        strSQL = strSQL & " ,Situacao " & vbCrLf
        strSQL = strSQL & " ,FormaPagamento " & vbCrLf
        strSQL = strSQL & " FROM PedidoLocacao " & vbCrLf
        strSQL = strSQL & " INNER JOIN Cacambas ON PedidoLocacao.idCacambas = Cacambas.idCacambas" & vbCrLf
        strSQL = strSQL & " INNER JOIN Clientes ON PedidoLocacao.idClientes = Clientes.idClientes" & vbCrLf
        strSQL = strSQL & " WHERE PedidoLocacao.idCaixa = " & vIdCaixa
        
        With flxPedidoLocacaoCaixa
                Set rsPedidoLocacao = New ADODB.Recordset
                rsPedidoLocacao.Open strSQL, dbCacamba, adOpenStatic, adLockReadOnly
                While Not rsPedidoLocacao.EOF
                        .AddItem ""
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.idPedidoLocacao) = NVL(rsPedidoLocacao!idPedidoLocacao, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.NumeroPedidoLocacao) = NVL(rsPedidoLocacao!numeroControle, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.NumeroCacamba) = NVL(rsPedidoLocacao!NumeroCacamba, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.nome) = NVL(rsPedidoLocacao!nome, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.DataLocacao) = NVL(rsPedidoLocacao!DataLocacao, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.dataRealRetirada) = NVL(rsPedidoLocacao!dataRealRetirada, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.ValorServico) = Format(NVL(rsPedidoLocacao!ValorServico, "0"), "###,##0.00")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.ValorDesconto) = Format(NVL(rsPedidoLocacao!ValorDesconto, "0"), "###,##0.00")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.ValorTotal) = Format(NVL(rsPedidoLocacao!ValorTotal, "0"), "###,##0.00")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.Situacao) = NVL(rsPedidoLocacao!Situacao, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.FormaPagamento) = NVL(rsPedidoLocacao!FormaPagamento, "")
                        
                        rsPedidoLocacao.MoveNext
                Wend
        End With
End Sub


Private Sub inclirPedidoLocacao()
        
        With flxPedidoLocacaoSemCaixa
                If .Rows = 1 Then
                        MsgBox "Nenhum Pedido de Locação Selecionado!", vbInformation, "Caçambas"
                        Exit Sub
                End If
                
                flxPedidoLocacaoCaixa.AddItem ""
                flxPedidoLocacaoCaixa.TextMatrix(flxPedidoLocacaoCaixa.Rows - 1, colGridPedidoLocacao.idPedidoLocacao) = .TextMatrix(.Row, colGridPedidoLocacao.idPedidoLocacao)
                flxPedidoLocacaoCaixa.TextMatrix(flxPedidoLocacaoCaixa.Rows - 1, colGridPedidoLocacao.NumeroPedidoLocacao) = .TextMatrix(.Row, colGridPedidoLocacao.NumeroPedidoLocacao)
                flxPedidoLocacaoCaixa.TextMatrix(flxPedidoLocacaoCaixa.Rows - 1, colGridPedidoLocacao.NumeroCacamba) = .TextMatrix(.Row, colGridPedidoLocacao.NumeroCacamba)
                flxPedidoLocacaoCaixa.TextMatrix(flxPedidoLocacaoCaixa.Rows - 1, colGridPedidoLocacao.nome) = .TextMatrix(.Row, colGridPedidoLocacao.nome)
                flxPedidoLocacaoCaixa.TextMatrix(flxPedidoLocacaoCaixa.Rows - 1, colGridPedidoLocacao.DataLocacao) = .TextMatrix(.Row, colGridPedidoLocacao.DataLocacao)
                flxPedidoLocacaoCaixa.TextMatrix(flxPedidoLocacaoCaixa.Rows - 1, colGridPedidoLocacao.dataRealRetirada) = .TextMatrix(.Row, colGridPedidoLocacao.dataRealRetirada)
                flxPedidoLocacaoCaixa.TextMatrix(flxPedidoLocacaoCaixa.Rows - 1, colGridPedidoLocacao.ValorServico) = .TextMatrix(.Row, colGridPedidoLocacao.ValorServico)
                flxPedidoLocacaoCaixa.TextMatrix(flxPedidoLocacaoCaixa.Rows - 1, colGridPedidoLocacao.ValorDesconto) = .TextMatrix(.Row, colGridPedidoLocacao.ValorDesconto)
                flxPedidoLocacaoCaixa.TextMatrix(flxPedidoLocacaoCaixa.Rows - 1, colGridPedidoLocacao.ValorTotal) = .TextMatrix(.Row, colGridPedidoLocacao.ValorTotal)
                flxPedidoLocacaoCaixa.TextMatrix(flxPedidoLocacaoCaixa.Rows - 1, colGridPedidoLocacao.Situacao) = .TextMatrix(.Row, colGridPedidoLocacao.Situacao)
                flxPedidoLocacaoCaixa.TextMatrix(flxPedidoLocacaoCaixa.Rows - 1, colGridPedidoLocacao.FormaPagamento) = .TextMatrix(.Row, colGridPedidoLocacao.FormaPagamento)
                
                If .Rows = 2 Then
                        geraGridPedidoLocacaoSemCaixa
                Else
                        .RemoveItem .Rows - 1
                End If
        End With
        
        calculaValorCaixa
End Sub


Private Sub excluiPedidoLocacao()
        
        With flxPedidoLocacaoCaixa
                If .Rows = 1 Then
                        MsgBox "Nenhum Pedido de Locação Selecionado!", vbInformation, "Caçambas"
                        Exit Sub
                End If
                
                flxPedidoLocacaoSemCaixa.AddItem ""
                flxPedidoLocacaoSemCaixa.TextMatrix(flxPedidoLocacaoSemCaixa.Rows - 1, colGridPedidoLocacao.idPedidoLocacao) = .TextMatrix(.Row, colGridPedidoLocacao.idPedidoLocacao)
                flxPedidoLocacaoSemCaixa.TextMatrix(flxPedidoLocacaoSemCaixa.Rows - 1, colGridPedidoLocacao.NumeroPedidoLocacao) = .TextMatrix(.Row, colGridPedidoLocacao.NumeroPedidoLocacao)
                flxPedidoLocacaoSemCaixa.TextMatrix(flxPedidoLocacaoSemCaixa.Rows - 1, colGridPedidoLocacao.NumeroCacamba) = .TextMatrix(.Row, colGridPedidoLocacao.NumeroCacamba)
                flxPedidoLocacaoSemCaixa.TextMatrix(flxPedidoLocacaoSemCaixa.Rows - 1, colGridPedidoLocacao.nome) = .TextMatrix(.Row, colGridPedidoLocacao.nome)
                flxPedidoLocacaoSemCaixa.TextMatrix(flxPedidoLocacaoSemCaixa.Rows - 1, colGridPedidoLocacao.DataLocacao) = .TextMatrix(.Row, colGridPedidoLocacao.DataLocacao)
                flxPedidoLocacaoSemCaixa.TextMatrix(flxPedidoLocacaoSemCaixa.Rows - 1, colGridPedidoLocacao.dataRealRetirada) = .TextMatrix(.Row, colGridPedidoLocacao.dataRealRetirada)
                flxPedidoLocacaoSemCaixa.TextMatrix(flxPedidoLocacaoSemCaixa.Rows - 1, colGridPedidoLocacao.ValorServico) = .TextMatrix(.Row, colGridPedidoLocacao.ValorServico)
                flxPedidoLocacaoSemCaixa.TextMatrix(flxPedidoLocacaoSemCaixa.Rows - 1, colGridPedidoLocacao.ValorDesconto) = .TextMatrix(.Row, colGridPedidoLocacao.ValorDesconto)
                flxPedidoLocacaoSemCaixa.TextMatrix(flxPedidoLocacaoSemCaixa.Rows - 1, colGridPedidoLocacao.ValorTotal) = .TextMatrix(.Row, colGridPedidoLocacao.ValorTotal)
                flxPedidoLocacaoSemCaixa.TextMatrix(flxPedidoLocacaoSemCaixa.Rows - 1, colGridPedidoLocacao.Situacao) = .TextMatrix(.Row, colGridPedidoLocacao.Situacao)
                flxPedidoLocacaoSemCaixa.TextMatrix(flxPedidoLocacaoSemCaixa.Rows - 1, colGridPedidoLocacao.FormaPagamento) = .TextMatrix(.Row, colGridPedidoLocacao.FormaPagamento)
                
                If .Rows = 2 Then
                        geraGridPedidoLocacaoComCaixa
                Else
                        .RemoveItem .Rows - 1
                End If
        End With
        
        calculaValorCaixa
        
End Sub

Private Sub calculaValorCaixa()
        
        Dim vliAuxiliar                         As Long
        Dim vlcValorCaixa                       As Currency
        
        vlcValorCaixa = 0
        
        With flxPedidoLocacaoCaixa
                For vliAuxiliar = 1 To .Rows - 1
                        vlcValorCaixa = vlcValorCaixa + .TextMatrix(vliAuxiliar, colGridPedidoLocacao.ValorTotal)
                Next
        End With
        
        mskValorCaixa.Text = vlcValorCaixa
End Sub


Private Sub gravaPedidoLocacaoCaixa(pIdCaixa As Integer)
        
        Dim vliAuxiliar                         As Long
        Dim strSQL                              As String
        
        With flxPedidoLocacaoCaixa
                If .Rows = 1 Then
                        Exit Sub
                End If
                
                strSQL = Empty
                strSQL = strSQL & " UPDATE " & vbCrLf
                strSQL = strSQL & " PedidoLocacao " & vbCrLf
                strSQL = strSQL & " SET idCaixa = null"
                strSQL = strSQL & " WHERE idCaixa = " & pIdCaixa
                
                dbCacamba.Execute strSQL
                        
                For vliAuxiliar = 1 To .Rows - 1
                        strSQL = Empty
                        strSQL = strSQL & " UPDATE " & vbCrLf
                        strSQL = strSQL & " PedidoLocacao " & vbCrLf
                        strSQL = strSQL & " SET idCaixa = " & pIdCaixa
                        strSQL = strSQL & " WHERE idPedidoLocacao = " & .TextMatrix(vliAuxiliar, colGridPedidoLocacao.idPedidoLocacao)
                        
                        dbCacamba.Execute strSQL
                Next
        End With
End Sub

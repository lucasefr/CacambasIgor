VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPedidoLocacaoGerenciamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gerenciamento de Pedido de Locação"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16275
   Icon            =   "frmGerenciamentoPedidoLocacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   16275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdPedido 
      Caption         =   "Ir para o Pedido"
      Height          =   735
      Left            =   13320
      TabIndex        =   36
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   16215
      Begin VB.Frame fraTipoPedido 
         Caption         =   "Tipo de Locação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4560
         TabIndex        =   37
         Top             =   1320
         Width           =   3615
         Begin VB.OptionButton optTipoLocacaoTroca 
            Caption         =   "Troca"
            Height          =   195
            Left            =   2280
            TabIndex        =   40
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optTipoLocacaoAluguel 
            Caption         =   "Aluguel"
            Height          =   195
            Left            =   1200
            TabIndex        =   39
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optTipoLocacaoTodos 
            Caption         =   "Todos"
            Height          =   195
            Left            =   240
            TabIndex        =   38
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Data Entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   8280
         TabIndex        =   33
         Top             =   240
         Width           =   3615
         Begin MSMask.MaskEdBox mskDataEntregaInicial 
            Height          =   375
            Left            =   480
            TabIndex        =   5
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDataEntregaFinal 
            Height          =   375
            Left            =   2280
            TabIndex        =   6
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label7 
            Caption         =   "De"
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
            TabIndex        =   35
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label5 
            Caption         =   "Até"
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
            Left            =   1800
            TabIndex        =   34
            Top             =   600
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Data Real Retirada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   8280
         TabIndex        =   30
         Top             =   1440
         Width           =   3615
         Begin MSMask.MaskEdBox mskDataRealRetiradaInicial 
            Height          =   375
            Left            =   480
            TabIndex        =   7
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDataRealRetiradaFinal 
            Height          =   375
            Left            =   2280
            TabIndex        =   8
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label4 
            Caption         =   "De"
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
            TabIndex        =   32
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "Até"
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
            Left            =   1800
            TabIndex        =   31
            Top             =   600
            Width           =   375
         End
      End
      Begin VB.Frame fraDataLocacao 
         Caption         =   "Data de Locação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4560
         TabIndex        =   27
         Top             =   240
         Width           =   3615
         Begin MSMask.MaskEdBox mskDataLocacaoInicial 
            Height          =   375
            Left            =   480
            TabIndex        =   3
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskDataLocacaoFinal 
            Height          =   375
            Left            =   2280
            TabIndex        =   4
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   661
            _Version        =   393216
            MaxLength       =   10
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            Caption         =   "Até"
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
            Left            =   1800
            TabIndex        =   29
            Top             =   600
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "De"
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
            TabIndex        =   28
            Top             =   600
            Width           =   375
         End
      End
      Begin VB.Frame fraSituacao 
         Caption         =   "Situação"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   14160
         TabIndex        =   26
         Top             =   240
         Width           =   1935
         Begin VB.CheckBox chkSituacaoPagoRetirada 
            Caption         =   "Pago Retirada"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1080
            Width           =   1455
         End
         Begin VB.CheckBox chkSituacaoAreceber 
            Caption         =   "A Receber"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox chkSituacaoPagoEntrega 
            Caption         =   "Pago Entrega"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   1455
         End
         Begin VB.CheckBox chkSituacaoCortesia 
            Caption         =   "Cortesia"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CheckBox chkSituacaoDevedor 
            Caption         =   "Devedor"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1800
            Width           =   1215
         End
      End
      Begin VB.Frame fraFormaPagamento 
         Caption         =   "Forma de Pagamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   12000
         TabIndex        =   25
         Top             =   240
         Width           =   2055
         Begin VB.CheckBox chkFormaPagBoleto 
            Caption         =   "Boleto"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CheckBox chkFormaPagCartao 
            Caption         =   "Cartão"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1440
            Width           =   975
         End
         Begin VB.CheckBox chkFormaPagTransferencia 
            Caption         =   "Transferência"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CheckBox chkFormaPagCheque 
            Caption         =   "Cheque"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkFormaPagDinheiro 
            Caption         =   "Dinheiro"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.TextBox txtNumeroControle 
         Height          =   375
         Left            =   120
         MaxLength       =   10
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton cmdPesquisa 
         Caption         =   "Pesquisa"
         Height          =   495
         Left            =   15000
         TabIndex        =   20
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdSelecionaCliente 
         Height          =   375
         Left            =   6120
         Picture         =   "frmGerenciamentoPedidoLocacao.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtCliente 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   21
         Top             =   2160
         Width           =   6375
      End
      Begin VB.Label lblCNPJCPF 
         Caption         =   "Número do Controle"
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
         TabIndex        =   24
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "Alocadas no Cliente"
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
         TabIndex        =   22
         Top             =   1920
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   735
      Left            =   15105
      TabIndex        =   0
      Top             =   7800
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid flxPesquisa 
      Height          =   4215
      Left            =   0
      TabIndex        =   23
      Top             =   3480
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   7435
      _Version        =   393216
      BackColorBkg    =   12648447
      GridColor       =   16777215
   End
End
Attribute VB_Name = "frmPedidoLocacaoGerenciamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vIdClientePedidoLocacao                 As Integer

Private Enum colGridPedidoLocacao
        idPedidoLocacao = 0
        NumeroPedidoLocacao = 1
        TipoPedidoLocacao = 2
        NumeroCacamba = 3
        nome = 4
        cpfcnpj = 5
        DataLocacao = 6
        dataEntrega = 7
        dataRealRetirada = 8
        PrevisaoRetirada = 9
        ValorServico = 10
        ValorDesconto = 11
        ValorTotal = 12
        Situacao = 13
        FormaPagamento = 14
        Endereco = 15
        Numero = 16
        Bairro = 17
        Cidade = 18
End Enum

Private Sub cmdPedido_Click()
        selecionaPedido
End Sub

Private Sub cmdPesquisa_Click()
        
        Dim strSQL                              As String
        Dim strFiltro                           As String
        Dim rsPedidoLocacao                     As ADODB.Recordset
        
        geraGridPesquisa
        
        If Not validaDataPesquisa Then Exit Sub
        
        strSQL = Empty
        strSQL = strSQL & " SELECT " & vbCrLf
        strSQL = strSQL & " PedidoLocacao.idPedidoLocacao " & vbCrLf
        strSQL = strSQL & " ,Clientes.Nome " & vbCrLf
        strSQL = strSQL & " ,Clientes.CNPJCpf " & vbCrLf
        strSQL = strSQL & " ,Cacambas.Numero as NumeroCacamba " & vbCrLf
        strSQL = strSQL & " ,NumeroControle " & vbCrLf
        strSQL = strSQL & " ,DataLocacao " & vbCrLf
        strSQL = strSQL & " ,DataRealRetirada " & vbCrLf
        strSQL = strSQL & " ,PrevisaoRetirada " & vbCrLf
        strSQL = strSQL & " ,DataEntrega " & vbCrLf
        strSQL = strSQL & " ,ValorServico " & vbCrLf
        strSQL = strSQL & " ,ValorDesconto " & vbCrLf
        strSQL = strSQL & " ,ValorTotal " & vbCrLf
        strSQL = strSQL & " ,Situacao " & vbCrLf
        strSQL = strSQL & " ,FormaPagamento " & vbCrLf
        strSQL = strSQL & " ,PedidoLocacao.Endereco " & vbCrLf
        strSQL = strSQL & " ,PedidoLocacao.EnderecoNumero " & vbCrLf
        strSQL = strSQL & " ,PedidoLocacao.Bairro " & vbCrLf
        strSQL = strSQL & " ,PedidoLocacao.Cidade " & vbCrLf
        strSQL = strSQL & " ,PedidoLocacao.TipoPedidoLocacao " & vbCrLf
        strSQL = strSQL & " FROM PedidoLocacao " & vbCrLf
        strSQL = strSQL & " INNER JOIN Cacambas ON PedidoLocacao.idCacambas = Cacambas.idCacambas" & vbCrLf
        strSQL = strSQL & " INNER JOIN Clientes ON PedidoLocacao.idClientes = Clientes.idClientes" & vbCrLf
        strSQL = strSQL & " INNER JOIN Usuarios ON PedidoLocacao.idUsuarios = Usuarios.idUsuarios" & vbCrLf
        
        strFiltro = Empty
        
        If Val(txtNumeroControle.Text) <> 0 Then
                strFiltro = " WHERE NumeroControle = " & txtNumeroControle.Text & vbCrLf
        End If
        
        If Trim(txtCliente.Text) <> "" Then
                If strFiltro = Empty Then
                        strFiltro = " WHERE PedidoLocacao.idClientes = " & vIdClientePedidoLocacao
                Else
                        strFiltro = strFiltro & " AND PedidoLocacao.idClientes = " & vIdClientePedidoLocacao
                End If
        End If
        
        If Not optTipoLocacaoTodos.Value Then
            If optTipoLocacaoAluguel.Value Then
                If strFiltro = Empty Then
                        strFiltro = " WHERE PedidoLocacao.TipoPedidoLocacao = " & e_TipoPedidoLocacao.aluguel
                Else
                        strFiltro = strFiltro & " PedidoLocacao.TipoPedidoLocacao = " & e_TipoPedidoLocacao.aluguel
                End If
            Else
                If strFiltro = Empty Then
                        strFiltro = " WHERE PedidoLocacao.TipoPedidoLocacao = " & e_TipoPedidoLocacao.troca
                Else
                        strFiltro = strFiltro & " AND  PedidoLocacao.TipoPedidoLocacao = " & e_TipoPedidoLocacao.troca
                End If
            End If
        End If
        
        If IsDate(mskDataLocacaoInicial.FormattedText) And IsDate(mskDataLocacaoFinal.FormattedText) Then
                If strFiltro = Empty Then
                        strFiltro = " WHERE DataLocacao BETWEEN '" & Format(mskDataLocacaoInicial.FormattedText, "yyyy-mm-dd") & "' AND '" & Format(mskDataLocacaoFinal.FormattedText, "yyyy-mm-dd") & "'" & vbCrLf
                Else
                        strFiltro = strFiltro & " AND DataLocacao BETWEEN '" & Format(mskDataLocacaoInicial.FormattedText, "yyyy-mm-dd") & "' AND '" & Format(mskDataLocacaoFinal.FormattedText, "yyyy-mm-dd") & "'" & vbCrLf
                End If
        End If
        
        If IsDate(mskDataRealRetiradaInicial.FormattedText) And IsDate(mskDataRealRetiradaFinal.FormattedText) Then
                If strFiltro = Empty Then
                        strFiltro = " WHERE DataRealRetirada BETWEEN '" & Format(mskDataRealRetiradaInicial.FormattedText, "yyyy-mm-dd") & "' AND '" & Format(mskDataRealRetiradaFinal.FormattedText, "yyyy-mm-dd") & "'" & vbCrLf
                Else
                        strFiltro = strFiltro & " AND DataRealRetirada BETWEEN '" & Format(mskDataRealRetiradaInicial.FormattedText, "yyyy-mm-dd") & "' AND '" & Format(mskDataRealRetiradaFinal.FormattedText, "yyyy-mm-dd") & "'" & vbCrLf
                End If
        End If
        
        If IsDate(mskDataEntregaInicial.FormattedText) And IsDate(mskDataEntregaFinal.FormattedText) Then
                If strFiltro = Empty Then
                        strFiltro = " WHERE DataEntrega BETWEEN '" & Format(mskDataEntregaInicial.FormattedText, "yyyy-mm-dd") & "' AND '" & Format(mskDataEntregaFinal.FormattedText, "yyyy-mm-dd") & "'" & vbCrLf
                Else
                        strFiltro = strFiltro & " AND DataEntrega BETWEEN '" & Format(mskDataEntregaInicial.FormattedText, "yyyy-mm-dd") & "' AND '" & Format(mskDataEntregaFinal.FormattedText, "yyyy-mm-dd") & "'" & vbCrLf
                End If
        End If

        If strFiltro = "" Then
                strFiltro = filtraFormaPagamento(True)
        Else
                strFiltro = strFiltro & filtraFormaPagamento(False)
        End If
        
        If strFiltro = "" Then
                strFiltro = filtraSituacao(True)
        Else
                strFiltro = strFiltro & filtraSituacao(False)
        End If
        
        strSQL = strSQL & strFiltro
        
        With flxPesquisa
                Set rsPedidoLocacao = New ADODB.Recordset
                rsPedidoLocacao.Open strSQL, dbCacamba, adOpenStatic, adLockReadOnly
                While Not rsPedidoLocacao.EOF
                        .AddItem ""
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.idPedidoLocacao) = NVL(rsPedidoLocacao!idPedidoLocacao, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.NumeroPedidoLocacao) = NVL(rsPedidoLocacao!numeroControle, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.NumeroCacamba) = NVL(rsPedidoLocacao!NumeroCacamba, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.TipoPedidoLocacao) = IIf(NVL(rsPedidoLocacao!TipoPedidoLocacao, 0) = e_TipoPedidoLocacao.aluguel, "Aluguel", "Troca")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.nome) = NVL(rsPedidoLocacao!nome, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.cpfcnpj) = NVL(rsPedidoLocacao!CNPJCPF, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.DataLocacao) = NVL(rsPedidoLocacao!DataLocacao, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.dataRealRetirada) = NVL(rsPedidoLocacao!dataRealRetirada, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.dataEntrega) = NVL(rsPedidoLocacao!dataEntrega, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.PrevisaoRetirada) = NVL(rsPedidoLocacao!PrevisaoRetirada, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.ValorServico) = Format(NVL(rsPedidoLocacao!ValorServico, "0"), "###,##0.00")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.ValorDesconto) = Format(NVL(rsPedidoLocacao!ValorDesconto, "0"), "###,##0.00")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.ValorTotal) = Format(NVL(rsPedidoLocacao!ValorTotal, "0"), "###,##0.00")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.Situacao) = NVL(rsPedidoLocacao!Situacao, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.FormaPagamento) = NVL(rsPedidoLocacao!FormaPagamento, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.Endereco) = NVL(rsPedidoLocacao!Endereco, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.Numero) = NVL(rsPedidoLocacao!EnderecoNumero, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.Bairro) = NVL(rsPedidoLocacao!Bairro, "")
                        .TextMatrix(.Rows - 1, colGridPedidoLocacao.Cidade) = NVL(rsPedidoLocacao!Cidade, "")
                        
                        rsPedidoLocacao.MoveNext
                Wend
        End With


End Sub
Private Function filtraFormaPagamento(pPrimeiroFiltro As Boolean) As String
        
        Dim strFiltro                           As String
        Dim strFiltroCondicao                   As String
        Dim vliQuantiadeFiltro                  As Integer
        
        vliQuantiadeFiltro = 0
        strFiltro = ""
        
        If chkFormaPagCartao.Value = vbChecked Then
                vliQuantiadeFiltro = 1
        End If
        
        If chkFormaPagCheque.Value = vbChecked Then
                vliQuantiadeFiltro = vliQuantiadeFiltro + 1
        End If
        
        If chkFormaPagDinheiro.Value = vbChecked Then
                vliQuantiadeFiltro = vliQuantiadeFiltro + 1
        End If
        
        If chkFormaPagTransferencia.Value = vbChecked Then
                vliQuantiadeFiltro = vliQuantiadeFiltro + 1
        End If
        
        If chkFormaPagBoleto.Value = vbChecked Then
                vliQuantiadeFiltro = vliQuantiadeFiltro + 1
        End If
        
        If vliQuantiadeFiltro = 0 Then
                Exit Function
        
        ElseIf vliQuantiadeFiltro = 1 Then
                If pPrimeiroFiltro Then
                        strFiltro = " WHERE " & vbCrLf
                Else
                        strFiltro = " AND " & vbCrLf
                End If
                
                If chkFormaPagCartao.Value = vbChecked Then
                        strFiltro = strFiltro & " FormaPagamento = 'Cartão'"
                End If
                
                If chkFormaPagCheque.Value = vbChecked Then
                        strFiltro = strFiltro & " FormaPagamento = 'Cheque'"
                End If
                
                If chkFormaPagDinheiro.Value = vbChecked Then
                        strFiltro = strFiltro & " FormaPagamento = 'Dinheiro'"
                End If
                
                If chkFormaPagTransferencia.Value = vbChecked Then
                        strFiltro = strFiltro & " FormaPagamento = 'Transferência'"
                End If
                
                If chkFormaPagBoleto.Value = vbChecked Then
                        strFiltro = strFiltro & " FormaPagamento = 'Boleto'"
                End If
        Else
                
                If pPrimeiroFiltro Then
                        strFiltro = " WHERE (" & vbCrLf
                Else
                        strFiltro = " AND (" & vbCrLf
                End If
                
                If chkFormaPagCartao.Value = vbChecked Then
                        If strFiltroCondicao = "" Then
                                strFiltroCondicao = strFiltroCondicao & " FormaPagamento = 'Cartão' " & vbCrLf
                        Else
                                strFiltroCondicao = strFiltroCondicao & " OR FormaPagamento = 'Cartão' " & vbCrLf
                        End If
                End If
                
                If chkFormaPagCheque.Value = vbChecked Then
                        If strFiltroCondicao = "" Then
                                strFiltroCondicao = strFiltroCondicao & " FormaPagamento = 'Cheque'" & vbCrLf
                        Else
                                strFiltroCondicao = strFiltroCondicao & " OR FormaPagamento = 'Cheque'" & vbCrLf
                        End If
                End If
                
                If chkFormaPagDinheiro.Value = vbChecked Then
                        If strFiltroCondicao = "" Then
                                strFiltroCondicao = strFiltroCondicao & " FormaPagamento = 'Dinheiro'" & vbCrLf
                        Else
                                strFiltroCondicao = strFiltroCondicao & " OR  FormaPagamento = 'Dinheiro'" & vbCrLf
                        End If
                End If
                
                If chkFormaPagTransferencia.Value = vbChecked Then
                        If strFiltroCondicao = "" Then
                                strFiltroCondicao = strFiltroCondicao & "  FormaPagamento = 'Transferência'" & vbCrLf
                        Else
                                strFiltroCondicao = strFiltroCondicao & " OR   FormaPagamento = 'Transferência'" & vbCrLf
                        End If
                End If
                
                If chkFormaPagBoleto.Value = vbChecked Then
                        If strFiltroCondicao = "" Then
                                strFiltroCondicao = strFiltroCondicao & "  FormaPagamento = 'Boleto'" & vbCrLf
                        Else
                                strFiltroCondicao = strFiltroCondicao & " OR    FormaPagamento = 'Boleto'" & vbCrLf
                        End If
                End If
                
                strFiltroCondicao = strFiltroCondicao & ")"
        End If
                
        filtraFormaPagamento = strFiltro & strFiltroCondicao
        
End Function

Private Function filtraSituacao(pPrimeiroFiltro As Boolean) As String
        
        Dim strFiltro                           As String
        Dim strFiltroCondicao                   As String
        Dim vliQuantiadeFiltro                  As Integer
        
        vliQuantiadeFiltro = 0
        strFiltro = ""
        
        If chkSituacaoAreceber.Value = vbChecked Then
                vliQuantiadeFiltro = 1
        End If
        
        If chkSituacaoPagoEntrega.Value = vbChecked Then
                vliQuantiadeFiltro = vliQuantiadeFiltro + 1
        End If
        
        If chkSituacaoPagoRetirada.Value = vbChecked Then
                vliQuantiadeFiltro = vliQuantiadeFiltro + 1
        End If
        
        If chkSituacaoDevedor.Value = vbChecked Then
                vliQuantiadeFiltro = vliQuantiadeFiltro + 1
        End If
        
        If chkSituacaoCortesia.Value = vbChecked Then
                vliQuantiadeFiltro = vliQuantiadeFiltro + 1
        End If
        
        If vliQuantiadeFiltro = 0 Then
                Exit Function
                
        ElseIf vliQuantiadeFiltro = 1 Then
                If pPrimeiroFiltro Then
                        strFiltro = " WHERE " & vbCrLf
                Else
                        strFiltro = " AND " & vbCrLf
                End If
                
                If chkSituacaoAreceber.Value = vbChecked Then
                        strFiltro = strFiltro & " Situacao = 'A Receber'"
                End If
                
                If chkSituacaoPagoEntrega.Value = vbChecked Then
                        strFiltro = strFiltro & " Situacao = 'Pago na Entrega'"
                End If
                
                If chkSituacaoPagoRetirada.Value = vbChecked Then
                        strFiltro = strFiltro & " Situacao = 'Pago na Retirada'"
                End If
                
                If chkSituacaoDevedor.Value = vbChecked Then
                        strFiltro = strFiltro & " Situacao = 'Devedor'"
                End If
                
                If chkSituacaoCortesia.Value = vbChecked Then
                        strFiltro = strFiltro & " Situacao = 'Cortesia'"
                End If
        Else
                If pPrimeiroFiltro Then
                        strFiltro = " WHERE (" & vbCrLf
                Else
                        strFiltro = " AND (" & vbCrLf
                End If
                
                If chkSituacaoAreceber.Value = vbChecked Then
                        If strFiltroCondicao = "" Then
                                strFiltroCondicao = strFiltroCondicao & " Situacao = 'A Receber' " & vbCrLf
                        Else
                                strFiltroCondicao = strFiltroCondicao & " OR Situacao = 'A Receber' " & vbCrLf
                        End If
                End If
                
                If chkSituacaoPagoEntrega.Value = vbChecked Then
                        If strFiltroCondicao = "" Then
                                strFiltroCondicao = strFiltroCondicao & " Situacao = 'Pago na Entrega' " & vbCrLf
                        Else
                                strFiltroCondicao = strFiltroCondicao & " OR Situacao = 'Pago na Entrega' " & vbCrLf
                        End If
                End If
                
                If chkSituacaoPagoRetirada.Value = vbChecked Then
                        If strFiltroCondicao = "" Then
                                strFiltroCondicao = strFiltroCondicao & " Situacao = 'Pago na Retirada' " & vbCrLf
                        Else
                                strFiltroCondicao = strFiltroCondicao & " OR Situacao = 'Pago na Retirada' " & vbCrLf
                        End If
                End If
                
                If chkSituacaoDevedor.Value = vbChecked Then
                        If strFiltroCondicao = "" Then
                                strFiltroCondicao = strFiltroCondicao & " Situacao = 'Devedor' " & vbCrLf
                        Else
                                strFiltroCondicao = strFiltroCondicao & " OR Situacao = 'Devedor' " & vbCrLf
                        End If
                End If
                
                If chkSituacaoCortesia.Value = vbChecked Then
                        If strFiltroCondicao = "" Then
                                strFiltroCondicao = strFiltroCondicao & " Situacao = 'Cortesia' " & vbCrLf
                        Else
                                strFiltroCondicao = strFiltroCondicao & " OR Situacao = 'Cortesia' " & vbCrLf
                        End If
                End If
                strFiltroCondicao = strFiltroCondicao & ")"
        End If
                
        filtraSituacao = strFiltro & strFiltroCondicao
        
End Function

Private Sub cmdSair_Click()
        Unload Me
End Sub

Private Sub cmdSelecionaCliente_Click()
        Dim strSQL                      As String
        Dim vColunas                    As String
        
        strSQL = Empty
        strSQL = strSQL & " SELECT " & vbCrLf
        strSQL = strSQL & " idClientes" & vbCrLf
        strSQL = strSQL & " ,Nome " & vbCrLf
        strSQL = strSQL & " ,CnpjCPF " & vbCrLf
        strSQL = strSQL & " ,Email " & vbCrLf
        strSQL = strSQL & " ,TelefoneFixo" & vbCrLf
        strSQL = strSQL & " ,TelefoneComercial" & vbCrLf
        strSQL = strSQL & " ,TelefoneCelular" & vbCrLf
        strSQL = strSQL & " FROM Clientes"
        
        vColunas = "Código,Nome, CPF/CNPJ, Email, TelefoneFixo, TelefoneComercial, TelefoneCelular"
        
        frmPesquisa.carregaGridPesquisa 7, strSQL, vColunas, 1
        frmPesquisa.Show vbModal
        
        If vgRetornoConsulta <> 0 Then
                vIdClientePedidoLocacao = vgRetornoConsulta
                txtCliente.Text = recuperaDescricao("Clientes", "Nome", "idClientes", CInt(vIdClientePedidoLocacao), False)
        Else
                vIdClientePedidoLocacao = 0
                txtCliente.Text = ""
        End If


End Sub

Private Sub Form_Load()
        geraGridPesquisa
End Sub



Private Sub mskDataEntregaFinal_LostFocus()
        If Not IsDate(mskDataEntregaFinal.FormattedText) And mskDataEntregaFinal.Text <> "__/__/____" Then
                MsgBox "Data de Entrega Final inválida!", vbInformation, "Caçambas"
                mskDataEntregaFinal.Text = "__/__/____"
                mskDataEntregaFinal.SetFocus
                Exit Sub
        End If
        
        If IsDate(mskDataEntregaInicial.FormattedText) And IsDate(mskDataEntregaFinal.FormattedText) Then
                If CDate(mskDataEntregaInicial.FormattedText) > CDate(mskDataEntregaFinal.FormattedText) Then
                        MsgBox "A data de Entrega Final deve ser maior ou igual a data de Entrega Inicial!", vbInformation, "STM Cargas"
                        mskDataEntregaInicial.Text = "__/__/____"
                        mskDataEntregaFinal.Text = "__/__/____"
                        mskDataEntregaInicial.SetFocus
                        Exit Sub
                End If
        End If
End Sub

Private Sub mskDataEntregaInicial_LostFocus()
        If Not IsDate(mskDataEntregaInicial.FormattedText) And mskDataEntregaInicial.Text <> "__/__/____" Then
                MsgBox "Data de Entrega Inicial inválida!", vbInformation, "Caçambas"
                mskDataEntregaInicial.Text = "__/__/____"
                mskDataEntregaInicial.SetFocus
                Exit Sub
        End If
        
        If IsDate(mskDataEntregaInicial.FormattedText) And IsDate(mskDataEntregaFinal.FormattedText) Then
                If CDate(mskDataEntregaInicial.FormattedText) > CDate(mskDataEntregaFinal.FormattedText) Then
                        MsgBox "A data de Entrega Final deve ser maior ou igual a data de Entrega Inicial!", vbInformation, "STM Cargas"
                        mskDataEntregaInicial.Text = "__/__/____"
                        mskDataEntregaFinal.Text = "__/__/____"
                        mskDataEntregaInicial.SetFocus
                        Exit Sub
                End If
        End If
End Sub

Private Sub mskDataLocacaoFinal_LostFocus()
        If Not IsDate(mskDataLocacaoFinal.FormattedText) And mskDataLocacaoFinal.Text <> "__/__/____" Then
                MsgBox "Data Locação Final inválida!", vbInformation, "Caçambas"
                mskDataLocacaoFinal.Text = "__/__/____"
                mskDataLocacaoFinal.SetFocus
                Exit Sub
        End If
        
        If IsDate(mskDataLocacaoInicial.FormattedText) And IsDate(mskDataLocacaoFinal.FormattedText) Then
                If CDate(mskDataLocacaoInicial.FormattedText) > CDate(mskDataLocacaoFinal.FormattedText) Then
                        MsgBox "A data de Locação Final deve ser maior ou igual a data de Locacação Inicial!", vbInformation, "STM Cargas"
                        mskDataLocacaoFinal.Text = "__/__/____"
                        mskDataLocacaoInicial.Text = "__/__/____"
                        mskDataLocacaoInicial.SetFocus
                        Exit Sub
                End If
        End If
End Sub

Private Sub mskDataLocacaoInicial_LostFocus()
        If Not IsDate(mskDataLocacaoInicial.FormattedText) And mskDataLocacaoInicial.Text <> "__/__/____" Then
                MsgBox "Data Locação Inicial inválida!", vbInformation, "Caçambas"
                mskDataLocacaoInicial.Text = "__/__/____"
                mskDataLocacaoInicial.SetFocus
                Exit Sub
        End If
        
        If IsDate(mskDataLocacaoInicial.FormattedText) And IsDate(mskDataLocacaoFinal.FormattedText) Then
                If CDate(mskDataLocacaoInicial.FormattedText) > CDate(mskDataLocacaoFinal.FormattedText) Then
                        MsgBox "A data de Locação Final deve ser maior ou igual a data de Locacação Inicial!", vbInformation, "STM Cargas"
                        mskDataLocacaoInicial.Text = "__/__/____"
                        mskDataLocacaoFinal.Text = "__/__/____"
                        mskDataLocacaoInicial.SetFocus
                        Exit Sub
                End If
        End If
End Sub


Private Sub mskDataRealRetiradaFinal_LostFocus()
        If Not IsDate(mskDataRealRetiradaFinal.FormattedText) And mskDataRealRetiradaFinal.Text <> "__/__/____" Then
                MsgBox "Data de Real de Retirada Inicial inválida!", vbInformation, "Caçambas"
                mskDataRealRetiradaFinal.Text = "__/__/____"
                mskDataRealRetiradaFinal.SetFocus
                Exit Sub
        End If
        
        If IsDate(mskDataRealRetiradaInicial.FormattedText) And IsDate(mskDataRealRetiradaFinal.FormattedText) Then
                If CDate(mskDataRealRetiradaInicial.FormattedText) > CDate(mskDataRealRetiradaFinal.FormattedText) Then
                        MsgBox "A data de Retirada Final deve ser maior ou igual a data de Retirada Inicial!", vbInformation, "STM Cargas"
                        mskDataRealRetiradaInicial.Text = "__/__/____"
                        mskDataRealRetiradaFinal.Text = "__/__/____"
                        mskDataRealRetiradaInicial.SetFocus
                        Exit Sub
                End If
        End If
End Sub

Private Sub mskDataRealRetiradaInicial_LostFocus()
        If Not IsDate(mskDataRealRetiradaInicial.FormattedText) And mskDataRealRetiradaInicial.Text <> "__/__/____" Then
                MsgBox "Data de Real de Retirada Inicial inválida!", vbInformation, "Caçambas"
                mskDataRealRetiradaInicial.Text = "__/__/____"
                mskDataRealRetiradaInicial.SetFocus
                Exit Sub
        End If
        
        If IsDate(mskDataRealRetiradaInicial.FormattedText) And IsDate(mskDataRealRetiradaFinal.FormattedText) Then
                If CDate(mskDataRealRetiradaInicial.FormattedText) > CDate(mskDataRealRetiradaFinal.FormattedText) Then
                        MsgBox "A data de Retirada Final deve ser maior ou igual a data de Retirada Inicial!", vbInformation, "STM Cargas"
                        mskDataRealRetiradaInicial.Text = "__/__/____"
                        mskDataRealRetiradaFinal.Text = "__/__/____"
                        mskDataRealRetiradaInicial.SetFocus
                        Exit Sub
                End If
        End If
End Sub

Private Sub txtNumeroControle_KeyPress(KeyAscii As Integer)
        
        KeyAscii = SoNumeros(KeyAscii)
        
        If KeyAscii = 0 Then
                Exit Sub
        End If
        
End Sub

Private Sub geraGridPesquisa()
        
        With flxPesquisa
                .Rows = 1
                .Cols = 19
                .FixedCols = 0
                
                .TextMatrix(0, colGridPedidoLocacao.idPedidoLocacao) = ""
                .TextMatrix(0, colGridPedidoLocacao.NumeroPedidoLocacao) = "Nº Pedido"
                .TextMatrix(0, colGridPedidoLocacao.NumeroCacamba) = "Nº Caçamba"
                .TextMatrix(0, colGridPedidoLocacao.TipoPedidoLocacao) = "Tipo Pedido"
                .TextMatrix(0, colGridPedidoLocacao.nome) = "Cliente"
                .TextMatrix(0, colGridPedidoLocacao.cpfcnpj) = "CPF/CNPJ"
                .TextMatrix(0, colGridPedidoLocacao.DataLocacao) = "Data Locação"
                .TextMatrix(0, colGridPedidoLocacao.dataRealRetirada) = "Data Retirada"
                .TextMatrix(0, colGridPedidoLocacao.dataEntrega) = "Data Entrega"
                .TextMatrix(0, colGridPedidoLocacao.PrevisaoRetirada) = "Previsão Retirada"
                .TextMatrix(0, colGridPedidoLocacao.ValorServico) = "Valor"
                .TextMatrix(0, colGridPedidoLocacao.ValorDesconto) = "Desconto"
                .TextMatrix(0, colGridPedidoLocacao.ValorTotal) = "Total"
                .TextMatrix(0, colGridPedidoLocacao.Situacao) = "Situação"
                .TextMatrix(0, colGridPedidoLocacao.FormaPagamento) = "Forma Pagamento"
                .TextMatrix(0, colGridPedidoLocacao.Endereco) = "Endereço"
                .TextMatrix(0, colGridPedidoLocacao.Numero) = "Número"
                .TextMatrix(0, colGridPedidoLocacao.Bairro) = "Bairro"
                .TextMatrix(0, colGridPedidoLocacao.Cidade) = "Cidade"
                
                .ColWidth(colGridPedidoLocacao.idPedidoLocacao) = 0
                .ColWidth(colGridPedidoLocacao.NumeroPedidoLocacao) = 1000
                .ColWidth(colGridPedidoLocacao.TipoPedidoLocacao) = 1500
                .ColWidth(colGridPedidoLocacao.NumeroCacamba) = 1000
                .ColWidth(colGridPedidoLocacao.nome) = 2500
                .ColWidth(colGridPedidoLocacao.cpfcnpj) = 1200
                .ColWidth(colGridPedidoLocacao.DataLocacao) = 1200
                .ColWidth(colGridPedidoLocacao.dataRealRetirada) = 1200
                .ColWidth(colGridPedidoLocacao.dataEntrega) = 1200
                .ColWidth(colGridPedidoLocacao.PrevisaoRetirada) = 1400
                .ColWidth(colGridPedidoLocacao.ValorServico) = 1000
                .ColWidth(colGridPedidoLocacao.ValorDesconto) = 1000
                .ColWidth(colGridPedidoLocacao.ValorTotal) = 1000
                .ColWidth(colGridPedidoLocacao.Situacao) = 1600
                .ColWidth(colGridPedidoLocacao.FormaPagamento) = 1400
                .ColWidth(colGridPedidoLocacao.Endereco) = 2000
                .ColWidth(colGridPedidoLocacao.Numero) = 900
                .ColWidth(colGridPedidoLocacao.Bairro) = 1100
                .ColWidth(colGridPedidoLocacao.Cidade) = 1100
                
                
                .SelectionMode = flexSelectionByRow
                .GridLines = flexGridInset
        End With
End Sub


Private Function validaDataPesquisa() As Boolean
        
        validaDataPesquisa = False
        
        If mskDataEntregaInicial.Text = "__/__/____" And IsDate(mskDataEntregaFinal.FormattedText) Then
                MsgBox "Preencha a data de Entrega Inicial", vbInformation, "Caçambas"
                mskDataEntregaInicial.SetFocus
                Exit Function
        End If
        
        If mskDataEntregaFinal.Text = "__/__/____" And IsDate(mskDataEntregaInicial.FormattedText) Then
                MsgBox "Preencha a data de Entrega Final", vbInformation, "Caçambas"
                mskDataEntregaFinal.SetFocus
                Exit Function
        End If
        
        If mskDataLocacaoInicial.Text = "__/__/____" And IsDate(mskDataLocacaoFinal.FormattedText) Then
                MsgBox "Preencha a data de Locação Inicial", vbInformation, "Caçambas"
                mskDataLocacaoInicial.SetFocus
                Exit Function
        End If
        
        If mskDataLocacaoFinal.Text = "__/__/____" And IsDate(mskDataLocacaoInicial.FormattedText) Then
                MsgBox "Preencha a data de Locação Final", vbInformation, "Caçambas"
                mskDataLocacaoFinal.SetFocus
                Exit Function
        End If
        
        If mskDataRealRetiradaInicial.Text = "__/__/____" And IsDate(mskDataRealRetiradaFinal.FormattedText) Then
                MsgBox "Preencha a data Real de Retirada Inicial", vbInformation, "Caçambas"
                mskDataRealRetiradaInicial.SetFocus
                Exit Function
        End If
        
        If mskDataRealRetiradaFinal.Text = "__/__/____" And IsDate(mskDataRealRetiradaInicial.FormattedText) Then
                MsgBox "Preencha a data Real de Retirada Final", vbInformation, "Caçambas"
                mskDataRealRetiradaFinal.SetFocus
                Exit Function
        End If
        
        validaDataPesquisa = True
End Function

Private Sub selecionaPedido()
        
        If flxPesquisa.Rows = 1 Then
                MsgBox "Nenhum pedido selcionado!", vbInformation, "Caçamba"
                Exit Sub
        End If
        
        With frmPedidoLocacao
                Load frmPedidoLocacao
                .vIdPedidoLocacao = flxPesquisa.TextMatrix(flxPesquisa.Row, colGridPedidoLocacao.idPedidoLocacao)
                .carregaPedidoLocacao flxPesquisa.TextMatrix(flxPesquisa.Row, colGridPedidoLocacao.idPedidoLocacao)
                .Show vbModal
                cmdPesquisa_Click
        End With
End Sub

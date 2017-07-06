VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPedidoLocacao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedido de Locação"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11880
   Icon            =   "frmPedidoLocacao.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraPedidoLocacao 
      Height          =   7935
      Left            =   120
      TabIndex        =   35
      Top             =   0
      Width           =   11655
      Begin VB.CommandButton cmdSelecionaBanco 
         Height          =   375
         Left            =   4680
         Picture         =   "frmPedidoLocacao.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   7320
         Width           =   375
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo Locação"
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
         Left            =   4680
         TabIndex        =   62
         Top             =   240
         Width           =   2295
         Begin VB.OptionButton optTroca 
            Caption         =   "Troca"
            Height          =   195
            Left            =   1200
            TabIndex        =   64
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optAluguel 
            Caption         =   "Aluguel"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin MSMask.MaskEdBox mskDataPagamento 
         Height          =   375
         Left            =   7200
         TabIndex        =   26
         Top             =   7320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskValorTotal 
         Height          =   375
         Left            =   4080
         TabIndex        =   23
         Top             =   6600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskValorDesconto 
         Height          =   375
         Left            =   2040
         TabIndex        =   22
         Top             =   6600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskValorServico 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   6600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDataRealRetirada 
         Height          =   375
         Left            =   5280
         TabIndex        =   6
         Top             =   1440
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDataEntrega 
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtTotalDiasAlocado 
         Enabled         =   0   'False
         Height          =   375
         Left            =   7080
         MaxLength       =   3
         TabIndex        =   7
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox txtUsuario 
         Enabled         =   0   'False
         Height          =   405
         Left            =   9120
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   28
         Top             =   7320
         Width           =   2295
      End
      Begin VB.ComboBox cboSituacao 
         Height          =   315
         Left            =   5160
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   7320
         Width           =   1935
      End
      Begin VB.ComboBox cboFormaPagamento 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   7320
         Width           =   2535
      End
      Begin VB.CommandButton cmdOutroEndereco 
         Caption         =   "Outro Endereço"
         Height          =   375
         Left            =   6720
         TabIndex        =   10
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelecionaCliente 
         Height          =   375
         Left            =   6120
         Picture         =   "frmPedidoLocacao.frx":03E1
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtCliente 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   8
         Top             =   2160
         Width           =   6375
      End
      Begin VB.TextBox txtPrevisaoRetirada 
         Height          =   375
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton cmdSelecionaCacamba 
         Height          =   375
         Left            =   3960
         Picture         =   "frmPedidoLocacao.frx":07B6
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtNumeroControle 
         Height          =   375
         Left            =   120
         MaxLength       =   10
         TabIndex        =   0
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtNumeroCacamba 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtTelefoneFixo 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   12
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox txtEmailCliente 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   11
         Top             =   2880
         Width           =   6375
      End
      Begin VB.TextBox txtEndereco 
         Height          =   375
         Left            =   120
         MaxLength       =   150
         TabIndex        =   15
         Top             =   4440
         Width           =   7455
      End
      Begin VB.TextBox txtEnderecoNumero 
         Height          =   375
         Left            =   7680
         MaxLength       =   15
         TabIndex        =   16
         Top             =   4440
         Width           =   1695
      End
      Begin VB.TextBox txtEnderecoBairro 
         Height          =   375
         Left            =   120
         MaxLength       =   45
         TabIndex        =   17
         Top             =   5160
         Width           =   3375
      End
      Begin VB.TextBox txtEnderecoCidade 
         Height          =   375
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   18
         Top             =   5160
         Width           =   3375
      End
      Begin VB.ComboBox cboEnderecoEstado 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   5880
         Width           =   855
      End
      Begin VB.TextBox txtEnderecoComplemento 
         Height          =   375
         Left            =   1200
         MaxLength       =   45
         TabIndex        =   20
         Top             =   5880
         Width           =   3375
      End
      Begin VB.TextBox txtTelefoneComercial 
         Height          =   375
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   13
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox txtTelefoneCelular 
         Height          =   375
         Left            =   4440
         Locked          =   -1  'True
         MaxLength       =   14
         TabIndex        =   14
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox txtBanco 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2760
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   66
         Top             =   7320
         Width           =   2295
      End
      Begin VB.Label lblBanco 
         Caption         =   "Caçamba"
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
         Left            =   2760
         TabIndex        =   67
         Top             =   7080
         Width           =   1455
      End
      Begin VB.Label Label16 
         Caption         =   "Data do Pagamento"
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
         Left            =   7200
         TabIndex        =   61
         Top             =   7080
         Width           =   1695
      End
      Begin VB.Label Label14 
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
         Height          =   255
         Left            =   3600
         TabIndex        =   60
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "Total Dias alocado"
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
         Left            =   7080
         TabIndex        =   59
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Usuario"
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
         Left            =   9120
         TabIndex        =   58
         Top             =   7080
         Width           =   735
      End
      Begin VB.Label Label11 
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
         Height          =   375
         Left            =   5160
         TabIndex        =   57
         Top             =   7080
         Width           =   1455
      End
      Begin VB.Label Label10 
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
         Height          =   375
         Left            =   120
         TabIndex        =   56
         Top             =   7080
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Total"
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
         Left            =   4080
         TabIndex        =   55
         Top             =   6360
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Desconto"
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
         Left            =   2040
         TabIndex        =   54
         Top             =   6360
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Valor do Serviço"
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
         TabIndex        =   53
         Top             =   6360
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Cliente"
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
         TabIndex        =   52
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label5 
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
         Height          =   255
         Left            =   5280
         TabIndex        =   51
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Dias"
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
         Left            =   2400
         TabIndex        =   50
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Previsão de Retirada"
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
         Left            =   1560
         TabIndex        =   49
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Data Locação"
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
         TabIndex        =   48
         Top             =   1200
         Width           =   1455
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
         TabIndex        =   47
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblNomeCliente 
         Caption         =   "Caçamba"
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
         TabIndex        =   46
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblTelefoneCliente 
         Caption         =   "Telefone Fixo"
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
         TabIndex        =   45
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label lblEmailCliente 
         Caption         =   "Email"
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
         TabIndex        =   44
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label lblEndereco 
         Caption         =   "Endereço"
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
         TabIndex        =   43
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label lblNumero 
         Caption         =   "Número"
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
         Left            =   7680
         TabIndex        =   42
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label lblEnderecoBairro 
         Caption         =   "Bairro"
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
         TabIndex        =   41
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label lblEnderecoCidade 
         Caption         =   "Cidade"
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
         Left            =   3600
         TabIndex        =   40
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label lblEnderecoEstado 
         Caption         =   "UF"
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
         TabIndex        =   39
         Top             =   5640
         Width           =   975
      End
      Begin VB.Label lblEnderecoComplemento 
         Caption         =   "Complemento"
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
         Left            =   1200
         TabIndex        =   38
         Top             =   5640
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Telefone Comercial"
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
         Left            =   2280
         TabIndex        =   37
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label lblTelefoneCelular 
         Caption         =   "Celular"
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
         Left            =   4440
         TabIndex        =   36
         Top             =   3480
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "Localizar"
      Height          =   675
      Left            =   120
      TabIndex        =   31
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   675
      Left            =   1305
      TabIndex        =   29
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   675
      Left            =   3660
      TabIndex        =   27
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Height          =   675
      Left            =   4845
      TabIndex        =   32
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   675
      Left            =   6015
      TabIndex        =   33
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   675
      Left            =   7200
      TabIndex        =   34
      Top             =   8040
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   675
      Left            =   2475
      TabIndex        =   30
      Top             =   8040
      Width           =   1095
   End
End
Attribute VB_Name = "frmPedidoLocacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vIdClientePedidoLocacao                  As Integer
Public vIdPedidoLocacao                         As Integer
Private vIdCacambas                             As Integer
Private vIdBancos                               As Integer
Private vOperacao                               As operacao
                             

Private Sub cboSituacao_Click()
        If UCase(cboSituacao.Text) = UCase("Pago na Entrega") Or UCase(cboSituacao.Text) = UCase("Pago na Retirada") Then
                If vOperacao <> consulta Then
                        mskDataPagamento.Enabled = True
                        mskDataPagamento.SetFocus
                End If
        Else
                mskDataPagamento.Enabled = False
                mskDataPagamento.Text = "__/__/____"
        End If
End Sub

Private Sub cmdOutroEndereco_Click()
        txtEndereco.Text = ""
        txtEnderecoNumero.Text = ""
        txtEnderecoBairro.Text = ""
        txtEnderecoCidade.Text = ""
        txtEnderecoComplemento.Text = ""
        cboEnderecoEstado.ListIndex = 0
        txtEndereco.SetFocus
End Sub

Private Sub cmdSair_Click()
        Unload Me
End Sub

Private Sub cmdSelecionaBanco_Click()
        Dim strSQL                      As String
        Dim vColunas                    As String
        
        strSQL = Empty
        strSQL = strSQL & " SELECT " & vbCrLf
        strSQL = strSQL & " idBancos" & vbCrLf
        strSQL = strSQL & " ,nomeBanco " & vbCrLf
        strSQL = strSQL & " ,numeroBanco " & vbCrLf
        strSQL = strSQL & " FROM Bancos"
        
        vColunas = "Código,Nome, Número"
        
        frmPesquisa.carregaGridPesquisa 3, strSQL, vColunas, 1
        frmPesquisa.Show vbModal
        
        If vgRetornoConsulta <> 0 Then
                vIdBancos = vgRetornoConsulta
                txtBanco.Text = recuperaDescricao("bancos", "nomeBanco", "idBancos", vIdBancos)
        Else
                vIdBancos = 0
                txtBanco.Text = ""
        End If

End Sub

Private Sub cmdSelecionaCacamba_Click()
        Dim strSQL                      As String
        Dim vColunas                    As String
        Dim vNumeroControleAux          As Integer
        
        strSQL = Empty
        strSQL = strSQL & " SELECT " & vbCrLf
        strSQL = strSQL & " idCacambas " & vbCrLf
        strSQL = strSQL & " ,numero" & vbCrLf
        strSQL = strSQL & " FROM Cacambas"
        
        vColunas = "Código,Numero"
        
        frmPesquisa.carregaGridPesquisa 2, strSQL, vColunas, 1
        frmPesquisa.Show vbModal
        
        vNumeroControleAux = 0
        
        If vgRetornoConsulta <> 0 Then
                vNumeroControleAux = verificaCacambaAlocada(vgRetornoConsulta)
                
                If vNumeroControleAux > 0 Then
                        MsgBox "A caçamba selecionada está alocada no pedido de locação Nº " & vNumeroControleAux & vbCrLf _
                                & "Você deve fazer a retiradda da mesma antes de fazer uma nova locação!"
                        vIdCacambas = 0
                        txtNumeroCacamba.Text = ""
                        Exit Sub
                End If
                vIdCacambas = vgRetornoConsulta
                txtNumeroCacamba = recuperaDescricao("Cacambas", "Numero", "idcacambas", CInt(vgRetornoConsulta), False)
        Else
                vIdCacambas = 0
                txtNumeroCacamba.Text = ""
        End If

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
                carregaInformacaoCliente vIdClientePedidoLocacao
        Else
                vIdClientePedidoLocacao = 0
                txtCliente.Text = ""
                txtEmailCliente.Text = ""
                txtTelefoneCelular.Text = ""
                txtTelefoneFixo.Text = ""
                txtTelefoneComercial = ""
                txtEndereco.Text = ""
                txtEnderecoNumero.Text = ""
                txtEnderecoBairro.Text = ""
                txtEnderecoCidade.Text = ""
                txtEnderecoComplemento.Text = ""
                cboEnderecoEstado.ListIndex = 0
        End If

End Sub

Public Sub carregaInformacaoCliente(pIdCliente As Integer)
        
        Dim strSQL                      As String
        Dim rsClientes                  As ADODB.Recordset
        
        strSQL = Empty
        strSQL = strSQL & " SELECT " & vbCrLf
        strSQL = strSQL & " idClientes" & vbCrLf
        strSQL = strSQL & " ,Nome " & vbCrLf
        strSQL = strSQL & " ,CnpjCPF " & vbCrLf
        strSQL = strSQL & " ,Email " & vbCrLf
        strSQL = strSQL & " ,TelefoneFixo" & vbCrLf
        strSQL = strSQL & " ,TelefoneComercial" & vbCrLf
        strSQL = strSQL & " ,TelefoneCelular" & vbCrLf
        strSQL = strSQL & " ,Endereco" & vbCrLf
        strSQL = strSQL & " ,Numero" & vbCrLf
        strSQL = strSQL & " ,Bairro" & vbCrLf
        strSQL = strSQL & " ,Cidade" & vbCrLf
        strSQL = strSQL & " ,UF" & vbCrLf
        strSQL = strSQL & " ,Complemento" & vbCrLf
        strSQL = strSQL & " FROM Clientes"
        strSQL = strSQL & " WHERE idClientes = " & pIdCliente
        
        Set rsClientes = New ADODB.Recordset
        
        With rsClientes
                .Open strSQL, dbCacamba, adOpenStatic, adLockReadOnly
                If Not .EOF Then
                        txtCliente.Text = NVL(!nome, "")
                        txtEmailCliente.Text = NVL(!Email, "")
                        txtTelefoneFixo.Text = NVL(!TelefoneFixo, "")
                        txtTelefoneComercial.Text = NVL(!TelefoneComercial, "")
                        txtTelefoneCelular.Text = NVL(!TelefoneCelular, "")
                        txtEndereco.Text = NVL(!Endereco, "")
                        txtEnderecoNumero.Text = NVL(!Numero, "")
                        txtEnderecoBairro.Text = NVL(!Bairro, "")
                        txtEnderecoCidade.Text = NVL(!Cidade, "")
                        cboEnderecoEstado.Text = NVL(!UF, "")
                        txtEnderecoComplemento.Text = NVL(!Complemento, "")
                End If
        End With
End Sub
Private Sub Form_Load()
        
        CarregaComboUf cboEnderecoEstado
        carregaComboFormaPagamento cboFormaPagamento
        carregaComboSituacao cboSituacao
        controlaBotoes Me, operacao.consulta
        carregaPedidoLocacao
        txtUsuario.Text = recuperaDescricao("Usuarios", "Nome", "idUsuarios", CInt(vgIdUsuarioLogado), False)
        fraPedidoLocacao.Enabled = False
        
End Sub
Private Sub calculaValorServico()
        mskValorTotal = mskValorServico - mskValorDesconto
End Sub


Private Sub mskData_LostFocus()
        If Not IsDate(mskData.FormattedText) Then
                MsgBox "Data de locação inválida!", vbInformation, "Caçambas"
                mskData.Text = "__/__/____"
                mskData.SetFocus
                Exit Sub
        End If
End Sub



Private Sub mskDataPagamento_LostFocus()
        If mskDataPagamento.Text <> "__/__/____" And Not IsDate(mskDataPagamento.FormattedText) Then
                MsgBox "Data de Pagamento inválida!", vbInformation, "Caçambas"
                mskDataPagamento.Text = "__/__/____"
                mskDataPagamento.SetFocus
                Exit Sub
        End If
End Sub

Private Sub mskDataRealRetirada_LostFocus()
        If Not IsDate(mskDataRealRetirada.FormattedText) Then
                MsgBox "Data Real de Retirada inválida!", vbInformation, "Caçambas"
                mskDataRealRetirada.Text = "__/__/____"
                mskDataRealRetirada.SetFocus
                Exit Sub
        End If
        
        If IsDate(mskData.FormattedText) And IsDate(mskDataRealRetirada.FormattedText) Then
                If CDate(mskData.FormattedText) > CDate(mskDataRealRetirada.FormattedText) Then
                        MsgBox "A data real de retirada deve ser maior ou igual a data de locação!", vbInformation, "STM Cargas"
                        mskDataRealRetirada.Text = "__/__/____"
                        mskDataRealRetirada.SetFocus
                        Exit Sub
                End If
        End If
        
        txtTotalDiasAlocado.Text = DateDiff("d", mskData.FormattedText, mskDataRealRetirada.FormattedText)
End Sub

Private Sub mskValorDesconto_LostFocus()
        calculaValorServico
End Sub

Private Sub mskValorServico_LostFocus()
        calculaValorServico
End Sub


Public Sub carregaPedidoLocacao(Optional pIdPedidoLocacao As Integer)

        Dim strSQL                      As String
        Dim rsPedidoLocacao                   As ADODB.Recordset
        
        strSQL = Empty
        If pIdPedidoLocacao = 0 Then
                strSQL = strSQL & " SELECT " & vbCrLf
                strSQL = strSQL & " idPedidoLocacao " & vbCrLf
                strSQL = strSQL & " ,Clientes.idClientes " & vbCrLf
                strSQL = strSQL & " ,Clientes.Nome " & vbCrLf
                strSQL = strSQL & " ,Usuarios.Nome as NomeUsuario " & vbCrLf
                strSQL = strSQL & " ,Cacambas.idCacambas " & vbCrLf
                strSQL = strSQL & " ,Cacambas.Numero as NumeroCacamba " & vbCrLf
                strSQL = strSQL & " ,NumeroControle " & vbCrLf
                strSQL = strSQL & " ,DataLocacao " & vbCrLf
                strSQL = strSQL & " ,PrevisaoRetirada " & vbCrLf
                strSQL = strSQL & " ,DataRealRetirada " & vbCrLf
                strSQL = strSQL & " ,EnderecoNumero " & vbCrLf
                strSQL = strSQL & " ,PedidoLocacao.Bairro " & vbCrLf
                strSQL = strSQL & " ,PedidoLocacao.Cidade " & vbCrLf
                strSQL = strSQL & " ,PedidoLocacao.UF " & vbCrLf
                strSQL = strSQL & " ,PedidoLocacao.Complemento " & vbCrLf
                strSQL = strSQL & " ,PedidoLocacao.Endereco " & vbCrLf
                strSQL = strSQL & " ,PedidoLocacao.DataPagamento " & vbCrLf
                strSQL = strSQL & " ,PedidoLocacao.DataEntrega " & vbCrLf
                strSQL = strSQL & " ,ValorServico " & vbCrLf
                strSQL = strSQL & " ,ValorDesconto " & vbCrLf
                strSQL = strSQL & " ,ValorTotal " & vbCrLf
                strSQL = strSQL & " ,FormaPagamento " & vbCrLf
                strSQL = strSQL & " ,Situacao " & vbCrLf
                strSQL = strSQL & " ,Clientes.email " & vbCrLf
                strSQL = strSQL & " ,Clientes.telefoneCelular " & vbCrLf
                strSQL = strSQL & " ,Clientes.telefoneFixo " & vbCrLf
                strSQL = strSQL & " ,Clientes.telefoneComercial " & vbCrLf
                strSQL = strSQL & " ,PedidoLocacao.TipoPedidoLocacao " & vbCrLf
                strSQL = strSQL & " ,Bancos.idBancos " & vbCrLf
                strSQL = strSQL & " ,Bancos.nomeBanco " & vbCrLf
                strSQL = strSQL & " FROM PedidoLocacao " & vbCrLf
                strSQL = strSQL & " INNER JOIN Clientes ON PedidoLocacao.idClientes = Clientes.idClientes" & vbCrLf
                strSQL = strSQL & " INNER JOIN Cacambas ON PedidoLocacao.idCacambas = Cacambas.idCacambas" & vbCrLf
                strSQL = strSQL & " INNER JOIN Usuarios ON PedidoLocacao.idUsuarios = Usuarios.idUsuarios" & vbCrLf
                strSQL = strSQL & " LEFT JOIN Bancos ON PedidoLocacao.idBancos = Bancos.idBancos" & vbCrLf
                strSQL = strSQL & " ORDER BY idPedidoLocacao DESC "
                strSQL = strSQL & " LIMIT 1"
        Else
                strSQL = strSQL & " SELECT " & vbCrLf
                strSQL = strSQL & " idPedidoLocacao " & vbCrLf
                strSQL = strSQL & " ,Clientes.idClientes " & vbCrLf
                strSQL = strSQL & " ,Clientes.Nome " & vbCrLf
                strSQL = strSQL & " ,Usuarios.Nome as NomeUsuario " & vbCrLf
                strSQL = strSQL & " ,Cacambas.idCacambas " & vbCrLf
                strSQL = strSQL & " ,Cacambas.Numero as NumeroCacamba " & vbCrLf
                strSQL = strSQL & " ,NumeroControle " & vbCrLf
                strSQL = strSQL & " ,DataLocacao " & vbCrLf
                strSQL = strSQL & " ,PrevisaoRetirada " & vbCrLf
                strSQL = strSQL & " ,DataRealRetirada " & vbCrLf
                strSQL = strSQL & " ,EnderecoNumero " & vbCrLf
                strSQL = strSQL & " ,PedidoLocacao.Bairro " & vbCrLf
                strSQL = strSQL & " ,PedidoLocacao.Cidade " & vbCrLf
                strSQL = strSQL & " ,PedidoLocacao.UF " & vbCrLf
                strSQL = strSQL & " ,PedidoLocacao.Complemento " & vbCrLf
                strSQL = strSQL & " ,PedidoLocacao.Endereco " & vbCrLf
                strSQL = strSQL & " ,PedidoLocacao.DataPagamento " & vbCrLf
                strSQL = strSQL & " ,PedidoLocacao.DataEntrega " & vbCrLf
                strSQL = strSQL & " ,ValorServico " & vbCrLf
                strSQL = strSQL & " ,ValorDesconto " & vbCrLf
                strSQL = strSQL & " ,ValorTotal " & vbCrLf
                strSQL = strSQL & " ,FormaPagamento " & vbCrLf
                strSQL = strSQL & " ,Situacao " & vbCrLf
                strSQL = strSQL & " ,Clientes.email " & vbCrLf
                strSQL = strSQL & " ,Clientes.telefoneCelular " & vbCrLf
                strSQL = strSQL & " ,Clientes.telefoneFixo " & vbCrLf
                strSQL = strSQL & " ,Clientes.telefoneComercial " & vbCrLf
                strSQL = strSQL & " ,PedidoLocacao.TipoPedidoLocacao " & vbCrLf
                strSQL = strSQL & " ,Bancos.idBancos " & vbCrLf
                strSQL = strSQL & " ,Bancos.nomeBanco " & vbCrLf
                strSQL = strSQL & " FROM PedidoLocacao " & vbCrLf
                strSQL = strSQL & " INNER JOIN Clientes ON PedidoLocacao.idClientes = Clientes.idClientes" & vbCrLf
                strSQL = strSQL & " INNER JOIN Cacambas ON PedidoLocacao.idCacambas = Cacambas.idCacambas" & vbCrLf
                strSQL = strSQL & " INNER JOIN Usuarios ON PedidoLocacao.idUsuarios = Usuarios.idUsuarios" & vbCrLf
                strSQL = strSQL & " LEFT JOIN Bancos ON PedidoLocacao.idBancos = Bancos.idBancos" & vbCrLf
                strSQL = strSQL & " WHERE idPedidoLocacao = " & pIdPedidoLocacao
        End If
        
        Set rsPedidoLocacao = New ADODB.Recordset
        
        With rsPedidoLocacao
                .Open strSQL, dbCacamba, adOpenStatic, adLockReadOnly
                If Not .EOF Then
                        vIdPedidoLocacao = !idPedidoLocacao
                        txtCliente.Text = NVL(!nome, "")
                        vIdClientePedidoLocacao = NVL(!idClientes, 0)
                        txtUsuario.Text = NVL(!NomeUsuario, "")
                        vIdCacambas = NVL(!idCacambas, 0)
                        txtNumeroCacamba.Text = NVL(!NumeroCacamba, "")
                        txtNumeroControle.Text = NVL(!numeroControle, "")
                        mskData = NVL(!DataLocacao, "")
                        mskDataEntrega.Text = NVL(!dataEntrega, "")
                        txtPrevisaoRetirada.Text = NVL(!PrevisaoRetirada, "")
                        mskDataRealRetirada.Text = NVL(!dataRealRetirada, "__/__/____")
                        
                        txtEnderecoNumero.Text = NVL(!EnderecoNumero, "")
                        txtEnderecoBairro.Text = NVL(!Bairro, "")
                        txtEnderecoCidade.Text = NVL(!Cidade, "")
                        cboEnderecoEstado.Text = NVL(!UF, "")
                        txtEnderecoComplemento.Text = NVL(!Complemento, "")
                        txtEndereco.Text = NVL(!Endereco, "")
                        mskValorServico = NVL(!ValorServico, 0)
                        mskValorDesconto = NVL(!ValorDesconto, 0)
                        mskValorTotal = NVL(!ValorTotal, 0)
                        cboFormaPagamento.Text = NVL(!FormaPagamento, "")
                        vOperacao = consulta
                        cboSituacao.Text = NVL(!Situacao, "")
                        
                        txtEmailCliente.Text = NVL(!Email, "")
                        txtTelefoneCelular.Text = NVL(!TelefoneCelular, "")
                        txtTelefoneComercial.Text = NVL(!TelefoneComercial, "")
                        txtTelefoneFixo.Text = NVL(!TelefoneFixo, "")
                        mskDataPagamento.Text = NVL(!DataPagamento, "__/__/____")
                        If IsDate(mskDataRealRetirada.FormattedText) Then
                                txtTotalDiasAlocado.Text = DateDiff("d", mskData.FormattedText, mskDataRealRetirada.FormattedText)
                        End If
                        If NVL(!TipoPedidoLocacao, 0) = 0 Then
                            optAluguel.Value = True
                        Else
                            optTroca.Value = True
                        End If
                        
                        vIdBancos = NVL(!idBancos, 0)
                        txtBanco.Text = NVL(!nomeBanco, "")
                        
                End If
        End With
End Sub

Private Function validaGravacao() As Boolean

        validaGravacao = False
        
        If Trim(txtNumeroCacamba.Text) = "" Or vIdCacambas = 0 Then
                MsgBox "Selecione a caçamba!", vbInformation, "Caçambas"
                cmdSelecionaCacamba.SetFocus
                Exit Function
        End If
        
        If Trim(txtCliente.Text) = "" Or vIdClientePedidoLocacao = 0 Then
                MsgBox "Infomre o cliente!", vbInformation, "Caçambas"
                cmdSelecionaCliente.SetFocus
                Exit Function
        End If
        
        If Trim(txtNumeroControle.Text) = "" Then
                MsgBox "Infomre o número da caçamba!", vbInformation, "Caçambas"
                txtNumeroCacamba.SetFocus
                Exit Function
        End If
        
        If Not IsDate(mskData) Then
                MsgBox "Infomre a data de locação!", vbInformation, "Caçambas"
                mskData.SetFocus
                Exit Function
        End If
        
        If Trim(txtPrevisaoRetirada.Text) = "" Then
                MsgBox "Infomre a previsão de retirada!", vbInformation, "Caçambas"
                If txtNumeroCacamba.Enabled Then txtNumeroCacamba.SetFocus
                Exit Function
        End If
        
        If Trim(txtEnderecoNumero.Text) = "" Then
                MsgBox "Infomre o número do endereço!", vbInformation, "Caçambas"
                txtEnderecoNumero.SetFocus
                Exit Function
        End If
        
        If Trim(txtEnderecoBairro.Text) = "" Then
                MsgBox "Infomre o bairro!", vbInformation, "Caçambas"
                txtEnderecoBairro.SetFocus
                Exit Function
        End If
        
        If Trim(txtEnderecoCidade.Text) = "" Then
                MsgBox "Infomre o bairro!", vbInformation, "Caçambas"
                txtEnderecoCidade.SetFocus
                Exit Function
        End If
        
        If Trim(cboEnderecoEstado.Text) = "" Then
                MsgBox "Infomre o bairro!", vbInformation, "Caçambas"
                cboEnderecoEstado.SetFocus
                Exit Function
        End If
        
        If Val(mskValorServico.FormattedText) = 0 Then
                MsgBox "Informe o valor do serviço!", vbInformation, "Caçambas"
                mskValorServico.SetFocus
                Exit Function
        End If
        
        If Trim(cboFormaPagamento.Text) = "" Then
                MsgBox "Infomre a forma de pagamento!", vbInformation, "Caçambas"
                cboFormaPagamento.SetFocus
                Exit Function
        End If
        
        If Trim(cboSituacao.Text) = "" Then
                MsgBox "Infomre a situação!", vbInformation, "Caçambas"
                cboSituacao.SetFocus
                Exit Function
        End If
        
        If UCase(cboSituacao.Text) = UCase("Pago na Entrega") Or UCase(cboSituacao.Text) = "Pago na Retirada" Then
                If mskDataPagamento.Text = "__/__/____" Then
                        MsgBox "Informe a data de pagamento!", vbInformation, "Caçambas"
                        mskDataPagamento.SetFocus
                        Exit Function
                End If
        End If
        
        
        validaGravacao = True
        
End Function

Private Sub txtNumeroControle_KeyPress(KeyAscii As Integer)
                
        KeyAscii = SoNumeros(KeyAscii)
        
        If KeyAscii = 0 Then
                Exit Sub
        End If
End Sub

Private Sub txtNumeroControle_LostFocus()
        Dim vNumeroControle                             As Integer
        
        If Trim(txtNumeroControle.Text) = "" Then Exit Sub
        
        vNumeroControle = NVL(recuperaDescricao("PedidoLocacao", "idPedidoLocacao", "NumeroControle", txtNumeroControle.Text, True), 0)
        
        If vNumeroControle <> 0 Then
                MsgBox "Número de Controle já cadastrado!", vbInformation, "Caçambas"
                txtNumeroControle.Text = ""
                txtNumeroControle.SetFocus
                Exit Sub
        End If
End Sub

Private Sub txtPrevisaoRetirada_KeyPress(KeyAscii As Integer)
                
        KeyAscii = SoNumeros(KeyAscii)
        
        If KeyAscii = 0 Then
                Exit Sub
        End If
End Sub

Private Sub limpaCampos()
        
        vIdClientePedidoLocacao = 0
        txtCliente.Text = ""
        txtNumeroControle.Text = ""
        txtNumeroCacamba.Text = ""
        vIdCacambas = 0
        mskData.Text = Format(Now, "dd/mm/yyyy")
        mskDataRealRetirada.Text = "__/__/____"
        mskDataPagamento.Text = "__/__/____"
        txtPrevisaoRetirada.Text = ""
        txtEmailCliente.Text = ""
        txtTelefoneCelular.Text = ""
        txtTelefoneComercial.Text = ""
        txtTelefoneFixo.Text = ""
        txtEndereco.Text = ""
        txtEnderecoNumero.Text = ""
        txtEnderecoBairro.Text = ""
        txtEnderecoComplemento.Text = ""
        cboEnderecoEstado.ListIndex = 0
        mskValorServico = 0
        mskValorDesconto = 0
        mskValorTotal = 0
        cboFormaPagamento.ListIndex = 0
        cboSituacao.ListIndex = 0
        vIdBancos = 0
        txtBanco.Text = ""
        
End Sub

Private Sub cmdLocalizar_Click()
        Dim strSQL                      As String
        Dim vColunas                    As String
        
        strSQL = strSQL & " SELECT " & vbCrLf
        strSQL = strSQL & " idPedidoLocacao " & vbCrLf
        strSQL = strSQL & " ,Clientes.Nome " & vbCrLf
        strSQL = strSQL & " ,Cacambas.Numero as NumeroCacamba " & vbCrLf
        strSQL = strSQL & " ,NumeroControle " & vbCrLf
        strSQL = strSQL & " ,DataLocacao " & vbCrLf
        strSQL = strSQL & " ,PrevisaoRetirada " & vbCrLf
        strSQL = strSQL & " ,DataRealRetirada " & vbCrLf
        strSQL = strSQL & " ,ValorServico " & vbCrLf
        strSQL = strSQL & " ,ValorDesconto " & vbCrLf
        strSQL = strSQL & " ,ValorTotal " & vbCrLf
        strSQL = strSQL & " ,FormaPagamento " & vbCrLf
        strSQL = strSQL & " ,Situacao " & vbCrLf
        strSQL = strSQL & " FROM PedidoLocacao " & vbCrLf
        strSQL = strSQL & " INNER JOIN Clientes ON PedidoLocacao.idClientes = Clientes.idClientes" & vbCrLf
        strSQL = strSQL & " INNER JOIN Cacambas ON PedidoLocacao.idCacambas = Cacambas.idCacambas" & vbCrLf
        strSQL = strSQL & " INNER JOIN Usuarios ON PedidoLocacao.idUsuarios = Usuarios.idUsuarios" & vbCrLf
        strSQL = strSQL & " ORDER BY idPedidoLocacao DESC "
        
        vColunas = "Código,Cliente, Nº Caçamba, Nº Controle, Data Locação, Previsão Retirada, Data Real Retirada, Vlr Serviço, Vlr Desconto, Vlr Total,Forma Pagamento, Situação"
        
        frmPesquisa.carregaGridPesquisa 12, strSQL, vColunas, 1
        frmPesquisa.Show vbModal
        
        If vgRetornoConsulta <> 0 Then
                vIdPedidoLocacao = vgRetornoConsulta
                carregaPedidoLocacao vIdPedidoLocacao
        End If
        
        
End Sub

Public Sub cmdIncluir_Click()
        controlaBotoes Me, operacao.Inclusao
        fraPedidoLocacao.Enabled = True
        limpaCampos
        vIdPedidoLocacao = 0
        mskDataRealRetirada.Enabled = False
        vOperacao = Inclusao
        optAluguel.Value = True
        If txtNumeroControle.Visible Then txtNumeroControle.SetFocus
End Sub

Private Sub cmdGravar_Click()
        Dim strSQL                     As String
        
        If Not validaGravacao Then Exit Sub
        
        strSQL = Empty
        
        If vOperacao = Inclusao Then
        
                strSQL = strSQL & " INSERT INTO PedidoLocacao " & vbCrLf
                strSQL = strSQL & " (idCacambas " & vbCrLf
                strSQL = strSQL & " ,idClientes " & vbCrLf
                strSQL = strSQL & " ,idUsuarios " & vbCrLf
                strSQL = strSQL & " ,numeroControle " & vbCrLf
                strSQL = strSQL & " ,dataLocacao " & vbCrLf
                strSQL = strSQL & " ,previsaoRetirada " & vbCrLf
                strSQL = strSQL & " ,Endereco " & vbCrLf
                strSQL = strSQL & " ,EnderecoNumero " & vbCrLf
                strSQL = strSQL & " ,Bairro " & vbCrLf
                strSQL = strSQL & " ,Cidade " & vbCrLf
                strSQL = strSQL & " ,UF " & vbCrLf
                strSQL = strSQL & " ,Complemento " & vbCrLf
                strSQL = strSQL & " ,ValorServico " & vbCrLf
                strSQL = strSQL & " ,ValorDesconto " & vbCrLf
                strSQL = strSQL & " ,ValorTotal " & vbCrLf
                strSQL = strSQL & " ,FormaPagamento " & vbCrLf
                strSQL = strSQL & " ,Situacao " & vbCrLf
                strSQL = strSQL & " ,DataUltimaAlteracao " & vbCrLf
                strSQL = strSQL & " ,TipoPedidoLocacao " & vbCrLf
                If IsDate(mskDataPagamento.FormattedText) Then
                        strSQL = strSQL & " ,DataPagamento " & vbCrLf
                End If
                
                If IsDate(mskDataEntrega.FormattedText) Then
                        strSQL = strSQL & " ,DataEntrega " & vbCrLf
                End If
                
                If vIdBancos <> 0 Then
                    strSQL = strSQL & " , idBancos " & vbCrLf
                End If
                strSQL = strSQL & ") VALUES ( "
                strSQL = strSQL & vIdCacambas & vbCrLf
                strSQL = strSQL & "," & vIdClientePedidoLocacao & vbCrLf
                strSQL = strSQL & "," & vgIdUsuarioLogado & vbCrLf
                strSQL = strSQL & "," & txtNumeroControle & vbCrLf
                strSQL = strSQL & ",'" & Format(mskData.FormattedText, "yyyymmdd") & "'" & vbCrLf
                strSQL = strSQL & "," & txtPrevisaoRetirada.Text & vbCrLf
                strSQL = strSQL & ",'" & Replace(txtEndereco.Text, "'", "''") & "'" & vbCrLf
                strSQL = strSQL & ",'" & txtEnderecoNumero.Text & "'" & vbCrLf
                strSQL = strSQL & ",'" & Replace(txtEnderecoBairro.Text, "'", "''") & "'" & vbCrLf
                strSQL = strSQL & ",'" & Replace(txtEnderecoCidade.Text, "'", "''") & "'" & vbCrLf
                strSQL = strSQL & ",'" & cboEnderecoEstado.Text & "'" & vbCrLf
                strSQL = strSQL & ",'" & Replace(txtEnderecoComplemento.Text, "'", "''") & "'" & vbCrLf
                strSQL = strSQL & "," & Replace(mskValorServico, ",", ".") & vbCrLf
                strSQL = strSQL & "," & Replace(mskValorDesconto, ",", ".") & vbCrLf
                strSQL = strSQL & "," & Replace(mskValorTotal, ",", ".") & vbCrLf
                strSQL = strSQL & ",'" & cboFormaPagamento.Text & "'" & vbCrLf
                strSQL = strSQL & ",'" & cboSituacao.Text & "'" & vbCrLf
                strSQL = strSQL & ",'" & Format(Now, "yyyy-mm-dd HH:MM:ss") & "'" & vbCrLf
                strSQL = strSQL & "," & IIf(optAluguel.Value, e_TipoPedidoLocacao.aluguel, e_TipoPedidoLocacao.troca) & vbCrLf
                If IsDate(mskDataPagamento.FormattedText) Then
                        strSQL = strSQL & ",'" & Format(mskDataPagamento.FormattedText, "yyyy-mm-dd HH:MM:ss") & "'" & vbCrLf
                End If
                If IsDate(mskDataEntrega.FormattedText) Then
                        strSQL = strSQL & ",'" & Format(mskDataEntrega.FormattedText, "yyyy-mm-dd") & "'" & vbCrLf
                End If
                If vIdBancos <> 0 Then
                    strSQL = strSQL & " ," & vIdBancos & vbCrLf
                End If
                strSQL = strSQL & ")" & vbCrLf
        
        Else
                        
                strSQL = strSQL & " UPDATE PedidoLocacao " & vbCrLf
                strSQL = strSQL & " SET " & vbCrLf
                strSQL = strSQL & " idCacambas = " & vIdCacambas & vbCrLf
                strSQL = strSQL & " ,idClientes = " & vIdClientePedidoLocacao & vbCrLf
                strSQL = strSQL & " ,idUsuarios = " & vgIdUsuarioLogado & vbCrLf
                strSQL = strSQL & " ,NumeroControle = " & txtNumeroControle.Text & vbCrLf
                strSQL = strSQL & " ,dataLocacao = '" & Format(mskData.FormattedText, "yyyymmdd") & "'" & vbCrLf
                strSQL = strSQL & " ,PrevisaoRetirada = " & txtPrevisaoRetirada.Text & vbCrLf
                strSQL = strSQL & " ,EnderecoNumero = '" & txtEnderecoNumero.Text & "'" & vbCrLf
                strSQL = strSQL & " ,Endereco = '" & Replace(txtEndereco.Text, "'", "''") & "'" & vbCrLf
                strSQL = strSQL & " ,EnderecoNumero = '" & txtEnderecoNumero.Text & "'" & vbCrLf
                strSQL = strSQL & " ,Bairro = '" & Replace(txtEnderecoBairro.Text, "'", "''") & "'" & vbCrLf
                strSQL = strSQL & " ,Cidade = '" & Replace(txtEnderecoCidade.Text, "'", "''") & "'" & vbCrLf
                strSQL = strSQL & " ,UF = '" & cboEnderecoEstado.Text & "'" & vbCrLf
                strSQL = strSQL & " ,Complemento = '" & Replace(txtEnderecoComplemento.Text, "'", "''") & "'" & vbCrLf
                strSQL = strSQL & ", ValorServico = " & Replace(mskValorServico, ",", ".") & vbCrLf
                strSQL = strSQL & ", ValorDesconto =" & Replace(mskValorDesconto, ",", ".") & vbCrLf
                strSQL = strSQL & ", ValorTotal = " & Replace(mskValorTotal, ",", ".") & vbCrLf
                strSQL = strSQL & ", FormaPagamento = '" & cboFormaPagamento.Text & "'" & vbCrLf
                strSQL = strSQL & ", Situacao = '" & cboSituacao.Text & "'" & vbCrLf
                strSQL = strSQL & ", TipoPedidoLocacao = " & IIf(optAluguel.Value, e_TipoPedidoLocacao.aluguel, e_TipoPedidoLocacao.troca)
                If IsDate(mskDataRealRetirada.FormattedText) Then
                        strSQL = strSQL & " ,dataRealRetirada = '" & Format(mskDataRealRetirada.FormattedText, "yyyymmdd") & "'" & vbCrLf
                End If
                strSQL = strSQL & " ,dataUltimaAlteracao = '" & Format(Now, "yyyy-mm-dd HH:MM:ss") & "'" & vbCrLf
                If IsDate(mskDataPagamento.FormattedText) Then
                        strSQL = strSQL & " ,dataPagamento = '" & Format(mskDataPagamento.FormattedText, "yyyymmdd") & "'" & vbCrLf
                End If
                If IsDate(mskDataEntrega.FormattedText) Then
                        strSQL = strSQL & " ,dataEntrega = '" & Format(mskDataEntrega.FormattedText, "yyyymmdd") & "'" & vbCrLf
                End If
                If vIdBancos <> 0 Then
                    strSQL = strSQL & " , idBancos = " & vIdBancos & vbCrLf
                End If
                strSQL = strSQL & " WHERE idPedidoLocacao = " & vIdPedidoLocacao
        
        End If
        
        dbCacamba.Execute strSQL


        controlaBotoes Me, operacao.consulta
        vOperacao = consulta
        fraPedidoLocacao.Enabled = False
End Sub

Private Sub cmdExcluir_Click()
        Dim strSQL                      As String
        
        If MsgBox("Confirma exlusão?", vbQuestion + vbYesNo + vbDefaultButton2, "Caçambas") = vbNo Then
                Exit Sub
        End If
        
        strSQL = Empty
        strSQL = strSQL & "DELETE FROM PedidoLocacao WHERE idPedidoLocacao = " & vIdPedidoLocacao
        
        dbCacamba.Execute strSQL
        
        controlaBotoes Me, operacao.consulta
        fraPedidoLocacao.Enabled = False
        carregaPedidoLocacao
        vOperacao = consulta
End Sub


Private Sub cmdCancelar_Click()
        controlaBotoes Me, operacao.consulta
        fraPedidoLocacao.Enabled = False
        vOperacao = consulta
End Sub

Private Sub cmdAlterar_Click()
        controlaBotoes Me, operacao.Inclusao
        fraPedidoLocacao.Enabled = True
        vOperacao = Alteracao
        mskDataRealRetirada.Enabled = True
        
        If UCase(cboSituacao.Text) = UCase("Pago na Entrega") Or UCase(cboSituacao.Text) = UCase("Pago da Retirada") Then
                If vOperacao <> consulta Then
                        mskDataPagamento.Enabled = True
                End If
        Else
                mskDataPagamento.Enabled = False
                mskDataPagamento.Text = "__/__/____"
        End If
End Sub

Private Function verificaCacambaAlocada(pIdCacamba As Integer) As Integer
        
        Dim strSQL                      As String
        Dim rsCacamba                   As ADODB.Recordset
        
        strSQL = Empty
        strSQL = strSQL & " SELECT NumeroControle " & vbCrLf
        strSQL = strSQL & " FROM PedidoLocacao" & vbCrLf
        strSQL = strSQL & " WHERE idCacambas = " & pIdCacamba
        strSQL = strSQL & " AND DataRealRetirada IS NULL"
        
        verificaCacambaAlocada = 0
        
        Set rsCacamba = New ADODB.Recordset
        
        With rsCacamba
                .Open strSQL, dbCacamba, adOpenStatic, adLockReadOnly
                If Not .EOF Then
                        verificaCacambaAlocada = !numeroControle
                End If
        End With

        
End Function



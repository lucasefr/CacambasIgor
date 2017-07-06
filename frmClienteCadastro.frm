VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmClienteCadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10710
   Icon            =   "frmClienteCadastro.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGerarPedidoLocacao 
      Caption         =   "Gerar Pedido"
      Height          =   735
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   735
      Left            =   4880
      TabIndex        =   17
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   735
      Left            =   9600
      TabIndex        =   21
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   735
      Left            =   8420
      TabIndex        =   20
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Height          =   735
      Left            =   7240
      TabIndex        =   19
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   735
      Left            =   6060
      TabIndex        =   18
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   735
      Left            =   3700
      TabIndex        =   15
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "Localizar"
      Height          =   735
      Left            =   2520
      TabIndex        =   16
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Frame fraCadastroCliente 
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10695
      Begin MSMask.MaskEdBox mskCNPJCPF 
         Height          =   390
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   688
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtTelefoneCelular 
         Height          =   375
         Left            =   4440
         MaxLength       =   14
         TabIndex        =   8
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox txtTelefoneComercial 
         Height          =   375
         Left            =   2280
         MaxLength       =   14
         TabIndex        =   7
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox txtEnderecoComplemento 
         Height          =   375
         Left            =   1200
         MaxLength       =   45
         TabIndex        =   14
         Top             =   5880
         Width           =   3375
      End
      Begin VB.ComboBox cboEnderecoEstado 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   5880
         Width           =   855
      End
      Begin VB.TextBox txtEnderecoCidade 
         Height          =   375
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   12
         Top             =   5160
         Width           =   3375
      End
      Begin VB.TextBox txtEnderecoBairro 
         Height          =   375
         Left            =   120
         MaxLength       =   45
         TabIndex        =   11
         Top             =   5160
         Width           =   3375
      End
      Begin VB.TextBox txtEnderecoNumero 
         Height          =   375
         Left            =   7680
         MaxLength       =   15
         TabIndex        =   10
         Top             =   4440
         Width           =   1695
      End
      Begin VB.TextBox txtEndereco 
         Height          =   375
         Left            =   120
         MaxLength       =   150
         TabIndex        =   9
         Top             =   4440
         Width           =   7455
      End
      Begin VB.TextBox txtEmailCliente 
         Height          =   375
         Left            =   120
         MaxLength       =   150
         TabIndex        =   5
         Top             =   2880
         Width           =   6855
      End
      Begin VB.TextBox txtTelefoneFixo 
         Height          =   375
         Left            =   120
         MaxLength       =   14
         TabIndex        =   6
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox txtNomeCliente 
         Height          =   375
         Left            =   120
         MaxLength       =   100
         TabIndex        =   4
         Top             =   2160
         Width           =   9255
      End
      Begin VB.Frame fraTipoPessoa 
         Caption         =   "Tipo de Pessoa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2895
         Begin VB.OptionButton optPessoaJuridica 
            Caption         =   "Jurídica"
            Height          =   495
            Left            =   1440
            TabIndex        =   2
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optPessoaFisica 
            Caption         =   "Fisica"
            Height          =   495
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
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
         TabIndex        =   35
         Top             =   3480
         Width           =   1815
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
         TabIndex        =   34
         Top             =   5640
         Width           =   2055
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
         TabIndex        =   33
         Top             =   5640
         Width           =   975
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
         TabIndex        =   32
         Top             =   4920
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
         TabIndex        =   31
         Top             =   4920
         Width           =   735
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
         TabIndex        =   30
         Top             =   4200
         Width           =   735
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
         TabIndex        =   29
         Top             =   4200
         Width           =   975
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
         TabIndex        =   28
         Top             =   2640
         Width           =   2055
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
         TabIndex        =   26
         Top             =   3480
         Width           =   3255
      End
      Begin VB.Label lblNomeCliente 
         Caption         =   "Nome"
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
         TabIndex        =   25
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblCNPJCPF 
         Caption         =   "CPF"
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
         Top             =   1200
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Telefone"
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   3840
      Width           =   975
   End
End
Attribute VB_Name = "frmClienteCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private vOperacao               As operacao
Private vIdCliente              As Integer

Private Enum maxLenghtTipoPessoa
        Juridica = 14
        Fisica = 11
End Enum

Private Sub cmdAlterar_Click()
        controlaBotoes Me, operacao.Inclusao
        fraCadastroCliente.Enabled = True
        vOperacao = Alteracao
        cmdGerarPedidoLocacao.Enabled = False
End Sub

Private Sub cmdCancelar_Click()
        controlaBotoes Me, operacao.consulta
        fraCadastroCliente.Enabled = False
        vOperacao = consulta
        cmdGerarPedidoLocacao.Enabled = True
End Sub

Private Sub cmdExcluir_Click()
        Dim strSQL                      As String
        
        If MsgBox("Confirma exlusão?", vbQuestion + vbYesNo + vbDefaultButton2, "Caçambas") = vbNo Then
                Exit Sub
        End If
        
        strSQL = Empty
        strSQL = strSQL & "DELETE FROM Clientes WHERE idUsuarios = " & vIdCliente
        
        dbCacamba.Execute strSQL
        
        controlaBotoes Me, operacao.consulta
        fraCadastroCliente.Enabled = False
        carregaCliente
        vOperacao = consulta
        cmdGerarPedidoLocacao.Enabled = True
End Sub

Private Sub cmdGerarPedidoLocacao_Click()

        If vIdCliente = 0 Then
                MsgBox "Selecione o cliente!", vbInformation, "Caçambas"
                Exit Sub
        End If
        
        With frmPedidoLocacao
                .cmdIncluir_Click
                .vIdClientePedidoLocacao = vIdCliente
                .carregaInformacaoCliente vIdCliente
                .Show vbModal
        End With
End Sub

Private Sub cmdGravar_Click()
        Dim strSQL                     As String
        
        If Not validaGravacao Then Exit Sub
        
        strSQL = Empty
        
        If vOperacao = Inclusao Then
        
                strSQL = strSQL & " INSERT INTO Clientes " & vbCrLf
                strSQL = strSQL & " (nome " & vbCrLf
                strSQL = strSQL & " ,TipoCliente " & vbCrLf
                strSQL = strSQL & " ,CnpjCPF " & vbCrLf
                strSQL = strSQL & " ,email " & vbCrLf
                strSQL = strSQL & " ,telefoneFixo " & vbCrLf
                strSQL = strSQL & " ,telefoneComercial " & vbCrLf
                strSQL = strSQL & " ,telefoneCelular " & vbCrLf
                strSQL = strSQL & " ,Endereco " & vbCrLf
                strSQL = strSQL & " ,Numero " & vbCrLf
                strSQL = strSQL & " ,Bairro " & vbCrLf
                strSQL = strSQL & " ,Cidade " & vbCrLf
                strSQL = strSQL & " ,UF " & vbCrLf
                strSQL = strSQL & " ,Complemento " & vbCrLf
                strSQL = strSQL & " ,DataCadastro " & vbCrLf
                strSQL = strSQL & ") VALUES ( "
                strSQL = strSQL & "'" & Replace(txtNomeCliente.Text, "'", "''") & "'" & vbCrLf
                
                If optPessoaFisica.Value Then
                        strSQL = strSQL & "," & tipoPessoa.Fisica & vbCrLf
                Else
                        strSQL = strSQL & "," & tipoPessoa.Juridica & vbCrLf
                End If
                
                strSQL = strSQL & ",'" & GetOnlyNumbers(mskCNPJCPF.Text) & "'" & vbCrLf
                strSQL = strSQL & ",'" & txtEmailCliente.Text & "'" & vbCrLf
                strSQL = strSQL & ",'" & txtTelefoneFixo.Text & "'" & vbCrLf
                strSQL = strSQL & ",'" & txtTelefoneComercial.Text & "'" & vbCrLf
                strSQL = strSQL & ",'" & txtTelefoneCelular.Text & "'" & vbCrLf
                strSQL = strSQL & ",'" & Replace(txtEndereco.Text, "'", "''") & "'" & vbCrLf
                strSQL = strSQL & ",'" & txtEnderecoNumero.Text & "'" & vbCrLf
                strSQL = strSQL & ",'" & Replace(txtEnderecoBairro.Text, "'", "''") & "'" & vbCrLf
                strSQL = strSQL & ",'" & Replace(txtEnderecoCidade.Text, "'", "''") & "'" & vbCrLf
                strSQL = strSQL & ",'" & cboEnderecoEstado.Text & "'" & vbCrLf
                strSQL = strSQL & ",'" & Replace(txtEnderecoComplemento.Text, "'", "''") & "'" & vbCrLf
                strSQL = strSQL & ",'" & Format(Now, "yyyymmdd") & "'" & vbCrLf
                strSQL = strSQL & ")" & vbCrLf
        
        Else
                        
                strSQL = strSQL & " UPDATE Clientes " & vbCrLf
                strSQL = strSQL & " SET " & vbCrLf
                strSQL = strSQL & " nome = '" & Replace(txtNomeCliente.Text, "'", "''") & "'" & vbCrLf
                If optPessoaFisica.Value Then
                        strSQL = strSQL & " ,TipoCliente = " & tipoPessoa.Fisica & vbCrLf
                Else
                        strSQL = strSQL & " ,TipoCliente = " & tipoPessoa.Juridica & vbCrLf
                End If
                strSQL = strSQL & " ,CNPJCPF = '" & GetOnlyNumbers(mskCNPJCPF.Text) & "'" & vbCrLf
                strSQL = strSQL & " ,Email = '" & txtEmailCliente.Text & "'" & vbCrLf
                strSQL = strSQL & " ,TelefoneFixo = '" & txtTelefoneFixo.Text & "'" & vbCrLf
                strSQL = strSQL & " ,TelefoneComercial = '" & txtTelefoneComercial.Text & "'" & vbCrLf
                strSQL = strSQL & " ,TelefoneCelular = '" & txtTelefoneCelular.Text & "'" & vbCrLf
                strSQL = strSQL & " ,Endereco = '" & Replace(txtEndereco.Text, "'", "''") & "'" & vbCrLf
                strSQL = strSQL & " ,Numero = '" & txtEnderecoNumero.Text & "'" & vbCrLf
                strSQL = strSQL & " ,Bairro = '" & Replace(txtEnderecoBairro.Text, "'", "''") & "'" & vbCrLf
                strSQL = strSQL & " ,Cidade = '" & Replace(txtEnderecoCidade.Text, "'", "''") & "'" & vbCrLf
                strSQL = strSQL & " ,UF = '" & cboEnderecoEstado.Text & "'" & vbCrLf
                strSQL = strSQL & " ,Complemento = '" & Replace(txtEnderecoComplemento.Text, "'", "''") & "'" & vbCrLf
                strSQL = strSQL & " WHERE idClientes = " & vIdCliente
        
        End If
        
        dbCacamba.Execute strSQL

        carregaCliente 0
        controlaBotoes Me, operacao.consulta
        vOperacao = consulta
        cmdGerarPedidoLocacao.Enabled = True
        cmdGerarPedidoLocacao.SetFocus
        fraCadastroCliente.Enabled = False
End Sub

Private Sub cmdIncluir_Click()
        controlaBotoes Me, operacao.Inclusao
        fraCadastroCliente.Enabled = True
        limpaCampos
        vIdCliente = 0
        vOperacao = Inclusao
        optPessoaFisica.SetFocus
        cmdGerarPedidoLocacao.Enabled = False
        txtEnderecoCidade.Text = "Nova Lima"
End Sub

Private Sub cmdLocalizar_Click()
        
        frmClientePesquisa.Show vbModal
        
        If vgRetornoConsulta <> 0 Then
                vIdCliente = vgRetornoConsulta
                carregaCliente vIdCliente
        End If
        
        
End Sub

Private Sub cmdSair_Click()
        Unload Me
End Sub

Private Sub Form_Load()
        controlaBotoes Me, operacao.consulta
        CarregaComboUf cboEnderecoEstado
        fraCadastroCliente.Enabled = False
        carregaCliente
        
End Sub


Private Sub mskCNPJCPF_LostFocus()
        If optPessoaFisica.Value Then
                If EntradaCPFinvalida(mskCNPJCPF) Then
                        mskCNPJCPF.Text = "___.___.___-__"
                        If mskCNPJCPF.Enabled Then mskCNPJCPF.SetFocus
                        Exit Sub
                End If
        Else
                If EntradaCNPJinvalida(mskCNPJCPF) Then
                        mskCNPJCPF.Text = "__.___.___/____-__"
                        If mskCNPJCPF.Enabled Then mskCNPJCPF.SetFocus
                        Exit Sub
                End If
        End If
End Sub

Private Sub optPessoaFisica_Click()
        lblCNPJCPF.Caption = "CPF"
        mskCNPJCPF.Mask = "###.###.###-##"
        mskCNPJCPF.Format = "###.###.###-##"
End Sub

Private Sub optPessoaJuridica_Click()
        lblCNPJCPF.Caption = "CNPJ"
        mskCNPJCPF.Mask = "##.###.###/####-##"
        mskCNPJCPF.Format = "##.###.###/####-##"
End Sub

Private Sub limpaCampos()
        
        txtNomeCliente.Text = ""
        'txtCNPJCPF.Text = ""
        mskCNPJCPF.Text = ""
        txtEmailCliente.Text = ""
        txtTelefoneCelular.Text = ""
        txtTelefoneComercial.Text = ""
        txtTelefoneFixo.Text = ""
        txtEndereco.Text = ""
        txtEnderecoNumero.Text = ""
        txtEnderecoBairro.Text = ""
        txtEnderecoComplemento.Text = ""
        cboEnderecoEstado.ListIndex = 0
        txtEnderecoCidade.Text = ""
        optPessoaFisica.Value = True
End Sub

Private Function validaGravacao() As Boolean

        validaGravacao = False
        
        If Trim(txtNomeCliente.Text) = "" Then
                MsgBox "Informe o nome do cliente!", vbInformation, "Caçambas"
                Exit Function
        End If
        
        If Trim(mskCNPJCPF.Text) = "" Then
                MsgBox "Infomre o CNPJ ou CPF do usuário!", vbInformation, "Caçambas"
                Exit Function
        End If
        
        validaGravacao = True
        
End Function


Private Sub carregaCliente(Optional pIdCliente As Integer)

        Dim strSQL                      As String
        Dim rsCliente                   As ADODB.Recordset
        
        strSQL = Empty
        If pIdCliente = 0 Then
                strSQL = strSQL & " SELECT * FROM Clientes ORDER BY idClientes DESC LIMIT 1"
                
        Else
                strSQL = strSQL & " SELECT * FROM Clientes " & vbCrLf
                strSQL = strSQL & " WHERE idClientes = " & pIdCliente
        End If
        
        Set rsCliente = New ADODB.Recordset
        
        With rsCliente
                .Open strSQL, dbCacamba, adOpenStatic, adLockReadOnly
                If Not .EOF Then
                        vIdCliente = !idClientes
                        txtNomeCliente.Text = !nome
                        If !tipoCliente = tipoPessoa.Fisica Then
                                optPessoaFisica.Value = True
                        Else
                                optPessoaJuridica.Value = True
                        End If
                        mskCNPJCPF.Text = !CNPJCPF
                        txtEmailCliente.Text = NVL(!Email, "")
                        txtTelefoneCelular.Text = NVL(!TelefoneCelular, "")
                        txtTelefoneComercial.Text = NVL(!TelefoneComercial, "")
                        txtTelefoneFixo.Text = NVL(!TelefoneFixo, "")
                        txtEndereco.Text = NVL(!Endereco, "")
                        txtEnderecoNumero.Text = NVL(!Numero, "")
                        txtEnderecoBairro.Text = NVL(!Bairro, "")
                        txtEnderecoComplemento.Text = NVL(!Complemento, "")
                        cboEnderecoEstado.Text = NVL(!UF, "")
                        
                End If
        End With
End Sub

Private Sub txtCNPJCPF_LostFocus()
        If optPessoaFisica.Value Then
                If EntradaCPFinvalida(mskCNPJCPF) Then
                        mskCNPJCPF.Text = ""
                        If mskCNPJCPF.Enabled Then mskCNPJCPF.SetFocus
                        Exit Sub
                End If
        Else
                If EntradaCNPJinvalida(mskCNPJCPF) Then
                        mskCNPJCPF.Text = ""
                        If mskCNPJCPF.Enabled Then mskCNPJCPF.SetFocus
                        Exit Sub
                End If
        End If
End Sub

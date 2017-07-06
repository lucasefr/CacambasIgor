VERSION 5.00
Begin VB.Form frmUsuarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Usuarios"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8355
   Icon            =   "frmUsuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraCadastroUsuarios 
      Height          =   2655
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   8175
      Begin VB.TextBox txtSenha 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2160
         Width           =   2655
      End
      Begin VB.TextBox txtLogin 
         Height          =   375
         Left            =   120
         MaxLength       =   12
         TabIndex        =   2
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtNome 
         Height          =   375
         Left            =   120
         MaxLength       =   45
         TabIndex        =   1
         Top             =   600
         Width           =   6975
      End
      Begin VB.Label Senha 
         Caption         =   "Senha"
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
         TabIndex        =   13
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Login"
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
         TabIndex        =   12
         Top             =   1200
         Width           =   1455
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
         TabIndex        =   11
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "Localizar"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   735
      Left            =   1305
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   735
      Left            =   3660
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Height          =   735
      Left            =   4845
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   735
      Left            =   6015
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   735
      Left            =   7200
      TabIndex        =   5
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   735
      Left            =   2475
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vOperacao                   As operacao
Dim vIdUsuario                  As Integer

Private Sub cmdAlterar_Click()
        controlaBotoes Me, operacao.Alteracao
        fraCadastroUsuarios.Enabled = True
        vOperacao = Alteracao
End Sub

Private Sub cmdCancelar_Click()
        controlaBotoes Me, operacao.consulta
        fraCadastroUsuarios.Enabled = False
End Sub

Private Sub cmdExcluir_Click()

        Dim strSQL                      As String
        
        If MsgBox("Confirma exlusão?", vbQuestion + vbYesNo + vbDefaultButton2, "Caçambas") = vbNo Then
                Exit Sub
        End If
        
        strSQL = Empty
        strSQL = strSQL & "DELETE FROM Usuarios WHERE idUsuarios = " & vIdUsuario
        
        dbCacamba.Execute strSQL
        
        controlaBotoes Me, operacao.consulta
        fraCadastroUsuarios.Enabled = False
        carregaUsuario
End Sub

Private Sub cmdGravar_Click()
        Dim strSQL                     As String
        
        If Not validaGravacao Then Exit Sub
        
        strSQL = Empty
        
        If vOperacao = Inclusao Then
        
                strSQL = strSQL & " INSERT INTO Usuarios " & vbCrLf
                strSQL = strSQL & " (nome " & vbCrLf
                strSQL = strSQL & " ,login " & vbCrLf
                strSQL = strSQL & " ,senha " & vbCrLf
                strSQL = strSQL & ") VALUES ( "
                strSQL = strSQL & "'" & txtNome.Text & "'" & vbCrLf
                strSQL = strSQL & ",'" & txtLogin.Text & "'" & vbCrLf
                strSQL = strSQL & ",'" & txtSenha.Text & "'" & vbCrLf
                strSQL = strSQL & ")" & vbCrLf
        
        Else
                        
                strSQL = strSQL & " UPDATE Usuarios " & vbCrLf
                strSQL = strSQL & " SET " & vbCrLf
                strSQL = strSQL & " nome = '" & txtNome.Text & "'" & vbCrLf
                strSQL = strSQL & " ,login = '" & txtLogin.Text & "'" & vbCrLf
                strSQL = strSQL & " ,senha = '" & txtSenha.Text & "'" & vbCrLf
                strSQL = strSQL & " WHERE idUsuarios = " & vIdUsuario
        
        End If
        
        dbCacamba.Execute strSQL


        controlaBotoes Me, operacao.consulta
        vOperacao = consulta
        fraCadastroUsuarios.Enabled = False
End Sub

Private Sub cmdIncluir_Click()
        controlaBotoes Me, operacao.Inclusao
        fraCadastroUsuarios.Enabled = True
        limpaCampos
        vIdUsuario = 0
        vOperacao = Inclusao
End Sub

Private Sub cmdLocalizar_Click()
        Dim strSQL                      As String
        Dim vColunas                    As String
        
        strSQL = Empty
        strSQL = strSQL & " SELECT * FROM Usuarios"
        
        vColunas = "Código,Nome, Login, Senha"
        
        frmPesquisa.carregaGridPesquisa 3, strSQL, vColunas, 1
        frmPesquisa.Show vbModal
        
        If vgRetornoConsulta <> 0 Then
                vIdUsuario = vgRetornoConsulta
                carregaUsuario vIdUsuario
        End If
        
        
End Sub

Private Sub cmdSair_Click()
        Unload Me
End Sub

Private Sub Form_Load()
        controlaBotoes Me, operacao.consulta
        fraCadastroUsuarios.Enabled = False
        carregaUsuario
End Sub

Private Sub carregaUsuario(Optional pIdUsuario As Integer)

        Dim strSQL                      As String
        Dim rsUsuario                   As ADODB.Recordset
        
        strSQL = Empty
        If pIdUsuario = 0 Then
                strSQL = strSQL & " SELECT * FROM Usuarios ORDER BY idUsuarios DESC LIMIT 1"
        Else
                strSQL = strSQL & " SELECT * FROM Usuarios " & vbCrLf
                strSQL = strSQL & " WHERE idUsuarios = " & pIdUsuario
        End If
        
        Set rsUsuario = New ADODB.Recordset
        
        With rsUsuario
                .Open strSQL, dbCacamba, adOpenStatic, adLockReadOnly
                If Not .EOF Then
                        vIdUsuario = !idUsuarios
                        txtNome.Text = !nome
                        txtLogin.Text = !login
                        txtSenha.Text = !Senha
                End If
        End With
End Sub

Private Function validaGravacao() As Boolean

        validaGravacao = False
        
        If Trim(txtNome.Text) = "" Then
                MsgBox "Informe o nome do usuário!", vbInformation, "Caçambas"
                Exit Function
        End If
        
        If Trim(txtLogin.Text) = "" Then
                MsgBox "Infomre o login do usuário!", vbInformation, "Caçambas"
                Exit Function
        End If
        
        If Trim(txtSenha.Text) = "" Then
                MsgBox "Infomr a senha do usuário!", vbInformation, "Caçambas"
                Exit Function
        End If
        
        validaGravacao = True
        
End Function

Private Sub limpaCampos()
        
        txtNome.Text = ""
        txtLogin.Text = ""
        txtSenha.Text = ""
End Sub

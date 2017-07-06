VERSION 5.00
Begin VB.Form frmBancoCadastro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Bancos"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8250
   Icon            =   "frmBancoCadastro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmBancoCadastro.frx":000C
   ScaleHeight     =   2730
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraCadastroBanco 
      Height          =   1815
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   8175
      Begin VB.TextBox txtNumeroBanco 
         Height          =   375
         Left            =   120
         MaxLength       =   3
         TabIndex        =   10
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtNomeBanco 
         Height          =   375
         Left            =   120
         MaxLength       =   45
         TabIndex        =   8
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Número do Banco"
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
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label lblNomeCliente 
         Caption         =   "Nome do Banco"
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
         TabIndex        =   9
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "Localizar"
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   735
      Left            =   1185
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   735
      Left            =   3540
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Height          =   735
      Left            =   4725
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   735
      Left            =   5895
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   735
      Left            =   7080
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   735
      Left            =   2355
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
End
Attribute VB_Name = "frmBancoCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vIdBanco                As Integer
Private vOperacao               As operacao

Private Sub cmdAlterar_Click()
        controlaBotoes Me, operacao.Inclusao
        fraCadastroBanco.Enabled = True
        vOperacao = Alteracao
End Sub

Private Sub cmdCancelar_Click()
        controlaBotoes Me, operacao.consulta
        fraCadastroBanco.Enabled = False
        carregaBancos
        vOperacao = consulta
End Sub

Private Sub cmdExcluir_Click()
        Dim strSQL                      As String
        
        If MsgBox("Confirma exlusão?", vbQuestion + vbYesNo + vbDefaultButton2, "Caçambas") = vbNo Then
                Exit Sub
        End If
        
        strSQL = Empty
        strSQL = strSQL & "DELETE FROM bancos WHERE idBancos = " & vIdBanco
        
        dbCacamba.Execute strSQL
        
        controlaBotoes Me, operacao.consulta
        fraCadastroBanco.Enabled = False
        carregaBancos
        vOperacao = consulta

End Sub

Private Sub cmdGravar_Click()
        Dim strSQL                     As String
        
        If Not validaGravacao Then Exit Sub
        
        strSQL = Empty
        
        If vOperacao = Inclusao Then
        
                strSQL = strSQL & " INSERT INTO bancos " & vbCrLf
                strSQL = strSQL & " (nomeBanco " & vbCrLf
                strSQL = strSQL & " ,numeroBanco " & vbCrLf
                strSQL = strSQL & ") VALUES ( "
                strSQL = strSQL & "'" & txtNomeBanco.Text & "'" & vbCrLf
                strSQL = strSQL & ",'" & txtNumeroBanco.Text & "'" & vbCrLf
                strSQL = strSQL & ")" & vbCrLf
        
        Else
                        
                strSQL = strSQL & " UPDATE bancos " & vbCrLf
                strSQL = strSQL & " SET " & vbCrLf
                strSQL = strSQL & " nomeBanco = '" & txtNomeBanco.Text & "'" & vbCrLf
                strSQL = strSQL & " ,numeroBanco =  '" & txtNumeroBanco.Text & "'" & vbCrLf
                strSQL = strSQL & " WHERE idBancos = " & vIdBanco
        
        End If
        
        dbCacamba.Execute strSQL
        
        controlaBotoes Me, operacao.consulta
        vOperacao = consulta
        fraCadastroBanco.Enabled = False
End Sub

Private Sub cmdIncluir_Click()
        controlaBotoes Me, operacao.Inclusao
        fraCadastroBanco.Enabled = True
        limpaCampos
        vIdBanco = 0
        vOperacao = Inclusao
        txtNomeBanco.SetFocus
End Sub

Private Sub cmdLocalizar_Click()
        Dim strSQL                      As String
        Dim vColunas                    As String
        
        strSQL = Empty
        strSQL = strSQL & " SELECT " & vbCrLf
        strSQL = strSQL & " idBancos " & vbCrLf
        strSQL = strSQL & " ,nomeBanco" & vbCrLf
        strSQL = strSQL & " ,numeroBanco" & vbCrLf
        strSQL = strSQL & " FROM Bancos"
        
        vColunas = "Código,Nome,Nimero"
        
        frmPesquisa.carregaGridPesquisa 3, strSQL, vColunas, 1
        frmPesquisa.Show vbModal
        
        If vgRetornoConsulta <> 0 Then
                vIdBanco = vgRetornoConsulta
                carregaBancos vIdBanco
        End If
End Sub

Private Sub cmdSair_Click()
        Unload Me
End Sub

Private Sub carregaBancos(Optional pIdCacamba As Integer)

        Dim strSQL                      As String
        Dim rsCacamba                   As ADODB.Recordset
        
        strSQL = Empty
        If pIdCacamba = 0 Then
                strSQL = strSQL & " SELECT  * "
                strSQL = strSQL & " FROM Bancos ORDER BY idBancos DESC "
                strSQL = strSQL & " LIMIT 1"
        Else
                strSQL = strSQL & " SELECT * FROM Bancos " & vbCrLf
                strSQL = strSQL & " WHERE idBancos = " & pIdCacamba
        End If
        
        Set rsCacamba = New ADODB.Recordset
        
        With rsCacamba
                .Open strSQL, dbCacamba, adOpenStatic, adLockReadOnly
                If Not .EOF Then
                        vIdBanco = !idBancos
                        txtNomeBanco.Text = !nomeBanco
                        txtNumeroBanco.Text = !numeroBanco
                Else
                    limpaCampos
                End If
        End With
End Sub

Private Function validaGravacao() As Boolean

        validaGravacao = False
        
        If Trim(txtNomeBanco.Text) = "" Then
                MsgBox "Informe o nome do banco!", vbInformation, "Caçambas"
                Exit Function
        End If
        
        If Trim(txtNumeroBanco.Text) = "" Then
                MsgBox "Informe o número do banco!", vbInformation, "Caçambas"
                Exit Function
        End If
        
        validaGravacao = True
        
End Function

Private Sub limpaCampos()
        
        txtNomeBanco.Text = ""
        txtNumeroBanco.Text = ""
        
End Sub

Private Sub Form_Load()
    controlaBotoes Me, operacao.consulta
    fraCadastroBanco.Enabled = False
    carregaBancos
End Sub

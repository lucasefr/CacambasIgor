VERSION 5.00
Begin VB.Form frmCacambaCadasto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Caçambas"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8250
   Icon            =   "frmCacambaCadasto.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   735
      Left            =   2355
      TabIndex        =   4
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   735
      Left            =   7080
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      Height          =   735
      Left            =   5895
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdAlterar 
      Caption         =   "Alterar"
      Height          =   735
      Left            =   4725
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   735
      Left            =   3540
      TabIndex        =   1
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "Incluir"
      Height          =   735
      Left            =   1185
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CommandButton cmdLocalizar 
      Caption         =   "Localizar"
      Height          =   735
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Frame fraCadastroCacambas 
      Height          =   1215
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   8175
      Begin VB.TextBox txtNumeroCacamba 
         Height          =   375
         Left            =   120
         MaxLength       =   45
         TabIndex        =   0
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label lblNomeCliente 
         Caption         =   "Número da Caçamba"
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
End
Attribute VB_Name = "frmCacambaCadasto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private vIdCacamba              As Integer
Private vOperacao               As operacao

Private Sub cmdSair_Click()
        Unload Me
End Sub

Private Sub Form_Load()
        controlaBotoes Me, operacao.consulta
        fraCadastroCacambas.Enabled = False
        carregaCacamba
End Sub

Private Sub txtNumeroCacamba_KeyPress(KeyAscii As Integer)
        
        KeyAscii = SoNumeros(KeyAscii)
        
        If KeyAscii = 0 Then
                Exit Sub
        End If
End Sub


Private Sub carregaCacamba(Optional pIdCacamba As Integer)

        Dim strSQL                      As String
        Dim rsCacamba                   As ADODB.Recordset
        
        strSQL = Empty
        If pIdCacamba = 0 Then
                strSQL = strSQL & " SELECT  * "
                strSQL = strSQL & " FROM Cacambas ORDER BY idCacambas DESC "
                strSQL = strSQL & " LIMIT 1"
        Else
                strSQL = strSQL & " SELECT * FROM Cacambas " & vbCrLf
                strSQL = strSQL & " WHERE idCacambas = " & pIdCacamba
        End If
        
        Set rsCacamba = New ADODB.Recordset
        
        With rsCacamba
                .Open strSQL, dbCacamba, adOpenStatic, adLockReadOnly
                If Not .EOF Then
                        vIdCacamba = !idCacambas
                        txtNumeroCacamba.Text = !Numero
                End If
        End With
End Sub

Private Function validaGravacao() As Boolean

        validaGravacao = False
        
        If Trim(txtNumeroCacamba.Text) = "" Then
                MsgBox "Informe o número da caçamba!", vbInformation, "Caçambas"
                Exit Function
        End If
        
        If recuperaDescricao("Cacambas", "Numero", "Numero", txtNumeroCacamba, False) <> "" Then
                MsgBox "Número de caçamba já cadastrado!", vbInformation, "Caçambas"
                Exit Function
        End If
        
        validaGravacao = True
        
End Function

Private Sub limpaCampos()
        
        txtNumeroCacamba.Text = ""
        
End Sub

Private Sub cmdLocalizar_Click()
        Dim strSQL                      As String
        Dim vColunas                    As String
        
        strSQL = Empty
        strSQL = strSQL & " SELECT " & vbCrLf
        strSQL = strSQL & " idCacambas " & vbCrLf
        strSQL = strSQL & " ,numero" & vbCrLf
        strSQL = strSQL & " FROM Cacambas"
        
        vColunas = "Código,Numero"
        
        frmPesquisa.carregaGridPesquisa 2, strSQL, vColunas, 1
        frmPesquisa.Show vbModal
        
        If vgRetornoConsulta <> 0 Then
                vIdCacamba = vgRetornoConsulta
                carregaCacamba vIdCacamba
        End If
        
        
End Sub

Private Sub cmdIncluir_Click()
        controlaBotoes Me, operacao.Inclusao
        fraCadastroCacambas.Enabled = True
        limpaCampos
        vIdCacamba = 0
        vOperacao = Inclusao
        txtNumeroCacamba.SetFocus
End Sub


Private Sub cmdGravar_Click()
        Dim strSQL                     As String
        
        If Not validaGravacao Then Exit Sub
        
        strSQL = Empty
        
        If vOperacao = Inclusao Then
        
                strSQL = strSQL & " INSERT INTO Cacambas " & vbCrLf
                strSQL = strSQL & " (numero " & vbCrLf
                strSQL = strSQL & ") VALUES ( "
                strSQL = strSQL & txtNumeroCacamba.Text & vbCrLf
                strSQL = strSQL & ")" & vbCrLf
        
        Else
                        
                strSQL = strSQL & " UPDATE Cacambas " & vbCrLf
                strSQL = strSQL & " SET " & vbCrLf
                strSQL = strSQL & " numero = " & txtNumeroCacamba.Text & vbCrLf
                strSQL = strSQL & " WHERE idCacambas = " & vIdCacamba
        
        End If
        
        dbCacamba.Execute strSQL

        controlaBotoes Me, operacao.consulta
        vOperacao = consulta
        fraCadastroCacambas.Enabled = False
End Sub


Private Sub cmdExcluir_Click()
        Dim strSQL                      As String
        
        If MsgBox("Confirma exlusão?", vbQuestion + vbYesNo + vbDefaultButton2, "Caçambas") = vbNo Then
                Exit Sub
        End If
        
        strSQL = Empty
        strSQL = strSQL & "DELETE FROM Cacambas WHERE idCacambas = " & vIdCacamba
        
        dbCacamba.Execute strSQL
        
        controlaBotoes Me, operacao.consulta
        fraCadastroCacambas.Enabled = False
        carregaCacamba
        vOperacao = consulta
End Sub

Private Sub cmdCancelar_Click()
        controlaBotoes Me, operacao.consulta
        fraCadastroCacambas.Enabled = False
        vOperacao = consulta
End Sub


Private Sub cmdAlterar_Click()
        controlaBotoes Me, operacao.Inclusao
        fraCadastroCacambas.Enabled = True
        vOperacao = Alteracao
End Sub


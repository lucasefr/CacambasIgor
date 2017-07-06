VERSION 5.00
Begin VB.MDIForm frmCacambas 
   BackColor       =   &H8000000C&
   Caption         =   "Caçambas"
   ClientHeight    =   5910
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10155
   Icon            =   "frmCacambas.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuCadastros 
      Caption         =   "Cadastros"
      Begin VB.Menu mnuUsuarios 
         Caption         =   "Usuários"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuClientes 
         Caption         =   "Clientes"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCacambas 
         Caption         =   "Caçambas"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuCaixa 
         Caption         =   "Caixas"
      End
      Begin VB.Menu mnuBanco 
         Caption         =   "Bancos"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu mnuDocumentos 
      Caption         =   "Documentos"
      Begin VB.Menu mnuPedidoLocacao 
         Caption         =   "Pedido de Locação"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnuGerenciamento 
      Caption         =   "Gerenciamento"
      Begin VB.Menu mnuGerenciamentoCacambas 
         Caption         =   "Caçambas"
      End
      Begin VB.Menu mnuPedidoLocacaoGerenciamento 
         Caption         =   "Pedidos de Locação"
      End
   End
End
Attribute VB_Name = "frmCacambas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuBanco_Click()
        frmBancoCadastro.Show vbModal
End Sub

Private Sub mnuCacambas_Click()
        frmCacambaCadasto.Show vbModal
End Sub

Private Sub mnuCaixa_Click()
        frmCaixa.Show vbModal
End Sub

Private Sub mnuClientes_Click()
        frmClienteCadastro.Show vbModal
End Sub

Private Sub mnuGerenciamentoCacambas_Click()
        frmCacambaGerenciamento.Show vbModal
End Sub

Private Sub mnuPedidoLocacao_Click()
        frmPedidoLocacao.Show vbModal
End Sub

Private Sub mnuPedidoLocacaoGerenciamento_Click()
        frmPedidoLocacaoGerenciamento.Show vbModal
End Sub

Private Sub mnuUsuarios_Click()
        frmUsuarios.Show vbModal
End Sub

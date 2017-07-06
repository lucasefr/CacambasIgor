VERSION 5.00
Begin VB.MDIForm frmCacambas 
   BackColor       =   &H8000000C&
   Caption         =   "Ca�ambas"
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
         Caption         =   "Usu�rios"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuClientes 
         Caption         =   "Clientes"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCacambas 
         Caption         =   "Ca�ambas"
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
         Caption         =   "Pedido de Loca��o"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnuGerenciamento 
      Caption         =   "Gerenciamento"
      Begin VB.Menu mnuGerenciamentoCacambas 
         Caption         =   "Ca�ambas"
      End
      Begin VB.Menu mnuPedidoLocacaoGerenciamento 
         Caption         =   "Pedidos de Loca��o"
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

VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmClientePesquisa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pesquisa Cliente"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13470
   Icon            =   "frmPesquisaCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   13470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Height          =   1815
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   13215
      Begin VB.CommandButton cmdPesquisa 
         Caption         =   "Pesquisa"
         Height          =   495
         Left            =   12000
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtCNPJCPF 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtNomeCliente 
         Height          =   375
         Left            =   120
         MaxLength       =   100
         TabIndex        =   1
         Top             =   480
         Width           =   5175
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
         TabIndex        =   8
         Top             =   960
         Width           =   1695
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
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdSeleciona 
      Caption         =   "Seleciona"
      Height          =   735
      Left            =   11040
      TabIndex        =   4
      Top             =   7920
      Width           =   1095
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   735
      Left            =   12225
      TabIndex        =   0
      Top             =   7920
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid flxPesquisa 
      Height          =   5775
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   10186
      _Version        =   393216
      BackColorBkg    =   12648447
      GridColor       =   16777215
   End
End
Attribute VB_Name = "frmClientePesquisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum colGridCliente
        idCliente = 0
        nome = 1
        CPJCNPJ = 2
        Email = 3
        TelefoneFixo = 4
        TelefoneComercial = 5
        TelefoneCelular = 6
End Enum

Private Sub cmdPesquisa_Click()
        
        Dim strSQL                              As String
        Dim strFiltro                           As String
        Dim rsCliente                           As ADODB.Recordset
        
        geraGridPesquisa
        
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
        
        strFiltro = Empty
        
        If Trim(txtCNPJCPF.Text) <> "" Then
                strFiltro = " WHERE CnpjCPF = '" & txtCNPJCPF.Text & "'" & vbCrLf
        End If
        
        If Trim(txtNomeCliente.Text) <> "" Then
                If strFiltro = Empty Then
                        strFiltro = " WHERE Nome LIKE '%" & txtNomeCliente.Text & "%'"
                Else
                        strFiltro = strFiltro & " AND Nome LIKE '%" & txtNomeCliente.Text & "%'"
                End If
        End If
        
        strSQL = strSQL & strFiltro
        
        Set rsCliente = New ADODB.Recordset
        
        With flxPesquisa
                rsCliente.Open strSQL, dbCacamba, adOpenStatic, adLockReadOnly
                While Not rsCliente.EOF
                        .AddItem ""
                        .TextMatrix(.Rows - 1, colGridCliente.idCliente) = NVL(rsCliente!idClientes, "")
                        .TextMatrix(.Rows - 1, colGridCliente.nome) = NVL(rsCliente!nome, "")
                        .TextMatrix(.Rows - 1, colGridCliente.Email) = NVL(rsCliente!Email, "")
                        .TextMatrix(.Rows - 1, colGridCliente.CPJCNPJ) = NVL(rsCliente!CNPJCPF, "")
                        .TextMatrix(.Rows - 1, colGridCliente.Email) = NVL(rsCliente!Email, "")
                        .TextMatrix(.Rows - 1, colGridCliente.TelefoneFixo) = NVL(rsCliente!TelefoneFixo, "")
                        .TextMatrix(.Rows - 1, colGridCliente.TelefoneComercial) = NVL(rsCliente!TelefoneComercial, "")
                        .TextMatrix(.Rows - 1, colGridCliente.TelefoneCelular) = NVL(rsCliente!TelefoneCelular, "")
                        rsCliente.MoveNext
                Wend
        End With
        
        
End Sub


Private Sub geraGridPesquisa()
        
        With flxPesquisa
                .Rows = 1
                .Cols = 7
                .FixedCols = 0
                .TextMatrix(0, colGridCliente.idCliente) = "Codigo"
                .TextMatrix(0, colGridCliente.nome) = "Cliente"
                .TextMatrix(0, colGridCliente.CPJCNPJ) = "CPF/CNPJ"
                .TextMatrix(0, colGridCliente.Email) = "Email"
                .TextMatrix(0, colGridCliente.TelefoneFixo) = "Telefone Fixo"
                .TextMatrix(0, colGridCliente.TelefoneComercial) = "Telefone Comercial"
                .TextMatrix(0, colGridCliente.TelefoneCelular) = "Telefone Celular"
                
                .ColWidth(colGridCliente.idCliente) = 1200
                .ColWidth(colGridCliente.nome) = 3000
                .ColWidth(colGridCliente.CPJCNPJ) = 1200
                .ColWidth(colGridCliente.Email) = 3000
                .ColWidth(colGridCliente.TelefoneFixo) = 1500
                .ColWidth(colGridCliente.TelefoneComercial) = 1500
                .ColWidth(colGridCliente.TelefoneCelular) = 1500
                
                .SelectionMode = flexSelectionByRow
                .GridLines = flexGridInset
        End With
        
End Sub


Private Sub cmdSair_Click()
        Unload Me
End Sub

Private Sub cmdSeleciona_Click()
        
        If flxPesquisa.Rows = 1 Then
                MsgBox "Selecione o cliente!", vbInformation, "Caçambas"
                Exit Sub
        End If
        
        vgRetornoConsulta = flxPesquisa.TextMatrix(flxPesquisa.Row, colGridCliente.idCliente)
        
        Unload Me
        
        
End Sub

Private Sub Form_Load()
        geraGridPesquisa
End Sub


VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5715
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox txtLogin 
      Height          =   375
      Left            =   120
      MaxLength       =   12
      TabIndex        =   0
      Top             =   4560
      Width           =   2655
   End
   Begin VB.TextBox txtSenha 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   4560
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   4
      Top             =   60
      Width           =   7080
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   120
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   675
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   8
         Top             =   3060
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   7
         Top             =   3270
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         Caption         =   "Warning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   3660
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5970
         TabIndex        =   9
         Top             =   2700
         Width           =   885
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2160
         TabIndex        =   11
         Top             =   1140
         Width           =   2430
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "LicenseTo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "CompanyProduct"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2115
         TabIndex        =   10
         Top             =   705
         Visible         =   0   'False
         Width           =   3000
      End
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
      TabIndex        =   13
      Top             =   4320
      Width           =   1455
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
      Left            =   4560
      TabIndex        =   12
      Top             =   4320
      Width           =   1455
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdOK_Click()
        Dim strSQL                              As String
        Dim rsUsuario                           As ADODB.Recordset
        
        If Trim(txtLogin.Text) = "" Then
                MsgBox "Informe o login!", vbInformation, "Caçambas"
                txtLogin.SetFocus
                Exit Sub
        End If
        
        If Trim(txtSenha.Text) = "" Then
                MsgBox "Informe a senha!", vbInformation, "Caçamvas"
                txtSenha.SetFocus
                Exit Sub
        End If
        
        strSQL = Empty
        strSQL = strSQL & " SELECT * " & vbCrLf
        strSQL = strSQL & " FROM Usuarios " & vbCrLf
        strSQL = strSQL & " WHERE " & vbCrLf
        strSQL = strSQL & " Login ='" & txtLogin.Text & "'" & vbCrLf
        strSQL = strSQL & " AND Senha = '" & txtSenha.Text & "'"
        
        Set rsUsuario = New ADODB.Recordset
        
        With rsUsuario
                .Open strSQL, dbCacamba, adOpenStatic, adLockReadOnly
                If Not .EOF Then
                        vgIdUsuarioLogado = !idUsuarios
                        Unload Me
                Else
                        MsgBox "Login ou senha incorretos!", vbInformation, "Caçambas"
                        Exit Sub
                End If
        End With
        
End Sub

Private Sub cmdSair_Click()

        End

End Sub


Private Sub Form_Load()
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = "Fera Caçambas"
    lblVersion.Caption = "1.0.0"
    conectaBanco
End Sub



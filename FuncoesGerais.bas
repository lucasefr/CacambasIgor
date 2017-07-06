Attribute VB_Name = "FuncoesGerais"
Option Explicit

Public dbCacamba                        As ADODB.Connection
Public vgRetornoConsulta                As Integer
Public vgIdUsuarioLogado                As Integer

Public Enum tipoPessoa
        Fisica = 0
        Juridica = 1
End Enum

Public Enum operacao
        Inclusao = 0
        Alteracao = 1
        consulta = 2
End Enum

Public Enum e_TipoPedidoLocacao
    aluguel = 0
    troca = 1
End Enum


Public Sub CarregaComboUf(pCombo As ComboBox)
        
        pCombo.AddItem "MG"
        pCombo.AddItem "AC"
        pCombo.AddItem "AL"
        pCombo.AddItem "AP"
        pCombo.AddItem "AM"
        pCombo.AddItem "BA"
        pCombo.AddItem "CE"
        pCombo.AddItem "DF"
        pCombo.AddItem "ES"
        pCombo.AddItem "GO"
        pCombo.AddItem "MA"
        pCombo.AddItem "MT"
        pCombo.AddItem "MS"
        pCombo.AddItem "PA"
        pCombo.AddItem "PB"
        pCombo.AddItem "PR"
        pCombo.AddItem "PE"
        pCombo.AddItem "PI"
        pCombo.AddItem "RJ"
        pCombo.AddItem "RN"
        pCombo.AddItem "RS"
        pCombo.AddItem "RO"
        pCombo.AddItem "RR"
        pCombo.AddItem "SC"
        pCombo.AddItem "SP"
        pCombo.AddItem "SE"
        pCombo.AddItem "TO"
End Sub

Public Sub carregaComboFormaPagamento(pCombo As ComboBox)

        pCombo.AddItem "Dinheiro"
        pCombo.AddItem "Cheque"
        pCombo.AddItem "Transferência"
        pCombo.AddItem "Cartão"
        pCombo.AddItem "Boleto"
        
End Sub
Public Sub carregaComboSituacao(pCombo As ComboBox)

        pCombo.AddItem "A Receber"
        pCombo.AddItem "Pago na Entrega"
        pCombo.AddItem "Pago na Retirada"
        pCombo.AddItem "Devedor"
        pCombo.AddItem "Cortesia"
        
End Sub


Public Sub controlaBotoes(pForm As Form, pOperacao As Integer)

        Select Case pOperacao
                Case operacao.Inclusao
                        pForm.cmdIncluir.Enabled = False
                        pForm.cmdLocalizar.Enabled = False
                        pForm.cmdCancelar.Enabled = True
                        pForm.cmdExcluir.Enabled = False
                        pForm.cmdGravar.Enabled = True
                        pForm.cmdAlterar.Enabled = False

                Case operacao.Alteracao
                        pForm.cmdIncluir.Enabled = False
                        pForm.cmdLocalizar.Enabled = False
                        pForm.cmdCancelar.Enabled = True
                        pForm.cmdExcluir.Enabled = False
                        pForm.cmdGravar.Enabled = True
                        pForm.cmdAlterar.Enabled = False

                Case operacao.consulta
                        pForm.cmdIncluir.Enabled = True
                        pForm.cmdLocalizar.Enabled = True
                        pForm.cmdCancelar.Enabled = False
                        pForm.cmdExcluir.Enabled = True
                        pForm.cmdGravar.Enabled = False
                        pForm.cmdAlterar.Enabled = True
                        
        End Select

End Sub

Public Sub conectaBanco()
                
        Set dbCacamba = New ADODB.Connection
        dbCacamba.Open "Provider=MSDASQL;Driver=MySQL ODBC 5.1 Driver;server=localhost;uid=root;pwd=hamatullametal;database=cacambas;port=3306"
        
End Sub

Public Function NVL(pValor As Variant, pRetorno As String) As String
        
        If IsNull(pValor) Or pValor = "" Then
                NVL = pRetorno
        Else
                NVL = pValor
        End If
        
End Function

Public Function SoNumeros(ByVal KeyAscii As Integer) As Integer


        If InStr("1234567890", Chr(KeyAscii)) = 0 Then
                SoNumeros = 0
        Else
                SoNumeros = KeyAscii
        End If

        Select Case KeyAscii
                Case 8
                        SoNumeros = KeyAscii
                Case 13
                        SoNumeros = KeyAscii
                Case 32
                        SoNumeros = KeyAscii
        End Select
End Function

Public Function recuperaDescricao(pNomeTabela As String, pCampoRecuperado As String, pCampoChave As String, pValorChave As String, pValorChaveString As Boolean) As String
        
        Dim strSQL                          As String
        Dim rsPesquisa                      As ADODB.Recordset
        
        strSQL = Empty
        strSQL = strSQL & " SELECT " & vbCrLf
        strSQL = strSQL & pCampoRecuperado & " as descricao " & vbCrLf
        strSQL = strSQL & " FROM " & pNomeTabela & vbCrLf
        If pValorChaveString Then
                strSQL = strSQL & " WHERE " & pCampoChave & " = '" & pValorChave & "'"
        Else
                strSQL = strSQL & " WHERE " & pCampoChave & " = " & pValorChave
        End If
        
        Set rsPesquisa = New ADODB.Recordset
        
        With rsPesquisa
                .Open strSQL, dbCacamba, adOpenStatic, adLockReadOnly
                If Not .EOF Then
                        recuperaDescricao = !descricao
                End If
        End With
        
End Function

Public Function ValidarCpf(ByVal pCPF As String, Optional ByVal pPermiteZeros As Boolean = False) As Boolean
    Dim X                                                   As Integer
    Dim mult1                                               As Integer
    Dim mult2                                               As Integer
    Dim dig1                                                As Integer
    Dim dig2                                                As Integer
    Dim strAux                                              As String

    ' Monta string contendo 11 caracteres numéricos
    strAux = Right$(String$(11, "0") & GetOnlyNumbers(pCPF), 11)

    mult1 = 10
    mult2 = 11
    For X = 1 To 9
        dig1 = dig1 + (CInt(Mid$(strAux, X, 1)) * mult1)
        mult1 = mult1 - 1
    Next X
    For X = 1 To 10
        dig2 = dig2 + (CInt(Mid$(strAux, X, 1)) * mult2)
        mult2 = mult2 - 1
    Next X
    dig1 = (dig1 * 10) Mod 11
    dig2 = (dig2 * 10) Mod 11
    If dig1 = 10 Then dig1 = 0
    If dig2 = 10 Then dig2 = 0
    If CInt(Mid$(strAux, 10, 1)) <> dig1 Then
        Exit Function
    End If
    If CInt(Mid$(strAux, 11, 1)) <> dig2 Or (Val(strAux) = 0 And Not pPermiteZeros) Then
        Exit Function
    End If

    ValidarCpf = True
End Function

Public Function EntradaCNPJinvalida(vcampo As Variant) As Boolean
    EntradaCNPJinvalida = False
    If Not ValidarCnpj(vcampo.Text, True) Then
        EntradaCNPJinvalida = True
        MsgBox "Numero CNPJ Invalido.", vbExclamation, "Atenção"
        If vcampo.Enabled And vcampo.Visible Then vcampo.SetFocus
        Exit Function
    End If
End Function

Public Function EntradaCPFinvalida(vcampo As Variant, Optional bShowMsg As Boolean = True) As Boolean
    EntradaCPFinvalida = False
    If Not ValidarCpf(vcampo.Text, True) Then
        EntradaCPFinvalida = True
        MsgBox "Numero CPF Invalido.", vbExclamation, "Atenção"
        If vcampo.Enabled And vcampo.Visible Then vcampo.SetFocus
        Exit Function
    End If
End Function

Public Function ValidarCnpj(ByVal pCnpj As String, Optional ByVal pPermiteZeros As Boolean = False) As Boolean
    Dim mult1                                               As String
    Dim mult2                                               As String
    Dim dig1                                                As Integer
    Dim dig2                                                As Integer
    Dim X                                                   As Integer
    Dim strAux                                              As String

    ' Monta string contendo 14 caracteres numéricos
    strAux = Right$(String$(14, "0") & GetOnlyNumbers(pCnpj), 14)

    mult1 = "543298765432"
    mult2 = "6543298765432"
    For X = 1 To 12
        dig1 = dig1 + (CInt(Mid$(strAux, X, 1)) * CInt(Mid$(mult1, X, 1)))
    Next X
    For X = 1 To 13
        dig2 = dig2 + (CInt(Mid$(strAux, X, 1)) * CInt(Mid$(mult2, X, 1)))
    Next X

    dig1 = (dig1 * 10) Mod 11
    dig2 = (dig2 * 10) Mod 11
    If dig1 = 10 Then dig1 = 0
    If dig2 = 10 Then dig2 = 0

    If dig1 <> CInt(Mid$(strAux, 13, 1)) Then
        Exit Function
    End If
    If dig2 <> CInt(Mid$(strAux, 14, 1)) Or (Val(strAux) = 0 And Not pPermiteZeros) Then
        Exit Function
    End If

    ValidarCnpj = True
End Function


Public Function GetOnlyNumbers(ByVal pStr As String) As String

        Dim i                           As Integer
        Dim strNum                      As String
        
        For i = 1 To Len(pStr)
                If Mid$(pStr, i, 1) Like "[0-9]" Then
                        strNum = strNum & Mid$(pStr, i, 1)
                End If
        Next i
        
        GetOnlyNumbers = strNum

End Function

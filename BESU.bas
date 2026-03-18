Attribute VB_Name = "Módulo5"
Option Explicit

Function BuscarEletricoUnificado(valorRef As Double, unidadeRef As String, tabela As Range, colValor As Integer, colUnidade As Integer, colRetorno As Integer) As Variant
    ' ==============================================================================
    ' Parâmetros:
    ' valorRef:    O valor medido (ex: 500)
    ' unidadeRef:  A unidade desse valor (ex: "mO", "kO", "O")
    ' tabela:      O intervalo da tabela de dados
    ' colValor:    Coluna dos limites numéricos
    ' colUnidade:  Coluna das unidades
    ' colRetorno:  Coluna do valor a ser retornado (Erro/Resolução)
    ' ==============================================================================

    Dim i As Long
    
    ' Variáveis da Referência (Entrada do usuário)
    Dim valRefAbsoluto As Double
    Dim multRef As Double
    
    ' Variáveis da Tabela (Dados do equipamento)
    Dim valTab As Double
    Dim uniTab As String
    Dim valTabAbsoluto As Double
    Dim multTab As Double
    Dim valorLeitura As Double
    
    ' 1. Define o Multiplicador da Referência (Entrada)
    ' Se unidadeRef for "kO", multRef será 1000
    multRef = ObterMultiplicador(unidadeRef)
    
    ' Converte para valor absoluto (Base) para comparação
    valRefAbsoluto = valorRef * multRef
    
    ' 2. Varre a tabela
    For i = 1 To tabela.Rows.Count
        
        ' Lê os dados da linha
        valTab = tabela.Cells(i, colValor).Value
        uniTab = Trim(tabela.Cells(i, colUnidade).Value)
        
        ' Define o Multiplicador da Linha Atual da Tabela
        multTab = ObterMultiplicador(uniTab)
        
        ' Converte limite da tabela para absoluto
        valTabAbsoluto = valTab * multTab
        
        ' 3. Comparação: O range cobre o valor medido?
        If valTabAbsoluto >= valRefAbsoluto Then
            
            ' Pega o valor cru da tabela (que está na unidade da tabela)
            valorLeitura = tabela.Cells(i, colRetorno).Value
            
            ' 4. CÁLCULO FINAL DE DIMENSÃO (O "Pulo do Gato")
            ' Transformamos o valor lido em absoluto (* multTab)
            ' E dividimos pelo multiplicador da referência (/ multRef)
            ' Assim, o número retornado estará na mesma grandeza da entrada.
            
            BuscarEletricoUnificado = (valorLeitura * multTab) / multRef
            
            Exit Function
        End If
        
    Next i

    BuscarEletricoUnificado = "Fora de Range"

End Function

' ==============================================================================
' Função Auxiliar: Define o fator de multiplicação (10^x)
' Aceita "K" e "k" como 1000. "O" isolado como 1.
' ==============================================================================
Private Function ObterMultiplicador(txtUnidade As String) As Double
    Dim prefixo As String
    txtUnidade = Trim(txtUnidade)
    
    ' Se estiver vazio ou for apenas o símbolo Ohm, é unidade base (x1)
    If Len(txtUnidade) = 0 Or txtUnidade = "O" Then
        ObterMultiplicador = 1
        Exit Function
    End If
    
    ' Pega a primeira letra para checar o prefixo
    prefixo = Left(txtUnidade, 1)
    
    ' Verifica o prefixo (Binário para diferenciar m de M)
    Select Case True
        ' Micro (u ou símbolo µ)
        Case prefixo = "u" Or prefixo = "µ": ObterMultiplicador = 10 ^ -6
        
        ' Mili (m minúsculo) - Cuidado para não confundir com Mega
        Case StrComp(prefixo, "m", vbBinaryCompare) = 0: ObterMultiplicador = 10 ^ -3
        
        ' Kilo (Aceita tanto "k" quanto "K" conforme solicitado)
        Case StrComp(prefixo, "k", vbBinaryCompare) = 0 Or StrComp(prefixo, "K", vbBinaryCompare) = 0: ObterMultiplicador = 10 ^ 3
        
        ' Mega (M maiúsculo)
        Case StrComp(prefixo, "M", vbBinaryCompare) = 0: ObterMultiplicador = 10 ^ 6
        
        ' Giga
        Case StrComp(prefixo, "G", vbBinaryCompare) = 0: ObterMultiplicador = 10 ^ 9
        
        ' Nano
        Case StrComp(prefixo, "n", vbBinaryCompare) = 0: ObterMultiplicador = 10 ^ -9
        
        'Tera
         Case StrComp(prefixo, "T", vbBinaryCompare) = 0: ObterMultiplicador = 10 ^ 12
            
        ' Caso não ache prefixo (ex: "V", "A"), multiplica por 1
        Case Else: ObterMultiplicador = 1
    End Select
End Function

Attribute VB_Name = "conferir"
'VALIDAR_tabela
'
'Validar o tabela seguindo uma série de regras de negócios, que podem aumentar o desativar algum
'A lógica consistem em percorrer cada registro da tbl_tabela e ver se encaixa uma regra na tbl_regras_tabela


'Declaração de Objetos
Private dados_tabela As ListObject
Private dados_regras As ListObject
Private dictColunas_dados As Dictionary
Private dictColunas_regras As Dictionary
Private colErros As New Collection

'Declaração de Variáveis
Private arrDados As Variant
Private arrRegras As Variant
Private valorCampo_regras As String
Private linhatabela As Long
Private linhaRegras As Long
Private notificarErro As Boolean


Sub validar_tabela()


Dim contador As Long
Dim inconsistencia As Boolean 'Critério para executar function que separa registro que está errado
Dim nomeCampo_regras As String
Dim celula As Range
Dim valorCampo_dados As String
notificarErro = False

'Setar Objetos
Set dados_tabela = GarantirTabela("Tabela", "tbl_tabela")
Set dados_regras = GarantirTabela("Lista de Regras", "tbl_regras_tabela")
Set dictColunas_dados = New Dictionary
Set dictColunas_regras = New Dictionary

'Resetar campo Já Agrupado
For Each celula In dados_regras.ListColumns("Já foi Agrupado?").DataBodyRange
    celula.Value = "Não"
Next celula


'Validar e Adicionar Registros
If dados_tabela.DataBodyRange Is Nothing Then
    MsgBox "Sem registro na tabela para validar."
    Exit Sub
Else
    'Inserir Registros
    arrDados = dados_tabela.DataBodyRange.Value
    arrRegras = dados_regras.DataBodyRange.Value
    
    'Inserir Cabeçaho
    For contador = 1 To dados_tabela.HeaderRowRange.Columns.Count
        dictColunas_dados.Add dados_tabela.HeaderRowRange.Cells(1, contador).Value, contador
    Next contador
    For contador = 1 To dados_regras.HeaderRowRange.Columns.Count
        dictColunas_regras.Add dados_regras.HeaderRowRange.Cells(1, contador).Value, contador
    Next contador
    
    
    
    
End If


'Processar Regras de Negócio
For linhatabela = 1 To UBound(arrDados, 1) 'Cada registro
    For linhaRegras = 1 To UBound(arrRegras, 1) 'Cada regra
        If arrRegras(linhaRegras, dictColunas_regras("Ativo")) = "Sim" Then
            
            'Preparação para próxima etapa
            inconsistencia = False
            
            'Definindo o campo de busca da tabela Regras
            nomeCampo_regras = arrRegras(linhaRegras, dictColunas_regras("Campo Alvo"))
            valorCampo_regras = arrRegras(linhaRegras, dictColunas_regras("Valor Alvo"))
            valorCampo_dados = arrDados(linhatabela, dictColunas_dados(nomeCampo_regras))
            
            
            'Validando se verica inconsistência
            If valorCampo_regras = valorCampo_dados Then
                inconsistencia = validar_registro(linhatabela, linhaRegras)
                If inconsistencia = True Then
                    'Colar os registros no colErros
                    Call registrar_erro(linhaRegras, linhatabela)
                End If
            End If
            
        
        End If
    Next linhaRegras
Next linhatabela

'Colar dados do colErros uma nova planilha
Call planilha_erro

'Resetar o campo Já Agrupado
For Each celula In dados_regras.ListColumns("Já foi Agrupado?").DataBodyRange
    celula.Value = "Não"
Next celula

'Limpar Objetos
Set colErros = Nothing
Set dados_tabela = Nothing
Set dados_regras = Nothing
Set dictColunas_dados = Nothing
Set dictColunas_regras = Nothing
notificarErro = False

'Aviso de Finalização
MsgBox "Macro finalizada!"


End Sub


Private Function validar_registro(lintabela As Long, linRegras As Long) As Boolean

On Error GoTo erro_regra

'Declaração Variáveis
Dim lista_valores() As String
Dim conjunto As String
Dim valor_tabela As String
Dim nomeCampo_regras As String
Dim operador As String, agrupado As String
Dim valorEncontrado As Boolean, executarCalculo As Boolean
Dim i As Long

'Regras de Pertencimento
conjunto = Replace(Replace(arrRegras(linRegras, dictColunas_regras("Valor_Condição1")), Chr(10), ""), Chr(13), "")
lista_valores = Split(conjunto, ";")
operador = arrRegras(linRegras, dictColunas_regras("Operador_Condição1"))
If operador <> "Combinação" Then 'Se operador for Combinação, não será utilizado esses campos dessa forma
    nomeCampo_regras = arrRegras(linRegras, dictColunas_regras("Campo_Condição1"))
    valor_tabela = arrDados(lintabela, dictColunas_dados(nomeCampo_regras))
End If
agrupado = arrRegras(linRegras, dictColunas_regras("Tipo Agrupamento"))
valorEncontrado = False
executarCalculo = False


Select Case Trim(operador)
    
    Case "<>", "=" 'Operadores de Igual ou Diferente
        
        For i = LBound(lista_valores) To UBound(lista_valores)
            If lista_valores(i) = valor_tabela Then
                valorEncontrado = True
                Exit For
            End If
        Next i
        
        'Debug.Print "Registro:[" & valor_tabela & "], Encontrado: [" & valorEncontrado & "]"
        'Lançamento Correto
        If Trim(operador) = "<>" Then
            If valorEncontrado = False Then
                validar_registro = True
                
                'Se tem Calculo e Agrupamento
                If agrupado <> "" And arrRegras(linRegras, dictColunas_regras("Já foi Agrupado?")) = "Não" Then
                    dados_regras.DataBodyRange.Cells(linRegras, dictColunas_regras("Já foi Agrupado?")).Value = "Sim"
                    GoTo erro_regra
                End If
                
            End If
        ElseIf Trim(operador) = "=" Then
            If valorEncontrado = True Then
                validar_registro = True
                
                'Se tem Calculo e Agrupamento
                If agrupado <> "" And arrRegras(linRegras, dictColunas_regras("Já foi Agrupado?")) = "Não" Then
                    dados_regras.DataBodyRange.Cells(linRegras, dictColunas_regras("Já foi Agrupado?")).Value = "Sim"
                    GoTo erro_regra
                End If
                
            End If
        End If
        
        
    
    
    Case "Contém", "Não Contém" 'Operadores de Igual ou Diferente
        
        For i = LBound(lista_valores) To UBound(lista_valores)
            If UCase(valor_tabela) Like "*" & UCase(lista_valores(i)) & "*" Then
                valorEncontrado = True
                Exit For
            End If
        Next i
        
        'Debug.Print "Registro:[" & valor_tabela & "], Encontrado: [" & valorEncontrado & "]"
        'Lançamento Correto
        If Trim(operador) = "Não Contém" Then
            If valorEncontrado = False Then
                validar_registro = True
                
                'Se tem Calculo e Agrupamento
                If agrupado <> "" And arrRegras(linRegras, dictColunas_regras("Já foi Agrupado?")) = "Não" Then
                    dados_regras.DataBodyRange.Cells(linRegras, dictColunas_regras("Já foi Agrupado?")).Value = "Sim"
                    GoTo erro_regra
                End If
                
            End If
        ElseIf Trim(operador) = "Contém" Then
            If valorEncontrado = True Then
                validar_registro = True
                
                'Se tem Calculo e Agrupamento
                If agrupado <> "" And arrRegras(linRegras, dictColunas_regras("Já foi Agrupado?")) = "Não" Then
                    dados_regras.DataBodyRange.Cells(linRegras, dictColunas_regras("Já foi Agrupado?")).Value = "Sim"
                    GoTo erro_regra
                End If
                
            End If
        End If

    
    
    Case ">", "<" 'Operadores Maior e Menor (somente com valores númericos)
        
        'Validar campo númerico ou data
        If Not IsDate(valor_tabela) Then
            If Not IsNumeric(valor_tabela) Then
                Debug.Print "Data Credito " & valor_tabela
                Debug.Print "Coluna " & nomeCampo_regras & " não é do tipo númerico. Transformar dados da coluna."
                validar_registro = False
                'MsgBox "Operador > não atribuido ao campo Valor.Corrigir a regra Erro VBA."
                dados_regras.DataBodyRange.Cells(linRegras, dictColunas_regras("Ativo")).Value = "Erro VBA"
                Exit Function
            End If
        End If
    
                
        'Buscar valores
        If Trim(operador) = ">" Then
            For i = LBound(lista_valores) To UBound(lista_valores)
                If IsDate(valor_tabela) Then
                    If CDate(valor_tabela) > CDate(lista_valores(i)) Then
                        valorEncontrado = True
                        Exit For
                    End If
                Else
                    If CDbl(valor_tabela) > CDbl(lista_valores(i)) Then
                        valorEncontrado = True
                        Exit For
                    End If
                End If
            Next i
            
            'Debug.Print "Registro:[" & valor_tabela & "], Encontrado: [" & valorEncontrado & "]"
            'Lançamento Correto
            If valorEncontrado = True Then
                validar_registro = True
                
                'Se tem Calculo e Agrupamento
                If agrupado <> "" And arrRegras(linRegras, dictColunas_regras("Já foi Agrupado?")) = "Não" Then
                    dados_regras.DataBodyRange.Cells(linRegras, dictColunas_regras("Já foi Agrupado?")).Value = "Sim"
                    GoTo erro_regra
                End If
            End If
        
        ElseIf Trim(operador) = "<" Then
            For i = LBound(lista_valores) To UBound(lista_valores)
                If IsDate(valor_tabela) Then
                    If CDate(valor_tabela) < CDate(lista_valores(i)) Then
                        valorEncontrado = True
                        Exit For
                    End If
                Else
                    If CDbl(valor_tabela) < CDbl(lista_valores(i)) Then
                        valorEncontrado = True
                        Exit For
                    End If
                End If
            Next i
        
            'Debug.Print "Registro:[" & valor_tabela & "], Encontrado: [" & valorEncontrado & "]"
            'Lançamento Correto
            If valorEncontrado = True Then
                validar_registro = True
                
                'Se tem Calculo e Agrupamento
                If agrupado <> "" And arrRegras(linRegras, dictColunas_regras("Já foi Agrupado?")) = "Não" Then
                    dados_regras.DataBodyRange.Cells(linRegras, dictColunas_regras("Já foi Agrupado?")).Value = "Sim"
                    GoTo erro_regra
                End If
            End If
        
        End If
        
        
        
    
    Case "Filtro" 'Seleciona primeiro registro encontrado no parâmetro lista_valores e envia para Agrupamento
        For i = LBound(lista_valores) To UBound(lista_valores)
            If lista_valores(i) = valor_tabela Then
                valorEncontrado = True
                Exit For
            End If
        Next i
        
        'Lançamento Correto
        If valorEncontrado = True Then
            'Se tem Calculo e Agrupamento
            If agrupado <> "" And dados_regras.DataBodyRange.Cells(linRegras, dictColunas_regras("Já foi Agrupado?")).Value = "Não" Then
                dados_regras.DataBodyRange.Cells(linRegras, dictColunas_regras("Já foi Agrupado?")).Value = "Sim"
                executarCalculo = True
            End If
        End If
        
        
    
    Case "Encontrar" 'É inconsistência se não encontrar nenhum lançamento no parâmetro lista_valores
        If dados_regras.DataBodyRange.Cells(linRegras, dictColunas_regras("Já foi Agrupado?")).Value = "Não" Then
            For i = LBound(lista_valores) To UBound(lista_valores)
                If lista_valores(i) = valor_tabela Then
                    dados_regras.DataBodyRange.Cells(linRegras, dictColunas_regras("Já foi Agrupado?")).Value = "Sim"
                    Exit Function
                Else
                    ' Regra "Encontrar": se não achou nenhuma ocorrência do valor esperado,
                    ' reaproveita a lógica de "Contar" simulando "quantidade < 1" para acusar erro.
                    ' Dados dessa regra no arrRegras são sobrescritos para  exectuar o agrupamento Contar.
                    dados_regras.DataBodyRange.Cells(linRegras, dictColunas_regras("Já foi Agrupado?")).Value = "Sim"
                    agrupado = "Contar"
                    arrRegras(linRegras, dictColunas_regras("Operador_Condição2")) = "<"
                    arrRegras(linRegras, dictColunas_regras("Valor_Condição2")) = 1
                    valor_tabela = conjunto
                    executarCalculo = True
                End If
            Next i
            
            'Se tem Calculo e Agrupamento
            If agrupado <> "" And dados_regras.DataBodyRange.Cells(linRegras, dictColunas_regras("Já foi Agrupado?")).Value = "Não" Then
                dados_regras.DataBodyRange.Cells(linRegras, dictColunas_regras("Já foi Agrupado?")).Value = "Sim"
                GoTo erro_regra
            End If
        End If
    
    
    
    Case "Combinação" 'É inconsistência encontrar lançamento no parâmetro lista_valores
        Dim qntParametro As Long, contador As Long, a As Long
        qntParametro = UBound(lista_valores) - LBound(lista_valores) + 1
        contador = 0
        
        For i = 1 To dados_tabela.HeaderRowRange.Columns.Count
            valor_tabela = arrDados(lintabela, i)
            For a = LBound(lista_valores) To UBound(lista_valores)
                If lista_valores(a) = valor_tabela Then
                    contador = contador + 1
                    Exit For
                End If
            Next a
        Next i
        
        
        If contador = qntParametro Then
            validar_registro = True
            
            'Se tem Calculo e Agrupamento
            If agrupado <> "" And dados_regras.DataBodyRange.Cells(linRegras, dictColunas_regras("Já foi Agrupado?")).Value = "Não" Then
                dados_regras.DataBodyRange.Cells(linRegras, dictColunas_regras("Já foi Agrupado?")).Value = "Sim"
                GoTo erro_regra
            End If
            Exit Function
        End If
        
    
    
    
    
        
End Select


'Regras de Agrupamento
If executarCalculo = True Then
    
    'Variáveis Exclusivas do Agrupamento
    Dim calculo As Double
    Dim limite As Double
    Dim k As Long
    Dim op_agrupado As String, campoAlvo As String
    Dim erroLocalizado As Boolean
    calculo = 0
    validar_registro = False 'Agrupamentos chamam anotar_erro() sem passar pela inconsistência = true
    op_agrupado = arrRegras(linRegras, dictColunas_regras("Operador_Condição2"))
    limite = arrRegras(linRegras, dictColunas_regras("Valor_Condição2"))
    erroLocalizado = False
    campoAlvo = arrRegras(linRegras, dictColunas_regras("Campo Alvo"))
    
    
    'Calcular Valores
    If agrupado = "Somar" Then 'Exemplo do prestador com valor fixo + comissão
        Dim nomeCampo_Calculo As String
        nomeCampo_Calculo = arrRegras(linRegras, dictColunas_regras("Campo_Condição2"))
        For k = 1 To UBound(arrDados, 1)
            If arrDados(k, dictColunas_dados(nomeCampo_regras)) = valor_tabela Then
                If valorCampo_regras = arrDados(k, dictColunas_dados(campoAlvo)) Then
                    calculo = calculo + arrDados(k, dictColunas_dados(nomeCampo_Calculo))
                End If
            End If
        Next k

    
    ElseIf agrupado = "Contar" Then 'Exemplo de 3 pagamentos e não sabe nome do fornecedor
        For k = 1 To UBound(arrDados, 1)
            If arrDados(k, dictColunas_dados(nomeCampo_regras)) = valor_tabela Then
                If valorCampo_regras = arrDados(k, dictColunas_dados(campoAlvo)) Then
                    calculo = calculo + 1
                End If
            End If
        Next k
    
    ElseIf agrupado = "Média" Then 'Exemplo do Ticket Médio
        Dim nomeCampo_Calc As String, total As Double, qntde As Long
        nomeCampo_Calc = arrRegras(linRegras, dictColunas_regras("Campo_Condição2"))
        total = 0
        qntde = 0
        For k = 1 To UBound(arrDados, 1)
            If arrDados(k, dictColunas_dados(nomeCampo_regras)) = valor_tabela Then
                If valorCampo_regras = arrDados(k, dictColunas_dados(campoAlvo)) Then
                    total = total + arrDados(k, dictColunas_dados(nomeCampo_Calc))
                    qntde = qntde + 1
                End If
            End If
        Next k
        If total > 0 Then
            calculo = total / qntde
        Else
            calculo = 0
        End If

    
    Else
        validar_registro = False
        MsgBox "Operador ainda não programado no vba. Revisar a function!"
        dados_regras.DataBodyRange.Cells(linRegras, dictColunas_regras("Ativo")).Value = "Erro VBA"
        Debug.Print "Função IF do somar, contar."
    End If
               
               
    'Validar Condicional
    Debug.Print calculo
    Select Case Trim(op_agrupado)
            
        Case "<"
            If calculo < limite Then
                erroLocalizado = True
            End If
        
        Case ">"
            If calculo > limite Then
                erroLocalizado = True
            End If
        
        Case "="
            If calculo = limite Then
                erroLocalizado = True
            End If
        
        Case "<>"
            If calculo <> limite Then
                erroLocalizado = True
            End If
                
        Case Else
            MsgBox "Operador ainda não programado no vba. Revisar a function!"
            dados_regras.DataBodyRange.Cells(linRegras, dictColunas_regras("Ativo")).Value = "Erro VBA"
            Debug.Print "Select Case Agrupamento"
         
    End Select
    
    
    'Inserir registros errados (se condicional = true)
    If erroLocalizado = True Then
        If operador = "Encontrar" Then
            Call registrar_erro(linRegras)
        Else
            For k = 1 To UBound(arrDados, 1)
                If arrDados(k, dictColunas_dados(nomeCampo_regras)) = valor_tabela Then 'Campo Condicao_1"
                    If valorCampo_regras = arrDados(k, dictColunas_dados(campoAlvo)) Then 'Campo Alvo
                        'Colar os registros no arrErros()
                        Call registrar_erro(linRegras, k)
                    End If
                End If
            Next k
        End If
    End If
End If

Exit Function


erro_regra:
        'Operador não programado no VBA
        Debug.Print "Nome da regra: " & arrRegras(linRegras, dictColunas_regras("Nome da Regra"))
        Debug.Print "Operador 1: " & Trim(operador) & " | Operador 2: " & agrupado & " | Operador 3: " & op_agrupado
        validar_registro = False
        If notificarErro = False Then 'Avisar somente uma vez quando der erro em alguma regra
            notificarErro = True
            MsgBox "Há no mínimo uma regra que deu erro ao executar. Verificar regras com nome Erro VBA!" & vbCrLf & _
                    "1º Revise se a regra foi escrita corretamente." & vbCrLf & _
                    "2º Revise se a coluna da tabela foi formatada corretamente (tipo data, tipo número, espaçamentos, células vazias, etc.)", vbExclamation
        End If
        dados_regras.DataBodyRange.Cells(linRegras, dictColunas_regras("Ativo")).Value = "Erro VBA"
       
       

End Function


Private Sub registrar_erro(regras As Long, Optional dados As Long)
    
    'Criando o registro
    Dim linErro() As Variant
    Dim v As Long, qntColunas As Long
    
    qntColunas = dados_tabela.HeaderRowRange.Columns.Count
    ReDim linErro(1 To qntColunas + 2)
    
    If dados = 0 Then
        For v = 1 To qntColunas
            linErro(v) = ""
        Next v
    Else
        For v = 1 To qntColunas
            linErro(v) = arrDados(dados, v)
        Next v
    End If
    linErro(qntColunas + 1) = arrRegras(regras, dictColunas_regras("Nome da Regra"))
    linErro(qntColunas + 2) = arrRegras(regras, dictColunas_regras("Mensagem do Erro"))
    
    
    'Salvando em colErros
    colErros.Add linErro

End Sub



Private Sub planilha_erro()

    'Validar Collection
    If colErros.Count = 0 Then
        MsgBox "Todos estão corretamente tratado." & vbCrLf & "Não há erros de tabela nessa tabela."
        Exit Sub
    End If
    
    
    'Criar planilha nova
    Dim planilhaNova As Worksheet
    Set planilhaNova = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count))
    
    
    'Criar o cabeçalho
    Dim totalColunas As Long, c As Long, r As Long
    totalColunas = dados_tabela.HeaderRowRange.Columns.Count + 2
    
    For c = 1 To dados_tabela.HeaderRowRange.Columns.Count
        planilhaNova.Cells(1, c).Value = dados_tabela.HeaderRowRange.Cells(1, c).Value
    Next c
    planilhaNova.Cells(1, totalColunas - 1).Value = "Nome da Regra"
    planilhaNova.Cells(1, totalColunas).Value = "Mensagem de Erro"
    
    
    'Colar dados do colErros
    Dim itemErro As Variant
    Dim formatValor As Variant
    r = 2
    For Each itemErro In colErros
        For c = 1 To totalColunas
            formatValor = itemErro(c)
            If IsDate(formatValor) Then
                planilhaNova.Cells(r, c).Value = CDate(itemErro(c))
            ElseIf IsNumeric(formatValor) Then
                planilhaNova.Cells(r, c).Value = CDbl(itemErro(c))
            Else
                planilhaNova.Cells(r, c).Value = CStr(itemErro(c))
            End If
        Next c
        r = r + 1
    Next itemErro
    
    'Formatações na planilha
    planilhaNova.Columns.AutoFit
    
    'Classificar por Nome da Regra
    Dim qntLin As Long, coluna As Long
    Dim intervalo As Range, classificado As Range
    
    With planilhaNova
        coluna = totalColunas - 1
        qntLin = .Cells(.Rows.Count, coluna).End(xlUp).Row
        Set intervalo = .Range(.Cells(1, 1), .Cells(qntLin, totalColunas))
        Set classificado = .Cells(1, coluna)
        
        intervalo.Sort Key1:=classificado, _
                  Order1:=xlAscending, _
                  Header:=xlYes, _
                  DataOption1:=xlSortNormal
    End With
    
    'Limpar Objetos
    Set planilhaNova = Nothing
    
End Sub


Public Function GarantirTabela(ByVal nomeAba As String, ByVal nomeTabela As String) As ListObject
'Garante que a aba possua uma tabela com o nome esperado.
'Se a tabela existir com outro nome (ex.: Tabela1), ela é renomeada.

    Dim ws As Worksheet
    Dim lo As ListObject

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nomeAba)
    On Error GoTo 0

    If ws Is Nothing Then
        Err.Raise vbObjectError + 1000, , "A planilha '" & nomeAba & "' não foi encontrada."
    End If

    'Caso 1: já existe com o nome correto
    On Error Resume Next
    Set lo = ws.ListObjects(nomeTabela)
    On Error GoTo 0

    If lo Is Nothing Then

        'Caso 2: existe uma única tabela com outro nome -> restaura o nome
        Select Case ws.ListObjects.Count
            Case 0
                Err.Raise vbObjectError + 1001, , "Não há nenhuma tabela na aba '" & nomeAba & "'."
            Case 1
                Set lo = ws.ListObjects(1)
                lo.Name = nomeTabela
            Case Else
                Err.Raise vbObjectError + 1002, , "Há mais de uma tabela na aba '" & nomeAba & "'." & _
                    vbCrLf & "Mantenha apenas a tabela '" & nomeTabela & "'."
        End Select

    End If

    Set GarantirTabela = lo

End Function

Attribute VB_Name = "Módulo6"

Sub PreencherDatasSemanais()

    ' ==============================================================================

    ' VERSÃO "RASTREADOR" (MAIS SEGURA)

    ' Em vez de pular linhas cegamente, este código procura onde está escrito "DE"

    ' e preenche apenas as células necessárias.

    ' ==============================================================================

    Dim planilha As Worksheet

    Set planilha = ActiveSheet

    

    ' CONFIGURAÇÕES

    Dim colRotulo As String: colRotulo = "B" ' Coluna onde está escrito "DE" / "PARA"

    Dim colData As String: colData = "C"     ' Coluna onde as datas serão colocadas

    Dim linhaLimite As Integer: linhaLimite = 21 ' Até onde procurar (para não rodar eternamente)

    

    Dim entradaUsuario As String

    Dim dataBase As Date

    Dim primeiroDia As Date

    Dim ultimoDia As Date

    

    ' 1. Pergunta o Mês/Ano

    entradaUsuario = InputBox("Digite o Mês e Ano (ex: 11/2025)                                                                                                                     Mas lembre-se de Verificar/Ajustar a 1° Semana, Caso for necessário!:", "Gerador Inteligente")

    If entradaUsuario = "" Then Exit Sub

    

    ' Validação

    On Error Resume Next

    dataBase = CDate("01/" & entradaUsuario)

    If Err.Number <> 0 Then

        MsgBox "Data inválida. Use MM/AAAA.", vbExclamation

        Exit Sub

    End If

    On Error GoTo 0

    

    ' Define datas

    primeiroDia = dataBase

    ultimoDia = DateSerial(Year(dataBase), Month(dataBase) + 1, 0)

    

    ' Variáveis de controle

    Dim i As Integer

    Dim contadorSemanas As Integer

    Dim linhaAnteriorPara As Integer ' Guarda a linha do "PARA" da semana anterior

    

    contadorSemanas = 0

    linhaAnteriorPara = 0

    

    ' ==============================================================================

    ' LOOP DE RASTREAMENTO

    ' ==============================================================================

    For i = 1 To linhaLimite

        ' Verifica se a célula na coluna B contém exatamente "DE" (ignorando espaços)

        If UCase(Trim(planilha.Range(colRotulo & i).Value)) = "DE" Then

            

            contadorSemanas = contadorSemanas + 1

            

            ' --- LÓGICA DA PRIMEIRA SEMANA (VALORES FIXOS) ---

            If contadorSemanas = 1 Then

                ' Preenche "DE" (Início Fixo)

                planilha.Range(colData & i).Value = primeiroDia

                

                ' Calcula "PARA" (Fim da primeira semana)

                Dim fimPrimeiraSemana As Date

                fimPrimeiraSemana = primeiroDia + (7 - Weekday(primeiroDia, vbSunday))

                If fimPrimeiraSemana > ultimoDia Then fimPrimeiraSemana = ultimoDia

                

                ' Preenche "PARA" na linha de baixo (i + 1)

                planilha.Range(colData & i + 1).Value = fimPrimeiraSemana

                

            ' --- LÓGICA DAS SEMANAS SEGUINTES (FÓRMULAS) ---

            Else

                ' Fórmula DE: Pega o PARA da semana anterior + 1

                planilha.Range(colData & i).Formula = _

                    "=IF(" & colData & linhaAnteriorPara & ">=DATE(" & Year(dataBase) & "," & Month(dataBase) + 1 & ",0),""""," & colData & linhaAnteriorPara & "+1)"

                

                ' Fórmula PARA: Pega o DE desta semana + 6

                planilha.Range(colData & i + 1).Formula = _

                    "=IF(" & colData & i & "="""","""",MIN(DATE(" & Year(dataBase) & "," & Month(dataBase) + 1 & ",0)," & colData & i & "+6))"

            End If

            

            ' Atualiza a referência para a próxima vez que encontrarmos um "DE"

            ' O "PARA" atual está na linha (i + 1)

            linhaAnteriorPara = i + 1

            

        End If

    Next i



    MsgBox "Concluído! Encontrei e preenchi " & contadorSemanas & " semanas." & vbCrLf & _

           "As linhas sem 'DE' (como totais e outras tabelas) foram preservadas.", vbInformation



End Sub




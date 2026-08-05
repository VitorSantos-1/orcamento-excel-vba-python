Attribute VB_Name = "Módulo5"
Sub Adicionar_Dados_Dinamico()
    Dim wsDest As Worksheet
    Dim wsFonte As Worksheet
    Dim nomeFonte As String
    
    ' 1. Identifica em qual aba você clicou no botão e define a aba Fonte
    Set wsDest = ActiveSheet
    
    Select Case UCase(wsDest.Name)
        Case "LOJA CENTRAL"
            nomeFonte = "Dados Loja Central"
        Case "LOJA NORTE"
            nomeFonte = "Dados Loja Norte"
        Case "LOJA SUL"
            nomeFonte = "Dados Loja Sul"
        Case "LOJA OESTE"
            nomeFonte = "Dados Loja Oeste"
        Case Else
            MsgBox "Execute este botão a partir de uma das abas de loja (loja central, loja norte, etc.).", vbExclamation
            Exit Sub
    End Select

    ' Verifica se a aba de dados existe para evitar erros
    On Error Resume Next
    Set wsFonte = Sheets(nomeFonte)
    On Error GoTo 0
    
    If wsFonte Is Nothing Then
        MsgBox "Não encontrei a aba chamada: " & nomeFonte, vbCritical
        Exit Sub
    End If
    
    ' Desativa a tela piscando para ser super rápido
    Application.ScreenUpdating = False
    
    Dim colDest As Integer
    Dim nomeCategoria As String
    Dim colFonte As Integer
    Dim c As Range
    Dim nomeBusca As String
    Dim nomeCelula As String
    Dim encontrouAlgum As Boolean
    
    encontrouAlgum = False
    
    ' O loop começa na Coluna 2 (B) e pula de 4 em 4 colunas: B(2), F(6), J(10), N(14)...
    For colDest = 2 To 100 Step 4
        ' Pega o nome da categoria que está na linha 22, uma coluna para a direita
        nomeCategoria = wsDest.Cells(22, colDest + 1).Value
        colFonte = 0
        
        If Trim(nomeCategoria) <> "" Then
            ' Limpeza de texto: tira quebra de linha e espaços duplos
            nomeBusca = UCase(Trim(Replace(Replace(Replace(nomeCategoria, vbCrLf, " "), vbLf, " "), Chr(160), " ")))
            Do While InStr(nomeBusca, "  ") > 0
                nomeBusca = Replace(nomeBusca, "  ", " ")
            Loop
            
            ' Procura o nome da categoria nas linhas 1 e 2 da aba Fonte
            For Each c In wsFonte.Range("A1:ZZ2")
                If Trim(c.Value) <> "" Then
                    nomeCelula = UCase(Trim(Replace(Replace(Replace(c.Value, vbCrLf, " "), vbLf, " "), Chr(160), " ")))
                    Do While InStr(nomeCelula, "  ") > 0
                        nomeCelula = Replace(nomeCelula, "  ", " ")
                    Loop
                    
                    ' Se achou o nome exato da categoria
                    If nomeCelula = nomeBusca Then
                        colFonte = c.Column
                        Exit For
                    End If
                End If
            Next c
            
            ' Se encontrou a categoria, vamos achar as colunas corretas dinamicamente
            If colFonte > 0 Then
                encontrouAlgum = True
                
                ' Fallback padrão
                Dim colDataF As Integer: colDataF = colFonte - 1
                Dim colDiaF As Integer: colDiaF = colFonte
                Dim colMetaF As Integer: colMetaF = colFonte + 4
                
                Dim checkCol As Integer
                ' Procura no bloco da categoria (colFonte - 2 até colFonte + 8)
                For checkCol = colFonte - 2 To colFonte + 8
                    If checkCol > 0 Then
                        Dim val1 As String: val1 = UCase(Trim(Replace(Replace(wsFonte.Cells(1, checkCol).Value, vbCrLf, " "), vbLf, " ")))
                        Dim val2 As String: val2 = UCase(Trim(Replace(Replace(wsFonte.Cells(2, checkCol).Value, vbCrLf, " "), vbLf, " ")))
                        
                        ' Procura coluna de Data
                        If val2 = "DATA" Then
                            colDataF = checkCol
                        End If
                        
                        ' Procura coluna de Dia
                        If val2 = "DIA" Or val2 = "DIA " Then
                            colDiaF = checkCol
                        End If
                        
                        ' Procura coluna de Meta / Vendas
                        If val1 = "META DIARIA" Or val1 = "META DIÁRIA" Or val2 = "DE VENDAS" Then
                            colMetaF = checkCol
                        End If
                    End If
                Next checkCol
                
                ' Insere as fórmulas apontando para as colunas encontradas
                wsDest.Range(wsDest.Cells(24, colDest), wsDest.Cells(54, colDest)).FormulaR1C1 = "='" & nomeFonte & "'!R[-21]C" & colDataF
                wsDest.Range(wsDest.Cells(24, colDest + 1), wsDest.Cells(54, colDest + 1)).FormulaR1C1 = "='" & nomeFonte & "'!R[-21]C" & colDiaF
                wsDest.Range(wsDest.Cells(24, colDest + 2), wsDest.Cells(54, colDest + 2)).FormulaR1C1 = "='" & nomeFonte & "'!R[-21]C" & colMetaF
            End If
        End If
    Next colDest
    
    ' Recalcula para garantir que as fórmulas atualizem
    Calculate
    wsDest.Range("B1").Select
    Application.ScreenUpdating = True
    
    ' Feedback para saber o que aconteceu
    If encontrouAlgum Then
        MsgBox "Dados de " & nomeFonte & " adicionados com sucesso!", vbInformation, "Automação Inteligente"
    Else
        MsgBox "O código rodou, mas não encontrou nenhuma categoria correspondente. Verifique se os nomes das categorias no painel estão iguais aos da base de dados.", vbExclamation, "Aviso"
    End If
End Sub

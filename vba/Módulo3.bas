Attribute VB_Name = "Módulo3"

Sub LimparOrcamento()

    Dim ws As Worksheet

    Set ws = ActiveSheet

    

    ' Desativa a atualização de tela para a macro rodar instantaneamente

    Application.ScreenUpdating = False

    

    ' Define as colunas que precisam ser limpas

    Dim colunas As Variant

    colunas = Array("E", "AC", "AG", "AK", "AO", "AS")

    

    ' Loop que passa por cada coluna limpando os blocos de linhas (incluindo a linha 5 de Bebidas)

    Dim col As Variant

    For Each col In colunas

        ws.Range(col & "4:" & col & "27").ClearContents

        ws.Range(col & "37:" & col & "56").ClearContents

        ws.Range(col & "66:" & col & "87").ClearContents

        ws.Range(col & "97:" & col & "117").ClearContents

    Next col

    

    ' Retorna a visualização para o topo da planilha

    ws.Range("A1").Select

    

    ' Reativa a atualização de tela

    Application.ScreenUpdating = True

    

    MsgBox "Orçamento limpo com sucesso!", vbInformation, "Concluído"

End Sub


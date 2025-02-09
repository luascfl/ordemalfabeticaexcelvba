# Ordenador de Planilhas no Excel VBA
Este projeto fornece uma macro em VBA para ordenar as planilhas em uma pasta de trabalho do Excel em ordem alfabética, seja em ordem crescente ou decrescente.

## Recursos Principais
* Ordena as planilhas do Excel em ordem alfabética.
* Permite a ordenação em ordem crescente ou decrescente.
* Simples e fácil de usar.

## Tecnologias Utilizadas
* Microsoft Excel VBA

## Pré-requisitos
* Microsoft Excel (qualquer versão que suporte VBA)

## Instalação
1. **Abra sua pasta de trabalho do Excel.**
2. **Abra o editor de VBA:** Pressione Alt + F11.
3. **Insira um novo módulo:** No editor de VBA, vá para Inserir > Módulo.
4. **Copie e cole o código:** Copie o código abaixo do arquivo `Ordem alfabética Excel VBA.vba` para o módulo.

```vba
Sub Sort_Active_Book()
    Dim i As Integer
    Dim j As Integer
    Dim iAnswer As VbMsgBoxResult

    ' Pergunta ao usuário em qual direção deseja ordenar as planilhas.
    iAnswer = MsgBox("Ordenar planilhas em ordem crescente?" & Chr(10) _
        & "Clicar em Não ordenará em ordem decrescente", _
        vbYesNoCancel + vbQuestion + vbDefaultButton1, "Ordenar Planilhas")

    For i = 1 To Sheets.Count
        For j = 1 To Sheets.Count - 1
            ' Se a resposta for Sim, ordena em ordem crescente.
            If iAnswer = vbYes Then
                If UCase$(Sheets(j).Name) > UCase$(Sheets(j + 1).Name) Then
                    Sheets(j).Move After:=Sheets(j + 1)
                End If
            ' Se a resposta for Não, ordena em ordem decrescente.
            ElseIf iAnswer = vbNo Then
                If UCase$(Sheets(j).Name) < UCase$(Sheets(j + 1).Name) Then
                    Sheets(j).Move After:=Sheets(j + 1)
                End If
            End If
        Next j
    Next i
End Sub
```
5. **Feche o editor de VBA.**

## Uso
1. Abra a pasta de trabalho do Excel contendo as planilhas que você deseja ordenar.
2. Execute a macro `Sort_Active_Book`. Uma caixa de mensagem aparecerá perguntando se você deseja ordenar em ordem crescente (Sim) ou decrescente (Não).
3. As planilhas serão reorganizadas com base na sua seleção.

## Estrutura do Projeto
O projeto consiste em dois arquivos:
* `Ordem alfabética Excel VBA.vba`: Contém o código VBA para ordenar as planilhas.
* `LICENSE`: Contém as informações da Licença MIT.

## Contribuição
Contribuições são bem-vindas! Sinta-se à vontade para enviar pull requests.

## Licença
Este projeto está licenciado sob a Licença MIT - consulte o arquivo [LICENSE](LICENSE) para obter detalhes.

## Tratamento de Erros
A macro inclui tratamento básico de erros usando uma `MsgBox` para solicitar ao usuário a direção da ordenação. No entanto, não há tratamento específico de erros para problemas inesperados no próprio código VBA. Em caso de comportamento inesperado, revise seu arquivo do Excel e o código VBA para quaisquer inconsistências. Você pode potencialmente melhorar o código com tratamento de erros mais robusto (por exemplo, `On Error Resume Next` e verificações de `Err.Number`) para uma solução pronta para produção.
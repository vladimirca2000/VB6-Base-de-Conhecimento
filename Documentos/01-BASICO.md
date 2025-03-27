# 🟢 Perguntas Básicas

## 1. O que é o Option Explicit e por que é importante?

Resposta:


Option Explicit força a declaração explícita de variáveis. Sem ele, variáveis não declaradas são tratadas como Variant, o que pode levar a erros difíceis de rastrear.

``` vb
Option Explicit
Dim x As Integer  ' Declaração obrigatória
x = 10
```

## 2. Como declarar uma variável do tipo Integer em VB6?

Resposta:

Dim idade As Integer. Tipos primitivos incluem String, Boolean, Date, etc. Variáveis não declaradas são Variant (consomem mais memória).

``` vb
Dim idade As Integer
```

## 3.Como exibir uma mensagem ao usuário?

Resposta:

MsgBox "Mensagem". Parâmetros adicionais controlam botões e ícones.

``` vb
' Exemplo 1: Mensagem simples
MsgBox "Olá, usuário!"

' Exemplo 2: Mensagem com título
MsgBox "Operação concluída com sucesso!", vbInformation, "Informação"

' Exemplo 3: Mensagem com botões de sim e não
Dim resposta As Integer
resposta = MsgBox("Deseja continuar?", vbYesNo + vbQuestion, "Confirmação")
If resposta = vbYes Then
    MsgBox "Você escolheu continuar."
Else
    MsgBox "Você escolheu não continuar."
End If

' Exemplo 4: Mensagem com ícone de erro
MsgBox "Ocorreu um erro!", vbCritical, "Erro"
```

## 4. Para que serve o evento Click de um CommandButton?

Resposta:

Executa código quando o botão é clicado. 

``` vb
Private Sub Command1_Click()
    MsgBox "Botão pressionado!"
End Sub
```

## 5. O que é um TextBox e como acessar seu conteúdo?

Resposta:

Controle para entrada de texto. Use Text1.Text para ler/definir o valor.

``` vb
Private Sub Command1_Click()
    If Text1.Text = "" Then
        MsgBox "O TextBox está vazio. Por favor, digite algo."
    Else
        MsgBox "Você digitou: " & Text1.Text
    End If
End Sub


Private Sub Command2_Click()
    Text1.Text = "Texto padrão"
End Sub
```

## 6. O que é o evento Load de um formulário?

Resposta:

É acionado quando o formulário é carregado na memória, mas ainda não é exibido na tela. Esse evento é útil para inicializar variáveis, configurar controles ou realizar outras tarefas de preparação antes que o formulário seja mostrado ao usuário.

``` vb
Private Sub Form_Load()
    Text1.Text = "Digite algo aqui"
    
    Me.Caption = "Formulário de Exemplo"
End Sub
```

## 7. Como comentar uma linha em VB6?

Resposta:

Use ' ou Rem.

``` vb
' Este é um comentário
```

## 8. Qual a diferença entre Dim, Private e Public?

Resposta:

Dim: Escopo local (procedimento).
Private: Acessível apenas no módulo.
Public: Acessível globalmente (em módulos .bas).

``` vb
' Exemplo de uso de Dim, Private e Public em VB6

' Declaração de uma variável pública
Public contadorGlobal As Integer

' Declaração de uma variável privada no módulo
Private contadorModulo As Integer

' Procedimento que usa uma variável local
Private Sub ExemploDim()
    ' Declaração de uma variável local
    Dim contadorLocal As Integer
    contadorLocal = 0
    MsgBox "Contador Local: " & contadorLocal
End Sub

' Procedimento que incrementa o contador global
Private Sub IncrementarContadorGlobal()
    contadorGlobal = contadorGlobal + 1
    MsgBox "Contador Global: " & contadorGlobal
End Sub

' Procedimento que incrementa o contador do módulo
Private Sub IncrementarContadorModulo()
    contadorModulo = contadorModulo + 1
    MsgBox "Contador do Módulo: " & contadorModulo
End Sub
```

## 9. Para que serve e como usar um loop For...Next?

É usado para repetir um bloco de código um número específico de vezes. Ele é útil quando você sabe com antecedência quantas vezes deseja executar o bloco de código.

``` vb
' Exemplo de uso do loop For...Next em VB6
Private Sub ExemploForNext()
    Dim i As Integer
    Dim resultado As String
    resultado = ""

    ' Loop de 1 a 10
    For i = 1 To 10
        resultado = resultado & "Número: " & i & vbCrLf
    Next i

    ' Exibe o resultado em uma MsgBox
    MsgBox resultado
End Sub
```

## 10. O que é o If...Then...Else?

permite executar diferentes blocos de código com base em uma condição. Se a condição for verdadeira, o bloco de código após Then é executado. Se a condição for falsa, o bloco de código após Else (se presente) é executado.

``` vb
' Exemplo de uso do If...Then...Else em VB6
Private Sub ExemploIfThenElse()
    Dim idade As Integer
    idade = 20

    ' Verifica a idade e exibe uma mensagem apropriada
    If idade < 18 Then
        MsgBox "Você é menor de idade."
    ElseIf idade >= 18 And idade < 65 Then
        MsgBox "Você é adulto."
    Else
        MsgBox "Você é idoso."
    End If
End Sub
```
## 🟢 Perguntas Básicas

### 1. O que é o Option Explicit e por que é importante?

Resposta:


Option Explicit força a declaração explícita de variáveis. Sem ele, variáveis não declaradas são tratadas como Variant, o que pode levar a erros difíceis de rastrear.

``` vb
Option Explicit
Dim x As Integer  ' Declaração obrigatória
x = 10
```

### 2. Como declarar uma variável do tipo Integer em VB6?

Resposta:

Dim idade As Integer. Tipos primitivos incluem String, Boolean, Date, etc. Variáveis não declaradas são Variant (consomem mais memória).

```vb
Dim idade As Integer
```

# Documenta��o


## Classe FBR.Arquivo

### FBR.Arquivo.GravarArquivoTexto()

Faz a grava��o de um texto (string) em um arquivo com o Encoding iso-8859-1 (Por default adiciona no final)

**Sintaxe:**
```vb.net

 Public Shared Function GravarArquivoTexto(ByVal NomeArquivo As String, ByVal strTextoAGravar As String,
        Optional ByVal modoArquivo As FileMode = FileMode.Append,
        Optional ByVal modoAcessoArquivo As FileAccess = FileAccess.Write,
        Optional ByVal TipoCodificacao As Codificacoes = Codificacoes.UTF_8) As Boolean
```

**Par�metros:**

- NomeArquivo: Local onde ser� gravado o arquivo, incluindo a extens�o;
- strTextoAGravar: Texto que ser� adicionado dentro do arquivo;
- modoArquivo (padr�o Append): Diz se o texto deve ser acrescido no final (padr�o) ou arquivo deve ser limpo e sobrescrito;
- modoAcessoArquivo (padr�o Write): Descreve a forma como o arquivo ser� aberto;
- TipoCodificacao: O formato do arquivo (� baseado na estrutura de Codifica��es) - exemplos mais adiante

**Retorno**
- Traz um resultado booleano (true ou false) dizendo se houve sucesso
- N�o gera exce��o 


**Exemplos:**

```vb.net
FBR.Arquivo.GravarArquivoTexto("C:\temp\arquivo.txt", "Uma frase curta que ser� inserido dentro do arquivo indicado")
```

```vb.net
FBR.Arquivo.GravarArquivoTexto("C:\temp\arquivo.txt", "Um outro teste no mesmo arquivo, sobrescrevendo os dados anteriores", IO.FileMode.Create)
```
```vb.net
 If (FBR.Arquivo.GravarArquivoTexto("C:\temp\arquivo.txt", "Adicionar Dados a um arquivo existente")) Then
     Console.WriteLine("Dados gravados com sucesso")
 Else
     Console.WriteLine("N�o foi poss�vel gravar os dados")
End If
```

```vb.net
FBR.Arquivo.GravarArquivoTexto("C:\temp\novoArquivo.txt", "Um novo Arquivo formato diferente", , , Arquivo.Codificacoes.UTF_8)
```


### FBR.Arquivo.CriarGravarArquivoTexto
Fun��o que grava um texto dentro de um arquivo (Cria se necess�rio, deleta conteudo e regrava)

**Sintaxe:**
```vb.net
Public Shared Function CriarGravarArquivoTexto(ByVal NomeArquivo As String, ByVal strTextoAGravar As String) As Boolean
 ```

Exemplo:
```vb.net
FBR.Arquivo.CriarGravarArquivoTexto("C:\temp\arquivoExemplo.txt", "Esse texto sera colocado dentro do arquivo, sobrepondo os dados anteriores")
 ```

# FBR.Arquivo.ProcessarFormatoDiretorio

 Fun��o que processa o formato do diret�rio para garantir que a saida contenha um Path Valido e terminado por "\\"

**Sintaxe:**
```vb.net
Public Shared Function ProcessarFormatoDiretorio(ByVal diretorio As String) As String
```

Exemplo
```vb.net
Console.WriteLine(FBR.Arquivo.ProcessarFormatoDiretorio("C:\TEMP"))
' Retorna C:\TEMP\
``` 

```vb.net
'Exemplo de um "problema" 
Dim ArquivoLOG As String
ArquivoLOG = Environment.CurrentDirectory & "meuprograma.log"
Console.WriteLine(ArquivoLOG)
' Retorna C:\Meus Programas\TesteFBR\bin\Debug\net5.0meuprograma.log  (veja que o erro � a falta da barra)
``` 

```vb.net
' Exemplo de um Bom uso!
Dim ArquivoLOG As String
ArquivoLOG = FBR.Arquivo.ProcessarFormatoDiretorio(Environment.CurrentDirectory) & "meuprograma.log"
Console.WriteLine(ArquivoLOG)
' Retorna C:\Meus Programas\TesteFBR\bin\Debug\net5.0\meuprograma.log  (resultado OK)
 ``` 


### FBR.Arquivo.PegarTamanhoArquivo

Obtem o tamanho em Bytes de um Arquivo em Disco

**Sintaxe:**
```vb.net
Public Shared Function PegarTamanhoArquivo(ByVal NomeArquivo As String) As Long
```

**Par�metros**
- NomeArquivo: Local completo do arquivo (incluindo diretorio, nome e extens�o)

**Retorno**
- Traz um resultado da quantidade de bytes do arquivo (n�o uso de disco, mas o tamanho do arquivo bruto)
- Retorna -1 se der erro (arquivo inexistente).
- N�o retorna exce��o

**Exemplo:**
```vb.net
Console.WriteLine(FBR.Arquivo.PegarTamanhoArquivo("C:\temp\arquivo.txt"))
' Retorna 108
```

### FBR.Arquivo.SHA1Arquivo

Obtem uma string representando o hexadecimal min�sculo do SHA1 do resultado do arquivo

**Sintaxe:**
```vb.net
 Public Shared Function SHA1Arquivo(ByVal LocalArquivo As String) As String
 ```

**Par�metros**
- LocalArquivo: Local completo do arquivo (incluindo diretorio, nome e extens�o)

**Retorno**
- Retorna uma string representando um hexadecimal do SHA1 do conteudo do arquivo
- Retorna "" (vazio) se der erro .
- N�o retorna exce��o

**Exemplo:**
```vb.net
Console.WriteLine(FBR.Arquivo.SHA1Arquivo("C:\temp\arquivo.txt"))
' Retorna c4125e77200868c3622418ae99975f49dc570534
```


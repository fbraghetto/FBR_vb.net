# Documentação


## Classe FBR.Arquivo

### FBR.Arquivo.GravarArquivoTexto()

Faz a gravação de um texto (string) em um arquivo com o Encoding iso-8859-1 (Por default adiciona no final)

**Sintaxe:**
```vb.net

 Public Shared Function GravarArquivoTexto(ByVal NomeArquivo As String, ByVal strTextoAGravar As String,
        Optional ByVal modoArquivo As FileMode = FileMode.Append,
        Optional ByVal modoAcessoArquivo As FileAccess = FileAccess.Write,
        Optional ByVal TipoCodificacao As Codificacoes = Codificacoes.UTF_8) As Boolean
```

**Parâmetros:**

- NomeArquivo: Local onde será gravado o arquivo, incluindo a extensão;
- strTextoAGravar: Texto que será adicionado dentro do arquivo;
- modoArquivo (padrão Append): Diz se o texto deve ser acrescido no final (padrão) ou arquivo deve ser limpo e sobrescrito;
- modoAcessoArquivo (padrão Write): Descreve a forma como o arquivo será aberto;
- TipoCodificacao: O formato do arquivo (é baseado na estrutura de Codificações) - exemplos mais adiante

**Retorno**
- Traz um resultado booleano (true ou false) dizendo se houve sucesso
- Não gera exceção 


**Exemplos:**

```vb.net
FBR.Arquivo.GravarArquivoTexto("C:\temp\arquivo.txt", "Uma frase curta que será inserido dentro do arquivo indicado")
```

```vb.net
FBR.Arquivo.GravarArquivoTexto("C:\temp\arquivo.txt", "Um outro teste no mesmo arquivo, sobrescrevendo os dados anteriores", IO.FileMode.Create)
```
```vb.net
 If (FBR.Arquivo.GravarArquivoTexto("C:\temp\arquivo.txt", "Adicionar Dados a um arquivo existente")) Then
     Console.WriteLine("Dados gravados com sucesso")
 Else
     Console.WriteLine("Não foi possível gravar os dados")
End If
```

```vb.net
FBR.Arquivo.GravarArquivoTexto("C:\temp\novoArquivo.txt", "Um novo Arquivo formato diferente", , , Arquivo.Codificacoes.UTF_8)
```


### FBR.Arquivo.CriarGravarArquivoTexto
Função que grava um texto dentro de um arquivo (Cria se necessário, deleta conteudo e regrava)

**Sintaxe:**
```vb.net
Public Shared Function CriarGravarArquivoTexto(ByVal NomeArquivo As String, ByVal strTextoAGravar As String) As Boolean
 ```

Exemplo:
```vb.net
FBR.Arquivo.CriarGravarArquivoTexto("C:\temp\arquivoExemplo.txt", "Esse texto sera colocado dentro do arquivo, sobrepondo os dados anteriores")
 ```

# FBR.Arquivo.ProcessarFormatoDiretorio

 Função que processa o formato do diretório para garantir que a saida contenha um Path Valido e terminado por "\\"

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
' Retorna C:\Meus Programas\TesteFBR\bin\Debug\net5.0meuprograma.log  (veja que o erro é a falta da barra)
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

**Parâmetros**
- NomeArquivo: Local completo do arquivo (incluindo diretorio, nome e extensão)

**Retorno**
- Traz um resultado da quantidade de bytes do arquivo (não uso de disco, mas o tamanho do arquivo bruto)
- Retorna -1 se der erro (arquivo inexistente).
- Não retorna exceção

**Exemplo:**
```vb.net
Console.WriteLine(FBR.Arquivo.PegarTamanhoArquivo("C:\temp\arquivo.txt"))
' Retorna 108
```

### FBR.Arquivo.SHA1Arquivo

Obtem uma string representando o hexadecimal minúsculo do SHA1 do resultado do arquivo

**Sintaxe:**
```vb.net
 Public Shared Function SHA1Arquivo(ByVal LocalArquivo As String) As String
 ```

**Parâmetros**
- LocalArquivo: Local completo do arquivo (incluindo diretorio, nome e extensão)

**Retorno**
- Retorna uma string representando um hexadecimal do SHA1 do conteudo do arquivo
- Retorna "" (vazio) se der erro .
- Não retorna exceção

**Exemplo:**
```vb.net
Console.WriteLine(FBR.Arquivo.SHA1Arquivo("C:\temp\arquivo.txt"))
' Retorna c4125e77200868c3622418ae99975f49dc570534
```


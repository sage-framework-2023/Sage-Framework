# SageTypes

Tipos para o VBA escritos em C#, no lugar das classes do `Sage.xlam` (legado, que será aposentado). Funcionam em qualquer pasta de trabalho sem precisar do suplemento.

O VBA os vê como a biblioteca **Sage**, então o código continua escrevendo `Sage.StringS`, como já fazia com o `Sage.xlam`.

Por enquanto:
- **StringS**, com a mesma API do `StringS` do `Sage.xlam`;
- **DictionaryS**, com a API do `DictionaryS` do `Sage.xlam` e o comportamento e os métodos do dicionário do Python;
- **ListS**, com a API do `ListS` do `Sage.xlam` e o comportamento e os métodos da lista do Python.
- **DataFrame**, tabela no estilo do pandas sobre o DuckDB, para milhões de linhas;
- **DateTimeS**, data e hora como o `datetime` do Python, com fuso horário;
- **Json**, o módulo `json` do Python;
- **Requests**, a biblioteca `requests` do Python, para chamar APIs (com **Response** e **Session**).

## Instalar

1. Feche o Excel.
2. Rode:

   ```
   powershell -ExecutionPolicy Bypass -File install.ps1
   ```

   O script compila `src\*.cs` com o `csc` do .NET Framework 4, que já vem no Windows, copia o DLL e a `duckdb.dll` para `%LOCALAPPDATA%\Sage\Types`, gera a biblioteca de tipos (`SageTypes.tlb`, com o nome de biblioteca `Sage`) e registra tudo só no usuário atual (HKCU), sem administrador.
3. Com o SageEditor instalado, a referência é marcada sozinha quando você abre o código de um projeto (*Editor: Referência Automática* nas Configurações). Sem ele: **Ferramentas > Referências > Sage Framework**.

Uma pasta de trabalho não pode referenciar ao mesmo tempo esta biblioteca e o projeto `Sage` do `Sage.xlam`, porque os nomes são iguais. Remova a referência ao `Sage.xlam` antes de adicionar esta.

Para remover: `install.ps1 -Uninstall`.

A `duckdb.dll` (64 bits, assinada pela DuckDB Foundation) vem em `lib\` no pacote da release. Num clone do git ela não existe, e o `install.ps1` baixa a versão fixada e confere o hash SHA-256. Para montar o instalador da release (um `.exe` só, com tudo dentro): `build-release.ps1 -Version x.y.z` na raiz do repositório; ele gera `dist\SageFramework-Setup-x.y.z.exe` (licença do DuckDB em `THIRD-PARTY-NOTICES.txt`).

## Uso

```vb
Dim s As Sage.StringS
Set s = New Sage.StringS
s = "  olá\tmundo  "                  ' membro padrão (Value)
Debug.Print s.Upper.Replace("MUNDO", "VBA")

Dim t As Object                       ' sem referência (late binding)
Set t = CreateObject("Sage.StringS")
```

O nome precisa do prefixo `Sage.` no `Dim` e no `New`. O VBA não diferencia maiúsculas de minúsculas, e `StringS` é igual a `Strings`, um módulo da própria biblioteca VBA, que sempre vem antes das outras referências. Nas chamadas de métodos (`s.Upper`) o prefixo não é necessário.

## Valores dentro de ListS e DictionaryS

Os tipos devolvem os tipos do Sage sempre que possível, como o `Sage.xlam` fazia:

```vb
Dim lst As New Sage.ListS
lst = Array("Teste 2", "Teste 1", Array(1, 2))
Debug.Print lst(0).Upper()            ' TESTE 2: texto volta como StringS
lst(2).Append 3                       ' array guardado vira ListS e é alterado no lugar
```

- **Ao ler** (`l(i)`, `d(chave)`, `Item`, `GetItem`, `Pop`, `Min`, `Max`, `For Each`...), texto volta como `StringS`. Onde o VBA espera texto (`Debug.Print`, `x = lst(0)`, `lst(0) = "a"`, `Len`, `&`), ele usa o `Value` do `StringS`; `TypeName(lst(0))` é `"StringS"`.
- **Ao guardar**, array vira `ListS` (os de dentro também; um array 2D, como `Range.Value`, vira lista de linhas), e um `StringS` é guardado como texto.
- **Ficam como estão:** números, datas e objetos (`Scripting.Dictionary`, `Collection`, `Range`...), guardados por referência.
- **VBA "puro":** `l.Value` sem índice, `l.ToArray` e `d.Value` sem chave devolvem arrays e texto comuns, com as listas internas também como arrays, para `UBound`, `Join` do VBA etc.

## DictionaryS

```vb
Dim d As New Sage.DictionaryS
d("nome") = "Ana"                     ' membro padrão (Value)
d(1) = 10                             ' chave de qualquer tipo
Set d("filho") = New Sage.DictionaryS ' objetos com Set
Debug.Print d.ToString                ' {'nome': 'Ana', 1: 10, 'filho': {}}

Dim k As Variant
For Each k In d                       ' chaves, na ordem de inserção
    Debug.Print k, d.GetItem(k)
Next k
```

| Python | DictionaryS |
|---|---|
| `d[k]`, `d[k] = v` | `d(k)`, `d(k) = v` / `Set d(k) = objeto` |
| `d[k]` com chave inexistente: `KeyError` | erro 9, "KeyError: 'k'" |
| `k in d`, `len(d)` | `d.Contains(k)` (ou `Exists`), `d.Count` (ou `Length`) |
| `d.get(k, padrão)` | `d.GetItem(k, padrão)` |
| `d.keys()`, `d.values()`, `d.items()` | `d.Keys`, `d.Values` (ou `Itens`), `d.Items` (pares `Array(k, v)`) |
| `d.pop(k[, padrão])`, `d.popitem()` | `d.Pop(k[, padrão])`, `d.PopItem` |
| `d.setdefault(k[, v])`, `d.update(outro)` | `d.SetDefault(k[, v])`, `d.Update(outro)` |
| `dict.fromkeys(chaves[, v])`, `d.copy()`, `d.clear()` | `d.FromKeys(chaves[, v])`, `d.Copy`, `d.Clear` |
| `del d[k]` | `d.Remove k` (sem erro se não existir) |
| `repr(d)` | `d.ToString` |

- **Chaves:** como no Python, números são comparados pelo valor (`1`, `1#` e `CLng(1)` são a mesma chave) e texto diferencia maiúsculas; `"1"` é outra chave. Um `StringS` como chave vira o texto dele. Arrays não podem ser chave (erro 13).
- **Fontes aceitas** por `Update`, `Append` e `d = ...`: outro `DictionaryS`, um `Scripting.Dictionary`, `Array(k1, v1, k2, v2...)` ou `Array(Array(k1, v1), ...)`.
- **Sem chave:** `d = fonte` substitui todo o conteúdo; ler `d.Value` devolve um `Scripting.Dictionary` com tudo, como no `Sage.xlam`.

### Diferenças em relação ao DictionaryS do Sage.xlam

- Chave inexistente gera erro (`KeyError`) em vez de devolver `Empty`, como no Python. Para ler com valor padrão: `GetItem(chave, padrão)`.
- As chaves mantêm o tipo (o `Sage.xlam` convertia tudo para texto): `d(1)` e `d("1")` são chaves diferentes.
- Arrays guardados viram `ListS` (no `Sage.xlam`, só ao ler); `Items` e `PopItem` devolvem pares como `ListS`.

## ListS

```vb
Dim l As New Sage.ListS
l = Array(3, 1, 2)                    ' membro padrão (Value)
l.Append 4                            ' altera a própria lista
Debug.Print l(0), l(-1)               ' 3  4 (negativo conta do fim)
Debug.Print l.Slice(1, 3).ToString    ' [1, 2]
Debug.Print l.Sort.ToString           ' [1, 2, 3, 4]
Debug.Print l.Join(", ")              ' 1, 2, 3, 4 (StringS)
```

| Python | ListS |
|---|---|
| `l[i]`, `l[i] = v`, `l[-1]` | `l(i)`, `l(i) = v` / `Set l(i) = objeto`, `l(-1)` |
| índice fora do intervalo: `IndexError` | erro 9, "IndexError: list index out of range" |
| `len(l)`, `x in l` | `l.Count` (ou `Length`), `l.Contains(x)` |
| `l.append(x)`, `l.extend(it)`, `l.insert(i, x)` | `l.Append x`, `l.Extend it`, `l.Insert i, x` |
| `l.pop([i])`, `l.remove(x)`, `del l[i]`, `l.clear()` | `l.Pop([i])`, `l.RemoveValue x`, `l.Remove i`, `l.Clear` |
| `l.index(x[, ini[, fim]])`, `l.count(x)` | `l.Index(x[, ini[, fim]])`, `l.CountOf(x)` |
| `l.sort(reverse=True)`, `l.reverse()`, `l.copy()` | `l.Sort(True)`, `l.Reverse`, `l.Copy` |
| `l[ini:fim:passo]` | `l.Slice(ini, fim, passo)` (qualquer um pode ser omitido: `l.Slice(, , -1)`) |
| `l1 + l2`, `l * n` | `l1.Concat(l2)`, `l.Repeat(n)` |
| `sum(l)`, `min(l)`, `max(l)` | `l.Sum`, `l.Min`, `l.Max` |
| `repr(l)`, `tuple(l)` | `l.ToString`, `l.ArrayType = SgTuple` |

- **Métodos que alteram a lista** (`Append`, `Remove`, `Extend`, `Insert`, `RemoveValue`, `Clear`, `Sort`, `Reverse`) também a devolvem. Funciona tanto `l.Append x` quanto o encadeamento do `Sage.xlam`: `.Split(vbNewLine).Remove(-1).Join(vbNewLine)`.
- **Iteráveis** aceitos por `l = ...`, `Extend` e `Concat`: array (2D: cada linha vira uma `ListS`), outro `ListS`, `DictionaryS` (as chaves), `Collection`, `Range` e texto (os caracteres).
- **Comparação** (`Contains`, `Index`, `CountOf`, `RemoveValue`): números pelo valor (`21 = 21#`), texto diferenciando maiúsculas, objetos pela identidade.
- **Ordenação** estável, como no Python: números pelo valor, texto pela ordem dos caracteres (maiúsculas antes das minúsculas); texto e número misturados geram erro 13.
- **Tupla:** com `ArrayType = SgTuple`, qualquer alteração gera erro 13; `ToString` mostra `(1, 2)`.
- `l.Value` sem índice (ou `l.ToArray`) devolve um array do VBA, para `UBound`, `Join` do VBA etc.

### Diferenças em relação ao ListS do Sage.xlam

- `Append` e `Remove` alteram a própria lista (no `Sage.xlam` devolviam uma cópia e a original ficava igual). Como continuam devolvendo a lista, `l = l.Append(x)` dá o mesmo resultado.
- Índice fora do intervalo gera `IndexError` (erro 9) com a mensagem do Python.
- Arrays viram `ListS` ao guardar (no `Sage.xlam`, uma cópia a cada leitura), então alterar uma lista interna altera a lista guardada.

## DataFrame

Tabela no estilo do pandas, feita para volumes grandes (dezenas de milhões de linhas). Os dados ficam no [DuckDB](https://duckdb.org), um banco analítico embutido (`duckdb.dll`, junto do `SageTypes.dll`), e não em arrays do VBA. Cada operação vira SQL, executado em paralelo e só quando o resultado é pedido. O banco é um arquivo temporário por processo do Excel (`%TEMP%\SageDuckDB`, apagado ao fechar), com as tabelas comprimidas no disco: só o que está em uso fica na memória, limitada a 40% da RAM.

```vb
Dim df As New Sage.DataFrame
df.ReadCsv "C:\dados\vendas.csv"
Debug.Print df                                   ' tabela como no pandas (5 primeiras + 5 últimas)

Dim resumo As Sage.DataFrame
Set resumo = df.Query("Valor > 100 And Cidade = ""SP""") _
               .GroupBy("Produto", "Total = sum(Valor), N = count(*)") _
               .SortValues("Total", False)
resumo.ToRange Sheets("Resumo").Range("A1")

df.Eval "Total = Valor * Qtd"                    ' coluna nova (ou substituída), em SQL
df("Ativo") = True                               ' coluna inteira

df.MoveFirst                                     ' cursor, linha a linha
Do Until df.EOF
    df("Valor") = df("Valor") / 100
    df.MoveNext
Loop
```

**Carregar:** `ReadCsv(Caminho, [Cabeçalho], [Delimitador], [SeparadorDecimal], [Codificação])`, `ReadParquet`, `ReadExcel(Caminho, [Planilha], [Cabeçalho])`, `FromRange`, `FromArray` (2D ou array de linhas), `FromRecords` (lista de `DictionaryS`), `Sql("SELECT ... FROM {0} JOIN {1} ...", outroDf)`.

**Gravar e ver:** `ToCsv`, `ToParquet`, `ToExcel` (formato pela extensão), `ToRange`, `ToArray`, `ToString` / `Debug.Print df`, `Show` (janela com grade, lida aos poucos: abre na hora mesmo com milhões de linhas).

**Explorar:** `Count`, `Columns`, `DTypes`, `Shape`, `Info`, `Describe`, `Head`, `Tail`, `Slice`, `Sample`, `Col`, `Unique`, `ValueCounts`, `Sum`, `Mean`, `Min`, `Max`, `Median`, `Std`, `NUnique`.

**Transformar** (devolvem um DataFrame novo; o original fica igual, menos em `Eval` com `=` e atribuições): `Query`, `Select` (aceita `"Nova = expressão"`), `Drop`, `Rename`, `SortValues`, `DropDuplicates`, `DropNA`, `FillNA`, `GroupBy(Chaves, "Nome = agregação, ...")` ou com um `DictionaryS` `{coluna: "sum"}`, `Pivot`, `Merge(Outro, On, How, LeftOn, RightOn)` (sufixos `_x`/`_y`), `Concat` (por nome de coluna), `Copy`.

Nas expressões (`Query`, `Eval`, `Select`, `GroupBy`), texto vai entre aspas duplas como no VBA (`"SP"`) e nomes com espaço entre colchetes (`[Valor Total]`); o resto é SQL do DuckDB (`round`, `year(Data)`, `Is Null`, `Like`...).

**Linhas e células:** `df.At(linha, "Coluna")` lê e grava uma célula; `df.Loc(linha)` devolve a linha como `DictionaryS`; `df.Loc(df.Count) = Array(...)` acrescenta uma linha, como `df.loc[len(df)] = [...]`; `For Each linha In df` percorre as linhas como `DictionaryS`. O cursor (`MoveFirst`, `MoveNext`, `MovePrevious`, `MoveLast`, `Move`, `EOF`, `BOF`, `Index`) lê blocos de 65.536 linhas e grava as alterações em lote, mas ainda passa por cada linha no VBA: para milhões de linhas, prefira `Eval`/`Query`/`GroupBy`, que rodam no DuckDB.

`df("Coluna")` devolve a coluna como `ListS` fora do cursor e a célula da linha atual dentro dele. Uma coluna inteira recebe um valor (`df("X") = 0`) ou um array/`ListS` com um valor por linha.

### No lugar das funções do DataFrame do Sage.xlam

| Sage.xlam | DataFrame |
|---|---|
| `AddColumn` | `df("Nova") = valor`, `df("Nova") = array`, `df.Eval "Nova = expressão"` |
| `AddRow` | `df.Loc(df.Count) = Array(...)`; muitas linhas: `df.Concat(outro)` |
| remover coluna/linhas | `df.Drop("Col")`, `df.Query("condição")` |
| `Data` (array interno) | `df.ToArray` / `df.FromArray` |

O tipo de uma coluna se alarga sozinho: um inteiro que recebe um decimal vira `DOUBLE`; um número que recebe texto vira `VARCHAR`.

## DateTimeS

```vb
Dim d As New Sage.DateTimeS
d = Now                                          ' ou Date, "2024-01-05T10:30:00Z", "15/03/2024"
Debug.Print d.StrFTime("%d/%m/%Y %H:%M")         ' códigos do strftime do Python
Debug.Print d.AddMonths(1).IsoFormat             ' 31/01 + 1 mês = 29/02
Debug.Print d.Add(Days:=7, Hours:=-2).ToString
Debug.Print d.UtcNow.AsTimeZone("E. South America Standard Time").IsoFormat   ' ...-03:00
```

Como no Python, uma data não muda: os métodos devolvem um `DateTimeS` novo (`Set d = d.AddMonths(1)`). O membro padrão (`Value`) é um `Date` do VBA, então `d > DateSerial(2024, 1, 1)`, `Format(d, ...)` e `Range("A1") = d.Value` funcionam.

| Python | DateTimeS |
|---|---|
| `datetime.now()`, `date.today()`, `datetime.now(timezone.utc)` | `d.Now`, `d.Today`, `d.UtcNow` |
| `datetime(2024, 1, 5, 10, 30)` | `d.Create(2024, 1, 5, 10, 30)` |
| `strptime(texto, "%d/%m/%Y")`, `strftime(...)` | `d.StrPTime(texto, "%d/%m/%Y")`, `d.StrFTime(...)` |
| `fromisoformat(...)`, `isoformat()`, `str(d)` | `d.FromIsoFormat(...)`, `d.IsoFormat`, `d.ToString` |
| `fromtimestamp(s)`, `timestamp()` | `d.FromTimestamp(s)`, `d.Timestamp` |
| `d + timedelta(days=1)`, `d + relativedelta(months=1)` | `d.Add(Days:=1)`, `d.AddMonths(1)` ou `d.Add(Months:=1)` |
| `(d1 - d2).total_seconds()` | `d1.Diff(d2, "seconds")` (padrão: dias) |
| `d.replace(day=1)`, `d.astimezone(tz)`, `d.utcoffset()` | `d.Replace(Day:=1)`, `d.AsTimeZone(tz)`, `d.UtcOffset` (horas) |
| `d.year`, `d.weekday()`, `d.isocalendar()` | `d.Year`, `d.Weekday` (segunda = 0), `d.IsoCalendar` |
| pandas: `quarter`, `days_in_month`, `normalize()` | `d.Quarter`, `d.DaysInMonth`, `d.Normalize` |

- **Fuso horário:** sem fuso (`TimeZone = ""`) a data é "ingênua", como no Python. Fusos aceitos: `"UTC"`, `"local"`, deslocamentos (`"-03:00"`) e nomes do Windows (`"E. South America Standard Time"`; lista com `tzutil /l`). Os nomes da IANA (`"America/Sao_Paulo"`) não existem no .NET Framework.
- **Nomes de mês e dia** (`%b`, `%A`) seguem o idioma do Windows, ao escrever e ao ler.
- **Texto** atribuído (`d = "..."`) é lido como ISO 8601; se não for, no formato do Windows (dd/mm/aaaa no pt-BR).

## Json

Global, como o módulo do Python: usado sem `Dim` (ou como `Sage.Json`).

```vb
Dim d As Sage.DictionaryS
Set d = Json.Loads("{""nome"": ""Ana"", ""itens"": [1, 2.5, null]}")
Debug.Print d("nome").Upper, d("itens")(1)       ' ANA  2,5
Debug.Print Json.Dumps(d, Indent:=2, SortKeys:=True)
Json.Dump d, "C:\dados\saida.json"               ' UTF-8
Set d = Json.Load("C:\dados\saida.json")
```

| JSON | VBA (Loads) |
|---|---|
| objeto, lista | `DictionaryS`, `ListS` (na ordem do texto) |
| texto | `StringS` |
| número inteiro / decimal | `Long` (ou `LongLong`, se não couber) / `Double` |
| `true`, `false`, `null` | `True`, `False`, `Empty` |

`Dumps` aceita também `Scripting.Dictionary`, arrays (2D: lista de linhas), `Collection`, `Date` e `DateTimeS` (texto ISO 8601) e `DataFrame` (lista de linhas). Erros no texto geram `JSONDecodeError` com linha e coluna, como no Python. Diferença do Python: `EnsureAscii` é `False` por padrão (acentos ficam como estão); com `EnsureAscii:=True`, saem como `\u00e7`.

## Requests

Global, como o módulo do Python: `Requests.Get(...)`, sem `Dim`. Devolve um `Sage.Response`.

```vb
Dim r As Sage.Response
Set r = Requests.Get("https://economia.awesomeapi.com.br/json/last/USD-BRL")
r.RaiseForStatus                                 ' erro se o status for 4xx ou 5xx
Debug.Print Val(r.Json()("USDBRL")("bid"))       ' cotação do dólar (texto com ponto: Val, não CDbl)

Dim dados As New Sage.DictionaryS
dados("nome") = "Ana"
Set r = Requests.Post("https://httpbin.org/post", Json:=dados, _
                      Headers:=Array("Authorization", "Bearer " & token), Timeout:=30)
```

| Python | Sage |
|---|---|
| `requests.get(url, params=..., headers=..., auth=..., timeout=...)` | `Requests.Get(url, Params:=..., Headers:=..., Auth:=..., Timeout:=...)` |
| `requests.post(url, data=..., json=...)` (e `put`, `patch`) | `Requests.Post(url, Data:=..., Json:=...)` (e `Put`, `Patch`) |
| `requests.delete`, `head`, `request(método, url, ...)` | `Requests.Delete`, `Head`, `Request(método, url, ...)` |
| `r.status_code`, `r.ok`, `r.reason` | `r.StatusCode`, `r.Ok`, `r.Reason` |
| `r.text`, `r.content`, `r.json()` | `r.Text` (`StringS`), `r.Content` (`Byte()`), `r.Json()` |
| `json.dumps(r.json(), indent=2)` | `r.JsonString` (uma linha) ou `r.JsonString(Indent:=2)` (formatado); também `SortKeys`, `EnsureAscii` |
| `r.headers["content-type"]` | `r.Header("content-type")` (sem diferenciar maiúsculas) ou `r.Headers` (`DictionaryS`) |
| `r.url`, `r.elapsed`, `r.encoding`, `r.raise_for_status()` | `r.Url`, `r.Elapsed` (segundos), `r.Encoding`, `r.RaiseForStatus` |
| `s = requests.Session()` | `Set s = Requests.Session()` (`Sage.Session`) |

- **Params, Headers e Data** aceitam `DictionaryS`, `Scripting.Dictionary` ou `Array(chave1, valor1, ...)`. Um valor lista repete a chave (`?id=1&id=2`).
- **Corpo:** `Json:=` envia JSON (`application/json`); `Data:=` com dicionário envia formulário; com texto ou `Byte()`, envia como está.
- **Auth:** `Array("usuário", "senha")` (autenticação básica). Tokens vão em `Headers`.
- **Status 4xx/5xx** não geram erro, como no requests: confira `r.Ok` ou chame `r.RaiseForStatus`. Falha de rede gera `ConnectionError`; tempo esgotado, `Timeout` (erro 5).
- **Session:** `s.Headers("X-Api-Key") = chave`, `s.Params`, `s.Auth` e `s.Timeout` valem para todas as chamadas da sessão, e os cookies (login) ficam guardados entre elas (`s.Cookies` os lista).
- **Downloads:** `r.Save "C:\arquivo.pdf"`.
- Usa TLS 1.2/1.3 e o proxy do Windows (com o usuário logado), como o navegador. A chamada é síncrona: o Excel espera a resposta.

## Diferenças em relação ao StringS do Sage.xlam

- `FString` e `Join` aceitam até 30 argumentos, e não uma `ParamArray` ilimitada: o VBA recusa a `ParamArray` exportada pelo .NET.
- `Count("an")` conta ocorrências (2 em "banana"); a versão VBA contava caracteres removidos (4).
- `Mid(3)` sem tamanho vai até o fim; na versão VBA o padrão `"End"` dava erro de tipo.
- `Proper` segue as mesmas regras do PROPER do Excel, sem chamar o Excel.

Mantidos de propósito, como no original: ao atribuir um valor, `\n` vira quebra de linha, `\t` vira espaços até a próxima coluna múltipla de 4 e o texto perde espaços e quebras das pontas (por isso `", "` vira `","`).

## Para acrescentar membros

Todos os nomes (tipos, membros, parâmetros, enums) seguem o PascalCase, sem sublinhado: `ReadCsv`, `SortValues`, `SgTuple`.

As interfaces (`_StringS`, `_DictionaryS`, `_ListS`, `_DataFrame`, `_DateTimeS`, `_Json`, `_Requests`, `_Session`, `_Response`, `_Globals`) são duais: o VBA as chama pela vtable. Acrescente membros **sempre no fim**, com o próximo `DispId`, e nunca reordene nem remova os existentes, senão o código VBA já compilado chama o método errado. Uma classe nova precisa de `[Guid]`, `[ProgId]`, interface própria e uma linha em `$Classes` no `install.ps1`.

Limitações do exportador do .NET (`TypeLibConverter`) e como contorná-las:
- **Propriedade Variant com `Let`:** o .NET exporta o setter só como `Property Set`. Declare também um método `LetNome(...)` com `[PropertyLet("Nome")]`, os mesmos parâmetros e mais o valor; o `install.ps1` o transforma no `Property Let` de `Nome` ao gerar o `.tlb`.
- **Propriedade com parâmetros** (como `At(linha, coluna)`): o .NET só exporta o indexador. Declare um método `GetNome(...)` com `[PropertyGet("Nome")]`; no VBA ele é lido como `Nome(...)`.
- **`ParamArray`:** o `params` do C# não é aceito pelo VBA; use parâmetros `[Optional]`.
- **Membros de enum:** o .NET os exporta como `Enum_Membro` (`SgArrayTypes_SgTuple`); o `install.ps1` tira o prefixo, e o VBA vê `SgTuple`.
- **Membros globais** (`Json`, `Requests`): ficam na classe `Globals`, com `[AppObject]`; o `install.ps1` a marca como *app object* no `.tlb`, e o VBA a cria sozinho.
- **Nome igual a algo da biblioteca VBA** (como `Strings`): o VBA acha o dele primeiro, por isso o prefixo `Sage.`.

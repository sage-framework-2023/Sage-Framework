# SageTypes

Tipos para o VBA escritos em C#, no lugar das classes do `Sage.xlam` (legado, que será aposentado). Funcionam em qualquer pasta de trabalho sem precisar do suplemento.

O VBA os vê como a biblioteca **Sage**, então o código continua escrevendo `Sage.StringS`, como já fazia com o `Sage.xlam`.

Por enquanto:
- **StringS**, com a mesma API do `StringS` do `Sage.xlam`;
- **DictionaryS**, com a API do `DictionaryS` do `Sage.xlam` e o comportamento e os métodos do dicionário do Python;
- **ListS**, com a API do `ListS` do `Sage.xlam` e o comportamento e os métodos da lista do Python.
- **DataFrame**, tabela no estilo do pandas sobre o DuckDB, para milhões de linhas;
- **DatetimeS**, data e hora como o `datetime` do Python, com fuso horário;
- **Json**, o módulo `json` do Python;
- **Requests**, a biblioteca `requests` do Python, para chamar APIs (com **Response** e **Session**);
- **SqlEngine**, conexão a bancos de dados (PostgreSQL, MySQL, SQLite, SQL Server, Access, ODBC), como o `Engine` do SQLAlchemy usado pelo pandas.

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
| `d | outro` (Python 3.9) | `d.Union(outro)` (dicionário novo; `Update` altera o próprio) |
| `Scripting.Dictionary`: `Add`, `Exists`, `RemoveAll` | `d.Add k, v` (erro 457 se a chave já existir), `d.Exists(k)`, `d.RemoveAll` |
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
| `sorted(l)`, `list(reversed(l))` | `l.Sorted` / `l.Sorted(True)`, `l.Reversed` (listas novas; a original fica igual) |
| `l.sort(key=lambda r: r[2])` (lista de linhas) | `l.SortBy(2)`; várias colunas: `l.SortBy(Array(2, 0))`; linhas que são dicionários: `l.SortBy("nome")` |
| `list(dict.fromkeys(l))` (sem repetidos, na ordem) | `l.Unique` |
| `l[ini:fim:passo]` | `l.Slice(ini, fim, passo)` (qualquer um pode ser omitido: `l.Slice(, , -1)`) |
| `l1 + l2`, `l * n` | `l1.Concat(l2)`, `l.Repeat(n)` |
| `sum(l)`, `min(l)`, `max(l)` | `l.Sum`, `l.Min`, `l.Max` |
| `repr(l)`, `tuple(l)` | `l.ToString`, `l.ArrayType = sgTuple` |

- **Métodos que alteram a lista** (`Append`, `Remove`, `Extend`, `Insert`, `RemoveValue`, `Clear`, `Sort`, `Reverse`) também a devolvem. Funciona tanto `l.Append x` quanto o encadeamento do `Sage.xlam`: `.Split(vbNewLine).Remove(-1).Join(vbNewLine)`.
- **Iteráveis** aceitos por `l = ...`, `Extend` e `Concat`: array (2D: cada linha vira uma `ListS`), outro `ListS`, `DictionaryS` (as chaves), `Collection`, `Range` e texto (os caracteres).
- **Comparação** (`Contains`, `Index`, `CountOf`, `RemoveValue`): números pelo valor (`21 = 21#`), texto diferenciando maiúsculas, objetos pela identidade.
- **Ordenação** estável, como no Python: números pelo valor, texto pela ordem dos caracteres (maiúsculas antes das minúsculas); texto e número misturados geram erro 13.
- **Tupla:** com `ArrayType = sgTuple`, qualquer alteração gera erro 13; `ToString` mostra `(1, 2)`.
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

**Carregar:** `ReadCsv(Caminho, [Cabeçalho], [Delimitador], [SeparadorDecimal], [Codificação])`, `ReadParquet`, `ReadExcel(Caminho, [Planilha], [Cabeçalho])`, `FromRange`, `FromArray` (2D ou array de linhas), `FromRecords` (lista de `DictionaryS`), `Sql("SELECT ... FROM {0} JOIN {1} ...", outroDf)`. Os `Read` aceitam caminho relativo (`ReadCsv "vendas.csv"`, `"dados\vendas.csv"`), a partir da pasta da pasta de trabalho ativa (se ela não foi salva, da pasta atual do Excel).

**Gravar e ver:** `ToCsv`, `ToParquet`, `ToExcel` (formato pela extensão), `ToSql` (banco de dados, pelo `SqlEngine`), `ToRange`, `ToArray`, `ToString` / `Debug.Print df`, `Show` (mostra o DataFrame na aba *Resultado do DataFrame* da janela Terminal do editor, numa grade lida aos poucos: abre na hora mesmo com milhões de linhas, e a macro continua; sem o SageEditor, abre uma janela própria).

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

## DatetimeS

```vb
Dim d As New Sage.DatetimeS
d = Now                                          ' ou Date, "2024-01-05T10:30:00Z", "15/03/2024"
Debug.Print d.StrFTime("%d/%m/%Y %H:%M")         ' códigos do strftime do Python
Debug.Print d.AddMonths(1).IsoFormat             ' 31/01 + 1 mês = 29/02
Debug.Print d.Add(Days:=7, Hours:=-2).ToString
Debug.Print d.UtcNow.AsTimeZone("E. South America Standard Time").IsoFormat   ' ...-03:00
```

Como no Python, uma data não muda: os métodos devolvem um `DatetimeS` novo (`Set d = d.AddMonths(1)`). O membro padrão (`Value`) é um `Date` do VBA, então `d > DateSerial(2024, 1, 1)`, `Format(d, ...)` e `Range("A1") = d.Value` funcionam.

| Python | DatetimeS |
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
| `pd.offsets.MonthEnd` / `Period.end_time` | `d.EndOf(sgMonth)` (último instante do período: 2024-08-31 23:59:59.999999) |
| Excel: `WORKDAY(d, n, feriados)`; numpy: `busday_offset` | `d.AddBusinessDays(n, Feriados)` (pula sábado, domingo e os feriados; `n` negativo volta) |
| numpy: `busday_count(d, outra, holidays=...)` | `d.BusinessDays(outra, Feriados)` (dias úteis em [d, outra); negativo se outra vier antes) |
| SQL Server: `DATETRUNC(month, d)`; DuckDB: `date_trunc('month', d)` | `d.Trunc(sgMonth)` (`sgYear`, `sgQuarter`, `sgMonth`, `sgWeek`, `sgDay`, `sgHour`, `sgMinute`, `sgSecond`, `sgMillisecond`, `sgMicrosecond`; a semana começa na segunda) |

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

`Dumps` aceita também `Scripting.Dictionary`, arrays (2D: lista de linhas), `Collection`, `Date` e `DatetimeS` (texto ISO 8601) e `DataFrame` (lista de linhas). Erros no texto geram `JSONDecodeError` com linha e coluna, como no Python. Diferença do Python: `EnsureAscii` é `False` por padrão (acentos ficam como estão); com `EnsureAscii:=True`, saem como `\u00e7`.

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

## Funções globais: GetUser, Sleep, OfficeTheme e Resizable

Como o `Json` e o `Requests`, usadas sem `Dim` (ou com o prefixo `Sage.`):

| Python | Sage |
|---|---|
| `getpass.getuser()` | `GetUser()`: o usuário logado no Windows, pela API do sistema (não lê a variável `USERNAME`). `GetUser(True)` devolve também o domínio: `EMPRESA\rodrigo` (fora de domínio, o nome do computador). Devolve `StringS` |
| `time.sleep(1.5)` | `Sleep 1.5`: espera em segundos, com frações. O Excel continua redesenhando a tela e não aparece como "Não respondendo"; cliques e teclas esperam o fim da macro, como em qualquer macro ocupada. Valor negativo: `ValueError` (erro 5) |
| — | `OfficeTheme()`: o tema do Office em uso (*Arquivo > Conta > Tema do Office*), em nomes fixos em inglês, para comparar no código sem depender do idioma: `"Colorful"` (Colorido), `"Dark Gray"` (Cinza-Escuro), `"Black"` (Preto) ou `"White"` (Branco). Com *Usar configuração do sistema*, devolve o tema que o Office aplica: `"White"` com o Windows no modo claro e `"Black"` no escuro. Devolve `StringS` |

Um `Declare ... Sleep Lib "kernel32"` no próprio módulo tem prioridade sobre o `Sleep` do Sage.

### Resizable: UserForm redimensionável

```vb
Private Sub UserForm_Initialize()
    Resizable Me, Anchors:=Array( _
        "gridDados", "Fill", _          ' cresce para os dois lados
        "lblInfo", "Width", _           ' acompanha a largura
        "btnFechar", "Bottom Right")    ' fica no canto de baixo à direita
End Sub
```

O form ganha borda para redimensionar e os botões de maximizar e minimizar; os controles ancorados acompanham o tamanho, a partir das posições do designer:

| Âncora | |
|---|---|
| `Left`, `Top` | Ficam onde estão (o padrão) |
| `Right`, `Bottom` | Acompanham a borda direita / de baixo (o controle se move) |
| `Width`, `Height` | Esticam junto com o form |
| `Fill` | `Width Height` |

Separe as âncoras com espaço (`"Bottom Right"`). `Anchors` aceita `Array("controle", "âncora", ...)` ou um `DictionaryS`; controle que não existe no form: `KeyError` (erro 9). `MinWidth` e `MinHeight` (em pontos, como `Width`/`Height` do form) limitam o tamanho mínimo, que por padrão é o do designer; `MaxButton:=False` e `MinButton:=False` tiram os botões. Funciona com qualquer controle (do MSForms ou do Sage); o `UserForm_Resize` continua disparando. A janela é a do próprio form (não procurada pelo título); tudo é solto quando o form fecha.

## TextBox (controle para UserForms)

Caixa de texto com correção ortográfica, para colocar em UserForms: no designer, o **Sage TextBox** já aparece na página **Sage** da Caixa de ferramentas (o SageShortcuts a cria uma vez por usuário, com o Excel fechado). Se tiver sido desmarcado: clique direito na página > *Controles adicionais* > marque **Sage Framework TextBox**. As palavras erradas ficam sublinhadas em vermelho enquanto se digita; o clique direito mostra as sugestões, *Ignorar* e *Adicionar ao dicionário*. A correção é a do Windows (a do Edge e do Bloco de Notas): os idiomas são os instalados em *Configurações > Hora e idioma*, e as palavras adicionadas valem para todos os programas.

```vb
Private Sub txtObs_Change()
    lblTamanho.Caption = Len(txtObs.Value)
End Sub

txtObs.Language = "en-US"                    ' padrão: o idioma do Windows
Debug.Print txtObs.Misspellings.Join(", ")   ' ListS com as palavras erradas
```

| Membro | |
|---|---|
| `Value` (padrão), `Text` | O texto |
| `SpellCheck`, `Language` | Liga a correção (padrão: ligada) e o idioma (`"pt-BR"`; vazio = o do Windows). Idioma não instalado: `ValueError` (erro 5) |
| `Misspellings()` | `ListS` com as palavras erradas, sem repetir |
| `Placeholder` | Texto de dica, em cinza, enquanto a caixa está vazia |
| `MultiLine`, `WordWrap`, `ScrollBars`, `MaxLength`, `PasswordChar`, `Locked`, `Enabled`, `TextAlign` | Como no TextBox do MSForms (`sgScrollBars`, `sgTextAlign`). Com `PasswordChar`, sem correção |
| `BackColor`, `ForeColor`, `BorderColor`, `BorderStyle` | Cores e borda (`sgBorderStyleSingle`, fina; com o foco, azul) |
| `FontName`, `FontSize`, `FontBold`, `FontItalic` | A fonte (padrão: Segoe UI 9) |
| `SelStart`, `SelLength`, `SelText`, `TextLength`, `Copy`, `Cut`, `Paste`, `SelectAll` | Seleção e área de transferência |
| Eventos | `Change`, `KeyDown(KeyCode, Shift)`, `KeyUp`, `KeyPress(KeyAscii)` (zerar o código cancela a tecla), `DblClick`; `Enter`, `Exit`, `BeforeUpdate`, `AfterUpdate` vêm do UserForm |

`Left`, `Top`, `Width`, `Height`, `Visible`, `TabIndex`, `TabStop`, `Tag`, `ControlTipText` e `SetFocus` também vêm do UserForm. Tab, Shift+Tab, Enter (numa linha só) e Esc passam ao form, como nos controles do MSForms; Ctrl+A seleciona tudo.

**Configuração do Office:** o Microsoft 365 passou a vir com *Desabilitar todos os controles sem notificação* (Arquivo > Opções > Central de Confiabilidade > Configurações > Configurações do ActiveX), que bloqueia qualquer controle ActiveX que não seja da Microsoft, inclusive este. Com a opção *Avisar-me antes de habilitar controles UFI com restrições adicionais e controles SFI com restrições mínimas* (o padrão antigo), o Sage TextBox funciona sem aviso: ele é marcado como seguro para inicialização e para scripts. A pasta de trabalho com o controle só abre o form em computadores com o Sage instalado.

## DataGrid (controle para UserForms)

Mostra um `DataFrame` num UserForm (*Sage Framework DataGrid* na Caixa de ferramentas). Só as linhas à vista são lidas do DuckDB, então funciona com milhões de linhas.

```vb
Set gridVendas.DataFrame = df
gridVendas.ColumnFormats = Array("valor", """R$"" #,##0.00", "percentual", "0.0%", "data_cadastro", "dd/mm/yyyy")

Private Sub gridVendas_DblClick()
    MsgBox gridVendas.Cell(gridVendas.SelectedRow, "Cliente")
End Sub
```

**Paginação automática** (`AutoPaging`, padrão `True`): a grade calcula quantas linhas cabem na altura do controle. Se todas couberem, mostra todas; senão, mostra uma página por vez, com a barra *linhas 1–20 de 1.234  ◀ Página 1 de 62 ▶* embaixo, à direita, e refaz a conta quando o controle muda de tamanho. Com `AutoPaging = False`, a grade rola.

| Membro | |
|---|---|
| `DataFrame` | O DataFrame mostrado (`Set`); `Nothing` limpa |
| `AutoPaging` | Paginação automática (padrão) ou barra de rolagem |
| `Page`, `PageCount`, `RowsPerPage`, `NextPage`, `PreviousPage` | Página atual (a partir de 1), quantas há, quantas linhas cabem |
| `SelectedRow`, `SelectedColumn`, `RowCount`, `Cell(Linha, Coluna)` | Linha e coluna da célula ativa (linha a partir de 0, como o índice; `-1` e `""` sem seleção; atribuir `SelectedRow` seleciona a linha inteira), total de linhas e o valor de uma célula (coluna pelo nome ou pela posição; índices negativos contam do fim) |
| `Copy` | Copia a seleção, como o Ctrl+C |
| `ColumnFormats` | Formato de colunas, com os formatos da função `Format` do VBA (o mesmo motor, com as configurações regionais do Windows): um `DictionaryS` ou `Array("coluna", "formato", ...)`, como o `DictionaryS` recebe. Os nomes têm que ser colunas do DataFrame (senão, `KeyError`, erro 9). Vale para a grade e o Ctrl+C; `Cell` devolve o valor original. Fica só no código, não no form |
| `ShowIndex`, `FontName`, `FontSize`, `BackColor`, `ForeColor`, `BorderColor`, `HeaderBackColor`, `SelectionColor` | Aparência (`SelectionColor`: a cor da seleção, padrão `RGB(204, 228, 247)`; sem o foco, a mesma cor acinzentada) |
| Eventos | `Click`, `DblClick`, `SelChange`, `PageChange` |

**Seleção**: clique numa célula (a célula ativa fica com contorno azul); arraste ou Shift+clique para um retângulo (o Shift+clique vale entre páginas); clique no cabeçalho para a coluna inteira e no índice para a linha inteira (com Shift, estende). Teclado: setas movem a célula ativa e Shift+setas estendem; Ctrl+setas vão à borda; Home/End, à primeira/última coluna; Ctrl+Home/End, à primeira/última célula; Page Up/Down, uma página; Ctrl+A, tudo. **Ctrl+C** copia a seleção separada por tabulações (cola em planilhas), com os `ColumnFormats`; NULL vira vazio; seleções grandes (uma coluna inteira) são lidas do DuckDB em blocos. Tab, Enter e Esc passam ao form.

Valores nulos aparecem como **Vazio** (português) ou **Empty** (outros idiomas), em roxo; números alinhados à direita.

## DateBox (controle para UserForms)

Caixa de data com calendário (*Sage Framework DateBox* na Caixa de ferramentas). `Value` é um `Date` de verdade, ou `Empty` sem data.

```vb
dtEntrega.Value = Date
dtEntrega.MinDate = Date                 ' Empty: sem limite
dtEntrega.Format = "dd/mmm/yyyy"         ' fora do foco; vazio = data abreviada do Windows

Private Sub dtEntrega_Change()           ' só quando a data muda (não a cada tecla)
    lblPrazo.Caption = dtEntrega.Value - Date & " dias"
End Sub
```

**Digitação:** só números e separadores; o separador entra sozinho depois do dia e do mês. Conferida ao sair da caixa ou no Enter: `12/10/2026`, `12-10-26`, `12102026`, `121026`, `1503` (ano atual), `5` (mês e ano atuais), na ordem da data abreviada do Windows. Inválida ou fora dos limites: bipe e volta a data anterior. Com o foco, a data aparece como na digitação (`dd/mm/aaaa`); sem ele, no `Format` (o motor do `Format` do VBA). Seta para cima/baixo: um dia.

**Calendário:** o botão à direita, F4 ou Alt+Seta para baixo. É uma janela própria (`CalendarPopup`), desenhada à mão, que não ativa (`WS_EX_NOACTIVATE`, `MA_NOACTIVATE`) e fica com a captura do mouse enquanto aberta, como a lista de uma ComboBox: o foco continua na caixa, as teclas de navegação vão para o calendário (setas, Page Up/Down, Ctrl+Page Up/Down, Home/End, Enter, Esc), e o clique fora o fecha. Vistas de dias, meses e anos (clique no título); *Hoje* e *Limpar* no rodapé; primeiro dia da semana e nomes da cultura do Windows; dias fora de `MinDate`/`MaxDate` apagados.

| Membro | |
|---|---|
| `Value` (padrão) | `Date` ou `Empty`; aceita texto de data. Fora dos limites: `ValueError` (erro 380); texto que não é data: `TypeError` (erro 13) |
| `Text` | O texto da caixa; atribuir funciona como digitar |
| `Format`, `MinDate`, `MaxDate`, `Placeholder` | Formato fora do foco, limites (`Empty` = sem limite) e texto de dica |
| `DropDown` | Abre o calendário |
| `Locked`, `Enabled`, `TextAlign`, `BackColor`, `ForeColor`, `BorderColor`, `BorderStyle`, `FontName`, `FontSize`, `FontBold`, `FontItalic` | Como no TextBox |
| Eventos | `Change`, `DropButtonClick`, `KeyDown(KeyCode, Shift)`, `KeyPress(KeyAscii)` |

As datas são gravadas no form como texto ISO (`yyyy-MM-dd`).

### Como é feito

A classe (`TextBox.cs`; o `DataGrid.cs` e o `DateBox.cs` seguem o mesmo molde) deriva de `OleControl` (`OleControl.cs`), que implementa à mão o protocolo ActiveX (`IOleObject`, `IOleInPlaceObject`, `IOleInPlaceActiveObject`, `IOleControl`, `IViewObject2`, `IPersistStreamInit`, `IObjectSafety`), como o ATL do C++. O `Control` do WinForms, que traz esse suporte pronto, derruba o `FM20.DLL` ao ser inserido. Quando o form ativa o controle, ele cria uma janela filha (`OleHost`) com uma caixa `EDIT` do Windows dentro; sem janela (no designer), desenha pelo `IViewObject::Draw`. A correção ortográfica é a API `ISpellChecker` do Windows (`Spelling.cs`), com o sublinhado desenhado por cima da caixa.

Pontos que fazem a diferença no UserForm:
- O stream recebido em `IPersistStreamInit::Save`/`Load` é liberado na hora (`Marshal.FinalReleaseComObject`). O .NET o manteria aberto até o coletor de lixo, e o form, ao rodar, abre o mesmo stream (`o`) em modo exclusivo: daria *Erro de acesso a caminho/arquivo*.
- `IMarshal` e `IAgileObject` são recusados (`ICustomQueryInterface`): o UserForm embrulha o controle (extender) e repassaria a pergunta ao .NET, que responderia com o empacotador livre de threads; pela automação (outro processo) chegaria o controle sem o embrulho, e o `FM20.DLL` cai.
- O tamanho do controle vai junto com as propriedades gravadas (o form o pede com `GetExtent` depois de carregar).
- As propriedades são gravadas com nome e tipo (`SaveProperties`/`LoadProperties`); as desconhecidas são ignoradas ao ler, então versões novas podem acrescentar propriedades. Não mude os nomes.
- O `install.ps1` marca a classe com `[control]` no `.tlb` e tira a interface `_Object` do .NET; o registro leva as chaves de controle (`Control`, `MiscStatus`, `TypeLib`, `Version`, `ToolboxBitmap32`, categorias de controle, seguro para scripts e para inicialização, `PersistsToStreamInit`). No `com.txt`, a linha do controle termina em `|control`.
- A Caixa de ferramentas (`Microsoft\Forms\EXCEL.box`, no perfil do usuário) recebe os controles numa página **Sage** própria, pelo SageShortcuts (`ToolboxSetup`), uma vez por usuário (os já postos ficam em `%APPDATA%\Sage\toolbox.txt`), e só com o Excel fechado: o VBE mantém o arquivo aberto e o regrava ao fechar. O arquivo é um form do MSForms com um MultiPage; a página é criada como o VBE cria a dele (*Nova página*): um site no `f`, uma aba no `o` (TabStrip), o ID no `x` e uma pasta `iNN` com a lista de controles. O código fica em `SageShortcuts\src\ToolboxFile.cs`, usado também pelo desinstalador. Arquivo fora do formato conferido: os controles vão para a primeira página, como antes.
- O ícone da Caixa de ferramentas é um BITMAP do `SageToolbox.dll` (1 TextBox, 2 DataGrid, 3 DateBox; `res\toolbox.res`, 16x16, compilado pelo `build.cmd`).
- O MSForms guarda um cache (`.exd`) da biblioteca de tipos de cada controle, em `%TEMP%\VBE` e em `Microsoft\Forms` no perfil (`%APPDATA%` ou a pasta `Application Data` do perfil, que pode ser uma pasta de verdade em vez do atalho); o instalador apaga os do Sage, para não ficar um cache de uma versão anterior.

## SqlEngine

Conexão a um banco de dados, como o `Engine` do SQLAlchemy que o pandas usa em `pd.read_sql(consulta, engine)` e `df.to_sql("tabela", engine)`. As consultas vão no SQL do próprio banco; os parâmetros, com `?`, num `Array`.

```vb
Dim db As New Sage.SqlEngine
db.Connect "mssql://servidor/Vendas"

Dim df As Sage.DataFrame
Set df = db.ReadSql("SELECT * FROM Pedidos WHERE Ano = ? AND Status = ?", Array(2024, "Aberto"))
Debug.Print db.Execute("UPDATE Pedidos SET Status = 'OK' WHERE Id = ?", Array(10))   ' linhas afetadas
df.ToSql "PedidosCopia", db, "replace"
db.Close
```

| Banco | `Connect` | Como acessa |
|---|---|---|
| PostgreSQL | `postgresql://usuário:senha@servidor:5432/banco` ou `host=... dbname=...` | DuckDB (extensão `postgres`) |
| MySQL / MariaDB | `mysql://usuário:senha@servidor:3306/banco` | DuckDB (extensão `mysql`) |
| SQLite | `C:\dados\base.db` (`.sqlite`, `.sqlite3`) ou `sqlite:C:\dados\base.db` | DuckDB (extensão `sqlite`) |
| Arquivo do DuckDB | `C:\dados\base.duckdb` | DuckDB |
| SQL Server | `mssql://servidor/banco` (usuário do Windows), `mssql://usuário:senha@servidor/banco` ou `Server=...;Database=...` | cliente do .NET (sem driver ODBC) |
| Access | `C:\dados\base.accdb` (`.mdb`) | OLE DB do Office (ACE) |
| Outros | `Driver={...};...` ou `DSN=...` (ODBC), `Provider=...;...` (OLE DB) | drivers do Windows |

| pandas / SQLAlchemy | SqlEngine |
|---|---|
| `engine = create_engine(url)` | `db.Connect url` |
| `pd.read_sql(sql, engine, params=[...])` | `db.ReadSql(sql, Array(...))` (ou só o nome da tabela) |
| `conn.execute(sql, params)` | `db.Execute(sql, Array(...))` (devolve as linhas afetadas; -1 quando o banco não informa) |
| `df.to_sql("t", engine, if_exists="replace")` | `df.ToSql "t", db, "replace"` (`"fail"`, o padrão; `"replace"`; `"append"`) |
| `inspect(engine).get_table_names()` | `db.Tables` |
| `engine.dispose()` | `db.Close` |

- **PostgreSQL, MySQL, SQLite e DuckDB** passam pelo mesmo DuckDB do DataFrame: os dados chegam direto no formato dele, sem passar linha a linha pelo VBA. As extensões vêm no pacote (não precisam de internet).
- **SQL Server** grava com *bulk copy* (20 mil linhas em menos de meio segundo no LocalDB) e lê pelo `SqlClient` do .NET, que já vem no Windows.
- **Access e ODBC** gravam linha a linha numa transação; para tabelas grandes, prefira o SQL Server ou o PostgreSQL.
- **Tipos:** inteiros viram `BIGINT`, decimais `DOUBLE`, datas `DATE`/`TIMESTAMP`, sim/não `BOOLEAN`, o resto texto. Ao gravar (`ToSql`), a tabela é criada com os tipos equivalentes do banco (no SQL Server, `BIGINT`, `FLOAT`, `DATE`, `DATETIME2`, `BIT` e `NVARCHAR(MAX)`; no Access, `LONG`, `DOUBLE`, `DATETIME`, `YESNO` e `TEXT`/`LONGTEXT`).
- **Erros** chegam como `DatabaseError` ou `ConnectionError` (erro 5), com a mensagem do banco. A senha nunca aparece na mensagem: ela é trocada por `***`.

## Diferenças em relação ao StringS do Sage.xlam

- `FString` e `Join` aceitam até 30 argumentos, e não uma `ParamArray` ilimitada: o VBA recusa a `ParamArray` exportada pelo .NET.
- `Count("an")` conta ocorrências (2 em "banana"); a versão VBA contava caracteres removidos (4).
- `Mid(3)` sem tamanho vai até o fim; na versão VBA o padrão `"End"` dava erro de tipo.
- `Proper` segue as mesmas regras do PROPER do Excel, sem chamar o Excel.

Mantidos de propósito, como no original: ao atribuir um valor, `\n` vira quebra de linha, `\t` vira espaços até a próxima coluna múltipla de 4 e o texto perde espaços e quebras das pontas (por isso `", "` vira `","`).

## Para acrescentar membros

Os nomes (classes, membros, parâmetros) seguem o PascalCase, sem sublinhado: `DatetimeS`, `ReadCsv`, `SortValues`. Os enums e seus membros seguem o camelCase com o prefixo `sg`, como as constantes do próprio VBA (`vbCrLf`, `xlUp`): `sgDateParts` (`sgMonth`...), `sgArrayTypes` (`sgList`, `sgTuple`).

As interfaces (`_StringS`, `_DictionaryS`, `_ListS`, `_DataFrame`, `_DatetimeS`, `_Json`, `_Requests`, `_Session`, `_Response`, `_Globals`, `_SqlEngine`) são duais: o VBA as chama pela vtable. Acrescente membros **sempre no fim**, com o próximo `DispId`, e nunca reordene nem remova os existentes, senão o código VBA já compilado chama o método errado. Uma classe nova precisa de `[Guid]`, `[ProgId]`, interface própria e uma linha em `$Classes` no `install.ps1`.

Limitações do exportador do .NET (`TypeLibConverter`) e como contorná-las:
- **Propriedade Variant com `Let`:** o .NET exporta o setter só como `Property Set`. Declare também um método `LetNome(...)` com `[PropertyLet("Nome")]`, os mesmos parâmetros e mais o valor; o `install.ps1` o transforma no `Property Let` de `Nome` ao gerar o `.tlb`.
- **Propriedade com parâmetros** (como `At(linha, coluna)`): o .NET só exporta o indexador. Declare um método `GetNome(...)` com `[PropertyGet("Nome")]`; no VBA ele é lido como `Nome(...)`.
- **`ParamArray`:** o `params` do C# não é aceito pelo VBA; use parâmetros `[Optional]`.
- **Membros de enum:** o .NET os exporta como `Enum_Membro` (`sgArrayTypes_sgTuple`); o `install.ps1` tira o prefixo, e o VBA vê `sgTuple`.
- **Membros globais** (`Json`, `Requests`, `GetUser`, `Sleep`): ficam na classe `Globals`, com `[AppObject]`; o `install.ps1` a marca como *app object* no `.tlb`, e o VBA a cria sozinho.
- **Nome igual a algo da biblioteca VBA** (como `Strings`): o VBA acha o dele primeiro, por isso o prefixo `Sage.`.

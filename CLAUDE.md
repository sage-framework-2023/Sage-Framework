# Contexto: refatoração das classes do add-in Sage (VBA)

Documento para continuar o trabalho no Claude Code (app desktop), em uma máquina Windows com Excel.

## 1. Projeto

- Projeto VBA **"Sage"**, com cerca de 23 mil linhas e 59 componentes. As classes são `Exposed` e ficam disponíveis a outros add-ins via `PublicClasses.New_*`.
- Arquivo analisado na conversa: `Sage.xlam`.
- Arquivo de desenvolvimento na máquina do usuário: `C:\Users\rodrigo.bechara.CSG\Documents\Programas\Sage Excel\Sage.xlsm`.
  - **Pode estar mais atualizado que o `.xlam`.** Não sobrescreva o `General` sem comparar antes.
- O projeto VBA é protegido por senha.
  - O **usuário** desbloqueia manualmente no VBE (Alt+F11), na mesma sessão do Excel, antes da automação.
  - O Claude não deve digitar nem usar a senha.
- Codepage do projeto: 1252. Os fontes usam CRLF.
- Pré-requisito: *Arquivo > Opções > Central de Confiabilidade > Configurações de Macro > "Confiar no acesso ao modelo de objeto do projeto do VBA"*.

## 2. O que foi feito na conversa

1. **Auditoria do que está inacabado.** Pendências fora do escopo atual:
   - `DataFrame.WriteTable` usa `Self.ColumnType` sem índice.
   - `DataFrame.GroupBy` chama `LoadArray` sem o ponto.
   - `Network.MapNetworkDrive` nunca cria o `fso`.
   - `Google` e `APIClient` estão inacabados.
   - Há três implementações de atalhos de teclado.
   - `uKeyboardShortcuts` mistura eventos `df_` com controles `tb_`.
   - `uTvwBox` está 100% comentado.
2. **Revisão de bugs** das classes StringS, ListS, DatetimeS e DictionaryS (resumo na seção 4).
3. **Reescrita** das classes StringS, ListS, DatetimeS, DictionaryS e JsonParser, com edição de trechos do `General.bas`.
4. **Gravação direta no `vbaProject.bin` do `Sage.xlam`:**
   - Todos os módulos foram purgados de p-code (`MODULEOFFSET = 0`).
   - O `_VBA_PROJECT` foi reduzido ao mínimo (`CC 61 FF FF 00 00 00`), o que força a recompilação.
   - Os streams `__SRP_*` foram removidos.
   - Os 53 módulos não alterados mantêm o fonte idêntico byte a byte.
5. **Nada foi compilado nem executado no Excel.** A validação foi feita só com um lint estático e com a releitura do arquivo por olefile/olevba.

## 3. Arquivos entregues

| Arquivo | Conteúdo |
|---|---|
| `Sage.xlam` | Add-in com o código novo (sem p-code; o Excel recompila ao abrir) |
| `modulos/StringS.cls`, `ListS.cls`, `DatetimeS.cls`, `DictionaryS.cls`, `JsonParser.cls` | Classes reescritas por completo (cp1252 + CRLF, com cabeçalho `VERSION/BEGIN` para importação) |
| `modulos/General.bas` | General do `.xlam` com as edições aplicadas |
| `modulos/TestSageTypes.bas` | Cerca de 150 verificações; `RunAllTests` escreve o resultado na Janela Imediata |

Cabeçalhos preservados: `StringS` tem `VB_PredeclaredId = True`; as demais classes têm `False`. Todas têm `VB_Exposed = True`.

## 4. Mudanças por componente

### StringS (imutável; cada método devolve uma nova instância)
- `Value` guarda o texto como recebido, sem converter `\n`/`\t`. Para isso agora existe `.Unescape`.
- `Search(texto, [sgStart|sgEnd], [Compare=vbTextCompare])`:
  - Usa `InStr`/`InStrRev` em base 0 e devolve **-1** quando não encontra.
  - Não usa mais curingas.
  - Com `sgEnd`, devolve o início da última ocorrência.
- `IsDigit`: só aceita 0-9. `IsNumeric`: usa `VBA.IsNumeric`. **Os dois trocaram de comportamento** em relação à versão anterior, e texto vazio agora retorna False.
- `Trim([texto])`: remove espaços, tabs, quebras de linha e NBSP, além de repetições do texto indicado.
- Métodos novos: `Strip`, `LStrip` e `RStrip` (conjunto de caracteres, como no Python).
- `Count` corrigido para substrings com mais de um caractere.
- `Replace` aceita arrays ou ListS (par a par), com `Compare` opcional.
- `Left`/`Right`: valores negativos removem caracteres; valor não numérico gera erro 13.
- `Mid(Start, [Length])` continua em base 1.
- `FString` mantém a concatenação. O novo `Format` substitui `{0}`, `{1}`… e aceita `{{ }}` para chaves literais.
- `Join` aceita ParamArray, array ou ListS; o separador é o próprio texto.
- Outros métodos novos: `Lines`, `IndexOf`, `Contains`, `StartsWith`, `EndsWith`, `Equals`, `PadLeft`, `PadRight`, `Repeat` e `ExpandTabs`.

### ListS
- Armazenamento em `InItems()` com capacidade que dobra; `Add` é O(1) amortizado.
- `Value` sem índice **sempre** devolve array. Isso corrige o "colapso" de listas com 1 elemento.
- `Value(i)` devolve texto como StringS, array como ListS e o resto como está. `Item(i)` devolve o valor bruto.
- Índices negativos funcionam também com objetos. Índice inexistente gera erro 9, inclusive em `Remove`.
- Métodos que retornam uma nova lista: `Append`, `Remove`, `Slice`, `Reverse`, `Sorted` e `Copy`.
- Métodos que alteram a própria lista: `Add`, `Insert`, `RemoveAt`, `Pop`, `Extend`, `Clear` e `Sort`.
- Consultas: `IndexOf`, `Contains`, `CountOf`, `Count`/`Length`, `ToArray` e `ToCollection`.
- Suporta `For Each` (`NewEnum`).
- Carga aceita array 1D, array 2D (lido linha a linha), ListS, Collection e escalar (vira lista de 1 elemento).
- Tupla (`ArrayType = sgTuple`) continua somente leitura.

### DictionaryS
- Implementada sobre `Scripting.Dictionary`: operações O(1) e sem a dependência circular com o DataFrame.
- `CompareMode` só pode ser alterado com o dicionário vazio.
- Chaves mantêm o tipo:
  - `1` e `"1"` são chaves diferentes.
  - Inteiros viram Long e os demais números viram Double.
  - StringS vira String e DatetimeS vira texto ISO.
- `Value(k)` devolve texto como StringS, data como DatetimeS e array como ListS; chave inexistente devolve Empty. `Item(k)` devolve o valor bruto e também aceita Let/Set.
- `Value` sem chave:
  - Na leitura, devolve uma cópia em `Scripting.Dictionary`.
  - Na gravação, carrega pares (array, ListS, array 2D de 2 colunas, Scripting.Dictionary ou DictionaryS).
- `GetItem(k, padrão)` não tenta mais adivinhar datas.
- `Append` não sobrescreve chaves existentes; o novo `Update` sobrescreve.
- Outros métodos: `Pop`, `Remove` (sem erro se a chave não existe), `Keys`, `Items` (`Itens` mantido como alias), `Copy`, `Clear`, `Key` e `For Each` sobre as chaves.
- Para copiar objetos, use `Set d.Value = outro`.

### DatetimeS
- Armazenamento:
  - Data e hora em segundos inteiros.
  - Microssegundos em um Long.
  - Deslocamento UTC em minutos.
- O membro padrão `Value` devolve um **Date nativo**. Para o texto use:
  - `ToIso([Sep="T"], [Digits=6], [IncludeOffset])`.
  - `ToString`, que mantém o formato antigo `yyyy-mm-dd hh:mm:ss.ffffff+hh:mm`.
- Parse ISO: aceita T, fração, `Z`, `±hh:mm`, `±hhmm` e `±hh`; outros formatos caem no `CDate`. Offsets negativos agora são lidos corretamente.
- `Timestamp` segue o padrão Unix, em UTC.
- Setters de ano, mês, dia, hora, minuto e segundo validam os valores e ajustam o fim do mês.
- `TimeZone` (Let) só troca o rótulo do fuso. Para converter o horário use `ToTimeZone`, `ToUTC` ou `ToLocal`.
- `DateAdd`: anos e meses precisam ser inteiros; as demais partes são somadas em microssegundos, com virada correta de dia.
- `DateDiff(outro, [parte=sgSecond])` agora **devolve Double** e considera os fusos.
- `StrFTime` trata `f…` e `z` como literais e **não aplica mais o deslocamento de fuso**.
- Métodos novos: `IsoWeek`, `CompareTo`, `Copy` e `OffsetMinutes`.
- `Now` usa `Date`/`Timer` e não depende mais do separador decimal do Windows.
- Membros `Friend`: `SetRaw`, `AddMicroseconds` e `UtcSeconds`.

### JsonParser
- Parser que percorre o texto uma única vez, por índice:
  - objeto → DictionaryS; array → ListS; texto → String; número → Long ou Double (via `Val`, independente do locale); `true`/`false`/`null`.
- Suporta escapes `\" \\ \/ \b \f \n \r \t \uXXXX` e ignora BOM.
- Erros informam linha, coluna e posição.
- Métodos novos: `Parse`, `ToJson([indent])` e `Stringify(valor, [indent])`.

### General.bas
- Novo `GetTimeZoneMinutes()`, usando a API `GetTimeZoneInformation` (com declaração `#If VBA7` e considerando horário de verão).
- `GetTimeZone()` devolve `+hh:mm`. A variável `strTimeZone` e o uso de WMI foram removidos.
- `TypeName`: para objetos usa `VBA.TypeName`; qualquer array devolve `"Array"`.
- `ToDatetime`:
  - "Now" sem diferenciar maiúsculas.
  - DatetimeS recebido é copiado.
- `ToString`, `ToList` e `ToDictionary`: devolvem a mesma instância quando recebem o próprio tipo; outros objetos entram via `Set`.

## 5. Mudanças incompatíveis (revisar nos add-ins que usam o Sage)

1. DatetimeS:
   - `Value` devolve Date.
   - `DateDiff` devolve Double.
   - `Timestamp` é UTC.
   - `StrFTime` não aplica o deslocamento de fuso.
2. StringS:
   - `IsDigit` e `IsNumeric` trocaram de comportamento.
   - `Search` devolve -1 em vez de gerar erro, e `sgEnd` devolve o início da ocorrência.
   - `Value`, `Replace` e `Trim` não convertem `\n`/`\t`.
3. ListS: `Remove` fora do intervalo gera erro 9.
4. DictionaryS:
   - As chaves mantêm o tipo.
   - `GetItem` não adivinha datas.
   - Copiar objetos exige `Set d.Value = x`.
5. `Sage.TypeName` de um array com 1 elemento devolve `"Array"`.

## 6. Cuidados de VBA adotados no código novo (manter ao corrigir)

- **Evite atribuir sem `Set` um valor simples a uma Variant que contém um objeto.** Isso pode acionar o membro padrão do objeto. O código novo contorna o problema de três formas:
  - Recria arrays em vez de sobrescrever posições (`RebuildWithout`, `SetItemAt`).
  - Usa variáveis locais novas a cada valor lido (`StoreInDictionary`, `StoreInList`).
  - Guarda `InJson` em um array de 1 posição.
- Uma atribuição `x.Value = obj` (Let) passa o **membro padrão** de `obj`, não o objeto. Para objetos, use `Set x.Value = obj`.
- Dentro das classes, sempre qualifique as funções nativas (`VBA.Left$`, `VBA.Year`, `VBA.DateAdd`, `VBA.Replace`…), porque elas colidem com nomes de métodos das próprias classes. Use `Sage.ToString` em vez de `ToString`.

## 7. Próximos passos (no Claude Code)

1. Faça backup do `Sage.xlsm`.
2. Confirme que o projeto está desbloqueado pelo usuário e que o acesso ao modelo de objeto do VBA está habilitado.
3. Substitua as classes StringS, ListS, DatetimeS, DictionaryS e JsonParser pelos arquivos de `modulos/`, usando `VBComponents.Remove` e `.Import`.
4. Com `General`, **não substitua o módulo inteiro**:
   - Exporte o atual e compare com `modulos/General.bas`.
   - Aplique só as mudanças da seção 4 (declarações da API de fuso, `GetTimeZone`, `GetTimeZoneMinutes`, `ToDatetime`, `ToDictionary`, `ToString`, `ToList` e `TypeName`).
   - Remova `strTimeZone`.
5. Importe `TestSageTypes.bas`, compile (*Depurar > Compilar VBAProject*) e rode `RunAllTests`. Corrija até passar tudo.
6. Verifique os usos que dependem do comportamento antigo:
   - `General.FormatDouble` e `EOMonth`.
   - `Terminal.ProgressBar`.
   - `DataFrame`: `Column Get` usa `ToDatetime`, `ToString` e `ToList`; `SQLServer_Record` usa `Dt.StrFTime.Replace`.
   - Os outros add-ins que consomem o Sage.
7. Salve e gere o `.xlam`.
8. Opcional: corrija as pendências da auditoria listadas na seção 2.

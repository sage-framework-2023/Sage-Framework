using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace SageEditor
{
    // Problemas de um módulo VBA encontrados lendo o texto, sem compilar (aba Problemas):
    //   - blocos desencontrados: If/End If, For/Next, Do/Loop, While/Wend, With/End With,
    //     Select Case/End Select, Sub/Function/Property, Type, Enum e #If/#End If;
    //   - Next com a variável de outro For, Case e Else fora do bloco, Exit For/Do fora do laço,
    //     Exit Sub numa Function;
    //   - rótulo de GoTo, GoSub e Resume que não existe no procedimento;
    //   - declaração duplicada (variável, constante, parâmetro) e nome ambíguo (procedimento);
    //   - código fora de procedimento e declarações depois dos procedimentos;
    //   - sintaxe (VbaSyntax.cs): cabeçalho de procedimento, declarações, parênteses, If sem Then,
    //     For sem = ou To, dois valores seguidos numa expressão.
    // As mensagens são as do VBA. Só o que dá para afirmar com certeza pelo texto: o que depende
    // das bibliotecas referenciadas (variável não definida, tipos) fica para o compilador.
    //
    // Sem erros em cascata: um bloco que fica aberto é apontado na linha dele, e o fechamento de
    // um bloco mais de fora (o Next de um For com um With aberto dentro) fecha o que ficou aberto
    // sem acusar o fechamento. Com #If/#Else, só o primeiro trecho é lido (o outro costuma repetir
    // as mesmas declarações para outra versão do Office).
    static partial class VbaLinter
    {
        public sealed class Problem
        {
            public int Line, Column, Length;  // linha e coluna (base 1) e o tamanho do trecho
            public string Message;
            public string Procedure;           // o procedimento (vazio: nas declarações do módulo)
        }

        enum Kind { Procedure, If, For, Do, While, With, Select, Type, Enum }

        sealed class Block
        {
            public Kind Kind;
            public int Line, Column, Length;
            public string Variable;  // For: a variável de controle
            public string ProcKind;  // Procedure: "Sub", "Function" ou "Property"
            public string Name;      // Procedure: o nome
        }

        // Um comando: o texto (sem comentários; o conteúdo das strings trocado por espaços, para
        // as palavras não serem encontradas dentro delas) e onde ele começa
        sealed class Statement
        {
            public string Text;
            public int Line, Column;
            public Func<int, Statement> At; // a posição (no texto da linha lógica) em linha e coluna
            public int Offset;              // onde o comando começa na linha lógica
        }

        static readonly RegexOptions Ci = RegexOptions.IgnoreCase | RegexOptions.CultureInvariant;
        static readonly Regex ProcStart = new Regex(@"^(?:(?:Public|Private|Friend|Global)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_À-ɏ][\wÀ-ɏ]*)", Ci);
        static readonly Regex ProcEnd = new Regex(@"^End\s+(Sub|Function|Property)\b", Ci);
        static readonly Regex TypeStart = new Regex(@"^(?:(?:Public|Private)\s+)?(Type|Enum)\s+[A-Za-z_]", Ci);
        static readonly Regex TypeEnd = new Regex(@"^End\s+(Type|Enum)\b", Ci);
        static readonly Regex ForStart = new Regex(@"^For\s+(?:Each\s+)?([A-Za-z_À-ɏ][\wÀ-ɏ]*)", Ci);
        static readonly Regex Then = new Regex(@"\bThen\b", Ci);
        static readonly Regex Exit = new Regex(@"\bExit\s+(For|Do|Sub|Function|Property)\b", Ci);
        static readonly Regex Jump = new Regex(@"\b(GoTo|GoSub)\s+([A-Za-z_À-ɏ][\wÀ-ɏ]*|\d+)", Ci);
        static readonly Regex OnJump = new Regex(@"\bOn\s+(?!Error\b)(?:.+?)\b(?:GoTo|GoSub)\s+(.+)$", Ci);
        static readonly Regex Resume = new Regex(@"\bResume\s+([A-Za-z_À-ɏ][\wÀ-ɏ]*|\d+)", Ci);
        static readonly Regex Label = new Regex(@"^\s*([A-Za-z_À-ɏ][\wÀ-ɏ]*|\d+):(?!=)", Ci);
        static readonly Regex LineNumber = new Regex(@"^\s*(\d+)\s", Ci);
        static readonly Regex FirstWord = new Regex(@"^\s*([A-Za-z_#À-ɏ][\wÀ-ɏ]*)", Ci);
        static readonly Regex Identifier = new Regex(@"^\s*(?:WithEvents\s+)?([A-Za-z_À-ɏ][\wÀ-ɏ]*)", Ci);
        static readonly Regex DeclareProc = new Regex(@"^(?:(?:Public|Private)\s+)?Declare\s+(?:PtrSafe\s+)?(?:Sub|Function)\s+([A-Za-z_À-ɏ][\wÀ-ɏ]*)", Ci);
        static readonly Regex Variables = new Regex(@"^(?:(Dim|Static|Const|Private|Public|Global)\s+)(?:(Const)\s+)?(?!Sub\b|Function\b|Property\b|Declare\b|Type\b|Enum\b|Event\b)(.+)$", Ci);
        static readonly Regex Directive = new Regex(@"^\s*#(If|ElseIf|Else|End\s*If|Const)\b", Ci);

        // Palavras que não podem ser rótulo ("Else:" é Else seguido de outro comando)
        static readonly HashSet<string> NotLabels = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "Else", "Rem" };

        public static List<Problem> Analyze(string code)
        {
            return new Analysis(code ?? "").Run();
        }

        sealed partial class Analysis
        {
            readonly string[] lines;
            readonly List<Problem> problems = new List<Problem>();
            readonly List<Block> stack = new List<Block>();
            Block procedure;                    // o procedimento aberto
            bool afterProcedures;               // já houve um procedimento (só comentários depois dele)
            string typeBlock;                   // dentro de Type ou Enum: as linhas são membros
            int directiveDepth;                 // dentro de #If
            int skipDepth;                      // dentro de um #Else/#ElseIf: não lido
            readonly List<int> directiveLines = new List<int>();
            readonly List<bool> directiveSkipping = new List<bool>();

            // Nomes do módulo (variáveis e procedimentos) e do procedimento aberto
            readonly Dictionary<string, string> moduleNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            readonly Dictionary<string, bool> localNames = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            readonly Dictionary<string, bool> labels = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            readonly List<Statement> jumps = new List<Statement>(); // GoTo/Resume: o rótulo no Text

            public Analysis(string code)
            {
                lines = code.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            }

            public List<Problem> Run()
            {
                int i = 0;
                while (i < lines.Length)
                {
                    int first = i;
                    // Linha lógica: as continuações (" _" no fim) juntam as linhas físicas; cada
                    // trecho guarda de que linha física veio
                    List<int> segLine = new List<int>(), segStart = new List<int>();
                    StringBuilder logical = new StringBuilder();
                    while (true)
                    {
                        string clean = Clean(lines[i]);
                        segLine.Add(i + 1);
                        segStart.Add(logical.Length);
                        string trimmed = clean.TrimEnd();
                        bool continues = trimmed.EndsWith(" _") || trimmed == "_";
                        logical.Append(continues ? trimmed.Substring(0, trimmed.Length - 1) : clean);
                        i++;
                        if (!continues || i >= lines.Length) break;
                    }
                    Line(logical.ToString(), segLine, segStart);
                }
                Finish();
                problems.Sort(delegate(Problem a, Problem b) { return a.Line != b.Line ? a.Line.CompareTo(b.Line) : a.Column.CompareTo(b.Column); });
                return problems;
            }

            // O texto sem o comentário e com o conteúdo das strings trocado por espaços (as colunas
            // continuam as mesmas)
            static string Clean(string line)
            {
                StringBuilder sb = new StringBuilder(line.Length);
                bool inString = false;
                for (int k = 0; k < line.Length; k++)
                {
                    char c = line[k];
                    if (inString)
                    {
                        if (c == '"')
                        {
                            if (k + 1 < line.Length && line[k + 1] == '"') { sb.Append("  "); k++; continue; }
                            inString = false;
                            sb.Append('"');
                        }
                        else sb.Append(' ');
                        continue;
                    }
                    if (c == '"') { inString = true; sb.Append('"'); continue; }
                    if (c == '\'') break;
                    sb.Append(c);
                }
                // Rem: comentário quando é o começo de um comando
                string s = sb.ToString();
                Match rem = Regex.Match(s, @"(^|:)\s*Rem(\s|$)", Ci);
                if (rem.Success) s = s.Substring(0, rem.Index + rem.Groups[1].Length);
                return s;
            }

            // Uma linha lógica: o rótulo do começo, as diretivas e os comandos (separados por ":")
            void Line(string text, List<int> segLine, List<int> segStart)
            {
                if (text.Trim().Length == 0) return;
                Func<int, Statement> at = delegate(int index)
                {
                    int seg = 0;
                    while (seg + 1 < segStart.Count && segStart[seg + 1] <= index) seg++;
                    return new Statement { Line = segLine[seg], Column = index - segStart[seg] + 1 };
                };

                Match directive = Directive.Match(text);
                if (directive.Success) { DirectiveLine(directive.Groups[1].Value, at(directive.Groups[1].Index - 1)); return; }
                if (skipDepth > 0) return;

                int start = 0;
                Match label = Label.Match(text);
                if (label.Success && !NotLabels.Contains(label.Groups[1].Value))
                {
                    if (procedure != null) labels[label.Groups[1].Value] = true;
                    start = label.Length;
                }
                else
                {
                    Match number = LineNumber.Match(text);
                    if (number.Success)
                    {
                        if (procedure != null) labels[number.Groups[1].Value] = true;
                        start = number.Length;
                    }
                }

                // Os comandos, separados por ":" (fora de strings, que já estão em branco); um If de
                // uma linha só leva o resto da linha com ele
                int k = start;
                while (k < text.Length)
                {
                    while (k < text.Length && char.IsWhiteSpace(text[k])) k++;
                    if (k >= text.Length) break;
                    int end = k;
                    bool ifLine = Regex.IsMatch(text.Substring(k), @"^If\b", Ci);
                    if (ifLine) end = text.Length;
                    else
                        while (end < text.Length && !(text[end] == ':' && (end + 1 >= text.Length || text[end + 1] != '='))) end++;
                    string body = text.Substring(k, end - k).TrimEnd();
                    Statement st = at(k);
                    st.Text = body;
                    st.At = at;
                    st.Offset = k;
                    if (body.Length > 0) Command(st);
                    k = end + 1;
                }
            }

            void DirectiveLine(string word, Statement st)
            {
                string w = word.Replace(" ", "").ToLowerInvariant();
                if (w == "if")
                {
                    directiveDepth++;
                    directiveLines.Add(st.Line);
                    directiveSkipping.Add(false);
                    return;
                }
                if (w == "const") return;
                if (directiveDepth == 0)
                {
                    Add(st, 1 + word.Length, string.Format(Strings.LintWithout, "#" + DirectiveName(w), "#If"));
                    return;
                }
                if (w == "else" || w == "elseif")
                {
                    // Só o primeiro trecho é lido: os outros são pulados até o #End If
                    if (!directiveSkipping[directiveSkipping.Count - 1]) { directiveSkipping[directiveSkipping.Count - 1] = true; skipDepth++; }
                    return;
                }
                // #End If
                if (directiveSkipping[directiveSkipping.Count - 1]) skipDepth--;
                directiveSkipping.RemoveAt(directiveSkipping.Count - 1);
                directiveLines.RemoveAt(directiveLines.Count - 1);
                directiveDepth--;
            }

            static string DirectiveName(string w)
            {
                return w == "endif" ? "End If" : w == "elseif" ? "ElseIf" : w == "else" ? "Else" : w;
            }

            void Command(Statement st)
            {
                string s = st.Text;
                string word = FirstWordOf(s);

                // Type e Enum: as linhas de dentro são membros, até o End Type/End Enum
                if (typeBlock != null)
                {
                    Match te = TypeEnd.Match(s);
                    if (te.Success)
                    {
                        if (!te.Groups[1].Value.Equals(typeBlock, StringComparison.OrdinalIgnoreCase))
                            Add(st, s.Length, string.Format(Strings.LintWithout, "End " + Proper(te.Groups[1].Value), Proper(te.Groups[1].Value)));
                        else Pop(Kind.Type);
                        typeBlock = null;
                    }
                    return;
                }

                List<Token> tokens = Tokens(s);
                bool balanced = Balanced(st, tokens); // com parênteses desencontrados, o resto da sintaxe não é conferido

                Match proc = ProcStart.Match(s);
                if (proc.Success && !Regex.IsMatch(s, @"^(?:(?:Public|Private)\s+)?Declare\b", Ci))
                {
                    StartProcedure(st, proc);
                    if (balanced) Header(st, tokens, proc.Groups[2].Index);
                    return;
                }
                Match pend = ProcEnd.Match(s);
                if (pend.Success) { EndProcedure(st, Proper(pend.Groups[1].Value)); return; }

                if (procedure == null)
                {
                    ModuleLevel(st, word);
                    if (balanced && !afterProcedures) ModuleDeclarations(st, tokens);
                    return;
                }

                // Dentro de um procedimento
                Match ts = TypeStart.Match(s);
                if (ts.Success || DeclareProc.IsMatch(s))
                {
                    Add(st, word.Length, Strings.LintInsideProcedure);
                    if (ts.Success) typeBlock = Proper(ts.Groups[1].Value);
                    return;
                }
                bool declaration = Regex.IsMatch(s, @"^(Dim|Static|Const)\b", Ci);
                if (declaration)
                {
                    LocalDeclarations(st);
                    if (balanced) Declarations(st, tokens, 1);
                }
                Jumps(st);
                if (balanced && !declaration) Syntax(st, tokens);

                switch (word.ToLowerInvariant())
                {
                    case "if":
                        // Bloco If: nada depois do Then (o comentário já saiu)
                        Match then = Then.Match(s);
                        if (!then.Success)
                        {
                            // Sem Then (acusado pelo Syntax): sem GoTo, vale como bloco, para o End If
                            // de depois não virar outro erro
                            if (!Regex.IsMatch(s, @"\bGoTo\b", Ci)) Push(Kind.If, st, 2);
                            return;
                        }
                        if (s.Substring(then.Index + then.Length).Trim().Length == 0) Push(Kind.If, st, 2);
                        else foreach (Match ex in Exit.Matches(s)) ExitCheck(st, ex);
                        return;
                    case "elseif":
                        if (Top() == null || Top().Kind != Kind.If) Add(st, 6, string.Format(Strings.LintWithout, "ElseIf", Strings.LintBlockIfAfter));
                        return;
                    case "else":
                        if (Top() == null || Top().Kind != Kind.If) Add(st, 4, string.Format(Strings.LintWithout, "Else", "If"));
                        return;
                    case "endif":
                        Close(Kind.If, st, "End If", Strings.LintBlockIfAfter);
                        return;
                    case "end":
                        Match endWhat = Regex.Match(s, @"^End\s+(If|With|Select)\b", Ci);
                        if (!endWhat.Success) return; // End (encerra o programa)
                        string what = endWhat.Groups[1].Value.ToLowerInvariant();
                        if (what == "if") Close(Kind.If, st, "End If", Strings.LintBlockIfAfter);
                        else if (what == "with") Close(Kind.With, st, "End With", "With");
                        else Close(Kind.Select, st, "End Select", "Select Case");
                        return;
                    case "for":
                        Match f = ForStart.Match(s);
                        Block b = Push(Kind.For, st, 3);
                        if (f.Success) b.Variable = f.Groups[1].Value;
                        return;
                    case "next":
                        NextStatement(st);
                        return;
                    case "do":
                        Push(Kind.Do, st, 2);
                        return;
                    case "loop":
                        Close(Kind.Do, st, "Loop", "Do");
                        return;
                    case "while":
                        Push(Kind.While, st, 5);
                        return;
                    case "wend":
                        Close(Kind.While, st, "Wend", "While");
                        return;
                    case "with":
                        Push(Kind.With, st, 4);
                        return;
                    case "select":
                        if (Regex.IsMatch(s, @"^Select\s+Case\b", Ci)) Push(Kind.Select, st, 6);
                        return;
                    case "case":
                        if (Top() == null || Top().Kind != Kind.Select) Add(st, 4, string.Format(Strings.LintWithout, "Case", "Select Case"));
                        return;
                    case "exit":
                        Match e = Exit.Match(s);
                        if (e.Success && e.Index == 0) ExitCheck(st, e);
                        return;
                }
            }

            static string FirstWordOf(string s)
            {
                Match m = FirstWord.Match(s);
                return m.Success ? m.Groups[1].Value : "";
            }

            static string Proper(string s)
            {
                s = Regex.Replace(s.Trim(), @"\s+", " ");
                StringBuilder sb = new StringBuilder();
                foreach (string part in s.Split(' '))
                {
                    if (sb.Length > 0) sb.Append(' ');
                    sb.Append(char.ToUpperInvariant(part[0])).Append(part.Substring(1).ToLowerInvariant());
                }
                return sb.ToString();
            }

            // ----------------------------------------------------------------
            // Procedimentos
            // ----------------------------------------------------------------

            void StartProcedure(Statement st, Match proc)
            {
                if (procedure != null)
                {
                    // Um procedimento dentro de outro: o de antes ficou sem o End
                    CloseOpen(0);
                    Add(procedure, string.Format(Strings.LintExpected, "End " + procedure.ProcKind));
                    LabelsCheck();
                    procedure = null;
                    stack.Clear();
                }
                string kind = Proper(proc.Groups[1].Value.Split(' ')[0]); // Sub, Function, Property
                string full = Proper(proc.Groups[1].Value);               // Property Get...
                string name = proc.Groups[2].Value;
                Block b = new Block { Kind = Kind.Procedure, Line = st.Line, Column = st.Column, Length = st.Text.Length, ProcKind = kind, Name = name };
                // Nomes no módulo: Property Get, Let e Set com o mesmo nome podem (um de cada); o
                // resto (outro procedimento ou variável com o nome) é ambíguo
                string existing;
                bool clash = false;
                if (moduleNames.TryGetValue(name, out existing))
                    foreach (string part in existing.Split('|'))
                        if (part == full || !(part.StartsWith("Property") && full.StartsWith("Property"))) clash = true;
                if (clash && directiveDepth == 0) Add(st, st.Text.Length, string.Format(Strings.LintAmbiguous, name));
                else if (!clash) moduleNames[name] = existing == null ? full : existing + "|" + full;

                procedure = b;
                stack.Clear();
                stack.Add(b);
                localNames.Clear();
                labels.Clear();
                jumps.Clear();
                if (kind == "Function" || full == "Property Get") localNames[name] = true; // o nome da função é uma variável dela
                Parameters(st, proc);
            }

            void EndProcedure(Statement st, string kind)
            {
                if (procedure == null)
                {
                    Add(st, st.Text.Length, string.Format(Strings.LintWithout, "End " + kind, kind));
                    return;
                }
                if (procedure.ProcKind != kind)
                {
                    // End Function num Sub: o VBA pede o End do procedimento aberto
                    Add(st, st.Text.Length, string.Format(Strings.LintExpected, "End " + procedure.ProcKind));
                }
                CloseOpen(0);
                LabelsCheck();
                procedure = null;
                stack.Clear();
                afterProcedures = true;
            }

            // Os parâmetros: nomes repetidos são declaração duplicada
            void Parameters(Statement st, Match proc)
            {
                string s = st.Text;
                int open = s.IndexOf('(', proc.Index + proc.Length);
                if (open < 0) return;
                int depth = 0, close = -1;
                for (int k = open; k < s.Length; k++)
                {
                    if (s[k] == '(') depth++;
                    else if (s[k] == ')') { depth--; if (depth == 0) { close = k; break; } }
                }
                if (close < 0) return;
                foreach (KeyValuePair<int, string> item in SplitTop(s, open + 1, close))
                {
                    string p = Regex.Replace(item.Value, @"^\s*(?:(?:Optional|ByVal|ByRef|ParamArray)\s+)*", "", Ci);
                    Match id = Identifier.Match(p);
                    if (!id.Success) continue;
                    Declare(localNames, id.Groups[1].Value, st, item.Key);
                }
            }

            // ----------------------------------------------------------------
            // Declarações
            // ----------------------------------------------------------------

            void ModuleLevel(Statement st, string word)
            {
                string s = st.Text;
                if (afterProcedures)
                {
                    Add(st, s.Length, Strings.LintAfterProcedures);
                    return;
                }
                Match ts = TypeStart.Match(s);
                if (ts.Success)
                {
                    typeBlock = Proper(ts.Groups[1].Value);
                    stack.Add(new Block { Kind = Kind.Type, Line = st.Line, Column = st.Column, Length = s.Length, Name = typeBlock });
                    return;
                }
                Match te = TypeEnd.Match(s);
                if (te.Success) { Add(st, s.Length, string.Format(Strings.LintWithout, "End " + Proper(te.Groups[1].Value), Proper(te.Groups[1].Value))); return; }
                Match decl = DeclareProc.Match(s);
                if (decl.Success)
                {
                    string name = decl.Groups[1].Value;
                    if (directiveDepth == 0 && moduleNames.ContainsKey(name)) Add(st, s.Length, string.Format(Strings.LintAmbiguous, name));
                    else moduleNames[name] = "Declare";
                    return;
                }
                if (Regex.IsMatch(s, @"^(?:(?:Public|Private)\s+)?Event\b", Ci)) return; // Public Event Mudou()
                switch (word.ToLowerInvariant())
                {
                    case "option": case "implements": case "event": case "attribute":
                    case "defbool": case "defbyte": case "defint": case "deflng": case "deflnglng": case "deflngptr":
                    case "defcur": case "defsng": case "defdbl": case "defdate": case "defstr": case "defobj": case "defvar":
                        return;
                }
                Match vars = Variables.Match(s);
                if (vars.Success)
                {
                    int bodyStart = vars.Groups[3].Index;
                    foreach (KeyValuePair<int, string> item in SplitTop(s, bodyStart, s.Length))
                    {
                        Match id = Identifier.Match(item.Value);
                        if (!id.Success) continue;
                        string name = id.Groups[1].Value;
                        if (directiveDepth > 0) continue;
                        if (moduleNames.ContainsKey(name))
                        {
                            string kind = moduleNames[name];
                            Statement where = new Statement { Line = st.Line, Column = st.Column + item.Key + (item.Value.Length - item.Value.TrimStart().Length) };
                            AddAt(where, name.Length, kind == "Var" ? string.Format(Strings.LintDuplicate, name) : string.Format(Strings.LintAmbiguous, name));
                        }
                        else moduleNames[name] = "Var";
                    }
                    return;
                }
                Add(st, s.Length, Strings.LintOutsideProcedure);
            }

            void LocalDeclarations(Statement st)
            {
                string s = st.Text;
                Match m = Regex.Match(s, @"^(?:Dim|Static|Const)\s+(?:Const\s+)?", Ci);
                foreach (KeyValuePair<int, string> item in SplitTop(s, m.Length, s.Length))
                {
                    Match id = Identifier.Match(item.Value);
                    if (!id.Success) continue;
                    Declare(localNames, id.Groups[1].Value, st, item.Key + (item.Value.Length - item.Value.TrimStart().Length));
                }
            }

            void Declare(Dictionary<string, bool> scope, string name, Statement st, int offset)
            {
                if (directiveDepth > 0) return; // #If/#Else costuma declarar o mesmo nome nos dois trechos
                if (scope.ContainsKey(name))
                    AddAt(new Statement { Line = st.Line, Column = st.Column + offset }, name.Length, string.Format(Strings.LintDuplicate, name));
                else scope[name] = true;
            }

            // As partes separadas por vírgula fora de parênteses, com a posição de cada uma
            static List<KeyValuePair<int, string>> SplitTop(string s, int from, int to)
            {
                List<KeyValuePair<int, string>> parts = new List<KeyValuePair<int, string>>();
                int depth = 0, start = from;
                for (int k = from; k < to; k++)
                {
                    char c = s[k];
                    if (c == '(') depth++;
                    else if (c == ')') depth--;
                    else if (c == ',' && depth == 0)
                    {
                        parts.Add(new KeyValuePair<int, string>(start, s.Substring(start, k - start)));
                        start = k + 1;
                    }
                }
                if (to > start) parts.Add(new KeyValuePair<int, string>(start, s.Substring(start, to - start)));
                return parts;
            }

            // ----------------------------------------------------------------
            // Blocos
            // ----------------------------------------------------------------

            Block Top() { return stack.Count > 1 ? stack[stack.Count - 1] : null; } // o procedimento não conta

            Block Push(Kind kind, Statement st, int length)
            {
                Block b = new Block { Kind = kind, Line = st.Line, Column = st.Column, Length = length };
                stack.Add(b);
                return b;
            }

            void Pop(Kind kind)
            {
                for (int k = stack.Count - 1; k >= 0; k--)
                    if (stack[k].Kind == kind) { stack.RemoveAt(k); return; }
            }

            // Fecha o bloco do tipo pedido. Se houver blocos abertos dentro dele, eles é que estão
            // errados (ficaram sem o fechamento); sem o bloco, o fechamento está sobrando.
            void Close(Kind kind, Statement st, string closer, string opener)
            {
                int index = -1;
                for (int k = stack.Count - 1; k >= 1; k--)
                    if (stack[k].Kind == kind) { index = k; break; }
                if (index < 0)
                {
                    Add(st, closer.Length, string.Format(Strings.LintWithout, closer, opener));
                    return;
                }
                CloseOpen(index + 1);
                stack.RemoveAt(index);
            }

            // Os blocos abertos a partir de "from" ficaram sem fechamento: acusa cada um na linha dele
            void CloseOpen(int from)
            {
                for (int k = stack.Count - 1; k >= Math.Max(1, from); k--)
                {
                    Block b = stack[k];
                    Add(b, Unclosed(b.Kind));
                    stack.RemoveAt(k);
                }
            }

            static string Unclosed(Kind kind)
            {
                switch (kind)
                {
                    case Kind.If: return string.Format(Strings.LintWithout, Strings.LintBlockIf, "End If");
                    case Kind.For: return string.Format(Strings.LintWithout, "For", "Next");
                    case Kind.Do: return string.Format(Strings.LintWithout, "Do", "Loop");
                    case Kind.While: return string.Format(Strings.LintWithout, "While", "Wend");
                    case Kind.With: return string.Format(Strings.LintWithout, "With", "End With");
                    case Kind.Select: return string.Format(Strings.LintWithout, "Select Case", "End Select");
                    default: return "";
                }
            }

            // Next [variável][, variável...]: cada variável fecha um For, de dentro para fora
            void NextStatement(Statement st)
            {
                string rest = st.Text.Substring(4).Trim();
                List<string> vars = new List<string>();
                if (rest.Length > 0) foreach (string v in rest.Split(',')) vars.Add(v.Trim());
                if (vars.Count == 0) vars.Add(null);
                foreach (string v in vars)
                {
                    Block top = Top();
                    if (top != null && top.Kind == Kind.For)
                    {
                        if (v != null && top.Variable != null && !v.Equals(top.Variable, StringComparison.OrdinalIgnoreCase))
                        {
                            Add(st, st.Text.Length, Strings.LintNextVariable);
                            // Next i com o For j aberto: fecha o For i (de fora) e o For j continua
                            // aberto, para o Next j de depois não virar outro erro
                            for (int k = stack.Count - 2; k >= 1; k--)
                                if (stack[k].Kind == Kind.For && v.Equals(stack[k].Variable, StringComparison.OrdinalIgnoreCase)) { stack.RemoveAt(k); return; }
                            stack.RemoveAt(stack.Count - 1);
                            return;
                        }
                        stack.RemoveAt(stack.Count - 1);
                        continue;
                    }
                    Close(Kind.For, st, "Next", "For");
                }
            }

            void ExitCheck(Statement st, Match e)
            {
                string what = Proper(e.Groups[1].Value);
                Statement where = new Statement { Line = st.Line, Column = st.Column + e.Index };
                if (what == "For" || what == "Do")
                {
                    Kind kind = what == "For" ? Kind.For : Kind.Do;
                    for (int k = stack.Count - 1; k >= 1; k--) if (stack[k].Kind == kind) return;
                    AddAt(where, e.Length, string.Format(Strings.LintExitOutside, "Exit " + what, what == "For" ? "For...Next" : "Do...Loop"));
                    return;
                }
                if (procedure != null && procedure.ProcKind != what)
                    AddAt(where, e.Length, string.Format(Strings.LintExitNotAllowed, what, procedure.ProcKind));
            }

            // ----------------------------------------------------------------
            // Rótulos
            // ----------------------------------------------------------------

            void Jumps(Statement st)
            {
                string s = st.Text;
                foreach (Match m in Jump.Matches(s))
                {
                    if (Regex.IsMatch(s.Substring(0, m.Index), @"\bOn\s+Error\s*$", Ci) && (m.Groups[2].Value == "0")) continue;
                    jumps.Add(new Statement { Line = st.Line, Column = st.Column + m.Groups[2].Index, Text = m.Groups[2].Value });
                }
                Match on = OnJump.Match(s);
                if (on.Success)
                    foreach (KeyValuePair<int, string> item in SplitTop(s, on.Groups[1].Index, s.Length))
                    {
                        string name = item.Value.Trim();
                        if (name.Length > 0 && !jumps.Exists(j => j.Line == st.Line && j.Text.Equals(name, StringComparison.OrdinalIgnoreCase)))
                            jumps.Add(new Statement { Line = st.Line, Column = st.Column + item.Key + (item.Value.Length - item.Value.TrimStart().Length), Text = name });
                    }
                foreach (Match m in Resume.Matches(s))
                {
                    string name = m.Groups[1].Value;
                    if (name.Equals("Next", StringComparison.OrdinalIgnoreCase) || name == "0") continue;
                    jumps.Add(new Statement { Line = st.Line, Column = st.Column + m.Groups[1].Index, Text = name });
                }
            }

            void LabelsCheck()
            {
                foreach (Statement j in jumps)
                    if (!labels.ContainsKey(j.Text)) AddAt(j, j.Text.Length, string.Format(Strings.LintLabel, j.Text));
                jumps.Clear();
                labels.Clear();
            }

            // ----------------------------------------------------------------

            void Finish()
            {
                if (procedure != null)
                {
                    CloseOpen(0);
                    Add(procedure, string.Format(Strings.LintExpected, "End " + procedure.ProcKind));
                    LabelsCheck();
                    procedure = null;
                }
                if (typeBlock != null)
                    foreach (Block b in stack) if (b.Kind == Kind.Type) Add(b, string.Format(Strings.LintExpected, "End " + b.Name));
                for (int k = 0; k < directiveLines.Count; k++)
                    problems.Add(new Problem { Line = directiveLines[k], Column = 1, Length = 3, Message = string.Format(Strings.LintWithout, "#If", "#End If") });
            }

            string CurrentProcedure { get { return procedure != null ? procedure.Name : ""; } }

            void Add(Statement st, int length, string message) { AddAt(st, length, message); }

            void AddAt(Statement st, int length, string message)
            {
                problems.Add(new Problem { Line = st.Line, Column = st.Column, Length = Math.Max(1, length), Message = message, Procedure = CurrentProcedure });
            }

            void Add(Block b, string message)
            {
                problems.Add(new Problem { Line = b.Line, Column = b.Column, Length = Math.Max(1, b.Length), Message = message,
                    Procedure = procedure != null ? procedure.Name : b.Name ?? "" });
            }
        }
    }
}

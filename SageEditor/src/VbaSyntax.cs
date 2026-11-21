using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace SageEditor
{
    // A sintaxe de cada comando, para a aba Problemas (os erros que o VBE marca em vermelho):
    //   - cabeçalho de Sub/Function/Property: depois do nome, só (parâmetros) e As Tipo;
    //   - declarações (Dim, Static, Const, Public, Private, Global e os parâmetros):
    //     nome [(limites)] [As [New] Tipo [* n]], e o = do Const;
    //   - parênteses desencontrados;
    //   - If/ElseIf sem Then, For sem = ou To, For Each sem In;
    //   - dois valores seguidos, sem operador entre eles, numa atribuição ou condição (x = a b).
    // Só acusa o que tem certeza: um caractere que a leitura não conhece faz a verificação
    // daquele trecho parar, em vez de arriscar um alarme falso.
    static partial class VbaLinter
    {
        enum Tk { Name, Number, Text, Date, Op, Open, Close, Comma, Dot, Bang, Other }

        sealed class Token
        {
            public Tk Kind;
            public int Start, Length;  // no texto do comando
            public string Value;
            public bool Is(string word) { return Kind == Tk.Name && Value.Equals(word, StringComparison.OrdinalIgnoreCase); }
            public bool IsOp(string op) { return Kind == Tk.Op && Value == op; }
        }

        // Palavras que ligam dois valores (ou vêm antes de um) numa expressão
        static readonly HashSet<string> OperatorWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "And", "Or", "Not", "Xor", "Eqv", "Imp", "Mod", "Is", "Like", "New", "TypeOf", "AddressOf", "ByVal", "ByRef",
        };

        static bool NameStart(char c) { return char.IsLetter(c) || c == '_'; }
        static bool NameChar(char c) { return char.IsLetterOrDigit(c) || c == '_'; }
        static bool Hex(char c) { return char.IsDigit(c) || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F'); }

        // As partes do comando (o texto já vem sem comentário e com as strings em branco)
        static List<Token> Tokens(string s)
        {
            List<Token> t = new List<Token>();
            int k = 0;
            bool afterOperand = false;
            while (k < s.Length)
            {
                char c = s[k];
                if (char.IsWhiteSpace(c)) { k++; continue; }
                int start = k;
                Tk kind;
                if (NameStart(c))
                {
                    while (k < s.Length && NameChar(s[k])) k++;
                    // O sufixo de tipo (a%, b$, c#...); o ! seguido de nome é o operador (Forms!Nome)
                    if (k < s.Length && "%&$!#@^".IndexOf(s[k]) >= 0)
                    {
                        bool bang = s[k] == '!' && k + 1 < s.Length && (NameStart(s[k + 1]) || s[k + 1] == '[');
                        bool hex = s[k] == '&' && k + 1 < s.Length && "HhOo".IndexOf(s[k + 1]) >= 0;
                        if (!bang && !hex) k++;
                    }
                    kind = Tk.Name;
                }
                else if (char.IsDigit(c) || (c == '.' && !afterOperand && k + 1 < s.Length && char.IsDigit(s[k + 1])))
                {
                    while (k < s.Length && char.IsDigit(s[k])) k++;
                    if (k < s.Length && s[k] == '.') { k++; while (k < s.Length && char.IsDigit(s[k])) k++; }
                    if (k < s.Length && "eEdD".IndexOf(s[k]) >= 0)
                    {
                        int e = k + 1;
                        if (e < s.Length && (s[e] == '+' || s[e] == '-')) e++;
                        if (e < s.Length && char.IsDigit(s[e])) { k = e; while (k < s.Length && char.IsDigit(s[k])) k++; }
                    }
                    if (k < s.Length && "%&!#@^".IndexOf(s[k]) >= 0) k++;
                    kind = Tk.Number;
                }
                else if (c == '&' && !afterOperand && k + 1 < s.Length && "HhOo".IndexOf(s[k + 1]) >= 0)
                {
                    k += 2;
                    while (k < s.Length && Hex(s[k])) k++;
                    if (k < s.Length && "&%^".IndexOf(s[k]) >= 0) k++;
                    kind = Tk.Number;
                }
                else if (c == '"')
                {
                    int close = s.IndexOf('"', k + 1);
                    k = close < 0 ? s.Length : close + 1;
                    kind = Tk.Text;
                }
                else if (c == '#' && !afterOperand)
                {
                    int close = s.IndexOf('#', k + 1);
                    if (close < 0) { k++; kind = Tk.Other; } // Print #1, Input(1, #1)...
                    else { k = close + 1; kind = Tk.Date; }
                }
                else if (c == '[')
                {
                    int close = s.IndexOf(']', k + 1);
                    k = close < 0 ? s.Length : close + 1;
                    kind = Tk.Name;
                }
                else if (c == '(') { k++; kind = Tk.Open; }
                else if (c == ')') { k++; kind = Tk.Close; }
                else if (c == ',') { k++; kind = Tk.Comma; }
                else if (c == '.') { k++; kind = Tk.Dot; }
                else if (c == '!') { k++; kind = Tk.Bang; }
                else if (c == ':' && k + 1 < s.Length && s[k + 1] == '=') { k += 2; kind = Tk.Op; }
                else if ("+-*/\\^&=<>;".IndexOf(c) >= 0) { k++; kind = Tk.Op; }
                else { k++; kind = Tk.Other; }
                t.Add(new Token { Kind = kind, Start = start, Length = k - start, Value = s.Substring(start, k - start) });
                afterOperand = kind == Tk.Name || kind == Tk.Number || kind == Tk.Text || kind == Tk.Date || kind == Tk.Close;
            }
            return t;
        }

        sealed partial class Analysis
        {
            void Expected(Statement st, int offset, int length, string what)
            {
                Statement where = st.At != null ? st.At(st.Offset + offset) : new Statement { Line = st.Line, Column = st.Column + offset };
                AddAt(where, length, string.Format(Strings.LintExpected, what));
            }

            void Expected(Statement st, Token at, string what) { Expected(st, at.Start, at.Length, what); }

            // No fim do comando (falta alguma coisa): aponta o comando inteiro
            void ExpectedAtEnd(Statement st, string what) { Expected(st, 0, st.Text.Length, what); }

            // Parênteses: um ) a mais acusa fim da instrução; um ( sem par acusa o )
            bool Balanced(Statement st, List<Token> t)
            {
                List<Token> open = new List<Token>();
                foreach (Token tok in t)
                {
                    if (tok.Kind == Tk.Open) open.Add(tok);
                    else if (tok.Kind == Tk.Close)
                    {
                        if (open.Count == 0) { Expected(st, tok, Strings.LintEndOfStatement); return false; }
                        open.RemoveAt(open.Count - 1);
                    }
                }
                if (open.Count > 0) { Expected(st, open[open.Count - 1], ")"); return false; }
                return true;
            }

            // O fim do trecho entre parênteses que começa em "open" (o índice do ")")
            static int Matching(List<Token> t, int open)
            {
                int depth = 0;
                for (int k = open; k < t.Count; k++)
                {
                    if (t[k].Kind == Tk.Open) depth++;
                    else if (t[k].Kind == Tk.Close && --depth == 0) return k;
                }
                return -1;
            }

            // Os itens separados por vírgula fora de parênteses: [início, fim)
            static List<int[]> Items(List<Token> t, int from, int to)
            {
                List<int[]> items = new List<int[]>();
                int depth = 0, start = from;
                for (int k = from; k < to; k++)
                {
                    if (t[k].Kind == Tk.Open) depth++;
                    else if (t[k].Kind == Tk.Close) depth--;
                    else if (t[k].Kind == Tk.Comma && depth == 0) { items.Add(new[] { start, k }); start = k + 1; }
                }
                items.Add(new[] { start, to });
                return items;
            }

            // ------------------------------------------------------------
            // Cabeçalho e declarações
            // ------------------------------------------------------------

            // Sub Nome [(parâmetros)] [As Tipo]
            void Header(Statement st, List<Token> t, int nameStart)
            {
                int i = t.FindIndex(x => x.Start == nameStart);
                if (i < 0) return;
                i++;
                if (i < t.Count && t[i].Kind == Tk.Open)
                {
                    int close = Matching(t, i);
                    if (close < 0) return;
                    if (close > i + 1)
                        foreach (int[] item in Items(t, i + 1, close))
                            if (!Item(st, t, item[0], item[1], ItemMode.Parameter)) break;
                    i = close + 1;
                }
                if (i < t.Count && t[i].Is("As"))
                {
                    i = TypeName(st, t, i + 1, t.Count, false);
                    if (i < 0) return;
                    if (i + 1 < t.Count && t[i].Kind == Tk.Open && t[i + 1].Kind == Tk.Close) i += 2; // As Long()
                }
                if (i < t.Count) Expected(st, t[i], Strings.LintEndOfStatement);
            }

            enum ItemMode { Variable, Constant, Parameter }

            // Public/Private/Global/Dim/Static/Const no módulo (o ModuleLevel já viu que é declaração)
            void ModuleDeclarations(Statement st, List<Token> t)
            {
                Match vars = Variables.Match(st.Text);
                if (!vars.Success) return;
                bool constant = vars.Groups[1].Value.Equals("Const", StringComparison.OrdinalIgnoreCase) || vars.Groups[2].Success;
                int from = t.FindIndex(x => x.Start >= vars.Groups[3].Index);
                if (from < 0) return;
                DeclarationItems(st, t, from, constant);
            }

            // Dim/Static/Const num procedimento
            void Declarations(Statement st, List<Token> t, int from)
            {
                bool constant = t[0].Is("Const");
                if (from < t.Count && t[from].Is("Const")) { constant = true; from++; }
                DeclarationItems(st, t, from, constant);
            }

            void DeclarationItems(Statement st, List<Token> t, int from, bool constant)
            {
                if (from >= t.Count) { ExpectedAtEnd(st, Strings.LintIdentifier); return; }
                foreach (int[] item in Items(t, from, t.Count))
                    if (!Item(st, t, item[0], item[1], constant ? ItemMode.Constant : ItemMode.Variable)) return;
            }

            // Um item: [modificadores] nome [(limites)] [As [New] Tipo [* n]] [= valor]
            // Devolve falso quando acusou (o resto do comando não é conferido)
            bool Item(Statement st, List<Token> t, int i, int to, ItemMode mode)
            {
                if (i >= to)
                {
                    // Vírgula sem nada depois (Dim a, ) ou antes
                    Token at = i < t.Count ? t[i] : t[t.Count - 1];
                    Expected(st, at, Strings.LintIdentifier);
                    return false;
                }
                if (mode == ItemMode.Parameter)
                    while (i < to && (t[i].Is("Optional") || t[i].Is("ByVal") || t[i].Is("ByRef") || t[i].Is("ParamArray"))) i++;
                else if (mode == ItemMode.Variable && t[i].Is("WithEvents")) i++;
                if (i >= to || t[i].Kind != Tk.Name || OperatorWords.Contains(t[i].Value))
                {
                    Expected(st, i < to ? t[i] : t[to - 1], Strings.LintIdentifier);
                    return false;
                }
                i++;
                if (i < to && t[i].Kind == Tk.Open)
                {
                    if (mode == ItemMode.Constant) { Expected(st, t[i], "="); return false; }
                    int close = Matching(t, i);
                    if (close < 0 || close >= to) return false;
                    if (mode == ItemMode.Variable && close > i + 1 && !Expression(st, t, i + 1, close, true)) return false;
                    i = close + 1;
                }
                if (i < to && t[i].Is("As"))
                {
                    i = TypeName(st, t, i + 1, to, mode == ItemMode.Variable);
                    if (i < 0) return false;
                }
                if (mode == ItemMode.Constant)
                {
                    if (i >= to || !t[i].IsOp("=")) { if (i < to) Expected(st, t[i], "="); else Expected(st, t[to - 1], "="); return false; }
                    if (i + 1 >= to) { Expected(st, t[i], Strings.LintExpression); return false; }
                    return Expression(st, t, i + 1, to, false);
                }
                if (mode == ItemMode.Parameter && i < to && t[i].IsOp("="))
                {
                    if (i + 1 >= to) { Expected(st, t[i], Strings.LintExpression); return false; }
                    return Expression(st, t, i + 1, to, false);
                }
                if (i < to) { Expected(st, t[i], Strings.LintEndOfStatement); return false; }
                return true;
            }

            // O tipo depois do As: [New] Nome[.Nome] [* n]. Devolve o índice depois do tipo (-1: acusou)
            int TypeName(Statement st, List<Token> t, int i, int to, bool allowNew)
            {
                if (allowNew && i < to && t[i].Is("New")) i++;
                if (i >= to || t[i].Kind != Tk.Name) { Expected(st, i < to ? t[i] : t[to - 1], Strings.LintIdentifier); return -1; }
                i++;
                while (i + 1 < to && t[i].Kind == Tk.Dot && t[i + 1].Kind == Tk.Name) i += 2;
                if (i < to && t[i].IsOp("*")) // String * 10
                {
                    if (i + 1 >= to || (t[i + 1].Kind != Tk.Number && t[i + 1].Kind != Tk.Name)) { Expected(st, t[i], Strings.LintExpression); return -1; }
                    i += 2;
                }
                return i;
            }

            // ------------------------------------------------------------
            // Comandos
            // ------------------------------------------------------------

            void Syntax(Statement st, List<Token> t)
            {
                if (t.Count == 0) return;
                Token first = t[0];
                if (first.Is("If") || first.Is("ElseIf"))
                {
                    int then = t.FindIndex(x => x.Is("Then"));
                    if (then < 0)
                    {
                        // If x GoTo rótulo também vale (só no If)
                        int jump = first.Is("If") ? t.FindIndex(x => x.Is("GoTo")) : -1;
                        if (jump < 0) { ExpectedAtEnd(st, first.Is("If") ? Strings.LintThenOrGoTo : "Then"); return; }
                        then = jump;
                    }
                    if (then == 1) { Expected(st, t[1], Strings.LintExpression); return; }
                    Expression(st, t, 1, then, false, first.Is("If") ? Strings.LintThenOrGoTo : "Then");
                    return;
                }
                if (first.Is("For")) { ForSyntax(st, t); return; }
                if (first.Is("Do") || first.Is("Loop"))
                {
                    if (t.Count > 1 && (t[1].Is("While") || t[1].Is("Until")))
                    {
                        if (t.Count == 2) { Expected(st, t[1], Strings.LintExpression); return; }
                        Expression(st, t, 2, t.Count, false);
                    }
                    return;
                }
                if (first.Is("While"))
                {
                    if (t.Count == 1) { Expected(st, first, Strings.LintExpression); return; }
                    Expression(st, t, 1, t.Count, false);
                    return;
                }
                if (first.Is("Select"))
                {
                    if (t.Count > 1 && t[1].Is("Case"))
                    {
                        if (t.Count == 2) { Expected(st, t[1], Strings.LintExpression); return; }
                        Expression(st, t, 2, t.Count, false);
                    }
                    return;
                }
                Assignment(st, t);
            }

            // For var = início To fim [Step passo] / For Each var In coleção
            void ForSyntax(Statement st, List<Token> t)
            {
                if (t.Count > 1 && t[1].Is("Each"))
                {
                    if (t.Count < 3 || t[2].Kind != Tk.Name) { Expected(st, t.Count > 2 ? t[2] : t[1], Strings.LintIdentifier); return; }
                    if (t.Count < 4 || !t[3].Is("In")) { Expected(st, t.Count > 3 ? t[3] : t[2], "In"); return; }
                    if (t.Count == 4) { Expected(st, t[3], Strings.LintExpression); return; }
                    Expression(st, t, 4, t.Count, false);
                    return;
                }
                if (t.Count < 2 || t[1].Kind != Tk.Name) { Expected(st, t.Count > 1 ? t[1] : t[0], Strings.LintIdentifier); return; }
                if (t.Count > 2 && (t[2].Kind == Tk.Dot || t[2].Kind == Tk.Open || t[2].Kind == Tk.Bang)) return; // não é uma variável simples
                if (t.Count < 3 || !t[2].IsOp("=")) { Expected(st, t.Count > 2 ? t[2] : t[1], "="); return; }
                int to = -1, step = -1, depth = 0;
                for (int k = 3; k < t.Count; k++)
                {
                    if (t[k].Kind == Tk.Open) depth++;
                    else if (t[k].Kind == Tk.Close) depth--;
                    else if (depth == 0 && to < 0 && t[k].Is("To")) to = k;
                    else if (depth == 0 && to >= 0 && t[k].Is("Step")) { step = k; break; }
                }
                if (to < 0)
                {
                    // For i = 1 3: o To era esperado no lugar do segundo valor
                    Token extra = Adjacent(t, 3, t.Count, false);
                    if (extra != null) Expected(st, extra, "To"); else ExpectedAtEnd(st, "To");
                    return;
                }
                if (to == 3) { Expected(st, t[to], Strings.LintExpression); return; }
                if (!Expression(st, t, 3, to, false, "To")) return;
                int end = step < 0 ? t.Count : step;
                if (end == to + 1) { Expected(st, t[to], Strings.LintExpression); return; }
                if (!Expression(st, t, to + 1, end, false)) return;
                if (step >= 0)
                {
                    if (step + 1 >= t.Count) { Expected(st, t[step], Strings.LintExpression); return; }
                    Expression(st, t, step + 1, t.Count, false);
                }
            }

            // [Set|Let] alvo = expressão: o alvo é nome[(...)][.membro[(...)]]...
            void Assignment(Statement st, List<Token> t)
            {
                int i = 0;
                if ((t[0].Is("Set") || t[0].Is("Let")) && t.Count > 1 && (t[1].Kind == Tk.Name || t[1].Kind == Tk.Dot)) i = 1;
                if (i < t.Count && t[i].Kind == Tk.Dot) i++;
                if (i >= t.Count || t[i].Kind != Tk.Name) return;
                i++;
                while (i < t.Count)
                {
                    if (t[i].Kind == Tk.Open) { int close = Matching(t, i); if (close < 0) return; i = close + 1; }
                    else if ((t[i].Kind == Tk.Dot || t[i].Kind == Tk.Bang) && i + 1 < t.Count && t[i + 1].Kind == Tk.Name) i += 2;
                    else break;
                }
                if (i >= t.Count || !t[i].IsOp("=")) return; // não é atribuição (chamada, Print...)
                if (i + 1 >= t.Count) { Expected(st, t[i], Strings.LintExpression); return; }
                Expression(st, t, i + 1, t.Count, false);
            }

            // Dois valores seguidos sem operador entre eles (x = a b, If a b Then). Nas listas entre
            // parênteses, "To" separa limites (Dim a(1 To 3)). Devolve falso quando acusou.
            bool Expression(Statement st, List<Token> t, int from, int to, bool bounds)
            {
                return Expression(st, t, from, to, bounds, Strings.LintEndOfStatement);
            }

            // "expected": o que o VBA pede no lugar do valor a mais (If a b Then: Then ou GoTo)
            bool Expression(Statement st, List<Token> t, int from, int to, bool bounds, string expected)
            {
                Token extra = Adjacent(t, from, to, bounds);
                if (extra == null) return true;
                Expected(st, extra, expected);
                return false;
            }

            // O primeiro valor que vem logo depois de outro, sem operador entre eles (null: nenhum)
            static Token Adjacent(List<Token> t, int from, int to, bool bounds)
            {
                bool prevOperand = false, afterMember = false;
                for (int k = from; k < to; k++)
                {
                    Token tok = t[k];
                    if (tok.Kind == Tk.Other) return null; // algo que a leitura não conhece: não arrisca
                    if (tok.Kind == Tk.Dot || tok.Kind == Tk.Bang) { afterMember = true; prevOperand = false; continue; }
                    bool keyword = tok.Kind == Tk.Name && !afterMember &&
                        (OperatorWords.Contains(tok.Value) || (bounds && tok.Is("To")));
                    bool operand = (tok.Kind == Tk.Name && !keyword) || tok.Kind == Tk.Number || tok.Kind == Tk.Text || tok.Kind == Tk.Date;
                    if (operand && prevOperand) return tok;
                    afterMember = false;
                    prevOperand = operand || tok.Kind == Tk.Close;
                }
                return null;
            }
        }
    }
}

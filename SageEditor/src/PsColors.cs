using System;
using System.Collections.Generic;
using System.Drawing;
using System.Management.Automation.Language;

namespace SageEditor
{
    // Cores da linha de comando da aba Terminal, pelo analisador do próprio PowerShell:
    // comandos (python, git, Get-Date) em amarelo, parâmetros (-m, --upgrade) em cinza,
    // textos, variáveis, números, palavras-chave e comentários com cores próprias.
    static class PsColors
    {
        public struct Span
        {
            public int Start, Length;
            public Color Color;
            public Span(int start, int length, Color color) { Start = start; Length = length; Color = color; }
        }

        sealed class Palette
        {
            public Color Command, Parameter, Text, Variable, Number, Keyword, Comment, Operator, Type;
        }

        static readonly Palette Dark = new Palette
        {
            Command = Color.FromArgb(0xF9, 0xF1, 0xA5),
            Parameter = Color.FromArgb(0x8A, 0x8A, 0x8A),
            Text = Color.FromArgb(0x3A, 0x96, 0xDD),
            Variable = Color.FromArgb(0x16, 0xC6, 0x0C),
            Number = Color.FromArgb(0xB5, 0xCE, 0xA8),
            Keyword = Color.FromArgb(0xC5, 0x86, 0xC0),
            Comment = Color.FromArgb(0x6A, 0x99, 0x55),
            Operator = Color.FromArgb(0x8A, 0x8A, 0x8A),
            Type = Color.FromArgb(0x4E, 0xC9, 0xB0),
        };

        static readonly Palette Light = new Palette
        {
            Command = Color.FromArgb(0x79, 0x5E, 0x26),
            Parameter = Color.FromArgb(0x6E, 0x6E, 0x6E),
            Text = Color.FromArgb(0xA3, 0x15, 0x15),
            Variable = Color.FromArgb(0x00, 0x10, 0x80),
            Number = Color.FromArgb(0x09, 0x86, 0x58),
            Keyword = Color.FromArgb(0x00, 0x00, 0xFF),
            Comment = Color.FromArgb(0x00, 0x80, 0x00),
            Operator = Color.FromArgb(0x6E, 0x6E, 0x6E),
            Type = Color.FromArgb(0x26, 0x7F, 0x99),
        };

        // Trechos coloridos do texto; o que não aparece fica na cor normal
        public static List<Span> Colorize(string text, bool dark)
        {
            List<Span> spans = new List<Span>();
            if (string.IsNullOrEmpty(text)) return spans;
            Palette p = dark ? Dark : Light;
            Token[] tokens;
            ParseError[] errors;
            try { Parser.ParseInput(text, out tokens, out errors); }
            catch (Exception) { return spans; }
            foreach (Token token in tokens) Add(spans, token, p, text.Length);
            return spans;
        }

        // Cor de cada caractere (null: a cor normal); os trechos de dentro (variáveis num texto)
        // vencem os de fora
        public static Color?[] Map(string text, bool dark)
        {
            Color?[] map = new Color?[text == null ? 0 : text.Length];
            foreach (Span s in Colorize(text, dark))
                for (int i = s.Start; i < s.Start + s.Length && i < map.Length; i++) map[i] = s.Color;
            return map;
        }

        static void Add(List<Span> spans, Token token, Palette p, int length)
        {
            Color? color = ColorOf(token, p);
            int start = token.Extent.StartOffset, end = Math.Min(token.Extent.EndOffset, length);
            if (color != null && end > start) spans.Add(new Span(start, end - start, color.Value));
            // Variáveis dentro de textos com aspas duplas ("Olá $nome")
            StringExpandableToken expandable = token as StringExpandableToken;
            if (expandable != null && expandable.NestedTokens != null)
                foreach (Token nested in expandable.NestedTokens) Add(spans, nested, p, length);
        }

        static Color? ColorOf(Token token, Palette p)
        {
            TokenFlags f = token.TokenFlags;
            if ((f & TokenFlags.CommandName) != 0) return p.Command;
            // --upgrade, --version: opções de programas externos (pip, git), que o PowerShell
            // trata como texto comum
            if (token.Kind == TokenKind.Generic && token.Text.Length > 2 && token.Text.StartsWith("--")) return p.Parameter;
            switch (token.Kind)
            {
                case TokenKind.Parameter: return p.Parameter;
                case TokenKind.Variable:
                case TokenKind.SplattedVariable: return p.Variable;
                case TokenKind.Number: return p.Number;
                case TokenKind.Comment: return p.Comment;
                case TokenKind.StringLiteral:
                case TokenKind.StringExpandable:
                case TokenKind.HereStringLiteral:
                case TokenKind.HereStringExpandable: return p.Text;
                case TokenKind.Pipe:
                case TokenKind.Redirection:
                case TokenKind.RedirectInStd: return p.Operator;
            }
            if ((f & TokenFlags.Keyword) != 0) return p.Keyword;
            if ((f & TokenFlags.TypeName) != 0) return p.Type;
            if ((f & (TokenFlags.BinaryOperator | TokenFlags.UnaryOperator | TokenFlags.AssignmentOperator)) != 0) return p.Operator;
            return null;
        }
    }
}

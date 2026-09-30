using System;
using System.Collections.Generic;
using System.Drawing;

namespace SageVBE
{
    // Tema de cores. Default (Colors == null) deixa o VBE como ele é.
    sealed class Theme
    {
        public string Name;
        public bool IsDark;

        // Índice COLOR_* do GetSysColor -> COLORREF, válido só nas janelas do VBE
        public Dictionary<int, int> SysColors;

        // Paleta de 16 cores do editor (Preto, Azul-marinho, Verde, Azul-petróleo, Marrom,
        // Roxo, Oliva, Prata, Cinza, Azul, Verde-limão, Ciano, Vermelho, Magenta, Amarelo,
        // Branco), usada nas opções "Formato do editor": uma versão para a cor do texto e
        // outra para a cor de fundo. null = paleta original.
        public int[] Palette, BackPalette;

        // Cores da tela de Configurações
        public Color Background, Sidebar, Foreground, Muted, Border, Input, Accent, Hover;

        readonly Dictionary<int, IntPtr> brushes = new Dictionary<int, IntPtr>();
        HashSet<int> own;

        // Cor que já pertence ao tema (evita converter duas vezes)
        public bool OwnsColor(int colorRef)
        {
            if (own == null)
            {
                HashSet<int> set = new HashSet<int>(SysColors.Values);
                own = set;
            }
            return own.Contains(colorRef);
        }

        public bool IsDefault { get { return SysColors == null; } }

        public IntPtr Brush(int colorRef)
        {
            IntPtr brush;
            lock (brushes)
            {
                if (!brushes.TryGetValue(colorRef, out brush))
                {
                    brush = Native.CreateSolidBrush(colorRef);
                    brushes[colorRef] = brush; // nunca liberados: podem estar em uso pelo VBE
                }
            }
            return brush;
        }

        public int Window { get { return SysColors[COLOR_WINDOW]; } }
        public int WindowText { get { return SysColors[COLOR_WINDOWTEXT]; } }
        public int Face { get { return SysColors[COLOR_BTNFACE]; } }
        public int Caption { get { return SysColors[COLOR_ACTIVECAPTION]; } }
        public int CaptionText { get { return SysColors[COLOR_CAPTIONTEXT]; } }
        public int FrameBorder { get { return SysColors[COLOR_WINDOWFRAME]; } }

        // ------------------------------------------------------------------

        public const string DefaultName = "Padrão do VBE";

        public static readonly Theme[] All = new Theme[]
        {
            CreateDefault(),
            Create("Dark Modern", true,
                editor: "#1F1F1F", sidebar: "#181818", text: "#CCCCCC", muted: "#9D9D9D",
                border: "#2B2B2B", selection: "#264F78", selectionText: "#FFFFFF",
                input: "#313131", accent: "#0078D4", hover: "#2A2D2E",
                palette: DarkPalette(), backPalette: DarkBackPalette("#1F1F1F")),
            Create("Dark+", true,
                editor: "#1E1E1E", sidebar: "#252526", text: "#D4D4D4", muted: "#9D9D9D",
                border: "#3C3C3C", selection: "#264F78", selectionText: "#FFFFFF",
                input: "#3C3C3C", accent: "#007ACC", hover: "#2A2D2E",
                palette: DarkPalette(), backPalette: DarkBackPalette("#1E1E1E")),
            Create("Sage", true,
                editor: "#2B352D", sidebar: "#243027", text: "#DCE5D8", muted: "#9AA894",
                border: "#3B4A3E", selection: "#4E6E56", selectionText: "#FFFFFF",
                input: "#34403A", accent: "#8DAA85", hover: "#33403A",
                palette: SagePalette(), backPalette: DarkBackPalette("#2B352D")),
            Create("Light Modern", false,
                editor: "#FFFFFF", sidebar: "#F8F8F8", text: "#3B3B3B", muted: "#717171",
                border: "#E5E5E5", selection: "#ADD6FF", selectionText: "#000000",
                input: "#FFFFFF", accent: "#005FB8", hover: "#F2F2F2",
                palette: LightPalette(), backPalette: LightPalette()),
        };

        public static Theme Find(string name)
        {
            foreach (Theme t in All)
                if (string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase)) return t;
            return All[0];
        }

        static Theme CreateDefault()
        {
            Theme t = new Theme();
            t.Name = DefaultName;
            t.Background = SystemColors.Window;
            t.Sidebar = SystemColors.Control;
            t.Foreground = SystemColors.ControlText;
            t.Muted = SystemColors.GrayText;
            t.Border = SystemColors.ControlDark;
            t.Input = SystemColors.Window;
            t.Accent = SystemColors.Highlight;
            t.Hover = SystemColors.ControlLight;
            return t;
        }

        const int COLOR_SCROLLBAR = 0, COLOR_ACTIVECAPTION = 2, COLOR_INACTIVECAPTION = 3, COLOR_MENU = 4,
            COLOR_WINDOW = 5, COLOR_WINDOWFRAME = 6, COLOR_MENUTEXT = 7, COLOR_WINDOWTEXT = 8, COLOR_CAPTIONTEXT = 9,
            COLOR_ACTIVEBORDER = 10, COLOR_INACTIVEBORDER = 11, COLOR_APPWORKSPACE = 12, COLOR_HIGHLIGHT = 13,
            COLOR_HIGHLIGHTTEXT = 14, COLOR_BTNFACE = 15, COLOR_BTNSHADOW = 16, COLOR_GRAYTEXT = 17, COLOR_BTNTEXT = 18,
            COLOR_INACTIVECAPTIONTEXT = 19, COLOR_BTNHIGHLIGHT = 20, COLOR_3DDKSHADOW = 21, COLOR_3DLIGHT = 22,
            COLOR_INFOTEXT = 23, COLOR_INFOBK = 24, COLOR_HOTLIGHT = 26, COLOR_GRADIENTACTIVECAPTION = 27,
            COLOR_GRADIENTINACTIVECAPTION = 28, COLOR_MENUHILIGHT = 29, COLOR_MENUBAR = 30;

        static Theme Create(string name, bool dark, string editor, string sidebar, string text, string muted,
            string border, string selection, string selectionText, string input, string accent, string hover,
            int[] palette, int[] backPalette)
        {
            Theme t = new Theme();
            t.Name = name;
            t.IsDark = dark;
            t.Palette = palette;
            t.BackPalette = backPalette;
            t.Background = ColorTranslator.FromHtml(editor);
            t.Sidebar = ColorTranslator.FromHtml(sidebar);
            t.Foreground = ColorTranslator.FromHtml(text);
            t.Muted = ColorTranslator.FromHtml(muted);
            t.Border = ColorTranslator.FromHtml(border);
            t.Input = ColorTranslator.FromHtml(input);
            t.Accent = ColorTranslator.FromHtml(accent);
            t.Hover = ColorTranslator.FromHtml(hover);

            int ed = Ref(editor), side = Ref(sidebar), fg = Ref(text), mute = Ref(muted), bd = Ref(border);
            Dictionary<int, int> c = new Dictionary<int, int>();
            c[COLOR_SCROLLBAR] = side;
            c[COLOR_ACTIVECAPTION] = side;
            c[COLOR_GRADIENTACTIVECAPTION] = side;
            c[COLOR_INACTIVECAPTION] = side;
            c[COLOR_GRADIENTINACTIVECAPTION] = side;
            c[COLOR_CAPTIONTEXT] = fg;
            c[COLOR_INACTIVECAPTIONTEXT] = mute;
            c[COLOR_MENU] = side;
            c[COLOR_MENUBAR] = side;
            c[COLOR_MENUTEXT] = fg;
            c[COLOR_MENUHILIGHT] = Ref(accent);
            c[COLOR_WINDOW] = ed;
            c[COLOR_WINDOWTEXT] = fg;
            c[COLOR_WINDOWFRAME] = bd;
            c[COLOR_ACTIVEBORDER] = bd;
            c[COLOR_INACTIVEBORDER] = bd;
            c[COLOR_APPWORKSPACE] = ed;
            c[COLOR_HIGHLIGHT] = Ref(selection);
            c[COLOR_HIGHLIGHTTEXT] = Ref(selectionText);
            c[COLOR_BTNFACE] = side;
            c[COLOR_BTNTEXT] = fg;
            c[COLOR_BTNSHADOW] = bd;
            c[COLOR_BTNHIGHLIGHT] = bd;
            c[COLOR_3DLIGHT] = bd;
            c[COLOR_3DDKSHADOW] = dark ? Ref("#0F0F0F") : Ref("#C8C8C8");
            c[COLOR_GRAYTEXT] = mute;
            c[COLOR_INFOBK] = side;
            c[COLOR_INFOTEXT] = fg;
            c[COLOR_HOTLIGHT] = Ref(accent);
            t.SysColors = c;
            return t;
        }

        // Cor do texto no tema escuro
        static int[] DarkPalette()
        {
            return new int[]
            {
                Ref("#D4D4D4"), // Preto        -> texto claro
                Ref("#569CD6"), // Azul-marinho -> palavras-chave
                Ref("#6A9955"), // Verde        -> comentários
                Ref("#4EC9B0"), // Azul-petróleo
                Ref("#D16969"), // Marrom
                Ref("#C586C0"), // Roxo
                Ref("#DCDCAA"), // Oliva
                Ref("#C0C0C0"), // Prata
                Ref("#9D9D9D"), // Cinza
                Ref("#3794FF"), // Azul
                Ref("#B5CEA8"), // Verde-limão
                Ref("#9CDCFE"), // Ciano
                Ref("#F44747"), // Vermelho     -> erro de sintaxe
                Ref("#D670D6"), // Magenta
                Ref("#DCDCAA"), // Amarelo
                Ref("#FFFFFF"), // Branco       -> texto do ponto de interrupção
            };
        }

        // Cor de fundo no tema escuro: tons escuros, e branco vira o fundo do editor
        static int[] DarkBackPalette(string editor)
        {
            return new int[]
            {
                Ref("#000000"), // Preto
                Ref("#1B2B4B"), // Azul-marinho
                Ref("#1E3A1E"), // Verde
                Ref("#1B3B3B"), // Azul-petróleo
                Ref("#6E1B1B"), // Marrom       -> ponto de interrupção
                Ref("#3B1E3B"), // Roxo
                Ref("#3B3B1E"), // Oliva
                Ref("#3C3C3C"), // Prata
                Ref("#505050"), // Cinza
                Ref("#1F3F6F"), // Azul
                Ref("#2F4F2F"), // Verde-limão  -> retorno de chamada
                Ref("#1F4E5C"), // Ciano        -> indicador
                Ref("#5A1D1D"), // Vermelho
                Ref("#4F1F4F"), // Magenta
                Ref("#4B4B18"), // Amarelo      -> ponto de execução
                Ref(editor),    // Branco       -> fundo do editor
            };
        }

        // Cor do texto no tema Sage: verdes-sálvia e tons terrosos suaves
        static int[] SagePalette()
        {
            return new int[]
            {
                Ref("#D3DDD0"), // Preto        -> texto claro
                Ref("#A3C9A8"), // Azul-marinho -> palavras-chave (sálvia clara)
                Ref("#7B8C76"), // Verde        -> comentários (sálvia apagada)
                Ref("#7FB8A4"), // Azul-petróleo
                Ref("#D18A7A"), // Marrom
                Ref("#B79BC2"), // Roxo
                Ref("#D6C98E"), // Oliva
                Ref("#C0C8BD"), // Prata
                Ref("#98A393"), // Cinza
                Ref("#8FB4C9"), // Azul
                Ref("#B8D4A0"), // Verde-limão
                Ref("#9ED0C8"), // Ciano
                Ref("#E27D6F"), // Vermelho     -> erro de sintaxe
                Ref("#C79AC7"), // Magenta
                Ref("#E0D48F"), // Amarelo
                Ref("#FFFFFF"), // Branco       -> texto do ponto de interrupção
            };
        }

        static int[] LightPalette()
        {
            return new int[]
            {
                Ref("#000000"), Ref("#0000FF"), Ref("#008000"), Ref("#267F99"),
                Ref("#A31515"), Ref("#AF00DB"), Ref("#795E26"), Ref("#C0C0C0"),
                Ref("#808080"), Ref("#0070C1"), Ref("#C8E6C9"), Ref("#B3E5FC"),
                Ref("#E51400"), Ref("#AF00DB"), Ref("#FFF59D"), Ref("#FFFFFF"),
            };
        }

        // "#RRGGBB" -> COLORREF (0x00BBGGRR)
        public static int Ref(string hex)
        {
            Color c = ColorTranslator.FromHtml(hex);
            return c.R | (c.G << 8) | (c.B << 16);
        }
    }
}

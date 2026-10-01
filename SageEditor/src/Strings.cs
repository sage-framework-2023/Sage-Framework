using System;
using System.Globalization;
using Microsoft.Win32;

namespace SageEditor
{
    // Textos da interface no idioma do Office (o mesmo do Excel e do VBE). Para outro
    // idioma: um método com os textos traduzidos e uma linha em Load.
    //
    // Idioma: "locale" no settings.json (como no VS Code), senão o idioma da interface
    // do Office (UILanguageTag), senão o do Windows. Português é o padrão; os demais
    // idiomas usam inglês até ganharem tradução própria.
    static class Strings
    {
        public static string Language; // "pt" ou "en"

        // Menu Sage
        public static string MenuSettings;

        // Configurações
        public static string SettingsTitle, SearchCue, UserTab, NoResults;
        public static string NavCommon, NavAppearance, NavEditor;
        public static string On, Off;
        public static string AppearanceCategory, ColorTheme, ColorThemeDescription, ColorThemeKeywords;
        public static string EditorCategory, LineNumbers, LineNumbersDescription, LineNumbersKeywords;
        public static string ShowTabs, ShowTabsDescription, ShowTabsKeywords;
        public static string DefaultTheme;

        // Abas
        public static string Close, CloseOthers, CloseToRight, CloseAll;

        public static void Load()
        {
            Language = Detect();
            if (Language == "pt") Portuguese();
            else English();
        }

        static string Detect()
        {
            string tag = Settings.Get(Settings.LocaleKey, "");
            if (tag.Length == 0) tag = OfficeLanguage();
            if (tag.Length == 0) tag = CultureInfo.CurrentUICulture.Name;
            return tag.StartsWith("pt", StringComparison.OrdinalIgnoreCase) ? "pt" : "en";
        }

        // HKCU\Software\Microsoft\Office\<versão>\Common\LanguageResources: UILanguageTag
        // (ex.: "pt-br"), a menos que o Office siga o idioma do Windows
        static string OfficeLanguage()
        {
            foreach (string version in new string[] { "16.0", "15.0" })
            {
                try
                {
                    using (RegistryKey k = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\" + version + @"\Common\LanguageResources"))
                    {
                        if (k == null) continue;
                        object follow = k.GetValue("FollowSystemUILanguage");
                        if (follow is int && (int)follow == 1) return "";
                        string tag = k.GetValue("UILanguageTag") as string;
                        if (!string.IsNullOrEmpty(tag)) return tag;
                    }
                }
                catch (Exception ex) { Log.Error(ex); }
            }
            return "";
        }

        static void Portuguese()
        {
            MenuSettings = "&Configurações...";

            SettingsTitle = "Configurações - Sage";
            SearchCue = "Pesquisar configurações";
            UserTab = "Usuário";
            NoResults = "Nenhuma configuração encontrada";
            NavCommon = "Comumente Usado";
            NavAppearance = "Aparência";
            NavEditor = "Editor";
            On = "Ativado";
            Off = "Desativado";

            AppearanceCategory = "Aparência: ";
            ColorTheme = "Tema de Cores";
            ColorThemeDescription = "Especifica o tema de cores de todo o editor do VBA: menus, barras de ferramentas, " +
                "janelas de Projeto, Propriedades e Verificação imediata, a Caixa de ferramentas e o código.";
            ColorThemeKeywords = "tema cores aparência theme color workbench dark light escuro claro";

            EditorCategory = "Editor: ";
            LineNumbers = "Números de Linha";
            LineNumbersDescription = "Mostra o número de cada linha à esquerda do código, com a linha atual em destaque. " +
                "Funciona com os temas do Sage (não com o Padrão do VBE).";
            LineNumbersKeywords = "linha linhas números numeração line numbers editor margem";
            ShowTabs = "Mostrar Abas";
            ShowTabsDescription = "Mostra as janelas abertas (módulos, classes e formulários) como abas no topo da área de código. " +
                "Clique para ativar, arraste para mudar de lugar, \"×\" ou botão do meio para fechar e botão direito para mais opções.";
            ShowTabsKeywords = "abas aba guias tabs janelas abertas módulos workbench editor show tabs";

            DefaultTheme = "Padrão do VBE";

            Close = "Fechar";
            CloseOthers = "Fechar Outras";
            CloseToRight = "Fechar à Direita";
            CloseAll = "Fechar Todas";
        }

        static void English()
        {
            MenuSettings = "&Settings...";

            SettingsTitle = "Settings - Sage";
            SearchCue = "Search settings";
            UserTab = "User";
            NoResults = "No settings found";
            NavCommon = "Commonly Used";
            NavAppearance = "Appearance";
            NavEditor = "Editor";
            On = "On";
            Off = "Off";

            AppearanceCategory = "Appearance: ";
            ColorTheme = "Color Theme";
            ColorThemeDescription = "Specifies the color theme of the whole VBA editor: menus, toolbars, " +
                "Project, Properties and Immediate windows, the Toolbox and the code.";
            ColorThemeKeywords = "theme color appearance workbench dark light";

            EditorCategory = "Editor: ";
            LineNumbers = "Line Numbers";
            LineNumbersDescription = "Shows the number of each line to the left of the code, with the current line highlighted. " +
                "Works with the Sage themes (not with VBE Default).";
            LineNumbersKeywords = "line numbers editor gutter margin";
            ShowTabs = "Show Tabs";
            ShowTabsDescription = "Shows the open windows (modules, classes and forms) as tabs at the top of the code area. " +
                "Click to activate, drag to reorder, \"×\" or middle button to close and right button for more options.";
            ShowTabsKeywords = "tabs open windows modules workbench editor show tabs";

            DefaultTheme = "VBE Default";

            Close = "Close";
            CloseOthers = "Close Others";
            CloseToRight = "Close to the Right";
            CloseAll = "Close All";
        }
    }
}

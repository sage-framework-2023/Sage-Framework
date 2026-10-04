using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace SageEditor
{
    // %APPDATA%\Sage\settings.json, no formato do VS Code (só valores texto por enquanto):
    //   { "workbench.colorTheme": "Dark Modern" }
    static class Settings
    {
        public static readonly string Folder = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Sage");
        public static readonly string FilePath = Path.Combine(Folder, "settings.json");

        public const string ColorThemeKey = "workbench.colorTheme";
        public const string LineNumbersKey = "editor.lineNumbers";
        public const string EditorTabsKey = "workbench.editor.showTabs";
        public const string AutoReferenceKey = "sage.autoReference";
        public const string MultiCursorKey = "sage.multiCursor";
        public const string LocaleKey = "locale"; // opcional: força o idioma ("pt-BR", "en"); senão, o do Office

        static readonly object sync = new object();
        static readonly SortedDictionary<string, string> values = new SortedDictionary<string, string>();

        public static string ColorTheme
        {
            get { return Get(ColorThemeKey, Theme.DefaultName); }
            set { Set(ColorThemeKey, value); }
        }

        // "on" (padrão) ou "off", como no VS Code
        public static bool LineNumbers
        {
            get { return !string.Equals(Get(LineNumbersKey, "on"), "off", StringComparison.OrdinalIgnoreCase); }
            set { Set(LineNumbersKey, value ? "on" : "off"); }
        }

        // "on" (padrão) ou "off": referência ao Sage Framework nas pastas de trabalho novas (AutoReference)
        public static bool AutoReference
        {
            get { return !string.Equals(Get(AutoReferenceKey, "on"), "off", StringComparison.OrdinalIgnoreCase); }
            set { Set(AutoReferenceKey, value ? "on" : "off"); }
        }

        // "on" (padrão) ou "off": vários cursores (MultiCursor)
        public static bool MultiCursor
        {
            get { return !string.Equals(Get(MultiCursorKey, "on"), "off", StringComparison.OrdinalIgnoreCase); }
            set { Set(MultiCursorKey, value ? "on" : "off"); }
        }

        // "multiple" (padrão) ou "none", como no VS Code
        public static bool EditorTabs
        {
            get { return !string.Equals(Get(EditorTabsKey, "multiple"), "none", StringComparison.OrdinalIgnoreCase); }
            set { Set(EditorTabsKey, value ? "multiple" : "none"); }
        }

        public static string Get(string key, string fallback)
        {
            lock (sync)
            {
                string v;
                return values.TryGetValue(key, out v) ? v : fallback;
            }
        }

        public static void Set(string key, string value)
        {
            lock (sync)
            {
                values[key] = value;
                Save();
            }
        }

        public static void Load()
        {
            lock (sync)
            {
                values.Clear();
                if (!File.Exists(FilePath)) return;
                string json = File.ReadAllText(FilePath, Encoding.UTF8);
                foreach (Match m in Regex.Matches(json, "\"((?:[^\"\\\\]|\\\\.)*)\"\\s*:\\s*\"((?:[^\"\\\\]|\\\\.)*)\""))
                    values[Unescape(m.Groups[1].Value)] = Unescape(m.Groups[2].Value);
            }
        }

        static void Save()
        {
            StringBuilder sb = new StringBuilder("{\r\n");
            int i = 0;
            foreach (KeyValuePair<string, string> kv in values)
            {
                sb.Append("    \"").Append(Escape(kv.Key)).Append("\": \"").Append(Escape(kv.Value)).Append('"');
                sb.Append(++i < values.Count ? ",\r\n" : "\r\n");
            }
            sb.Append("}\r\n");
            Directory.CreateDirectory(Folder);
            File.WriteAllText(FilePath, sb.ToString(), new UTF8Encoding(false));
        }

        static string Escape(string s)
        {
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"");
        }

        static string Unescape(string s)
        {
            return Regex.Replace(s, "\\\\(.)", "$1");
        }
    }
}

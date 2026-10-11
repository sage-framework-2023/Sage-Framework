using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace SageTypes
{
    // Correção ortográfica do Windows (a mesma do Edge e do Bloco de Notas): os idiomas são os
    // instalados no Windows (Configurações > Hora e idioma), e as palavras adicionadas ficam no
    // dicionário do usuário do Windows, que vale para todos os programas. Um corretor por idioma,
    // compartilhado entre os controles: "Ignorar tudo" vale para todos até o Excel fechar.
    static class Spelling
    {
        static ISpellCheckerFactory factory;
        static bool unavailable;
        static readonly Dictionary<string, ISpellChecker> checkers = new Dictionary<string, ISpellChecker>(StringComparer.OrdinalIgnoreCase);

        internal struct Error { public int Start, Length; }

        static ISpellCheckerFactory Factory()
        {
            if (factory == null && !unavailable)
            {
                try { factory = (ISpellCheckerFactory)Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("7AB36653-1796-484B-BDFA-E74F1DB7C1DC"))); }
                catch (Exception) { unavailable = true; }
            }
            return factory;
        }

        public static bool IsSupported(string language)
        {
            ISpellCheckerFactory f = Factory();
            if (f == null || string.IsNullOrEmpty(language)) return false;
            try
            {
                int supported;
                f.IsSupported(language, out supported);
                return supported != 0;
            }
            catch (Exception) { return false; }
        }

        // Idioma padrão: o do Windows, ou o primeiro instalado
        public static string DefaultLanguage()
        {
            foreach (CultureInfo c in new[] { CultureInfo.CurrentUICulture, CultureInfo.CurrentCulture })
                if (IsSupported(c.Name)) return c.Name;
            if (IsSupported("en-US")) return "en-US";
            List<string> all = Languages();
            return all.Count > 0 ? all[0] : "";
        }

        public static List<string> Languages()
        {
            List<string> result = new List<string>();
            ISpellCheckerFactory f = Factory();
            if (f == null) return result;
            try
            {
                IEnumString e;
                f.get_SupportedLanguages(out e);
                result.AddRange(Strings(e, 100));
            }
            catch (Exception) { }
            return result;
        }

        static ISpellChecker Checker(string language)
        {
            if (string.IsNullOrEmpty(language)) return null;
            ISpellChecker checker;
            if (checkers.TryGetValue(language, out checker)) return checker;
            checker = null;
            if (IsSupported(language))
            {
                try { Factory().CreateSpellChecker(language, out checker); }
                catch (Exception) { checker = null; }
            }
            checkers[language] = checker;
            return checker;
        }

        public static List<Error> Check(string language, string text)
        {
            List<Error> result = new List<Error>();
            ISpellChecker checker = Checker(language);
            if (checker == null || string.IsNullOrEmpty(text)) return result;
            try
            {
                IEnumSpellingError errors;
                checker.Check(text, out errors);
                while (true)
                {
                    ISpellingError error;
                    if (errors.Next(out error) != 0 || error == null) break;
                    uint start, length;
                    error.get_StartIndex(out start);
                    error.get_Length(out length);
                    result.Add(new Error { Start = (int)start, Length = (int)length });
                    Marshal.ReleaseComObject(error);
                }
                Marshal.ReleaseComObject(errors);
            }
            catch (Exception) { }
            return result;
        }

        public static List<string> Suggest(string language, string word, int max)
        {
            ISpellChecker checker = Checker(language);
            if (checker == null || string.IsNullOrEmpty(word)) return new List<string>();
            try
            {
                IEnumString e;
                checker.Suggest(word, out e);
                return Strings(e, max);
            }
            catch (Exception) { return new List<string>(); }
        }

        public static void Ignore(string language, string word)
        {
            ISpellChecker checker = Checker(language);
            if (checker != null) checker.Ignore(word);
        }

        public static void Add(string language, string word)
        {
            ISpellChecker checker = Checker(language);
            if (checker != null) checker.Add(word);
        }

        static List<string> Strings(IEnumString e, int max)
        {
            List<string> result = new List<string>();
            if (e == null) return result;
            string[] one = new string[1];
            while (result.Count < max && e.Next(1, one, IntPtr.Zero) == 0) result.Add(one[0]);
            Marshal.ReleaseComObject(e);
            return result;
        }

        [ComImport, Guid("8E018A9D-2415-4677-BF08-794EA61F94BB"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface ISpellCheckerFactory
        {
            void get_SupportedLanguages(out IEnumString languages);
            void IsSupported([MarshalAs(UnmanagedType.LPWStr)] string language, out int supported);
            void CreateSpellChecker([MarshalAs(UnmanagedType.LPWStr)] string language, out ISpellChecker checker);
        }

        [ComImport, Guid("B6FD0B71-E2BC-4653-8D05-F197E412770B"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface ISpellChecker
        {
            void get_LanguageTag([MarshalAs(UnmanagedType.LPWStr)] out string tag);
            void Check([MarshalAs(UnmanagedType.LPWStr)] string text, out IEnumSpellingError errors);
            void Suggest([MarshalAs(UnmanagedType.LPWStr)] string word, out IEnumString suggestions);
            void Add([MarshalAs(UnmanagedType.LPWStr)] string word);
            void Ignore([MarshalAs(UnmanagedType.LPWStr)] string word);
        }

        [ComImport, Guid("803E3BD4-2828-4410-8290-418D1D73C762"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IEnumSpellingError
        {
            [PreserveSig] int Next(out ISpellingError error);
        }

        [ComImport, Guid("B7C82D61-FBE8-4B47-9B27-6C0D2E0DE0A3"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface ISpellingError
        {
            void get_StartIndex(out uint start);
            void get_Length(out uint length);
            void get_CorrectiveAction(out int action);
            void get_Replacement([MarshalAs(UnmanagedType.LPWStr)] out string replacement);
        }
    }
}

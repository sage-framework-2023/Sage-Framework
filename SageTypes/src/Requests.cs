using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;

namespace SageTypes
{
    // Requests: a biblioteca requests do Python, para chamar APIs. Global, como no Python (sem Dim):
    //
    //   Dim r As Sage.Response
    //   Set r = Requests.Get("https://api.exemplo.com/itens", Params:=Array("pagina", 2))
    //   r.RaiseForStatus                              ' erro se o status for 4xx/5xx
    //   Debug.Print r.Json()("total")
    //
    //   Set r = Requests.Post(url, Json:=dados, Headers:=Array("Authorization", "Bearer " & token))
    //
    // Session guarda cabeçalhos, autenticação e cookies entre as chamadas (login, depois as consultas).
    //
    // Interfaces duais: acrescente membros só no fim, com o próximo DispId (ver StringS).
    [ComVisible(true), Guid("9CC2C2FF-5C73-44CC-8614-B4A15860D2D4"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface _Requests
    {
        [DispId(1)] Response Get(string Url, [Optional] object Params, [Optional] object Headers, [Optional] object Auth,
            [Optional] object Timeout, [Optional] object AllowRedirects);
        [DispId(2)] Response Post(string Url, [Optional] object Data, [Optional] object Json, [Optional] object Params,
            [Optional] object Headers, [Optional] object Auth, [Optional] object Timeout);
        [DispId(3)] Response Put(string Url, [Optional] object Data, [Optional] object Json, [Optional] object Params,
            [Optional] object Headers, [Optional] object Auth, [Optional] object Timeout);
        [DispId(4)] Response Patch(string Url, [Optional] object Data, [Optional] object Json, [Optional] object Params,
            [Optional] object Headers, [Optional] object Auth, [Optional] object Timeout);
        [DispId(5)] Response Delete(string Url, [Optional] object Params, [Optional] object Headers, [Optional] object Auth,
            [Optional] object Timeout);
        [DispId(6)] Response Head(string Url, [Optional] object Params, [Optional] object Headers, [Optional] object Auth,
            [Optional] object Timeout);
        [DispId(7)] Response Request(string Method, string Url, [Optional] object Params, [Optional] object Data, [Optional] object Json,
            [Optional] object Headers, [Optional] object Auth, [Optional] object Timeout, [Optional] object AllowRedirects);
        [DispId(8)] Session Session();
    }

    [ComVisible(true), Guid("18A710C0-DD90-407E-92A8-56E2E1FBB35C"), ProgId("Sage.Requests")]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(_Requests))]
    public sealed class Requests : _Requests
    {
        public Requests() { }

        // Cada chamada usa uma sessão nova (sem cookies guardados), como requests.get
        public Response Get(string Url, object Params, object Headers, object Auth, object Timeout, object AllowRedirects)
        { return new Session().Get(Url, Params, Headers, Auth, Timeout, AllowRedirects); }
        public Response Post(string Url, object Data, object Json, object Params, object Headers, object Auth, object Timeout)
        { return new Session().Post(Url, Data, Json, Params, Headers, Auth, Timeout); }
        public Response Put(string Url, object Data, object Json, object Params, object Headers, object Auth, object Timeout)
        { return new Session().Put(Url, Data, Json, Params, Headers, Auth, Timeout); }
        public Response Patch(string Url, object Data, object Json, object Params, object Headers, object Auth, object Timeout)
        { return new Session().Patch(Url, Data, Json, Params, Headers, Auth, Timeout); }
        public Response Delete(string Url, object Params, object Headers, object Auth, object Timeout)
        { return new Session().Delete(Url, Params, Headers, Auth, Timeout); }
        public Response Head(string Url, object Params, object Headers, object Auth, object Timeout)
        { return new Session().Head(Url, Params, Headers, Auth, Timeout); }
        public Response Request(string Method, string Url, object Params, object Data, object Json, object Headers, object Auth, object Timeout, object AllowRedirects)
        { return new Session().Request(Method, Url, Params, Data, Json, Headers, Auth, Timeout, AllowRedirects); }
        public Session Session() { return new Session(); }
    }

    [ComVisible(true), Guid("9B170703-92C8-47BF-9C0A-B7DFC1DCE26F"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface _Session
    {
        [DispId(1)] Response Get(string Url, [Optional] object Params, [Optional] object Headers, [Optional] object Auth,
            [Optional] object Timeout, [Optional] object AllowRedirects);
        [DispId(2)] Response Post(string Url, [Optional] object Data, [Optional] object Json, [Optional] object Params,
            [Optional] object Headers, [Optional] object Auth, [Optional] object Timeout);
        [DispId(3)] Response Put(string Url, [Optional] object Data, [Optional] object Json, [Optional] object Params,
            [Optional] object Headers, [Optional] object Auth, [Optional] object Timeout);
        [DispId(4)] Response Patch(string Url, [Optional] object Data, [Optional] object Json, [Optional] object Params,
            [Optional] object Headers, [Optional] object Auth, [Optional] object Timeout);
        [DispId(5)] Response Delete(string Url, [Optional] object Params, [Optional] object Headers, [Optional] object Auth,
            [Optional] object Timeout);
        [DispId(6)] Response Head(string Url, [Optional] object Params, [Optional] object Headers, [Optional] object Auth,
            [Optional] object Timeout);
        [DispId(7)] Response Request(string Method, string Url, [Optional] object Params, [Optional] object Data, [Optional] object Json,
            [Optional] object Headers, [Optional] object Auth, [Optional] object Timeout, [Optional] object AllowRedirects);

        // Valem para todas as chamadas da sessão (os da chamada são somados / têm prioridade)
        [DispId(10)] DictionaryS Headers { get; }
        [DispId(11)] DictionaryS Params { get; }
        [DispId(12)] object Auth { get; set; }
        [DispId(13)] object Timeout { get; set; }
        [DispId(14)] DictionaryS Cookies([Optional] object Url);

        [DispId(1000), PropertyLet("Auth")] void LetAuth(object Value);
        [DispId(1001), PropertyLet("Timeout")] void LetTimeout(object Value);
    }

    [ComVisible(true), Guid("FCB13B55-C345-4703-89B6-F4635B7C5A9A"), ProgId("Sage.Session")]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(_Session))]
    public sealed class Session : _Session
    {
        readonly DictionaryS headers = new DictionaryS();
        readonly DictionaryS parameters = new DictionaryS();
        readonly CookieContainer cookies = new CookieContainer();
        object auth;
        object timeout;

        static Session()
        {
            // O Excel não diz ao .NET qual versão dele usar, e o .NET fica no modo compatível
            // com o 4.0: TLS 1.0, recusado pelas APIs ("A conexão subjacente estava fechada").
            // Nesse modo nem o SystemDefault vale; por isso TLS 1.2 e 1.3 explícitos (o 1.3
            // só existe a partir do .NET 4.8)
            const SecurityProtocolType Tls12 = (SecurityProtocolType)3072, Tls13 = (SecurityProtocolType)12288;
            try { ServicePointManager.SecurityProtocol = Tls12 | Tls13; }
            catch (Exception) { ServicePointManager.SecurityProtocol = Tls12; }
            // Proxy de empresa: o do Windows, com o usuário logado
            if (WebRequest.DefaultWebProxy != null) WebRequest.DefaultWebProxy.Credentials = CredentialCache.DefaultCredentials;
        }

        public Session() { }

        public DictionaryS Headers { get { return headers; } }
        public DictionaryS Params { get { return parameters; } }
        public object Auth { get { return auth; } set { LetAuth(value); } }
        public void LetAuth(object Value) { auth = Interop.IsMissing(Value) ? null : Value; }
        public object Timeout { get { return timeout; } set { LetTimeout(value); } }
        public void LetTimeout(object Value) { timeout = Interop.IsMissing(Value) ? null : Value; }

        // Cookies guardados (de todos os sites, ou só dos de Url)
        public DictionaryS Cookies(object Url)
        {
            DictionaryS result = new DictionaryS();
            IEnumerable<Cookie> all;
            if (!Interop.IsMissing(Url)) all = cookies.GetCookies(new Uri(Interop.Text(Url))).Cast<Cookie>();
            else
            {
                // CookieContainer não lista tudo: lê a tabela interna por reflexão
                List<Cookie> list = new List<Cookie>();
                System.Collections.Hashtable table = (System.Collections.Hashtable)typeof(CookieContainer).GetField("m_domainTable",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).GetValue(cookies);
                foreach (object domain in table.Values)
                {
                    System.Collections.SortedList paths = (System.Collections.SortedList)domain.GetType().GetField("m_list",
                        System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance).GetValue(domain);
                    foreach (CookieCollection c in paths.Values) list.AddRange(c.Cast<Cookie>());
                }
                all = list;
            }
            foreach (Cookie c in all) result.LetValue(c.Name, c.Value);
            return result;
        }

        public Response Get(string Url, object Params, object Headers, object Auth, object Timeout, object AllowRedirects)
        { return Send("GET", Url, Params, null, null, Headers, Auth, Timeout, AllowRedirects); }
        public Response Post(string Url, object Data, object Json, object Params, object Headers, object Auth, object Timeout)
        { return Send("POST", Url, Params, Data, Json, Headers, Auth, Timeout, null); }
        public Response Put(string Url, object Data, object Json, object Params, object Headers, object Auth, object Timeout)
        { return Send("PUT", Url, Params, Data, Json, Headers, Auth, Timeout, null); }
        public Response Patch(string Url, object Data, object Json, object Params, object Headers, object Auth, object Timeout)
        { return Send("PATCH", Url, Params, Data, Json, Headers, Auth, Timeout, null); }
        public Response Delete(string Url, object Params, object Headers, object Auth, object Timeout)
        { return Send("DELETE", Url, Params, null, null, Headers, Auth, Timeout, null); }
        // Como no requests, Head não segue redirecionamentos
        public Response Head(string Url, object Params, object Headers, object Auth, object Timeout)
        { return Send("HEAD", Url, Params, null, null, Headers, Auth, Timeout, false); }
        public Response Request(string Method, string Url, object Params, object Data, object Json, object Headers, object Auth, object Timeout, object AllowRedirects)
        { return Send(Method.ToUpperInvariant(), Url, Params, Data, Json, Headers, Auth, Timeout, AllowRedirects); }

        // ------------------------------------------------------------------
        // Envio
        // ------------------------------------------------------------------

        Response Send(string method, string url, object queryParams, object data, object json, object callHeaders, object callAuth,
            object callTimeout, object allowRedirects)
        {
            // Parâmetros da sessão + os da chamada, na URL (?a=1&b=2)
            List<KeyValuePair<string, string>> query = Pairs(parameters).ToList();
            query.AddRange(Pairs(queryParams));
            Uri uri;
            if (!Uri.TryCreate(AppendQuery(url, query), UriKind.Absolute, out uri) || (uri.Scheme != "http" && uri.Scheme != "https"))
                throw Interop.Error(5, "InvalidURL: " + Interop.Repr(url) + " (use http:// ou https://)");

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
            request.Method = method;
            request.CookieContainer = cookies;
            request.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
            request.AllowAutoRedirect = Missing(allowRedirects) || Convert.ToBoolean(allowRedirects, CultureInfo.InvariantCulture);
            request.UserAgent = "Sage-Framework/1.0";
            request.Accept = "*/*";
            object seconds = !Missing(callTimeout) ? callTimeout : timeout;
            if (seconds != null)
            {
                int ms = (int)Math.Round(Convert.ToDouble(Interop.Unwrap(seconds), CultureInfo.InvariantCulture) * 1000);
                request.Timeout = ms;
                request.ReadWriteTimeout = ms;
            }

            // Autenticação básica: Auth = Array("usuário", "senha")
            object a = !Missing(callAuth) ? callAuth : auth;
            if (a != null)
            {
                List<object> pair = Interop.Items(a).ToList();
                if (pair.Count != 2) throw Interop.Error(5, "ValueError: Auth deve ser Array(\"usuário\", \"senha\").");
                string token = Convert.ToBase64String(Encoding.UTF8.GetBytes(Interop.Text(pair[0]) + ":" + Interop.Text(pair[1])));
                request.Headers[HttpRequestHeader.Authorization] = "Basic " + token;
            }

            // Corpo: Json (application/json), Data como dicionário (formulário), texto ou bytes
            byte[] body = null;
            string contentType = null;
            if (!Missing(json))
            {
                body = new UTF8Encoding(false).GetBytes(SageTypes.Json.Serialize(json, null, null, null));
                contentType = "application/json";
            }
            else if (!Missing(data))
            {
                object raw = Interop.Unwrap(data);
                if (raw is byte[]) body = (byte[])raw;
                else if (raw is string) body = Encoding.UTF8.GetBytes((string)raw);
                else
                {
                    body = Encoding.UTF8.GetBytes(Encode(Pairs(data)));
                    contentType = "application/x-www-form-urlencoded";
                }
            }

            // Cabeçalhos: os da sessão, depois os da chamada
            foreach (KeyValuePair<string, string> h in Pairs(headers).Concat(Pairs(callHeaders)))
            {
                if (string.Equals(h.Key, "Content-Type", StringComparison.OrdinalIgnoreCase)) contentType = h.Value;
                else SetHeader(request, h.Key, h.Value);
            }
            if (contentType != null) request.ContentType = contentType;

            Stopwatch clock = Stopwatch.StartNew();
            try
            {
                if (body != null)
                {
                    request.ContentLength = body.Length;
                    using (Stream s = request.GetRequestStream()) s.Write(body, 0, body.Length);
                }
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                    return new Response(response, method, clock.Elapsed);
            }
            catch (WebException e)
            {
                // 4xx/5xx: o requests devolve a resposta (RaiseForStatus gera o erro)
                HttpWebResponse response = e.Response as HttpWebResponse;
                if (response != null)
                    using (response) return new Response(response, method, clock.Elapsed);
                if (e.Status == WebExceptionStatus.Timeout) throw Interop.Error(5, "Timeout: " + url + " não respondeu a tempo.");
                throw Interop.Error(5, "ConnectionError: " + e.Message + " (" + url + ")");
            }
        }

        static bool Missing(object value) { return value == null || Interop.IsMissing(value); }

        // Cabeçalhos com propriedade própria no HttpWebRequest (o resto vai em Headers)
        static void SetHeader(HttpWebRequest request, string name, string value)
        {
            switch (name.ToLowerInvariant())
            {
                case "accept": request.Accept = value; break;
                case "user-agent": request.UserAgent = value; break;
                case "referer": request.Referer = value; break;
                case "connection": if (value.ToLowerInvariant() == "close") request.KeepAlive = false; else request.Connection = value; break;
                case "expect": request.Expect = value; break;
                case "host": request.Host = value; break;
                case "if-modified-since": request.IfModifiedSince = DateTime.Parse(value, CultureInfo.InvariantCulture); break;
                case "content-length": break; // calculado
                default: request.Headers[name] = value; break;
            }
        }

        // DictionaryS, Scripting.Dictionary ou Array(k1, v1, ...); valor lista repete a chave (a=1&a=2)
        static IEnumerable<KeyValuePair<string, string>> Pairs(object source)
        {
            if (Missing(source)) yield break;
            foreach (KeyValuePair<object, object> pair in Interop.PairsOf(source))
            {
                object value = pair.Value;
                if (value is ListS || value is Array)
                {
                    foreach (object item in Interop.Items(value)) yield return new KeyValuePair<string, string>(Interop.Text(pair.Key), Text(item));
                }
                else if (value != null) yield return new KeyValuePair<string, string>(Interop.Text(pair.Key), Text(value));
            }
        }

        // Texto para URL/cabeçalho: números com ponto, datas em ISO, True/False como no Python
        static string Text(object value)
        {
            object v = Interop.Unwrap(value);
            if (v is bool) return (bool)v ? "True" : "False";
            if (v is DateTime) return DatetimeS.From((DateTime)v, null).Iso8601("T", "auto");
            DatetimeS d = v as DatetimeS;
            if (d != null) return d.Iso8601("T", "auto");
            IFormattable f = v as IFormattable;
            if (f != null) return f.ToString(null, CultureInfo.InvariantCulture);
            return Interop.Text(v);
        }

        static string Encode(IEnumerable<KeyValuePair<string, string>> pairs)
        {
            return string.Join("&", pairs.Select(p => Escape(p.Key) + "=" + Escape(p.Value)));
        }

        static string Escape(string s)
        {
            // EscapeDataString tem limite de tamanho no .NET 4: em pedaços
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < s.Length; i += 30000) sb.Append(Uri.EscapeDataString(s.Substring(i, Math.Min(30000, s.Length - i))));
            return sb.ToString().Replace("%20", "+");
        }

        static string AppendQuery(string url, List<KeyValuePair<string, string>> query)
        {
            if (query.Count == 0) return url;
            int hash = url.IndexOf('#');
            string fragment = hash >= 0 ? url.Substring(hash) : "";
            if (hash >= 0) url = url.Substring(0, hash);
            return url + (url.Contains("?") ? (url.EndsWith("?") || url.EndsWith("&") ? "" : "&") : "?") + Encode(query) + fragment;
        }
    }

    // ----------------------------------------------------------------------
    // Resposta
    // ----------------------------------------------------------------------

    [ComVisible(true), Guid("9DEE6C77-AEF8-4815-BCD4-F032100D7ACA"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface _Response
    {
        [DispId(1)] int StatusCode { get; }
        [DispId(2)] bool Ok { get; }
        [DispId(3)] StringS Reason { get; }
        [DispId(4)] StringS Text { get; }
        [DispId(5)] object Content { get; }
        [DispId(6)] object Json();
        [DispId(7)] DictionaryS Headers { get; }
        [DispId(8)] StringS Url { get; }
        [DispId(9)] double Elapsed { get; }
        [DispId(10)] StringS Encoding { get; }
        [DispId(11)] void RaiseForStatus();
        [DispId(12)] string ToString();
        [DispId(13)] StringS Header(string Name);
        [DispId(14)] void Save(string Path);
        [DispId(15)] StringS JsonString([Optional] object Indent, [Optional] object SortKeys, [Optional] object EnsureAscii);
    }

    [ComVisible(true), Guid("B028ADB0-40B2-461C-8AD1-B4236852C17E"), ProgId("Sage.Response")]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(_Response))]
    public sealed class Response : _Response
    {
        readonly int status;
        readonly string reason, url, encoding, method;
        readonly byte[] content;
        readonly DictionaryS headers = new DictionaryS();
        readonly double elapsed;
        string text;

        public Response() { }

        internal Response(HttpWebResponse response, string method, TimeSpan elapsed)
        {
            this.method = method;
            status = (int)response.StatusCode;
            reason = response.StatusDescription;
            url = response.ResponseUri.ToString();
            foreach (string name in response.Headers.AllKeys) headers.LetValue(name, response.Headers[name]);
            using (Stream s = response.GetResponseStream())
            using (MemoryStream m = new MemoryStream())
            {
                s.CopyTo(m);
                content = m.ToArray();
            }
            this.elapsed = elapsed.TotalSeconds;
            // Codificação do Content-Type (charset=...); sem ela, UTF-8 (o padrão das APIs JSON)
            string charset = response.CharacterSet;
            encoding = string.IsNullOrEmpty(charset) || response.ContentType.IndexOf("charset", StringComparison.OrdinalIgnoreCase) < 0 ? "utf-8" : charset;
        }

        public int StatusCode { get { return status; } }
        public bool Ok { get { return status > 0 && status < 400; } }
        public StringS Reason { get { return Wrap(reason ?? ""); } }

        public StringS Text
        {
            get
            {
                if (text == null)
                {
                    System.Text.Encoding e;
                    try { e = System.Text.Encoding.GetEncoding(encoding); }
                    catch (ArgumentException) { e = new UTF8Encoding(false); }
                    text = content == null ? "" : e.GetString(content);
                    if (text.Length > 0 && text[0] == '﻿') text = text.Substring(1);
                }
                return Wrap(text);
            }
        }

        // Bytes (Byte() no VBA), para arquivos e imagens
        public object Content { get { return content ?? new byte[0]; } }

        public object Json()
        {
            try { return Interop.Wrap(SageTypes.Json.Parse(Text.Value)); }
            catch (COMException e)
            {
                throw Interop.Error(5, e.Message.Replace("JSONDecodeError", "JSONDecodeError (resposta " + status + ")"));
            }
        }

        // O JSON da resposta como texto: uma linha (padrão) ou formatado (Indent:=2), como
        // Json.Dumps(r.Json(), ...); r.Text é o texto exatamente como a API mandou
        public StringS JsonString(object Indent, object SortKeys, object EnsureAscii)
        {
            object value;
            try { value = SageTypes.Json.Parse(Text.Value); }
            catch (COMException e)
            {
                throw Interop.Error(5, e.Message.Replace("JSONDecodeError", "JSONDecodeError (resposta " + status + ")"));
            }
            return Wrap(SageTypes.Json.Serialize(value, Indent, SortKeys, EnsureAscii));
        }

        public DictionaryS Headers { get { return headers; } }

        // Um cabeçalho pelo nome, sem diferenciar maiúsculas (como r.headers["content-type"]); "" se não houver
        public StringS Header(string Name)
        {
            foreach (KeyValuePair<object, object> pair in headers.Pairs())
                if (string.Equals(pair.Key as string, Name, StringComparison.OrdinalIgnoreCase)) return Wrap(Interop.Text(pair.Value));
            return Wrap("");
        }

        public StringS Url { get { return Wrap(url ?? ""); } }
        public double Elapsed { get { return elapsed; } }
        public StringS Encoding { get { return Wrap(encoding ?? ""); } }

        // HTTPError do requests: "404 Client Error: Not Found for url: ..."
        public void RaiseForStatus()
        {
            if (status < 400) return;
            string kind = status < 500 ? "Client Error" : "Server Error";
            throw Interop.Error(5, "HTTPError: " + status + " " + kind + ": " + reason + " for url: " + url);
        }

        // Grava o conteúdo num arquivo (downloads)
        public void Save(string Path) { File.WriteAllBytes(Path, content ?? new byte[0]); }

        public override string ToString() { return "<Response [" + status + "]>"; }

        static StringS Wrap(string s) { return (StringS)Interop.Wrap(s); }
    }
}

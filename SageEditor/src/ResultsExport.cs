using System;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;

namespace SageEditor
{
    // "Salvar" da aba Resultado do DataFrame: grava o DataFrame mostrado (todas as linhas, não só
    // as da tela) em CSV, Excel, JSON, Markdown ou XML. As linhas vêm da mesma função que a grade
    // usa (o SageTypes as manda como texto: números no formato do Windows, datas em ISO, True/False
    // e null para NULL); os números e as datas voltam a ser números e datas onde o formato tem tipos.
    static class ResultsExport
    {
        public enum Format { Csv, Excel, Json, Markdown, Xml }

        public const long ExcelMaxRows = 1048575; // sem o cabeçalho
        const int Block = 10000;

        public static string Extension(Format f)
        {
            switch (f)
            {
                case Format.Csv: return ".csv";
                case Format.Excel: return ".xlsx";
                case Format.Json: return ".json";
                case Format.Markdown: return ".md";
                default: return ".xml";
            }
        }

        public static void Save(Format format, string path, string[] columns, string[] types, long rows, Func<long, int, string[][]> fetch)
        {
            switch (format)
            {
                case Format.Csv: Csv(path, columns, rows, fetch); break;
                case Format.Excel: Excel(path, columns, types, rows, fetch); break;
                case Format.Json: Json(path, columns, types, rows, fetch); break;
                case Format.Markdown: Markdown(path, columns, types, rows, fetch); break;
                case Format.Xml: Xml(path, columns, types, rows, fetch); break;
            }
        }

        // Todas as linhas, em blocos
        static void ForEachRow(long rows, Func<long, int, string[][]> fetch, Action<string[]> action)
        {
            for (long start = 0; start < rows; start += Block)
            {
                string[][] block = fetch(start, (int)Math.Min(Block, rows - start));
                if (block == null || block.Length == 0) break;
                foreach (string[] row in block) action(row);
            }
        }

        // ------------------------------------------------------------------
        // Tipos
        // ------------------------------------------------------------------

        enum Kind { Text, Number, Bool, Date, DateTime }

        static Kind KindOf(string type)
        {
            string t = (type ?? "").ToUpperInvariant();
            if (t.EndsWith("INT") || t == "INTEGER" || t == "DOUBLE" || t == "FLOAT" || t == "REAL" || t.StartsWith("DECIMAL") || t == "HUGEINT") return Kind.Number;
            if (t == "BOOLEAN") return Kind.Bool;
            if (t == "DATE") return Kind.Date;
            if (t.StartsWith("TIMESTAMP")) return Kind.DateTime;
            return Kind.Text;
        }

        static Kind[] Kinds(string[] types)
        {
            Kind[] k = new Kind[types.Length];
            for (int i = 0; i < types.Length; i++) k[i] = KindOf(types[i]);
            return k;
        }

        // O texto do número (formato do Windows) como número; false se não for
        static bool Number(string text, out double value)
        {
            return double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value);
        }

        static bool Date(string text, out DateTime value)
        {
            return DateTime.TryParseExact(text, new[] { "yyyy-MM-dd", "yyyy-MM-dd HH:mm:ss" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out value);
        }

        // ------------------------------------------------------------------
        // CSV: separador de lista do Windows (";" no pt-BR), como o df.ToCsv; UTF-8 com BOM,
        // para o Excel abrir com os acentos e os decimais certos
        // ------------------------------------------------------------------

        static void Csv(string path, string[] columns, long rows, Func<long, int, string[][]> fetch)
        {
            string sep = CultureInfo.CurrentCulture.TextInfo.ListSeparator;
            using (StreamWriter w = new StreamWriter(path, false, new UTF8Encoding(true)))
            {
                w.WriteLine(CsvLine(columns, sep));
                ForEachRow(rows, fetch, row => w.WriteLine(CsvLine(row, sep)));
            }
        }

        static string CsvLine(string[] values, string sep)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < values.Length; i++)
            {
                if (i > 0) sb.Append(sep);
                string v = values[i];
                if (v == null) continue; // NULL: vazio
                if (v.Contains(sep) || v.Contains("\"") || v.Contains("\n") || v.Contains("\r"))
                    sb.Append('"').Append(v.Replace("\"", "\"\"")).Append('"');
                else sb.Append(v);
            }
            return sb.ToString();
        }

        // ------------------------------------------------------------------
        // JSON: lista de objetos
        // ------------------------------------------------------------------

        static void Json(string path, string[] columns, string[] types, long rows, Func<long, int, string[][]> fetch)
        {
            Kind[] kinds = Kinds(types);
            using (StreamWriter w = new StreamWriter(path, false, new UTF8Encoding(false)))
            {
                w.Write("[");
                bool first = true;
                ForEachRow(rows, fetch, row =>
                {
                    w.Write(first ? "\n  {" : ",\n  {");
                    first = false;
                    for (int c = 0; c < columns.Length; c++)
                    {
                        if (c > 0) w.Write(", ");
                        w.Write(JsonString(columns[c]));
                        w.Write(": ");
                        w.Write(JsonValue(c < row.Length ? row[c] : null, kinds[c]));
                    }
                    w.Write("}");
                });
                w.Write(first ? "]\n" : "\n]\n");
            }
        }

        static string JsonValue(string v, Kind kind)
        {
            if (v == null) return "null";
            double d;
            if (kind == Kind.Number && Number(v, out d) && !double.IsNaN(d) && !double.IsInfinity(d)) return d.ToString("R", CultureInfo.InvariantCulture);
            if (kind == Kind.Bool) return v == "True" ? "true" : "false";
            return JsonString(v);
        }

        static string JsonString(string s)
        {
            StringBuilder sb = new StringBuilder("\"");
            foreach (char ch in s)
            {
                switch (ch)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (ch < 0x20) sb.Append("\\u").Append(((int)ch).ToString("x4"));
                        else sb.Append(ch);
                        break;
                }
            }
            return sb.Append('"').ToString();
        }

        // ------------------------------------------------------------------
        // Markdown: tabela, com os números alinhados à direita
        // ------------------------------------------------------------------

        static void Markdown(string path, string[] columns, string[] types, long rows, Func<long, int, string[][]> fetch)
        {
            Kind[] kinds = Kinds(types);
            using (StreamWriter w = new StreamWriter(path, false, new UTF8Encoding(false)))
            {
                w.WriteLine("| " + string.Join(" | ", Array.ConvertAll(columns, MdCell)) + " |");
                StringBuilder rule = new StringBuilder("|");
                foreach (Kind k in kinds) rule.Append(k == Kind.Number ? " ---: |" : " --- |");
                w.WriteLine(rule.ToString());
                ForEachRow(rows, fetch, row => w.WriteLine("| " + string.Join(" | ", Array.ConvertAll(row, MdCell)) + " |"));
            }
        }

        static string MdCell(string v)
        {
            if (v == null) return "";
            return v.Replace("\\", "\\\\").Replace("|", "\\|").Replace("\r\n", "<br>").Replace("\n", "<br>").Replace("\r", "<br>");
        }

        // ------------------------------------------------------------------
        // XML: <rows><row><Coluna>valor</Coluna></row></rows>; NULL com xsi:nil. Números e
        // booleanos no formato do XML Schema (1234.5, true), como no JSON
        // ------------------------------------------------------------------

        const string Xsi = "http://www.w3.org/2001/XMLSchema-instance";

        static void Xml(string path, string[] columns, string[] types, long rows, Func<long, int, string[][]> fetch)
        {
            Kind[] kinds = Kinds(types);
            string[] names = Array.ConvertAll(columns, c => XmlConvert.EncodeLocalName(string.IsNullOrEmpty(c) ? "column" : c));
            XmlWriterSettings settings = new XmlWriterSettings { Indent = true, Encoding = new UTF8Encoding(false) };
            using (XmlWriter w = XmlWriter.Create(path, settings))
            {
                w.WriteStartDocument();
                w.WriteStartElement("rows");
                w.WriteAttributeString("xmlns", "xsi", null, Xsi);
                ForEachRow(rows, fetch, row =>
                {
                    w.WriteStartElement("row");
                    for (int c = 0; c < names.Length; c++)
                    {
                        w.WriteStartElement(names[c]);
                        string v = c < row.Length ? row[c] : null;
                        double d;
                        if (v == null) w.WriteAttributeString("nil", Xsi, "true");
                        else if (kinds[c] == Kind.Number && Number(v, out d) && !double.IsNaN(d) && !double.IsInfinity(d)) w.WriteString(d.ToString("R", CultureInfo.InvariantCulture));
                        else if (kinds[c] == Kind.Bool) w.WriteString(v == "True" ? "true" : "false");
                        else w.WriteString(XmlSafe(v));
                        w.WriteEndElement();
                    }
                    w.WriteEndElement();
                });
                w.WriteEndElement();
                w.WriteEndDocument();
            }
        }

        // ------------------------------------------------------------------
        // Excel (.xlsx): o pacote Open XML montado à mão (sem abrir o Excel), com números como
        // números, datas como datas (com formato), True/False como booleanos e NULL vazio
        // ------------------------------------------------------------------

        static void Excel(string path, string[] columns, string[] types, long rows, Func<long, int, string[][]> fetch)
        {
            Kind[] kinds = Kinds(types);
            if (File.Exists(path)) File.Delete(path);
            using (FileStream file = new FileStream(path, FileMode.CreateNew))
            using (ZipArchive zip = new ZipArchive(file, ZipArchiveMode.Create))
            {
                Entry(zip, "[Content_Types].xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                    "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
                    "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                    "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
                    "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>" +
                    "</Types>");
                Entry(zip, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                    "</Relationships>");
                Entry(zip, "xl/workbook.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                    "<sheets><sheet name=\"DataFrame\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
                Entry(zip, "xl/_rels/workbook.xml.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                    "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" +
                    "</Relationships>");
                // Estilos: 0 padrão, 1 cabeçalho em negrito, 2 data (14), 3 data e hora (22)
                Entry(zip, "xl/styles.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
                    "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
                    "<fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font><font><b/><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>" +
                    "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>" +
                    "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>" +
                    "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>" +
                    "<cellXfs count=\"4\">" +
                    "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>" +
                    "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>" +
                    "<xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +
                    "<xf numFmtId=\"22\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>" +
                    "</cellXfs></styleSheet>");

                ZipArchiveEntry sheet = zip.CreateEntry("xl/worksheets/sheet1.xml", CompressionLevel.Optimal);
                using (Stream s = sheet.Open())
                using (XmlWriter w = XmlWriter.Create(s, new XmlWriterSettings { Encoding = new UTF8Encoding(false) }))
                {
                    const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                    w.WriteStartDocument(true);
                    w.WriteStartElement("worksheet", ns);
                    // Cabeçalho congelado
                    w.WriteStartElement("sheetViews", ns);
                    w.WriteStartElement("sheetView", ns);
                    w.WriteAttributeString("workbookViewId", "0");
                    w.WriteStartElement("pane", ns);
                    w.WriteAttributeString("ySplit", "1");
                    w.WriteAttributeString("topLeftCell", "A2");
                    w.WriteAttributeString("activePane", "bottomLeft");
                    w.WriteAttributeString("state", "frozen");
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndElement();
                    // Largura de cada coluna pelo cabeçalho e pelas primeiras linhas (senão as datas
                    // viram ########), entre 8 e 60 caracteres
                    string[][] sample = rows > 0 ? fetch(0, (int)Math.Min(200, rows)) : new string[0][];
                    w.WriteStartElement("cols", ns);
                    for (int c = 0; c < columns.Length; c++)
                    {
                        int chars = Math.Max(columns[c].Length, kinds[c] == Kind.Bool ? 10 : 0); // VERDADEIRO
                        foreach (string[] row in sample ?? new string[0][])
                            if (c < row.Length && row[c] != null) chars = Math.Max(chars, kinds[c] == Kind.DateTime ? 19 : kinds[c] == Kind.Date ? 10 : row[c].Length);
                        string min = (c + 1).ToString(CultureInfo.InvariantCulture);
                        w.WriteStartElement("col", ns);
                        w.WriteAttributeString("min", min);
                        w.WriteAttributeString("max", min);
                        w.WriteAttributeString("width", Math.Max(8, Math.Min(60, chars + 3)).ToString(CultureInfo.InvariantCulture));
                        w.WriteAttributeString("customWidth", "1");
                        w.WriteEndElement();
                    }
                    w.WriteEndElement();
                    w.WriteStartElement("sheetData", ns);

                    long r = 1;
                    w.WriteStartElement("row", ns);
                    w.WriteAttributeString("r", "1");
                    for (int c = 0; c < columns.Length; c++) TextCell(w, ns, Ref(c, r), columns[c], 1);
                    w.WriteEndElement();

                    ForEachRow(rows, fetch, row =>
                    {
                        r++;
                        w.WriteStartElement("row", ns);
                        w.WriteAttributeString("r", r.ToString(CultureInfo.InvariantCulture));
                        for (int c = 0; c < columns.Length; c++)
                        {
                            string v = c < row.Length ? row[c] : null;
                            if (v == null) continue; // NULL: célula vazia
                            string at = Ref(c, r);
                            double d;
                            DateTime dt;
                            if (kinds[c] == Kind.Number && Number(v, out d) && !double.IsNaN(d) && !double.IsInfinity(d))
                                ValueCell(w, ns, at, null, d.ToString("R", CultureInfo.InvariantCulture), 0);
                            else if (kinds[c] == Kind.Bool)
                                ValueCell(w, ns, at, "b", v == "True" ? "1" : "0", 0);
                            else if ((kinds[c] == Kind.Date || kinds[c] == Kind.DateTime) && Date(v, out dt) && dt.Year >= 1900)
                                ValueCell(w, ns, at, null, dt.ToOADate().ToString("R", CultureInfo.InvariantCulture), dt.TimeOfDay == TimeSpan.Zero && kinds[c] == Kind.Date ? 2 : 3);
                            else
                                TextCell(w, ns, at, v, 0);
                        }
                        w.WriteEndElement();
                    });

                    w.WriteEndElement(); // sheetData
                    w.WriteEndElement(); // worksheet
                    w.WriteEndDocument();
                }
            }
        }

        static void Entry(ZipArchive zip, string name, string content)
        {
            using (StreamWriter w = new StreamWriter(zip.CreateEntry(name, CompressionLevel.Optimal).Open(), new UTF8Encoding(false)))
                w.Write(content);
        }

        static void ValueCell(XmlWriter w, string ns, string at, string type, string value, int style)
        {
            w.WriteStartElement("c", ns);
            w.WriteAttributeString("r", at);
            if (type != null) w.WriteAttributeString("t", type);
            if (style != 0) w.WriteAttributeString("s", style.ToString(CultureInfo.InvariantCulture));
            w.WriteElementString("v", ns, value);
            w.WriteEndElement();
        }

        static void TextCell(XmlWriter w, string ns, string at, string text, int style)
        {
            w.WriteStartElement("c", ns);
            w.WriteAttributeString("r", at);
            w.WriteAttributeString("t", "inlineStr");
            if (style != 0) w.WriteAttributeString("s", style.ToString(CultureInfo.InvariantCulture));
            w.WriteStartElement("is", ns);
            w.WriteStartElement("t", ns);
            if (text.Length > 0 && (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1])))
                w.WriteAttributeString("xml", "space", null, "preserve");
            w.WriteString(XmlSafe(text.Length > 32767 ? text.Substring(0, 32767) : text));
            w.WriteEndElement();
            w.WriteEndElement();
            w.WriteEndElement();
        }

        // Caracteres de controle não são aceitos no XML
        static string XmlSafe(string s)
        {
            StringBuilder sb = null;
            for (int i = 0; i < s.Length; i++)
            {
                char ch = s[i];
                bool ok = ch == '\t' || ch == '\n' || ch == '\r' || ch >= 0x20;
                if (ok && sb == null) continue;
                if (sb == null) sb = new StringBuilder(s, 0, i, s.Length);
                if (ok) sb.Append(ch);
            }
            return sb == null ? s : sb.ToString();
        }

        // Coluna 0, linha 1 -> "A1"
        static string Ref(int column, long row)
        {
            string letters = "";
            for (int c = column + 1; c > 0; c = (c - 1) / 26) letters = (char)('A' + (c - 1) % 26) + letters;
            return letters + row.ToString(CultureInfo.InvariantCulture);
        }
    }
}

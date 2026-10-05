using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace SageTypes
{
    // DatetimeS: data e hora como o datetime do Python (com fuso opcional), no lugar do Date do VBA.
    //
    //   Dim d As New Sage.DatetimeS
    //   d = Now                                    ' membro padrão (Value): Date, texto ISO, outro DatetimeS
    //   Debug.Print d.AddMonths(1).StrFTime("%d/%m/%Y %H:%M")
    //   Set d = d.FromIsoFormat("2024-01-05T10:30:00Z")
    //   Debug.Print d.AsTimeZone("E. South America Standard Time").IsoFormat
    //
    // Os métodos não alteram o objeto: devolvem um DatetimeS novo (como no Python, datetime é
    // imutável). Sem fuso (TimeZone = "") a data é "ingênua", como no Python; com fuso, ela sabe
    // a diferença para o UTC e IsoFormat a mostra (+00:00).
    //
    // Interface dual: acrescente membros só no fim, com o próximo DispId (ver StringS).
    // Partes de uma data, para Trunc (as do DATETRUNC do SQL Server)
    [ComVisible(true), Guid("6CF3A3DD-8AFA-47E2-BBE6-22DD273E7AD8")]
    public enum sgDateParts
    {
        sgYear = 1,
        sgQuarter = 2,
        sgMonth = 3,
        sgWeek = 4,
        sgDay = 5,
        sgHour = 6,
        sgMinute = 7,
        sgSecond = 8,
        sgMillisecond = 9,
        sgMicrosecond = 10,
    }

    [ComVisible(true), Guid("EB5D4D4D-6B68-42B6-9A54-B169C43424CA"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface _DatetimeS
    {
        // Lê como Date do VBA; d = valor é o LetValue abaixo
        [DispId(0)] object Value { get; set; }

        // Partes
        [DispId(1)] int Year { get; }
        [DispId(2)] int Month { get; }
        [DispId(3)] int Day { get; }
        [DispId(4)] int Hour { get; }
        [DispId(5)] int Minute { get; }
        [DispId(6)] int Second { get; }
        [DispId(7)] int Microsecond { get; }
        [DispId(8)] int Weekday();
        [DispId(9)] int IsoWeekday();
        [DispId(10)] ListS IsoCalendar();
        [DispId(11)] int DayOfYear { get; }
        [DispId(12)] int Quarter { get; }
        [DispId(13)] int DaysInMonth { get; }
        [DispId(14)] bool IsLeapYear { get; }

        // Criar (devolvem um DatetimeS novo)
        [DispId(20)] DatetimeS Now([Optional] object TimeZone);
        [DispId(21)] DatetimeS Today();
        [DispId(22)] DatetimeS UtcNow();
        [DispId(23)] DatetimeS Create(int Year, int Month, int Day, [Optional] object Hour, [Optional] object Minute,
            [Optional] object Second, [Optional] object Microsecond, [Optional] object TimeZone);
        [DispId(24)] DatetimeS StrPTime(string Text, string Format);
        [DispId(25)] DatetimeS FromIsoFormat(string Text);
        [DispId(26)] DatetimeS FromTimestamp(double Seconds, [Optional] object TimeZone);
        [DispId(27)] DatetimeS Parse(object Text);

        // Texto e número
        [DispId(30)] StringS StrFTime(string Format);
        [DispId(31)] StringS IsoFormat([Optional] object Sep, [Optional] object TimeSpec);
        [DispId(32)] double Timestamp();
        [DispId(33)] string ToString();

        // Contas (devolvem um DatetimeS novo)
        [DispId(40)] DatetimeS Replace([Optional] object Year, [Optional] object Month, [Optional] object Day, [Optional] object Hour,
            [Optional] object Minute, [Optional] object Second, [Optional] object Microsecond, [Optional] object TimeZone);
        [DispId(41)] DatetimeS Add([Optional] object Days, [Optional] object Hours, [Optional] object Minutes, [Optional] object Seconds,
            [Optional] object Weeks, [Optional] object Months, [Optional] object Years, [Optional] object Milliseconds);
        [DispId(42)] DatetimeS AddMonths(int Months);
        [DispId(43)] double Diff(object Other, [Optional] object Unit);
        [DispId(44)] DatetimeS Normalize();

        // Fuso horário
        [DispId(50)] StringS TimeZone { get; }
        [DispId(51)] DatetimeS AsTimeZone([Optional] object TimeZone);
        [DispId(52)] object UtcOffset();

        [DispId(1000), PropertyLet("Value")] void LetValue(object Value);

        // DATETRUNC do SQL Server / date_trunc do DuckDB: início do ano, trimestre, mês, semana...
        [DispId(53)] DatetimeS Trunc(sgDateParts Part);
        // Dias úteis (segunda a sexta, menos os feriados): WORKDAY do Excel / busday_offset do numpy
        [DispId(54)] DatetimeS AddBusinessDays(int Days, [Optional] object Holidays);
        // Dias úteis entre esta data (inclusive) e a outra (exclusive), como o busday_count do numpy
        [DispId(55)] int BusinessDays(object Other, [Optional] object Holidays);
        // Último instante do período (end_time do pandas): o par do Trunc
        [DispId(56)] DatetimeS EndOf(sgDateParts Part);
    }

    [ComVisible(true), Guid("992489F3-7D1F-4D2F-B13A-05AB53F1A745"), ProgId("Sage.DatetimeS")]
    [ClassInterface(ClassInterfaceType.None), ComDefaultInterface(typeof(_DatetimeS))]
    public sealed class DatetimeS : _DatetimeS
    {
        DateTime value = DateTime.Today;   // relógio local do fuso (Kind Unspecified)
        TimeZoneInfo zone;                 // null: sem fuso ("ingênua")

        public DatetimeS() { }

        internal static DatetimeS From(DateTime value, TimeZoneInfo zone)
        {
            DatetimeS d = new DatetimeS();
            d.value = DateTime.SpecifyKind(value, DateTimeKind.Unspecified);
            d.zone = zone;
            return d;
        }

        internal DateTime Raw { get { return value; } }

        // ------------------------------------------------------------------
        // Valor
        // ------------------------------------------------------------------

        public object Value
        {
            get { return value; }
            set { LetValue(value); }
        }

        public void LetValue(object Value)
        {
            DatetimeS parsed = Coerce(Value);
            value = parsed.value;
            zone = parsed.zone;
        }

        // Date, texto (ISO ou no formato do Windows), número (data serial) ou outro DatetimeS
        internal static DatetimeS Coerce(object source)
        {
            DatetimeS other = source as DatetimeS;
            if (other != null) return From(other.value, other.zone);
            object v = Interop.Unwrap(source);
            if (v is DateTime) return From((DateTime)v, null);
            string text = v as string;
            if (text != null) return ParseText(text);
            if (PyOrder.IsNumber(v)) return From(DateTime.FromOADate(Convert.ToDouble(v, CultureInfo.InvariantCulture)), null);
            throw Interop.Error(13, "TypeError: não é possível converter " + Interop.TypeLabel(v) + " em data.");
        }

        // ------------------------------------------------------------------
        // Partes
        // ------------------------------------------------------------------

        public int Year { get { return value.Year; } }
        public int Month { get { return value.Month; } }
        public int Day { get { return value.Day; } }
        public int Hour { get { return value.Hour; } }
        public int Minute { get { return value.Minute; } }
        public int Second { get { return value.Second; } }
        public int Microsecond { get { return (int)(value.Ticks % TimeSpan.TicksPerSecond / 10); } }

        // Segunda = 0 ... domingo = 6, como no Python
        public int Weekday() { return ((int)value.DayOfWeek + 6) % 7; }
        public int IsoWeekday() { return Weekday() + 1; }

        // (ano ISO, semana ISO, dia da semana ISO), como tupla
        public ListS IsoCalendar()
        {
            int week = IsoWeek(value);
            int year = value.Year;
            if (week >= 52 && value.Month == 1) year--;
            else if (week == 1 && value.Month == 12) year++;
            ListS t = ListS.From(new object[] { year, week, IsoWeekday() });
            t.ArrayType = sgArrayTypes.sgTuple;
            return t;
        }

        static int IsoWeek(DateTime d)
        {
            DayOfWeek day = CultureInfo.InvariantCulture.Calendar.GetDayOfWeek(d);
            if (day >= DayOfWeek.Monday && day <= DayOfWeek.Wednesday) d = d.AddDays(3);
            return CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(d, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
        }

        public int DayOfYear { get { return value.DayOfYear; } }
        public int Quarter { get { return (value.Month - 1) / 3 + 1; } }
        public int DaysInMonth { get { return DateTime.DaysInMonth(value.Year, value.Month); } }
        public bool IsLeapYear { get { return DateTime.IsLeapYear(value.Year); } }

        // ------------------------------------------------------------------
        // Criar
        // ------------------------------------------------------------------

        // Sem fuso: hora local do Windows, sem fuso (datetime.now()); com fuso: a hora de lá
        public DatetimeS Now(object TimeZone)
        {
            if (Interop.IsMissing(TimeZone)) return From(DateTime.Now, null);
            TimeZoneInfo tz = Zone(TimeZone);
            return From(TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, tz), tz);
        }

        public DatetimeS Today() { return From(DateTime.Today, null); }

        // Com fuso UTC (o Python recomenda datetime.now(timezone.utc) no lugar do utcnow())
        public DatetimeS UtcNow() { return From(DateTime.UtcNow, TimeZoneInfo.Utc); }

        // datetime(ano, mês, dia, hora, minuto, segundo, microssegundo, tzinfo)
        public DatetimeS Create(int Year, int Month, int Day, object Hour, object Minute, object Second, object Microsecond, object TimeZone)
        {
            try
            {
                DateTime d = new DateTime(Year, Month, Day, Part(Hour), Part(Minute), Part(Second)).AddTicks(Part(Microsecond) * 10L);
                return From(d, Interop.IsMissing(TimeZone) ? null : Zone(TimeZone));
            }
            catch (ArgumentOutOfRangeException)
            {
                throw Interop.Error(5, "ValueError: data inválida.");
            }
        }

        static int Part(object value) { return Interop.IsMissing(value) ? 0 : Interop.Integer(value); }

        // ------------------------------------------------------------------
        // Texto -> data
        // ------------------------------------------------------------------

        // ISO 8601: 2024-01-05, 2024-01-05T10:30, 2024-01-05 10:30:00.123456+03:00, ...Z
        static readonly Regex Iso = new Regex(
            @"^\s*(\d{4})-?(\d{2})-?(\d{2})(?:[T ](\d{2})(?::?(\d{2})(?::?(\d{2})(?:[.,](\d{1,7}))?)?)?)?\s*(Z|[+-]\d{2}(?::?\d{2})?)?\s*$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        public DatetimeS FromIsoFormat(string Text)
        {
            DatetimeS d = TryIso(Text);
            if (d == null) throw Interop.Error(5, "ValueError: Invalid isoformat string: " + Interop.Repr(Text));
            return d;
        }

        static DatetimeS TryIso(string text)
        {
            Match m = Iso.Match(text ?? "");
            if (!m.Success) return null;
            try
            {
                DateTime d = new DateTime(Int(m.Groups[1]), Int(m.Groups[2]), Int(m.Groups[3]),
                    Int(m.Groups[4]), Int(m.Groups[5]), Int(m.Groups[6]));
                if (m.Groups[7].Success) d = d.AddTicks(long.Parse(m.Groups[7].Value.PadRight(7, '0'), CultureInfo.InvariantCulture));
                return From(d, m.Groups[8].Success ? Offset(m.Groups[8].Value) : null);
            }
            catch (ArgumentOutOfRangeException)
            {
                return null;
            }
        }

        static int Int(Group g) { return g.Success ? int.Parse(g.Value, CultureInfo.InvariantCulture) : 0; }

        // Como o dateutil.parser.parse: ISO, depois o formato do Windows (dd/mm/aaaa no pt-BR),
        // depois o formato invariável (inglês)
        public DatetimeS Parse(object Text) { return ParseText(Interop.Text(Text)); }

        static DatetimeS ParseText(string text)
        {
            DatetimeS iso = TryIso(text);
            if (iso != null) return iso;
            DateTime d;
            foreach (CultureInfo culture in new[] { CultureInfo.CurrentCulture, CultureInfo.InvariantCulture })
                if (DateTime.TryParse(text, culture, DateTimeStyles.AllowWhiteSpaces, out d)) return From(d, null);
            throw Interop.Error(5, "ValueError: data não reconhecida: " + Interop.Repr(text));
        }

        // Com os códigos do strftime do Python: %d/%m/%Y %H:%M:%S, %b (jan), %z (+0300)...
        public DatetimeS StrPTime(string Text, string Format)
        {
            bool hasZone;
            string[] formats = NetFormats(Format, out hasZone);
            string text = Text;
            if (hasZone) text = Regex.Replace(text, @"(?<=\d)\s*(Z|([+-])(\d{2}):?(\d{2}))\s*$",
                m => m.Groups[1].Value.ToUpperInvariant() == "Z" ? "+00:00" : m.Groups[2].Value + m.Groups[3].Value + ":" + m.Groups[4].Value);
            foreach (CultureInfo culture in new[] { CultureInfo.CurrentCulture, CultureInfo.InvariantCulture })
            {
                if (hasZone)
                {
                    DateTimeOffset o;
                    if (DateTimeOffset.TryParseExact(text, formats, culture, DateTimeStyles.AllowWhiteSpaces, out o))
                        return From(o.DateTime, FixedZone(o.Offset));
                }
                else
                {
                    DateTime d;
                    if (DateTime.TryParseExact(text, formats, culture, DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.NoCurrentDateDefault, out d))
                    {
                        // Sem data no formato: 1900-01-01, como no Python
                        if (d.Date == DateTime.MinValue.Date) d = new DateTime(1900, 1, 1) + d.TimeOfDay;
                        return From(d, null);
                    }
                }
            }
            throw Interop.Error(5, "ValueError: time data " + Interop.Repr(Text) + " does not match format " + Interop.Repr(Format));
        }

        // Formato do Python -> formatos do .NET (vários, porque %f aceita de 1 a 6 dígitos)
        static string[] NetFormats(string format, out bool hasZone)
        {
            hasZone = false;
            StringBuilder sb = new StringBuilder();
            bool fraction = false;
            for (int i = 0; i < format.Length; i++)
            {
                char c = format[i];
                if (c != '%' || i == format.Length - 1)
                {
                    sb.Append(Literal(c));
                    continue;
                }
                char code = format[++i];
                if (code == '-' && i < format.Length - 1) code = format[++i]; // %-d (sem zero à esquerda)
                switch (code)
                {
                    case 'Y': sb.Append("yyyy"); break;
                    case 'y': sb.Append("yy"); break;
                    case 'm': sb.Append("M"); break;      // 1 ou 2 dígitos, como o Python aceita
                    case 'd': sb.Append("d"); break;
                    case 'H': sb.Append("H"); break;
                    case 'I': sb.Append("h"); break;
                    case 'M': sb.Append("m"); break;
                    case 'S': sb.Append("s"); break;
                    case 'f': sb.Append("\u0001"); fraction = true; break;
                    case 'p': sb.Append("tt"); break;
                    case 'b': sb.Append("MMM"); break;
                    case 'B': sb.Append("MMMM"); break;
                    case 'a': sb.Append("ddd"); break;
                    case 'A': sb.Append("dddd"); break;
                    case 'z': sb.Append("zzz"); hasZone = true; break;
                    case '%': sb.Append("\\%"); break;
                    default: throw Interop.Error(5, "ValueError: código %" + code + " não suportado em StrPTime.");
                }
            }
            string net = sb.ToString();
            if (net.Length == 1) net = "%" + net; // "d" sozinho seria um formato padrão do .NET
            if (!fraction) return new[] { net };
            List<string> all = new List<string>();
            for (int digits = 1; digits <= 6; digits++) all.Add(net.Replace("\u0001", new string('f', digits)));
            return all.ToArray();
        }

        static string Literal(char c)
        {
            return char.IsLetterOrDigit(c) || "\\'\":/%.,-".IndexOf(c) >= 0 ? "\\" + c : c.ToString();
        }

        public DatetimeS FromTimestamp(double Seconds, object TimeZone)
        {
            DateTime utc = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddTicks((long)Math.Round(Seconds * TimeSpan.TicksPerSecond));
            if (Interop.IsMissing(TimeZone)) return From(utc.ToLocalTime(), null);
            TimeZoneInfo tz = Zone(TimeZone);
            return From(TimeZoneInfo.ConvertTimeFromUtc(utc, tz), tz);
        }

        // ------------------------------------------------------------------
        // Data -> texto
        // ------------------------------------------------------------------

        // Códigos do strftime do Python; nomes de mês e dia no idioma do Windows
        public StringS StrFTime(string Format)
        {
            CultureInfo culture = CultureInfo.CurrentCulture;
            DateTimeFormatInfo f = culture.DateTimeFormat;
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < Format.Length; i++)
            {
                char c = Format[i];
                if (c != '%' || i == Format.Length - 1) { sb.Append(c); continue; }
                char code = Format[++i];
                bool pad = true;
                if (code == '-' && i < Format.Length - 1) { pad = false; code = Format[++i]; }
                string two = pad ? "00" : "0";
                switch (code)
                {
                    case 'Y': sb.Append(value.Year.ToString("0000", CultureInfo.InvariantCulture)); break;
                    case 'y': sb.Append((value.Year % 100).ToString(two, CultureInfo.InvariantCulture)); break;
                    case 'm': sb.Append(value.Month.ToString(two, CultureInfo.InvariantCulture)); break;
                    case 'd': sb.Append(value.Day.ToString(two, CultureInfo.InvariantCulture)); break;
                    case 'H': sb.Append(value.Hour.ToString(two, CultureInfo.InvariantCulture)); break;
                    case 'I': sb.Append(((value.Hour + 11) % 12 + 1).ToString(two, CultureInfo.InvariantCulture)); break;
                    case 'M': sb.Append(value.Minute.ToString(two, CultureInfo.InvariantCulture)); break;
                    case 'S': sb.Append(value.Second.ToString(two, CultureInfo.InvariantCulture)); break;
                    case 'f': sb.Append(Microsecond.ToString("000000", CultureInfo.InvariantCulture)); break;
                    case 'p': sb.Append(value.Hour < 12 ? Designator(f.AMDesignator, "AM") : Designator(f.PMDesignator, "PM")); break;
                    case 'a': sb.Append(f.GetAbbreviatedDayName(value.DayOfWeek)); break;
                    case 'A': sb.Append(f.GetDayName(value.DayOfWeek)); break;
                    case 'b': sb.Append(f.GetAbbreviatedMonthName(value.Month)); break;
                    case 'B': sb.Append(f.GetMonthName(value.Month)); break;
                    case 'j': sb.Append(value.DayOfYear.ToString(pad ? "000" : "0", CultureInfo.InvariantCulture)); break;
                    case 'w': sb.Append((int)value.DayOfWeek); break;
                    case 'u': sb.Append(IsoWeekday()); break;
                    case 'U': sb.Append(WeekOfYear(DayOfWeek.Sunday).ToString(two, CultureInfo.InvariantCulture)); break;
                    case 'W': sb.Append(WeekOfYear(DayOfWeek.Monday).ToString(two, CultureInfo.InvariantCulture)); break;
                    case 'V': sb.Append(IsoWeek(value).ToString(two, CultureInfo.InvariantCulture)); break;
                    case 'G': sb.Append(IsoCalendar().Items[0]); break;
                    case 'c': sb.Append(value.ToString(f.ShortDatePattern + " " + f.LongTimePattern, culture)); break;
                    case 'x': sb.Append(value.ToString(f.ShortDatePattern, culture)); break;
                    case 'X': sb.Append(value.ToString(f.LongTimePattern, culture)); break;
                    case 'z': if (zone != null) sb.Append(OffsetText(Offset(), false)); break;
                    case 'Z': if (zone != null) sb.Append(ZoneName(zone)); break;
                    case '%': sb.Append('%'); break;
                    default: sb.Append('%').Append(code); break;
                }
            }
            return Wrap(sb.ToString());
        }

        static string Designator(string culture, string fallback) { return string.IsNullOrEmpty(culture) ? fallback : culture; }

        // %U / %W do Python: semanas que começam no domingo/segunda; dias antes do primeiro são a semana 0
        int WeekOfYear(DayOfWeek first)
        {
            int jan1 = (int)new DateTime(value.Year, 1, 1).DayOfWeek;
            int offset = ((int)first - jan1 + 7) % 7;  // dias até o primeiro "first" do ano
            int day = value.DayOfYear - 1;
            return day < offset ? 0 : (day - offset) / 7 + 1;
        }

        // 2024-01-05T10:30:00, com .ffffff se houver microssegundos e +00:00 se houver fuso
        public StringS IsoFormat(object Sep, object TimeSpec)
        {
            return Wrap(Iso8601(Interop.IsMissing(Sep) ? "T" : Interop.Text(Sep), Interop.IsMissing(TimeSpec) ? "auto" : Interop.Text(TimeSpec)));
        }

        internal string Iso8601(string sep, string timeSpec)
        {
            string time;
            switch (timeSpec.ToLowerInvariant())
            {
                case "auto": time = Microsecond != 0 ? "HH:mm:ss.ffffff" : "HH:mm:ss"; break;
                case "hours": time = "HH"; break;
                case "minutes": time = "HH:mm"; break;
                case "seconds": time = "HH:mm:ss"; break;
                case "milliseconds": time = "HH:mm:ss.fff"; break;
                case "microseconds": time = "HH:mm:ss.ffffff"; break;
                default: throw Interop.Error(5, "ValueError: Unknown timespec value");
            }
            string text = value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) + sep + value.ToString(time, CultureInfo.InvariantCulture);
            if (zone != null) text += OffsetText(Offset(), true);
            return text;
        }

        // Segundos desde 1970-01-01 UTC; sem fuso, a hora é a local do Windows
        public double Timestamp()
        {
            return (ToUtc() - new DateTime(1970, 1, 1)).TotalSeconds;
        }

        // str(datetime) do Python: 2024-01-05 10:30:00
        public override string ToString() { return Iso8601(" ", "auto"); }

        // ------------------------------------------------------------------
        // Contas
        // ------------------------------------------------------------------

        public DatetimeS Replace(object Year, object Month, object Day, object Hour, object Minute, object Second, object Microsecond, object TimeZone)
        {
            try
            {
                DateTime d = new DateTime(
                    Interop.IsMissing(Year) ? value.Year : Interop.Integer(Year),
                    Interop.IsMissing(Month) ? value.Month : Interop.Integer(Month),
                    Interop.IsMissing(Day) ? value.Day : Interop.Integer(Day),
                    Interop.IsMissing(Hour) ? value.Hour : Interop.Integer(Hour),
                    Interop.IsMissing(Minute) ? value.Minute : Interop.Integer(Minute),
                    Interop.IsMissing(Second) ? value.Second : Interop.Integer(Second));
                d = d.AddTicks((Interop.IsMissing(Microsecond) ? this.Microsecond : Interop.Integer(Microsecond)) * 10L);
                TimeZoneInfo tz = zone;
                if (!Interop.IsMissing(TimeZone)) tz = Interop.Text(TimeZone) == "" ? null : Zone(TimeZone);
                return From(d, tz);
            }
            catch (ArgumentOutOfRangeException)
            {
                throw Interop.Error(5, "ValueError: data inválida.");
            }
        }

        // timedelta (dias, horas...) e relativedelta (meses, anos: 31/01 + 1 mês = 29/02)
        public DatetimeS Add(object Days, object Hours, object Minutes, object Seconds, object Weeks, object Months, object Years, object Milliseconds)
        {
            DateTime d = value;
            int months = (Interop.IsMissing(Years) ? 0 : Interop.Integer(Years) * 12) + (Interop.IsMissing(Months) ? 0 : Interop.Integer(Months));
            if (months != 0) d = d.AddMonths(months);
            double days = Number(Days) + Number(Weeks) * 7 + Number(Hours) / 24 + Number(Minutes) / 1440 + Number(Seconds) / 86400 + Number(Milliseconds) / 86400000;
            d = d.AddTicks((long)Math.Round(days * TimeSpan.TicksPerDay));
            return From(d, zone);
        }

        static double Number(object value)
        {
            if (Interop.IsMissing(value)) return 0;
            object v = Interop.Unwrap(value);
            if (!PyOrder.IsNumber(v)) throw Interop.Error(13, "TypeError: esperado um número, não " + Interop.TypeLabel(v));
            return Convert.ToDouble(v, CultureInfo.InvariantCulture);
        }

        public DatetimeS AddMonths(int Months) { return From(value.AddMonths(Months), zone); }

        // Diferença Me - Other, em dias (padrão), "weeks", "hours", "minutes", "seconds" ou "milliseconds"
        public double Diff(object Other, object Unit)
        {
            DatetimeS other = Coerce(Other);
            TimeSpan span = ToUtc() - other.ToUtc();
            switch (Interop.IsMissing(Unit) ? "days" : Interop.Text(Unit).ToLowerInvariant())
            {
                case "days": case "d": return span.TotalDays;
                case "weeks": case "w": return span.TotalDays / 7;
                case "hours": case "h": return span.TotalHours;
                case "minutes": case "min": return span.TotalMinutes;
                case "seconds": case "s": return span.TotalSeconds;
                case "milliseconds": case "ms": return span.TotalMilliseconds;
                default: throw Interop.Error(5, "ValueError: unidade inválida: use days, weeks, hours, minutes, seconds ou milliseconds.");
            }
        }

        // Meia-noite do mesmo dia (Timestamp.normalize do pandas)
        public DatetimeS Normalize() { return From(value.Date, zone); }

        // Zera o que vem depois da parte pedida, como o DATETRUNC do SQL Server. A semana começa
        // na segunda-feira (ISO), como no date_trunc do DuckDB.
        public DatetimeS Trunc(sgDateParts Part)
        {
            DateTime d = value;
            switch (Part)
            {
                case sgDateParts.sgYear: return From(new DateTime(d.Year, 1, 1), zone);
                case sgDateParts.sgQuarter: return From(new DateTime(d.Year, (d.Month - 1) / 3 * 3 + 1, 1), zone);
                case sgDateParts.sgMonth: return From(new DateTime(d.Year, d.Month, 1), zone);
                case sgDateParts.sgWeek: return From(d.Date.AddDays(-Weekday()), zone);
                case sgDateParts.sgDay: return From(d.Date, zone);
                case sgDateParts.sgHour: return From(new DateTime(d.Year, d.Month, d.Day, d.Hour, 0, 0), zone);
                case sgDateParts.sgMinute: return From(new DateTime(d.Year, d.Month, d.Day, d.Hour, d.Minute, 0), zone);
                case sgDateParts.sgSecond: return From(new DateTime(d.Year, d.Month, d.Day, d.Hour, d.Minute, d.Second), zone);
                case sgDateParts.sgMillisecond: return From(new DateTime(d.Ticks - d.Ticks % TimeSpan.TicksPerMillisecond), zone);
                case sgDateParts.sgMicrosecond: return From(new DateTime(d.Ticks - d.Ticks % 10), zone);
                default: throw Interop.Error(5, "ValueError: parte inválida (" + (int)Part + "): use sgYear, sgQuarter, sgMonth, sgWeek, sgDay, sgHour, sgMinute, sgSecond, sgMillisecond ou sgMicrosecond.");
            }
        }

        // Pula sábados, domingos e os feriados (Array/ListS de datas); Days negativo volta. Com 0, a
        // própria data, como o WORKDAY do Excel. A hora fica a mesma.
        public DatetimeS AddBusinessDays(int Days, object Holidays)
        {
            HashSet<DateTime> holidays = HolidaySet(Holidays);
            DateTime d = value;
            int step = Days < 0 ? -1 : 1;
            for (int left = Math.Abs(Days); left > 0; )
            {
                d = d.AddDays(step);
                if (IsBusinessDay(d, holidays)) left--;
            }
            return From(d, zone);
        }

        // [esta data, Other): negativo quando Other vem antes. Só as datas contam (não as horas).
        public int BusinessDays(object Other, object Holidays)
        {
            HashSet<DateTime> holidays = HolidaySet(Holidays);
            DateTime start = value.Date, end = Coerce(Other).value.Date;
            int sign = 1;
            if (end < start) { DateTime t = start; start = end; end = t; sign = -1; }
            int count = 0;
            for (DateTime d = start; d < end; d = d.AddDays(1))
                if (IsBusinessDay(d, holidays)) count++;
            return sign * count;
        }

        static bool IsBusinessDay(DateTime d, HashSet<DateTime> holidays)
        {
            return d.DayOfWeek != DayOfWeek.Saturday && d.DayOfWeek != DayOfWeek.Sunday && !holidays.Contains(d.Date);
        }

        static HashSet<DateTime> HolidaySet(object holidays)
        {
            HashSet<DateTime> set = new HashSet<DateTime>();
            if (Interop.IsMissing(holidays) || holidays == null) return set;
            object raw = Interop.Unwrap(holidays);
            IEnumerable<object> items = raw is Array || holidays is ListS ? Interop.Items(holidays) : (IEnumerable<object>)new[] { holidays };
            foreach (object h in items)
                if (h != null && !(h is DBNull)) set.Add(Coerce(h).value.Date);
            return set;
        }

        public DatetimeS EndOf(sgDateParts Part)
        {
            DateTime start = Trunc(Part).value, next;
            switch (Part)
            {
                case sgDateParts.sgYear: next = start.AddYears(1); break;
                case sgDateParts.sgQuarter: next = start.AddMonths(3); break;
                case sgDateParts.sgMonth: next = start.AddMonths(1); break;
                case sgDateParts.sgWeek: next = start.AddDays(7); break;
                case sgDateParts.sgDay: next = start.AddDays(1); break;
                case sgDateParts.sgHour: next = start.AddHours(1); break;
                case sgDateParts.sgMinute: next = start.AddMinutes(1); break;
                case sgDateParts.sgSecond: next = start.AddSeconds(1); break;
                case sgDateParts.sgMillisecond: next = start.AddTicks(TimeSpan.TicksPerMillisecond); break;
                default: return From(start, zone); // microssegundo: já é o último instante
            }
            return From(next.AddTicks(-10), zone); // 1 microssegundo antes do próximo período
        }

        // ------------------------------------------------------------------
        // Fuso horário
        // ------------------------------------------------------------------

        public StringS TimeZone { get { return Wrap(zone == null ? "" : ZoneName(zone)); } }

        // Converte para outro fuso (padrão: o do Windows); sem fuso, a data é tida como local
        public DatetimeS AsTimeZone(object TimeZone)
        {
            TimeZoneInfo tz = Interop.IsMissing(TimeZone) ? TimeZoneInfo.Local : Zone(TimeZone);
            return From(TimeZoneInfo.ConvertTimeFromUtc(DateTime.SpecifyKind(ToUtc(), DateTimeKind.Utc), tz), tz);
        }

        // Diferença para o UTC em horas (-3 em Brasília); Empty sem fuso, como None no Python
        public object UtcOffset()
        {
            if (zone == null) return null;
            return Offset().TotalHours;
        }

        TimeSpan Offset() { return zone.GetUtcOffset(value); }

        DateTime ToUtc()
        {
            if (zone == null) return TimeZoneInfo.ConvertTimeToUtc(DateTime.SpecifyKind(value, DateTimeKind.Local));
            return value - Offset();
        }

        // "UTC", "local", deslocamento ("-03:00", "+0530") ou o nome do Windows
        // ("E. South America Standard Time"; lista: tzutil /l)
        static TimeZoneInfo Zone(object name)
        {
            string text = Interop.Text(name).Trim();
            switch (text.ToUpperInvariant())
            {
                case "UTC": case "Z": case "GMT": return TimeZoneInfo.Utc;
                case "LOCAL": return TimeZoneInfo.Local;
            }
            if (Regex.IsMatch(text, @"^[+-]\d{2}(:?\d{2})?$")) return Offset(text);
            try { return TimeZoneInfo.FindSystemTimeZoneById(text); }
            catch (TimeZoneNotFoundException)
            {
                throw Interop.Error(5, "ValueError: fuso desconhecido: " + Interop.Repr(text) +
                    " (use \"UTC\", \"local\", \"-03:00\" ou um nome do Windows, como \"E. South America Standard Time\").");
            }
        }

        static TimeZoneInfo Offset(string text)
        {
            if (text.ToUpperInvariant() == "Z") return TimeZoneInfo.Utc;
            int sign = text[0] == '-' ? -1 : 1;
            string digits = text.Substring(1).Replace(":", "");
            int hours = int.Parse(digits.Substring(0, 2), CultureInfo.InvariantCulture);
            int minutes = digits.Length >= 4 ? int.Parse(digits.Substring(2, 2), CultureInfo.InvariantCulture) : 0;
            return FixedZone(new TimeSpan(sign * hours, sign * minutes, 0));
        }

        static TimeZoneInfo FixedZone(TimeSpan offset)
        {
            if (offset == TimeSpan.Zero) return TimeZoneInfo.Utc;
            string name = OffsetText(offset, true);
            return TimeZoneInfo.CreateCustomTimeZone(name, offset, name, name);
        }

        static string ZoneName(TimeZoneInfo tz) { return tz == TimeZoneInfo.Utc ? "UTC" : tz.Id; }

        static string OffsetText(TimeSpan offset, bool colon)
        {
            string sign = offset < TimeSpan.Zero ? "-" : "+";
            offset = offset.Duration();
            return sign + offset.Hours.ToString("00", CultureInfo.InvariantCulture) + (colon ? ":" : "") + offset.Minutes.ToString("00", CultureInfo.InvariantCulture);
        }

        static StringS Wrap(string text) { return (StringS)Interop.Wrap(text); }
    }
}

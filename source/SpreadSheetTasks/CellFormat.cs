using System;

namespace SpreadSheetTasks
{
    /// <summary>
    /// A cell value paired with an Excel number format string. Pass a <see cref="FormattedCell"/>
    /// in a data source (DataTable, object[][], or List) to apply a per-cell number format when
    /// writing to Excel. The <see cref="Format"/> follows Excel's number-format syntax
    /// (e.g. <see cref="F.THOUSANDS_SEP"/>, <see cref="F.DATE_ISO"/>).
    /// </summary>
    public readonly struct FormattedCell
    {
        /// <summary>The cell's value. Stored as <see cref="object"/> to support any .NET type.</summary>
        public object Value { get; }

        /// <summary>The Excel number format string to apply to this cell.</summary>
        public string Format { get; }

        /// <summary>
        /// Creates a formatted cell with the supplied value and Excel number format.
        /// </summary>
        /// <param name="value">The cell's value.</param>
        /// <param name="format">The Excel number format string. Must not be null.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="format"/> is null.</exception>
        public FormattedCell(object value, string format)
        {
            Value = value;
            Format = format ?? throw new ArgumentNullException(nameof(format));
        }
    }

    /// <summary>
    /// A set of standard Excel number format string constants. Use these with
    /// <see cref="FormattedCell"/> or any API that accepts a raw Excel format string.
    /// </summary>
    public static class F
    {
        /// <summary>Thousands separator, e.g. <c>1,234,567</c>.</summary>
        public const string THOUSANDS_SEP = "#,##0";
        /// <summary>Polish złoty currency, e.g. <c>1,234.56 zł</c>.</summary>
        public const string CURRENCY_PLN = "#,##0.00 \"z\u0142\"";
        /// <summary>Euro currency, e.g. <c>1,234.56 €</c>.</summary>
        public const string CURRENCY_EUR = "#,##0.00 \u20AC";
        /// <summary>Percentage with no decimals, e.g. <c>25%</c>.</summary>
        public const string PERCENTAGE = "0%";
        /// <summary>Scientific notation, e.g. <c>1.23E+04</c>.</summary>
        public const string SCIENTIFIC = "0.00E+00";
        /// <summary>Two-decimal number with thousands separator, e.g. <c>1,234.56</c>.</summary>
        public const string TWO_DECIMALS = "#,##0.00";
        /// <summary>Treat the cell as plain text (no number formatting).</summary>
        public const string TEXT = "@";
        /// <summary>Nine-digit zero-padded number, e.g. <c>000123456</c>.</summary>
        public const string LEADING_ZEROS = "000000000";

        /// <summary>Short date, e.g. <c>02.06.2026</c>.</summary>
        public const string DATE_SHORT = "dd.mm.yyyy";
        /// <summary>Long date, e.g. <c>2 June 2026</c>.</summary>
        public const string DATE_LONG = "d mmmm yyyy";
        /// <summary>Day-month-year date, e.g. <c>02-06-2026</c>.</summary>
        public const string DATE_DAY_MONTH_YEAR = "dd-mm-yyyy";
        /// <summary>ISO 8601 date, e.g. <c>2026-06-02</c>.</summary>
        public const string DATE_ISO = "yyyy-mm-dd";
        /// <summary>Month and year, e.g. <c>June 2026</c>.</summary>
        public const string DATE_MONTH_YEAR = "mmmm yyyy";
        /// <summary>Full date with weekday, e.g. <c>Tuesday, 2 June 2026</c>.</summary>
        public const string DATE_WEEKDAY = "dddd, d mmmm yyyy";
        /// <summary>Day and month, e.g. <c>2 June</c>.</summary>
        public const string DATE_DAY_MONTH = "d mmmm";
        /// <summary>Year only, e.g. <c>2026</c>.</summary>
        public const string DATE_YEAR_ONLY = "yyyy";

        /// <summary>Short datetime, e.g. <c>02.06.2026 14:34</c>.</summary>
        public const string DATETIME_SHORT = "dd.mm.yyyy hh:mm";
        /// <summary>Long datetime, e.g. <c>2 June 2026 14:34:56</c>.</summary>
        public const string DATETIME_LONG = "d mmmm yyyy hh:mm:ss";
        /// <summary>24-hour time, e.g. <c>14:34</c>.</summary>
        public const string TIME_HH_MM = "hh:mm";
        /// <summary>24-hour time with seconds, e.g. <c>14:34:56</c>.</summary>
        public const string TIME_HH_MM_SS = "hh:mm:ss";
        /// <summary>12-hour time with AM/PM, e.g. <c>2:34 PM</c>.</summary>
        public const string TIME_12H = "h:mm AM/PM";
        /// <summary>24-hour datetime, e.g. <c>02.06.2026 14:34:56</c>.</summary>
        public const string DATETIME_24H = "dd.mm.yyyy hh:mm:ss";
        /// <summary>ISO 8601 datetime, e.g. <c>2026-06-02T14:34:56</c>.</summary>
        public const string DATETIME_ISO = "yyyy-mm-dd\"T\"hh:mm:ss";
        /// <summary>Time with milliseconds, e.g. <c>14:34:56.123</c>.</summary>
        public const string TIME_MS = "hh:mm:ss.000";

        /// <summary>Alias of <see cref="DATE_SHORT"/>.</summary>
        public const string SHORT_DATE = "dd.mm.yyyy";
        /// <summary>Alias of <see cref="DATE_LONG"/>.</summary>
        public const string LONG_DATE = "d mmmm yyyy";
        /// <summary>Alias of <see cref="DATE_ISO"/>.</summary>
        public const string ISO_DATE = "yyyy-mm-dd";
        /// <summary>Alias of <see cref="DATETIME_SHORT"/>.</summary>
        public const string SHORT_DATETIME = "dd.mm.yyyy hh:mm";
        /// <summary>Alias of <see cref="DATETIME_LONG"/>.</summary>
        public const string LONG_DATETIME = "d mmmm yyyy hh:mm:ss";
        /// <summary>Alias of <see cref="DATETIME_ISO"/>.</summary>
        public const string ISO_DATETIME = "yyyy-mm-dd\"T\"hh:mm:ss";
    }
}

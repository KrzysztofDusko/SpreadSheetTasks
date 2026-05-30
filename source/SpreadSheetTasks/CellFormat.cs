using System;

namespace SpreadSheetTasks
{
    public readonly struct FormattedCell
    {
        public object Value { get; }
        public string Format { get; }

        public FormattedCell(object value, string format)
        {
            Value = value;
            Format = format ?? throw new ArgumentNullException(nameof(format));
        }
    }

    public static class F
    {
        public const string THOUSANDS_SEP = "#,##0";
        public const string CURRENCY_PLN = "#,##0.00 \"z\u0142\"";
        public const string CURRENCY_EUR = "#,##0.00 \u20AC";
        public const string PERCENTAGE = "0%";
        public const string SCIENTIFIC = "0.00E+00";
        public const string TWO_DECIMALS = "#,##0.00";
        public const string TEXT = "@";
        public const string LEADING_ZEROS = "000000000";

        public const string DATE_SHORT = "dd.mm.yyyy";
        public const string DATE_LONG = "d mmmm yyyy";
        public const string DATE_DAY_MONTH_YEAR = "dd-mm-yyyy";
        public const string DATE_ISO = "yyyy-mm-dd";
        public const string DATE_MONTH_YEAR = "mmmm yyyy";
        public const string DATE_WEEKDAY = "dddd, d mmmm yyyy";
        public const string DATE_DAY_MONTH = "d mmmm";
        public const string DATE_YEAR_ONLY = "yyyy";

        public const string DATETIME_SHORT = "dd.mm.yyyy hh:mm";
        public const string DATETIME_LONG = "d mmmm yyyy hh:mm:ss";
        public const string TIME_HH_MM = "hh:mm";
        public const string TIME_HH_MM_SS = "hh:mm:ss";
        public const string TIME_12H = "h:mm AM/PM";
        public const string DATETIME_24H = "dd.mm.yyyy hh:mm:ss";
        public const string DATETIME_ISO = "yyyy-mm-dd\"T\"hh:mm:ss";
        public const string TIME_MS = "hh:mm:ss.000";

        public const string SHORT_DATE = "dd.mm.yyyy";
        public const string LONG_DATE = "d mmmm yyyy";
        public const string ISO_DATE = "yyyy-mm-dd";
        public const string SHORT_DATETIME = "dd.mm.yyyy hh:mm";
        public const string LONG_DATETIME = "d mmmm yyyy hh:mm:ss";
        public const string ISO_DATETIME = "yyyy-mm-dd\"T\"hh:mm:ss";
    }
}

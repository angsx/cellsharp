using System.Globalization;

namespace CellSharp.Internal;

internal static class CellValueConverter
{
    private const NumberStyles DecimalStyles = NumberStyles.AllowLeadingWhite
        | NumberStyles.AllowTrailingWhite
        | NumberStyles.AllowLeadingSign
        | NumberStyles.AllowDecimalPoint;

    private static readonly string[] DateTimeFormats =
    {
        "O",
        "yyyy-MM-dd",
        "yyyy-MM-dd'T'HH:mm:ss",
        "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF",
    };

    internal static bool Supports(Type type)
    {
        var valueType = Nullable.GetUnderlyingType(type) ?? type;

        return valueType == typeof(string)
            || valueType == typeof(byte)
            || valueType == typeof(sbyte)
            || valueType == typeof(short)
            || valueType == typeof(ushort)
            || valueType == typeof(int)
            || valueType == typeof(uint)
            || valueType == typeof(long)
            || valueType == typeof(ulong)
            || valueType == typeof(float)
            || valueType == typeof(double)
            || valueType == typeof(decimal)
            || valueType == typeof(bool)
            || valueType == typeof(DateTime)
            || valueType == typeof(Guid);
    }

    internal static bool TryConvert(
        string? value,
        Type targetType,
        bool isDateCell,
        CultureInfo culture,
        bool emptyStringAsNull,
        out object? converted,
        out ExcelReadErrorCode errorCode)
    {
        if (value is null || (emptyStringAsNull && value.Length == 0))
        {
            if (targetType == typeof(string) || Nullable.GetUnderlyingType(targetType) is not null)
            {
                converted = null;
                errorCode = default;
                return true;
            }

            converted = null;
            errorCode = ExcelReadErrorCode.RequiredValueMissing;
            return false;
        }

        var valueType = Nullable.GetUnderlyingType(targetType) ?? targetType;
        if (valueType == typeof(string))
        {
            converted = value;
            errorCode = default;
            return true;
        }

        if (valueType == typeof(bool))
        {
            if (value == "1" || value.Equals("true", StringComparison.OrdinalIgnoreCase))
            {
                converted = true;
                errorCode = default;
                return true;
            }

            if (value == "0" || value.Equals("false", StringComparison.OrdinalIgnoreCase))
            {
                converted = false;
                errorCode = default;
                return true;
            }

            converted = null;
            errorCode = ExcelReadErrorCode.InvalidValue;
            return false;
        }

        if (valueType == typeof(DateTime))
        {
            if (isDateCell && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var serial))
            {
                try
                {
                    converted = DateTime.FromOADate(serial);
                    errorCode = default;
                    return true;
                }
                catch (ArgumentException)
                {
                }
            }

            if (DateTime.TryParseExact(
                value,
                DateTimeFormats,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.RoundtripKind,
                out var dateTime))
            {
                converted = dateTime;
                errorCode = default;
                return true;
            }

            if (DateTime.TryParse(value, culture, DateTimeStyles.AllowWhiteSpaces, out dateTime))
            {
                converted = dateTime;
                errorCode = default;
                return true;
            }

            converted = null;
            errorCode = ExcelReadErrorCode.InvalidValue;
            return false;
        }

        if (valueType == typeof(Guid))
        {
            if (Guid.TryParse(value, out var guid))
            {
                converted = guid;
                errorCode = default;
                return true;
            }

            converted = null;
            errorCode = ExcelReadErrorCode.InvalidValue;
            return false;
        }

        if (TryConvertNumber(value, valueType, culture, out converted))
        {
            errorCode = default;
            return true;
        }

        errorCode = ExcelReadErrorCode.InvalidValue;
        return false;
    }

    private static bool TryConvertNumber(string value, Type type, CultureInfo culture, out object? converted)
    {
        if (type == typeof(byte) && TryParse<byte>(value, NumberStyles.Integer, culture, byte.TryParse, out var byteValue))
        {
            converted = byteValue;
            return true;
        }

        if (type == typeof(sbyte) && TryParse<sbyte>(value, NumberStyles.Integer, culture, sbyte.TryParse, out var sbyteValue))
        {
            converted = sbyteValue;
            return true;
        }

        if (type == typeof(short) && TryParse<short>(value, NumberStyles.Integer, culture, short.TryParse, out var shortValue))
        {
            converted = shortValue;
            return true;
        }

        if (type == typeof(ushort) && TryParse<ushort>(value, NumberStyles.Integer, culture, ushort.TryParse, out var ushortValue))
        {
            converted = ushortValue;
            return true;
        }

        if (type == typeof(int) && TryParse<int>(value, NumberStyles.Integer, culture, int.TryParse, out var intValue))
        {
            converted = intValue;
            return true;
        }

        if (type == typeof(uint) && TryParse<uint>(value, NumberStyles.Integer, culture, uint.TryParse, out var uintValue))
        {
            converted = uintValue;
            return true;
        }

        if (type == typeof(long) && TryParse<long>(value, NumberStyles.Integer, culture, long.TryParse, out var longValue))
        {
            converted = longValue;
            return true;
        }

        if (type == typeof(ulong) && TryParse<ulong>(value, NumberStyles.Integer, culture, ulong.TryParse, out var ulongValue))
        {
            converted = ulongValue;
            return true;
        }

        if (type == typeof(float) && TryParse<float>(value, NumberStyles.Float, culture, float.TryParse, out var floatValue))
        {
            converted = floatValue;
            return true;
        }

        if (type == typeof(double) && TryParse<double>(value, NumberStyles.Float, culture, double.TryParse, out var doubleValue))
        {
            converted = doubleValue;
            return true;
        }

        if (type == typeof(decimal) && TryParse<decimal>(value, DecimalStyles, culture, decimal.TryParse, out var decimalValue))
        {
            converted = decimalValue;
            return true;
        }

        converted = null;
        return false;
    }

    private delegate bool NumberParser<T>(string value, NumberStyles styles, IFormatProvider provider, out T parsed);

    private static bool TryParse<T>(
        string value,
        NumberStyles styles,
        CultureInfo culture,
        NumberParser<T> parser,
        out T parsed)
    {
        if (parser(value, styles, CultureInfo.InvariantCulture, out parsed))
        {
            return true;
        }

        return !string.Equals(culture.Name, CultureInfo.InvariantCulture.Name, StringComparison.Ordinal)
            && parser(value, styles, culture, out parsed);
    }
}

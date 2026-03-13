#ifndef FORMULAMETADATA_H
#define FORMULAMETADATA_H

#include <QString>
#include <QVector>
#include <QMap>

struct FormulaParamInfo {
    QString name;
    bool optional = false;
    QString description;
};

struct FormulaFuncInfo {
    QString name;
    QString description;
    QString syntax;
    QVector<FormulaParamInfo> params;
};

inline const QMap<QString, FormulaFuncInfo>& formulaRegistry() {
    static const QMap<QString, FormulaFuncInfo> reg = {
        // === Aggregate ===
        {"SUM", {"SUM", "Adds all the numbers in a range of cells",
            "SUM(number1, [number2], ...)",
            {{"number1", false, "The first number, cell reference, or range to add"},
             {"number2", true, "Additional numbers, references, or ranges to add"}}}},
        {"AVERAGE", {"AVERAGE", "Returns the average (arithmetic mean) of its arguments",
            "AVERAGE(number1, [number2], ...)",
            {{"number1", false, "The first number, cell reference, or range for the average"},
             {"number2", true, "Additional numbers, references, or ranges"}}}},
        {"COUNT", {"COUNT", "Counts the number of cells that contain numbers",
            "COUNT(value1, [value2], ...)",
            {{"value1", false, "The first cell reference or range to count numbers in"},
             {"value2", true, "Additional references or ranges to count"}}}},
        {"COUNTA", {"COUNTA", "Counts the number of non-empty cells in a range",
            "COUNTA(value1, [value2], ...)",
            {{"value1", false, "The first cell reference or range to count non-empty cells"},
             {"value2", true, "Additional references or ranges to count"}}}},
        {"MIN", {"MIN", "Returns the smallest value in a set of values",
            "MIN(number1, [number2], ...)",
            {{"number1", false, "The first number, cell reference, or range to find the minimum"},
             {"number2", true, "Additional numbers, references, or ranges"}}}},
        {"MAX", {"MAX", "Returns the largest value in a set of values",
            "MAX(number1, [number2], ...)",
            {{"number1", false, "The first number, cell reference, or range to find the maximum"},
             {"number2", true, "Additional numbers, references, or ranges"}}}},

        // === Conditional ===
        {"IF", {"IF", "Checks a condition and returns one value if TRUE, another if FALSE",
            "IF(logical_test, value_if_true, [value_if_false])",
            {{"logical_test", false, "The condition to evaluate (e.g., A1>10)"},
             {"value_if_true", false, "The value returned when the condition is TRUE"},
             {"value_if_false", true, "The value returned when the condition is FALSE. Defaults to FALSE"}}}},
        {"IFERROR", {"IFERROR", "Returns a value if no error; otherwise returns an alternate value",
            "IFERROR(value, value_if_error)",
            {{"value", false, "The value or formula to check for an error"},
             {"value_if_error", false, "The value to return if the first argument produces an error"}}}},
        {"AND", {"AND", "Returns TRUE if all arguments evaluate to TRUE",
            "AND(logical1, [logical2], ...)",
            {{"logical1", false, "The first condition to evaluate"},
             {"logical2", true, "Additional conditions. All must be TRUE for AND to return TRUE"}}}},
        {"OR", {"OR", "Returns TRUE if any argument evaluates to TRUE",
            "OR(logical1, [logical2], ...)",
            {{"logical1", false, "The first condition to evaluate"},
             {"logical2", true, "Additional conditions. At least one must be TRUE"}}}},
        {"NOT", {"NOT", "Reverses the logic of its argument",
            "NOT(logical)",
            {{"logical", false, "A value or expression that evaluates to TRUE or FALSE"}}}},

        // === String ===
        {"CONCAT", {"CONCAT", "Joins several text strings into one text string",
            "CONCAT(text1, [text2], ...)",
            {{"text1", false, "The first text string, cell reference, or range to join"},
             {"text2", true, "Additional text strings to join"}}}},
        {"CONCATENATE", {"CONCATENATE", "Joins several text strings into one text string",
            "CONCATENATE(text1, [text2], ...)",
            {{"text1", false, "The first text string to join"},
             {"text2", true, "Additional text strings to join (up to 255)"}}}},
        {"LEN", {"LEN", "Returns the number of characters in a text string",
            "LEN(text)",
            {{"text", false, "The text string whose length you want to find"}}}},
        {"UPPER", {"UPPER", "Converts all characters in a text string to uppercase",
            "UPPER(text)",
            {{"text", false, "The text to convert to uppercase"}}}},
        {"LOWER", {"LOWER", "Converts all characters in a text string to lowercase",
            "LOWER(text)",
            {{"text", false, "The text to convert to lowercase"}}}},
        {"TRIM", {"TRIM", "Removes extra spaces from text, leaving single spaces between words",
            "TRIM(text)",
            {{"text", false, "The text from which to remove extra spaces"}}}},
        {"LEFT", {"LEFT", "Returns the specified number of characters from the start of a text string",
            "LEFT(text, [num_chars])",
            {{"text", false, "The text string from which to extract characters"},
             {"num_chars", true, "Number of characters to extract. Defaults to 1"}}}},
        {"RIGHT", {"RIGHT", "Returns the specified number of characters from the end of a text string",
            "RIGHT(text, [num_chars])",
            {{"text", false, "The text string from which to extract characters"},
             {"num_chars", true, "Number of characters to extract. Defaults to 1"}}}},
        {"MID", {"MID", "Returns characters from the middle of a text string, given a starting position and length",
            "MID(text, start_num, num_chars)",
            {{"text", false, "The text string containing the characters to extract"},
             {"start_num", false, "The position of the first character to extract (1-based)"},
             {"num_chars", false, "The number of characters to extract"}}}},
        {"FIND", {"FIND", "Finds one text string within another (case-sensitive). Returns the starting position",
            "FIND(find_text, within_text, [start_num])",
            {{"find_text", false, "The text to find"},
             {"within_text", false, "The text to search within"},
             {"start_num", true, "Position to start searching from. Defaults to 1"}}}},
        {"SEARCH", {"SEARCH", "Finds one text string within another (case-insensitive). Supports wildcards",
            "SEARCH(find_text, within_text, [start_num])",
            {{"find_text", false, "The text to find. Supports ? and * wildcards"},
             {"within_text", false, "The text to search within"},
             {"start_num", true, "Position to start searching from. Defaults to 1"}}}},
        {"SUBSTITUTE", {"SUBSTITUTE", "Replaces occurrences of old text with new text in a string",
            "SUBSTITUTE(text, old_text, new_text, [instance_num])",
            {{"text", false, "The text or cell reference containing text to replace"},
             {"old_text", false, "The text to find and replace"},
             {"new_text", false, "The replacement text"},
             {"instance_num", true, "Which occurrence to replace. If omitted, replaces all"}}}},
        {"TEXT", {"TEXT", "Converts a number to text in a specified number format",
            "TEXT(value, format_text)",
            {{"value", false, "The number to format"},
             {"format_text", false, "The format pattern (e.g., \"$#,##0.00\", \"MM/DD/YYYY\")"}}}},
        {"PROPER", {"PROPER", "Capitalizes the first letter of each word in a text string",
            "PROPER(text)",
            {{"text", false, "The text to capitalize"}}}},
        {"REPT", {"REPT", "Repeats text a given number of times",
            "REPT(text, number_times)",
            {{"text", false, "The text to repeat"},
             {"number_times", false, "The number of times to repeat the text"}}}},
        {"EXACT", {"EXACT", "Checks whether two text strings are exactly identical (case-sensitive)",
            "EXACT(text1, text2)",
            {{"text1", false, "The first text string to compare"},
             {"text2", false, "The second text string to compare"}}}},
        {"VALUE", {"VALUE", "Converts a text string that represents a number to a numeric value",
            "VALUE(text)",
            {{"text", false, "The text enclosed in quotes or a cell reference containing the text to convert"}}}},

        // === Math ===
        {"ROUND", {"ROUND", "Rounds a number to a specified number of decimal places",
            "ROUND(number, num_digits)",
            {{"number", false, "The number to round"},
             {"num_digits", false, "Number of decimal places. Negative values round to left of decimal"}}}},
        {"ROUNDUP", {"ROUNDUP", "Rounds a number up (away from zero) to a specified number of digits",
            "ROUNDUP(number, num_digits)",
            {{"number", false, "The number to round up"},
             {"num_digits", false, "Number of decimal places to round up to"}}}},
        {"ROUNDDOWN", {"ROUNDDOWN", "Rounds a number down (toward zero) to a specified number of digits",
            "ROUNDDOWN(number, num_digits)",
            {{"number", false, "The number to round down"},
             {"num_digits", false, "Number of decimal places to round down to"}}}},
        {"ABS", {"ABS", "Returns the absolute value of a number (number without its sign)",
            "ABS(number)",
            {{"number", false, "The number whose absolute value you want"}}}},
        {"SQRT", {"SQRT", "Returns the positive square root of a number",
            "SQRT(number)",
            {{"number", false, "The number to find the square root of. Must be >= 0"}}}},
        {"POWER", {"POWER", "Returns the result of a number raised to a power",
            "POWER(number, power)",
            {{"number", false, "The base number"},
             {"power", false, "The exponent to raise the base number to"}}}},
        {"MOD", {"MOD", "Returns the remainder after a number is divided by a divisor",
            "MOD(number, divisor)",
            {{"number", false, "The number to divide"},
             {"divisor", false, "The number to divide by"}}}},
        {"INT", {"INT", "Rounds a number down to the nearest integer",
            "INT(number)",
            {{"number", false, "The number to round down to the nearest integer"}}}},
        {"CEILING", {"CEILING", "Rounds a number up to the nearest specified multiple",
            "CEILING(number, significance)",
            {{"number", false, "The number to round up"},
             {"significance", false, "The multiple to round up to"}}}},
        {"FLOOR", {"FLOOR", "Rounds a number down to the nearest specified multiple",
            "FLOOR(number, significance)",
            {{"number", false, "The number to round down"},
             {"significance", false, "The multiple to round down to"}}}},
        {"LOG", {"LOG", "Returns the logarithm of a number to a specified base",
            "LOG(number, [base])",
            {{"number", false, "The positive number to calculate the logarithm for"},
             {"base", true, "The base of the logarithm. Defaults to 10"}}}},
        {"LN", {"LN", "Returns the natural logarithm (base e) of a number",
            "LN(number)",
            {{"number", false, "The positive number to calculate the natural log for"}}}},
        {"EXP", {"EXP", "Returns e (2.71828...) raised to the power of a given number",
            "EXP(number)",
            {{"number", false, "The exponent applied to the base e"}}}},
        {"RAND", {"RAND", "Returns a random decimal number between 0 and 1. Recalculates on each change",
            "RAND()", {}}},
        {"RANDBETWEEN", {"RANDBETWEEN", "Returns a random integer between two specified values",
            "RANDBETWEEN(bottom, top)",
            {{"bottom", false, "The smallest integer the function can return"},
             {"top", false, "The largest integer the function can return"}}}},

        // === Statistical ===
        {"COUNTIF", {"COUNTIF", "Counts the number of cells in a range that meet a given condition",
            "COUNTIF(range, criteria)",
            {{"range", false, "The range of cells to evaluate"},
             {"criteria", false, "The condition that determines which cells to count (e.g., \">10\", \"Apple\")"}}}},
        {"SUMIF", {"SUMIF", "Adds the values in a range of cells that meet a specified condition",
            "SUMIF(range, criteria, [sum_range])",
            {{"range", false, "The range of cells to evaluate against the criteria"},
             {"criteria", false, "The condition that determines which cells to sum (e.g., \">10\")"},
             {"sum_range", true, "The cells to sum. If omitted, the range cells are summed"}}}},
        {"AVERAGEIF", {"AVERAGEIF", "Returns the average of cells in a range that meet a specified condition",
            "AVERAGEIF(range, criteria, [average_range])",
            {{"range", false, "The range of cells to evaluate against the criteria"},
             {"criteria", false, "The condition that determines which cells to average"},
             {"average_range", true, "The cells to average. If omitted, uses range"}}}},
        {"COUNTBLANK", {"COUNTBLANK", "Counts the number of empty cells in a specified range",
            "COUNTBLANK(range)",
            {{"range", false, "The range in which to count blank cells"}}}},
        {"SUMPRODUCT", {"SUMPRODUCT", "Multiplies corresponding elements in arrays, then returns the sum of those products",
            "SUMPRODUCT(array1, [array2], ...)",
            {{"array1", false, "The first array or range whose elements you want to multiply then add"},
             {"array2", true, "Additional arrays or ranges. Must have the same dimensions"}}}},
        {"MEDIAN", {"MEDIAN", "Returns the median (middle value) of the given numbers",
            "MEDIAN(number1, [number2], ...)",
            {{"number1", false, "The first number, cell reference, or range"},
             {"number2", true, "Additional numbers, references, or ranges"}}}},
        {"MODE", {"MODE", "Returns the most frequently occurring value in a data set",
            "MODE(number1, [number2], ...)",
            {{"number1", false, "The first number, cell reference, or range"},
             {"number2", true, "Additional numbers, references, or ranges"}}}},
        {"STDEV", {"STDEV", "Estimates standard deviation based on a sample (ignores text and logicals)",
            "STDEV(number1, [number2], ...)",
            {{"number1", false, "The first number, cell reference, or range of the sample"},
             {"number2", true, "Additional numbers, references, or ranges"}}}},
        {"VAR", {"VAR", "Estimates variance based on a sample (ignores text and logicals)",
            "VAR(number1, [number2], ...)",
            {{"number1", false, "The first number, cell reference, or range of the sample"},
             {"number2", true, "Additional numbers, references, or ranges"}}}},
        {"LARGE", {"LARGE", "Returns the k-th largest value in a data set",
            "LARGE(array, k)",
            {{"array", false, "The array or range of data"},
             {"k", false, "The position (from largest) in the array to return. 1 = largest"}}}},
        {"SMALL", {"SMALL", "Returns the k-th smallest value in a data set",
            "SMALL(array, k)",
            {{"array", false, "The array or range of data"},
             {"k", false, "The position (from smallest) in the array to return. 1 = smallest"}}}},
        {"RANK", {"RANK", "Returns the rank of a number in a list of numbers",
            "RANK(number, ref, [order])",
            {{"number", false, "The number whose rank you want to find"},
             {"ref", false, "The array or range of numbers to rank against"},
             {"order", true, "0 or omitted = descending (largest is rank 1). Non-zero = ascending"}}}},
        {"PERCENTILE", {"PERCENTILE", "Returns the k-th percentile of values in a range",
            "PERCENTILE(array, k)",
            {{"array", false, "The array or range of data"},
             {"k", false, "The percentile value between 0 and 1 (e.g., 0.25 for 25th percentile)"}}}},

        // === Date/Time ===
        {"NOW", {"NOW", "Returns the current date and time. Updates on recalculation",
            "NOW()", {}}},
        {"TODAY", {"TODAY", "Returns the current date (without time). Updates daily",
            "TODAY()", {}}},
        {"YEAR", {"YEAR", "Returns the year component of a date as a 4-digit number",
            "YEAR(serial_number)",
            {{"serial_number", false, "The date from which to extract the year"}}}},
        {"MONTH", {"MONTH", "Returns the month component of a date (1-12)",
            "MONTH(serial_number)",
            {{"serial_number", false, "The date from which to extract the month"}}}},
        {"DAY", {"DAY", "Returns the day component of a date (1-31)",
            "DAY(serial_number)",
            {{"serial_number", false, "The date from which to extract the day"}}}},
        {"DATE", {"DATE", "Creates a date value from individual year, month, and day components",
            "DATE(year, month, day)",
            {{"year", false, "The year (1900-9999)"},
             {"month", false, "The month (1-12). Values > 12 roll into the next year"},
             {"day", false, "The day (1-31). Values > days in month roll into the next month"}}}},
        {"HOUR", {"HOUR", "Returns the hour component of a time value (0-23)",
            "HOUR(serial_number)",
            {{"serial_number", false, "The time value from which to extract the hour"}}}},
        {"MINUTE", {"MINUTE", "Returns the minute component of a time value (0-59)",
            "MINUTE(serial_number)",
            {{"serial_number", false, "The time value from which to extract the minutes"}}}},
        {"SECOND", {"SECOND", "Returns the second component of a time value (0-59)",
            "SECOND(serial_number)",
            {{"serial_number", false, "The time value from which to extract the seconds"}}}},
        {"DATEDIF", {"DATEDIF", "Calculates the difference between two dates in days, months, or years",
            "DATEDIF(start_date, end_date, unit)",
            {{"start_date", false, "The start date"},
             {"end_date", false, "The end date. Must be >= start_date"},
             {"unit", false, "The unit of time: \"Y\" (years), \"M\" (months), \"D\" (days)"}}}},
        {"NETWORKDAYS", {"NETWORKDAYS", "Returns the number of whole working days between two dates (excludes weekends)",
            "NETWORKDAYS(start_date, end_date, [holidays])",
            {{"start_date", false, "The start date"},
             {"end_date", false, "The end date"},
             {"holidays", true, "A range of dates to exclude as holidays"}}}},
        {"WEEKDAY", {"WEEKDAY", "Returns the day of the week corresponding to a date (1-7)",
            "WEEKDAY(serial_number, [return_type])",
            {{"serial_number", false, "The date to find the day of the week for"},
             {"return_type", true, "1 = Sun(1)-Sat(7), 2 = Mon(1)-Sun(7), 3 = Mon(0)-Sun(6)"}}}},
        {"EDATE", {"EDATE", "Returns a date that is a specified number of months before or after a given date",
            "EDATE(start_date, months)",
            {{"start_date", false, "The starting date"},
             {"months", false, "Number of months before (negative) or after (positive) start_date"}}}},
        {"EOMONTH", {"EOMONTH", "Returns the last day of the month that is a specified number of months before or after a date",
            "EOMONTH(start_date, months)",
            {{"start_date", false, "The starting date"},
             {"months", false, "Number of months before (negative) or after (positive) start_date"}}}},
        {"DATEVALUE", {"DATEVALUE", "Converts a date stored as text to a date serial number",
            "DATEVALUE(date_text)",
            {{"date_text", false, "Text representing a date (e.g., \"1/15/2024\", \"January 15, 2024\")"}}}},

        // === Lookup ===
        {"VLOOKUP", {"VLOOKUP", "Searches for a value in the first column of a table and returns a value in the same row from another column",
            "VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])",
            {{"lookup_value", false, "The value to search for in the first column of the table"},
             {"table_array", false, "The range of cells that contains the data (e.g., A1:D10)"},
             {"col_index_num", false, "The column number in the table from which to return a value (1-based)"},
             {"range_lookup", true, "TRUE = approximate match (default, data must be sorted). FALSE = exact match"}}}},
        {"HLOOKUP", {"HLOOKUP", "Searches for a value in the first row of a table and returns a value in the same column from another row",
            "HLOOKUP(lookup_value, table_array, row_index_num, [range_lookup])",
            {{"lookup_value", false, "The value to search for in the first row of the table"},
             {"table_array", false, "The range of cells that contains the data"},
             {"row_index_num", false, "The row number in the table from which to return a value (1-based)"},
             {"range_lookup", true, "TRUE = approximate match (default). FALSE = exact match"}}}},
        {"XLOOKUP", {"XLOOKUP", "Searches a range for a match and returns the corresponding item from a second range. More flexible than VLOOKUP",
            "XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode])",
            {{"lookup_value", false, "The value to search for"},
             {"lookup_array", false, "The range to search in"},
             {"return_array", false, "The range to return a value from"},
             {"if_not_found", true, "Value to return if no match is found. Default is #N/A"},
             {"match_mode", true, "0 = exact match (default), -1 = exact or next smaller, 1 = exact or next larger"}}}},
        {"INDEX", {"INDEX", "Returns the value of a cell at the intersection of a given row and column in a range",
            "INDEX(array, row_num, [column_num])",
            {{"array", false, "The range of cells or array constant"},
             {"row_num", false, "The row number in the array from which to return a value"},
             {"column_num", true, "The column number in the array. Required for multi-column ranges"}}}},
        {"MATCH", {"MATCH", "Returns the relative position of a value in a range that matches a specified value",
            "MATCH(lookup_value, lookup_array, [match_type])",
            {{"lookup_value", false, "The value to find in the lookup array"},
             {"lookup_array", false, "The range of cells to search (single row or column)"},
             {"match_type", true, "1 = largest value <= lookup (default, sorted asc). 0 = exact match. -1 = smallest value >= lookup"}}}},

        // === Info ===
        {"ISBLANK", {"ISBLANK", "Returns TRUE if the referenced cell is empty, FALSE otherwise",
            "ISBLANK(value)",
            {{"value", false, "The cell reference to check for emptiness"}}}},
        {"ISERROR", {"ISERROR", "Returns TRUE if a value is any error (#N/A, #VALUE!, #REF!, etc.)",
            "ISERROR(value)",
            {{"value", false, "The value or formula to check for an error"}}}},
        {"ISNUMBER", {"ISNUMBER", "Returns TRUE if a value is a number, FALSE otherwise",
            "ISNUMBER(value)",
            {{"value", false, "The value to test whether it is a number"}}}},
        {"ISTEXT", {"ISTEXT", "Returns TRUE if a value is text, FALSE otherwise",
            "ISTEXT(value)",
            {{"value", false, "The value to test whether it is text"}}}},
        {"CHOOSE", {"CHOOSE", "Returns a value from a list of values based on an index number",
            "CHOOSE(index_num, value1, [value2], ...)",
            {{"index_num", false, "The index number (1-based) specifying which value to return"},
             {"value1", false, "The first value to choose from"},
             {"value2", true, "Additional values (up to 254)"}}}},
        {"SWITCH", {"SWITCH", "Evaluates an expression against a list of value/result pairs and returns the first match",
            "SWITCH(expression, value1, result1, [value2, result2], ..., [default])",
            {{"expression", false, "The value or expression to evaluate"},
             {"value1", false, "The first value to compare against the expression"},
             {"result1", false, "The result to return if expression matches value1"},
             {"default", true, "The default result if no values match"}}}},

        // === Dynamic Array ===
        {"FILTER", {"FILTER", "Filters an array based on a Boolean array",
            "FILTER(array, include, [if_empty])",
            {{"array", false, "The range or array to filter"},
             {"include", false, "A Boolean array whose height or width is the same as the array"},
             {"if_empty", true, "The value to return if all values in the include array are FALSE"}}}},
        {"SORT", {"SORT", "Sorts the contents of a range or array",
            "SORT(array, [sort_index], [sort_order], [by_col])",
            {{"array", false, "The range or array to sort"},
             {"sort_index", true, "A number indicating the row or column to sort by (default 1)"},
             {"sort_order", true, "1 for ascending (default), -1 for descending"},
             {"by_col", true, "FALSE to sort by row (default), TRUE to sort by column"}}}},
        {"UNIQUE", {"UNIQUE", "Returns a list of unique values from a range or array",
            "UNIQUE(array, [by_col], [exactly_once])",
            {{"array", false, "The range or array from which to return unique values"},
             {"by_col", true, "FALSE for unique rows (default), TRUE for unique columns"},
             {"exactly_once", true, "FALSE for distinct (default), TRUE for values appearing exactly once"}}}},
        {"SEQUENCE", {"SEQUENCE", "Generates an array of sequential numbers",
            "SEQUENCE(rows, [columns], [start], [step])",
            {{"rows", false, "The number of rows to return"},
             {"columns", true, "The number of columns to return (default 1)"},
             {"start", true, "The first number in the sequence (default 1)"},
             {"step", true, "The amount to increment each subsequent value (default 1)"}}}},
    };
    return reg;
}

// Helper: find which function the cursor is inside, and which parameter index
struct FormulaContext {
    QString funcName;
    int paramIndex = -1;  // 0-based, -1 if not inside a function
};

inline FormulaContext findFormulaContext(const QString& text, int cursorPos) {
    FormulaContext ctx;
    if (!text.startsWith("=")) return ctx;

    int depth = 0;
    int parenPos = -1;
    for (int i = cursorPos - 1; i >= 1; --i) {
        QChar ch = text[i];
        if (ch == ')') depth++;
        else if (ch == '(') {
            if (depth == 0) { parenPos = i; break; }
            depth--;
        }
    }
    if (parenPos < 0) return ctx;

    int nameEnd = parenPos;
    int nameStart = nameEnd - 1;
    while (nameStart >= 1 && (text[nameStart].isLetterOrNumber() || text[nameStart] == '_'))
        nameStart--;
    nameStart++;
    if (nameStart >= nameEnd) return ctx;

    ctx.funcName = text.mid(nameStart, nameEnd - nameStart).toUpper();

    int commas = 0;
    int d = 0;
    for (int i = parenPos + 1; i < cursorPos && i < text.length(); ++i) {
        QChar ch = text[i];
        if (ch == '(') d++;
        else if (ch == ')') d--;
        else if (ch == ',' && d == 0) commas++;
    }
    ctx.paramIndex = commas;
    return ctx;
}

// Build rich-text syntax string with current param highlighted
inline QString buildParamHintHtml(const FormulaFuncInfo& info, int paramIndex) {
    QString html = "<span style='color:#333; font-size:12px;'>";
    html += "<b>" + info.name + "</b>(";
    for (int i = 0; i < info.params.size(); ++i) {
        if (i > 0) html += ", ";
        const auto& p = info.params[i];
        QString pName = p.optional ? "[" + p.name + "]" : p.name;
        if (i == paramIndex) {
            html += "<b style='color:#0078D4;'>" + pName + "</b>";
        } else {
            html += "<span style='color:#888;'>" + pName + "</span>";
        }
    }
    if (!info.params.isEmpty() && info.params.last().optional) {
        html += ", ...";
    }
    html += ")</span>";
    return html;
}

// Build detailed HTML for the function detail panel
inline QString buildDetailHtml(const FormulaFuncInfo& info) {
    QString html;
    html += "<table cellspacing='0' cellpadding='0' style='font-family: -apple-system, SF Pro Text, Helvetica;'>";

    // Function name
    html += "<tr><td style='padding: 0 0 6px 0;'>"
            "<span style='font-size:16px; font-weight:bold; color:#1A73E8;'>" + info.name + "</span>"
            "</td></tr>";

    // Description
    html += "<tr><td style='padding: 0 0 10px 0;'>"
            "<span style='font-size:12px; color:#333; line-height:1.5;'>" + info.description + "</span>"
            "</td></tr>";

    // Separator
    html += "<tr><td style='padding: 0 0 10px 0; border-top: 1px solid #E0E0E0;'></td></tr>";

    // Syntax label
    html += "<tr><td style='padding: 0 0 4px 0;'>"
            "<span style='font-size:11px; color:#888; text-transform:uppercase; letter-spacing:0.5px;'>Syntax</span>"
            "</td></tr>";

    // Syntax value
    html += "<tr><td style='padding: 6px 10px; background:#F8F9FA; border-radius:4px;'>"
            "<span style='font-size:12px; color:#202124; font-family:SF Mono, Menlo, monospace;'>" + info.syntax + "</span>"
            "</td></tr>";

    // Spacer
    html += "<tr><td style='padding: 6px 0 0 0;'></td></tr>";

    // Parameters
    if (!info.params.isEmpty()) {
        html += "<tr><td style='padding: 0 0 6px 0;'>"
                "<span style='font-size:11px; color:#888; text-transform:uppercase; letter-spacing:0.5px;'>Parameters</span>"
                "</td></tr>";

        for (const auto& p : info.params) {
            QString pLabel = p.optional ? "[" + p.name + "]" : p.name;
            QString badge = p.optional
                ? "<span style='font-size:9px; color:#5F6368; background:#F1F3F4; padding:1px 5px; "
                  "border-radius:3px; margin-left:6px;'>optional</span>"
                : "<span style='font-size:9px; color:#D93025; background:#FCE8E6; padding:1px 5px; "
                  "border-radius:3px; margin-left:6px;'>required</span>";

            // Param name + badge
            html += "<tr><td style='padding: 6px 0 0 0;'>"
                    "<span style='font-size:12px; font-weight:bold; color:#1A73E8;'>" + pLabel + "</span>"
                    + badge + "</td></tr>";

            // Param description
            if (!p.description.isEmpty()) {
                html += "<tr><td style='padding: 2px 0 4px 12px;'>"
                        "<span style='font-size:11px; color:#5F6368; line-height:1.4;'>" + p.description + "</span>"
                        "</td></tr>";
            }
        }
    }

    html += "</table>";
    return html;
}

#endif // FORMULAMETADATA_H

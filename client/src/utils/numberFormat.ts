import { NumberFormatType } from '../types/spreadsheet';

/**
 * Currency definitions for the currency submenu.
 */
export interface CurrencyDef {
  code: string;
  symbol: string;
  label: string;
}

export const CURRENCIES: CurrencyDef[] = [
  { code: 'USD', symbol: '$', label: 'US Dollar ($)' },
  { code: 'EUR', symbol: '\u20AC', label: 'Euro (\u20AC)' },
  { code: 'GBP', symbol: '\u00A3', label: 'British Pound (\u00A3)' },
  { code: 'JPY', symbol: '\u00A5', label: 'Japanese Yen (\u00A5)' },
  { code: 'INR', symbol: '\u20B9', label: 'Indian Rupee (\u20B9)' },
  { code: 'CNY', symbol: '\u00A5', label: 'Chinese Yuan (\u00A5)' },
  { code: 'KRW', symbol: '\u20A9', label: 'Korean Won (\u20A9)' },
  { code: 'CAD', symbol: 'CA$', label: 'Canadian Dollar (CA$)' },
  { code: 'AUD', symbol: 'A$', label: 'Australian Dollar (A$)' },
  { code: 'CHF', symbol: 'CHF', label: 'Swiss Franc (CHF)' },
  { code: 'BRL', symbol: 'R$', label: 'Brazilian Real (R$)' },
  { code: 'MXN', symbol: 'MX$', label: 'Mexican Peso (MX$)' },
];

/**
 * Date format options for the date submenu.
 */
export interface DateFormatDef {
  id: string;
  label: string;
  example: string;
  options: Intl.DateTimeFormatOptions;
}

export const DATE_FORMATS: DateFormatDef[] = [
  { id: 'mm/dd/yyyy', label: 'MM/DD/YYYY', example: '02/16/2026', options: { month: '2-digit', day: '2-digit', year: 'numeric' } },
  { id: 'dd/mm/yyyy', label: 'DD/MM/YYYY', example: '16/02/2026', options: { day: '2-digit', month: '2-digit', year: 'numeric' } },
  { id: 'yyyy-mm-dd', label: 'YYYY-MM-DD', example: '2026-02-16', options: { year: 'numeric', month: '2-digit', day: '2-digit' } },
  { id: 'mmm d, yyyy', label: 'MMM D, YYYY', example: 'Feb 16, 2026', options: { month: 'short', day: 'numeric', year: 'numeric' } },
  { id: 'mmmm d, yyyy', label: 'MMMM D, YYYY', example: 'February 16, 2026', options: { month: 'long', day: 'numeric', year: 'numeric' } },
  { id: 'd-mmm-yy', label: 'D-MMM-YY', example: '16-Feb-26', options: { day: 'numeric', month: 'short', year: '2-digit' } },
  { id: 'mm/dd', label: 'MM/DD', example: '02/16', options: { month: '2-digit', day: '2-digit' } },
];

/**
 * Format a cell value based on the number format type and decimal places.
 * Optionally accepts a currency code and date format id for richer formatting.
 */
export function formatCellValue(
  value: string,
  numberFormat: NumberFormatType,
  decimalPlaces: number,
  currencyCode?: string,
  dateFormatId?: string,
): string {
  if (!value || numberFormat === 'general' || numberFormat === 'text') {
    return value;
  }

  const num = parseFloat(value);
  if (isNaN(num) && numberFormat !== 'date' && numberFormat !== 'time') {
    return value;
  }

  switch (numberFormat) {
    case 'number':
      return num.toLocaleString('en-US', {
        minimumFractionDigits: decimalPlaces,
        maximumFractionDigits: decimalPlaces,
      });

    case 'currency': {
      const cur = currencyCode || 'USD';
      return num.toLocaleString('en-US', {
        style: 'currency',
        currency: cur,
        minimumFractionDigits: decimalPlaces,
        maximumFractionDigits: decimalPlaces,
      });
    }

    case 'accounting': {
      const cur = currencyCode || 'USD';
      const formatted = Math.abs(num).toLocaleString('en-US', {
        style: 'currency',
        currency: cur,
        minimumFractionDigits: decimalPlaces,
        maximumFractionDigits: decimalPlaces,
      });
      return num < 0 ? `(${formatted})` : formatted;
    }

    case 'percentage':
      return (num * 100).toLocaleString('en-US', {
        minimumFractionDigits: decimalPlaces,
        maximumFractionDigits: decimalPlaces,
      }) + '%';

    case 'date': {
      const date = parseDate(value);
      if (!date) return value;
      const fmt = DATE_FORMATS.find(f => f.id === dateFormatId) ?? DATE_FORMATS[0];
      // For specific format patterns, use manual formatting to match exactly
      if (dateFormatId === 'yyyy-mm-dd') {
        const y = date.getFullYear();
        const m = String(date.getMonth() + 1).padStart(2, '0');
        const d = String(date.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
      }
      if (dateFormatId === 'dd/mm/yyyy') {
        const d = String(date.getDate()).padStart(2, '0');
        const m = String(date.getMonth() + 1).padStart(2, '0');
        const y = date.getFullYear();
        return `${d}/${m}/${y}`;
      }
      return date.toLocaleDateString('en-US', fmt.options);
    }

    case 'time': {
      const date = parseDate(value);
      if (date) {
        return date.toLocaleTimeString('en-US', {
          hour: '2-digit',
          minute: '2-digit',
          second: '2-digit',
        });
      }
      return value;
    }

    case 'custom':
      return value; // custom formatting handled externally

    default:
      return value;
  }
}

/**
 * Apply an Excel-style custom format string to a numeric value.
 * Supported patterns:
 *   #,##0      → 1,235
 *   #,##0.00   → 1,234.56
 *   0.00       → 1234.56
 *   0%         → 123456%
 *   0.00%      → 123456.00%
 *   $#,##0     → $1,235
 *   $#,##0.00  → $1,234.56
 *   [Red]      → (prefix, ignored for now)
 *   "text"0    → text1235
 */
export function applyCustomFormatString(value: string, formatStr: string): string {
  if (!formatStr) return value;

  const num = parseFloat(value);
  if (isNaN(num)) return value;

  // Split format for positive;negative;zero (Excel convention)
  const parts = formatStr.split(';');
  let fmt = parts[0];
  if (num < 0 && parts.length > 1) fmt = parts[1];
  else if (num === 0 && parts.length > 2) fmt = parts[2];

  // Strip color codes like [Red], [Blue] etc.
  fmt = fmt.replace(/\[[A-Za-z]+\]/g, '');

  // Handle percentage
  const isPercent = fmt.includes('%');
  let val = isPercent ? num * 100 : num;
  const absVal = Math.abs(val);

  // Check for prefix/suffix text (quoted strings)
  let prefix = '';
  let suffix = '';
  fmt = fmt.replace(/"([^"]*)"/g, (_, text) => {
    // Determine if it's before or after the number part
    prefix += text;
    return '';
  });

  // Extract currency/prefix symbols before the number pattern
  const prefixMatch = fmt.match(/^([^#0,.]+)/);
  if (prefixMatch) {
    prefix += prefixMatch[1].trim();
    fmt = fmt.slice(prefixMatch[0].length);
  }

  // Extract suffix after the number pattern
  const suffixMatch = fmt.match(/([^#0,.]+)$/);
  if (suffixMatch) {
    suffix = suffixMatch[1].trim();
    fmt = fmt.slice(0, -suffixMatch[0].length);
  }

  if (isPercent) suffix = '%' + suffix.replace('%', '');

  // Count decimal places
  const dotIdx = fmt.indexOf('.');
  let decimals = 0;
  if (dotIdx !== -1) {
    decimals = fmt.length - dotIdx - 1;
  }

  // Determine if we use thousands separator
  const useComma = fmt.includes(',');

  // Format the number
  const formatted = absVal.toLocaleString('en-US', {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
    useGrouping: useComma,
  });

  const sign = val < 0 ? '-' : '';
  return `${sign}${prefix}${formatted}${suffix}`;
}

function parseDate(value: string): Date | null {
  // Try ISO / standard date string
  const d = new Date(value);
  if (!isNaN(d.getTime())) return d;

  // Try Excel serial date (days since 1900-01-01)
  const num = parseFloat(value);
  if (!isNaN(num) && num > 0 && num < 200000) {
    const excelEpoch = new Date(1899, 11, 30); // Excel epoch
    const ms = excelEpoch.getTime() + num * 86400000;
    return new Date(ms);
  }

  return null;
}

export const NUMBER_FORMAT_LABELS: Record<NumberFormatType, string> = {
  general: 'General',
  number: 'Number',
  currency: 'Currency',
  accounting: 'Accounting',
  percentage: 'Percentage',
  date: 'Date',
  time: 'Time',
  text: 'Text',
  custom: 'Custom',
};

export const NUMBER_FORMAT_EXAMPLES: Record<NumberFormatType, string> = {
  general: '1234.56',
  number: '1,234.56',
  currency: '$1,234.56',
  accounting: '$1,234.56',
  percentage: '12.35%',
  date: '02/16/2026',
  time: '01:30:00 PM',
  text: '1234.56',
  custom: '#,##0.00',
};

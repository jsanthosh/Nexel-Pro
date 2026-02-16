/**
 * Fill Series — detects numeric, month, day, and text+number patterns
 * and generates a continuation series for the given count.
 */

const MONTHS_FULL  = ['January','February','March','April','May','June','July','August','September','October','November','December'];
const MONTHS_SHORT = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const DAYS_FULL    = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
const DAYS_SHORT   = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];

function matchCase(template: string, value: string): string {
  if (template === template.toUpperCase()) return value.toUpperCase();
  if (template[0] === template[0].toUpperCase()) return value[0].toUpperCase() + value.slice(1).toLowerCase();
  return value.toLowerCase();
}

/**
 * Given 1–2 seed values, generate `count` values starting from the first seed.
 * Returns an array of `count` strings representing the full series.
 */
export function generateSeries(seeds: string[], count: number): string[] {
  if (seeds.length === 0 || count === 0) return [];

  const s0 = seeds[0].trim();
  const s1 = seeds.length > 1 ? seeds[1].trim() : '';

  // --- Numeric ---
  const n0 = parseFloat(s0);
  if (!isNaN(n0) && s0 !== '') {
    const n1 = parseFloat(s1);
    const step = (!isNaN(n1) && seeds.length > 1) ? n1 - n0 : 1;
    const isInt = Number.isInteger(n0) && Number.isInteger(step);
    return Array.from({ length: count }, (_, i) => {
      const val = n0 + step * i;
      return isInt ? String(Math.round(val)) : String(parseFloat(val.toFixed(10)));
    });
  }

  // --- List-based (months / days) ---
  for (const list of [MONTHS_FULL, MONTHS_SHORT, DAYS_FULL, DAYS_SHORT]) {
    const lower = list.map(s => s.toLowerCase());
    const idx0 = lower.indexOf(s0.toLowerCase());
    if (idx0 === -1) continue;

    let step = 1;
    if (seeds.length > 1) {
      const idx1 = lower.indexOf(s1.toLowerCase());
      if (idx1 !== -1) step = ((idx1 - idx0) % list.length + list.length) % list.length || list.length;
    }

    return Array.from({ length: count }, (_, i) => {
      const item = list[(idx0 + step * i) % list.length];
      return matchCase(s0, item);
    });
  }

  // --- Text + trailing number suffix: "Item 1", "Q1", "Week 1" ---
  const m0 = s0.match(/^(.*?)(\d+)$/);
  if (m0) {
    const prefix = m0[1];
    const start  = parseInt(m0[2], 10);
    let step = 1;
    if (seeds.length > 1) {
      const m1 = s1.match(/^(.*?)(\d+)$/);
      if (m1 && m1[1] === prefix) step = parseInt(m1[2], 10) - start;
    }
    return Array.from({ length: count }, (_, i) => `${prefix}${start + step * i}`);
  }

  // --- Fallback: repeat first seed ---
  return Array(count).fill(s0);
}

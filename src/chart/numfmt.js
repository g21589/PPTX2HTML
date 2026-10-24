/**
 * Excel number format codes (subset sufficient for chart labels).
 */

const DATE_RE = /(^|[^"\\])(y{1,4}|m{1,5}|d{1,4}|h{1,2}|s{1,2}|AM\/PM|A\/P)/i;

function splitSections(code) {
  const out = [];
  let cur = '';
  let q = false;
  for (let i = 0; i < code.length; i++) {
    const c = code[i];
    if (c === '"') q = !q;
    if (c === '\\' && !q) { cur += c + (code[i + 1] || ''); i++; continue; }
    if (c === ';' && !q) { out.push(cur); cur = ''; continue; }
    cur += c;
  }
  out.push(cur);
  return out;
}

function serialToDate(v, date1904) {
  let days = Math.floor(v);
  const frac = v - days;
  if (date1904) days += 1462;
  // Excel's 1900 leap year bug.
  if (days > 59) days -= 1;
  const ms = Date.UTC(1899, 11, 31) + days * 86400000 + Math.round(frac * 86400000);
  return new Date(ms);
}

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

function formatDate(v, code, date1904) {
  const d = serialToDate(v, date1904);
  const Y = d.getUTCFullYear(), M = d.getUTCMonth(), D = d.getUTCDate(), wd = d.getUTCDay();
  const h = d.getUTCHours(), mi = d.getUTCMinutes(), s = d.getUTCSeconds();
  const ampm = /AM\/PM|A\/P/i.test(code);
  let out = '';
  const tokens = code.match(/"[^"]*"|\\.|yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|A\/P|\[[^\]]*\]|./gi) || [];
  let lastWasHour = false;
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    const tl = t.toLowerCase();
    if (t[0] === '"') { out += t.slice(1, -1); continue; }
    if (t[0] === '\\') { out += t[1]; continue; }
    if (t[0] === '[') continue;
    switch (tl) {
      case 'yyyy': out += Y; break;
      case 'yy': out += String(Y).slice(2); break;
      case 'mmmmm': out += MONTHS[M][0]; break;
      case 'mmmm': out += MONTHS[M]; break;
      case 'mmm': out += MONTHS[M].slice(0, 3); break;
      case 'mm':
      case 'm': {
        // Minutes when following hours or preceding seconds.
        const next = (tokens.slice(i + 1).find((x) => /[a-z]/i.test(x)) || '').toLowerCase();
        const isMin = lastWasHour || next.startsWith('s');
        out += isMin ? String(mi).padStart(tl.length, '0') : String(M + 1).padStart(tl.length, '0');
        break;
      }
      case 'dddd': out += DAYS[wd]; break;
      case 'ddd': out += DAYS[wd].slice(0, 3); break;
      case 'dd': out += String(D).padStart(2, '0'); break;
      case 'd': out += D; break;
      case 'hh': case 'h': {
        const hv = ampm ? ((h + 11) % 12) + 1 : h;
        out += String(hv).padStart(tl.length, '0');
        break;
      }
      case 'ss': case 's': out += String(s).padStart(tl.length, '0'); break;
      case 'am/pm': out += h < 12 ? 'AM' : 'PM'; break;
      case 'a/p': out += h < 12 ? 'A' : 'P'; break;
      default: out += t;
    }
    if (/^h/.test(tl)) lastWasHour = true;
    else if (/[a-z]/i.test(tl)) lastWasHour = false;
  }
  return out;
}

function groupThousands(s) {
  return s.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function formatNumberSection(v, sec) {
  // Literals and placeholders.
  let code = sec.replace(/\[[^\]]*\]/g, '');
  let percent = 0;
  let prefix = '';
  let suffix = '';
  // Extract quoted / escaped literals keeping position (prefix before first digit placeholder).
  const firstPh = code.search(/[0#?]/);
  if (firstPh < 0) {
    return code.replace(/"([^"]*)"/g, '$1').replace(/\\(.)/g, '$1').replace(/_.|\*./g, '');
  }
  let lastPh = -1;
  for (let i = code.length - 1; i >= 0; i--) if (/[0#?]/.test(code[i])) { lastPh = i; break; }
  prefix = code.slice(0, firstPh);
  let body = code.slice(firstPh, lastPh + 1);
  suffix = code.slice(lastPh + 1);
  const lit = (s) => s.replace(/"([^"]*)"/g, '$1').replace(/\\(.)/g, '$1').replace(/_.|\*./g, '');
  percent = ((prefix + suffix).match(/%/g) || []).length;
  for (let i = 0; i < percent; i++) v *= 100;
  // Scientific.
  const sci = /E[+-]/i.exec(body);
  if (sci) {
    const mant = body.slice(0, sci.index);
    const expDigits = body.slice(sci.index + 2).replace(/[^0#]/g, '').length || 1;
    const dec = (mant.split('.')[1] || '').replace(/[^0#]/g, '').length;
    let exp = v === 0 ? 0 : Math.floor(Math.log10(Math.abs(v)));
    let m = v / Math.pow(10, exp);
    let ms = m.toFixed(dec);
    if (Math.abs(Number(ms)) >= 10) { exp++; m = v / Math.pow(10, exp); ms = m.toFixed(dec); }
    return lit(prefix) + ms + 'E' + (exp < 0 ? '-' : '+') + String(Math.abs(exp)).padStart(expDigits, '0') + lit(suffix);
  }
  // Thousands scaling: commas right after the last placeholder.
  while (body.endsWith(',')) { v /= 1000; body = body.slice(0, -1); }
  while (suffix.startsWith(',')) { v /= 1000; suffix = suffix.slice(1); }
  const thousands = body.includes(',');
  const parts = body.replace(/,/g, '').split('.');
  const intPart = parts[0] || '';
  const decPart = parts[1] || '';
  const minDec = (decPart.match(/0/g) || []).length;
  const maxDec = (decPart.match(/[0#?]/g) || []).length;
  const minInt = (intPart.match(/0/g) || []).length;
  let s = Math.abs(v).toFixed(maxDec);
  if (maxDec > minDec) {
    // Trim optional trailing zeros.
    let [i, f] = s.split('.');
    f = f || '';
    while (f.length > minDec && f.endsWith('0')) f = f.slice(0, -1);
    s = f.length ? i + '.' + f : i;
  }
  let [ip, fp] = s.split('.');
  if (ip === '0' && minInt === 0) ip = '';
  ip = ip.padStart(minInt, '0');
  if (thousands) ip = groupThousands(ip);
  let res = ip + (fp !== undefined ? '.' + fp : (decPart.length && minDec ? '.' : ''));
  if (!res) res = '0';
  return lit(prefix) + res + lit(suffix);
}

/** General format: up to ~10 significant characters like Excel. */
export function formatGeneral(v) {
  if (typeof v !== 'number' || !Number.isFinite(v)) return String(v);
  if (Number.isInteger(v) && Math.abs(v) < 1e11) return String(v);
  const a = Math.abs(v);
  if (a !== 0 && (a >= 1e11 || a < 1e-9)) return v.toExponential(5).replace(/\.?0+e/, 'E').replace('e', 'E');
  let s = v.toPrecision(10);
  if (s.includes('e')) return v.toExponential(5).replace(/\.?0+e/, 'E').replace('e', 'E');
  if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '');
  return s;
}

/**
 * Format a value with an Excel format code.
 * @param {number|string} v
 * @param {string} code
 */
export function formatValue(v, code, date1904 = false) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v !== 'number') {
    const n = Number(v);
    if (Number.isNaN(n)) return String(v);
    v = n;
  }
  if (!code || /^general$/i.test(code.trim())) return formatGeneral(v);
  const sections = splitSections(code);
  let sec;
  let val = v;
  if (sections.length >= 2 && v < 0) { sec = sections[1]; val = -v; }
  else if (sections.length >= 3 && v === 0) sec = sections[2];
  else sec = sections[0];
  // Conditional sections ([>100]) are rare in charts; ignore the condition.
  if (/^@$/.test(sec.trim())) return formatGeneral(v);
  if (/general/i.test(sec)) return sec.replace(/general/i, formatGeneral(val)).replace(/"([^"]*)"/g, '$1');
  if (DATE_RE.test(sec.replace(/"[^"]*"/g, '')) && !/[0#?]/.test(sec.replace(/"[^"]*"/g, ''))) return formatDate(v, sec, date1904);
  const out = formatNumberSection(val, sec);
  return (v < 0 && sections.length < 2 && val === v) ? '-' + out : out;
}

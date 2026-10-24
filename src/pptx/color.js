/**
 * DrawingML colours: srgbClr, schemeClr, sysClr, prstClr, hslClr, scrgbClr
 * plus the full set of colour transforms.
 */
import { clamp } from '../util/units.js';

/** Preset colour names (ST_PresetColorVal). */
export const PRESET_COLORS = {
  aliceBlue: 'F0F8FF', antiqueWhite: 'FAEBD7', aqua: '00FFFF', aquamarine: '7FFFD4', azure: 'F0FFFF', beige: 'F5F5DC',
  bisque: 'FFE4C4', black: '000000', blanchedAlmond: 'FFEBCD', blue: '0000FF', blueViolet: '8A2BE2', brown: 'A52A2A',
  burlyWood: 'DEB887', cadetBlue: '5F9EA0', chartreuse: '7FFF00', chocolate: 'D2691E', coral: 'FF7F50',
  cornflowerBlue: '6495ED', cornsilk: 'FFF8DC', crimson: 'DC143C', cyan: '00FFFF', darkBlue: '00008B', darkCyan: '008B8B',
  darkGoldenrod: 'B8860B', darkGray: 'A9A9A9', darkGrey: 'A9A9A9', darkGreen: '006400', darkKhaki: 'BDB76B',
  darkMagenta: '8B008B', darkOliveGreen: '556B2F', darkOrange: 'FF8C00', darkOrchid: '9932CC', darkRed: '8B0000',
  darkSalmon: 'E9967A', darkSeaGreen: '8FBC8F', darkSlateBlue: '483D8B', darkSlateGray: '2F4F4F', darkSlateGrey: '2F4F4F',
  darkTurquoise: '00CED1', darkViolet: '9400D3', deepPink: 'FF1493', deepSkyBlue: '00BFFF', dimGray: '696969',
  dimGrey: '696969', dodgerBlue: '1E90FF', firebrick: 'B22222', floralWhite: 'FFFAF0', forestGreen: '228B22',
  fuchsia: 'FF00FF', gainsboro: 'DCDCDC', ghostWhite: 'F8F8FF', gold: 'FFD700', goldenrod: 'DAA520', gray: '808080',
  grey: '808080', green: '008000', greenYellow: 'ADFF2F', honeydew: 'F0FFF0', hotPink: 'FF69B4', indianRed: 'CD5C5C',
  indigo: '4B0082', ivory: 'FFFFF0', khaki: 'F0E68C', lavender: 'E6E6FA', lavenderBlush: 'FFF0F5', lawnGreen: '7CFC00',
  lemonChiffon: 'FFFACD', lightBlue: 'ADD8E6', lightCoral: 'F08080', lightCyan: 'E0FFFF', lightGoldenrodYellow: 'FAFAD2',
  lightGray: 'D3D3D3', lightGrey: 'D3D3D3', lightGreen: '90EE90', lightPink: 'FFB6C1', lightSalmon: 'FFA07A',
  lightSeaGreen: '20B2AA', lightSkyBlue: '87CEFA', lightSlateGray: '778899', lightSlateGrey: '778899',
  lightSteelBlue: 'B0C4DE', lightYellow: 'FFFFE0', lime: '00FF00', limeGreen: '32CD32', linen: 'FAF0E6',
  magenta: 'FF00FF', maroon: '800000', medAquamarine: '66CDAA', medBlue: '0000CD', medOrchid: 'BA55D3',
  medPurple: '9370DB', medSeaGreen: '3CB371', medSlateBlue: '7B68EE', medSpringGreen: '00FA9A',
  medTurquoise: '48D1CC', medVioletRed: 'C71585', mediumAquamarine: '66CDAA', mediumBlue: '0000CD',
  mediumOrchid: 'BA55D3', mediumPurple: '9370DB', mediumSeaGreen: '3CB371', mediumSlateBlue: '7B68EE',
  mediumSpringGreen: '00FA9A', mediumTurquoise: '48D1CC', mediumVioletRed: 'C71585', midnightBlue: '191970',
  mintCream: 'F5FFFA', mistyRose: 'FFE4E1', moccasin: 'FFE4B5', navajoWhite: 'FFDEAD', navy: '000080',
  oldLace: 'FDF5E6', olive: '808000', oliveDrab: '6B8E23', orange: 'FFA500', orangeRed: 'FF4500', orchid: 'DA70D6',
  paleGoldenrod: 'EEE8AA', paleGreen: '98FB98', paleTurquoise: 'AFEEEE', paleVioletRed: 'DB7093',
  papayaWhip: 'FFEFD5', peachPuff: 'FFDAB9', peru: 'CD853F', pink: 'FFC0CB', plum: 'DDA0DD', powderBlue: 'B0E0E6',
  purple: '800080', red: 'FF0000', rosyBrown: 'BC8F8F', royalBlue: '4169E1', saddleBrown: '8B4513', salmon: 'FA8072',
  sandyBrown: 'F4A460', seaGreen: '2E8B57', seaShell: 'FFF5EE', sienna: 'A0522D', silver: 'C0C0C0', skyBlue: '87CEEB',
  slateBlue: '6A5ACD', slateGray: '708090', slateGrey: '708090', snow: 'FFFAFA', springGreen: '00FF7F',
  steelBlue: '4682B4', tan: 'D2B48C', teal: '008080', thistle: 'D8BFD8', tomato: 'FF6347', turquoise: '40E0D0',
  violet: 'EE82EE', wheat: 'F5DEB3', white: 'FFFFFF', whiteSmoke: 'F5F5F5', yellow: 'FFFF00', yellowGreen: '9ACD32',
  dkBlue: '00008B', dkCyan: '008B8B', dkGoldenrod: 'B8860B', dkGray: 'A9A9A9', dkGrey: 'A9A9A9', dkGreen: '006400',
  dkKhaki: 'BDB76B', dkMagenta: '8B008B', dkOliveGreen: '556B2F', dkOrange: 'FF8C00', dkOrchid: '9932CC', dkRed: '8B0000',
  dkSalmon: 'E9967A', dkSeaGreen: '8FBC8B', dkSlateBlue: '483D8B', dkSlateGray: '2F4F4F', dkSlateGrey: '2F4F4F',
  dkTurquoise: '00CED1', dkViolet: '9400D3', ltBlue: 'ADD8E6', ltCoral: 'F08080', ltCyan: 'E0FFFF',
  ltGoldenrodYellow: 'FAFAD2', ltGray: 'D3D3D3', ltGrey: 'D3D3D3', ltGreen: '90EE90', ltPink: 'FFB6C1',
  ltSalmon: 'FFA07A', ltSeaGreen: '20B2AA', ltSkyBlue: '87CEFA', ltSlateGray: '778899', ltSlateGrey: '778899',
  ltSteelBlue: 'B0C4DE', ltYellow: 'FFFFE0',
};

/** Windows system colours (used when sysClr has no lastClr). */
export const SYSTEM_COLORS = {
  scrollBar: 'C8C8C8', background: '000000', activeCaption: '99B4D1', inactiveCaption: 'BFCDDB', menu: 'F0F0F0',
  window: 'FFFFFF', windowFrame: '646464', menuText: '000000', windowText: '000000', captionText: '000000',
  activeBorder: 'B4B4B4', inactiveBorder: 'F4F7FC', appWorkspace: 'ABABAB', highlight: '3399FF', highlightText: 'FFFFFF',
  btnFace: 'F0F0F0', btnShadow: 'A0A0A0', grayText: '6D6D6D', btnText: '000000', inactiveCaptionText: '434E54',
  btnHighlight: 'FFFFFF', '3dDkShadow': '696969', '3dLight': 'E3E3E3', infoText: '000000', infoBk: 'FFFFE1',
  hotLight: '0066CC', gradientActiveCaption: 'B9D1EA', gradientInactiveCaption: 'D7E4F2', menuHighlight: '3399FF',
  menuBar: 'F0F0F0',
};

export const COLOR_NODES = new Set(['a:srgbClr', 'a:schemeClr', 'a:sysClr', 'a:prstClr', 'a:hslClr', 'a:scrgbClr']);

export function hexToRgb(hex) {
  let h = String(hex || '').replace('#', '').trim();
  if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
  const n = parseInt(h.slice(0, 6), 16);
  if (Number.isNaN(n)) return { r: 0, g: 0, b: 0, a: 1 };
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255, a: 1 };
}

const toLin = (c) => {
  c /= 255;
  return c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
};
const fromLin = (c) => {
  c = clamp(c, 0, 1);
  return 255 * (c <= 0.0031308 ? c * 12.92 : 1.055 * Math.pow(c, 1 / 2.4) - 0.055);
};

export function rgbToHsl(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const l = (max + min) / 2;
  let h = 0;
  let s = 0;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    h *= 60;
  }
  return { h, s, l };
}

export function hslToRgb(h, s, l) {
  h = ((h % 360) + 360) % 360;
  s = clamp(s, 0, 1);
  l = clamp(l, 0, 1);
  if (s === 0) {
    const v = l * 255;
    return { r: v, g: v, b: v };
  }
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const hk = h / 360;
  const f = (t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return { r: f(hk + 1 / 3) * 255, g: f(hk) * 255, b: f(hk - 1 / 3) * 255 };
}

/**
 * Apply DrawingML transforms (child elements of a colour node) to a colour.
 * @param {{r:number,g:number,b:number,a:number}} c
 * @param {import('../util/xml.js').XNode[]} mods
 */
export function applyTransforms(c, mods) {
  let { r, g, b, a } = c;
  for (const m of mods) {
    const v = m.num('val', 0) / 100000;
    switch (m.local) {
      case 'alpha': a = clamp(v, 0, 1); break;
      case 'alphaOff': a = clamp(a + v, 0, 1); break;
      case 'alphaMod': a = clamp(a * v, 0, 1); break;
      case 'tint': {
        // Lighter version: mix with white in linear RGB.
        const f = (x) => fromLin(toLin(x) * v + (1 - v));
        r = f(r); g = f(g); b = f(b);
        break;
      }
      case 'shade': {
        const f = (x) => fromLin(toLin(x) * v);
        r = f(r); g = f(g); b = f(b);
        break;
      }
      case 'comp': {
        const hsl = rgbToHsl(r, g, b);
        ({ r, g, b } = hslToRgb(hsl.h + 180, hsl.s, hsl.l));
        break;
      }
      case 'inv': r = 255 - r; g = 255 - g; b = 255 - b; break;
      case 'gray': {
        const y = 0.2126 * r + 0.7152 * g + 0.0722 * b;
        r = g = b = y;
        break;
      }
      case 'gamma': r = fromLin(r / 255) ; g = fromLin(g / 255); b = fromLin(b / 255); break;
      case 'invGamma': r = toLin(r) * 255; g = toLin(g) * 255; b = toLin(b) * 255; break;
      case 'hue':
      case 'hueOff':
      case 'hueMod':
      case 'sat':
      case 'satOff':
      case 'satMod':
      case 'lum':
      case 'lumOff':
      case 'lumMod': {
        const hsl = rgbToHsl(r, g, b);
        const deg = m.num('val', 0) / 60000;
        switch (m.local) {
          case 'hue': hsl.h = deg; break;
          case 'hueOff': hsl.h += deg; break;
          case 'hueMod': hsl.h *= v; break;
          case 'sat': hsl.s = v; break;
          case 'satOff': hsl.s += v; break;
          case 'satMod': hsl.s *= v; break;
          case 'lum': hsl.l = v; break;
          case 'lumOff': hsl.l += v; break;
          case 'lumMod': hsl.l *= v; break;
          default: break;
        }
        ({ r, g, b } = hslToRgb(hsl.h, clamp(hsl.s, 0, 1), clamp(hsl.l, 0, 1)));
        break;
      }
      case 'red': r = fromLin(v); break;
      case 'redOff': r = fromLin(toLin(r) + v); break;
      case 'redMod': r = fromLin(toLin(r) * v); break;
      case 'green': g = fromLin(v); break;
      case 'greenOff': g = fromLin(toLin(g) + v); break;
      case 'greenMod': g = fromLin(toLin(g) * v); break;
      case 'blue': b = fromLin(v); break;
      case 'blueOff': b = fromLin(toLin(b) + v); break;
      case 'blueMod': b = fromLin(toLin(b) * v); break;
      default: break;
    }
  }
  return { r: clamp(r, 0, 255), g: clamp(g, 0, 255), b: clamp(b, 0, 255), a };
}

/** Default mapping from presentation scheme names to theme slots. */
export const DEFAULT_CLR_MAP = {
  bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2', accent1: 'accent1', accent2: 'accent2', accent3: 'accent3',
  accent4: 'accent4', accent5: 'accent5', accent6: 'accent6', hlink: 'hlink', folHlink: 'folHlink',
};

/**
 * Resolve a colour choice element to RGBA.
 * @param {import('../util/xml.js').XNode} node the colour element (a:srgbClr...)
 * @param {{scheme: object, clrMap: object, phClr?: object}} env
 */
export function resolveColorNode(node, env) {
  if (!node) return null;
  let base;
  switch (node.name) {
    case 'a:srgbClr':
      base = hexToRgb(node.attrs.val);
      break;
    case 'a:sysClr':
      base = hexToRgb(node.attrs.lastClr || SYSTEM_COLORS[node.attrs.val] || '000000');
      break;
    case 'a:prstClr':
      base = hexToRgb(PRESET_COLORS[node.attrs.val] || '000000');
      break;
    case 'a:hslClr': {
      const rgb = hslToRgb(node.num('hue', 0) / 60000, node.num('sat', 0) / 100000, node.num('lum', 0) / 100000);
      base = { ...rgb, a: 1 };
      break;
    }
    case 'a:scrgbClr':
      base = { r: fromLin(node.num('r', 0) / 100000), g: fromLin(node.num('g', 0) / 100000), b: fromLin(node.num('b', 0) / 100000), a: 1 };
      break;
    case 'a:schemeClr': {
      base = resolveScheme(node.attrs.val, env);
      if (!base) base = { r: 0, g: 0, b: 0, a: 1 };
      break;
    }
    default:
      return null;
  }
  if (node.children.length) return applyTransforms(base, node.children);
  return { r: base.r, g: base.g, b: base.b, a: base.a === undefined ? 1 : base.a };
}

export function resolveScheme(name, env) {
  if (name === 'phClr') {
    if (env.phClr) return env.phClr;
    name = 'tx1';
  }
  const map = env.clrMap || DEFAULT_CLR_MAP;
  const slot = map[name] || name;
  const scheme = env.scheme || {};
  const c = scheme[slot] || scheme[name];
  return c ? { ...c } : null;
}

/** Find the colour child in a container (e.g. <a:solidFill>) and resolve it. */
export function colorFrom(container, env) {
  if (!container) return null;
  for (const c of container.children) {
    if (COLOR_NODES.has(c.name)) return resolveColorNode(c, env);
  }
  return null;
}

export function colorChild(container) {
  if (!container) return null;
  for (const c of container.children) if (COLOR_NODES.has(c.name)) return c;
  return null;
}

const hex2 = (v) => Math.round(v).toString(16).padStart(2, '0');

/** CSS colour string. */
export function css(c, alphaMul = 1) {
  if (!c) return 'transparent';
  const a = (c.a === undefined ? 1 : c.a) * alphaMul;
  if (a >= 0.999) return '#' + hex2(c.r) + hex2(c.g) + hex2(c.b);
  return `rgba(${Math.round(c.r)},${Math.round(c.g)},${Math.round(c.b)},${Math.round(a * 1000) / 1000})`;
}

export function toHex(c) {
  return c ? hex2(c.r) + hex2(c.g) + hex2(c.b) : '000000';
}

export function luminance(c) {
  return (0.2126 * toLin(c.r) + 0.7152 * toLin(c.g) + 0.0722 * toLin(c.b));
}

export function mix(c1, c2, t) {
  return { r: c1.r + (c2.r - c1.r) * t, g: c1.g + (c2.g - c1.g) * t, b: c1.b + (c2.b - c1.b) * t, a: c1.a + (c2.a - c1.a) * t };
}

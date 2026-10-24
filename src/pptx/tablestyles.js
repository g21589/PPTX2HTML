/**
 * Built-in PowerPoint table styles. PowerPoint normally embeds the
 * definitions of the styles a deck uses in tableStyles.xml; this module
 * regenerates them when they are missing (e.g. files written by other tools).
 */

// Style id -> [family, accent] (ids as documented by Microsoft).
const IDS = {
  '{2D5ABB26-0587-4C30-8999-92F81FD0307C}': ['Themed1', ''],
  '{3C2FFA5D-87B4-456A-9821-1D502468CF0F}': ['Themed1', 'accent1'],
  '{284E427A-3D55-4303-BF80-6455036E1DE7}': ['Themed1', 'accent2'],
  '{69C7853C-536D-4A76-A0AE-DD22124D55A5}': ['Themed1', 'accent3'],
  '{775DCB02-9BB8-47FD-8907-85C794F793BA}': ['Themed1', 'accent4'],
  '{35758FB7-9AC5-4552-8A53-C91805E547FA}': ['Themed1', 'accent5'],
  '{08FB837D-C827-4EFA-A057-4D05807E0F7C}': ['Themed1', 'accent6'],
  '{5940675A-B579-460E-94D1-54222C63F5DA}': ['Themed2', ''],
  '{D113A9D2-9D6B-4929-AA2D-F23B5EE8CBE7}': ['Themed2', 'accent1'],
  '{18603FDC-E32A-4AB5-989C-0864C3EAD2B8}': ['Themed2', 'accent2'],
  '{306799F8-075E-4A3A-A7F6-7FBC6576F1A4}': ['Themed2', 'accent3'],
  '{E269D01E-BC32-4049-B463-5C60D7B0CCD2}': ['Themed2', 'accent4'],
  '{327F97BB-C833-4FB7-BDE5-3F7075034690}': ['Themed2', 'accent5'],
  '{638B1855-1B75-4FBE-930C-398BA8C253C6}': ['Themed2', 'accent6'],
  '{9D7B26C5-4107-4FEC-AEDC-1716B250A1EF}': ['Light1', ''],
  '{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}': ['Light1', 'accent1'],
  '{0E3FDE45-AF77-4B5C-9715-49D594BDF05E}': ['Light1', 'accent2'],
  '{C083E6E3-FA7D-4D7B-A595-EF9225AFEA82}': ['Light1', 'accent3'],
  '{D27102A9-8310-4765-A935-A1911B00CA55}': ['Light1', 'accent4'],
  '{5FD0F851-EC5A-4D38-B0AD-8093EC10F338}': ['Light1', 'accent5'],
  '{68D230F3-CF80-4859-8CE7-A43EE81993B5}': ['Light1', 'accent6'],
  '{7E9639D4-E3E2-4D34-9284-5A2195B3D0D7}': ['Light2', ''],
  '{69012ECD-51FC-41F1-AA8D-1B2483CD663E}': ['Light2', 'accent1'],
  '{72833802-FEF1-4C79-8D5D-14CF1EAF98D9}': ['Light2', 'accent2'],
  '{F2DE63D5-997A-4646-A377-4702673A728D}': ['Light2', 'accent3'],
  '{17292A2E-F333-43FB-9621-5CBBE7FDCDCB}': ['Light2', 'accent4'],
  '{5A111915-BE36-4E01-A7E5-04B1672EAD32}': ['Light2', 'accent5'],
  '{912C8C85-51F0-491E-9774-3900AFEF0FD7}': ['Light2', 'accent6'],
  '{616DA210-FB5B-4158-B5E0-FEB733F419BA}': ['Light3', ''],
  '{BC89EF96-8CEA-46FF-86C4-4CE0E7609802}': ['Light3', 'accent1'],
  '{5DA37D80-6434-44D0-A028-1B22A696006F}': ['Light3', 'accent2'],
  '{8799B23B-EC83-4686-B30A-512413B5E67A}': ['Light3', 'accent3'],
  '{ED083AE6-46FA-4A59-8FB0-9F97EB10719F}': ['Light3', 'accent4'],
  '{BDBED569-4797-4DF1-A0F4-6AAB3CD982D8}': ['Light3', 'accent5'],
  '{E8B1032C-EA38-4F05-BA0D-38AFFFC7BED3}': ['Light3', 'accent6'],
  '{793D81CF-94F2-401A-BA57-92F5A7B2D0C5}': ['Medium1', ''],
  '{B301B821-A1FF-4177-AEE7-76D212191A09}': ['Medium1', 'accent1'],
  '{9DCAF9ED-07DC-4A11-8D7F-57B35C25682E}': ['Medium1', 'accent2'],
  '{1FECB4D8-DB02-4DC6-A0A2-4F2EBAE1DC90}': ['Medium1', 'accent3'],
  '{1E171933-4619-4E11-9A3F-F7608DF75F80}': ['Medium1', 'accent4'],
  '{FABFCF23-3B69-468F-B69F-88F6DE6A72F2}': ['Medium1', 'accent5'],
  '{10A1B5D5-9B99-4C35-A422-299274C87663}': ['Medium1', 'accent6'],
  '{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}': ['Medium2', ''],
  '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}': ['Medium2', 'accent1'],
  '{21E4AEA4-8DFA-4A89-87EB-49C32662AFE0}': ['Medium2', 'accent2'],
  '{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}': ['Medium2', 'accent3'],
  '{00A15C55-8517-42AA-B614-E9B94910E393}': ['Medium2', 'accent4'],
  '{7DF18680-E054-41AD-8BC1-D1AEF772440D}': ['Medium2', 'accent5'],
  '{93296810-A885-4BE3-A3E7-6D5BEEA58F35}': ['Medium2', 'accent6'],
  '{8EC20E35-A176-4012-BC5E-935CFFF8708E}': ['Medium3', ''],
  '{6E25E649-3F16-4E02-A733-19D2CDBF48F0}': ['Medium3', 'accent1'],
  '{85BE263C-DBD7-4A20-BB59-AAB30ACAA65A}': ['Medium3', 'accent2'],
  '{EB344D84-9AFB-497E-A393-DC336BA19D2E}': ['Medium3', 'accent3'],
  '{EB9631B5-78F2-41C9-869B-9F39066F8104}': ['Medium3', 'accent4'],
  '{74C1A8A3-306A-4EB7-A6B1-4F7E0EB9C5D6}': ['Medium3', 'accent5'],
  '{2A488322-F2BA-4B5B-9748-0D474271808F}': ['Medium3', 'accent6'],
  '{D7AC3CCA-C797-4891-BE02-D94E43425B78}': ['Medium4', ''],
  '{69CF1AB2-1976-4502-BF36-3FF5EA218861}': ['Medium4', 'accent1'],
  '{8A107856-5554-42FB-B03E-39F5DBC370BA}': ['Medium4', 'accent2'],
  '{0505E3EF-67EA-436B-97B2-0124C06EBD24}': ['Medium4', 'accent3'],
  '{C4B1156A-380E-4F78-BDF5-A606A8083BF9}': ['Medium4', 'accent4'],
  '{22838BEF-8BB2-4498-84A7-C5851F593DF1}': ['Medium4', 'accent5'],
  '{16D9F66E-5EB9-4882-86FB-DCBF35E3C3E4}': ['Medium4', 'accent6'],
  '{E8034E78-7F5D-4C2E-B375-FC64B27BC917}': ['Dark1', ''],
  '{125E5076-3810-47DD-B79F-674D7AD40C01}': ['Dark1', 'accent1'],
  '{37CE84F3-28C3-443E-9E96-99CF82512B78}': ['Dark1', 'accent2'],
  '{D03447BB-5D67-496B-8E87-E561075AD55C}': ['Dark1', 'accent3'],
  '{E929F9F4-4A8F-4326-A1B4-22849713DDAB}': ['Dark1', 'accent4'],
  '{8FD4443E-F989-4FC4-A0C8-D5A2AF1F390B}': ['Dark1', 'accent5'],
  '{AF606853-7671-496A-8E4F-DF71F8EC918B}': ['Dark1', 'accent6'],
  '{5202B0CA-FC54-4496-8BCA-5EF66A818D29}': ['Dark2', ''],
  '{0660B408-B3CF-4A94-85FC-2B1E0A45F4A2}': ['Dark2', 'accent1'],
  '{91EBBBCC-DAD2-459C-BE2E-F6DE35CF9A28}': ['Dark2', 'accent3'],
  '{46F890A9-2807-4EBB-B81D-B2AA78EC7F39}': ['Dark2', 'accent5'],
};

/** Colour spec: [scheme, transformName?, value?] */
const clr = (scheme, mod, val) => {
  const inner = mod ? `<a:${mod} val="${val}"/>` : '';
  return `<a:schemeClr val="${scheme}">${inner}</a:schemeClr>`;
};
const solid = (c) => `<a:solidFill>${c}</a:solidFill>`;
const ln = (c, w = 12700) => `<a:ln w="${w}" cmpd="sng">${solid(c)}</a:ln>`;

/** Build a table style XML from a description. */
function build(id, desc) {
  let xml = `<a:tblStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" styleId="${id}" styleName="builtin">`;
  if (desc.tblBg) xml += `<a:tblBg>${solid(desc.tblBg)}</a:tblBg>`;
  for (const part of ['wholeTbl', 'band1H', 'band2H', 'band1V', 'band2V', 'lastCol', 'firstCol', 'lastRow', 'firstRow']) {
    const p = desc[part];
    if (!p) continue;
    xml += `<a:${part}>`;
    if (p.text || p.bold) {
      xml += `<a:tcTxStyle${p.bold ? ' b="on"' : ''}><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef>${p.text || ''}</a:tcTxStyle>`;
    }
    xml += '<a:tcStyle><a:tcBdr>';
    for (const side of ['left', 'right', 'top', 'bottom', 'insideH', 'insideV']) {
      if (p[side]) xml += `<a:${side}>${ln(p[side], p[side + 'W'])}</a:${side}>`;
    }
    xml += '</a:tcBdr>';
    if (p.fill) xml += `<a:fill>${solid(p.fill)}</a:fill>`;
    xml += '</a:tcStyle>';
    xml += `</a:${part}>`;
  }
  xml += '</a:tblStyle>';
  return xml;
}

function describe(family, accent) {
  const a = accent;
  const all = (c) => ({ left: c, right: c, top: c, bottom: c, insideH: c, insideV: c });
  switch (family) {
    case 'Themed1': {
      // "No Style, No Grid".
      if (!a) return { wholeTbl: { text: clr('tx1') } };
      const b = clr(a);
      return {
        wholeTbl: { text: clr('dk1'), ...all(b) },
        firstRow: { text: clr('lt1'), bold: true, fill: clr(a), left: b, right: b, top: b, bottom: clr('lt1') },
        lastRow: { bold: true, left: b, right: b, top: b, bottom: b },
        firstCol: { bold: true, left: b, right: b, top: b, bottom: b, insideH: b },
        lastCol: { bold: true, left: b, right: b, top: b, bottom: b, insideH: b },
        band1H: { fill: clr(a, 'alpha', 40000) },
        band1V: { fill: clr(a, 'alpha', 40000) },
      };
    }
    case 'Themed2': {
      if (!a) {
        const t = clr('tx1', 'tint', 50000);
        // "No Style, Table Grid".
        return { wholeTbl: { text: clr('tx1'), left: t, right: t, top: t, bottom: t, insideH: clr('tx1'), insideV: clr('tx1') } };
      }
      const t = clr(a, 'tint', 50000);
      return {
        tblBg: clr(a),
        wholeTbl: { text: clr('lt1'), left: t, right: t, top: t, bottom: t },
        firstRow: { text: clr('lt1'), bold: true, bottom: clr('lt1'), bottomW: 25400 },
        lastRow: { bold: true, top: clr('lt1'), topW: 25400 },
        firstCol: { bold: true, right: clr('lt1') },
        lastCol: { bold: true, left: clr('lt1') },
        band1H: { fill: clr('lt1', 'alpha', 20000) },
        band1V: { fill: clr('lt1', 'alpha', 20000) },
      };
    }
    case 'Light1': {
      const c = a || 'tx1';
      return {
        wholeTbl: { text: clr('tx1'), top: clr(c), bottom: clr(c) },
        firstRow: { text: clr('tx1'), bold: true, bottom: clr(c) },
        lastRow: { bold: true, top: clr(c) },
        firstCol: { bold: true },
        lastCol: { text: clr('tx1'), bold: true },
        band1H: { fill: clr(c, 'alpha', 20000) },
        band1V: { fill: clr(c, 'alpha', 20000) },
      };
    }
    case 'Light2': {
      const c = a || 'tx1';
      return {
        wholeTbl: { text: clr('tx1'), left: clr(c), right: clr(c), top: clr(c), bottom: clr(c) },
        firstRow: { text: clr('bg1'), bold: true, fill: clr(c) },
        lastRow: { bold: true, top: clr(c), topW: 38100 },
        firstCol: { bold: true },
        lastCol: { bold: true },
        band1H: { top: clr(c), bottom: clr(c) },
        band1V: { left: clr(c), right: clr(c) },
        band2V: { left: clr(c), right: clr(c) },
      };
    }
    case 'Light3': {
      const c = a || 'tx1';
      return {
        wholeTbl: { text: clr('tx1'), ...all(clr(c)) },
        firstRow: { text: clr(c), bold: true, bottom: clr(c), bottomW: 25400 },
        lastRow: { bold: true, top: clr(c), topW: 38100 },
        firstCol: { bold: true },
        lastCol: { bold: true },
        band1H: { fill: clr(c, 'alpha', 20000) },
        band1V: { fill: clr(c, 'alpha', 20000) },
      };
    }
    case 'Medium1': {
      const c = a || 'dk1';
      return {
        wholeTbl: { text: clr('dk1'), fill: clr('lt1'), left: clr(c), right: clr(c), top: clr(c), bottom: clr(c), insideH: clr(c) },
        firstRow: { text: clr('lt1'), bold: true, fill: clr(c) },
        lastRow: { bold: true, fill: clr('lt1'), top: clr(c), topW: 50800 },
        firstCol: { bold: true },
        lastCol: { bold: true },
        band1H: { fill: clr(c, 'tint', 20000) },
        band1V: { fill: clr(c, 'tint', 20000) },
      };
    }
    case 'Medium2': {
      const c = a || 'dk1';
      const w = clr('lt1');
      return {
        wholeTbl: { text: clr('dk1'), fill: clr(c, 'tint', 20000), ...all(w) },
        band1H: { fill: clr(c, 'tint', 40000) },
        band1V: { fill: clr(c, 'tint', 40000) },
        lastCol: { text: clr('lt1'), bold: true, fill: clr(c) },
        firstCol: { text: clr('lt1'), bold: true, fill: clr(c) },
        lastRow: { text: clr('lt1'), bold: true, fill: clr(c), top: w, topW: 38100 },
        firstRow: { text: clr('lt1'), bold: true, fill: clr(c), bottom: w, bottomW: 38100 },
      };
    }
    case 'Medium3': {
      const c = a || 'dk1';
      return {
        wholeTbl: { text: clr('dk1'), fill: clr('lt1'), top: clr('dk1'), topW: 25400, bottom: clr('dk1'), bottomW: 25400 },
        band1H: { fill: clr('dk1', 'tint', 20000) },
        band1V: { fill: clr('dk1', 'tint', 20000) },
        lastCol: { text: clr('lt1'), bold: true, fill: clr(c) },
        firstCol: { text: clr('lt1'), bold: true, fill: clr(c) },
        lastRow: { bold: true, fill: clr('lt1'), top: clr('dk1'), topW: 50800 },
        firstRow: { text: clr('lt1'), bold: true, fill: clr(c), bottom: clr('dk1'), bottomW: 25400 },
      };
    }
    case 'Medium4': {
      const c = a || 'dk1';
      return {
        wholeTbl: { text: clr('dk1'), fill: clr(c, 'tint', 20000), ...all(clr(c)) },
        band1H: { fill: clr(c, 'tint', 40000) },
        band1V: { fill: clr(c, 'tint', 40000) },
        lastCol: { bold: true },
        firstCol: { bold: true },
        lastRow: { bold: true, fill: clr('dk1', 'tint', 20000), top: clr('dk1'), topW: 25400 },
        firstRow: { text: clr(c), bold: true, fill: clr(c, 'tint', 20000) },
      };
    }
    case 'Dark1': {
      const c = a || 'dk1';
      const mod = a ? 'shade' : 'tint';
      return {
        wholeTbl: { text: clr('dk1'), fill: clr(c, mod, 20000) },
        band1H: { fill: clr(c, mod, 40000) },
        band1V: { fill: clr(c, mod, 40000) },
        lastCol: { text: clr('lt1'), bold: true, fill: clr(c, mod, 60000), left: clr('lt1'), leftW: 25400 },
        firstCol: { text: clr('lt1'), bold: true, fill: clr(c, mod, 60000), right: clr('lt1'), rightW: 25400 },
        lastRow: { text: clr('lt1'), bold: true, fill: clr(c, mod, 20000), top: clr('lt1'), topW: 25400 },
        firstRow: { text: clr('lt1'), bold: true, fill: clr('dk1'), bottom: clr('lt1'), bottomW: 38100 },
      };
    }
    case 'Dark2': {
      const c = a || 'dk1';
      const pair = { accent1: 'accent2', accent3: 'accent4', accent5: 'accent6' };
      return {
        wholeTbl: { text: clr('dk1'), fill: clr(c, 'tint', 20000) },
        band1H: { fill: clr(c, 'tint', 40000) },
        band1V: { fill: clr(c, 'tint', 40000) },
        lastCol: { bold: true },
        firstCol: { bold: true },
        lastRow: { bold: true, fill: clr(c, 'tint', 20000), top: clr('dk1'), topW: 25400 },
        firstRow: { text: clr('lt1'), bold: true, fill: clr(pair[a] || 'dk1') },
      };
    }
    default:
      return null;
  }
}

/** Return the XML of a built-in table style, or null. */
export function builtinTableStyle(id) {
  if (!id) return null;
  const key = id.toUpperCase();
  const spec = IDS[key];
  if (!spec) return null;
  const desc = describe(spec[0], spec[1]);
  return desc ? build(key, desc) : null;
}

export const BUILTIN_TABLE_STYLE_IDS = Object.keys(IDS);

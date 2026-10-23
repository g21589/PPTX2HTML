/** Standard OOXML templates (theme, master, layouts) for generated decks. */

export const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
export const NS_EXT = NS + ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main"';

export function themeXml(o = {}) {
  const c = {
    dk1: '<a:sysClr val="windowText" lastClr="000000"/>', lt1: '<a:sysClr val="window" lastClr="FFFFFF"/>',
    dk2: '<a:srgbClr val="44546A"/>', lt2: '<a:srgbClr val="E7E6E6"/>', accent1: '<a:srgbClr val="4472C4"/>',
    accent2: '<a:srgbClr val="ED7D31"/>', accent3: '<a:srgbClr val="A5A5A5"/>', accent4: '<a:srgbClr val="FFC000"/>',
    accent5: '<a:srgbClr val="5B9BD5"/>', accent6: '<a:srgbClr val="70AD47"/>', hlink: '<a:srgbClr val="0563C1"/>',
    folHlink: '<a:srgbClr val="954F72"/>', ...(o.colors || {}),
  };
  const major = o.major || 'Calibri Light';
  const minor = o.minor || 'Calibri';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${o.name || 'Office Theme'}"><a:themeElements>
<a:clrScheme name="Office">${Object.entries(c).map(([k, v]) => `<a:${k}>${v}</a:${k}>`).join('')}</a:clrScheme>
<a:fontScheme name="Office"><a:majorFont><a:latin typeface="${major}"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック Light"/><a:font script="Hant" typeface="新細明體"/></a:majorFont><a:minorFont><a:latin typeface="${minor}"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="游ゴシック"/><a:font script="Hant" typeface="新細明體"/></a:minorFont></a:fontScheme>
<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme>
</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
}

function lvl(n, o) {
  return `<a:lvl${n}pPr ${o.attrs || ''}>${o.inner || ''}<a:defRPr sz="${o.sz}" kern="1200">${o.fill || '<a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'}<a:latin typeface="${o.font || '+mn-lt'}"/><a:ea typeface="${o.ea || '+mn-ea'}"/><a:cs typeface="${o.cs || '+mn-cs'}"/></a:defRPr></a:lvl${n}pPr>`;
}

export function textStyleList(kind) {
  let s = '';
  for (let i = 1; i <= 9; i++) {
    if (kind === 'title') {
      if (i > 1) break;
      s += lvl(i, { sz: 4400, font: '+mj-lt', ea: '+mj-ea', cs: '+mj-cs', attrs: 'algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"', inner: '<a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/>' });
    } else if (kind === 'body') {
      const sz = [2800, 2400, 2000, 1800, 1800, 1800, 1800, 1800, 1800][i - 1];
      const marL = 228600 + (i - 1) * 457200;
      s += lvl(i, { sz, attrs: `marL="${marL}" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"`, inner: `<a:lnSpc><a:spcPct val="90000"/></a:lnSpc><a:spcBef><a:spcPts val="${i === 1 ? 1000 : 500}"/></a:spcBef><a:buFont typeface="Arial" panose="020B0604020202020204" pitchFamily="34" charset="0"/><a:buChar char="•"/>` });
    } else {
      s += lvl(i, { sz: 1800, attrs: `marL="${(i - 1) * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"` });
    }
  }
  return s;
}

const ph = (id, name, type, idx, x, y, cx, cy, body = '<a:bodyPr/>', lst = '<a:lstStyle/>', text = '') => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${type ? ` type="${type}"` : ''}${idx !== null ? ` idx="${idx}"` : ''}${type === 'dt' || type === 'ftr' || type === 'sldNum' ? ' sz="quarter"' : ''}/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody>${body}${lst}<a:p>${text}<a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;

export function masterXml(W, H, o = {}) {
  const m = 838200 * (W / 12192000);
  const tx = (k) => Math.round(k * (W / 12192000));
  const ty = (k) => Math.round(k * (H / 6858000));
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
${ph(2, 'Title Placeholder 1', 'title', null, Math.round(m), ty(365125), tx(10515600), ty(1325563), '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"><a:normAutofit/></a:bodyPr>', '<a:lstStyle/>', '<a:r><a:rPr lang="en-US"/><a:t>Click to edit Master title style</a:t></a:r>')}
${ph(3, 'Text Placeholder 2', 'body', 1, Math.round(m), ty(1825625), tx(10515600), ty(4351338), '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"><a:normAutofit/></a:bodyPr>', '<a:lstStyle/>', '<a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>Edit Master text styles</a:t></a:r>')}
${ph(4, 'Date Placeholder 3', 'dt', 2, Math.round(m), ty(6356350), tx(2743200), ty(365125), '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/>', '<a:lstStyle><a:lvl1pPr algn="l"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>', '<a:fld id="{B5E1B2B5-0F7B-4C38-9A3A-111111111111}" type="datetimeFigureOut"><a:rPr lang="en-US"/><a:t>1/1/2024</a:t></a:fld>')}
${ph(5, 'Footer Placeholder 4', 'ftr', 3, tx(4038600), ty(6356350), tx(4114800), ty(365125), '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/>', '<a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>')}
${ph(6, 'Slide Number Placeholder 5', 'sldNum', 4, tx(8610600), ty(6356350), tx(2743200), ty(365125), '<a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0" anchor="ctr"/>', '<a:lstStyle><a:lvl1pPr algn="r"><a:defRPr sz="1200"><a:solidFill><a:schemeClr val="tx1"><a:tint val="75000"/></a:schemeClr></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>', '<a:fld id="{B5E1B2B5-0F7B-4C38-9A3A-222222222222}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld>')}
${o.masterShapes || ''}
</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst>${o.layoutIds}</p:sldLayoutIdLst>
<p:txStyles><p:titleStyle>${textStyleList('title')}</p:titleStyle><p:bodyStyle>${textStyleList('body')}</p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>${textStyleList('other')}</p:otherStyle></p:txStyles></p:sldMaster>`;
}

export function layoutXml(kind, W, H) {
  const tx = (k) => Math.round(k * (W / 12192000));
  const ty = (k) => Math.round(k * (H / 6858000));
  const lay = (type, name, inner, extra = '') => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout ${NS} type="${type}" preserve="1"${extra}><p:cSld name="${name}"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${inner}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
  const lph = (id, name, type, idx, geom = '', body = '<a:bodyPr/>', lst = '<a:lstStyle/>') => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${type ? ` type="${type}"` : ''}${idx !== null ? ` idx="${idx}"` : ''}/></p:nvPr></p:nvSpPr><p:spPr>${geom}</p:spPr><p:txBody>${body}${lst}<a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`;
  const footers = lph(4, 'Date Placeholder 3', 'dt', 10) + lph(5, 'Footer Placeholder 4', 'ftr', 11) + lph(6, 'Slide Number Placeholder 5', 'sldNum', 12);
  switch (kind) {
    case 'title':
      return lay('title', 'Title Slide',
        lph(2, 'Title 1', 'ctrTitle', null, `<a:xfrm><a:off x="${tx(1524000)}" y="${ty(1122363)}"/><a:ext cx="${tx(9144000)}" cy="${ty(2387600)}"/></a:xfrm>`, '<a:bodyPr anchor="b"/>', '<a:lstStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="6000"/></a:lvl1pPr></a:lstStyle>') +
        lph(3, 'Subtitle 2', 'subTitle', 1, `<a:xfrm><a:off x="${tx(1524000)}" y="${ty(3602038)}"/><a:ext cx="${tx(9144000)}" cy="${ty(1655762)}"/></a:xfrm>`, '<a:bodyPr/>', '<a:lstStyle><a:lvl1pPr marL="0" indent="0" algn="ctr"><a:buNone/><a:defRPr sz="2400"/></a:lvl1pPr></a:lstStyle>') + footers);
    case 'content':
      return lay('obj', 'Title and Content', lph(2, 'Title 1', 'title', null) + lph(3, 'Content Placeholder 2', null, 1) + footers);
    case 'titleOnly':
      return lay('titleOnly', 'Title Only', lph(2, 'Title 1', 'title', null) + footers);
    case 'twoContent':
      return lay('twoObj', 'Two Content', lph(2, 'Title 1', 'title', null) +
        lph(3, 'Content Placeholder 2', null, 1, `<a:xfrm><a:off x="${tx(838200)}" y="${ty(1825625)}"/><a:ext cx="${tx(5181600)}" cy="${ty(4351338)}"/></a:xfrm>`) +
        lph(7, 'Content Placeholder 3', null, 2, `<a:xfrm><a:off x="${tx(6172200)}" y="${ty(1825625)}"/><a:ext cx="${tx(5181600)}" cy="${ty(4351338)}"/></a:xfrm>`) + footers);
    case 'blank':
    default:
      return lay('blank', 'Blank', footers);
  }
}

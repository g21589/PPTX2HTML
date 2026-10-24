import test from 'node:test';
import assert from 'node:assert/strict';
import zlib from 'node:zlib';
import { parseXml, decodeEntities } from '../../src/util/xml.js';
import { inflateRaw } from '../../src/util/inflate.js';
import { ZipArchive } from '../../src/util/zip.js';
import { zip } from '../fixtures/zipwriter.mjs';

test('xml: canonical namespace prefixes', () => {
  const doc = parseXml('<?xml version="1.0"?><x:sld xmlns:x="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:d="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:rel="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><x:cSld><d:t xml:space="preserve"> a &amp; b &#x41;&#66; </d:t><x:pic rel:embed="rId1"/></x:cSld></x:sld>');
  assert.equal(doc.name, 'p:sld');
  const t = doc.path('p:cSld', 'a:t');
  assert.equal(t.text, ' a & b AB ');
  assert.equal(doc.path('p:cSld', 'p:pic').attrs['r:embed'], 'rId1');
});

test('xml: default namespace, CDATA, comments, mc:Choice Requires', () => {
  const doc = parseXml('<root xmlns="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:q="http://schemas.microsoft.com/office/powerpoint/2010/main"><!-- c --><t><![CDATA[<x>]]></t><mc:AlternateContent><mc:Choice Requires="q"/></mc:AlternateContent></root>');
  assert.equal(doc.name, 'a:root');
  assert.equal(doc.child('a:t').text, '<x>');
  assert.equal(doc.first('mc:Choice').attrs.Requires, 'p14');
  assert.equal(decodeEntities('&lt;&gt;&quot;&apos;&#9731;'), '<>"\'☃');
});

test('inflate matches zlib on random data', () => {
  for (let k = 0; k < 40; k++) {
    const n = Math.floor(Math.random() * 50000);
    const src = Buffer.alloc(n);
    for (let i = 0; i < n; i++) src[i] = k % 2 ? Math.random() * 256 : (i * 7) % 13;
    const comp = zlib.deflateRawSync(src, { level: k % 10 });
    assert.deepEqual(Buffer.from(inflateRaw(new Uint8Array(comp))), src);
  }
});

test('inflate rejects truncated data instead of hanging', () => {
  const comp = zlib.deflateRawSync(Buffer.alloc(100000, 7));
  assert.throws(() => inflateRaw(new Uint8Array(comp.subarray(0, 20))));
});

test('zip: read entries, case-insensitive lookup, recovery without central directory', () => {
  const buf = zip([{ name: 'a.txt', data: 'hello' }, { name: 'Dir/B.xml', data: '<b/>'.repeat(200) }]);
  const z = new ZipArchive(buf);
  assert.equal(z.text('a.txt'), 'hello');
  assert.equal(z.text('/dir/b.xml').length, 800);
  const cdStart = buf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]), buf.length, 'binary');
  const firstCd = buf.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]));
  void cdStart;
  const truncated = buf.subarray(0, firstCd);
  const z2 = new ZipArchive(truncated);
  assert.equal(z2.text('a.txt'), 'hello');
  // Central directory offset pointing outside the file: local headers are used.
  const damaged = Buffer.from(buf);
  const eocd = damaged.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  damaged.writeUInt32LE(0x7ffffff0, eocd + 16);
  const z3 = new ZipArchive(damaged);
  assert.equal(z3.text('a.txt'), 'hello');
  assert.equal(z3.text('Dir/B.xml').length, 800);
});

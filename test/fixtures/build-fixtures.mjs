#!/usr/bin/env node
/**
 * Generates the feature fixture decks into test/fixtures/out and writes a
 * manifest (features -> deck, slide, checks) used by the verifier.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Deck } from './builder.mjs';
import * as drawing from './catalog/drawing.mjs';
import * as text from './catalog/text.mjs';
import * as objects from './catalog/objects.mjs';
import * as animation from './catalog/animation.mjs';

export const CATALOGS = [drawing, text, objects, animation];
const here = path.dirname(fileURLToPath(import.meta.url));

/** Build all fixture decks; returns [{name, title, file, features}]. */
export function buildFixtures(outDir = path.join(here, 'out')) {
  fs.mkdirSync(outDir, { recursive: true });
  const decks = [];
  for (const cat of CATALOGS) {
    const deck = new Deck({ title: cat.title });
    const features = cat.build(deck);
    const file = path.join(outDir, cat.name + '.pptx');
    fs.writeFileSync(file, deck.build());
    decks.push({ name: cat.name, title: cat.title, file, features, slideCount: deck.slides.length });
  }
  return decks;
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const decks = buildFixtures(process.argv[2]);
  for (const d of decks) console.log(`${d.name}: ${d.slideCount} slides, ${d.features.length} features -> ${path.relative(process.cwd(), d.file)}`);
}

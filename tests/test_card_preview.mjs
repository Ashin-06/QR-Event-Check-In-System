import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { spawnSync } from 'node:child_process';
const dashboard = await readFile(new URL('../templates/dashboard.html', import.meta.url), 'utf8');
const scripts = [...dashboard.matchAll(/<script type="module">([\s\S]*?)<\/script>/g)];
assert.equal(scripts.length, 1);
const syntax = spawnSync(process.execPath, ['--input-type=module', '--check'], {input:scripts[0][1], encoding:'utf8', timeout:10000});
assert.equal(syntax.status, 0, syntax.stderr);
console.log('PASS: dashboard JavaScript syntax');
const source = await readFile(new URL('../static/js/card-preview.js', import.meta.url), 'utf8');
const { createCardPreview } = await import('data:text/javascript;base64,' + Buffer.from(source).toString('base64'));
const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
const revoked = [];
let sequence = 0;
URL.createObjectURL = () => `blob:test-${++sequence}`;
URL.revokeObjectURL = url => revoked.push(url);
const success = () => ({ok: true, headers: new Headers({'content-type': 'image/png'}), blob: async () => new Blob(['image'])});
const controls = () => ({image: {style: {}, hidden: true}, status: {style: {}, setAttribute() {}}, getRequest: () => ({url: '/preview', body: {mode:'custom'}})});

{
  const ui = controls();
  let calls = 0;
  const preview = createCardPreview({...ui, fetcher: async () => { calls++; return success(); }});
  preview.refresh(); preview.refresh(); preview.refresh();
  await sleep(300);
  assert.equal(calls, 1, 'Rapid edits should only request one preview');
  assert.equal(ui.image.hidden, false);
  const first = ui.image.src;
  preview.refresh();
  await sleep(300);
  assert(revoked.includes(first), 'Superseded images must be released');
  preview.dispose();
  assert(revoked.includes(ui.image.src), 'Disposal must release the last image');
  console.log('PASS: debounce and object URL cleanup');
}
{
  const ui = controls();
  ui.image.hidden = false;
  const preview = createCardPreview({...ui, fetcher: async () => ({ok:false, headers:new Headers(), json:async()=>({message:'Please sign in again.'})})});
  preview.refresh();
  await sleep(300);
  assert.equal(ui.image.hidden, true);
  assert.equal(ui.status.textContent, 'Please sign in again.');
  preview.dispose();
  console.log('PASS: failed requests show an error, not a stale image');
}
{
  const ui = controls();
  const pending = [], signals = [];
  const preview = createCardPreview({...ui, fetcher: (url, options) => { signals.push(options.signal); return new Promise(resolve => pending.push(resolve)); }});
  preview.refresh(); await sleep(300);
  preview.refresh(); await sleep(300);
  assert.equal(signals[0].aborted, true);
  pending[1](success()); await sleep(0);
  const latest = ui.image.src;
  pending[0](success()); await sleep(0);
  assert.equal(ui.image.src, latest, 'An older response must never replace the latest design');
  preview.dispose();
  console.log('PASS: cancellation and out-of-order responses');
}

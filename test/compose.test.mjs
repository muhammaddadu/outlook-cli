import { test } from 'node:test';
import assert from 'node:assert/strict';
import { textToHtml, prependToBody, resolveDraftStyle, DEFAULT_DRAFT_STYLE } from '../src/compose.mjs';

test('textToHtml: one paragraph per line, blank lines kept, edges trimmed', () => {
  const html = textToHtml('\nHi,\n\nLine two\n\n', 'X');
  assert.equal(
    html,
    '<p style="margin:0px; X">Hi,</p><p style="margin:0px; X"><br></p><p style="margin:0px; X">Line two</p>',
  );
});

test('textToHtml: bullets become a list, HTML is escaped', () => {
  const html = textToHtml('Points:\n- a & b\n* <c>\nDone', 'X');
  assert.match(html, /<ul style="margin:0px"><li style="X">a &amp; b<\/li><li style="X">&lt;c&gt;<\/li><\/ul>/);
  assert.match(html, /<p style="margin:0px; X">Done<\/p>$/);
});

test('prependToBody inserts after <body> with a gap paragraph', () => {
  const out = prependToBody('<html><body class="a"><hr>quote</body></html>', '<p>new</p>', 'X');
  assert.equal(out, '<html><body class="a"><p>new</p><p style="margin:0px; X"><br></p><hr>quote</body></html>');
});

test('resolveDraftStyle: explicit, then env, then default', () => {
  const prev = process.env.OUTLOOK_DRAFT_STYLE;
  try {
    delete process.env.OUTLOOK_DRAFT_STYLE;
    assert.equal(resolveDraftStyle(), DEFAULT_DRAFT_STYLE);
    process.env.OUTLOOK_DRAFT_STYLE = 'font-family:Env';
    assert.equal(resolveDraftStyle(), 'font-family:Env');
    assert.equal(resolveDraftStyle('font-family:Flag'), 'font-family:Flag');
  } finally {
    if (prev === undefined) delete process.env.OUTLOOK_DRAFT_STYLE;
    else process.env.OUTLOOK_DRAFT_STYLE = prev;
  }
});

test('replaceReplyText swaps only the text above the reply header', async () => {
  const { replaceReplyText } = await import('../src/compose.mjs');
  const html = '<html><body><div>old one-line text</div><hr style="x"><div id="divRplyFwdMsg">From: A</div>quote</body></html>';
  const out = replaceReplyText(html, '<p>new</p>', 'X');
  assert.equal(out, '<html><body><p>new</p><p style="margin:0px; X"><br></p><hr style="x"><div id="divRplyFwdMsg">From: A</div>quote</body></html>');
  assert.equal(replaceReplyText('<html><body>no header</body></html>', '<p>n</p>'), null);
});

/**
 * Compose helpers: turn plain-text draft bodies into Outlook-web-style HTML.
 *
 * Why: Outlook's `Comment` field on createReply/createForward is plain text,
 * and the server drops its line breaks when it renders the HTML draft, so a
 * multi-line reply shows up as one line. We instead build the HTML ourselves,
 * matching what Outlook web writes when you type: one <p style="margin:0px">
 * per line, <p><br></p> for blank lines, and real <ul>/<li> for "- " bullets.
 */

/** Outlook web's default compose font (Aptos 12pt, black). */
export const DEFAULT_DRAFT_STYLE =
  'font-family:Aptos,Arial,Helvetica,sans-serif; font-size:12pt; color:rgb(0,0,0)';

/** Style from --style, else $OUTLOOK_DRAFT_STYLE, else the Outlook default. */
export function resolveDraftStyle(explicit) {
  const s = (explicit ?? process.env.OUTLOOK_DRAFT_STYLE ?? '').trim();
  return s || DEFAULT_DRAFT_STYLE;
}

export function escapeHtml(s) {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

const BULLET = /^\s*[-*•]\s+/;

/**
 * Plain text → HTML fragment. Each line becomes a paragraph; blank lines
 * become an empty paragraph; consecutive "- " / "* " / "• " lines become a list.
 */
export function textToHtml(text, style = DEFAULT_DRAFT_STYLE) {
  const lines = String(text).replace(/\r\n?/g, '\n').split('\n');
  // Trim leading/trailing blank lines; keep inner ones.
  while (lines.length && !lines[0].trim()) lines.shift();
  while (lines.length && !lines[lines.length - 1].trim()) lines.pop();

  const p = (inner) => `<p style="margin:0px; ${style}">${inner}</p>`;
  const out = [];
  let list = null;
  const flush = () => {
    if (list) {
      out.push(`<ul style="margin:0px">${list.join('')}</ul>`);
      list = null;
    }
  };
  for (const line of lines) {
    if (BULLET.test(line)) {
      list ??= [];
      list.push(`<li style="${style}">${escapeHtml(line.replace(BULLET, ''))}</li>`);
      continue;
    }
    flush();
    out.push(line.trim() ? p(escapeHtml(line)) : p('<br>'));
  }
  flush();
  return out.join('');
}

/**
 * Insert an HTML fragment at the top of an existing HTML document's <body>
 * (keeping the quoted thread below it), followed by one blank paragraph —
 * the same gap Outlook web leaves above the reply header.
 */
export function prependToBody(existingHtml, fragment, style = DEFAULT_DRAFT_STYLE) {
  const block = `${fragment}<p style="margin:0px; ${style}"><br></p>`;
  const html = existingHtml ?? '';
  const m = html.match(/<body[^>]*>/i);
  if (!m) return `<html><body>${block}${html}</body></html>`;
  const at = m.index + m[0].length;
  return html.slice(0, at) + block + html.slice(at);
}

/** Body override (Text or HTML) → HTML fragment in the given style. */
export function bodyToFragment(body, style = DEFAULT_DRAFT_STYLE) {
  const type = String(body.ContentType ?? 'Text').toLowerCase();
  return type === 'html' ? String(body.Content) : textToHtml(body.Content, style);
}

/**
 * Replace the new-text part of a reply/forward draft — everything between
 * <body> and the reply header (Outlook web's <hr> / #divRplyFwdMsg) — with a
 * fresh fragment, keeping the quoted thread. Returns null when the draft has
 * no reply header (the caller decides what to do).
 */
export function replaceReplyText(existingHtml, fragment, style = DEFAULT_DRAFT_STYLE) {
  const html = existingHtml ?? '';
  const body = html.match(/<body[^>]*>/i);
  if (!body) return null;
  const start = body.index + body[0].length;
  const rest = html.slice(start);
  const marks = [rest.search(/<hr[\s>]/i), rest.search(/<div[^>]*id="?divRplyFwdMsg/i)].filter(
    (i) => i >= 0,
  );
  if (!marks.length) return null;
  const end = start + Math.min(...marks);
  return html.slice(0, start) + `${fragment}<p style="margin:0px; ${style}"><br></p>` + html.slice(end);
}

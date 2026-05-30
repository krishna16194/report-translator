// Translation via Google's free (unofficial) endpoint, called with fetch.
//
// Workers cap each request at ~50 subrequests and ~10ms CPU, so we:
//   * de-duplicate strings,
//   * batch many short strings into one request separated by a sentinel,
//   * cap the total number of fetches (subrequests).
// Anything left over when the budget is exhausted is returned untranslated.

const ENDPOINT = "https://translate.googleapis.com/translate_a/single";
const SENTINEL = "\nZZZSEP\n";          // unlikely to be altered by translation
const SPLIT_RE = /\s*ZZZSEP\s*/;
const MAX_BATCH_CHARS = 3500;
const MAX_SINGLE_CHARS = 4500;
const MAX_FETCHES = 45;                  // stay safely under the 50 subrequest cap

async function translateChunk(text, target) {
  const url =
    `${ENDPOINT}?client=gtx&sl=auto&tl=${encodeURIComponent(target)}` +
    `&dt=t&q=${encodeURIComponent(text)}`;
  const res = await fetch(url, { headers: { "User-Agent": "Mozilla/5.0" } });
  if (!res.ok) throw new Error(`translate HTTP ${res.status}`);
  const data = await res.json();
  // data[0] is an array of [translatedSegment, originalSegment, ...].
  return (data[0] || []).map((seg) => seg[0]).join("");
}

export async function translateMany(strings, target) {
  const result = new Map();
  const unique = [...new Set(strings.filter((s) => s && s.trim()))];
  let fetches = 0;

  // Build batches that fit within the per-request character budget.
  const batches = [];
  let cur = [];
  let curLen = 0;
  for (const s of unique) {
    if (s.length >= MAX_BATCH_CHARS) {
      batches.push([s]); // oversized item handled on its own
      continue;
    }
    if (cur.length && curLen + s.length > MAX_BATCH_CHARS) {
      batches.push(cur);
      cur = [];
      curLen = 0;
    }
    cur.push(s);
    curLen += s.length + SENTINEL.length;
  }
  if (cur.length) batches.push(cur);

  for (const batch of batches) {
    if (fetches >= MAX_FETCHES) break; // budget exhausted; leave rest untranslated

    // Single oversized string: chunk it and translate sequentially.
    if (batch.length === 1 && batch[0].length >= MAX_BATCH_CHARS) {
      const original = batch[0];
      const pieces = chunkText(original, MAX_SINGLE_CHARS);
      const out = [];
      let ok = true;
      for (const piece of pieces) {
        if (fetches >= MAX_FETCHES) { ok = false; break; }
        fetches++;
        try {
          out.push(await translateChunk(piece, target));
        } catch {
          ok = false;
          break;
        }
      }
      if (ok) result.set(original, out.join(""));
      continue;
    }

    // Normal batch: one fetch, split on the sentinel.
    fetches++;
    try {
      const joined = batch.join(SENTINEL);
      const translated = await translateChunk(joined, target);
      const parts = translated.split(SPLIT_RE);
      if (parts.length === batch.length) {
        batch.forEach((orig, i) => result.set(orig, parts[i].trim()));
      } else {
        // Sentinel got mangled — fall back to per-item within budget.
        for (const orig of batch) {
          if (fetches >= MAX_FETCHES) break;
          fetches++;
          try {
            result.set(orig, await translateChunk(orig, target));
          } catch {
            /* leave untranslated */
          }
        }
      }
    } catch {
      /* whole batch failed; leave untranslated */
    }
  }

  return result;
}

function chunkText(text, limit) {
  const chunks = [];
  let rest = text;
  while (rest.length > limit) {
    let at = rest.lastIndexOf(" ", limit);
    if (at < limit / 2) at = limit;
    chunks.push(rest.slice(0, at));
    rest = rest.slice(at);
  }
  if (rest) chunks.push(rest);
  return chunks;
}

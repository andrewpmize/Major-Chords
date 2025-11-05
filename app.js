// ---------- helpers ----------
function rootToken(s) {
  if (!s) return "";
  const m = String(s).trim().match(/^[A-Ga-g](?:#|b)?/);
  return m ? m[0].toUpperCase() : "";
}
function parseTable(text) {
  const lines = text.trim().split(/\r?\n/).filter(Boolean);
  if (!lines.length) return { headers: [], rows: [] };
  const split = (line) => line.split(/\t|,|;/).map(x => x.trim());
  const headers = split(lines[0]);
  const rows = lines.slice(1).map(split);
  return { headers, rows };
}
function findKeyCol(headers, key) {
  return headers.findIndex(h => String(h).trim().toUpperCase() === String(key).trim().toUpperCase());
}
const safe = (s) => String(s ?? "").replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));

// --- pitch helpers for vii° and V/vii°
const SHARP_SCALE = ['C','C#','D','D#','E','F','F#','G','G#','A','A#','B'];
const FLAT_EQUIV  = { 'A#':'Bb','D#':'Eb','G#':'Ab' }; // minimal enharmonics we need
const FLAT_KEYS   = new Set(['F','Bb','Eb','Ab','Db','Gb']); // use flats in these keys

function toIndex(root) {
  const R = root.toUpperCase();
  const i = SHARP_SCALE.indexOf(R);
  if (i >= 0) return i;
  // allow flat inputs (Bb -> A#)
  if (/^[A-G]b$/.test(R)) {
    const sharp = Object.keys(FLAT_EQUIV).find(k => FLAT_EQUIV[k] === R);
    return SHARP_SCALE.indexOf(sharp);
  }
  return -1;
}
function nameFromIndex(i, preferFlats=false) {
  let n = SHARP_SCALE[(i+12)%12];
  if (preferFlats && FLAT_EQUIV[n]) n = FLAT_EQUIV[n];
  return n;
}
function transpose(root, semitones, preferFlats=false) {
  const i = toIndex(root);
  return i < 0 ? '' : nameFromIndex(i + semitones, preferFlats);
}


// ---------- core logic: port of Chord_File_Major ----------
function chordFileMajor_JS(key, CA) {
  const { headers, rows } = CA;
  const col = findKeyCol(headers, key);
  if (col < 0) throw new Error(`Key '${key}' not found in CA headers.`);

  const CAget = (row1Based, col0Based) => {
    const r = row1Based - 5; // first data row corresponds to VBA row 5
    const c = col0Based;
    if (r < 0 || r >= rows.length) return "";
    if (c < 0 || c >= headers.length) return "";
    return rows[r][c] ?? "";
  };

  const One   = CAget(5,  col);
  const Two   = CAget(6,  col);
  const Three = CAget(7,  col);
  const Four  = CAget(8,  col);
  const Five  = CAget(9,  col);
  const Six   = CAget(10, col);
 
  // --- compute vii° and V/vii° from theory
  const preferFlats = FLAT_KEYS.has(String(key).replace(/\s+.*$/,''));
  const sevenRoot   = transpose(key, -1, preferFlats);      // leading tone (−1 semitone)
  const SevenDim    = sevenRoot ? `${sevenRoot}°` : '';     // vii° triad label
  const FiveSeven   = sevenRoot ? `${transpose(sevenRoot, 7, preferFlats)}7` : ''; // V of vii°

  const FlatThree = CAget(14, col);
  const FlatFour  = CAget(15, col);
  const FlatSix   = CAget(17, col);
  const FlatSeven = CAget(18, col);

  const headerRoot = CA.headers ? CA.headers.map(h => rootToken(h)) : headers.map(h => rootToken(h));

  function twoStep(row1Based) {
    const src = CAget(row1Based, col);
    const r1  = rootToken(src);
    const matchCol = headerRoot.findIndex(hr => hr === r1);
    if (matchCol < 0) return "";
    const vChord = CAget(9, matchCol);   // row 9 holds the V chord for that column
    const vRoot  = rootToken(vChord);
    return vRoot ? (vRoot + "7") : "";
  }

 return {
  One, Two, Three, Four, Five, Six,
  FlatThree, FlatFour, FlatSix, FlatSeven,
  FiveOne: twoStep(5),
  FiveTwo: twoStep(6),
  FiveThree: twoStep(7),
  FiveFour: twoStep(8),
  FiveFive: twoStep(9),
  FiveSix: twoStep(10),

  // NEW CHORDS:
  SevenDim,   // this is the vii° chord
  FiveSeven   // this is the V of vii°
};

}

// ---------- built-in CA grid (demo / default) ----------
const DEMO_GRID = `C,D,E,F,G,A,B,C#,D#,F#,G#,Bb
C maj,D maj,E maj,F maj,G maj,A maj,B maj,C# maj,D# maj,F# maj,G# maj,Bb maj
D min,E min,F# min,G min,A min,B min,C# min,D# min,F min,G# min,A# min,C min
E min,F# min,G# min,A min,B min,C# min,D# min,F min,G min,A# min,C min,D min
F maj,G maj,A maj,Bb maj,C maj,D maj,E maj,F# maj,G# maj,B maj,C# maj,Eb maj
G maj,A maj,B maj,C maj,D maj,E maj,F# maj,G# maj,A# maj,C# maj,D# maj,F maj
A min,B min,C# min,D min,E min,F# min,G# min,A# min,C min,D# min,F min,G min
,,,,,,,,,,,,
,,,,,,,,,,,,
,,,,,,,,,,,,
Eb maj,F maj,G maj,Ab maj,Bb maj,C maj,D maj,E maj,F# maj,A maj,B maj,Db maj
F min,G min,A min,Bb min,C min,D min,E min,F# min,G# min,B min,C# min,Eb min
,,,,,,,,,,,,
Ab maj,Bb maj,C maj,Db maj,Eb maj,F maj,G maj,A maj,B maj,D maj,E maj,Gb maj
Bb maj,C maj,D maj,Eb maj,F maj,G maj,A maj,B maj,C# maj,E maj,F# maj,Ab maj`;


// ---------- UI wiring ----------
window.addEventListener('DOMContentLoaded', () => {
  const run = () => {
    const key = document.getElementById('key').value;
    const CA  = parseTable(DEMO_GRID);
    const out = chordFileMajor_JS(key, CA);
    paintBands(out);
  };
  document.getElementById('runBtn').addEventListener('click', run);
  run(); // populate once on load
});

// Fill the slots inside each band
function paintBands(o) {
  document.querySelectorAll('[data-slot]').forEach(el => {
    const k = el.getAttribute('data-slot');
    el.textContent = safe(o[k] ?? '');
  });
}

// ===== Simple "Match and Paste Chords (horizontal)" web version =====

// Use the same key list as your dropdown
const MAJOR_KEYS = ['C','D','E','F','G','A','B','C#','D#','F#','G#','Bb'];

// Normalize a chord token into a coarse label: root + (m/°) when clear.
// Treat maj/7/sus/add/etc. as major-ish (no 'm') for fit purposes.
// Examples: "Em7" -> "Em", "G7" -> "G", "C°" -> "C°", "Bbmaj7" -> "Bb"
function normalizeChord(tok) {
  if (!tok) return '';
  const m = tok.trim().match(/^([A-Ga-g](?:#|b)?)(.*)$/);
  if (!m) return '';
  const preferFlatsGuess = /b/.test(m[1]); // if user typed Bb, keep flats
  let root = nameFromIndex(toIndex(m[1]), preferFlatsGuess) || m[1].toUpperCase();
  const tail = (m[2] || '').toLowerCase();
  if (/dim|°/.test(tail)) return root + '°';
  if (/m(?!aj)/.test(tail)) return root + 'm';
  return root;
}

// Build the diatonic set & roman maps for a major key
function diatonicForKey(key) {
  const preferFlats = FLAT_KEYS.has(key);
  const I   = key;
  const ii  = transpose(key, 2,  preferFlats) + 'm';
  const iii = transpose(key, 4,  preferFlats) + 'm';
  const IV  = transpose(key, 5,  preferFlats);
  const V   = transpose(key, 7,  preferFlats);
  const vi  = transpose(key, 9,  preferFlats) + 'm';
  const vii = transpose(key, 11, preferFlats) + '°';
  const dia = new Set([I, ii, iii, IV, V, vi, vii]);

  const romanByChord = {
    [I]:'I', [ii]:'ii', [iii]:'iii', [IV]:'IV', [V]:'V', [vi]:'vi', [vii]:'vii°'
  };
  const chordByRoman = { I, ii, iii, IV, V, vi, 'vii°':vii };

  return { dia, romanByChord, chordByRoman, preferFlats };
}

// Some near-diatonic “credit”: common secondary dominants + modal borrows
function nearHitsForKey(key) {
  const { preferFlats } = diatonicForKey(key);
  return new Set([
    transpose(key,  2, preferFlats) + '7', // V/ii
    transpose(key,  4, preferFlats) + '7', // V/iii
    transpose(key,  5, preferFlats) + '7', // V/IV
    transpose(key,  7, preferFlats) + '7', // V/V
    transpose(key,  9, preferFlats) + '7', // V/vi
    transpose(key, -1, preferFlats) + '7', // V/vii°
    transpose(key,  3, preferFlats),       // ♭III
    transpose(key,  8, preferFlats) + 'm', // iv
    transpose(key,  8, preferFlats),       // ♭VI
    transpose(key, 10, preferFlats)        // ♭VII
  ].map(normalizeChord));
}

// Score a key against a list of normalized chords
function scoreKeyAgainstChords(key, chordsNorm) {
  const { dia } = diatonicForKey(key);
  const near = nearHitsForKey(key);
  let hit = 0, nearHit = 0;
  for (const c of chordsNorm) {
    if (!c) continue;
    if (dia.has(c)) hit++;
    else if (near.has(c)) nearHit++;
  }
  const denom = Math.max(1, chordsNorm.length);
  const pct = Math.round(((hit + 0.5 * nearHit) / denom) * 100);
  return { key, hit, near: nearHit, pct };
}

// Romanize a sequence for a given key (unknowns become '?')
function romanizeSequence(chordsNorm, key) {
  const { romanByChord } = diatonicForKey(key);
  return chordsNorm.map(c => romanByChord[c] || '?');
}

// Very light progression detector over roman numerals
function detectProgressions(romans) {
  const seq = romans.filter(x => x !== '?'); // ignore unknowns
  const found = [];
  const has = (...pat) => {
    for (let i = 0; i <= seq.length - pat.length; i++) {
      let ok = true;
      for (let j = 0; j < pat.length; j++) {
        if (seq[i + j] !== pat[j]) { ok = false; break; }
      }
      if (ok) return true;
    }
    return false;
  };

  if (has('ii','V','I')) found.push('ii–V–I');
  if (has('I','V','vi','IV')) found.push('I–V–vi–IV');
  if (has('I','vi','IV','V')) found.push('I–vi–IV–V');
  if (has('I','IV','V')) found.push('I–IV–V');
  if (has('vi','IV','I','V')) found.push('vi–IV–I–V');
  if (has('IV','V','I')) found.push('IV–V–I');

  return [...new Set(found)];
}

// Render helpers
function renderTopMatches(results) {
  return `
    <div class="panel">
      <strong>Top matches:</strong>
      <ul style="margin:6px 0 0 16px;">
        ${results.map(r => `<li>${r.key} major — ${r.pct}% (hits ${r.hit}, near ${r.near})</li>`).join('')}
      </ul>
    </div>`;
}
function renderMappingTable(bestKey) {
  const { chordByRoman } = diatonicForKey(bestKey);
  const order = ['I','ii','iii','IV','V','vi','vii°'];
  return `
    <div class="panel">
      <strong>${bestKey} major — diatonic map</strong>
      <div class="grid grid-7" style="margin-top:8px;">
        ${order.map(r => `<div class="box"><span class="roman">${r}</span><span>${chordByRoman[r]}</span></div>`).join('')}
      </div>
    </div>`;
}
function renderProgHints(bestKey, romans) {
  const hits = detectProgressions(romans);
  if (!hits.length) return '';
  return `
    <div class="panel">
      <strong>Common progressions spotted (as Roman numerals in ${bestKey}):</strong>
      <p style="margin:6px 0 0">${hits.join(', ')}</p>
    </div>`;
}

// Wire up the Analyze button
window.addEventListener('DOMContentLoaded', () => {
  const btn = document.getElementById('analyzeBtn');
  if (!btn) return;

  btn.addEventListener('click', () => {
    const raw = (document.getElementById('songChords').value || '')
      .split(/\s+/)
      .map(normalizeChord)
      .filter(Boolean);

    if (!raw.length) {
      document.getElementById('keyResults').innerHTML = '<em>No chords provided.</em>';
      document.getElementById('mappingTable').innerHTML = '';
      document.getElementById('progHints').innerHTML = '';
      return;
    }

    const results = MAJOR_KEYS
      .map(k => scoreKeyAgainstChords(k, raw))
      .sort((a,b) => b.pct - a.pct)
      .slice(0, 5);

    const best = results[0].key;
    const romans = romanizeSequence(raw, best);

    document.getElementById('keyResults').innerHTML  = renderTopMatches(results);
    document.getElementById('mappingTable').innerHTML = renderMappingTable(best);
    document.getElementById('progHints').innerHTML   = renderProgHints(best, romans);
  });
});
// ===== Simple "Match and Paste Chords (horizontal)" web version =====
const MAJOR_KEYS = ['C','D','E','F','G','A','B','C#','D#','F#','G#','Bb'];

function normalizeChord(tok) {
  if (!tok) return '';
  const m = tok.trim().match(/^([A-Ga-g](?:#|b)?)(.*)$/);
  if (!m) return '';
  const preferFlatsGuess = /b/.test(m[1]);
  let root = nameFromIndex(toIndex(m[1]), preferFlatsGuess) || m[1].toUpperCase();
  const tail = (m[2] || '').toLowerCase();
  if (/dim|°/.test(tail)) return root + '°';
  if (/m(?!aj)/.test(tail)) return root + 'm';
  return root;
}

function diatonicForKey(key) {
  const preferFlats = FLAT_KEYS.has(key);
  const I   = key;
  const ii  = transpose(key, 2,  preferFlats) + 'm';
  const iii = transpose(key, 4,  preferFlats) + 'm';
  const IV  = transpose(key, 5,  preferFlats);
  const V   = transpose(key, 7,  preferFlats);
  const vi  = transpose(key, 9,  preferFlats) + 'm';
  const vii = transpose(key, 11, preferFlats) + '°';
  const dia = new Set([I, ii, iii, IV, V, vi, vii]);
  const romanByChord = { [I]:'I', [ii]:'ii', [iii]:'iii', [IV]:'IV', [V]:'V', [vi]:'vi', [vii]:'vii°' };
  const chordByRoman = { I, ii, iii, IV, V, vi, 'vii°':vii };
  return { dia, romanByChord, chordByRoman, preferFlats };
}

function nearHitsForKey(key) {
  const { preferFlats } = diatonicForKey(key);
  return new Set([
    transpose(key,  2, preferFlats) + '7',
    transpose(key,  4, preferFlats) + '7',
    transpose(key,  5, preferFlats) + '7',
    transpose(key,  7, preferFlats) + '7',
    transpose(key,  9, preferFlats) + '7',
    transpose(key, -1, preferFlats) + '7',
    transpose(key,  3, preferFlats),
    transpose(key,  8, preferFlats) + 'm',
    transpose(key,  8, preferFlats),
    transpose(key, 10, preferFlats)
  ].map(normalizeChord));
}

function scoreKeyAgainstChords(key, chordsNorm) {
  const { dia } = diatonicForKey(key);
  const near = nearHitsForKey(key);
  let hit = 0, nearHit = 0;
  for (const c of chordsNorm) {
    if (!c) continue;
    if (dia.has(c)) hit++;
    else if (near.has(c)) nearHit++;
  }
  const denom = Math.max(1, chordsNorm.length);
  const pct = Math.round(((hit + 0.5 * nearHit) / denom) * 100);
  return { key, hit, near: nearHit, pct };
}

function romanizeSequence(chordsNorm, key) {
  const { romanByChord } = diatonicForKey(key);
  return chordsNorm.map(c => romanByChord[c] || '?');
}

function renderTopMatches(results) {
  return `
    <div class="panel">
      <strong>Top matches:</strong>
      <ul style="margin:6px 0 0 16px;">
        ${results.map(r => `<li>${r.key} major — ${r.pct}% (hits ${r.hit}, near ${r.near})</li>`).join('')}
      </ul>
    </div>`;
}
function renderMappingTable(bestKey) {
  const { chordByRoman } = diatonicForKey(bestKey);
  const order = ['I','ii','iii','IV','V','vi','vii°'];
  return `
    <div class="panel">
      <strong>${bestKey} major — diatonic map</strong>
      <div class="grid grid-7" style="margin-top:8px;">
        ${order.map(r => `<div class="box"><span class="roman">${r}</span><span>${chordByRoman[r]}</span></div>`).join('')}
      </div>
    </div>`;
}
function detectProgressions(romans) {
  const seq = romans.filter(x => x !== '?');
  const found = [];
  const has = (...pat) => {
    for (let i = 0; i <= seq.length - pat.length; i++) {
      let ok = true;
      for (let j = 0; j < pat.length; j++) if (seq[i+j] !== pat[j]) { ok=false; break; }
      if (ok) return true;
    }
    return false;
  };
  if (has('ii','V','I')) found.push('ii–V–I');
  if (has('I','V','vi','IV')) found.push('I–V–vi–IV');
  if (has('I','vi','IV','V')) found.push('I–vi–IV–V');
  if (has('I','IV','V')) found.push('I–IV–V');
  if (has('vi','IV','I','V')) found.push('vi–IV–I–V');
  if (has('IV','V','I')) found.push('IV–V–I');
  return [...new Set(found)];
}
function renderProgHints(bestKey, romans) {
  const hits = detectProgressions(romans);
  if (!hits.length) return '';
  return `
    <div class="panel">
      <strong>Common progressions spotted (as Roman numerals in ${bestKey}):</strong>
      <p style="margin:6px 0 0">${hits.join(', ')}</p>
    </div>`;
}

// Robust wiring (logs help confirm it's running)
(function wireAnalyzer() {
  function onAnalyze() {
    console.log('[analyzer] click');
    const raw = (document.getElementById('songChords').value || '')
      .split(/\s+/)
      .map(normalizeChord)
      .filter(Boolean);

    if (!raw.length) {
      document.getElementById('keyResults').innerHTML   = '<em>No chords provided.</em>';
      document.getElementById('mappingTable').innerHTML = '';
      document.getElementById('progHints').innerHTML    = '';
      return;
    }

    const results = MAJOR_KEYS
      .map(k => scoreKeyAgainstChords(k, raw))
      .sort((a,b) => b.pct - a.pct)
      .slice(0, 5);

    const best = results[0].key;
    const romans = romanizeSequence(raw, best);

    document.getElementById('keyResults').innerHTML   = renderTopMatches(results);
    document.getElementById('mappingTable').innerHTML = renderMappingTable(best);
    document.getElementById('progHints').innerHTML    = renderProgHints(best, romans);
  }
  function attach() {
    const btn = document.getElementById('analyzeBtn');
    if (!btn) return console.warn('[analyzer] button not in DOM yet');
    if (btn.dataset.wired) return;
    btn.dataset.wired = '1';
    btn.addEventListener('click', onAnalyze);
    console.log('[analyzer] wired');
  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', attach);
  } else {
    attach();
  }
})();





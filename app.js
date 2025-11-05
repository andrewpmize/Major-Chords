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




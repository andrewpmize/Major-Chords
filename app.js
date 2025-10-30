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

  const FlatThree = CAget(14, col);
  const FlatFour  = CAget(15, col);
  const FlatSix   = CAget(17, col);
  const FlatSeven = CAget(18, col);

  // build array of header root tokens once
  const headerRoot = headers.map(h => rootToken(h));

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
    FiveSix: twoStep(10)
  };
}

// ---------- UI wiring ----------
const demoCA = `C,D,Eb,E,F,F#,G,Ab,A,Bb,B
Cmaj,Dmin,Eb,Fmaj,Dmin,Gmaj,Emin,Ab,Am,Bb,B
Dmin,Emin,F,Gmin,Amin,B,C,D,Eb,F,G
E

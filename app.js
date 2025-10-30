// --- Utility: capture the leading note token: A–G with optional # or b (case-insensitive)
function rootToken(s) {
  if (!s) return "";
  const m = String(s).trim().match(/^[A-Ga-g](?:#|b)?/);
  return m ? m[0].toUpperCase() : "";
}

// Parse a TSV/CSV-ish block where first line is headers.
// We expect a header row like the CA sheet's row 4 (D:O), including your chosen key (e.g., "C", "G", "Bb"...)
function parseTable(text) {
  const lines = text.trim().split(/\r?\n/).filter(Boolean);
  if (!lines.length) return { headers: [], rows: [] };

  const split = (line) => line.split(/\t|,|;/).map(x => x.trim());
  const headers = split(lines[0]); // columns
  const rows = lines.slice(1).map(line => split(line));
  return { headers, rows };
}

// Find column index for a given key (exact match on header cell)
function findKeyCol(headers, key) {
  return headers.findIndex(h => String(h).trim().toUpperCase() === String(key).trim().toUpperCase());
}

// Given CA (headers + rows), emulate your VBA lookups and return an object with all outputs.
function chordFileMajor_JS(key, CA) {
  const { headers, rows } = CA;
  const col = findKeyCol(headers, key);
  if (col < 0) throw new Error(`Key '${key}' not found in CA headers.`);

  // Helper to read "CA.Cells(rowIdx, col)" where your VBA's CA row numbers start at 1.
  // Our "rows[0]" is the first *data* row under headers, so:
  // CA row 5 → rows[5 - 5] = rows[0], CA row 6 → rows[1], …
  const CAget = (row1Based, col0Based) => {
    const r = row1Based - 5; // rows start at VBA row 5 in your code-path
    const c = col0Based;     // 0-based column within our headers array
    if (r < 0 || r >= rows.length) return "";
    if (c < 0 || c >= headers.length) return "";
    return rows[r][c] ?? "";
  };

  // 1) Primary chords (VBA: rows 5..10 → One..Six)
  const One   = CAget(5,  col);
  const Two   = CAget(6,  col);
  const Three = CAget(7,  col);
  const Four  = CAget(8,  col);
  const Five  = CAget(9,  col);
  const Six   = CAget(10, col);

  // 2) Flats (VBA rows 14,15,17,18)
  const FlatThree = CAget(14, col);
  const FlatFour  = CAget(15, col);
  const FlatSix   = CAget(17, col);
  const FlatSeven = CAget(18, col);

  // 3) Secondary dominants (“FiveOne”..“FiveSix”) from source rows 5..10
  // VBA loop was 5..11, but only mapped cases 5..10; we’ll mirror the 6 outputs.
  const headerRoot = headers.map(h => rootToken(h));

  function twoStep(row1Based) {
    const src = CAget(row1Based, col);
    const r1 = rootToken(src);
    const matchCol = headerRoot.findIndex(hr => hr === r1);
    if (matchCol < 0) return "";

    // From that matched column, read row 9 (V chord), take its root, append "7"
    const vChord = CAget(9, matchCol);
    const vRoot  = rootToken(vChord);
    return vRoot ? (vRoot + "7") : "";
  }

  const FiveOne   = twoStep(5);
  const FiveTwo   = twoStep(6);
  const FiveThree = twoStep(7);
  const FiveFour  = twoStep(8);
  const FiveFive  = twoStep(9);
  const FiveSix   = twoStep(10);

  return {
    One, Two, Three, Four, Five, Six,
    FlatThree, FlatFour, FlatSix, FlatSeven,
    FiveOne, FiveTwo, FiveThree, FiveFour, FiveFive, FiveSix
  };
}

// ----- UI wiring -----
const demoCA = `C,D,Eb,E,F,F#,G,Ab,A,Bb,B
Cmaj,Dmin,Eb?,Fmaj,Fmaj7,F#?,Gmaj,Ab?,A?,Bb?,B?
Cmin,D?,Eb?,F?,F?,F#?,G?,Ab?,A?,Bb?,B?
E?,F?,G?,A?,B?,C?,D?,Eb?,F?,G?,A?
F?,G?,A?,Bb?,C?,D?,E?,F?,G?,A?,B?
G?,A?,B?,C?,D?,E?,F#?,G?,A?,Bb?,B?
A?,B?,C?,D?,E?,F#?,G#?,A?,B?,C?,D?
(…continue your CA rows exactly as in your sheet…)`;

window.addEventListener('DOMContentLoaded', () => {
  const caInput = document.getElementById('caInput');
  if (!caInput.value.trim()) caInput.value = demoCA;

  document.getElementById('runBtn').addEventListener('click', () => {
    const key = document.getElementById('key').value.trim();
    try {
      const CA = parseTable(caInput.value);
      const out = chordFileMajor_JS(key, CA);
      renderResults(out);
    } catch (err) {
      renderError(err.message || String(err));
    }
  });
});

function renderResults(o) {
  const el = document.getElementById('results');
  el.style.display = 'block';
  const secdom = [ ['V/ I', o.FiveOne], ['V/ ii', o.FiveTwo], ['V/ iii', o.FiveThree],
                   ['V/ IV', o.FiveFour], ['V/ V', o.FiveFive], ['V/ vi', o.FiveSix] ];
  el.innerHTML = `
    <h2>Results</h2>

    <div class="card">
      <div class="grid">
        <div><strong>I</strong> <span class="badge">${safe(o.One)}</span></div>
        <div><strong>ii</strong> <span class="badge">${safe(o.Two)}</span></div>
        <div><strong>iii</strong> <span class="badge">${safe(o.Three)}</span></div>
        <div><strong>IV</strong> <span class="badge">${safe(o.Four)}</span></div>
        <div><strong>V</strong> <span class="badge">${safe(o.Five)}</span></div>
        <div><strong>vi</strong> <span class="badge">${safe(o.Six)}</span></div>
      </div>
    </div>

    <div class="card">
      <h3>Secondary Dominants</h3>
      <div class="grid">
        ${secdom.map(([lab,val]) => `<div><strong>${lab}</strong> <span class="badge">${safe(val)}</span></div>`).join('')}
      </div>
    </div>

    <div class="card">
      <h3>Modal Interchange</h3>
      <div class="grid">
        <div><strong>♭III</strong> <span class="badge">${safe(o.FlatThree)}</span></div>
        <div><strong>♭IV</strong>  <span class="badge">${safe(o.FlatFour)}</span></div>
        <div><strong>♭VI</strong>  <span class="badge">${safe(o.FlatSix)}</span></div>
        <div><strong>♭VII</strong> <span class="badge">${safe(o.FlatSeven)}</span></div>
      </div>
    </div>
  `;
}
function renderError(msg) {
  const el = document.getElementById('results');
  el.style.display = 'block';
  el.innerHTML = `<h2>Error</h2><p>${safe(msg)}</p>`;
}
function safe(s) { return String(s ?? "").replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m])); }

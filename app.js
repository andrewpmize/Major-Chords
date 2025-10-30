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
    const r = row1Based - 5; // our first data row corresponds to VBA row 5
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

  const headerRoot = CA.headers ? CA.headers.map(h => rootToken(h)) : [];

  function twoStep(row1Based) {
    const src = CAget(row1Based, col);
    const r1  = rootToken(src);
    const matchCol = headerRoot.findIndex(hr => hr === r1);
    if (matchCol < 0) return "";
    const vChord = CAget(9, matchCol);     // row 9 holds the V chord in that column
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
Emin,F#,G,Amin,B,C,D,Eb,F,G,A
F,G,A,Bb,C,D,E,F,G,A,B
G,A,B,C,D,E,F#,G,A,Bb,B
A,B,C#,D,E,F#,G#,A,B,C#,D
Bb,C,D,Eb,F,G,A,Bb,C,D,E
(…paste your real grid here; first data row should correspond to VBA row 5…)`;

window.addEventListener('DOMContentLoaded', () => {
  const caInput = document.getElementById('caInput');
  if (!caInput.value.trim()) caInput.value = demoCA;

  document.getElementById('runBtn').addEventListener('click', () => {
    const key = document.getElementById('key').value.trim();
    try {
      const CA = parseTable(caInput.value);
      const out = chordFileMajor_JS(key, CA);
      paintBands(out);
      drawAllArrows();
    } catch (err) {
      alert(err.message || String(err));
    }
  });

  // run once for demo
  document.getElementById('runBtn').click();
});

// Fill the slots inside each band
function paintBands(o) {
  document.querySelectorAll('[data-slot]').forEach(el => {
    const k = el.getAttribute('data-slot');
    el.textContent = safe(o[k] ?? '');
  });
}

// ---------- Arrow drawing ----------
function centerBottom(el, rel) {
  const a = el.getBoundingClientRect();
  const b = rel.getBoundingClientRect();
  return { x: a.left + a.width/2 - b.left, y: a.bottom - b.top };
}
function centerTop(el, rel) {
  const a = el.getBoundingClientRect();
  const b = rel.getBoundingClientRect();
  return { x: a.left + a.width/2 - b.left, y: a.top - b.top };
}
function arrow(svg, p1, p2, opts={}) {
  const { doubleHead=false, dashed=false, bend=0 } = opts;
  const ns = 'http://www.w3.org/2000/svg';

  // marker
  if (!svg.querySelector('#arrowHead')) {
    const defs = document.createElementNS(ns,'defs');
    const m = document.createElementNS(ns,'marker');
    m.setAttribute('id','arrowHead');
    m.setAttribute('markerWidth','10'); m.setAttribute('markerHeight','8');
    m.setAttribute('refX','9'); m.setAttribute('refY','4');
    m.setAttribute('orient','auto');
    const path = document.createElementNS(ns,'path');
    path.setAttribute('d','M0,0 L10,4 L0,8 Z'); path.setAttribute('fill','#1f5fd1');
    m.appendChild(path); defs.appendChild(m); svg.appendChild(defs);
  }

  const path = document.createElementNS(ns,'path');
  const midY = (p1.y + p2.y)/2 + bend;
  const d = `M ${p1.x},${p1.y} C ${p1.x},${midY} ${p2.x},${midY} ${p2.x},${p2.y}`;
  path.setAttribute('d', d);
  path.setAttribute('fill','none');
  path.setAttribute('stroke','#1f5fd1');
  path.setAttribute('stroke-width','2.5');
  if (dashed) path.setAttribute('stroke-dasharray','6 6');
  path.setAttribute('marker-end','url(#arrowHead)');
  if (doubleHead) path.setAttribute('marker-start','url(#arrowHead)');
  svg.appendChild(path);
}

function drawAllArrows() {
  const container = document.getElementById('diagram');
  const svg = document.getElementById('wires');
  // size SVG to container
  const r = container.getBoundingClientRect();
  svg.setAttribute('width', r.width);
  svg.setAttribute('height', container.scrollHeight);

  // clear previous
  while (svg.firstChild) svg.removeChild(svg.firstChild);

  // ensure marker defs persist
  // (arrow() will recreate as needed)

  // 1) SecDom → Main (down arrows)
  const pairs = [
    ['sec-I','main-I'],
    ['sec-ii','main-ii'],
    ['sec-iii','main-iii'],
    ['sec-IV','main-IV'],
    ['sec-V','main-V'],
    ['sec-vi','main-vi'],
  ];
  pairs.forEach(([a,b])=>{
    arrow(svg, centerBottom(document.getElementById(a), container),
               centerTop(document.getElementById(b), container),
               { bend: 10 });
  });

  // 2) Main (I, IV, V) ↔ Modal (♭III, ♭VI, iv, ♭VII) double-ended
  const mains = ['main-I','main-IV','main-V'];
  const mods  = ['mod-bIII','mod-bVI','mod-iv','mod-bVII'];
  mains.forEach(m=>{
    mods.forEach(md=>{
      arrow(svg,
        centerBottom(document.getElementById(m), container),
        centerTop(document.getElementById(md), container),
        { doubleHead:true, bend: 20 });
    });
  });
}

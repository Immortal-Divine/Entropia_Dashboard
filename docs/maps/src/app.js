/* Entropia Node Maps â€” a standalone viewer for the harvesting maps.
 *
 * Everything here is drawn on one canvas: the world picture, the spawn
 * spots, the respawn points, the place names.  Coordinates are held in
 * the picture's own pixel space, because that is the only frame all three
 * worlds share without inventing an origin â€” a spot is stored as the
 * (ix, iy) the game's transform puts it at, and the view is a pan and a
 * zoom over that picture.  World metres are recovered for the readout
 * through the same transform the overlay uses.
 */
(function () {
'use strict';

var D = window.MAPDATA;
if (!D) { document.body.textContent = 'data.js is missing.'; return; }

var MAPS = D.maps, K = D.k, SPECIES = D.species;
var MAP = {}, SIDX = {};
MAPS.forEach(function (m, i) { MAP[m.key] = m; SIDX[m.key] = i; });

/* ---------------------------------------------------------------- state */
var mkey   = MAPS[0].key;
var img    = new Image();
var ready  = false;
var vx = 0, vy = 0, sc = 1;      /* image px at the canvas centre, px per image px */
var W = 1, H = 1, dpr = 1;
var on     = Object.create(null);   /* species name -> true while it is drawn */
var showSp = true, showPl = true;
var sel    = null;               /* {si, x, y} the detail card is about */
var hover  = -1;                 /* species index under the pointer */
var mx = 0, my = 0, overSpot = null;
var query  = '';
var MATCH = null;               /* what the search found, name -> true,
                                 * kept for the map until it changes */
var viewChosen = false;          /* has any view been fitted at least once */
var loadTok = 0;                 /* which picture load is the live one */
var GRID = Object.create(null);  /* map key -> loaded background Image */
var LAYERM = 200;                 /* scene margin around the viewport, css px */
var layer = null, layerCtx = null, layerW = 0, layerH = 0;
var layerVx = 0, layerVy = 0, layerSc = 0, layerKey = '';
var snap = null, snapCtx = null, snapSc = 0;   /* last full frame, zoom preview */
var gesture = false, settleT = 0;              /* wheel/pinch in progress */

var cv  = document.getElementById('cv');
var ctx = cv.getContext('2d');

SPECIES.forEach(function (s) { on[s.name] = true; });

/* Two further layers, both named.  A region file offers a line for a
 * creature standing in a pack and a line for a single object left on the
 * ground, and the tables those numbers were read from say what each one
 * is - so a filter can name them instead of counting them.  The switches
 * are keyed by name rather than by number for two reasons: the same
 * creature on two worlds is one switch rather than two, and a shared link
 * reads Snow Demon rather than 13498. */
var CKEY = '@c:', IKEY = '@i:', TKEY = '@t:', SKEY = '@s:', NKEY = '@n:';
var CRE = D.creatureNames || {}, ITM = D.itemNames || {};
/* The second language for every name the client carries in both: a
 * creature resolved through the Lang tables reads English on one client
 * and German on the other, so the row shows both and a search over the row
 * finds it by either - it does not matter which language the site is
 * read in.  Null where the tables carry only one. */
var CRE_ALT = D.creatureNamesAlt || {}, ITM_ALT = D.itemNamesAlt || {};
var COL = { mob: '#78dcff', boss: '#ff6ea8', item: '#ffb64a' };
/* The treasure maps are their own layer: a dig spot on Cardia (or in its
 * cave) with the map that points at it.  Each row is one map, and its
 * picture is the parchment the game hands out - so the row, the pin and the
 * tooltip all show the same torn map, and the riddle baked into it is read
 * straight off the picture. */
var TTM = D.treasureNames || {}, TIC = D.treasureIcons || {},
    TTF = D.treasureFull || {}, TTX = D.treasureText || {},
    TCAT = D.treasureCat || {}, TZN = D.treasureZone || {};
/* The bonus-stat statues are their own layer: a placed monument standing
 * where the character can read the stat off it.  statueNames carries each
 * statue's unique key - its name plus the area it stands in - so two
 * "Madrigal Statue"s are two switches; statueLabel is the bare name the row
 * shows, the area heading already printing where it is.  The game files
 * hold no picture of a statue, so the one it is drawn with is cut out of a
 * screenshot of the monument itself: sidebar, pin and card all show the
 * same cutout. */
var STM = D.statueNames || {}, SZL = D.statueLabel || {},
    SZN = D.statueZone || {};
function statueIcon(k) {
  /* The cutout is shared by every world; data.js paths are relative to the
   * atlas asset root, so this resolves to icons/shared/statues/statue.png. */
  return 'shared/statues/statue.png';
}
/* The NPCs are their own layer: guards, quest givers and townsfolk the
 * world files place by hand, one row per kind the way a monster is - the
 * fifteen town guards are one switch, not fifteen.  The game files hold
 * no picture of a person standing around, so an NPC is marked by colour
 * and shape the way an ore is: a diamond in the one green nothing else
 * on the map wears. */
var NTM = D.npcNames || {}, NZN = D.npcZone || {};
/* The rim around a picture is the colour of the thing it marks rather than
 * one black for all of them: a boss is edged in red, a plant in a dark
 * charcoal purple, an ore in a workaday grey, an item in yellow, and a
 * creature - which is what a map is mostly made of - the darkest of them, a
 * black with just enough in it to be seen on the dark ground as well as on
 * the bright, so the four kinds can be told apart at a glance without four
 * frames.  The sidebar filters the same colours over the same pictures,
 * so a row's edge and the mark it stands for agree. */
var RIMCOL = { mob: '#333333', boss: '#ff5a4d', item: '#ffdc4d',
               ore: '#8a9099', plant: '#5a4b70', treasure: '#d4953a',
               statue: '#4d9de0', npc: '#3fd16a' };

/* Boss is a class the game's own table gives a creature, so it is read as
 * one: the ordinary animals and the things a player would call a boss are
 * not the same row of that table, and a map which draws them the same way
 * has thrown away the only distinction the game makes. */
var BOS = {};
Object.keys(D.creatureBoss || {}).forEach(function (k) {
  var o = Object.create(null);
  (D.creatureBoss[k] || []).forEach(function (i) { o[i] = 1; });
  BOS[k] = o;
});
function creNames(k) { return CRE[k] || []; }
function itmNames(k) { return ITM[k] || []; }
function treNames(k) { return TTM[k] || []; }
function statueNames(k) { return STM[k] || []; }
function statueLabel(k, ref) { var a = SZL[k]; return (a && a[ref]) || STM[k][ref] || ''; }
function npcNames(k) { return NTM[k] || []; }
function isBoss(k, ref) { return !!(BOS[k] && BOS[k][ref]); }

/* A creature's picture is filed under its model rather than under the name
 * a player reads, and an item's under the file the game itself names, so
 * the build hands a path over beside each name and this viewer only has to
 * open it.  One image is created per path no matter how many points it
 * stands at: a creature may stand in twenty places and asking for it twenty
 * times would be twenty requests for one picture. */
/* Two pictures exist for the same creature and both are shipped: the
 * full-body art the in-game map draws, and the square its model was filed
 * under in the item folder.  Which one is showing is the reader's to
 * choose - the sprite reads further across a bright map, the square sits
 * tighter in a crowd - so nothing here assumes it knows which. */
var CIC = D.creatureIcons || {}, CIC2 = D.creatureSquare || {},
    IIC = D.itemIcons || {};
/* The same again for an ore or a plant: the game names a picture of what
 * lands in your bag for mining it, so the sidebar can show the thing
 * itself.  It is one array against the species list rather than a map per
 * world, because a species is the same on every world it stands on. */
var SPIC = D.speciesIcons || [];
/* One picture per creature (see creIcon): the Full-body/Item-icons
 * switch is gone, so there is no picture preference to keep. */
/* Which groups are collapsed survives rebuilding the list: the picture
 * switch has to rebuild it, and losing a collapse the reader made would
 * be a small theft every time they pressed the button. */
var GOPEN = {};
var CZN = D.creatureZone || {}, IZN = D.itemZone || {};
var ICAT = D.itemCat || {};
/* A monster's tier level (Platin <P 16>, Royal <RL. 31>) rides beside its
 * name in rows and tips, never inside the name the switches, links and
 * search match. Normal mobs carry null until the live level bank lands. */
var CLVL = D.creatureLevel || {};
function lvlText(code) {
  if (!code) return '';
  if (code.charAt(0) === 'P') return '&lt;P ' + esc(code.slice(1)) + '&gt;';
  if (code.slice(0, 2) === 'RL') return '&lt;RL. ' + esc(code.slice(2)) + '&gt;';
  return '';
}
/* A set heading is written out as the raw key the build files rows under,
 * which is fine for "quest" and means nothing to a reader for "buff" -
 * the group over the world buffs says what those rows are. */
var CATTITLE = { buff: 'World Buff Spawn' };
var ZONE = D.zones || {};
var ICON = Object.create(null);        /* path -> Image, filled on demand */
var icoGen = 0; /* bumped whenever a picture finishes: the layer holds pixels,
                 * so a newly arrived picture has to retire it */
var icoDirty = false;

/* data.js keeps the paths relative to the atlas asset root: "maps/..."
 * resolves below icons/ for the desktop atlas and the website alike.  The
 * website serves index.html from the folder above icons/, so every picture
 * URL has to pass through here; data.js stays module-compatible. */
function asset(p) {
  if (!p || /^(?:[a-z]+:)?\/\//i.test(p) || p.charAt(0) === '/' ||
      p.indexOf('icons/') === 0) return p;
  return 'icons/' + p;
}

/* Many pictures finish loading within the same tick, and redrawing once
 * per picture would repaint the whole map a hundred times over while it
 * settles.  One frame is asked for and every load that lands before it
 * shares it. */
function icoRedraw() {
  if (icoDirty) return;
  icoDirty = true;
  requestAnimationFrame(function () { icoDirty = false; draw(); });
}
function iconImg(p) {
  if (!p) return null;
  p = asset(p);
  var im = ICON[p];
  if (!im) {
    im = new Image();
    im.onload = function () { icoGen++; icoRedraw(); };
    im.onerror = function () {
      /* A saturated server can drop one request of a burst; the errored
       * image would otherwise sit in the cache showing its fallback
       * glyph until a full reload.  One retry, then it stays a glyph. */
      if (!im._retried) {
        im._retried = true;
        setTimeout(function () { im.src = p; }, 1200);
      }
    };
    im.src = p;
    ICON[p] = im;
  }
  return im;
}
function creIcon(k, ref) {
  /* One picture per creature, no switch: the full-body art the in-game map
   * draws where the game has it, the model square everywhere else. */
  var a = CIC[k];
  var v = a && a[ref];
  if (!v) v = CIC2[k] && CIC2[k][ref];
  return v || null;
}
function creIsSquare(k, ref) {
  return !(CIC[k] && CIC[k][ref]) && !!(CIC2[k] && CIC2[k][ref]);
}
function itmIcon(k, ref) { var a = IIC[k]; return a ? (a[ref] || null) : null; }

/* A world buff post is not bound to the buff that was seen standing at
 * it: the server rolls family and level per spawn, so one post can hold
 * Penya Rush one moment and Swiftwind the next.  The pin therefore turns
 * through the pictures of the buffs that roll on this island rather than
 * wearing whichever happened to be observed, and each pin takes its own
 * place in the turn, so two posts side by side can show different rolls
 * at the same instant. */
var BUFF_ROT = [], BUFF_ROTKEY = null, buffRotT = 0;
function buffRot(m) {
  if (BUFF_ROTKEY === m) return;
  BUFF_ROTKEY = m;
  BUFF_ROT = [];
  var cat = ICAT[m] || [];
  (itmNames(m) || []).forEach(function (n, ref) {
    if (cat[ref] !== 'buff') return;
    var ic = itmIcon(m, ref);
    if (ic && BUFF_ROT.indexOf(ic) < 0) BUFF_ROT.push(ic);
  });
}
function buffIconAt(x, y) {
  if (!BUFF_ROT.length) return null;
  if (BUFF_ROT.length < 2) return BUFF_ROT[0];
  var h = (Math.round(x) * 92837111 ^ Math.round(y) * 689287499) >>> 0;
  return BUFF_ROT[(h % BUFF_ROT.length + buffRotT) % BUFF_ROT.length];
}
/* The map draws when it is changed rather than running an animation loop,
 * so the turn asks for a repaint itself - and only while there is a pool
 * to turn through. */
setInterval(function () {
  if (BUFF_ROT.length < 2 || !showSp) return;
  buffRotT = (buffRotT + 1) % BUFF_ROT.length;
  draw();
}, 2400);
/* Two pictures per treasure: the cropped illustration the row and the pin
 * show (small, no parchment, no text), and the whole scroll the tooltip and
 * the detail card show, where the baked riddle is actually read. */
function treIcon(k, ref) { var a = TIC[k]; return a ? (a[ref] || null) : null; }
function treFull(k, ref) { var a = TTF[k]; return a ? (a[ref] || null) : null; }
function treText(k, ref) { var a = TTX[k]; return (a && a[ref]) || ''; }
function zonesOf(k) { return ZONE[k] || []; }

/* Open every picture this world will ask for as soon as the world is
 * chosen, so the first draw after the map arrives is not a hundred empty
 * squares that fill in one at a time.  Both creature sets are opened: the
 * switch should change the map, not start it waiting. */
function preloadIcons() {
  (creNames(mkey) || []).forEach(function (n, ref) {
    var p = creIcon(mkey, ref);
    if (p) iconImg(p);
  });
  [IIC[mkey] || [], TIC[mkey] || []].forEach(function (list) {
    for (var i = 0; i < list.length; i++) if (list[i]) iconImg(list[i]);
  });
  /* an ore or a plant is one picture each, held against the species list */
  speciesOfThisMap().forEach(function (s) {
    var p = SPIC[SPECIES.indexOf(s)];
    if (p) iconImg(p);
  });
}

/* Default every switch on, including those belonging to a world the visitor
 * has not opened yet: a link carries the state of one world only, and
 * anything the link does not mention has to mean "still as it was". */
SPECIES.forEach(function (s) { on[s.name] = true; });
Object.keys(CRE).forEach(function (k) {
  CRE[k].forEach(function (n) { on[CKEY + n] = true; });
});
Object.keys(ITM).forEach(function (k) {
  ITM[k].forEach(function (n) { on[IKEY + n] = true; });
});
Object.keys(TTM).forEach(function (k) {
  TTM[k].forEach(function (n) { on[TKEY + n] = true; });
});
Object.keys(STM).forEach(function (k) {
  STM[k].forEach(function (n) { on[SKEY + n] = true; });
});
Object.keys(NTM).forEach(function (k) {
  NTM[k].forEach(function (n) { on[NKEY + n] = true; });
});

/* --------------------------------------------------------- persistence */
var LS = 'entropia-maps.v1';
function save() {
  var off = [];
  for (var n in on) if (!on[n]) off.push(n);
  try {
    localStorage.setItem(LS, JSON.stringify({
      map: mkey, off: off, sp: showSp, pl: showPl, grp: mgrp
    }));
  } catch (e) { /* private mode: the URL still carries the view */ }
}
function load() {
  var s = null;
  try { s = JSON.parse(localStorage.getItem(LS) || 'null'); } catch (e) {}
  if (!s) return;
  if (MAP[s.map]) mkey = s.map;
  if (s.sp !== undefined) showSp = !!s.sp;
  if (s.pl !== undefined) showPl = !!s.pl;
  if (s.grp !== undefined && GROUPS.indexOf(s.grp) >= 0)
    mgrp = s.grp;
  if (s.off && s.off.length) {
    for (var i = 0; i < s.off.length; i++) on[s.off[i]] = false;
  }
}

/* The URL is the shareable half: a link opened in a forum thread has to
 * come back to the same world, the same window on it, the same filters
 * and the same search.  Filters are a compact bit list, not names: every
 * filter key has a stable index in the sorted key list, the off (or on)
 * indices are delta + varint packed and base64url written, so even a
 * hundred switched-off names stay tens of characters.  The common
 * "everything on, whole map" link stays just #m=<world>. */
function b64e(bytes) {
  var s = '';
  for (var i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i] & 255);
  return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}
function b64d(code) {
  code = String(code).replace(/-/g, '+').replace(/_/g, '/');
  while (code.length % 4) code += '=';
  var s = atob(code), out = [];
  for (var i = 0; i < s.length; i++) out.push(s.charCodeAt(i) & 255);
  return out;
}
/* Sorted filter indices to a short string: first index as-is, then each
 * gap minus one (gaps are >= 1, so minus one keeps them small), each as a
 * 7-bit varint. */
function encIdx(list) {
  if (!list.length) return '';
  var bytes = [], prev = 0, first = true, i, v, n, b;
  for (i = 0; i < list.length; i++) {
    v = first ? list[i] : (list[i] - prev - 1);
    first = false; prev = list[i];
    n = v >>> 0;
    do { b = n & 127; n >>>= 7; if (n) b |= 128; bytes.push(b); } while (n);
  }
  return b64e(bytes);
}
function decIdx(code) {
  if (!code) return [];
  var bytes = b64d(code), out = [], i = 0, prev = -1, v, sh, b;
  while (i < bytes.length) {
    v = 0; sh = 0;
    for (;;) {
      if (i >= bytes.length) return out;
      b = bytes[i++];
      v |= (b & 127) << sh; sh += 7;
      if (!(b & 128)) break;
    }
    var idx = out.length === 0 ? v : prev + 1 + v;
    out.push(idx); prev = idx;
  }
  return out;
}
function filterKeys() { return Object.keys(on).sort(); }
/* Whichever list is shorter carries the state: '0' + off indices, or
 * '1' + on indices.  Empty (everything on) is '' and omitted. */
function encodeFilter() {
  var keys = filterKeys(), pos = Object.create(null), i;
  for (i = 0; i < keys.length; i++) pos[keys[i]] = i;
  var off = [], onl = [], n, id;
  for (n in on) {
    id = pos[n];
    if (id === undefined) continue;
    if (on[n]) onl.push(id); else off.push(id);
  }
  off.sort(function (a, b) { return a - b; });
  onl.sort(function (a, b) { return a - b; });
  if (!off.length) return '';
  if (off.length <= onl.length) return '0' + encIdx(off);
  return '1' + encIdx(onl);
}
function decodeFilter(code) {
  if (!code) return;
  var mode = code.charAt(0), data = code.slice(1);
  var keys = filterKeys(), ids = decIdx(data), j;
  if (mode === '1') {
    var keep = Object.create(null);
    for (j = 0; j < ids.length; j++)
      if (keys[ids[j]] !== undefined) keep[keys[ids[j]]] = 1;
    for (var n in on) on[n] = !!keep[n];
  } else {
    for (var m in on) on[m] = true;
    for (j = 0; j < ids.length; j++)
      if (keys[ids[j]] !== undefined) on[keys[ids[j]]] = false;
  }
}
function readHash() {
  var h = location.hash.replace(/^#/, '');
  if (!h) return false;
  var q = {}, any = false;
  h.split('&').forEach(function (p) {
    var i = p.indexOf('=');
    if (i < 0) return;
    q[decodeURIComponent(p.slice(0, i))] = decodeURIComponent(p.slice(i + 1));
    any = true;
  });
  if (!any) return false;
  if (MAP[q.m]) mkey = q.m;
  var nx = q.x !== undefined ? q.x : q.vx;
  var ny = q.y !== undefined ? q.y : q.vy;
  var nz = q.z !== undefined ? q.z : q.sc;
  if (nx !== undefined) vx = +nx;
  if (ny !== undefined) vy = +ny;
  if (nz !== undefined && isFinite(+nz)) sc = +nz;
  if (q.f !== undefined) {
    decodeFilter(q.f);
  } else {
    if (q.on) {                 /* legacy long form, still read */
      var keep = Object.create(null);
      q.on.split('|').forEach(function (n) { if (n) keep[n] = true; });
      for (var n2 in on) on[n2] = !!keep[n2];
    }
    if (q.off) {
      q.off.split('|').forEach(function (n) { if (n) on[n] = false; });
    }
  }
  if (q.sp !== undefined) showSp = q.sp !== '0';
  if (q.pl !== undefined) showPl = q.pl !== '0';
  if (q.q !== undefined) {
    query = String(q.q).slice(0, 120);
    var elQ = document.getElementById('q');
    if (elQ && elQ.value !== query) elQ.value = query;
  } else if (query) {
    query = '';
    var elQ2 = document.getElementById('q');
    if (elQ2 && elQ2.value !== '') elQ2.value = '';
  }
  if (nx !== undefined && ny !== undefined) {
    viewChosen = true;      /* a link carries a window onto the world, so
                             * the very first fit must not replace it */
    return true;
  }
  return false;
}
var hashLock = false;
function writeHash() {
  if (hashLock) return;
  var p = ['m=' + encodeURIComponent(mkey)];
  var m = MAP[mkey];
  if (m && Math.abs(sc - fitScale()) > 1e-6) {
    p.push('x=' + vx.toFixed(1), 'y=' + vy.toFixed(1), 'z=' + sc.toFixed(3));
  }
  var f = encodeFilter();
  if (f) p.push('f=' + encodeURIComponent(f));
  if (query) p.push('q=' + encodeURIComponent(query));
  if (!showSp) p.push('sp=0');
  if (!showPl) p.push('pl=0');
  var u = '#' + p.join('&');
  if (u !== location.hash) {
    history.replaceState(null, '', u);
  }
}
/* ------------------------------------------------------------ geometry */
/* The picture is drawn at its own pixel size and is never resampled out of
 * shape.  A world's data box (m.w x m.h) is the frame the pins were placed
 * in, and the picture that ships with the world today can differ from it in
 * both size and shape - a picture replaced after its pins were placed is the
 * normal reason.  Both cover the same ground, so the box is mapped onto the
 * picture instead of the picture being pressed into the box: a pin at
 * (ix, iy) sits at picture pixel (ix * PPX, iy * PPY).  The view - the centre
 * (vx, vy) and the zoom (sc) - is therefore measured in picture pixels, and
 * a world whose picture does match its box behaves exactly as it always did
 * (both factors are 1). */
var PW = 1, PH = 1;              /* the picture's own pixels */
var PPX = 1, PPY = 1;            /* picture pixels per map unit */
function picMetrics() {
  var m = MAP[mkey];
  var iw = (img && img.complete && img.naturalWidth) || 0;
  var ih = (img && img.complete && img.naturalHeight) || 0;
  if (!iw || !ih) { iw = m.w; ih = m.h; }   /* the picture is not there yet */
  PW = iw; PH = ih;
  PPX = iw / m.w; PPY = ih / m.h;
}
function sx(ix) { return W / 2 + (ix * PPX - vx) * sc; }
function sy(iy) { return H / 2 + (iy * PPY - vy) * sc; }
function ixAt(px) { return (vx + (px - W / 2) / sc) / PPX; }
function iyAt(py) { return (vy + (py - H / 2) / sc) / PPY; }
function fitScale() { return Math.min(W / PW, H / PH); }

function fit() {
  sc = fitScale();
  vx = PW / 2; vy = PH / 2;
  viewChosen = true;
}
/* Fit the view around one entry's pins - a row click, a zone heading, a
 * search jump to another world - so what was picked is what ends up on
 * screen, at the closest zoom that still holds every pin with room. */
function fitPoints(pts) {
  if (!pts || !pts.length) return;
  var x0 = 1e18, y0 = 1e18, x1 = -1e18, y1 = -1e18, i;
  for (i = 0; i < pts.length; i++) {
    if (pts[i][0] < x0) x0 = pts[i][0];
    if (pts[i][1] < y0) y0 = pts[i][1];
    if (pts[i][0] > x1) x1 = pts[i][0];
    if (pts[i][1] > y1) y1 = pts[i][1];
  }
  /* the pins are in map units; the view is in picture pixels */
  var pad = 70, bw = Math.max(200, (x1 - x0) * PPX) + pad * 2,
      bh = Math.max(200, (y1 - y0) * PPY) + pad * 2;
  sc = Math.min(W / bw, H / bh, 8);
  vx = (x0 + x1) / 2 * PPX; vy = (y0 + y1) / 2 * PPY;
  clampView(); viewChosen = true; draw(); writeHash();
}
/* Picture pins behind one sidebar name, on this world or, for a search
 * jump, on the world being jumped to.  Every layer answers: creatures and
 * ground items by their row name, treasure maps and NPCs the same way,
 * statues by either their unique key or the bare label the row shows, and
 * ore/plant ground by the species name itself.  With an area given, only
 * the pins standing in that area answer, so a jump lands in the zone it
 * was picked from rather than framing every same-named pin on the map. */
/* A world whose picture was replaced after its data was built had its pins
 * moved by hand into the new picture's pixels - a scale and a shift, and a
 * nudge for the pins no scale and shift could put right.  The calibration
 * that did it is kept in the data, so the world a pin stands in is still
 * read in the game's own metres: the movement is undone before the
 * picture's pixels are turned into metres.  What the game calls a place
 * does not move because the picture was redrawn. */
var CAL = D.calib || {};
function calibOf(k) { return CAL[k] || null; }
/* The picture's own pixels to the metre, for the scale bar: a calibrated
 * world was drawn at another rate, so its bar is longer. */
function pxPerMetre() {
  var c = CAL[mkey];
  return K * (c ? c.sx : 1);
}
function pinsForSpecies(label, mapKey, zone) {
  var k = mapKey || mkey, out = [], i;
  var arr = D.spots[k] || [];
  for (i = 0; i < arr.length; i++) {
    var sp = SPECIES[arr[i][2]];
    if (sp && sp.name === label && (zone == null || sp.zone === zone))
      out.push([arr[i][0], arr[i][1]]);
  }
  return out;
}
function pinsFor(pfx, label, mapKey, zone) {
  var k = mapKey || mkey, arr, names, zn, out = [], i;
  if (pfx === CKEY) { arr = D.mobs[k] || []; names = creNames(k); zn = CZN[k] || []; }
  else if (pfx === IKEY) { arr = D.items[k] || []; names = itmNames(k); zn = IZN[k] || []; }
  else if (pfx === TKEY) { arr = D.treasures[k] || []; names = treNames(k); zn = TZN[k] || []; }
  else if (pfx === NKEY) { arr = D.npcs[k] || []; names = npcNames(k); zn = NZN[k] || []; }
  else if (pfx === SKEY) {
    arr = D.statues[k] || [];
    var sz = SZN[k] || [];
    for (i = 0; i < arr.length; i++) {
      var ref = arr[i][2];
      if (((STM[k] && STM[k][ref] === label) || statueLabel(k, ref) === label) &&
          (zone == null || sz[ref] === zone))
        out.push([arr[i][0], arr[i][1]]);
    }
    return out;
  }
  else {
    /* A bare species name has no @ prefix, so a split caller lands here
     * with its first three letters as pfx - stitch it back together. */
    return pinsForSpecies((pfx || '') + (label || ''), k, zone);
  }
  for (i = 0; i < arr.length; i++)
    if (names[arr[i][2]] === label && (zone == null || zn[arr[i][2]] === zone))
      out.push([arr[i][0], arr[i][1]]);
  return out;
}
/* One sidebar row's pins, whatever layer it belongs to: a prefixed filter
 * key is split into its layer, a bare species name is looked up as-is. */
function pinsForRow(fullName, mapKey, zone) {
  if (!fullName) return [];
  if (fullName.charAt(0) === '@') return pinsFor(fullName.slice(0, 3), fullName.slice(3), mapKey, zone);
  return pinsForSpecies(fullName, mapKey, zone);
}
/* The filter switches behind one Other-worlds search hit.  Most layers are
 * one switch per name, so the key follows the label; statues are one switch
 * per placed monument, so every unique key wearing that label answers. */
function xrowKeys(r, mapKey) {
  if (r.grp === 'Items') return [IKEY + r.label];
  if (r.grp === 'Monsters' || r.grp === 'Bosses') return [CKEY + r.label];
  if (r.grp === 'Treasure Maps') return [TKEY + r.label];
  if (r.grp === 'NPCs') return [NKEY + r.label];
  if (r.grp === 'Statues') {
    var out = [], arr = STM[mapKey] || [], i;
    for (i = 0; i < arr.length; i++)
      if (statueLabel(mapKey, i) === r.label) out.push(SKEY + arr[i]);
    return out.length ? out : [SKEY + r.label];
  }
  return [r.label];
}
/* The pins behind one Other-worlds hit, scoped to the hit's own area: a
 * name standing in several zones frames the zone it was picked from rather
 * than every same-named pin on the map.  Without an area the whole map's
 * pins answer, as before. */
function xrowPoints(r, mapKey, zone) {
  if (r.grp === 'Items') return pinsFor(IKEY, r.label, mapKey, zone);
  if (r.grp === 'Monsters' || r.grp === 'Bosses') return pinsFor(CKEY, r.label, mapKey, zone);
  if (r.grp === 'Treasure Maps') return pinsFor(TKEY, r.label, mapKey, zone);
  if (r.grp === 'Statues') return pinsFor(SKEY, r.label, mapKey, zone);
  if (r.grp === 'NPCs') return pinsFor(NKEY, r.label, mapKey, zone);
  return pinsForSpecies(r.label, mapKey, zone);
}
function clampView() {
  var pad = 60;
  var lo, hi;
  lo = (pad - W / 2) / sc; hi = PW - (pad - W / 2) / sc;
  vx = lo > hi ? PW / 2 : Math.min(Math.max(vx, lo), hi);
  lo = (pad - H / 2) / sc; hi = PH - (pad - H / 2) / sc;
  vy = lo > hi ? PH / 2 : Math.min(Math.max(vy, lo), hi);
}
function zoomMath(px, py, f) {
  var a = ixAt(px), b = iyAt(py);   /* map units under the cursor */
  sc = Math.max(0.012, Math.min(24, sc * f));
  vx += (a - ixAt(px)) * PPX; vy += (b - iyAt(py)) * PPY;
  clampView();
}
function zoomAt(px, py, f) { zoomMath(px, py, f); draw(); queueHash(); }
/* A zoom gesture's URL update can wait for the gesture: replaceState on
 * every wheel tick was a history write per tick for no reader benefit. */
var hashTimer = 0;
function queueHash() {
  if (hashTimer) clearTimeout(hashTimer);
  hashTimer = setTimeout(function () { hashTimer = 0; writeHash(); }, 250);
}
/* Picture pixels to world metres. */
function toWorld(ix, iy) {
  var c = CAL[mkey], m = MAP[mkey];
  if (c) {
    return [(ix - c.ox) / (c.sx * K),
            (c.h - (iy - c.oy) / c.sy) / K];
  }
  return [ix / K, (m.h - iy) / K];
}

/* -------------------------------------------------------------- drawing */
var touched = false;            /* has the pointer been over the map yet */

function resize() {
  /* If the view was showing the whole map, a resize has to keep showing
   * the whole map; if it was a close crop the user chose, a resize must
   * not throw that away.  So the question is asked before the numbers
   * change, and answered with the new ones. */
  var wasFit = viewChosen && Math.abs(sc - fitScale()) < 1e-6;
  var r = cv.parentNode.getBoundingClientRect();
  dpr = window.devicePixelRatio || 1;
  W = Math.max(1, Math.round(r.width));
  H = Math.max(1, Math.round(r.height));
  cv.width = Math.round(W * dpr); cv.height = Math.round(H * dpr);
  cv.style.width = W + 'px'; cv.style.height = H + 'px';
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  layer = null; snap = null;
  if (!touched) { mx = W / 2; my = H / 2; }
  if (!viewChosen || wasFit) fit();
  clampView(); draw();
}
function spotR() { return Math.max(1.7, Math.min(6, 3.1 * sc)); }

function pathRun(arr, i, j, r, plant) {
  ctx.beginPath();
  for (var k = i; k < j; k++) {
    var X = sx(arr[k][0]), Y = sy(arr[k][1]);
    if (X < -30 || Y < -30 || X > W + 30 || Y > H + 30) continue;
    if (plant) {
      ctx.moveTo(X, Y - r * 1.3);
      ctx.lineTo(X + r * 1.05, Y);
      ctx.lineTo(X, Y + r * 1.3);
      ctx.lineTo(X - r * 1.05, Y);
      ctx.closePath();
    } else {
      ctx.moveTo(X + r, Y);
      ctx.arc(X, Y, r, 0, 6.2831853);
    }
  }
}
function drawSpots() {
  var arr = D.spots[mkey] || [];
  if (!arr.length) return;
  var r = spotR(), i = 0, n = arr.length, kinds = [];
  for (i = 0; i < n; i++) {
    var kk = SPECIES[arr[i][2]].kind;
    if (kinds.indexOf(kk) < 0) kinds.push(kk);
  }
  ctx.globalAlpha = 0.92;
  /* one pass per kind, so that a grey ore edge is merged with grey ore
   * edges and a green plant edge with green ones, and the two never
   * pretend to be one shape */
  kinds.forEach(function (kind) {
    beginRim();
    var a = 0;
    while (a < n) {
      var si = arr[a][2], j = a;
      while (j < n && arr[j][2] === si) j++;
      if (SPECIES[si].kind === kind && shown(SPECIES[si].name)) {
        /* the thing itself where the game has a picture of it, and the dot
         * the map has always drawn where it does not, or has not arrived */
        if (!SPIC[si] || !spotPic(arr, a, j, SPIC[si], RIMCOL[kind])) {
          var halo = r + 1.15, plant = kind === 'plant';
          pathRun(arr, a, j, halo, plant); ctx.fillStyle = 'rgba(0,0,0,.6)'; ctx.fill();
          pathRun(arr, a, j, r,     plant); ctx.fillStyle = SPECIES[si].color; ctx.fill();
        }
      }
      a = j;
    }
    endRim();
  });
  ctx.globalAlpha = 1;
}
/* The rim every picture on the map wears, and why it is not a box: it is
 * the picture's own silhouette, filled in the colour of what it marks and
 * grown by exactly one pixel in all eight directions - and nothing else.
 * What is left is a single opaque pixel wherever the shape ends - a dragon
 * edged along the dragon, a flower along the flower - and nothing at all
 * where the art is not, which is what a rectangular frame could never do.
 * The picture itself is a second copy, drawn a moment after it (see
 * paintRim), which is what lets marks of one colour be edged as the one
 * shape they make together.
 *
 * Three details carry the whole thing.  The silhouette is taken as solid
 * wherever the art is more than a third opaque, so a soft glow around a
 * figure does not become the figure.  It is grown in all eight
 * directions rather than four, so a vertical edge is not left with a
 * broken, stair-stepped outline.  And it is written as its own opaque
 * colour rather than reusing the art's alpha: a picture's outermost pixel
 * is anti-aliased to about half, and a rim drawn at half alpha reads as a
 * smudge rather than a line.
 *
 * Built once per picture per pixel size and kept, because a hundred and
 * thirty-nine lilies are one rim drawn a hundred and thirty-nine times,
 * and dropped once too many sizes have been tried for it to be worth the
 * memory. */
var RIM = Object.create(null), RIMB = Object.create(null), rimN = 0, rimGen = 0;
/* #rgb or #rrggbb to the three numbers the pixel buffer wants. */
function rimRGB(col) {
  var h = String(col).replace('#', '');
  if (h.length === 3) h = h[0] + h[0] + h[1] + h[1] + h[2] + h[2];
  return [parseInt(h.slice(0, 2), 16) || 0,
          parseInt(h.slice(2, 4), 16) || 0,
          parseInt(h.slice(4, 6), 16) || 0];
}
/* A picture keeps its own width-to-height wherever it is drawn: fitted
 * into the box centred, never stretched into a square.  The rim is grown
 * from the fitted art, so the edge follows the true silhouette. */
function fitBox(im, n) {
  var w = im.naturalWidth, h = im.naturalHeight;
  if (!w || !h) return [0, 0, n, n];
  var s = n / Math.max(w, h), dw = Math.max(1, Math.round(w * s)),
      dh = Math.max(1, Math.round(h * s));
  return [Math.floor((n - dw) / 2), Math.floor((n - dh) / 2), dw, dh];
}
function rimIcon(path, n, col) {
  col = col || RIMCOL.mob;
  var im = iconImg(path);
  if (!im || !im.complete || !im.naturalWidth) return null;
  n = Math.max(4, Math.round(n / 4) * 4);
  var key = path + '#' + n + '#' + col, c = RIM[key];
  if (c) return c;
  /* Sizes the map has left behind go, but only they: throwing the whole
   * cache away because it filled up meant that the third zoom step of a
   * gesture rebuilt every rim on the map at once, and that one step was
   * the hitch.  A sweep every few hundred builds, keeping what has been
   * used recently, costs nothing and never empties what is in use. */
  rimGen++;
  if (rimN > 640) {
    for (var rk in RIM) {
      if (RIM[rk].gen < rimGen - 400) {
        delete RIMB[rk]; delete RIM[rk]; rimN--;
      }
    }
  }
  /* the art at the size it will be drawn, read once for its alpha */
  var art = document.createElement('canvas');
  art.width = art.height = n;
  var g = art.getContext('2d', { willReadFrequently: true });
  var fb = fitBox(im, n);
  g.drawImage(im, fb[0], fb[1], fb[2], fb[3]);
  var d = g.getImageData(0, 0, n, n).data, x, y;
  var solid = new Uint8Array(n * n), i = 0;
  for (y = 0; y < n; y++)
    for (x = 0; x < n; x++, i++) solid[i] = d[i * 4 + 3] > 32 ? 1 : 0;
  /* one pixel of the rim colour around every solid pixel, in all eight
   * directions - black for a creature, red for a boss and so on */
  var rgb = rimRGB(col);
  c = document.createElement('canvas');
  c.width = c.height = n + 2;
  var h = c.getContext('2d');
  var out = h.createImageData(n + 2, n + 2), od = out.data;
  for (y = 0; y < n; y++) {
    for (x = 0; x < n; x++) {
      if (!solid[y * n + x]) continue;
      for (var dy = -1; dy <= 1; dy++) {
        var ty = y + dy + 1;
        if (ty < 0 || ty > n + 1) continue;
        for (var dx = -1; dx <= 1; dx++) {
          var tx = x + dx + 1;
          if (tx < 0 || tx > n + 1) continue;
          var q = (ty * (n + 2) + tx) * 4;
          od[q] = rgb[0]; od[q + 1] = rgb[1]; od[q + 2] = rgb[2];
          od[q + 3] = 255;
        }
      }
    }
  }
  h.putImageData(out, 0, 0);
  c.gen = rimGen;
  RIM[key] = c; RIMB[key] = art; rimN++;
  return c;
}
/* The picture on its own, at the size it is drawn, kept beside the rim
 * grown from it: the two are laid down as two passes over a whole list
 * rather than as one mark at a time, because the picture going down last
 * is what takes the edge away where two marks of one colour meet - and
 * because the same art that was read for the silhouette is the art that
 * goes back over it, the two are known to line up to the pixel. */
function bodyIcon(path, n, col) {
  n = Math.max(4, Math.round(n / 4) * 4);
  return RIMB[path + '#' + n + '#' + (col || RIMCOL.mob)] || null;
}

/* ---- one rim, however many marks make it --------------------------------
 * Each mark is edged in the colour of what it marks, but two marks of the
 * *same* colour whose edges touch used to leave a line between them - the
 * edge of each drawn across the other, and the pair reading as two outlined
 * things rather than as one shape with one edge round the pair.  What has
 * to be painted is the union: everything inside the union has no edge, and
 * only the outside does.
 *
 * That is what two passes give, without a scratch layer anywhere.  The
 * edges all go down first, and then every picture over them: an edge that
 * lies inside another mark's picture is an edge *between* two marks of one
 * colour, and the picture going down last takes it away - which is the
 * whole of the point.  An edge standing outside every picture has nothing
 * drawn over it and stays, so the outside of the pair keeps its line.  Two
 * marks standing clear of each other come out the same as they always did,
 * because edge-then-picture is what one mark was already drawn as.
 *
 * It costs two pictures per mark and nothing else: no clustering to find
 * which marks touch, no clearing of a layer, and - the expensive one -
 * no taking the pictures back out of their edges with a composite.  That
 * composite was the whole of the lag when a map was full and zoomed out.
 *
 * A rim of another colour is not part of any of this.  A red boss edge
 * crossing a black creature edge stays two edges, because that crossing is
 * the whole reason they are two colours. */
/* The pile a pass gathers is kept as flat arrays rather than as a pile of
 * little objects: three thousand marks every frame, each its own six-field
 * object, is a hundred and eighty kilobytes of garbage per frame and the
 * collecting of it lands in the middle of the very measurements that decide
 * how fast this is.  Two of the six are picture references, which is what
 * plain arrays are for. */
var BATCH = null;
function beginRim() {
  var b = BATCH;
  if (!b) b = BATCH = { x: new Float64Array(4096), y: new Float64Array(4096),
                        n: new Float64Array(4096), h: new Float64Array(4096),
                        r: new Array(4096),     b: new Array(4096), k: 0 };
  b.k = 0;
}
function endRim() {
  var b = BATCH; BATCH = null;
  if (b && b.k) paintRim(b);
}
/* One mark: into the pile for the pass being gathered, or straight onto
 * the map where nothing is gathering - where it draws the edge and the
 * picture over it, in that order, for the same reason as below. */
function rimMark(x, y, n, half, rim, body) {
  if (!BATCH) {
    ctx.drawImage(rim, x - half, y - half);
    if (body) ctx.drawImage(body, x - n / 2, y - n / 2);
    return;
  }
  var b = BATCH, i = b.k;
  if (i === b.x.length) {                 /* only ever on the first frame
                                           * that outgrows the first guess */
    var c = i * 2, t;
    t = new Float64Array(c); t.set(b.x); b.x = t;
    t = new Float64Array(c); t.set(b.y); b.y = t;
    t = new Float64Array(c); t.set(b.n); b.n = t;
    t = new Float64Array(c); t.set(b.h); b.h = t;
    t = new Array(c); t.set(b.r); b.r = t;
    t = new Array(c); t.set(b.b); b.b = t;
  }
  b.x[i] = x; b.y[i] = y; b.n[i] = n; b.h[i] = half;
  b.r[i] = rim; b.b[i] = body; b.k = i + 1;
}
/* Every edge of the pass first, every picture of the pass after it.  The
 * order is fixed across the whole pass, not per pair: a picture may only
 * ever cover an edge, and never another picture, so the sprites still
 * overlap the way the list draws them - and because nothing here decides
 * where anything goes, no mark is remembered as anything but six numbers
 * and two pictures. */
function paintRim(b) {
  var i, k = b.k;
  for (i = 0; i < k; i++) {
    ctx.drawImage(b.r[i], b.x[i] - b.h[i], b.y[i] - b.h[i]);
  }
  for (i = 0; i < k; i++) {
    if (b.b[i]) ctx.drawImage(b.b[i], b.x[i] - b.n[i] / 2, b.y[i] - b.n[i] / 2);
  }
}
/* An ore or a plant is drawn as the thing itself, plainly and the smallest
 * of the four marks: a field of them is what is being marked - a hundred
 * and thirty-nine lilies across one valley - so it has to stay under a
 * creature at every zoom rather than only at one scale, and it has been
 * taken down again since.  Its rim is the picture's own silhouette, one
 * pixel of the colour its kind is edged in - grey for an ore, green for a
 * plant - which keeps pale art from reading as a smudge without ever
 * standing the plant inside a square. */
function spotPic(arr, i, j, icon, col) {
  /* Icon sizes move in 4px steps: a zoom gesture crosses far fewer sizes,
   * so far fewer rims are ever built, and a step is invisible. */
  var s = Math.round(Math.max(6, Math.min(26, 2.2 + sc * 8)) / 4) * 4;
  var rim = rimIcon(icon, s, col);
  if (!rim) return false;
  var body = bodyIcon(icon, s, col);
  var half = s / 2 + 1, k, X, Y;
  for (k = i; k < j; k++) {
    X = sx(arr[k][0]); Y = sy(arr[k][1]);
    if (X < -s || Y < -s || X > W + s || Y > H + s) continue;
    rimMark(X, Y, s, half, rim, body);
  }
  return true;
}
/* One pass per creature and one per item, because a stroke style is per
 * path and the whole point of splitting them is that a hundred sheep must
 * not look like the dragon standing among them.
 *
 * Where the game has a picture for what stands here, the picture is drawn
 * with a rim rather than in a frame.  The rim is the picture's own
 * silhouette, one pixel thick and in the colour of its kind - yellow for
 * an item, black for a creature, red for a boss - what a frame was for, a
 * bright creature being the same as no marker at all on a bright map,
 * without a rectangle around every one of two hundred names.  What is left telling
 * them apart is what was telling them apart before the picture came: how
 * big they are, and the glyph below for anything that is still loading.
 * Three sizes, because the three things being marked are not equals - an
 * item is a single point a player is hunting for and is the largest of the
 * three, a creature is one of a hundred and is a little smaller than it,
 * and a boss has to be found among its own pack, so it is bigger than the
 * pack it hides in.
 *
 * A picture that has not arrived yet falls back to the glyph below, so a
 * slow request leaves the map drawn rather than blank: a cross for a
 * creature, a diamond for an object on the ground, and the same cross
 * inside a ring for a boss - the ring is the game's own class made
 * visible, and it is what keeps Snow Demon from hiding in its own pack. */
function drawRun(a, i, j, r, glyph, icon, sq, rimCol) {
  var im = icon && iconImg(icon);
  if (im && im.complete && im.naturalWidth) {
    var s = glyph === 0 ? Math.max(12, Math.min(52, 6 + sc * 16)) :
            glyph === 2 ? Math.max(15, Math.min(64, 7 + sc * 21)) :
            glyph === 3 ? Math.max(19, Math.min(78, 10 + sc * 25)) :
            /* The statue cutout is a tall monument, not a square parchment,
             * so it is drawn forty per cent over the same box the map gets -
             * floor, formula and ceiling all scaled together - and it carries
             * its own glyph so the parchment beside it keeps its own size.
             * The cutout is filed on a square canvas, or the rim would squash
             * it, and only 167 of its 199 rows hold art, so the box goes up by
             * 199/172 as well: together that leaves the monument the height it
             * had before the canvas was squared, at its true width. */
            glyph === 4 ? Math.max(31, Math.min(126, (14 + sc * 35) * 1.157)) :
                          Math.max(11, Math.min(50, 5 + sc * 16));
    /* The square a creature is filed under is cropped to the creature,
     * while the in-game picture has the ground it stands on around it, so
     * at one and the same box the square reads the larger of the two.  It
     * comes down a little for that - and only while it is the square being
     * drawn: the full-body art keeps the size it has always had, whichever
     * way the switch is set. */
    /* Icons stop growing past 48px (72 for the tall statue cutout): past
     * that a mark is big enough to read, and rim dilation is quadratic in
     * this number, so an uncapped zoom step is where a zoom hitch came
     * from.  Steps of 4px match the rim cache above. */
    s = glyph === 4 ? Math.min(72, Math.round(s / 4) * 4)
                    : Math.min(48, Math.round(s / 4) * 4);
    if (sq) s = Math.max(8, Math.round(s * 0.85 / 4) * 4);
    /* Far out, an icon is a few pixels anyway: one batched dot pass in the
     * kind's colour instead of thousands of rimmed pictures the eye can't
     * resolve.  This is what keeps busy maps draggable while zoomed out. */
    if (sc < 0.5) {
      var ds = Math.max(2.5, Math.min(6, sc * 10)), dh = ds / 2;
      ctx.fillStyle = rimCol ||
            (glyph === 0 ? COL.item :
             glyph === 2 ? COL.boss : COL.mob);
      ctx.beginPath();
      for (var q = i; q < j; q++) {
        var QX = sx(a[q][0]), QY = sy(a[q][1]);
        if (QX < -3 || QY < -3 || QX > W + 3 || QY > H + 3) continue;
        ctx.rect(QX - dh, QY - dh, ds, ds);
      }
      ctx.fill();
      return;
    }
    var n = Math.max(1, Math.round(s)),
        col = rimCol ||
              (glyph === 0 ? RIMCOL.item :
               glyph === 2 ? RIMCOL.boss : RIMCOL.mob),
        rim = rimIcon(icon, n, col),
        body = rim ? bodyIcon(icon, n, col) : null;
    var half = n / 2 + (rim ? 1 : 0), pts = [], k, X, Y;
    for (k = i; k < j; k++) {
      X = sx(a[k][0]); Y = sy(a[k][1]);
      if (X < -n || Y < -n || X > W + n || Y > H + n) continue;
      pts.push(X, Y);
    }
    for (k = 0; k < pts.length; k += 2) {
      if (rim) rimMark(pts[k], pts[k + 1], n, half, rim, body);
      else {
        var fb2 = fitBox(im, n);
        ctx.drawImage(im, pts[k] - half + fb2[0], pts[k + 1] - half + fb2[1],
                      fb2[2], fb2[3]);
      }
    }
    return;
  }
  /* No picture arrived: filled marks, never hollow.  A creature is a
   * filled red disc, a boss the same in pink inside its ring, and every
   * single object - item, treasure, statue, NPC - a filled diamond in its
   * own colour, the shape the sidebar rows already show. */
  ctx.fillStyle = glyph === 1 ? '#ff3b30' :
                  glyph === 2 ? COL.boss : ctx.strokeStyle;
  ctx.beginPath();
  for (var t = i; t < j; t++) {
    var PX = sx(a[t][0]), PY = sy(a[t][1]);
    if (PX < -24 || PY < -24 || PX > W + 24 || PY > H + 24) continue;
    if (glyph === 0 || glyph === 3 || glyph === 4) {
      ctx.moveTo(PX, PY - r * 1.45);
      ctx.lineTo(PX + r * 1.45, PY);
      ctx.lineTo(PX, PY + r * 1.45);
      ctx.lineTo(PX - r * 1.45, PY);
      ctx.closePath();
    } else {
      ctx.moveTo(PX + r, PY);
      ctx.arc(PX, PY, r, 0, 6.2831853);
    }
  }
  ctx.fill();
  if (glyph === 2) {                       /* the boss keeps its ring */
    ctx.beginPath();
    for (var u = i; u < j; u++) {
      var QX = sx(a[u][0]), QY = sy(a[u][1]);
      if (QX < -24 || QY < -24 || QX > W + 24 || QY > H + 24) continue;
      ctx.moveTo(QX + r * 2.1, QY);
      ctx.arc(QX, QY, r * 2.1, 0, 6.2831853);
    }
    ctx.stroke();
  }
}
function drawLayer(a, names, pfx, glyphFor, r, once) {
  var i = 0, n = a.length;
  /* once: several rows can carry the same ground - every world buff
   * stands at every post - and a point drawn under one of them is the
   * same point under the rest, so it is marked and never laid down
   * twice.  The mark is made only when the point is actually drawn, so
   * hiding one row still lets the same place come up under another. */
  var drawn = once ? Object.create(null) : null;
  while (i < n) {
    var ref = a[i][2], j = i;
    while (j < n && a[j][2] === ref) j++;
    var nm = names[ref];
    if (nm !== undefined && shown(pfx + nm)) {
      /* a pass may ask for only part of a list - the creatures without
       * their bosses, say - so that each colour can be edged, and merged,
       * on its own */
      var g = glyphFor(ref);
      if (g) {
        ctx.strokeStyle = g.c;
        /* A glyph can be a picture chosen per mark rather than per run -
         * the buff pins each take their own place in the turn - so the
         * run is broken wherever the picture, or the already-drawn
         * state, changes under it. */
        var turn = typeof g.icon === 'function', k = i;
        while (k < j) {
          var pth = turn ? g.icon(a[k][0], a[k][1]) : g.icon, e = k;
          while (e < j) {
            if (turn && g.icon(a[e][0], a[e][1]) !== pth) break;
            if (drawn) {
              var key = a[e][0] + ',' + a[e][1];
              if (drawn[key]) break;
              drawn[key] = 1;
            }
            e++;
          }
          if (e > k) drawRun(a, k, e, r, g.g, pth, g.sq, g.rim);
          k = e > k ? e : k + 1;         /* one already drawn, or moved on */
        }
      }
    }
    i = j;
  }
}
function drawMobs() {
  if (!showSp) return;
  if (sc < 0.11) return;               /* the packs are not worth 3000
                                         * hairlines from orbit */
  var r = sc > 1 ? 4.4 : 3.0;
  var m = mkey, boss = BOS[m] || {};
  ctx.lineWidth = 1.35;
  /* The world buff posts go down before any of it, under everything else on
   * the map.  A post is a piece of ground rather than a thing standing on
   * it - any buff that rolls on this island can appear at any of them - so
   * where a creature, a boss or a quest object happens to share the square,
   * that is what a reader wants to see on top. */
  beginRim();
  buffRot(m);
  drawLayer(D.items[m] || [], itmNames(m), IKEY, function (ref) {
    if ((ICAT[m] || [])[ref] !== 'buff') return null;
    return { c: 'rgba(255,182,74,.96)', g: 0,
             icon: BUFF_ROT.length ? buffIconAt : itmIcon(m, ref) };
  }, r, true);
  endRim();
  /* Three passes, one colour each, because an edge is only merged with an
   * edge of its own colour: the creatures, then their bosses over them - a
   * boss has to be found among its own pack - and the items over both, a
   * quest object being a single point a player is looking for and not
   * something to bury under whatever group happens to stand beside it. */
  beginRim();
  drawLayer(D.mobs[m] || [], creNames(m), CKEY, function (ref) {
    if (boss[ref]) return null;        /* the boss pass takes it */
    return { c: 'rgba(120,220,255,.78)', g: 1, icon: creIcon(m, ref),
             sq: creIsSquare(m, ref) };
  }, r);
  endRim();
  beginRim();
  drawLayer(D.mobs[m] || [], creNames(m), CKEY, function (ref) {
    if (!boss[ref]) return null;
    return { c: 'rgba(255,110,168,.95)', g: 2, icon: creIcon(m, ref),
             sq: creIsSquare(m, ref) };
  }, r);
  endRim();
  beginRim();
  drawLayer(D.items[m] || [], itmNames(m), IKEY, function (ref) {
    if ((ICAT[m] || [])[ref] === 'buff') return null;   /* already down */
    return { c: 'rgba(255,182,74,.96)', g: 0, icon: itmIcon(m, ref) };
  }, r, true);
  endRim();
}
/* The treasure maps are drawn over everything else, one torn parchment per
 * dig spot, edged in the same bronze the sidebar marks them with.  They are
 * not gated by the zoom cutoff the packs are: fifty-odd maps across a whole
 * world is not three thousand hairlines, and a map is what a player came
 * here to find, so it stays on the picture from orbit down. */
function drawTreasures() {
  if (!showSp) return;
  var m = mkey, r = sc > 1 ? 4.4 : 3.0;
  beginRim();
  drawLayer(D.treasures[m] || [], treNames(m), TKEY, function (ref) {
    return { c: 'rgba(212,149,58,.97)', g: 3, icon: treIcon(m, ref),
             rim: RIMCOL.treasure };
  }, r);
  endRim();
}
/* The statues are drawn with the monument's own cutout, edged in the one
 * blue nothing else on the map wears - a placed monument the character
 * reads a stat off, so it is a single point being looked for, and it stays
 * on the picture from orbit down the way a treasure map does. */
function drawStatues() {
  if (!showSp) return;
  var m = mkey, r = sc > 1 ? 4.6 : 3.2;
  ctx.lineWidth = 1.6;
  ctx.strokeStyle = 'rgba(77,157,224,.95)';
  drawLayer(D.statues[m] || [], statueNames(m), SKEY, function (ref) {
    return { c: 'rgba(77,157,224,.95)', g: 4, icon: statueIcon(m), rim: RIMCOL.statue };
  }, r);
}
/* The NPCs are drawn as the diamond the sidebar marks them with, edged
 * in the one green nothing else on the map wears - a person standing
 * around is a single point being looked for, the same shape and weight
 * as a statue, only in its own colour. */
function drawNpcs() {
  if (!showSp) return;
  var m = mkey, r = sc > 1 ? 4.6 : 3.2;
  ctx.lineWidth = 1.6;
  ctx.strokeStyle = 'rgba(63,209,106,.95)';
  drawLayer(D.npcs[m] || [], npcNames(m), NKEY, function (ref) {
    return { c: 'rgba(63,209,106,.95)', g: 0, icon: null, rim: RIMCOL.npc };
  }, r);
}
function drawRegions() {
  if (!showPl) return;
  var a = D.regions[mkey] || [];
  if (!a.length) return;
  var fs = Math.max(9.5, Math.min(17, 9 + sc * 2.4));
  if (sc < 0.09 && fs < 12) return;
  ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
  ctx.font = '600 ' + fs.toFixed(1) + 'px ' +
             '"Segoe UI",system-ui,-apple-system,Roboto,sans-serif';
  ctx.lineJoin = 'round';
  ctx.strokeStyle = 'rgba(0,0,0,.82)';
  ctx.fillStyle = 'rgba(226,238,250,.94)';
  ctx.lineWidth = Math.max(2.4, fs * 0.34);
  for (var i = 0; i < a.length; i++) {
    var X = sx(a[i].ix), Y = sy(a[i].iy);
    if (X < -140 || Y < -30 || X > W + 140 || Y > H + 30) continue;
    if (sc < 0.16 && !a[i].big) continue;
    ctx.strokeText(a[i].title, X, Y);
    ctx.fillText(a[i].title, X, Y);
  }
}
/* A scale bar only earns its place if the length it names is the length
 * it is: pick the largest round number of metres that fits in 110 px. */
var STEPS = [10, 25, 50, 100, 250, 500, 1000, 2000];
function drawScale() {
  var pxPerM = sc * pxPerMetre() * PPX, want = 0, len = 0;
  for (var i = STEPS.length - 1; i >= 0; i--) {
    var L = STEPS[i] * pxPerM;
    if (L <= 110) { want = STEPS[i]; len = L; break; }
  }
  if (!want) return;
  var x0 = 16, y0 = H - 18;
  ctx.strokeStyle = 'rgba(226,238,250,.8)'; ctx.lineWidth = 1.4;
  ctx.beginPath();
  ctx.moveTo(x0, y0 - 4); ctx.lineTo(x0, y0); ctx.lineTo(x0 + len, y0); ctx.lineTo(x0 + len, y0 - 4);
  ctx.stroke();
  ctx.textAlign = 'left'; ctx.textBaseline = 'bottom';
  ctx.font = '11px ' + '"Segoe UI",system-ui,sans-serif';
  ctx.fillStyle = 'rgba(226,238,250,.9)';
  ctx.fillText(want + ' m', x0, y0 - 6);
}
function drawSel() {
  if (!sel) return;
  var X = sx(sel.ix), Y = sy(sel.iy);
  if (X < -40 || Y < -40 || X > W + 40 || Y > H + 40) return;
  ctx.strokeStyle = '#fff'; ctx.lineWidth = 2;
  ctx.beginPath(); ctx.arc(X, Y, spotR() + 4.5, 0, 6.2831853); ctx.stroke();
  ctx.strokeStyle = 'rgba(0,0,0,.65)'; ctx.lineWidth = 1;
  ctx.beginPath(); ctx.arc(X, Y, spotR() + 6, 0, 6.2831853); ctx.stroke();
}
/* High-frequency input (wheel bursts, drags, icon landings) must never
 * stack full repaints: one frame, one paint.  Discrete actions call draw();
 * continuous gestures paint previews and settle into draw(). */
var paintQueued = false;
function draw() {
  if (paintQueued) return;
  paintQueued = true;
  requestAnimationFrame(function () { paintQueued = false; renderFrame(); });
}
/* What identifies the scene in the layer: world, toggles, picture set and
 * every filter generation, plus the buff turn (it repaints itself). */
function layerKeyNow() {
  return mkey + '|' + showSp + '|' + showPl + '|' +
         filterGen + '|' + buffRotT + '|' + (ready ? 1 : 0) + '|' + icoGen;
}
/* The expensive half: background plus every mark, drawn once into the
 * layer instead of on every pan frame.  drawSel/scale/status stay on the
 * main canvas - they are cheap and always current. */
function renderScene() {
  var m = MAP[mkey];
  ctx.fillStyle = '#0a0e13';
  ctx.fillRect(0, 0, W, H);
  /* ready is this window's own bookkeeping; the picture's is complete.
   * They should agree, and when they do not it is the picture that is
   * right, so the guard is on both - a half loaded image handed to
   * drawImage throws, and an exception here would abandon the rest of
   * the frame and leave a stale status bar claiming another world. */
  if (ready && img.complete && img.naturalWidth > 0) {
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(img, sx(0), sy(0), PW * sc, PH * sc);
    /* a hairline of the picture's own edge, so an empty sea still reads
       as the edge of the world rather than as more of the same */
    ctx.strokeStyle = 'rgba(90,176,255,.16)'; ctx.lineWidth = 1;
    ctx.strokeRect(sx(0) + .5, sy(0) + .5, PW * sc - 1, PH * sc - 1);
    drawRegions();
    drawMobs();
    drawTreasures();
    drawSpots();
    drawStatues();
    drawNpcs();
  }
}
/* Re-render the layer only when something it holds changed: another world,
 * another zoom, another filter, a buff turn, a resize - or the pan walked
 * past its margin.  The scene functions read the globals, so the layer is
 * rendered as a bigger viewport over the same centre. */
function ensureLayer() {
  var key = layerKeyNow();
  var need = !layer || layerKey !== key || layerSc !== sc ||
             layerW !== W + 2 * LAYERM || layerH !== H + 2 * LAYERM ||
             Math.abs(vx - layerVx) * sc > LAYERM ||
             Math.abs(vy - layerVy) * sc > LAYERM;
  if (!need) return;
  layerW = W + 2 * LAYERM; layerH = H + 2 * LAYERM;
  if (!layer) { layer = document.createElement('canvas'); layerCtx = layer.getContext('2d'); }
  layer.width = Math.max(1, Math.round(layerW * dpr));
  layer.height = Math.max(1, Math.round(layerH * dpr));
  var _ctx = ctx, _W = W, _H = H;
  ctx = layerCtx; W = layerW; H = layerH;
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  renderScene();
  ctx = _ctx; W = _W; H = _H;
  layerVx = vx; layerVy = vy; layerSc = sc; layerKey = key;
}
function blitLayer() {
  ctx.fillStyle = '#0a0e13';
  ctx.fillRect(0, 0, W, H);
  if (!layer) return;
  var dx = (layerVx - vx) * sc - LAYERM, dy = (layerVy - vy) * sc - LAYERM;
  ctx.drawImage(layer, 0, 0, layer.width, layer.height, dx, dy, layerW, layerH);
}
function paintOverlays() { drawSel(); drawScale(); status(); }
function snapMain() {
  if (!snap) { snap = document.createElement('canvas'); snapCtx = snap.getContext('2d'); }
  if (snap.width !== cv.width || snap.height !== cv.height) {
    snap.width = cv.width; snap.height = cv.height;
  }
  snapCtx.setTransform(1, 0, 0, 1, 0, 0);
  snapCtx.drawImage(cv, 0, 0);
  snapSc = sc;
}
function renderFrame() {
  ensureLayer();
  blitLayer();
  paintOverlays();
  snapMain();
}
function layerUsable() {
  return layer && layerKey === layerKeyNow() && layerSc === sc &&
         layerW === W + 2 * LAYERM && layerH === H + 2 * LAYERM;
}
/* A pan frame: blit the layer while it covers the view (same zoom, same
 * filters, inside its margin) and let the settle land the crisp frame.
 * Sustained drags - Madrigal especially - stay one blit per frame instead
 * of a full re-render every few hundred pixels. */
function previewPan() {
  if (layerUsable()) {
    blitLayer();
    paintOverlays();
    snapMain();   /* keep zoom previews honest after a drag */
    armSettle();
  } else {
    renderFrame();
  }
}
/* A zoom frame: scale the last full frame about the cursor.  One blit,
 * full icons included, slightly soft for a blink; the settle lands crisp. */
function previewZoom(px, py) {
  if (!snap || snapSc <= 0 || snap.width !== cv.width || cv.width < 2) {
    renderFrame(); return;
  }
  var k = sc / snapSc;
  ctx.fillStyle = '#0a0e13';
  ctx.fillRect(0, 0, W, H);
  ctx.drawImage(snap, px - px * k, py - py * k, W * k, H * k);
  drawScale(); status();
}
/* The crisp landing after a gesture stops moving. */
function armSettle() {
  gesture = true; elTip.hidden = true;
  if (settleT) clearTimeout(settleT);
  settleT = setTimeout(function () { settleT = 0; gesture = false; draw(); }, 140);
}
/* ------------------------------------------------------------ hit tests */
function nearest(px, py) {
  var arr = D.spots[mkey] || [], best = null, bd = 11 * 11;
  for (var i = 0; i < arr.length; i++) {
    if (!shown(SPECIES[arr[i][2]].name)) continue;
    var dx = sx(arr[i][0]) - px, dy = sy(arr[i][1]) - py, d = dx * dx + dy * dy;
    if (d <= bd) { bd = d; best = arr[i]; }
  }
  return best;
}

/* ------------------------------------------------------------- controls */
var elGroups  = document.getElementById('groups');
var elTip     = document.createElement('div');
elTip.className = 'tip'; elTip.hidden = true;
document.getElementById('stage').appendChild(elTip);
var elDetail = document.getElementById('detail');

/* No per-world buttons: the step buttons beside the jump list walk
 * the group filter's list, and the jump list holds the same worlds. */
function buildTabs() {
}
function syncTabs() {
  var vis = visibleMaps(), at = vis.map(function (m) {
    return m.key;
  }).indexOf(mkey);
  var pv = vis[at - 1], nx = vis[at + 1];
  if (elPrev) elPrev.title = pv ? 'Previous world: ' + pv.title
                                 : 'First world in this group';
  if (elNext) elNext.title = nx ? 'Next world: ' + nx.title
                                 : 'Last world in this group';
  syncJump();
}
/* The jump list beside the step buttons: the worlds the group filter
 * leaves in view, and nothing else, so the two dropdowns always agree -
 * narrowing the first narrows the second to the same list.  With every
 * group in view the worlds stand under their group headings, in the
 * same order the step buttons walk them. */
var elJump = document.getElementById('mapjump');
var elPrev = document.getElementById('mapprev');
var elNext = document.getElementById('mapnext');
var GRPTITLE = { main: 'Main Maps', dungeons: 'Dungeons', fields: 'Others',
                 event: 'Event', farm: 'Farm Maps',
                 platin: 'Platin Level Maps', royal: 'Royal Level Maps' };
var GROUPS = ['all', 'main', 'dungeons', 'fields', 'event',
              'farm', 'platin', 'royal'];
/* The tier filters list their worlds lowest-tier-first, so the step
 * buttons walk the tiers in order; Farm Maps keep run order; every other
 * filter keeps map order. */
function visibleMaps() {
  if (mgrp === 'farm') {
    var out = [];
    FARM_ORDER.forEach(function (k) { if (MAP[k]) out.push(MAP[k]); });
    return out;
  }
  if (mgrp === 'platin' || mgrp === 'royal') {
    var min = mgrp === 'platin' ? PLATIN_MIN : ROYAL_MIN;
    return MAPS.filter(function (m) { return inGroup(m.key); }).sort(function (a, b) {
      var x = min[a.key], y = min[b.key];
      if (x == null && y == null) return a.title.localeCompare(b.title);
      if (x == null) return 1;
      if (y == null) return -1;
      return (x - y) || a.title.localeCompare(b.title);
    });
  }
  return MAPS.filter(function (m) { return inGroup(m.key); });
}
function buildJumpOptions() {
  if (!elJump) return;
  var cur = elJump.value || mkey;
  elJump.textContent = '';
  /* Farm Maps and the tier filters are world sets, not world groups.
   * Farm splits into its penya and ECP tracks; the tiers list lowest
   * first.  Every set world carries its stage or tier beside the name
   * instead of a node count.  Sets never appear inside the All Worlds
   * jump list. */
  function jumpOption(m, text, parent) {
    var o = document.createElement('option');
    o.value = m.key;
    o.textContent = text;
    parent.appendChild(o);
  }
  if (mgrp === 'farm') {
    [['Penya', FARM_PENYA], ['ECP', FARM_ECP]].forEach(function (tr) {
      var og = document.createElement('optgroup');
      og.label = GRPTITLE.farm + ' - ' + tr[0];
      var any = false;
      tr[1].forEach(function (k) {
        if (!MAP[k] || !inGroup(k)) return;
        any = true;
        var fl = FARM_LABEL[k];
        jumpOption(MAP[k], fl ? MAP[k].title + ' [' + fl + ']' : MAP[k].title, og);
      });
      if (any) elJump.appendChild(og);
    });
  } else if (mgrp === 'platin' || mgrp === 'royal') {
    var set = visibleMaps();
    if (set.length) {
      var so = document.createElement('optgroup');
      so.label = GRPTITLE[mgrp];
      elJump.appendChild(so);
      var tp = mgrp === 'platin' ? 'P' : 'RL';
      set.forEach(function (m) {
        var tx = tierText(m.key, tp);
        jumpOption(m, tx ? m.title + ' ' + tx : m.title, so);
      });
    }
  } else {
  var groups = mgrp === 'all' ? ['main', 'dungeons', 'fields', 'event']
                              : [mgrp];
  groups.forEach(function (g) {
    var list = MAPS.filter(function (m) { return groupOf(m.key) === g; });
    if (!list.length) return;
    /* Dungeons split into difficulty tiers along display order, in every
     * view that lists them. */
    var chunks = g === 'dungeons'
      ? dungeonTiers(list).map(function (t) {
          return { label: GRPTITLE.dungeons + ' - ' + t.title, list: t.list };
        })
      : [{ label: GRPTITLE[g], list: list }];
    chunks.forEach(function (ch) {
      if (mgrp === 'all' || g === 'dungeons') {
        var og = document.createElement('optgroup');
        og.label = ch.label;
        elJump.appendChild(og);
        ch.list.forEach(function (m) { jumpOption(m, m.title, og); });
      } else {
        ch.list.forEach(function (m) { jumpOption(m, m.title, elJump); });
      }
    });
  });
  }
  if (elJump.querySelector('option[value="' + cur + '"]')) elJump.value = cur;
}
function buildJump() {
  if (!elJump) return;
  buildJumpOptions();
  elJump.addEventListener('change', function () {
    if (elJump.value) setMap(elJump.value);
  });
  if (elPrev) elPrev.addEventListener('click', function () { stepMap(-1); });
  if (elNext) elNext.addEventListener('click', function () { stepMap(1); });
}
function syncJump() {
  if (!elJump) return;
  if (elJump.value !== mkey) elJump.value = mkey;
  var vis = visibleMaps(), at = vis.map(function (m) {
    return m.key;
  }).indexOf(mkey);
  if (elPrev) elPrev.disabled = at <= 0;
  if (elNext) elNext.disabled = at < 0 || at >= vis.length - 1;
}
function stepMap(d) {
  var vis = visibleMaps();
  var at = vis.map(function (m) { return m.key; }).indexOf(mkey);
  var nx = vis[at + d];
  if (nx) setMap(nx.key);
}
/* Explicit world groups: the header lists Main Maps, Dungeons, Others
 * and Event in the order below.  Farm, Platin and Royal are sets instead
 * of groups - defined below - so a world keeps its group and additionally
 * joins a set.  Ruins of Desire closes the Dungeons list. */
var GROUP_OF = {};
[['main', ['cardia', 'cave', 'kogarashi', 'rifloatingcity', 'wdmadrigal']],
 ['fields', ['wdheaven01', 'wdheaven02', 'wdheaven03', 'wdheaven04',
              'wdheaven05', 'wdvolcane', 'wdvolcanered', 'wdcisland', 'wdkebaras',
             'wdrartesia', 'wddarkrartesia', 'duflmas', 'wdeden',
             'wdoldkebaras', 'wdaurania', 'duflvenux', 'wdveruxlevel01',
             'amethyst', 'dreadcove', 'wdoasisdesert', 'dudadk',
             'duflmasnew', 'dusatemple', 'wdguildwar']],
 ['dungeons', ['dupremiummap', 'duominous', 'duominous_1', 'durustia',
               'durustia_1', 'dubehamah', 'dukalgas', 'dudreadfulcave',
               'dusanpres', 'duherneos', 'ducontaminatedtrails',
               'dumysticdungeon', 'dudarkforest', 'duwaterdungeon',
               'duflcavenew', 'dumelind', 'duholytemplemap', 'dupandora',
               'dugolemtemple', 'dutrigremruins', 'duethralgarden',
               'duskydungeon01', 'dutempleofthedeep', 'duveruxwinter',
               'duveruxfiresmoke', 'duumbala01', 'dupiratecave',
               'duelysianisland', 'wdveruxjpinfinitycastle', 'dudracmaus',
               'duveruxruins']],
 ['event', ['wdpokeevent01', 'wdcandyworld', 'duhalloween',
            'wdeasterevent']]
].forEach(function (gd) {
  gd[1].forEach(function (k) { GROUP_OF[k] = gd[0]; });
});
function groupOf(k) {
  return GROUP_OF[k] || 'unsorted';
}
/* Which worlds hold a Platin- or Royal-tier monster, read off the live
 * level bank.  The sidebar already sorts monsters lowest-tier-first, so
 * stepping through either set walks every tiered monster in order.
 * Madrigal's single Royal entry is a data outlier the Royal set leaves
 * out on purpose. */
var PLATIN_MAPS = Object.create(null), ROYAL_MAPS = Object.create(null);
var PLATIN_MIN = Object.create(null), PLATIN_MAX = Object.create(null);
var ROYAL_MIN = Object.create(null), ROYAL_MAX = Object.create(null);
MAPS.forEach(function (m) {
  var arr = CLVL[m.key] || [];
  for (var i = 0; i < arr.length; i++) {
    var c = arr[i];
    if (!c) continue;
    var n = parseInt(String(c).replace(/[^0-9]/g, ''), 10);
    if (c.charAt(0) === 'P') {
      PLATIN_MAPS[m.key] = 1;
      if (!isNaN(n)) {
        if (PLATIN_MIN[m.key] == null || n < PLATIN_MIN[m.key]) PLATIN_MIN[m.key] = n;
        if (PLATIN_MAX[m.key] == null || n > PLATIN_MAX[m.key]) PLATIN_MAX[m.key] = n;
      }
    } else if (c.slice(0, 2) === 'RL' && m.key !== 'wdmadrigal') {
      ROYAL_MAPS[m.key] = 1;
      if (!isNaN(n)) {
        if (ROYAL_MIN[m.key] == null || n < ROYAL_MIN[m.key]) ROYAL_MIN[m.key] = n;
        if (ROYAL_MAX[m.key] == null || n > ROYAL_MAX[m.key]) ROYAL_MAX[m.key] = n;
      }
    }
  }
});
/* The dungeon difficulty tiers along display order: each tier runs
 * through its endpoint world - Beginner through the Catacombs, Shiny
 * through the Rhisis Trails, Black through the Golem Temple, Soulreaver
 * through the Deep Temple, Royal through the Ruins of Desire - and
 * anything past the last endpoint stays in the last tier.  Tiers are
 * contiguous slices, so the step buttons walk straight through them. */
var DUNGEON_TIERS = [
  { title: 'Beginner', until: 'dudreadfulcave' },
  { title: 'Shiny', until: 'ducontaminatedtrails' },
  { title: 'Black', until: 'dugolemtemple' },
  { title: 'Soulreaver', until: 'dutempleofthedeep' },
  { title: 'Royal', until: 'duveruxruins' }
];
function dungeonTiers(list) {
  var out = [], cur = [], ti = 0, last = DUNGEON_TIERS.length - 1;
  list.forEach(function (m) {
    cur.push(m);
    if (ti < last && m.key === DUNGEON_TIERS[ti].until) {
      out.push({ title: DUNGEON_TIERS[ti].title, list: cur });
      cur = []; ti++;
    }
  });
  if (cur.length) out.push({ title: DUNGEON_TIERS[ti].title, list: cur });
  return out;
}
/* The Farm Maps set: penya and ECP farming stages, each track in run
 * order - Penya from Tower 3 to Dread Cove, then ECP from Dark Rartesia
 * to Celestara - with the stage printed beside the world. */
var FARM_PENYA = ['wdheaven03', 'duflmas', 'wdoasisdesert', 'dreadcove'];
var FARM_ECP = ['wddarkrartesia', 'amethyst', 'rifloatingcity'];
var FARM_ORDER = FARM_PENYA.concat(FARM_ECP);
var FARM_LABEL = { 'wdheaven03': 'Penya Beginner', 'duflmas': 'Penya Beginner 2',
                   'wddarkrartesia': 'ECP Beginner', 'wdoasisdesert': 'Penya Midgame',
                   'amethyst': 'ECP Midgame', 'dreadcove': 'Penya Endgame',
                   'rifloatingcity': 'ECP Endgame' };
var FARM_MAPS = Object.create(null);
FARM_ORDER.forEach(function (k) { if (MAP[k]) FARM_MAPS[k] = 1; });
/* Whether a world belongs to the active header filter.  The farm and tier
 * filters are sets rather than groups, so membership is asked of the set -
 * a jump to a listed world must not fall back to that world's static group. */
function inGroup(k) {
  if (mgrp === 'all') return true;
  if (mgrp === 'farm') return !!FARM_MAPS[k];
  if (mgrp === 'platin') return !!PLATIN_MAPS[k];
  if (mgrp === 'royal') return !!ROYAL_MAPS[k];
  return groupOf(k) === mgrp;
}
/* The tier printed beside a tiered world, in the sidebar's own tier
 * writing: one tier reads <P 1> or <RL. 15>, a spread reads <P 6-12>. */
function tierText(k, tier) {
  var mn = tier === 'P' ? PLATIN_MIN[k] : ROYAL_MIN[k],
      mx = tier === 'P' ? PLATIN_MAX[k] : ROYAL_MAX[k];
  if (mn == null) return '';
  var p = tier === 'P' ? 'P ' : 'RL. ';
  if (mx == null || mx === mn) return '<' + p + mn + '>';
  return '<' + p + mn + '-' + mx + '>';
}
var mgrp = 'all';
var elGroup = document.getElementById('mapgroup');
function setGroup(g, rebuild) {
  if (GROUPS.indexOf(g) < 0) g = 'all';
  mgrp = g;
  if (elGroup) elGroup.value = g;
  if (rebuild !== false) { buildTabs(); buildJumpOptions(); syncTabs(); }
  save();
}

function speciesOfThisMap() {
  return SPECIES.filter(function (s) { return s.map === mkey; });
}
/* Every sub-level heading - an area, and a set of items over areas - is
 * indented and every one of them folds, whether or not the group above it
 * was split into sets.  A heading owns the rows under it and every heading
 * too small to out-rank it, up to the next one its own size: an area stops
 * at the next area or set, a set stops at the next set and takes its areas
 * with it.  The state is remembered for the session the way a group's own
 * open state is, and a search overrides it rather than hiding what it has
 * just found - a fold from before must not keep a match out of sight. */
var HOPEN = Object.create(null);
/* Which folds a search took out of the way, so that it can put them back:
 * a heading or a whole list it opened is listed here the once, and the
 * search closing closes it again.  A fold made by hand while the search
 * was up was never ours to move, and stays where the reader put it - both
 * through the typing that follows and after the box is cleared. */
var HUNF = [], GUNF = [];
function refold() {
  HUNF.forEach(function (k) { HOPEN[k] = false; }); HUNF = [];
  GUNF.forEach(function (t) { GOPEN[t] = false; }); GUNF = [];
}
/* A fold made with the search in the box is the reader's, not the search's,
 * so it is written down as one: the search opens what it found so that a
 * fold from before cannot keep a match out of sight, and without this line
 * it would take the reader's fold for an old one and spring a valley they
 * had just put away open again on the very next character.  Opening the
 * same thing again by hand is the reader's last word on it, so it comes off
 * the list and the search leaves it open when it goes.  Every fold is made
 * here - the arrow on a heading, the arrow on a list, the two switches over
 * the whole sidebar - so that a fold is the reader's wherever it was made. */
function readerFold(list, key, open) {
  var i = list.indexOf(key);
  if (!open) { if (query && i < 0) list.push(key); }
  else if (i >= 0) list.splice(i, 1);
}
function headLevel(el) { return el.classList.contains('cat') ? 2 : 1; }
function ownsOf(hd, lvl) {
  var kids = hd.parentNode.children, out = [], i = [].indexOf.call(kids, hd) + 1;
  while (i < kids.length) {
    var el = kids[i];
    if (el._stray) break;   /* entries without a category stand outside
      every fold, even the last category's */
    if ((el.classList.contains('zone') || el.classList.contains('cat')) &&
        headLevel(el) >= lvl) break;
    out.push(el); i++;
  }
  return out;
}
function setFold(hd, open) {
  hd.dataset.open = open ? 'true' : 'false';
  (hd._own || []).forEach(function (el) { el.classList.toggle('collapsed', !open); });
  /* the arrow at the end of the list's own line follows every fold under
   * it, so it is always saying whether this list's areas are standing up */
  var grp = hd.closest ? hd.closest('.grp') : null;
  if (grp && grp._subSync) grp._subSync();
}
function foldHead(hd, key, own) {
  hd._own = own;
  hd.addEventListener('click', function (ev) {
    if (ev.target.tagName === 'INPUT') return;   /* the switch, not the fold */
    var arrow = ev.target.closest && ev.target.closest('.arrow');
    var zoneFit = !arrow && hd.classList.contains('zone') &&
                  key.indexOf('xmap:') !== 0;
    if (zoneFit) {
      /* Clicking an area's line - anywhere but its arrow - fits the view
       * around everything standing in that area.  The arrow keeps folding;
       * sets and other worlds' headings keep folding everywhere. */
      var pts = [], seen = Object.create(null), i;
      for (i = 0; i < own.length; i++) {
        var nm = own[i].dataset && own[i].dataset.name;
        if (!nm || seen[nm]) continue;
        seen[nm] = 1;
        pts = pts.concat(pinsForRow(nm));
      }
      if (pts.length) fitPoints(pts);
      return;
    }
    HOPEN[key] = HOPEN[key] === false ? true : false;
    readerFold(HUNF, key, HOPEN[key] !== false);
    setFold(hd, HOPEN[key] !== false);
  });
  /* Every area and every set starts folded while the list over it starts
   * open, so the sidebar is five lists and their thirty-seven headings
   * rather than two hundred and sixty rows, and a reader opens what they
   * came for.  A heading with no opinion yet is a folded heading - and
   * only until the session has one: what the reader set is still what the
   * list is rebuilt with, so the picture switch does not unfold anything. */
  if (HOPEN[key] === undefined) HOPEN[key] = false;
  setFold(hd, HOPEN[key] !== false);
}
/* Folding the sidebar by hand is one arrow at a time, and a reader who
 * wants the whole thing out of the way - or back - should not have to click
 * forty times.  Collapse and Expand do exactly what the arrows do, over
 * every list and every area under it at once, and like them they touch not
 * a single switch: what is drawn is the reader's, and a fold is only about
 * how much list is on the screen.  The Other-worlds hits are one more list
 * in the walk, so a search stays tidy too.  The small arrow at the end of a
 * list's own line is the same thing scoped to that one list, which is what a
 * reader wants when the ores are what is being looked at and the areas of
 * every other list are just length. */
function foldAll(open) {
  [].forEach.call(elGroups.children, function (box) {
    if (!box._buckets) return;
    GOPEN[box._title] = open;
    box.dataset.open = open ? 'true' : 'false';
    readerFold(GUNF, box._title, open);
    foldSubs(box, open);
  });
}
function foldSubs(box, open) {
  (box._buckets || []).forEach(function (b) {
    HOPEN[b.key] = open;
    readerFold(HUNF, b.key, open);
    setFold(b.el, open);
  });
  if (box._subSync) box._subSync();
}
/* Every group in the sidebar is one widget - a tri-state header over named
 * rows, each with a count - so it is built once.  Species, creatures and
 * ground objects differ only in what their rows are called and what mark
 * carries them, and none of that justifies three copies of the same
 * arithmetic.  The switch is a name, so a creature and the same creature
 * found on another world are one switch, and so a shared link reads Snow
 * Demon rather than 13498. */
function makeGroup(title, rows, unit, zones, cats, place, total) {
  if (!rows.length) return null;
  /* Rows that share one pool of positions - every world buff stands at
   * every post - would each count the whole pool, and the list would
   * claim nine times the pins the map draws.  A place is counted once by
   * where it falls, so the line over the list reports what can be seen. */
  if (total == null) total = rows.reduce(function (a, r) { return a + r.n; }, 0);
  var box = document.createElement('div');
  box.className = 'grp';
  box.dataset.open = GOPEN[title] === false ? 'false' : 'true';
  box._title = title;              /* searching for a whole list by its name */

  var h = document.createElement('h3');
  h.innerHTML = '<span class="arrow">&#9662;</span>' +
                '<input type="checkbox">' +
                '<span>' + esc(title) + '</span>' +
                '<span class="n">' + rows.length + ' ' +
                (rows.length === 1 ? unit[0] : unit[1]) +
                ' &middot; ' + total + '</span>' +
                '<button type="button" class="subf" title="Collapse the ' +
                'areas in this list, or open them again">&#9662;</button>';
  var cb = h.querySelector('input');
  var subf = h.querySelector('.subf');
  h.addEventListener('click', function (ev) {
    if (ev.target === cb) {
      /* A switch takes in what the list is showing.  With a search in the
       * box, ticking Ore takes in the ores the search left on screen and
       * leaves the ones it hid exactly where they were - otherwise a
       * search for one valley would switch off every other valley too. */
      var hit = hitMap(box);
      rows.forEach(function (s) { if (hit[s.name]) on[s.name] = cb.checked; });
      syncRows(); save(); draw(); writeHash();
    } else if (ev.target.closest && ev.target.closest('.subf')) {
      /* the areas under this one list and nothing else - every list has
       * its own, and wanting them shorter is not wanting the rest of the
       * sidebar gone */
      var any = box._buckets.some(function (b) { return HOPEN[b.key] !== false; });
      foldSubs(box, !any);
    } else {
      box.dataset.open = box.dataset.open === 'true' ? 'false' : 'true';
      GOPEN[title] = box.dataset.open === 'true';
      readerFold(GUNF, title, GOPEN[title]);
    }
  });

  /* Where a list is zoned the areas run alphabetically with Mixed last,
   * which is where anything with several homes goes.  A row nothing zoned
   * is kept rather than dropped: a world with no named ground has not put
   * anything in several of them. */
  function byZone(list) {
    if (!zones || !zones.length) return [{ title: null, rows: list }];
    var out = [];
    zones.forEach(function (z) {
      var sub = list.filter(function (s) { return s.zone === z; });
      if (sub.length) out.push({ title: z, rows: sub });
    });
    var stray = list.filter(function (s) { return zones.indexOf(s.zone) < 0; });
    if (stray.length) out.push({ title: null, rows: stray });
    return out.length ? out : [{ title: null, rows: list }];
  }

  /* An item can belong to a set before it belongs to an area.  Cardia's
   * Omega glyphs are its first quest: four gems and four map notes found
   * on the ground - the fourth pickup of the run is the Strange Recipe
   * Book the four notes make, so it is carried as the fourth note rather
   * than as a row of its own next to the notes it is made of.  Lilly's
   * seven pickups are the second run and stand as their own set the same
   * way.  A set is a
   * heading over its own areas, not one
   * more area among them, because everything under it is carried by the
   * same quest.  A world whose items are all of one kind is not split at
   * all: Kogarashi has no glyphs, and a heading over the whole list would
   * be saying nothing.
   *
   * head:true is a heading that owns the entries after it rather than
   * carrying them, so it is drawn once and its rows come from the headings
   * under it. */
  var buckets = [], sets = (cats || []).filter(function (c) {
    return rows.some(function (s) { return s.cat === c; });
  });
  if (sets.length > 1) {
    sets.forEach(function (c) {
      var mine = rows.filter(function (s) { return s.cat === c; });
      if (!mine.length) return;
      buckets.push({ title: c, cls: 'cat', rows: mine, head: true,
                     n: place ? place[c] : null });
      byZone(mine).forEach(function (g) {
        buckets.push({ title: g.title, cls: 'zone', rows: g.rows });
      });
    });
  } else {
    byZone(rows).forEach(function (g) {
      buckets.push({ title: g.title, cls: 'zone', rows: g.rows });
    });
  }

  var wrap = document.createElement('div');
  wrap.className = 'rows';
  box._buckets = [];

  /* lvl is how deep the row hangs: nothing above it but the group (0), an
   * area or a set (2), or an area inside a set (3).  The switch is placed
   * by it, so a row's checkbox always lands just after the checkbox of the
   * heading it is listed under instead of far to the left of it. */
  function addRow(s, lvl, stray) {
    if (on[s.name] === undefined) on[s.name] = true;
    var r = document.createElement('label');
    r.className = 'row' + (lvl >= 2 ? ' in' : '') + (lvl >= 3 ? ' deep' : '');
    if (stray) r._stray = true;   /* rows with no heading of their own are
      owned by no fold: closing the last area must not take them away */
    r.dataset.name = s.name;
    r.innerHTML = '<input type="checkbox"' + (on[s.name] ? ' checked' : '') + '>' +
                  (s.icon ? '<img class="ic ic-' + s.sw + '" src="' +
                            esc(asset(s.icon)) + '" alt="">'
                          : '<span class="swatch ' + s.sw + '" style="background:' +
                            s.color + '"></span>') +
                  (s.tag ? '<span class="star" title="' + esc(s.tag) +
                           '">&#9733;</span>' : '') +
                  '<span class="nm" title="' + esc(s.label) + '">' +
                    esc(s.label) + '</span>' +
                  (s.lvl ? '<span class="lvl">' + lvlText(s.lvl) + '</span>' : '') +
                  /* The second language, where the client carries both:
                   * the same creature on the German client and on the
                   * English, shown so a reader of either finds it, and
                   * kept out of the way of the count and the switch. */
                  (s.alt ? '<span class="alt" title="' + esc(s.alt) + '">' +
                            esc(s.alt) + '</span>' : '') +
                  '<span class="n">' + s.n + '</span>';
    r.querySelector('input').addEventListener('change', function (ev) {
      on[s.name] = ev.target.checked;
      syncRows(); save(); draw(); writeHash();
    });
    /* Clicking the row itself - anywhere but its switch - fits the view
     * around everything standing under that name, so what was picked is
     * what ends up on screen.  The switch keeps working: this takes the
     * label's toggle for itself instead. */
    r.addEventListener('click', function (ev) {
      if (ev.target.tagName === 'INPUT') return;
      var pts = pinsForRow(s.name);
      if (!pts.length) return;
      ev.preventDefault();
      /* Clicking a switched-off row switches it back on first, so what was
       * picked is not only framed but also drawn - otherwise a search hit
       * would zoom to an empty patch of map and stay hidden. */
      if (!on[s.name]) { on[s.name] = true; syncRows(); save(); }
      fitPoints(pts);
    });
    wrap.appendChild(r);
  }

  /* A heading is the same tri-state over a shorter list, whether it names
   * an area or a whole set of items on one quest.  The arrow beside it is
   * the fold: clicking it, or the words, takes the heading and everything
   * under it out of the list without touching a single switch. */
  function addHead(title, cls, sub, n) {
    var hd = document.createElement('div');
    hd.className = cls;
    hd.innerHTML = '<span class="arrow">&#9662;</span>' +
                   '<input type="checkbox" checked>' +
                   '<span>' + esc(CATTITLE[title] || title) + '</span>' +
                   '<span class="n">' + (n != null ? n : sub.length) + '</span>';
    var zcb = hd.querySelector('input');
    zcb.addEventListener('change', function (ev) {
      /* the same rule as the group's switch above: an area takes in the
       * entries of the search, not the ones the search took out */
      var hit = hitMap(box);
      sub.forEach(function (s) { if (hit[s.name]) on[s.name] = ev.target.checked; });
      syncRows(); save(); draw(); writeHash();
    });
    wrap.appendChild(hd);
    box._buckets.push({ cb: zcb, el: hd, rows: sub, title: title,
                        key: box._title + '|' + cls + ':' + title });
  }

  /* Where each bucket sits in the tree, which is what puts its rows'
   * switches after the switches over them: an area's rows are one level
   * below the area, an area inside a set is two below, and rows with no
   * area of their own - only a set can have those - sit level with the
   * areas the set itself holds. */
  var inCat = false;
  buckets.forEach(function (b) {
    var lvl;
    if (b.title) {
      if (b.cls === 'cat') inCat = true;
      lvl = (b.cls === 'cat' ? 1 : (inCat ? 2 : 1)) + 1;
    } else {
      lvl = inCat ? 2 : 0;
    }
    if (b.title) addHead(b.title, b.cls || 'zone', b.rows, b.n);
    if (b.head) return;                 /* owns the headings that follow */
    b.rows.forEach(function (s) { addRow(s, lvl, !b.title); });
  });

  /* Nothing is folded until now, because a heading only knows what it owns
   * once every heading after it is in the list. */
  box._buckets.forEach(function (b) {
    foldHead(b.el, b.key, ownsOf(b.el, headLevel(b.el)));
  });

  /* The small arrow at the end of the list's own line answers for the areas
   * under it: it turns when they are all folded and stands when any one of
   * them is up, so it is always telling the truth about this list.  A list
   * with no areas has nothing for it to do, and is not given the space. */
  box._subSync = function () {
    if (!box._buckets.length) { subf.style.display = 'none'; return; }
    subf.style.display = '';
    subf.classList.toggle('closed', !box._buckets.some(function (b) {
      return HOPEN[b.key] !== false;
    }));
  };
  box._subSync();

  box.appendChild(h); box.appendChild(wrap);
  elGroups.appendChild(box);
  box._cb = cb; box._rows = rows;
  syncRows();
  return box;
}

/* Counts come from the points themselves rather than from the build, so a
 * row's number is always the number of things standing under that name and
 * cannot drift from what the map is drawing. */
function countsOf(arr) {
  var c = Object.create(null);
  for (var i = 0; i < arr.length; i++) c[arr[i][2]] = (c[arr[i][2]] || 0) + 1;
  return c;
}
function buildCreatureGroup() {
  var names = creNames(mkey), cnt = countsOf(D.mobs[mkey] || []), rows = [];
  var zn = CZN[mkey] || [], alt = CRE_ALT[mkey] || [];
  names.forEach(function (n, ref) {
    if (!cnt[ref]) return;
    var b = isBoss(mkey, ref);
    rows.push({ name: CKEY + n, label: n, n: cnt[ref],
                sw: b ? 'boss' : 'mob', color: b ? COL.boss : COL.mob,
                tag: b ? 'Boss' : '', icon: creIcon(mkey, ref),
                alt: alt[ref] || null, lvl: (CLVL[mkey] || [])[ref] || null,
                zone: zn[ref] || null });
  });
  /* Lowest level first inside each area, because that is the order a
   * reader hunts in; entries with no level yet sort after every leveled
   * one, alphabetical among themselves, and so do ties. The areas
   * themselves follow the world's own order with Mixed last. */
  function lvlNum(lvl) {
    if (!lvl) return null;
    var n = parseInt(lvl.replace(/[^0-9]/g, ''), 10);
    return isNaN(n) ? null : n;
  }
  rows.sort(function (a, b) {
    var x = lvlNum(a.lvl), y = lvlNum(b.lvl);
    if (x === null && y === null) return a.label.localeCompare(b.label);
    if (x === null) return 1;
    if (y === null) return -1;
    return (x - y) || a.label.localeCompare(b.label);
  });

  /* Bosses are their own group rather than a star against a name: eighteen
   * of them against two hundred ordinary spawns, and they are the reason
   * anyone opens this list rather than the map.  Mixed on either side, so
   * a wide ranging boss is not folded in with the pack it stands over. */
  var zs = zonesOf(mkey);
  makeGroup('Bosses', rows.filter(function (r) { return r.tag; }),
            ['boss', 'bosses'], zs);
  makeGroup('Monsters', rows.filter(function (r) { return !r.tag; }),
            ['creature', 'creatures'], zs);
}
function buildItemGroup() {
  var names = itmNames(mkey), cnt = countsOf(D.items[mkey] || []), rows = [];
  var zn = IZN[mkey] || [], cat = ICAT[mkey] || [], alt = ITM_ALT[mkey] || [];
  /* Counted by where a point falls rather than by how many rows carry it:
   * the world buffs each carry the whole pool of posts, and nine rows
   * over the same ground would otherwise claim nine times the pins. */
  var place = Object.create(null), seen = Object.create(null), total = 0;
  (D.items[mkey] || []).forEach(function (p) {
    var c = cat[p[2]] || 'quest', k = c + '|' + p[0] + ',' + p[1];
    if (seen[k]) return;
    seen[k] = 1;
    place[c] = (place[c] || 0) + 1;
    total++;
  });
  names.forEach(function (n, ref) {
    if (!cnt[ref]) return;
    rows.push({ name: IKEY + n, label: n, n: cnt[ref],
                sw: 'item', color: COL.item, icon: itmIcon(mkey, ref),
                alt: alt[ref] || null,
                zone: zn[ref] || null, cat: cat[ref] || 'quest' });
  });
  rows.sort(function (a, b) { return a.label.localeCompare(b.label); });
  buffRot(mkey);
  /* Omega glyphs first, then Lilly's, then everything else the quests on
   * this map ask for.  The build calls a set by its word; the split only
   * shows when two or more are present, so a world with no glyphs keeps
   * one plain list. */
  return makeGroup('Items', rows, ['kind', 'kinds'], zonesOf(mkey),
                   ['omega glyph', 'lilly', 'quest', 'buff'], place, total);
}
/* One row per treasure map, edged in bronze and carrying its own parchment
 * as the picture.  On Cardia the surface maps split by the area their dig
 * spot stands in - the same areas the rest of the site uses, so a map is
 * found the way an ore is - and the cave's four are a plain list, that
 * world having no named ground of its own.  A set heading (Surface / Cave /
 * Omega) only appears when two kinds share one world, which they do not, so
 * each world reads as the one list it is. */
function buildTreasureGroup() {
  var names = treNames(mkey), cnt = countsOf(D.treasures[mkey] || []), rows = [];
  var zn = TZN[mkey] || [], cat = TCAT[mkey] || [];
  names.forEach(function (n, ref) {
    if (!cnt[ref]) return;
    rows.push({ name: TKEY + n, label: n, n: cnt[ref],
                sw: 'treasure', color: RIMCOL.treasure, icon: treIcon(mkey, ref),
                zone: zn[ref] || null, cat: cat[ref] || 'surface',
                text: treText(mkey, ref) });
  });
  rows.sort(function (a, b) { return a.label.localeCompare(b.label); });
  return makeGroup('Treasure Maps', rows, ['map', 'maps'], zonesOf(mkey),
                   ['surface', 'cave', 'omega', 'cove']);
}
/* One row per statue, edged in the blue that is only theirs.  On Cardia a
 * statue is listed under the area it stands in - the same areas the rest of
 * the site uses, so a monument is found the way an ore is - and it never
 * lands in Mixed, because the build votes among the named grounds only.  A
 * world with no named ground keeps them a plain list, the way the cave's
 * four treasures read. */
function buildStatueGroup() {
  var keys = statueNames(mkey), cnt = countsOf(D.statues[mkey] || []), rows = [];
  var zn = SZN[mkey] || [];
  keys.forEach(function (n, ref) {
    if (!cnt[ref]) return;
    rows.push({ name: SKEY + n, label: statueLabel(mkey, ref), n: cnt[ref],
                sw: 'statue', color: RIMCOL.statue, icon: statueIcon(mkey),
                zone: zn[ref] || null, cat: null });
  });
  rows.sort(function (a, b) {
    return (a.zone || '').localeCompare(b.zone || '') ||
           a.label.localeCompare(b.label);
  });
  return makeGroup('Statues', rows, ['statue', 'statues'], zonesOf(mkey), []);
}
/* One row per NPC kind, edged in the green that is only theirs.  On Cardia
 * an NPC is listed under the area its posts stand in - a kind whose posts
 * disagree reads Mixed, the items' rule - and a world with no named
 * ground keeps them a plain list, the way the cave's statues read. */
function buildNpcGroup() {
  var keys = npcNames(mkey), cnt = countsOf(D.npcs[mkey] || []), rows = [];
  var zn = NZN[mkey] || [];
  keys.forEach(function (n, ref) {
    if (!cnt[ref]) return;
    rows.push({ name: NKEY + n, label: n, n: cnt[ref],
                sw: 'npc', color: RIMCOL.npc,
                zone: zn[ref] || null, cat: null });
  });
  rows.sort(function (a, b) {
    return (a.zone || '').localeCompare(b.zone || '') ||
           a.label.localeCompare(b.label);
  });
  return makeGroup('NPCs', rows, ['npc', 'npcs'], zonesOf(mkey), []);
}
/* Groups in reading order: what the site is for first, then what stands
 * on it.  The creatures used to be two rows no matter how many kinds, so
 * putting them above the ore cost nothing; at one row per creature they
 * would put sixty-five names between a harvester and Iron. */
function buildGroups() {
  elGroups.textContent = '';
  MATCH = null;                     /* the list it answers to has changed */
  /* The recommended ore sets belong to Cardia and Cardia Cave, so the
   * section stands only on those two worlds. */
  var elMining = document.getElementById('mining');
  if (elMining) elMining.hidden = !(mkey === 'cardia' || mkey === 'cave');
  var list = speciesOfThisMap();
  /* Ore and plant are read the way everything below them is: sorted
   * inside, and split by the area their ground belongs to when the world
   * has named ground to split by.  The build calls a species Mixed when
   * two of its grounds stand further apart than half a kilometre, and
   * otherwise gives it the area holding most of its ground.  A creature's
   * stands must all agree before one area names it; a species need only
   * have grounds that reach each other. */
  [['ore', 'Ore'], ['plant', 'Plant']].forEach(function (g) {
    var kind = g[0], rows = list.filter(function (s) { return s.kind === kind; });
    if (!rows.length) return;
    var mine = rows.map(function (s) {
      /* The picture is the thing itself where the game has one, which is
       * the same rule the items below follow: an ore's row in the item
       * table names a picture of the ore.  Without one the row keeps its
       * colour mark, which is the shape the map draws it as anyway. */
      return { name: s.name, label: s.name, n: s.n, sw: s.kind,
               color: s.color, zone: s.zone || null,
               icon: SPIC[SPECIES.indexOf(s)] || '' };
    });
    mine.sort(function (a, b) { return a.label.localeCompare(b.label); });
    makeGroup(g[1], mine, ['species', 'species'], zonesOf(mkey));
  });
  buildCreatureGroup();
  buildItemGroup();
  buildTreasureGroup();
  buildStatueGroup();
  buildNpcGroup();
  if (!elGroups.children.length) {
    var e = document.createElement('div');
    e.className = 'empty';
    e.textContent = 'No species have been recorded here yet.';
    elGroups.appendChild(e);
    renderX();
    return;
  }
  syncRows();
  renderX();
}
/* All and None are about everything the sidebar is offering on this world,
 * not about the top half of it: a button that says None and leaves two
 * hundred monsters standing has lied.  With a search in the box they are
 * about everything that search left showing, for the same reason the
 * switches over a single list are - a search for Snowland and then None
 * is asked for Snowland, not for the world. */
function allRows() {
  var out = [];
  [].forEach.call(elGroups.children, function (b) {
    if (!b._rows) return;
    var hit = hitMap(b);
    b._rows.forEach(function (s) { if (hit[s.name]) out.push(s); });
  });
  return out;
}
/* Every switch across every world that the current search holds up.  The
 * per-world All / None above read it off this world's visible rows; the
 * global buttons need the same answer for worlds whose lists are not built,
 * so the match is asked data-first instead of row-first: group, area, set,
 * both languages, tier, riddle text and the world's own name all count,
 * the same question the list and the Other-worlds hits ask - including
 * the riddle gate: a term naming a whole layer leaves riddle text out.
 * Shared names are one switch, so a key matching on any world is listed once. */
function globalMatchingKeys() {
  var terms = qterms();
  if (!terms.length) return Object.keys(on);
  var allowRiddle = riddleOn(terms);
  var seen = Object.create(null), out = [];
  function add(key) {
    if (key === undefined || on[key] === undefined || seen[key]) return;
    seen[key] = 1; out.push(key);
  }
  if (isDaily()) {
    Object.keys(DAILY).forEach(function (k) {
      var allow = DAILY[k];
      if (allow.indexOf('Monsters') >= 0 || allow.indexOf('Bosses') >= 0) {
        var dcnt = countsOf(D.mobs[k] || []);
        (CRE[k] || []).forEach(function (name, ref) {
          if (!dcnt[ref]) return;
          if (allow.indexOf(isBoss(k, ref) ? 'Bosses' : 'Monsters') >= 0)
            add(CKEY + name);
        });
      }
      if (allow.indexOf('Items') >= 0) {
        var dicnt = countsOf(D.items[k] || []);
        (ITM[k] || []).forEach(function (name, ref) {
          if (dicnt[ref]) add(IKEY + name);
        });
      }
    });
    return out;
  }
  function add(key) {
    if (key === undefined || on[key] === undefined || seen[key]) return;
    seen[key] = 1; out.push(key);
  }
  function hit(group, zone, cat, label, alt, lvl, text, unique) {
    var g = (group || '').toLowerCase(),
        lb = (label || '').toLowerCase(),
        uq = unique ? unique.toLowerCase() : null,
        al = alt ? String(alt).toLowerCase() : null,
        tx = text ? String(text).toLowerCase() : null,
        c = cat ? String(cat).toLowerCase() : null,
        mapped = c ? (CATW[c] || CATW[String(cat)]) : null,
        mp = mapped ? String(mapped).toLowerCase() : null;
    return terms.some(function (t) {
      if (g.indexOf(t) >= 0) return true;
      if (zone && zone.toLowerCase().indexOf(t) >= 0) return true;
      if (c && (c.indexOf(t) >= 0 || (mp && mp.indexOf(t) >= 0))) return true;
      if (lb.indexOf(t) >= 0) return true;
      if (uq && uq.indexOf(t) >= 0) return true;
      if (al && al.indexOf(t) >= 0) return true;
      if (allowRiddle && tx && tx.indexOf(t) >= 0) return true;
      if (lvl && lvlHit({ lvl: lvl }, t)) return true;
      return false;
    });
  }
  MAPS.forEach(function (m) {
    var k = m.key;
    var mapHit = terms.some(function (t) {
      return (m.title || k).toLowerCase().indexOf(t) >= 0;
    });
    SPECIES.forEach(function (s) {
      if (s.map !== k) return;
      if (mapHit || hit(s.kind === 'ore' ? 'Ore' : 'Plant',
                        s.zone, null, s.name, null, null, null, null))
        add(s.name);
    });
    var cnt = countsOf(D.mobs[k] || []),
        zn = CZN[k] || [], al = CRE_ALT[k] || [], lv = CLVL[k] || [];
    (CRE[k] || []).forEach(function (name, ref) {
      if (!cnt[ref]) return;
      if (mapHit || hit(isBoss(k, ref) ? 'Bosses' : 'Monsters',
                        zn[ref], null, name, al[ref], lv[ref], null, null))
        add(CKEY + name);
    });
    var icnt = countsOf(D.items[k] || []),
        izn = IZN[k] || [], icat = ICAT[k] || [], ial = ITM_ALT[k] || [];
    (ITM[k] || []).forEach(function (name, ref) {
      if (!icnt[ref]) return;
      if (mapHit || hit('Items', izn[ref], icat[ref] || 'quest',
                        name, ial[ref], null, null, null))
        add(IKEY + name);
    });
    var tcnt = countsOf(D.treasures[k] || []),
        tzn = TZN[k] || [], tcat = TCAT[k] || [];
    (TTM[k] || []).forEach(function (name, ref) {
      if (!tcnt[ref]) return;
      if (mapHit || hit('Treasure Maps', tzn[ref], tcat[ref],
                        name, null, null, treText(k, ref), null))
        add(TKEY + name);
    });
    var scnt = countsOf(D.statues[k] || []), szn = SZN[k] || [];
    (STM[k] || []).forEach(function (uniq, ref) {
      if (!scnt[ref]) return;
      if (mapHit || hit('Statues', szn[ref], null, statueLabel(k, ref),
                        null, null, null, uniq))
        add(SKEY + uniq);
    });
    var ncnt = countsOf(D.npcs[k] || []), nzn = NZN[k] || [];
    (NTM[k] || []).forEach(function (name, ref) {
      if (!ncnt[ref]) return;
      if (mapHit || hit('NPCs', nzn[ref], null, name, null, null, null, null))
        add(NKEY + name);
    });
  });
  return out;
}
function esc(s) {
  return String(s).replace(/[&<>"]/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
}
/* What the search is holding up in one list, answered for every row in it:
 * a row matches on what it is called, on what stands above it, and on what
 * that is called - the group's own name counts too.  Everything that has to
 * know what is on screen asks this one question: the drawing of the list,
 * and every switch in it, so a tick under a search lands on the entries
 * the search has left rather than reaching through it to the ones it took
 * out.  A search holds one or more |-separated terms and a row answers to
 * any of them; one term behaves exactly like the old substring search. */
function qterms() {
  return query.toLowerCase().split('|').map(function (t) { return t.trim(); })
    .filter(function (t) { return t !== ''; });
}
/* Every layer the sidebar can list, by the name its heading shows.  A term
 * naming a whole layer asks for that layer - "statue" means the statues -
 * so the riddle baked into a treasure map stops counting for that term:
 * without this, five Dread Cove riddles mentioning "a statue" would join
 * every statue search.  A term naming no layer ("lanterns") still reads
 * the riddles, which is how such a map stays findable by its text. */
var LAYER_TITLES = ['Ore', 'Plant', 'Bosses', 'Monsters', 'Items',
                    'Treasure Maps', 'Statues', 'NPCs'];
function riddleOn(terms) {
  return !terms.some(function (t) {
    return LAYER_TITLES.some(function (g) {
      return g.toLowerCase().indexOf(t) >= 0;
    });
  });
}
/* The Daily Quest worlds and the layers that count on each: Kebaria's,
 * Cardia's monsters, bosses and items, Aurania's and Kogarashi's monsters
 * and items.  A layer is a sidebar list, so "Monsters" never takes the
 * bosses box with it. */
var DAILY = {
  'wdoldkebaras': ['Monsters', 'Bosses', 'Items'],
  'wdaurania': ['Monsters', 'Items'],
  'cardia': ['Monsters', 'Bosses', 'Items'],
  'kogarashi': ['Monsters', 'Items']
};
function isDaily() { return query.trim().toLowerCase() === 'daily quest'; }
function hitMap(box) {
  if (isDaily()) {
    var allow = (DAILY[mkey] || []).indexOf(box._title) >= 0;
    var dm = Object.create(null);
    box._rows.forEach(function (s) { dm[s.name] = allow; });
    return dm;
  }
  var terms = qterms(), map = Object.create(null);
  if (!terms.length) {
    box._rows.forEach(function (s) { map[s.name] = true; });
    return map;
  }
  var titleHit = terms.some(function (t) {
    return (box._title || '').toLowerCase().indexOf(t) >= 0;
  });
  var under = Object.create(null);
  if (!titleHit) (box._buckets || []).forEach(function (b) {
    if (b.title && terms.some(function (t) {
        return b.title.toLowerCase().indexOf(t) >= 0;
      }))
      b.rows.forEach(function (s) { under[s.name] = 1; });
  });
  var allowRiddle = riddleOn(terms);
  box._rows.forEach(function (s) {
    var label = (s.label || s.name).toLowerCase(),
        alt = (s.alt || '').toLowerCase(),
        text = (s.text || '').toLowerCase();
    map[s.name] = titleHit || !!under[s.name] ||
                  terms.some(function (t) {
                    return label.indexOf(t) >= 0 || alt.indexOf(t) >= 0 ||
                           lvlHit(s, t) || (allowRiddle && text.indexOf(t) >= 0);
                  });   /* the riddle
                           * counts as the map's name too: "lanterns" is a way of
                           * asking for Treasure Map 01, and only for it - unless
                           * the search names a whole layer, when the statues
                           * are meant and the riddles stay out */
  });
  return map;
}
/* A Platin or Royal tier answers to its own name as well as its owner's:
 * "p16", "p 16" and "<P 16>" all find what wears <P 16>, and the same for
 * royal - while a bare number never worked, and still does not, so the
 * first chef and the second class keep answering for themselves. */
function lvlHit(s, q) {
  if (!s.lvl) return false;
  var terms = String(q).toLowerCase().split('|');
  for (var i = 0; i < terms.length; i++) {
    var nq = terms[i].replace(/[<>\.\s]/g, '');
    if (nq !== '' && nq === String(s.lvl).toLowerCase()) return true;
  }
  return false;
}
/* What the map is showing.  With nothing in the search box it is whatever
 * is switched on, as it has always been; with a search in the box it is
 * only what the search found, so asking for Midorite and following it to
 * another world leaves Midorite standing there and hides the two hundred
 * things that are not it - the list already does that, and a map still
 * covered in everything else is a map that has answered a question the
 * sidebar then contradicts.  The answer is kept rather than asked box by
 * box because a draw asks it once per creature, and it is thrown away the
 * moment the search or the list under it changes. */
function buildMatch() {
  var m = Object.create(null);
  [].forEach.call(elGroups.children, function (b) {
    if (!b._rows) return;
    var hit = hitMap(b);
    b._rows.forEach(function (s) { m[s.name] = hit[s.name]; });
  });
  return m;
}
function shown(name) {
  if (!on[name]) return false;
  if (!query) return true;
  if (!MATCH) MATCH = buildMatch();
  return !!MATCH[name];
}
/* The switches over a list, and the two over all of them, are read from
 * what the search left standing - so a tick under a search reports itself
 * over the entries still on screen rather than over the hundred that are
 * not, and going back to an empty box puts every count right again. */
function syncRows(unfold) {
  filterGen++;
  [].forEach.call(elGroups.children, function (box) {
    if (!box._rows) return;
    var hit = hitMap(box);
    var mine = box._rows.filter(function (s) { return hit[s.name]; });
    var all = mine.every(function (s) { return on[s.name]; });
    var some = mine.some(function (s) { return on[s.name]; });
    box._cb.checked = all;
    box._cb.indeterminate = !all && some;
    box._rows.forEach(function (s) {
      var r = box.querySelector('.row[data-name="' + cssq(s.name) + '"] input');
      if (r) r.checked = !!on[s.name];
    });
    /* an area heading is the same tri-state over a shorter list */
    (box._buckets || []).forEach(function (b) {
      var bm = b.rows.filter(function (s) { return hit[s.name]; });
      var ba = bm.every(function (s) { return on[s.name]; });
      var bs = bm.some(function (s) { return on[s.name]; });
      b.cb.checked = ba;
      b.cb.indeterminate = !ba && bs;
    });
    applyQuery(box, hit, unfold);
  });
}
function cssq(s) {
  return String(s).replace(/["\\]/g, '\\$&');
}
function applyQuery(box, hit, unfold) {
  var q = query.toLowerCase(), any = false;
  box._rows.forEach(function (s) {
    var h = !!hit[s.name];
    if (h) any = true;
    var r = box.querySelector('.row[data-name="' + cssq(s.name) + '"]');
    if (r) r.classList.toggle('hide', !h);
  });
  /* An area with nothing left showing under it would be a heading over an
   * empty list, so it goes with its rows.  What it does with the fold is
   * the whole of the rest: a search opens what it found, the once - a fold
   * from before must not keep a match out of sight - and then leaves the
   * fold alone.  A heading closed while the search is up therefore stays
   * closed through every tick afterwards instead of springing open again,
   * and the search closing puts back the folds it took out of the way. */
  (box._buckets || []).forEach(function (b) {
    var live = b.rows.some(function (s) {
      var r = box.querySelector('.row[data-name="' + cssq(s.name) + '"]');
      return r && !r.classList.contains('hide');
    });
    b.el.classList.toggle('hide', !live);
    if (unfold && q && live && HOPEN[b.key] === false &&
        HUNF.indexOf(b.key) < 0) {
      HOPEN[b.key] = true; HUNF.push(b.key);
    }
    setFold(b.el, HOPEN[b.key] !== false);
  });
  /* The list itself is folded the same way: opened once for the search,
   * folded by hand afterwards and it stays folded. */
  if (unfold && q && any && GOPEN[box._title] === false &&
      GUNF.indexOf(box._title) < 0) {
    GOPEN[box._title] = true; GUNF.push(box._title);
  }
  box.dataset.open = GOPEN[box._title] !== false ? 'true' : 'false';
  box.style.display = any ? '' : 'none';
}

/* The list only ever holds one world, so a search stopping at its edge
 * would answer "nothing" about a chest that is simply standing on another.
 * The other two are asked the same question the list asks - a name, an
 * area, a set, or the name of the group itself - and whatever comes back
 * is offered above the list, one heading per world.  Clicking a match
 * leaves for that world with the search still in the box, so the row that
 * was wanted is the row that ends up on screen. */
var XIDX = null;
var CATW = { 'omega glyph': 'Omega Glyph', lilly: 'Lilly', quest: 'Quest' };
function xindex() {
  if (XIDX) return XIDX;
  XIDX = Object.create(null);
  MAPS.forEach(function (m) {
    var k = m.key, rows = [];
    SPECIES.forEach(function (s) {
      if (s.map !== k) return;
      rows.push({ label: s.name, zone: s.zone || null, cat: null,
                  grp: s.kind === 'ore' ? 'Ore' : 'Plant' });
    });
    var cnt = countsOf(D.mobs[k] || []), zn = CZN[k] || [], lv = CLVL[k] || [];
    (CRE[k] || []).forEach(function (n, ref) {
      if (!cnt[ref]) return;
      rows.push({ label: n, zone: zn[ref] || null, cat: null,
                  lvl: lv[ref] || null,
                  grp: isBoss(k, ref) ? 'Bosses' : 'Monsters' });
    });
    var icnt = countsOf(D.items[k] || []), izn = IZN[k] || [], icat = ICAT[k] || [];
    (ITM[k] || []).forEach(function (n, ref) {
      if (!icnt[ref]) return;
      rows.push({ label: n, zone: izn[ref] || null, cat: icat[ref] || 'quest',
                  grp: 'Items' });
    });
    var tcnt = countsOf(D.treasures[k] || []), tzn = TZN[k] || [];
    (TTM[k] || []).forEach(function (n, ref) {
      if (!tcnt[ref]) return;
      rows.push({ label: n, zone: tzn[ref] || null, cat: null,
                  grp: 'Treasure Maps', text: treText(k, ref) });
    });
    var scnt = countsOf(D.statues[k] || []), szn = SZN[k] || [];
    (STM[k] || []).forEach(function (n, ref) {
      if (!scnt[ref]) return;
      rows.push({ label: statueLabel(k, ref), zone: szn[ref] || null, cat: null,
                  grp: 'Statues' });
    });
    var ncnt = countsOf(D.npcs[k] || []), nzn = NZN[k] || [];
    (NTM[k] || []).forEach(function (n, ref) {
      if (!ncnt[ref]) return;
      rows.push({ label: n, zone: nzn[ref] || null, cat: null,
                  grp: 'NPCs' });
    });
    XIDX[k] = { title: m.title || k, rows: rows };
  });
  return XIDX;
}
function xhit(r, q, mapHit, mk) {
  if (isDaily()) return ((DAILY[mk] || []).indexOf(r.grp) >= 0);
  if (mapHit) return true;
  var terms = String(q).split('|').map(function (t) { return t.trim(); })
    .filter(function (t) { return t !== ''; });
  if (!terms.length) return true;
  var allowRiddle = riddleOn(terms);
  return terms.some(function (t) {
    return r.label.toLowerCase().indexOf(t) >= 0 ||
           lvlHit(r, t) ||
           (r.zone && r.zone.toLowerCase().indexOf(t) >= 0) ||
           (r.cat && (CATW[r.cat] || r.cat).toLowerCase().indexOf(t) >= 0) ||
           r.grp.toLowerCase().indexOf(t) >= 0 ||
           (allowRiddle && (r.text || '').toLowerCase().indexOf(t) >= 0);
           /* riddle text makes a map findable on other worlds too, not just
            * by its label - unless the search names a whole layer */
  });
}
function renderX() {
  var old = document.getElementById('xmap');
  if (old) old.parentNode.removeChild(old);
  var q = query.toLowerCase();
  if (!q) return;
  var idx = xindex(), byMap = [];
  MAPS.forEach(function (m) {
    if (m.key === mkey) return;
    var e = idx[m.key];
    if (!e) return;
    var mapHit = qterms().some(function (t) {
      return (m.title || m.key).toLowerCase().indexOf(t) >= 0;
    });
    var mine = e.rows.filter(function (r) { return xhit(r, q, mapHit, m.key); });
    if (mine.length) byMap.push({ m: m, mine: mine });
  });
  var total = byMap.reduce(function (a, g) { return a + g.mine.length; }, 0);
  if (!total) return;
  var box = document.createElement('div');
  box.className = 'grp xmap'; box.id = 'xmap';
  box._title = 'Other worlds';
  box._buckets = [];
  box.dataset.open = GOPEN['Other worlds'] === false ? 'false' : 'true';
  var h = document.createElement('h3');
  h.innerHTML = '<span class="arrow">&#9662;</span>' +
                '<span>Other worlds</span>' +
                '<span class="n">' + total +
                (total === 1 ? ' match' : ' matches') + '</span>';
  h.addEventListener('click', function () {
    box.dataset.open = box.dataset.open === 'true' ? 'false' : 'true';
    GOPEN['Other worlds'] = box.dataset.open === 'true';
    readerFold(GUNF, 'Other worlds', GOPEN['Other worlds']);
  });
  box.appendChild(h);
  var wrap = document.createElement('div');
  wrap.className = 'rows';
  var LIMIT = 8;
  byMap.forEach(function (g) {
    var zh = document.createElement('div');
    zh.className = 'zone xz';
    zh.innerHTML = '<span class="arrow">&#9662;</span>' +
                   '<span>' + esc(g.m.title || g.m.key) + '</span>' +
                   '<span class="n">' + g.mine.length + '</span>';
    wrap.appendChild(zh);
    g.mine.slice(0, LIMIT).forEach(function (r) {
      var el = document.createElement('div');
      el.className = 'row xrow';
      el.title = (g.m.title || g.m.key) + ' \u00b7 ' + r.grp + ' \u00b7 ' + r.label +
                 (r.zone ? ' \u00b7 ' + r.zone : '');
      el.innerHTML = '<span class="xg">' + esc(r.grp) + '</span>' +
                     '<span class="nm">' + esc(r.label) + '</span>' +
                     (r.zone ? '<span class="xn" title="' + esc(r.zone) + '">' +
                               esc(r.zone) + '</span>' : '');
      el.addEventListener('click', function () {
        /* A jump to a switched-off hit switches it back on first, or the
         * new world would frame its pins without drawing them. */
        xrowKeys(r, g.m.key).forEach(function (k) { on[k] = true; });
        setMap(g.m.key, true);
        var pts = xrowPoints(r, g.m.key, r.zone);
        if (!pts.length && r.zone != null) pts = xrowPoints(r, g.m.key, null);
        if (pts.length) fitPoints(pts);
      });
      wrap.appendChild(el);
    });
    if (g.mine.length > LIMIT) {
      var more = document.createElement('div');
      more.className = 'more';
      more.textContent = '\u2026 and ' + (g.mine.length - LIMIT) +
                         ' more on this world';
      wrap.appendChild(more);
    }
    foldHead(zh, 'xmap:' + g.m.key, ownsOf(zh, 1));
    box._buckets.push({ el: zh, key: 'xmap:' + g.m.key });
  });
  box.appendChild(wrap);
  elGroups.insertBefore(box, elGroups.firstChild);
}

function setMap(k, keepView) {
  if (!MAP[k] || (k === mkey && ready)) return;
  mkey = k;
  ready = false;
  preloadIcons();           /* in parallel with the map's own picture */
  /* A picture that finishes loading after a newer one has been asked for
   * belongs to a view nobody is looking at any more.  Letting it through
   * would set ready while this map's picture is still blank, the next
   * draw would blow up on a half loaded image, and the whole window would
   * freeze showing the previous world with the new world's name on it. */
  var tok = ++loadTok;
  var el = document.getElementById('loading');
  /* Switching back to a world seen before reuses its picture from memory:
   * no second download, no second decode, no flash of the loader. */
  var cached = GRID[k];
  if (cached && cached.complete && cached.naturalWidth > 0) {
    img = cached;
    ready = true;
    picMetrics();
    el.hidden = true;
    if (!keepView) fit();
    clampView();
    if (!inGroup(k)) setGroup(groupOf(k));
    else { syncTabs(); }
    buildGroups();
    sel = null; elDetail.hidden = true;
    save(); draw(); writeHash();
    return;
  }
  el.hidden = false;
  el.textContent = 'Loading mapâ€¦';
  var pic = new Image();
  if ('decoding' in pic) pic.decoding = 'async';
  img = pic;
  pic.onload = function () {
    if (tok !== loadTok) return;      /* superseded - the newer switch won */
    GRID[k] = pic;
    ready = true;
    picMetrics();
    el.hidden = true;
    if (!keepView) fit();
    clampView(); draw(); writeHash();
  };
  pic.onerror = function () {
    if (tok !== loadTok) return;
    el.textContent = MAP[k].img + ' is missing.';
  };
  pic.src = asset(MAP[k].img);
  /* A search jump can land on a world outside the tab filter; the tabs
   * follow the map rather than hiding where it went. */
  if (!inGroup(k)) setGroup(groupOf(k));
  else { syncTabs(); }
  buildGroups();
  sel = null; elDetail.hidden = true;
  save();
}

/* ---------------------------------------------------------- detail card */
function showDetail(sp) {
  var s = SPECIES[sp[2]], w = toWorld(sp[0], sp[1]);
  sel = { si: sp[2], ix: sp[0], iy: sp[1] };
  elDetail.hidden = false;
  elDetail.classList.remove('wide');
  elDetail.innerHTML =
    '<div class="bar" style="background:' + s.color + '"></div>' +
    '<button class="x" title="Close">&times;</button>' +
    '<div class="in">' +
      '<h4><span class="swatch ' + s.kind + '" style="display:inline-block;width:11px;height:11px;' +
        (s.kind === 'plant' ? 'transform:rotate(45deg) scale(.86);' : 'border-radius:50%;') +
        'background:' + s.color + '"></span>' + esc(s.name) + '</h4>' +
      '<div class="kind">' + s.kind + ' &middot; ' + esc(MAP[mkey].title) + '</div>' +
      '<dl>' +
        '<dt>World x</dt><dd>' + w[0].toFixed(1) + '</dd>' +
        '<dt>World z</dt><dd>' + w[1].toFixed(1) + '</dd>' +
        '<dt>Known at</dt><dd>' + s.n + ' spot' + (s.n === 1 ? '' : 's') + ' here</dd>' +
      '</dl>' +
      '<div class="acts">' +
        '<button data-a="solo">Show only this</button>' +
        '<button data-a="centre">Centre</button>' +
      '</div>' +
    '</div>';
  elDetail.querySelector('.x').addEventListener('click', closeDetail);
  elDetail.querySelector('[data-a=solo]').addEventListener('click', function () {
    for (var n in on) on[n] = (n === s.name);
    syncRows(); save(); draw(); writeHash();
  });
  elDetail.querySelector('[data-a=centre]').addEventListener('click', function () {
    vx = sp[0]; vy = sp[1]; draw(); writeHash();
  });
  draw();
}
function closeDetail() { sel = null; elDetail.hidden = true; elDetail.classList.remove('wide'); draw(); }

/* Clicking a dig spot opens the map itself rather than a line of numbers:
 * the riddle is baked into the parchment, so the only way to read it is to
 * show the picture large.  The ring still marks the spot, and the world
 * coordinates are kept for anyone pacing it out. */
function nearestTreasure(px, py) {
  var arr = D.treasures[mkey] || [], names = treNames(mkey),
      best = null, bd = 13 * 13;
  for (var i = 0; i < arr.length; i++) {
    var nm = names[arr[i][2]];
    if (nm === undefined || !shown(TKEY + nm)) continue;
    var dx = sx(arr[i][0]) - px, dy = sy(arr[i][1]) - py, d = dx * dx + dy * dy;
    if (d <= bd) { bd = d; best = arr[i]; }
  }
  return best;
}
function showTreasureDetail(pt) {
  var ref = pt[2], nm = treNames(mkey)[ref], ic = treFull(mkey, ref);
  var w = toWorld(pt[0], pt[1]);
  sel = { si: -1, ix: pt[0], iy: pt[1] };
  elDetail.hidden = false;
  elDetail.classList.add('wide');
  elDetail.innerHTML =
    '<div class="bar" style="background:' + RIMCOL.treasure + '"></div>' +
    '<button class="x" title="Close">&times;</button>' +
    '<div class="in">' +
      '<h4>' + esc(nm) + '</h4>' +
      '<div class="kind">treasure dig spot &middot; ' + esc(MAP[mkey].title) + '</div>' +
      (ic ? '<img class="tmap" src="' + asset(ic) + '" alt="">' : '') +
      '<dl>' +
        '<dt>World x</dt><dd>' + w[0].toFixed(1) + '</dd>' +
        '<dt>World z</dt><dd>' + w[1].toFixed(1) + '</dd>' +
      '</dl>' +
      '<div class="acts">' +
        '<button data-a="centre">Centre</button>' +
      '</div>' +
    '</div>';
  elDetail.querySelector('.x').addEventListener('click', closeDetail);
  elDetail.querySelector('[data-a=centre]').addEventListener('click', function () {
    vx = pt[0]; vy = pt[1]; draw(); writeHash();
  });
  draw();
}

/* Clicking an NPC shows who it is and where it stands.  A person has no
 * parchment to enlarge, so the card is the name, the area, and the world
 * coordinates for walking to them. */
function nearestNpc(px, py) {
  var arr = D.npcs[mkey] || [], names = npcNames(mkey),
      best = null, bd = 12 * 12;
  for (var i = 0; i < arr.length; i++) {
    var nm = names[arr[i][2]];
    if (nm === undefined || !shown(NKEY + nm)) continue;
    var dx = sx(arr[i][0]) - px, dy = sy(arr[i][1]) - py, d = dx * dx + dy * dy;
    if (d <= bd) { bd = d; best = arr[i]; }
  }
  return best;
}
function showNpcDetail(pt) {
  var ref = pt[2], zone = (NZN[mkey] || [])[ref];
  var w = toWorld(pt[0], pt[1]);
  sel = { si: -1, ix: pt[0], iy: pt[1] };
  elDetail.hidden = false;
  elDetail.classList.remove('wide');
  elDetail.innerHTML =
    '<div class="bar" style="background:' + RIMCOL.npc + '"></div>' +
    '<button class="x" title="Close">&times;</button>' +
    '<div class="in">' +
      '<h4><span style="display:inline-block;width:11px;height:11px;' +
        'transform:rotate(45deg) scale(.86);background:' + RIMCOL.npc + ';' +
        'margin-right:7px"></span>' +
        esc(npcNames(mkey)[ref]) + '</h4>' +
      '<div class="kind">npc' + (zone ? ' &middot; ' + esc(zone) : '') +
        ' &middot; ' + esc(MAP[mkey].title) + '</div>' +
      '<dl>' +
        '<dt>World x</dt><dd>' + w[0].toFixed(1) + '</dd>' +
        '<dt>World z</dt><dd>' + w[1].toFixed(1) + '</dd>' +
      '</dl>' +
      '<div class="acts">' +
        '<button data-a="centre">Centre</button>' +
      '</div>' +
    '</div>';
  elDetail.querySelector('.x').addEventListener('click', closeDetail);
  elDetail.querySelector('[data-a=centre]').addEventListener('click', function () {
    vx = pt[0]; vy = pt[1]; draw(); writeHash();
  });
  draw();
}
/* Clicking a statue shows what it is and where it stands.  A monument has
 * no parchment to enlarge, so the card is the name it is filed under, the
 * area it stands in, and the world coordinates for pacing it out. */
function nearestStatue(px, py) {
  var arr = D.statues[mkey] || [], names = statueNames(mkey),
      best = null, bd = 12 * 12;
  for (var i = 0; i < arr.length; i++) {
    var nm = names[arr[i][2]];
    if (nm === undefined || !shown(SKEY + nm)) continue;
    var dx = sx(arr[i][0]) - px, dy = sy(arr[i][1]) - py, d = dx * dx + dy * dy;
    if (d <= bd) { bd = d; best = arr[i]; }
  }
  return best;
}
function showStatueDetail(pt) {
  var ref = pt[2], zone = (SZN[mkey] || [])[ref];
  var w = toWorld(pt[0], pt[1]);
  sel = { si: -1, ix: pt[0], iy: pt[1] };
  elDetail.hidden = false;
  elDetail.classList.remove('wide');
  elDetail.innerHTML =
    '<div class="bar" style="background:' + RIMCOL.statue + '"></div>' +
    '<button class="x" title="Close">&times;</button>' +
    '<div class="in">' +
      '<h4><img src="' + asset(statueIcon(mkey)) + '" style="height:26px;vertical-align:-7px;' +
        'margin-right:7px" alt="">' +
        esc(statueLabel(mkey, ref)) + '</h4>' +
      '<div class="kind">statue' + (zone ? ' &middot; ' + esc(zone) : '') +
        ' &middot; ' + esc(MAP[mkey].title) + '</div>' +
      '<dl>' +
        '<dt>World x</dt><dd>' + w[0].toFixed(1) + '</dd>' +
        '<dt>World z</dt><dd>' + w[1].toFixed(1) + '</dd>' +
      '</dl>' +
      '<div class="acts">' +
        '<button data-a="centre">Centre</button>' +
      '</div>' +
    '</div>';
  elDetail.querySelector('.x').addEventListener('click', closeDetail);
  elDetail.querySelector('[data-a=centre]').addEventListener('click', function () {
    vx = pt[0]; vy = pt[1]; draw(); writeHash();
  });
  draw();
}

/* -------------------------------------------------------------- pointer */
var dragging = false, dragMoved = false, lastX = 0, lastY = 0;
/* Hover walks every spawn on the map, so a mousemove storm must not run
 * it directly: one queued hover per frame at most. */
var hoverQueued = false;
function queueHover() {
  if (hoverQueued) return;
  hoverQueued = true;
  requestAnimationFrame(function () { hoverQueued = false; updateHover(); });
}

cv.addEventListener('pointerdown', function (e) {
  cv.setPointerCapture(e.pointerId);
  dragging = true; dragMoved = false;
  lastX = e.offsetX; lastY = e.offsetY;
  cv.classList.add('drag');
});
cv.addEventListener('pointermove', function (e) {
  mx = e.offsetX; my = e.offsetY; touched = true;
  if (dragging) {
    var dx = mx - lastX, dy = my - lastY;
    if (Math.abs(dx) + Math.abs(dy) > 0) dragMoved = true;
    vx -= dx / sc; vy -= dy / sc;
    lastX = mx; lastY = my;
    clampView(); previewPan();
    elTip.hidden = true;
  } else {
    queueHover();
  }
});
function endDrag(e) {
  if (!dragging) return;
  dragging = false;
  if (settleT) { clearTimeout(settleT); settleT = 0; gesture = false; }
  cv.classList.remove('drag');
  if (!dragMoved) {
    var sp = nearest(mx, my);
    if (sp) { showDetail(sp); }
    else {
      var tr = nearestTreasure(mx, my);
      if (tr) { showTreasureDetail(tr); }
      else {
        var st = nearestStatue(mx, my);
        if (st) showStatueDetail(st);
        else {
          var np = nearestNpc(mx, my);
          if (np) showNpcDetail(np); else closeDetail();
        }
      }
    }
  }
  if (dragMoved) draw();
  writeHash();
}
cv.addEventListener('pointerup', endDrag);
cv.addEventListener('pointercancel', endDrag);
cv.addEventListener('pointerleave', function () {
  elTip.hidden = true; hover = -1;
});

/* Wheel gestures arrive in bursts - a trackpad sends dozens of ticks per
 * flick.  Multiplying them into one factor and zooming once per frame
 * lands at the same scale with one repaint instead of dozens. */
var zoomPendF = 1, zoomPX = 0, zoomPY = 0, zoomQueued = false;
function flushZoom() {
  zoomQueued = false;
  if (zoomPendF !== 1) {
    var f = zoomPendF; zoomPendF = 1;
    zoomMath(zoomPX, zoomPY, f);
    previewZoom(zoomPX, zoomPY);
    armSettle(); queueHash();
  }
}
cv.addEventListener('wheel', function (e) {
  e.preventDefault();
  zoomPendF *= e.deltaY < 0 ? 1.18 : 1 / 1.18;
  zoomPX = e.offsetX; zoomPY = e.offsetY;
  if (!zoomQueued) { zoomQueued = true; requestAnimationFrame(flushZoom); }
}, { passive: false });

/* pinch â€” two pointers, scaled about their midpoint */
var pts = Object.create(null), pinchD = 0;
cv.addEventListener('pointerdown', function (e) { pts[e.pointerId] = e; });
cv.addEventListener('pointermove', function (e) {
  if (!pts[e.pointerId]) return;
  pts[e.pointerId] = e;
  var ids = Object.keys(pts);
  if (ids.length === 2) {
    var a = pts[ids[0]], b = pts[ids[1]];
    var d = Math.hypot(a.clientX - b.clientX, a.clientY - b.clientY);
    var r = cv.getBoundingClientRect();
    if (pinchD > 0 && d > 0) {
      var mpx = (a.clientX + b.clientX) / 2 - r.left,
          mpy = (a.clientY + b.clientY) / 2 - r.top;
      zoomMath(mpx, mpy, d / pinchD);
      previewZoom(mpx, mpy);
      armSettle(); queueHash();
    }
    pinchD = d; dragging = false;
  }
});
['pointerup', 'pointercancel', 'pointerleave'].forEach(function (t) {
  cv.addEventListener(t, function (e) { delete pts[e.pointerId]; pinchD = 0; });
});

/* A creature is a smaller target than a node, and the node is what this
 * site is for, so a spawn only answers the pointer when nothing else is
 * nearer - the same reasoning as the drawing order, kept in step so that
 * a pin drawn under a node also yields to it. */
function nearestSpawn(px, py) {
  if (!showSp) return null;
  var best = null, bd = 9 * 9;
  function look(arr, names, pfx, boss) {
    for (var i = 0; i < arr.length; i++) {
      var nm = names[arr[i][2]];
      /* shown(), not on[]: with a search in the box a switched-on mark the
       * search did not find is not drawn, so the pointer must not find it
       * either - otherwise invisible marks answer the mouse. */
      if (nm === undefined || !shown(pfx + nm)) continue;
      var dx = sx(arr[i][0]) - px, dy = sy(arr[i][1]) - py, d = dx * dx + dy * dy;
      if (d <= bd) {
        bd = d;
        var al = pfx === CKEY ? (CRE_ALT[mkey] || [])[arr[i][2]]
               : pfx === IKEY ? (ITM_ALT[mkey] || [])[arr[i][2]] : null;
        best = { x: arr[i][0], y: arr[i][1], name: nm, ref: arr[i][2],
                 pfx: pfx, boss: !!boss[arr[i][2]], alt: al || null };
      }
    }
  }
  look(D.mobs[mkey] || [], creNames(mkey), CKEY, BOS[mkey] || {});
  look(D.items[mkey] || [], itmNames(mkey), IKEY, {});
  look(D.treasures[mkey] || [], treNames(mkey), TKEY, {});
  look(D.statues[mkey] || [], statueNames(mkey), SKEY, {});
  look(D.npcs[mkey] || [], npcNames(mkey), NKEY, {});
  return best;
}
function spawnTip(s) {
  var w = toWorld(s.x, s.y);
  if (s.pfx === TKEY) {              /* hover shows the whole scroll: the
                                      * riddle is baked into it, so the
                                      * parchment shown whole is the whole
                                      * of the answer */
    var ic = treFull(mkey, s.ref);
    return (ic ? '<img src="' + asset(ic) + '" alt="">' : '') +
           '<b>' + esc(s.name) + '</b>' +
           '<i>' + w[0].toFixed(0) + ', ' + w[1].toFixed(0) + '</i>';
  }
  if (s.pfx === SKEY) {              /* a monument: its own cutout, the same
                                      * picture the row and the pin show */
    return '<img src="' + asset(statueIcon()) + '" style="height:17px;vertical-align:-4px" alt=""> ' +
           esc(s.name) +
           '<i>' + w[0].toFixed(0) + ', ' + w[1].toFixed(0) + '</i>';
  }
  if (s.pfx === NKEY) {              /* a person: the diamond it is drawn
                                      * as, in the green that is only theirs */
    return '<b style="color:' + RIMCOL.npc + '">&#9670;</b> ' + esc(s.name) +
           '<i>' + w[0].toFixed(0) + ', ' + w[1].toFixed(0) + '</i>';
  }
  var item = s.pfx === IKEY;
  var col = item ? COL.item : (s.boss ? COL.boss : '#ff3b30');
  var lvl = (!item && s.pfx === CKEY) ? ((CLVL[mkey] || [])[s.ref] || null) : null;
  return '<b style="color:' + col + '">' + (item ? '&#9670;' : '&#9679;') +
         '</b> ' + esc(s.name) +
         (lvl ? '<span class="lvl">' + lvlText(lvl) + '</span>' : '') +
         (s.alt ? '<i style="text-transform:none"> &middot; ' + esc(s.alt) + '</i>' : '') +
         (s.boss ? '<span class="tag">boss</span>' : '') +
         '<i>' + w[0].toFixed(0) + ', ' + w[1].toFixed(0) + '</i>';
}

function placeTip() {
  var x = mx + 8, y = my;
  if (x + elTip.offsetWidth > W - 8) x = mx - elTip.offsetWidth - 8;
  elTip.style.left = x + 'px';
  elTip.style.top = Math.max(10, Math.min(H - 10, y)) + 'px';
}

function updateHover() {
  var sp = nearest(mx, my);
  var si = sp ? sp[2] : -1;
  if (si !== hover) { hover = si; }
  /* the node wins the pointer, but a spawn still has to be findable when
   * the cursor is in open country, so the cursor follows whichever */
  var ss = sp ? null : nearestSpawn(mx, my);
  cv.classList.toggle('over', !!sp || !!ss);
  if (sp) {
    overSpot = sp;
    var s = SPECIES[sp[2]], w = toWorld(sp[0], sp[1]);
    elTip.hidden = false;
    elTip.className = 'tip';
    elTip.innerHTML = '<b style="color:' + s.color + '">&#9679;</b> ' + esc(s.name) +
                      '<i>' + w[0].toFixed(0) + ', ' + w[1].toFixed(0) + '</i>';
    placeTip();
    return;
  }
  overSpot = null;
  if (!ss) { elTip.hidden = true; return; }
  elTip.hidden = false;
  elTip.className = 'tip' + (ss.pfx === TKEY ? ' big' : '');
  elTip.innerHTML = spawnTip(ss);
  placeTip();
}

/* -------------------------------------------------------------- status */
var stMap = document.getElementById('stMap');
var stSpots = document.getElementById('stSpots');
var stSpawn = document.getElementById('stSpawn');
var stXY = document.getElementById('stXY');
var stScale = document.getElementById('stScale');
function visibleCount(arr, names, pfx) {
  var n = 0;
  for (var i = 0; i < arr.length; i++) {
    var nm = names[arr[i][2]];
    if (nm !== undefined && shown(pfx + nm)) n++;
  }
  return n;
}
/* The same, but counting a place once however many rows carry it.  Every
 * world buff stands at every post, so the raw count would be nine times
 * the pins the map is actually drawing and the bar would be lying about
 * what is on screen. */
function placeCount(arr, names, pfx) {
  var seen = Object.create(null), n = 0;
  for (var i = 0; i < arr.length; i++) {
    var nm = names[arr[i][2]];
    if (nm === undefined || !shown(pfx + nm)) continue;
    var k = arr[i][0] + ',' + arr[i][1];
    if (seen[k]) continue;
    seen[k] = 1; n++;
  }
  return n;
}
var lastSlowT = 0, lastSlowKey = '', cacheMap = '', cacheSpots = '', cacheSpawn = '';
var filterGen = 0; /* bumped on every filter change so the status bar refreshes at once */
function status(force) {
  /* The counts walk every point on the map, so they are not recomputed
   * every frame: the bar updates them at most twice a second, or when a
   * filter forces it.  Scale and cursor position stay per-frame. */
  var slowKey = mkey + '|' + showSp + '|' + filterGen;
  var now = (window.performance && performance.now()) || 0;
  if (force || slowKey !== lastSlowKey || now - lastSlowT > 500) {
    lastSlowT = now; lastSlowKey = slowKey;
    var m = MAP[mkey], n = 0, c = 0;
    var arr = D.spots[mkey] || [];
    for (var i = 0; i < arr.length; i++) if (shown(SPECIES[arr[i][2]].name)) n++;
    speciesOfThisMap().forEach(function (s) { if (shown(s.name)) c++; });
    cacheMap = '<b>' + esc(m.title) + '</b>';
    cacheSpots = n.toLocaleString() + ' spot' + (n === 1 ? '' : 's') +
                 ' Â· ' + c + ' species';
    cacheSpawn =
        visibleCount(D.mobs[mkey] || [], creNames(mkey), CKEY) + ' monsters Â· ' +
        placeCount(D.items[mkey] || [], itmNames(mkey), IKEY) + ' items Â· ' +
        visibleCount(D.treasures[mkey] || [], treNames(mkey), TKEY) + ' maps Â· ' +
        visibleCount(D.statues[mkey] || [], statueNames(mkey), SKEY) + ' statues Â· ' +
        visibleCount(D.npcs[mkey] || [], npcNames(mkey), NKEY) + ' npcs';
    if (stMap.innerHTML !== cacheMap) stMap.innerHTML = cacheMap;
    if (stSpots.textContent !== cacheSpots) stSpots.textContent = cacheSpots;
    if (stSpawn.textContent !== cacheSpawn) stSpawn.textContent = cacheSpawn;
  }
  var mpp = sc * pxPerMetre() * PPX;
  var sScale = mpp.toFixed(2) + ' px/m';
  if (stScale.textContent !== sScale) stScale.textContent = sScale;
  var w = toWorld(ixAt(mx), iyAt(my));
  var sXY = 'x ' + w[0].toFixed(0) + '   z ' + w[1].toFixed(0);
  if (stXY.textContent !== sXY) stXY.textContent = sXY;
}
/* ------------------------------------------------------------ shortcuts */
document.getElementById('zin').addEventListener('click', function () { zoomAt(W / 2, H / 2, 1.4); });
document.getElementById('zout').addEventListener('click', function () { zoomAt(W / 2, H / 2, 1 / 1.4); });
document.getElementById('bfit').addEventListener('click', function () { fit(); draw(); writeHash(); });
document.getElementById('bAll').addEventListener('click', function () {
  allRows().forEach(function (s) { on[s.name] = true; });
  syncRows(); save(); draw(); writeHash();
});
document.getElementById('bNone').addEventListener('click', function () {
  allRows().forEach(function (s) { on[s.name] = false; });
  syncRows(); save(); draw(); writeHash();
});
/* Global select: every switch at once, or - with a search in the box -
 * every switch the search holds up on any world.  All / None above are the
 * same question scoped to this world's visible rows; these two ask it of
 * all worlds, so clearing the box is never needed to bring one find back -
 * or to take one set of finds away. */
document.getElementById('bAllMaps').addEventListener('click', function () {
  globalMatchingKeys().forEach(function (n) { on[n] = true; });
  syncRows(); save(); draw(); writeHash();
});
document.getElementById('bNoneMaps').addEventListener('click', function () {
  globalMatchingKeys().forEach(function (n) { on[n] = false; });
  syncRows(); save(); draw(); writeHash();
});
/* Folding is the reader's view of the list and nothing else: these two
 * change no switch, move nothing on the map, and are remembered for the
 * session by the same state the arrows keep. */
document.getElementById('bFold').addEventListener('click', function () { foldAll(false); });
document.getElementById('bUnfold').addEventListener('click', function () { foldAll(true); });
function setQuery(nv) {
  nv = String(nv == null ? '' : nv).slice(0, 120);
  if (query && !nv) refold();        /* the search closing folds it all
                                      * back the way it was */
  query = nv;
  MATCH = null;                     /* the map draws what the search
                                     * found, and that has just changed */
  syncRows(!!query);                 /* a new search takes the folds out
                                      * of the way of what it found */
  renderX();
  draw();                            /* ...and takes everything else off
                                      * the map with it */
  syncGuides();
  syncMining();
  queueHash();
}
document.getElementById('q').addEventListener('input', function (e) {
  setQuery(e.target.value.trim());
});
var bSp = document.getElementById('tgspawns'), bPl = document.getElementById('tgplaces');
bSp.addEventListener('click', function () {
  showSp = !showSp; bSp.setAttribute('aria-pressed', showSp); draw(); save(); writeHash();
});
bPl.addEventListener('click', function () {
  showPl = !showPl; bPl.setAttribute('aria-pressed', showPl); draw(); save(); writeHash();
});
/* Guide shortcuts in the middle of the header: one press filters this
 * world to the guide's subject and lists every other world's hits above
 * the list, a second press clears it.  Each is just a saved search - the
 * Astral guide is the search for the Astral Crystal Notes and the
 * Celestial Treasure, written as one |-separated search - so every guide
 * stays shareable through the URL like any typed search.  Daily Quest is
 * the one exception: no substring spans its worlds and layers, so the
 * matchers answer "daily quest" from the DAILY table instead. */
var GUIDES = [
  ['tgstatue', 'statue'],
  ['tgore', '(ore)'],
  ['tgplant', '(plant)'],
  ['tgomega', 'omega'],
  ['tgastral', 'astral crystal|celestial map'],
  ['tgtreasure', 'treasure'],
  ['tgdaily', 'daily quest']
];
function syncGuides() {
  GUIDES.forEach(function (g) {
    var b = document.getElementById(g[0]);
    if (b) b.setAttribute('aria-pressed', query.toLowerCase() === g[1]);
  });
}
GUIDES.forEach(function (g) {
  var b = document.getElementById(g[0]);
  if (!b) return;
  b.addEventListener('click', function () {
    var elQ = document.getElementById('q');
    if (query.toLowerCase() === g[1]) {
      if (elQ) elQ.value = '';
      setQuery('');
    } else {
      if (elQ) elQ.value = g[1];
      setQuery(g[1]);
    }
  });
});
/* Recommended ore sets for Cardia and Cardia Cave: each set is one
 * |-separated search over exact "Name (Ore)" labels, so a set finds its
 * four ores and nothing whose name merely contains a fragment.  The label
 * spellings follow the data - Perl, Fairy Swing, Emeralnd. */
var MINING_SETS = [
  'adamantite (ore)|aquamarine (ore)|argentum (ore)|copper (ore)',
  'crusium (ore)|crystalwing (ore)|duckharukon (ore)|dureline (ore)',
  'emeralnd (ore)|fairy swing (ore)|hearthstone (ore)|iron (ore)',
  'lifestone (ore)|manastone (ore)|mithril (ore)|moonlightstone (ore)',
  'obsidian (ore)|opal (ore)|perl (ore)|regency (ore)',
  'retion (ore)|ruby (ore)|runecarore (ore)|sapphire (ore)',
  'adamantite (ore)|aquamarine (ore)|spiritstone (ore)|thunderboltstone (ore)'
];
function syncMining() {
  var bar = document.getElementById('mining');
  if (!bar) return;
  var qs = query.toLowerCase();
  [].forEach.call(bar.querySelectorAll('button[data-mine]'), function (b) {
    b.setAttribute('aria-pressed',
      qs === MINING_SETS[+b.getAttribute('data-mine')]);
  });
}
[].forEach.call(document.querySelectorAll('#mining button[data-mine]'), function (b) {
  b.addEventListener('click', function () {
    var q = MINING_SETS[+b.getAttribute('data-mine')];
    var elQ = document.getElementById('q');
    if (query.toLowerCase() === q) {
      if (elQ) elQ.value = '';
      setQuery('');
    } else {
      if (elQ) elQ.value = q;
      setQuery(q);
    }
  });
});
document.getElementById('burger').addEventListener('click', function () {
  document.getElementById('side').classList.toggle('open');
});
window.addEventListener('keydown', function (e) {
  if (/^(INPUT|TEXTAREA)$/.test(e.target.tagName)) {
    if (e.key === 'Escape') e.target.blur();
    return;
  }
  var k = e.key;
  if (k === 'f' || k === 'F') { fit(); draw(); writeHash(); }
  else if (k === '+' || k === '=') zoomAt(W / 2, H / 2, 1.4);
  else if (k === '-' || k === '_') zoomAt(W / 2, H / 2, 1 / 1.4);
  else if (k === 'Escape') closeDetail();
  else if (k >= '1' && k <= '9') { var m = MAPS[+k - 1]; if (m) setMap(m.key); }
  else if (k === 'ArrowLeft')  { vx -= 90 / sc; clampView(); previewPan(); armSettle(); writeHash(); }
  else if (k === 'ArrowRight') { vx += 90 / sc; clampView(); previewPan(); armSettle(); writeHash(); }
  else if (k === 'ArrowUp')    { vy -= 90 / sc; clampView(); previewPan(); armSettle(); writeHash(); }
  else if (k === 'ArrowDown')  { vy += 90 / sc; clampView(); previewPan(); armSettle(); writeHash(); }
  else return;
  e.preventDefault();
});
window.addEventListener('resize', resize);
window.addEventListener('hashchange', function () {
  hashLock = true;
  var prev = mkey;
  var had = readHash();
  syncGuides();
  syncMining();
  syncTabs(); buildGroups();
  if (mkey !== prev) setMap(mkey, had);
  else if (had) { clampView(); draw(); }
  hashLock = false;
});

/* ------------------------------------------------------------------ go */
/* The order matters and is the whole of the precedence rule: the stored
 * view is a default, and a link is a statement about this particular
 * visit, so the URL is read last and wins. */
load();
var fromHash = readHash();
if (elGroup) {
  elGroup.value = mgrp;
  elGroup.addEventListener('change', function () { setGroup(elGroup.value); });
}
buildTabs();
buildJump();
bSp.setAttribute('aria-pressed', showSp);
bPl.setAttribute('aria-pressed', showPl);
syncGuides();
syncMining();
resize();                       /* the canvas has its size before the
                                 * picture lands, so the first fit it
                                 * applies is against the real window */
setMap(mkey, fromHash);
})();

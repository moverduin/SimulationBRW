// Brandweer Soest — bezetting simulatie
// Alle verwerking gebeurt 100% client-side. Bron: Incidenten.xlsx.

// ---------- Default vehicle classification ----------
// Code -> { label, role, partial }
//   role:    HP-TS | NP-TS | HP-RV | HP-HA | OTHER
//   partial: 'never' | 'always' | 'nature'  (nature => alleen deelalarm bij natuur/heide/bos/duin)
const DEFAULT_VEHICLES = {
  '093334': { label: 'HP-TS (266)',           role: 'HP-TS', partial: 'never'  },
  '093341': { label: 'HP-TS oud (3341)',      role: 'HP-TS', partial: 'nature' },
  '093344': { label: 'HP-TS deelalarm (3344)',role: 'HP-TS', partial: 'always' },
  '093351': { label: 'HP-RV/HW (257)',        role: 'HP-RV', partial: 'never'  },
  '093352': { label: 'HP-AL (autoladder)',    role: 'HP-RV', partial: 'never'  },
  '093386': { label: 'HP-HA (3386)',          role: 'HP-HA', partial: 'never'  },
  '093364': { label: 'HP-HA (3364)',          role: 'HP-HA', partial: 'never'  },
  '093387': { label: 'HP-HA (3387)',          role: 'HP-HA', partial: 'never'  },
  '093431': { label: 'NP-TS (265)',           role: 'NP-TS', partial: 'never'  },
  '093441': { label: 'NP-TS oud (3441)',      role: 'NP-TS', partial: 'nature' },
  // Synthetic 'replacement vehicle' codes. Used wanneer kolom Voertuigen aangeeft
  // dat onze roepnummer (265/266/257/814) is uitgerukt, maar de body geen eigen
  // 0933xx/0934xx code bevat — onze bemanning reed dan op een vervanger.
  'REPL-HP-TS': { label: 'HP-TS (vervanger)', role: 'HP-TS', partial: 'never'  },
  'REPL-NP-TS': { label: 'NP-TS (vervanger)', role: 'NP-TS', partial: 'never'  },
  'REPL-HP-RV': { label: 'HP-RV (vervanger)', role: 'HP-RV', partial: 'never'  },
  'REPL-HP-HA': { label: 'HP-HA (vervanger)', role: 'HP-HA', partial: 'never'  },
};

// Roepnummer → rol. Deze nummers staan in kolom 'Voertuigen' wanneer onze post
// (Soest) is uitgerukt — onafhankelijk van het 0933xx/0934xx-codenummer in de body.
const ROEPNR_TO_ROLE = {
  '265': 'NP-TS',
  '266': 'HP-TS',
  '257': 'HP-RV',
  '814': 'HP-HA',
};
const REPL_CODE_FOR_ROLE = {
  'HP-TS': 'REPL-HP-TS',
  'NP-TS': 'REPL-NP-TS',
  'HP-RV': 'REPL-HP-RV',
  'HP-HA': 'REPL-HP-HA',
};

function parseRoepnummers(s) {
  if (!s) return [];
  const nums = String(s).match(/\d+/g) || [];
  return nums.filter(n => n in ROEPNR_TO_ROLE);
}

const NATURE_FIRE_RE = /\b(natuur|heide|bos|duin)/i;
const POST_OF_ROLE = { 'HP-TS': 'HP', 'HP-RV': 'HP', 'HP-HA': 'HP', 'NP-TS': 'NP' };

// Strip leading prio/pelo info en alle voertuigcodes uit body, normaliseer => address-key voor merge.
function bodyAddressKey(body) {
  let s = String(body || '');
  // verwijder voertuigcodes
  s = s.replace(/\b09\d{4}\b/g, ' ');
  // verwijder "P 1", "P 2", "(Pel. ...)", "(Uitbr.: ...)", "BR xxx", etc.
  s = s.replace(/\bP\s*\d\b/gi, ' ');
  s = s.replace(/\([^)]*\)/g, ' ');
  s = s.replace(/\bBR\b\s*\w+/gi, ' ');
  s = s.replace(/\bNP\b\s*\w*/gi, ' ');
  // collapse + lowercase
  s = s.toLowerCase().replace(/[^a-z0-9 ]+/g, ' ').replace(/\s+/g, ' ').trim();
  return s.slice(0, 80);
}

// Loaded data
let rawIncidents = [];          // unfiltered, all years
let allIncidents = [];          // filtered: current period + prio filter
let vehicleConfig = structuredClone(DEFAULT_VEHICLES);

// Charts
const charts = {};

// ---------- Excel loading ----------

const VEHICLE_CODE_RE = /\b09\d{4}\b/g;
const PRIO_RE = /^P\s*(\d)/i;
const INTREK_RE = /intrekken/i;

async function loadDefaultFile() {
  setStatus('Bestand laden…');
  try {
    // Try data/ first (works on GitHub Pages), then fall back to repo-root.
    let resp = await fetch('data/Incidenten.xlsx');
    if (!resp.ok) resp = await fetch('../Incidenten.xlsx');
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const buf = await resp.arrayBuffer();
    parseWorkbook(buf);
  } catch (err) {
    setStatus('Kon standaardbestand niet laden: ' + err.message + ' — upload het handmatig.');
  }
}

function loadFile(file) {
  setStatus('Bestand inlezen…');
  const reader = new FileReader();
  reader.onload = (e) => parseWorkbook(e.target.result);
  reader.onerror = () => setStatus('Fout bij inlezen.');
  reader.readAsArrayBuffer(file);
}

function parseWorkbook(arrayBuffer) {
  try {
    const wb = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
    if (!wb.SheetNames.includes('Incidenten')) {
      throw new Error("Werkblad 'Incidenten' niet gevonden.");
    }
    const sheet = wb.Sheets['Incidenten'];
    const rows = XLSX.utils.sheet_to_json(sheet, { defval: null, raw: true });

    // Pass 1: parse rows into atomic records (one per spreadsheet row).
    const recs = [];
    for (const r of rows) {
      const body = String(r.body ?? r.Location ?? '');
      if (!body || INTREK_RE.test(body)) continue;
      const codesRaw = body.match(VEHICLE_CODE_RE) || [];
      const seen = new Set();
      const codes = [];
      let haInRow = false;
      for (const c of codesRaw) {
        if (seen.has(c)) continue;
        // Binnen 1 regel telt een HA maar 1× (meerdere HA-codes = zelfde inzet, niet 2 trucks)
        if (vehicleConfig[c]?.role === 'HP-HA') {
          if (haInRow) continue;
          haInRow = true;
        }
        seen.add(c);
        codes.push(c);
      }
      const soest = codes.filter(c => c in vehicleConfig && !c.startsWith('REPL-'));

      // Uitruk volgens kolom Voertuigen (roepnummers 265/266/257/814 = onze posten).
      // Per regel: voor elke rol waarvoor een roepnummer is uitgerukt maar GEEN eigen
      // 0933xx/0934xx code in dezelfde regel-body staat, gaat onze bemanning op een
      // vervanger. Bewust per regel — binnen één incident kan post HP (266) meerdere
      // trucks sturen (bv. NBB 3344 én een vervanger HP-TS in een andere regel).
      //
      // Uitzondering: als de regel-body GEEN enkel 09xxxx-code bevat (typisch een
      // opschalingsregel "Zeer gr. BR" / realarm) dan zijn de roepnummers in
      // Voertuigen verwijzingen naar reeds gealarmeerde trucks — geen nieuwe dispatch.
      const rowRoles = new Set(soest.map(c => vehicleConfig[c]?.role).filter(Boolean));
      const replCodes = [];
      const isUpscaleRow = codesRaw.length === 0;
      if (!isUpscaleRow) {
        for (const n of parseRoepnummers(r.Voertuigen)) {
          const role = ROEPNR_TO_ROLE[n];
          if (!rowRoles.has(role)) replCodes.push(REPL_CODE_FOR_ROLE[role]);
        }
      }
      const finalCodes = [...soest, ...replCodes];
      if (finalCodes.length === 0) continue;

      let dt = r.Datum;
      if (!(dt instanceof Date)) dt = new Date(dt);
      const prio = (PRIO_RE.exec(String(r.Prio || body)) || [])[1];
      const prioN = prio ? Number(prio) : null;
      // P5 = oefening — uitsluiten
      if (prioN === 5) continue;

      recs.push({
        id: r.Incidentnummer,
        datetime: dt,
        prio: prioN,
        location: String(r.Location ?? '').trim(),
        body,
        bodyKey: bodyAddressKey(body),
        codes: finalCodes,
        dispatches: finalCodes.map(c => ({ code: c, dt })),
      });
    }
    recs.sort((a, b) => (a.datetime?.getTime() || 0) - (b.datetime?.getTime() || 0));

    // Pass 2: merge consecutive rows that belong to the same incident.
    // 1. Same Incidentnummer => same incident (sterkste signaal).
    // 2. Anders: same prio + <= 30 min + (gelijke location OF gelijke address-key uit body).
    const incidents = [];
    for (const r of recs) {
      const last = incidents[incidents.length - 1];
      const sameId = last && r.id && last.id && r.id === last.id;
      const closeInTime = last && Math.abs((r.datetime?.getTime() || 0) - (last.datetime?.getTime() || 0)) <= 30 * 60 * 1000;
      const samePrio = last && r.prio === last.prio;
      const sameLoc = last && r.location && last.location && r.location.toLowerCase() === last.location.toLowerCase();
      const sameBodyAddr = last && r.bodyKey && last.bodyKey && r.bodyKey === last.bodyKey;
      const sameIncident = sameId || (closeInTime && samePrio && (sameLoc || sameBodyAddr));
      if (sameIncident) {
        // Behoud volgorde van alarmering: voeg de nieuwe regel’s codes/dispatches toe.
        // Geen dedup over regels heen: post HP (266) kan in twee regels twee aparte
        // trucks sturen (bv. NBB én een vervanger).
        for (const c of r.codes) last.codes.push(c);
        for (const d of r.dispatches) last.dispatches.push(d);
        last.body += ' | ' + r.body;
        last.nature = last.nature || NATURE_FIRE_RE.test(r.body);
        continue;
      }
      incidents.push({
        id: r.id,
        datetime: r.datetime,
        year: r.datetime && !isNaN(r.datetime) ? r.datetime.getFullYear() : null,
        hour: r.datetime && !isNaN(r.datetime) ? r.datetime.getHours() : null,
        prio: r.prio,
        location: r.location,
        body: r.body,
        bodyKey: r.bodyKey,
        nature: NATURE_FIRE_RE.test(r.body),
        codes: [...r.codes],
        dispatches: [...r.dispatches],
      });
    }

    // Sorteer dispatches per incident chronologisch (eerste alarm bovenaan).
    for (const inc of incidents) {
      inc.dispatches.sort((a, b) => (a.dt?.getTime() || 0) - (b.dt?.getTime() || 0));
      inc.codes = inc.dispatches.map(d => d.code);
    }

    rawIncidents = incidents;
    const mergedDelta = recs.length - incidents.length;
    setStatus(`${incidents.length} incidenten geladen (${incidents[0]?.year}–${incidents[incidents.length-1]?.year}). ${mergedDelta} regels samengevoegd tot 1 incident. P5 (oefeningen) uitgesloten.`);
    applyFilters();
    renderVehicleTable();
    runSimulation();
  } catch (err) {
    setStatus('Fout: ' + err.message);
    console.error(err);
  }
}

function setStatus(msg) {
  document.getElementById('loadStatus').textContent = msg;
}

function applyFilters() {
  const range = +document.getElementById('yearRange').value; // 0=all, else last N years inclusive
  const prioVal = document.getElementById('prioFilter').value;
  const maxYear = rawIncidents.reduce((m, i) => Math.max(m, i.year || 0), 0);
  const minYear = range === 0 ? 0 : maxYear - range + 1;

  allIncidents = rawIncidents.filter(i => {
    if (i.year && i.year < minYear) return false;
    if (prioVal !== 'all' && i.prio !== +prioVal) return false;
    return true;
  });

  const info = document.getElementById('filterInfo');
  if (info) {
    if (range === 0) info.textContent = `${allIncidents.length} incidenten (alle jaren).`;
    else info.textContent = `${allIncidents.length} incidenten in ${minYear}–${maxYear}.`;
  }
  const sub = document.getElementById('sceneSub');
  if (sub) sub.textContent = range === 0 ? `· alle jaren · ${allIncidents.length} incidenten` : `· ${minYear}–${maxYear} · ${allIncidents.length} incidenten`;
}

// ---------- Vehicle table ----------
function renderVehicleTable() {
  const counts = {};
  for (const inc of allIncidents) for (const c of inc.codes) counts[c] = (counts[c] || 0) + 1;

  const tbody = document.querySelector('#vehicleTable tbody');
  tbody.innerHTML = '';
  const allCodes = new Set([...Object.keys(vehicleConfig), ...Object.keys(counts)]);
  // sort by count desc
  const sorted = [...allCodes].sort((a, b) => (counts[b] || 0) - (counts[a] || 0));

  for (const code of sorted) {
    const cfg = vehicleConfig[code] || { label: '', role: 'OTHER', partial: false };
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><code>${code}</code></td>
      <td><input type="text" data-code="${code}" data-prop="label" value="${cfg.label}"></td>
      <td>
        <select data-code="${code}" data-prop="role">
          ${['HP-TS','NP-TS','HP-RV','HP-HA','OTHER'].map(r =>
            `<option value="${r}" ${r===cfg.role?'selected':''}>${r}</option>`).join('')}
        </select>
        <select data-code="${code}" data-prop="partial" style="margin-left:8px">
          ${[['never','volledig'],['nature','deelalarm bij natuur'],['always','altijd deelalarm']].map(([v,t]) =>
            `<option value="${v}" ${v===cfg.partial?'selected':''}>${t}</option>`).join('')}
        </select>
      </td>
      <td>${counts[code] || 0}</td>`;
    tbody.appendChild(tr);
  }

  tbody.addEventListener('change', onVehicleEdit, { once: false });
}

function onVehicleEdit(e) {
  const t = e.target;
  const code = t.dataset.code;
  const prop = t.dataset.prop;
  if (!code || !prop) return;
  if (!vehicleConfig[code]) vehicleConfig[code] = { label: '', role: 'OTHER', partial: 'never' };
  vehicleConfig[code][prop] = t.value;
  // re-classify codes per incident (not changed because we keep all)
  runSimulation();
}

// ---------- Simulation ----------

function getSettings() {
  return {
    dedTS:  +document.getElementById('dedTS').value,
    dedRV:  +document.getElementById('dedRV').value,
    leaveBehind: +document.getElementById('leaveBehindOnPartial').value,
    capTS:  +document.getElementById('capTS').value,
    capRV:  +document.getElementById('capRV').value,
    capHA:  +document.getElementById('capHA').value,
    scenario: document.querySelector('input[name=scenario]:checked').value,
  };
}

/**
 * Simulate one incident. Returns:
 *   { piketUsed, nonPiketUsed, piketIdle, conflict, understaffed, vehicles }
 */
function simulateIncident(incident, s) {
  // Build alarm list in original (chronological) order from body.
  const alarms = [];
  const postsInvolved = new Set();
  const dispatches = incident.dispatches || incident.codes.map(c => ({ code: c, dt: incident.datetime }));
  for (let i = 0; i < dispatches.length; i++) {
    const d = dispatches[i];
    const cfg = vehicleConfig[d.code];
    if (!cfg) continue;
    const isPartial = cfg.partial === 'always' || (cfg.partial === 'nature' && incident.nature);
    const v = { code: d.code, ...cfg, isPartial, post: POST_OF_ROLE[cfg.role], dt: d.dt };
    if (v.post) postsInvolved.add(v.post);
    alarms.push(v);
  }

  // Pools
  let tsPiketAvail = s.dedTS;
  let rvPiketAvail = s.dedRV;
  let leftBehindTS = 0;
  let nonPiketUsed = 0;

  // Trackers
  let firstTsHandled = false;
  let firstTsPost = null;
  let secondTsDifferentPost = false;
  let rvIdleAtRuis = false;      // (huidig) HP-RV niet gealarmeerd, RV-piket zit thuis terwijl overige vrijw. opkomen
  let rvVrijeOpkomstFlag = false; // (vast eerst) HP-RV gealarmeerd terwijl al het piket op is → vrijwilligers vullen RV
  let haCount = 0;

  // Vast-eerst afspraak: als NP-TS in dezelfde uitruk staat, gaat TS-piket naar NP-TS
  // (operationele afspraak). RV-piket kan elke HP-truck (TS/RV/HA) bemannen, want
  // alle HP-trucks staan op hetzelfde posthuis. Daardoor is "RV zonder piket" geen
  // werkelijke ruis als RV-piket op een andere HP-truck zit.
  const npTsInAlarms = alarms.some(a => a.role === 'NP-TS');
  let tsPiketAssigned = false;

  // Walk alarms in alarm order
  for (const v of alarms) {
    let need;
    if (v.role === 'HP-TS' || v.role === 'NP-TS') {
      // Bezetting: deelalarm = capTS - leaveBehind, anders capTS
      need = v.isPartial ? Math.max(0, s.capTS - s.leaveBehind) : s.capTS;
      if (!firstTsHandled) {
        firstTsPost = v.post;
        if (v.isPartial) {
          // Deelalarm: alleen piket gaat, max (dedTS - leaveBehind). Niet aanvullen
          // met overige vrijwilligers — de TS rolt onderbemand.
          // In vast-eerst: als NP-TS verderop nog komt, blijft TS-piket gereserveerd
          // voor NP-TS en gaat het deelalarm met 0 piketleden (gewoon onderbemand).
          const reserveForNp = s.scenario === 'dedicatedFirst' && npTsInAlarms && v.role !== 'NP-TS';
          if (!reserveForNp) {
            const piketSent = Math.max(0, s.dedTS - s.leaveBehind);
            leftBehindTS = s.dedTS - piketSent;
            tsPiketAvail = leftBehindTS;
            tsPiketAssigned = true;
          }
          need = 0;
        } else {
          if (s.scenario === 'dedicatedFirst') {
            const tsPiketGoesHere = !tsPiketAssigned && (
              (npTsInAlarms && v.role === 'NP-TS') ||
              (!npTsInAlarms && v.role === 'HP-TS')
            );
            if (tsPiketGoesHere) {
              const t = Math.min(tsPiketAvail, need); tsPiketAvail -= t; need -= t;
              tsPiketAssigned = true;
            }
            // RV-piket vult elke HP-truck (zelfde posthuis)
            if (v.post === 'HP' && need > 0) {
              const t = Math.min(rvPiketAvail, need); rvPiketAvail -= t; need -= t;
            }
            // Restant TS-piket (mocht designated nog niet langs zijn geweest)
            if (need > 0) {
              const t = Math.min(tsPiketAvail, need); tsPiketAvail -= t; need -= t;
            }
            nonPiketUsed += need;
          } else {
            const goes = Math.min(tsPiketAvail, need);
            tsPiketAvail -= goes;
            need -= goes;
            nonPiketUsed += need;
          }
        }
        firstTsHandled = true;
      } else {
        const splitPost = v.post && firstTsPost && v.post !== firstTsPost;
        if (splitPost) secondTsDifferentPost = true;
        if (s.scenario === 'dedicatedFirst') {
          const tsPiketGoesHere = !tsPiketAssigned && (
            (npTsInAlarms && v.role === 'NP-TS') ||
            (!npTsInAlarms && v.role === 'HP-TS')
          );
          if (tsPiketGoesHere) {
            const t = Math.min(tsPiketAvail, need); tsPiketAvail -= t; need -= t;
            tsPiketAssigned = true;
          }
          // RV-piket vult elke HP-truck
          if (v.post === 'HP' && need > 0) {
            const t = Math.min(rvPiketAvail, need); rvPiketAvail -= t; need -= t;
          }
          // Restant TS-piket
          if (need > 0) {
            const t = Math.min(tsPiketAvail, need); tsPiketAvail -= t; need -= t;
          }
          nonPiketUsed += need;
        } else {
          // Huidig/custom: 2e TS — eerst rest-TS-piket, dan overige vrijwilligers
          const t = Math.min(tsPiketAvail, need); tsPiketAvail -= t; need -= t;
          nonPiketUsed += need;
        }
      }
    } else if (v.role === 'HP-RV') {
      need = s.capRV;
      const take = Math.min(rvPiketAvail, need); rvPiketAvail -= take; need -= take;
      if (s.scenario === 'dedicatedFirst') {
        // RV-piket kan al weg zijn (naar HP-TS in dit incident) — dan vult TS-piket
        // (als die nog over is) of overige vrijwilligers. Geen "stranded" flag meer:
        // RV-piket zit altijd op HP-post als die werk had.
        const t = Math.min(tsPiketAvail, need); tsPiketAvail -= t; need -= t;
        // Vrije opkomst: er is RV-bezetting nodig, maar zowel RV- als TS-piket zijn op.
        if (need > 0) rvVrijeOpkomstFlag = true;
      }
      nonPiketUsed += need;
    } else if (v.role === 'HP-HA') {
      // 1e HA = volledige bezetting, elke extra HA in zelfde incident = +2 personen
      haCount++;
      need = haCount === 1 ? s.capHA : 2;
      if (s.scenario === 'current') {
        // Huidig: HA is primair voor vrije opkomst — al het piket (TS \u00e9n RV)
        // blijft thuis, vrijwilligers vullen de HA volledig.
        nonPiketUsed += need;
      } else if (s.scenario === 'dedicatedFirst') {
        let t = Math.min(tsPiketAvail, need); tsPiketAvail -= t; need -= t;
        t     = Math.min(rvPiketAvail, need); rvPiketAvail -= t; need -= t;
        nonPiketUsed += need;
      } else {
        nonPiketUsed += need;
      }
    }
  }

  const piketIdle = tsPiketAvail + rvPiketAvail;
  const piketUsed = (s.dedTS + s.dedRV) - piketIdle;
  const conflict = piketIdle > 0 && nonPiketUsed > 0;

  // Huidig: RV-piket blijft thuis. Als RV niet alarmeerd én er overige vrijw. opgeroepen zijn → ruis-bron.
  if (s.scenario === 'current' && rvPiketAvail > 0 && nonPiketUsed > 0) rvIdleAtRuis = true;

  // ----- Eigenschap-labels (gelden in beide scenario's) -----
  const hasHpTs = alarms.some(a => a.role === 'HP-TS');
  const hasNpTs = alarms.some(a => a.role === 'NP-TS');
  const hasRv   = alarms.some(a => a.role === 'HP-RV');

  // Split-post: HP-TS én NP-TS gealarmeerd binnen 2 minuten van elkaar.
  const SPLIT_WINDOW_MS = 2 * 60 * 1000;
  let splitPost = false;
  if (hasHpTs && hasNpTs) {
    const hpTimes = alarms.filter(a => a.role === 'HP-TS' && a.dt).map(a => a.dt.getTime());
    const npTimes = alarms.filter(a => a.role === 'NP-TS' && a.dt).map(a => a.dt.getTime());
    for (const t1 of hpTimes) {
      for (const t2 of npTimes) {
        if (Math.abs(t1 - t2) <= SPLIT_WINDOW_MS) { splitPost = true; break; }
      }
      if (splitPost) break;
    }
  }

  // RV vrije opkomst (alleen vast-eerst): HP-RV gealarmeerd terwijl al het piket
  // (TS én RV) op is — vrijwilligers vullen de RV. Wordt in de HP-RV-branche geset.
  // In huidig blijft RV-piket altijd thuis tot RV gealarmeerd, dus daar bestaat
  // dit fenomeen niet.
  const rvVrijeOpkomst = rvVrijeOpkomstFlag;

  // Alleen HA: alleen HP-HA in alarmen, alle piket blijft thuis.
  const allRoles = new Set(alarms.map(a => a.role));
  const haOnlyIdle = allRoles.size === 1 && allRoles.has('HP-HA') && piketIdle === (s.dedTS + s.dedRV);

  // NBB met 4 thuis: TS-deelalarm uitgerukt (2 TS-piket thuis) én RV-piket volledig
  // thuis omdat RV niet gealarmeerd. Geldt in beide scenario's.
  const nbb4Idle = leftBehindTS === 2 && !hasRv && rvPiketAvail === s.dedRV;

  return {
    piketUsed, nonPiketUsed, piketIdle, conflict,
    splitPost, rvVrijeOpkomst, haOnlyIdle, nbb4Idle, rvIdleAtRuis,
    leftBehindTS,
    posts: [...postsInvolved],
    counts: {
      ts: alarms.filter(a => a.role === 'HP-TS' || a.role === 'NP-TS').length,
      rv: alarms.filter(a => a.role === 'HP-RV').length,
      ha: alarms.filter(a => a.role === 'HP-HA').length,
    },
  };
}

function runSimulation() {
  if (allIncidents.length === 0) return;
  const s = getSettings();

  // Run current chosen scenario AND the alternative for comparison
  const altScenario = s.scenario === 'dedicatedFirst' ? 'current' : 'dedicatedFirst';

  const main = allIncidents.map(inc => simulateIncident(inc, s));
  const alt = allIncidents.map(inc => simulateIncident(inc, { ...s, scenario: altScenario }));

  renderKpis(s, main, alt, altScenario);
  renderScene(main);
  renderOutcomeChart(main);
  renderYearChart(main);
  renderComboChart();
  renderHourChart(main);
  const SCEN_LABELS = { current: 'Huidig', dedicatedFirst: 'Vast eerst altijd', custom: 'Aangepast' };
  renderExamples(main, SCEN_LABELS[s.scenario] || s.scenario);
}

function setStat(name, val) {
  document.querySelectorAll(`[data-stat="${name}"]`).forEach(el => el.textContent = val);
}

function renderScene(results) {
  const counts = { 'HP-TS':0, 'NP-TS':0, 'HP-RV':0, 'HP-HA':0 };
  const conflicts = { 'HP-TS':0, 'NP-TS':0, 'HP-RV':0, 'HP-HA':0 };
  results.forEach((r, i) => {
    const seenRoles = new Set();
    for (const code of allIncidents[i].codes) {
      const role = vehicleConfig[code]?.role;
      if (role && role in counts && !seenRoles.has(role)) {
        counts[role]++;
        if (r.conflict) conflicts[role]++;
        seenRoles.add(role);
      }
    }
  });
  for (const role of Object.keys(counts)) {
    setStat(`${role}-count`, counts[role]);
    setStat(`${role}-conflict`, conflicts[role]);
  }
  setStat('HP-TS', counts['HP-TS']);
  setStat('HP-RV', counts['HP-RV']);
  setStat('HP-HA', counts['HP-HA']);
  setStat('NP-TS', counts['NP-TS']);
  setStat('HP-total', counts['HP-TS'] + counts['HP-RV'] + counts['HP-HA']);
  setStat('NP-total', counts['NP-TS']);
}

function pct(n, d) { return d === 0 ? '0%' : ((n / d) * 100).toFixed(1) + '%'; }

function renderKpis(s, main, alt, altName) {
  const total = main.length;
  const conflicts = main.filter(r => r.conflict).length;
  const altConflicts = alt.filter(r => r.conflict).length;
  const nonPiket = main.reduce((a, r) => a + r.nonPiketUsed, 0);
  const altNonPiket = alt.reduce((a, r) => a + r.nonPiketUsed, 0);
  const splitPost = main.filter(r => r.splitPost).length;
  const rvVrij = main.filter(r => r.rvVrijeOpkomst).length;
  const haOnlyIdle = main.filter(r => r.haOnlyIdle).length;
  const nbb4Idle = main.filter(r => r.nbb4Idle).length;
  const rvIdle = main.filter(r => r.rvIdleAtRuis).length;

  const labels = { current: 'Huidig', dedicatedFirst: 'Vast eerst altijd', custom: 'Aangepast' };
  const mainName = labels[s.scenario];
  const altLbl  = labels[altName];

  const dConf = altConflicts - conflicts;
  const dNon  = altNonPiket - nonPiket;
  const fmt = (n, suffix) => {
    if (n === 0) return `gelijk (${suffix})`;
    return n < 0 ? `${Math.abs(n)} minder ${suffix}` : `${n} meer ${suffix}`;
  };
  const conflictDeltaCls = dConf > 0 ? 'ok' : (dConf < 0 ? 'bad' : '');

  document.getElementById('kpis').innerHTML = `
    <div class="kpi"><div class="num">${total}</div><div class="lbl">incidenten geanalyseerd</div></div>
    <div class="kpi bad"><div class="num">${conflicts}</div><div class="lbl">ruis-incidenten in scenario «${mainName}» (${pct(conflicts, total)})</div></div>
    <div class="kpi"><div class="num">${nonPiket}</div><div class="lbl">oproepen overige vrijwilligers («${mainName}»)</div></div>
    <div class="kpi"><div class="num">${altConflicts}</div><div class="lbl">ruis-incidenten bij «${altLbl}» (${pct(altConflicts, total)})</div></div>
    <div class="kpi ${conflictDeltaCls}"><div class="num">${fmt(dConf, 'ruis')}</div><div class="lbl">«${altLbl}» t.o.v. «${mainName}»</div></div>
    <div class="kpi"><div class="num">${fmt(dNon, 'oproepen')}</div><div class="lbl">«${altLbl}» t.o.v. «${mainName}»</div></div>
    <div class="kpi bad"><div class="num">${splitPost}</div><div class="lbl">split-post: HP-TS én NP-TS samen gealarmeerd</div></div>
    <div class="kpi bad"><div class="num">${rvVrij}</div><div class="lbl">RV vrije opkomst («Vast eerst»): HP-RV gealarmeerd terwijl al het piket op is</div></div>
    <div class="kpi bad"><div class="num">${rvIdle}</div><div class="lbl">RV-piket zit thuis terwijl overige vrijw. opkomen («Huidig»)</div></div>
    <div class="kpi bad"><div class="num">${haOnlyIdle}</div><div class="lbl">alleen HA gealarmeerd — al het piket blijft thuis</div></div>
    <div class="kpi bad"><div class="num">${nbb4Idle}</div><div class="lbl">NBB-deelalarm: 2 TS-piket + 2 RV-piket blijven thuis</div></div>
  `;
}

// ---------- Charts ----------

function destroyChart(key) {
  if (charts[key]) { charts[key].destroy(); delete charts[key]; }
}

function renderOutcomeChart(results) {
  destroyChart('outcome');
  const ctx = document.getElementById('outcomeChart');
  const ok = results.filter(r => !r.conflict).length;
  const bad = results.filter(r => r.conflict).length;
  charts.outcome = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: ['Geen ruis', 'Ruis (piket thuis & overige vrijw. opgeroepen)'],
      datasets: [{ data: [ok, bad], backgroundColor: ['#4caf50', '#ef5350'] }],
    },
    options: { plugins: { legend: { labels: { color: '#e6edf3' } } } },
  });
}

function renderYearChart(results) {
  destroyChart('year');
  const byYear = {};
  results.forEach((r, i) => {
    const y = allIncidents[i].year;
    if (!y) return;
    byYear[y] ??= { ok: 0, bad: 0 };
    if (r.conflict) byYear[y].bad++; else byYear[y].ok++;
  });
  const years = Object.keys(byYear).sort();
  const ctx = document.getElementById('yearChart');
  charts.year = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: years,
      datasets: [
        { label: 'Geen ruis', data: years.map(y => byYear[y].ok), backgroundColor: '#4caf50' },
        { label: 'Ruis',      data: years.map(y => byYear[y].bad), backgroundColor: '#ef5350' },
      ],
    },
    options: {
      scales: {
        x: { stacked: true, ticks: { color: '#8aa0b6' } },
        y: { stacked: true, ticks: { color: '#8aa0b6' } },
      },
      plugins: { legend: { labels: { color: '#e6edf3' } } },
    },
  });
}

function renderComboChart() {
  destroyChart('combo');
  // Build combos: sorted role tuples
  const combos = {};
  for (const inc of allIncidents) {
    const roles = inc.codes.map(c => vehicleConfig[c]?.role || 'OTHER');
    const key = [...new Set(roles)].sort().join(' + ');
    combos[key] = (combos[key] || 0) + 1;
  }
  const sorted = Object.entries(combos).sort((a, b) => b[1] - a[1]).slice(0, 10);
  const ctx = document.getElementById('comboChart');
  charts.combo = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: sorted.map(([k]) => k),
      datasets: [{ label: 'Incidenten', data: sorted.map(([, v]) => v), backgroundColor: '#ff7a1a' }],
    },
    options: {
      indexAxis: 'y',
      scales: {
        x: { ticks: { color: '#8aa0b6' } },
        y: { ticks: { color: '#e6edf3' } },
      },
      plugins: { legend: { labels: { color: '#e6edf3' } } },
    },
  });
}

function renderHourChart(results) {
  destroyChart('hour');
  const buckets = Array.from({length: 24}, () => ({ ok: 0, bad: 0 }));
  results.forEach((r, i) => {
    const h = allIncidents[i].hour;
    if (h == null || h < 0 || h > 23) return;
    if (r.conflict) buckets[h].bad++; else buckets[h].ok++;
  });
  const ctx = document.getElementById('hourChart');
  charts.hour = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: buckets.map((_, i) => String(i).padStart(2,'0')),
      datasets: [
        { label: 'Geen ruis', data: buckets.map(b => b.ok), backgroundColor: '#4caf50' },
        { label: 'Ruis',      data: buckets.map(b => b.bad), backgroundColor: '#ef5350' },
      ],
    },
    options: {
      scales: {
        x: { stacked: true, ticks: { color: '#8aa0b6' } },
        y: { stacked: true, ticks: { color: '#8aa0b6' } },
      },
      plugins: { legend: { labels: { color: '#e6edf3' } } },
    },
  });
}

// Active filter for the incidents table.
let exampleFilter = 'all';

function renderExamples(results, scenarioLabel) {
  const tbody = document.querySelector('#examples tbody');
  tbody.innerHTML = '';

  // Sort by date descending (newest first)
  const idxs = results.map((_, i) => i)
    .sort((a, b) => (allIncidents[b].datetime?.getTime() || 0) - (allIncidents[a].datetime?.getTime() || 0));

  // Compute counts per filter
  const counts = { all: idxs.length, ruis: 0, ok: 0,
    splitPost: 0, rvVrijeOpkomst: 0, haOnlyIdle: 0, nbb4Idle: 0, rvIdleAtRuis: 0 };
  const isRuisR = (r) => r.conflict || r.splitPost || r.rvVrijeOpkomst || r.rvIdleAtRuis || r.haOnlyIdle || r.nbb4Idle;
  for (const i of idxs) {
    const r = results[i];
    if (isRuisR(r)) counts.ruis++; else counts.ok++;
    if (r.splitPost)         counts.splitPost++;
    if (r.rvVrijeOpkomst)    counts.rvVrijeOpkomst++;
    if (r.haOnlyIdle)        counts.haOnlyIdle++;
    if (r.nbb4Idle)          counts.nbb4Idle++;
    if (r.rvIdleAtRuis)      counts.rvIdleAtRuis++;
  }

  // Update filter button counts
  document.querySelectorAll('.examples-filters .filt').forEach(btn => {
    const k = btn.dataset.filter;
    const cnt = btn.querySelector('.cnt');
    if (cnt) cnt.textContent = counts[k] ?? 0;
    btn.classList.toggle('active', k === exampleFilter);
  });

  // Apply filter
  const passes = (r) => {
    switch (exampleFilter) {
      case 'all':              return true;
      case 'ruis':             return isRuisR(r);
      case 'ok':               return !isRuisR(r);
      case 'splitPost':        return r.splitPost;
      case 'rvVrijeOpkomst':   return r.rvVrijeOpkomst;
      case 'haOnlyIdle':       return r.haOnlyIdle;
      case 'nbb4Idle':         return r.nbb4Idle;
      case 'rvIdleAtRuis':     return r.rvIdleAtRuis;
      default: return true;
    }
  };

  const frag = document.createDocumentFragment();
  let shown = 0;
  for (const i of idxs) {
    const r = results[i];
    if (!passes(r)) continue;
    const inc = allIncidents[i];
    const tags = [];
    if (r.splitPost)         tags.push('<span class="tag tag-split">split-post</span>');
    if (r.rvVrijeOpkomst)    tags.push('<span class="tag tag-rvo">RV vrije opkomst</span>');
    if (r.haOnlyIdle)        tags.push('<span class="tag tag-ha">alleen HA</span>');
    if (r.nbb4Idle)          tags.push('<span class="tag tag-nbb">NBB — 4 piket thuis</span>');
    if (r.rvIdleAtRuis)      tags.push('<span class="tag tag-idle">RV-piket thuis</span>');
    const isRuis = isRuisR(r);
    const reason = isRuis
      ? `${r.nonPiketUsed} overige vrijw. opgeroepen, ${r.piketIdle} piket thuis ${tags.join(' ')}`
      : `<span class="muted">geen ruis (${r.nonPiketUsed} overige vrijw., ${r.piketIdle} piket thuis)</span>`;
    const veh = renderDispatches(inc.dispatches);
    const cleanLoc = stripVehicleCodes(inc.location);
    const tr = document.createElement('tr');
    tr.className = isRuis ? 'row-ruis' : 'row-ok';
    tr.innerHTML = `
      <td>${inc.datetime?.toISOString().slice(0,16).replace('T',' ') ?? ''}</td>
      <td>P${inc.prio ?? '?'}</td>
      <td>${veh}</td>
      <td>${reason}</td>
      <td>${escapeHtml(cleanLoc)}</td>`;
    frag.appendChild(tr);
    shown++;
  }
  tbody.appendChild(frag);

  // Scenario badge + hint
  const badge = document.getElementById('scenarioBadge');
  if (badge) badge.textContent = `· scenario: ${scenarioLabel}`;
  const hint = document.getElementById('examplesHint');
  if (hint) hint.textContent = `${shown} van ${idxs.length} incidenten getoond. Klik een filter om bij te stellen.`;
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

// Verwijder voertuigcodes (09xxxx) uit lokatietekst zodat alleen het adres/incidenttype overblijft.
function stripVehicleCodes(s) {
  return String(s || '').replace(/\b09\d{4}\b/g, '').replace(/\s+/g, ' ').trim();
}

// HH:MM van een Date.
function fmtTime(dt) {
  if (!(dt instanceof Date) || isNaN(dt)) return '';
  return dt.toISOString().slice(11, 16);
}

// Toon dispatches gestapeld met alarmtijd en Δt t.o.v. de eerste alarmering.
function renderDispatches(dispatches) {
  if (!dispatches || dispatches.length === 0) return '';
  const t0 = dispatches[0].dt?.getTime();
  const lines = dispatches.map((d, idx) => {
    const label = vehicleConfig[d.code]?.label || d.code;
    const time = fmtTime(d.dt);
    let delta = '';
    if (idx > 0 && t0 != null && d.dt) {
      const mins = Math.round((d.dt.getTime() - t0) / 60000);
      if (mins > 0) delta = ` <span class="muted">(+${mins} min)</span>`;
    }
    return `<div class="dispatch"><span class="dispatch-time">${time}</span> ${escapeHtml(label)}${delta}</div>`;
  });
  return lines.join('');
}

// ---------- Wire UI ----------

document.getElementById('btnLoadDefault').addEventListener('click', loadDefaultFile);
document.getElementById('fileInput').addEventListener('change', (e) => {
  const f = e.target.files[0];
  if (f) loadFile(f);
});

// Re-run simulation on any settings change
for (const id of ['dedTS','dedRV','leaveBehindOnPartial','capTS','capRV','capHA']) {
  document.getElementById(id).addEventListener('input', runSimulation);
}
document.querySelectorAll('input[name=scenario]').forEach(r => r.addEventListener('change', runSimulation));

for (const id of ['yearRange','prioFilter']) {
  document.getElementById(id).addEventListener('change', () => { applyFilters(); runSimulation(); });
}

// Filter knoppen voor incidententabel
document.querySelectorAll('.examples-filters .filt').forEach(btn => {
  btn.addEventListener('click', () => {
    exampleFilter = btn.dataset.filter;
    runSimulation();
  });
});

// Auto-load default file on first paint
window.addEventListener('DOMContentLoaded', loadDefaultFile);

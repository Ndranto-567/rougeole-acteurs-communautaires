// ═══════════════════════════════════════════════════════
//  ⚠️  CONFIGURATION — MODIFIEZ CETTE LIGNE UNIQUEMENT
//  Collez ici l'URL obtenue après le déploiement Apps Script
// ═══════════════════════════════════════════════════════
const GS_URL = 'https://script.google.com/macros/s/AKfycbxy63zxnxBwJX1dYpq46D-mv3UQrAvOsr0WPg4sGC-G6DYdlm1jaLIMn8B3kVr0MQwN4A/exec';

//  Exemple :
//  const GS_URL = 'https://script.google.com/macros/s/AKfycb.../exec';

// ═══════════════════════════════════════════════════════
//  DONNÉES RÉGIONS/DISTRICTS
// ═══════════════════════════════════════════════════════
const REGIONS_DISTRICTS = {
  "ALAOTRA MANGORO":["AMBATONDRAZAKA","AMPARAFARAVOLA","ANDILAMENA","ANOSIBE AN'ALA","MORAMANGA"],
  "AMORON'I MANIA":["AMBATOFINANDRAHANA","AMBOSITRA","FANDRIANA","MANANDRIANA"],
  "ANALAMANGA":["AMBOHIDRATRIMO","ANKAZOBE","ANDRAMASINA","ANTANANARIVO ATSIMONDRANO","ANJOZOROBE","ANTANANARIVO AVARADRANO","ANTANANARIVO RENIVOHITRA","MANJAKANDRIANA"],
  "ANALANJIROFO":["FENERIVE EST","VAVATENINA","SAINTE MARIE","SOANIERANA IVONGO","MANANARA NORD","MAROANTSETRA"],
  "ANDROY":["AMBOVOMBE","ANTANIMORA","BEKILY","BELOHA","TSIHOMBE"],
  "ANOSY":["AMBOASARY SUD","BETROKA","TOLAGNARO"],
  "ATSIMO ATSINANANA":["BEFOTAKA","MIDONGY DU SUD","FARAFANGANA","VANGAINDRANO","VONDROZO"],
  "ATSIMO-ANDREFANA":["AMPANIHY ANDREFANA","BETIOKY-SUD","BENENITRA","SAKARAHA","ANKAZOABO","BEROROHA","MOROMBE","TOLIARA II","TOLIARA I"],
  "ATSINANANA":["ANTANAMBAO MANAMPOTSY","MAHANORO","MAROLAMBO","TOAMASINA I","TOAMASINA II","BRICKAVILLE","VATOMANDRY"],
  "BETSIBOKA":["KANDREHO","MAEVATANANA","TSARATANANA NORD","TSARATANANA SUD"],
  "BOENI":["AMBATO-BOENI","MAROVOAY","MAHAJANGA I","MAHAJANGA II","MITSINJO","SOALALA"],
  "BONGOLAVA":["FENOARIVOBE","TSIROANOMANDIDY"],
  "DIANA":["AMBANJA","AMBILOBE","ANTSIRANANA I","ANTSIRANANA II","NOSY BE"],
  "FITOVINANY":["IKONGO","MANAKARA","VOHIPENO"],
  "HAUTE MATSIATRA":["AMBALAVAO","FII_VOHIBATO","AMBOHIMAHASOA","FIANARANTSOA I","FII_LALANGINA","FII_ISANDRA","IKALAMAVONY"],
  "IHOROMBE":["IAKORA","IHOSY","IVOHIBE"],
  "ITASY":["ARIVONIMAMO","MIARINARIVO","SOAVINANDRIANA"],
  "MELAKY":["AMBATOMAINTY","MORAFENOBE","ANTSALOVA","MAINTIRANO","BESALAMPY"],
  "MENABE":["BELO/TSIRIBIHINA","MORONDAVA","MANJA","MIANDRIVAZO","MAHABO"],
  "S A V A":["ANDAPA","ANTALAHA","SAMBAVA","VOHEMAR"],
  "SOFIA":["ANALALAVA","ANTSOHIHY","BEALALANA","MAMPIKONY","PORT-BERGER","MANDRITSARA","BEFANDRIANA NORD"],
  "VAKINANKARATRA":["AMBATOLAMPY","ANTANIFOTSY","ANTSIRABE I","ANTSIRABE II","FARATSIHO","BETAFO","MANDOTO"],
  "VATOVAVY":["IFANADIANA","MANANJARY","NOSY VARIKA"]
};

// ═══════════════════════════════════════════════════════
//  MAPPING COLONNES EXCEL → champs internes
// ═══════════════════════════════════════════════════════
const COL_MAP = {
  'region':          'region',
  'district':        'district',
  'csb':             'csb',
  'fokontany':       'fokontany',
  'nomet prenom':    'nom',
  'nom et prenom':   'nom',
  'nom etprenom':    'nom',
  'nom':             'nom',
  'poste':           'poste',
  'fonction':        'poste',
  'cin':             'cin',
  'nummvola':        'mvola',
  "numm'vola":       'mvola',
  'nummobile money': 'mvola',
  'mobile money':    'mvola',
  'mvola':           'mvola',
  'mvola/airtel':    'mvola'
};

// Regex de validation
const REGEX_CIN   = /^\d{3} \d{3} \d{3} \d{3}$/;  // 12 chiffres : XXX XXX XXX XXX
const REGEX_TEL   = /^\d{3} \d{2} \d{3} \d{2}$/;  // 10 chiffres : XXX XX XXX XX

// ═══════════════════════════════════════════════════════
//  ÉTAT
// ═══════════════════════════════════════════════════════
let rows = [];
let filteredRows = [];
let rowIdCounter = 0;
let currentPage = 1;
const PAGE_SIZE = 50;
let searchQuery = '';

// ═══════════════════════════════════════════════════════
//  INIT
// ═══════════════════════════════════════════════════════
// Init appelé directement depuis le HTML via onload, ou ici si DOM déjà prêt
function initPage() {
  var sel = document.getElementById('sel-region');
  if (!sel) return;
  Object.keys(REGIONS_DISTRICTS).sort().forEach(function(r) {
    sel.appendChild(new Option(r, r));
  });
  var tw = document.getElementById('table-wrap');
  if (tw) tw.addEventListener('paste', onPaste);
}

// ═══════════════════════════════════════════════════════
//  LOCALISATION
// ═══════════════════════════════════════════════════════
function onRegionChange() {
  const region = document.getElementById('sel-region').value;
  const sd = document.getElementById('sel-district');
  sd.innerHTML = '<option value="">— Sélectionner un district —</option>';
  if (region && REGIONS_DISTRICTS[region]) {
    REGIONS_DISTRICTS[region].forEach(d => sd.appendChild(new Option(d, d)));
    sd.disabled = false;
  } else {
    sd.disabled = true;
  }
  setErr('grp-region', false);
}

function onDistrictChange() {
  setErr('grp-district', false);
}

// ═══════════════════════════════════════════════════════
//  DRAG & DROP / IMPORT FICHIER
// ═══════════════════════════════════════════════════════
function onDragOver(e) {
  e.preventDefault();
  document.getElementById('import-zone').classList.add('drag-over');
}
function onDragLeave(e) {
  document.getElementById('import-zone').classList.remove('drag-over');
}
function onDrop(e) {
  e.preventDefault();
  document.getElementById('import-zone').classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file) processFile(file);
}
function onFileSelect(e) {
  const file = e.target.files[0];
  if (file) processFile(file);
  e.target.value = '';
}

// ═══════════════════════════════════════════════════════
//  TRAITEMENT FICHIER EXCEL / CSV
// ═══════════════════════════════════════════════════════
function processFile(file) {
  showProgress(true, 'Lecture du fichier…', 20);
  const reader = new FileReader();
  reader.onload = function(ev) {
    try {
      showProgress(true, 'Analyse des colonnes…', 50);
      let importedRows = [];
      if (file.name.endsWith('.csv')) {
        importedRows = parseCSV(ev.target.result);
      } else {
        const wb = XLSX.read(ev.target.result, { type: 'binary' });
        let sheetName = wb.SheetNames.find(n => normalizeKey(n).includes('acteurscommunautaire'))
                     || wb.SheetNames[0];
        const ws = wb.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(ws, { defval: '' });
        importedRows = mapExcelRows(jsonData);
      }
      showProgress(true, 'Import de ' + importedRows.length + ' ligne(s)…', 80);
      if (rows.length > 0) {
        if (confirm('Le tableau contient déjà ' + rows.length + ' ligne(s).\nCliquez OK pour AJOUTER, ou Annuler pour REMPLACER.')) {
          rows = rows.concat(importedRows);
        } else {
          rows = importedRows;
        }
      } else {
        rows = importedRows;
      }
      showProgress(true, 'Rendu du tableau…', 95);
      setTimeout(function() {
        renderTable();
        showProgress(false);
        showToast('✓ ' + importedRows.length + ' acteur(s) importé(s) depuis "' + file.name + '"', 'success');
      }, 100);
    } catch(err) {
      console.error(err);
      showProgress(false);
      showToast('Erreur lors de la lecture du fichier : ' + err.message, 'error');
    }
  };
  reader.onerror = function() { showProgress(false); showToast('Impossible de lire le fichier.', 'error'); };
  if (file.name.endsWith('.csv')) reader.readAsText(file, 'UTF-8');
  else reader.readAsBinaryString(file);
}

function normalizeKey(str) {
  return String(str).toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9']/g, '');
}

function mapExcelRows(jsonData) {
  return jsonData.map(function(row) {
    var mapped = { csb:'', fokontany:'', nom:'', poste:'', cin:'', mvola:'' };
    Object.keys(row).forEach(function(col) {
      var key = normalizeKey(col);
      var field = COL_MAP[key];
      if (field && field !== 'region' && field !== 'district') {
        mapped[field] = String(row[col] || '').trim();
      }
    });
    mapped._id = ++rowIdCounter;
    return mapped;
  }).filter(function(r) { return r.nom || r.csb || r.fokontany; });
}

function parseCSV(text) {
  var lines = text.split('\n').filter(function(l) { return l.trim(); });
  if (lines.length < 2) return [];
  var headers = lines[0].split(',').map(function(h) { return h.replace(/"/g,'').trim(); });
  return lines.slice(1).map(function(line) {
    var vals = line.split(',').map(function(v) { return v.replace(/"/g,'').trim(); });
    var row = {};
    headers.forEach(function(h, i) { row[h] = vals[i] || ''; });
    var mapped = mapExcelRows([row]);
    return mapped[0];
  }).filter(Boolean);
}

// ═══════════════════════════════════════════════════════
//  COLLER DEPUIS PRESSE-PAPIERS (Ctrl+V)
// ═══════════════════════════════════════════════════════
function onPaste(e) {
  var text = (e.clipboardData || window.clipboardData).getData('text');
  if (!text) return;
  var lines = text.split('\n').filter(function(l) { return l.trim(); });
  if (!lines.length) return;

  var firstLineLower = normalizeKey(lines[0]);
  if (firstLineLower.includes('nom') || firstLineLower.includes('csb') || firstLineLower.includes('region')) {
    var headers = lines[0].split('\t').map(function(h) { return h.trim(); });
    var dataLines = lines.slice(1).map(function(line) {
      var vals = line.split('\t');
      var row = {};
      headers.forEach(function(h, i) { row[h] = vals[i] || ''; });
      return row;
    });
    var mapped = mapExcelRows(dataLines);
    if (mapped.length) {
      rows = rows.concat(mapped);
      renderTable();
      showToast('✓ ' + mapped.length + ' ligne(s) collée(s) avec en-têtes', 'success');
      e.preventDefault();
      return;
    }
  }

  var newRows = lines.map(function(line) {
    var cols = line.split('\t');
    return {
      _id:       ++rowIdCounter,
      csb:       (cols[0] || '').trim(),
      fokontany: (cols[1] || '').trim(),
      nom:       (cols[2] || '').trim(),
      poste:     (cols[3] || '').trim(),
      cin:       (cols[4] || '').trim(),
      mvola:     (cols[5] || '').trim()
    };
  }).filter(function(r) { return r.nom || r.csb; });

  if (newRows.length) {
    rows = rows.concat(newRows);
    renderTable();
    showToast('✓ ' + newRows.length + ' ligne(s) collée(s)', 'success');
    e.preventDefault();
  }
}

// ═══════════════════════════════════════════════════════
//  RENDU DU TABLEAU
// ═══════════════════════════════════════════════════════
function renderTable() {
  if (searchQuery) {
    var q = searchQuery.toLowerCase();
    filteredRows = rows.filter(function(r) {
      return (r.nom||'').toLowerCase().includes(q) ||
             (r.csb||'').toLowerCase().includes(q) ||
             (r.fokontany||'').toLowerCase().includes(q) ||
             (r.poste||'').toLowerCase().includes(q) ||
             (r.cin||'').toLowerCase().includes(q);
    });
  } else {
    filteredRows = rows.slice();
  }

  var totalPages = Math.max(1, Math.ceil(filteredRows.length / PAGE_SIZE));
  if (currentPage > totalPages) currentPage = totalPages;

  var start    = (currentPage - 1) * PAGE_SIZE;
  var pageRows = filteredRows.slice(start, start + PAGE_SIZE);
  var tbody    = document.getElementById('ac-tbody');
  tbody.innerHTML = '';

  if (filteredRows.length === 0) {
    tbody.innerHTML =
      '<tr id="ac-empty-row"><td colspan="8" class="ac-empty">' +
      '<div class="ac-empty-icon">👥</div>' +
      '<div>' + (rows.length > 0 ? 'Aucun résultat pour cette recherche.' : 'Aucun acteur. Importez un fichier Excel ou ajoutez des lignes manuellement.') + '</div>' +
      '</td></tr>';
    document.getElementById('stats-bar').style.display = 'none';
    renderPagination(0, 0);
    updateCount();
    return;
  }

  pageRows.forEach(function(row, pageIdx) {
    var globalIdx  = rows.indexOf(row);
    var displayNum = start + pageIdx + 1;

    // Validation visuelle CIN et M'vola
    var cinVal    = (row.cin   || '').trim();
    var mvolaVal  = (row.mvola || '').trim();
    var cinClass  = cinVal   === '' ? '' : (REGEX_CIN.test(cinVal)   ? 'field-ok' : 'field-err');
    var mvolaClass= mvolaVal === '' ? '' : (REGEX_TEL.test(mvolaVal) ? 'field-ok' : 'field-err');

    var tr = document.createElement('tr');
    tr.dataset.idx = globalIdx;
    tr.innerHTML =
      '<td class="col-num">' + displayNum + '</td>' +
      '<td><input type="text" class="f-csb" value="' + esc(row.csb) + '" placeholder="Nom du CSB" onchange="updateRow(' + globalIdx + ',\'csb\',this.value)"></td>' +
      '<td><input type="text" class="f-fkt" value="' + esc(row.fokontany) + '" placeholder="Fokontany" onchange="updateRow(' + globalIdx + ',\'fokontany\',this.value)"></td>' +
      '<td><input type="text" class="f-nom" value="' + esc(row.nom) + '" placeholder="Nom et Prénom" onchange="updateRow(' + globalIdx + ',\'nom\',this.value)"></td>' +
      '<td><input type="text" class="f-pos" value="' + esc(row.poste) + '" placeholder="Poste/Fonction" onchange="updateRow(' + globalIdx + ',\'poste\',this.value)"></td>' +
      '<td><input type="text" class="f-cin ' + cinClass + '" value="' + esc(row.cin) + '" placeholder="XXX XXX XXX XXX" maxlength="15" onkeydown="onlyDigits(event)" oninput="formatCIN(this)" inputmode="numeric" autocomplete="off" title="12 chiffres : XXX XXX XXX XXX" onchange="updateRow(' + globalIdx + ',\'cin\',this.value);renderTable();"></td>' +
      '<td><input type="text" class="f-mvola ' + mvolaClass + '" value="' + esc(row.mvola) + '" placeholder="03X XX XXX XX" maxlength="13" onkeydown="onlyDigits(event)" oninput="formatMobileMoney(this)" inputmode="numeric" autocomplete="off" title="10 chiffres : XXX XX XXX XX" onchange="updateRow(' + globalIdx + ',\'mvola\',this.value);renderTable();"></td>' +
      '<td class="col-act"><button class="btn-del-row" onclick="deleteRow(' + globalIdx + ')">✕</button></td>';
    tbody.appendChild(tr);
  });

  updateStats();
  renderPagination(filteredRows.length, totalPages);
  updateCount();
  document.getElementById('stats-bar').style.display = 'flex';
}

function esc(str) {
  return String(str || '').replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function updateRow(idx, field, value) {
  if (rows[idx]) rows[idx][field] = value;
}

function deleteRow(idx) {
  rows.splice(idx, 1);
  renderTable();
}

function addRow() {
  rows.push({ _id: ++rowIdCounter, csb:'', fokontany:'', nom:'', poste:'', cin:'', mvola:'' });
  currentPage = Math.ceil(rows.length / PAGE_SIZE) || 1;
  renderTable();
  setTimeout(function() {
    var lastRow = document.querySelector('#ac-tbody tr:last-child input.f-csb');
    if (lastRow) lastRow.focus();
  }, 50);
}

function clearAll() {
  if (rows.length === 0) return;
  if (!confirm('Supprimer les ' + rows.length + ' acteur(s) du tableau ?')) return;
  rows = [];
  currentPage = 1;
  renderTable();
}

// ═══════════════════════════════════════════════════════
//  STATS & PAGINATION
// ═══════════════════════════════════════════════════════
function updateStats() {
  document.getElementById('stat-total').textContent  = filteredRows.length;
  document.getElementById('stat-cin').textContent    = filteredRows.filter(function(r) { return r.cin  && r.cin.trim();   }).length;
  document.getElementById('stat-mvola').textContent  = filteredRows.filter(function(r) { return r.mvola && r.mvola.trim(); }).length;
}

function updateCount() {
  var n = rows.length;
  document.getElementById('ac-count').textContent = '(' + n + ' acteur' + (n > 1 ? 's' : '') + ')';
}

function renderPagination(total, totalPages) {
  var pg   = document.getElementById('pagination');
  var info = document.getElementById('page-info');
  pg.innerHTML = '';
  if (totalPages <= 1) { info.textContent = total > 0 ? total + ' acteur(s)' : ''; return; }

  var start = (currentPage - 1) * PAGE_SIZE + 1;
  var end   = Math.min(currentPage * PAGE_SIZE, total);
  info.textContent = start + '–' + end + ' sur ' + total;

  var prev = document.createElement('button');
  prev.className = 'page-btn'; prev.textContent = '‹';
  prev.disabled = currentPage === 1;
  prev.onclick = function() { currentPage--; renderTable(); };
  pg.appendChild(prev);

  for (var p = 1; p <= totalPages; p++) {
    if (totalPages > 7 && Math.abs(p - currentPage) > 2 && p !== 1 && p !== totalPages) {
      if (p === 2 || p === totalPages - 1) {
        var s = document.createElement('span');
        s.textContent = '…'; s.style.color = 'var(--muted)'; s.style.padding = '0 4px';
        pg.appendChild(s);
      }
      continue;
    }
    var btn = document.createElement('button');
    btn.className = 'page-btn' + (p === currentPage ? ' active' : '');
    btn.textContent = p;
    btn.onclick = (function(pp) { return function() { currentPage = pp; renderTable(); }; })(p);
    pg.appendChild(btn);
  }

  var next = document.createElement('button');
  next.className = 'page-btn'; next.textContent = '›';
  next.disabled = currentPage === totalPages;
  next.onclick = function() { currentPage++; renderTable(); };
  pg.appendChild(next);
}

// ═══════════════════════════════════════════════════════
//  RECHERCHE
// ═══════════════════════════════════════════════════════
function onSearch() {
  searchQuery = document.getElementById('search-input').value.trim();
  currentPage = 1;
  renderTable();
}

// ═══════════════════════════════════════════════════════
//  VALIDATION FORMAT CIN & TEL
// ═══════════════════════════════════════════════════════
function validateFormatChamps() {
  var erreurs = [];
  rows.forEach(function(r, i) {
    var num  = i + 1;
    var nom  = r.nom || '(sans nom)';
    var cin  = (r.cin   || '').trim();
    var tel  = (r.mvola || '').trim();

    if (!cin) {
      erreurs.push('Ligne ' + num + ' – ' + nom + ' : CIN manquant.');
    } else if (!REGEX_CIN.test(cin)) {
      erreurs.push('Ligne ' + num + ' – ' + nom + ' : CIN invalide "' + cin + '" (format attendu : XXX XXX XXX XXX).');
    }

    if (!tel) {
      erreurs.push('Ligne ' + num + ' – ' + nom + ' : Numéro M\'vola manquant.');
    } else if (!REGEX_TEL.test(tel)) {
      erreurs.push('Ligne ' + num + ' – ' + nom + ' : Numéro invalide "' + tel + '" (format attendu : XXX XX XXX XX).');
    }
  });
  return erreurs;
}

// ═══════════════════════════════════════════════════════
//  VALIDATION GÉNÉRALE (région, district, lignes vides)
// ═══════════════════════════════════════════════════════
function validate() {
  var ok       = true;
  var region   = document.getElementById('sel-region').value;
  var district = document.getElementById('sel-district').value;
  setErr('grp-region',   !region);   if (!region)   ok = false;
  setErr('grp-district', !district); if (!district) ok = false;
  if (rows.length === 0) {
    showToast('Le tableau est vide. Ajoutez au moins un acteur.', 'error');
    ok = false;
  }
  var missing = rows.filter(function(r) { return !r.nom.trim(); }).length;
  if (missing > 0) {
    showToast(missing + ' ligne(s) sans nom. Veuillez les compléter ou supprimer.', 'warn');
    ok = false;
  }
  return ok;
}

// ═══════════════════════════════════════════════════════
//  ENREGISTRER DANS GOOGLE SHEET
// ═══════════════════════════════════════════════════════
async function submitData() {
  if (!validate()) return;

  // Validation format CIN et M'vola — bloquant
  var erreursFormat = validateFormatChamps();
  if (erreursFormat.length > 0) {
    showToast('⚠️ ' + erreursFormat.length + ' erreur(s) de format. Corrigez avant d\'enregistrer.', 'error');
    var msg = 'Erreurs de format détectées :\n\n' +
      erreursFormat.slice(0, 10).join('\n') +
      (erreursFormat.length > 10 ? '\n\n… et ' + (erreursFormat.length - 10) + ' autre(s) erreur(s).' : '');
    alert(msg);
    return;
  }

  var region   = document.getElementById('sel-region').value;
  var district = document.getElementById('sel-district').value;

  var acteurs = rows
    .map(function(r) {
      return {
        csb:       r.csb.trim(),
        fokontany: r.fokontany.trim(),
        nom:       r.nom.trim(),
        poste:     r.poste.trim(),
        cin:       r.cin.trim(),
        mvola:     r.mvola.trim()
      };
    })
    .filter(function(r) { return r.nom; });

  var payload = { action: 'saveActeursCommunautaires', region: region, district: district, acteurs: acteurs };

  showToast('Envoi en cours…', '');
  try {
    var res    = await fetch(GS_URL, { method:'POST', headers:{'Content-Type':'text/plain;charset=utf-8'}, body: JSON.stringify(payload) });
    var result = await res.json();
    if (result.status === 'ok') {
      showToast('✓ ' + acteurs.length + ' acteur(s) enregistré(s) dans Google Sheet !', 'success');
    } else {
      showToast(result.message || 'Erreur Apps Script', 'error');
    }
  } catch (err) {
    console.error(err);
    showToast('Erreur de connexion avec Google Sheet', 'error');
  }
}

// ═══════════════════════════════════════════════════════
//  CHARGER DEPUIS GOOGLE SHEET
// ═══════════════════════════════════════════════════════
async function loadData() {
  var region   = document.getElementById('sel-region').value;
  var district = document.getElementById('sel-district').value;
  if (!region || !district) { showToast('Sélectionnez une région et un district.', 'error'); return; }

  showToast('Chargement…', '');
  try {
    var url  = GS_URL + '?action=getActeursCommunautaires&region=' + encodeURIComponent(region) + '&district=' + encodeURIComponent(district);
    var res  = await fetch(url);
    var data = await res.json();
    if (data.status === 'ok' && data.acteurs && data.acteurs.length > 0) {
      rows = data.acteurs.map(function(a) {
        return { _id: ++rowIdCounter, csb: a.csb||'', fokontany: a.fokontany||'', nom: a.nom||'', poste: a.poste||'', cin: a.cin||'', mvola: a.mvola||'' };
      });
      currentPage = 1;
      renderTable();
      showToast('✓ ' + rows.length + ' acteur(s) chargé(s)', 'success');
    } else {
      showToast(data.message || 'Aucune donnée trouvée pour ce district.', 'warn');
    }
  } catch (err) {
    console.error(err);
    showToast('Erreur de connexion avec Google Sheet', 'error');
  }
}

// ═══════════════════════════════════════════════════════
//  EXPORTER CSV
// ═══════════════════════════════════════════════════════
function exportCSV() {
  if (rows.length === 0) { showToast('Aucune donnée à exporter.', 'warn'); return; }
  var region   = document.getElementById('sel-region').value   || 'REGION';
  var district = document.getElementById('sel-district').value || 'DISTRICT';
  var headers  = ['Région','District','CSB','Fokontany','Nom et Prénom','Poste','CIN',"Num M'vola"];
  var lines    = [headers.join(',')];
  rows.forEach(function(r) {
    lines.push([region, district, r.csb, r.fokontany, r.nom, r.poste, r.cin, r.mvola]
      .map(function(v) { return '"' + String(v||'').replace(/"/g,'""') + '"'; })
      .join(','));
  });
  var blob = new Blob(['\uFEFF' + lines.join('\n')], { type:'text/csv;charset=utf-8;' });
  var a    = document.createElement('a');
  a.href   = URL.createObjectURL(blob);
  a.download = 'ActeursCommunautaires_' + district + '_' + new Date().toISOString().slice(0,10) + '.csv';
  a.click();
  showToast('✓ Export CSV téléchargé', 'success');
}

// ═══════════════════════════════════════════════════════
//  RÉINITIALISER
// ═══════════════════════════════════════════════════════
function resetAll() {
  if (!confirm('Réinitialiser tout le formulaire ?')) return;
  rows = []; rowIdCounter = 0; currentPage = 1; searchQuery = '';
  document.getElementById('sel-region').value = '';
  document.getElementById('sel-district').innerHTML = '<option value="">— Sélectionner un district —</option>';
  document.getElementById('sel-district').disabled = true;
  document.getElementById('search-input').value = '';
  renderTable();
}

// ═══════════════════════════════════════════════════════
//  PROGRESSION IMPORT
// ═══════════════════════════════════════════════════════
function showProgress(visible, msg, pct) {
  var el = document.getElementById('upload-progress');
  if (!visible) { el.classList.remove('visible'); return; }
  el.classList.add('visible');
  document.getElementById('upload-msg').textContent  = msg || '';
  document.getElementById('progress-bar').style.width = (pct || 0) + '%';
}

// ═══════════════════════════════════════════════════════
//  UTILITAIRES
// ═══════════════════════════════════════════════════════
function setErr(groupId, hasError) {
  var g = document.getElementById(groupId);
  if (g) g.classList.toggle('has-error', hasError);
}

function showToast(msg, type) {
  var t = document.getElementById('toast');
  document.getElementById('toast-msg').textContent = msg;
  t.className = type ? 'show ' + type : 'show';
  clearTimeout(t._timer);
  t._timer = setTimeout(function() { t.className = type || ''; }, 4000);
}

function onlyDigits(event) {
  if (!/[0-9]/.test(event.key) && !['Backspace','Delete','ArrowLeft','ArrowRight','Tab'].includes(event.key)) {
    event.preventDefault();
  }
}

function formatCIN(input) {
  var v = input.value.replace(/\D/g, '').substring(0, 12);
  input.value = v.replace(/(\d{0,3})(\d{0,3})(\d{0,3})(\d{0,3})/,
    function(_, a, b, c, d) { return [a,b,c,d].filter(Boolean).join(' '); });
}

function formatMobileMoney(input) {
  var v   = input.value.replace(/\D/g, '').substring(0, 10);
  var res = '';
  if (v.length > 0) res += v.substring(0, 3);
  if (v.length > 3) res += ' ' + v.substring(3, 5);
  if (v.length > 5) res += ' ' + v.substring(5, 8);
  if (v.length > 8) res += ' ' + v.substring(8, 10);
  input.value = res;
}

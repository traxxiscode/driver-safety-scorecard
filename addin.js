/* =========================================================
   Baseline Driver Safety Scorecard — addin.js
   Firestore structure:
     Collection : driver_safety_scorecard
     Doc fields : database_name, ruleConfig, ruleWeights,
                  schedule (emails, enabled, freq, dataRange,
                  time, start), created_at, updated_at
   ========================================================= */

/* ── Firebase / Firestore ───────────────────────────────── */
var _db             = null;
var _docRef         = null;   // reference to this database's config doc
var _currentDatabase = null;

/**
 * Initialize Firestore — called once we have the database name from the
 * Geotab session.  Mirrors the HOS alerter: sign in anonymously, then
 * look up (or create) the config document for this database.
 */
function initFirestore(databaseName, cb) {
  _currentDatabase = databaseName;

  var config = {
    apiKey:            'AIzaSyCOMWmaflsbq2rqulJK11mbf_zqNrPH2Qc',
    authDomain:        'driver-safety-scorecard.firebaseapp.com',
    projectId:         'driver-safety-scorecard',
    storageBucket:     'driver-safety-scorecard.firebasestorage.app',
    messagingSenderId: '256203757490',
    appId:             '1:256203757490:web:27b0edc739a32b5bc7f6ab'
  };

  // Use the same multi-app-safe init pattern as the HOS alerter
  if (!firebase.apps.length) {
    firebase.initializeApp(config);
  }
  _db = firebase.firestore();

  // Anonymous auth — required for Firestore security rules
  var authCheck = new Promise(function (resolve, reject) {
    firebase.auth().onAuthStateChanged(function (user) {
      if (user) {
        resolve(user);
      } else {
        firebase.auth().signInAnonymously().then(resolve).catch(reject);
      }
    });
  });

  authCheck
    .then(function () {
      // Ensure a config document exists for this database
      return _db.collection('driver_safety_scorecard')
        .where('database_name', '==', _currentDatabase)
        .get();
    })
    .then(function (snap) {
      if (!snap.empty) {
        _docRef = snap.docs[0].ref;
        console.log('[Scorecard] Config doc found for', _currentDatabase);
      } else {
        return _db.collection('driver_safety_scorecard').add({
          database_name: _currentDatabase,
          ruleConfig:    {},
          ruleWeights:   {},
          schedule:      { enabled: false, emails: [] },
          created_at:    firebase.firestore.FieldValue.serverTimestamp(),
          updated_at:    firebase.firestore.FieldValue.serverTimestamp()
        }).then(function (ref) {
          _docRef = ref;
          console.log('[Scorecard] Created new config doc for', _currentDatabase);
        });
      }
    })
    .then(function () {
      if (typeof cb === 'function') cb(null);
    })
    .catch(function (e) {
      console.error('[Scorecard] Firestore init failed:', e);
      if (typeof cb === 'function') cb(e);
    });
}

/* Load the config doc and pass its data to cb(data) */
function fsLoad(cb) {
  if (!_docRef) { cb({}); return; }
  _docRef.get()
    .then(function (snap) { cb(snap.exists ? snap.data() : {}); })
    .catch(function (e)   { console.error('[Scorecard] Load failed:', e); cb({}); });
}

/* Merge-update the config doc with the provided fields */
function fsSave(fields, cb) {
  if (!_docRef) {
    if (typeof cb === 'function') cb(new Error('No Firestore doc reference'));
    return;
  }
  var payload = Object.assign({}, fields, {
    updated_at: firebase.firestore.FieldValue.serverTimestamp()
  });
  _docRef.update(payload)
    .then(function ()  { if (typeof cb === 'function') cb(null); })
    .catch(function (e) { console.error('[Scorecard] Save failed:', e); if (typeof cb === 'function') cb(e); });
}

/* ── Main Dashboard ─────────────────────────────────────── */
var safetyDash = (function () {
  var _api = null;
  var _rows = [], _rawData = [], _allRules = [], _ruleConfig = {}, _ruleWeights = {};
  var _days = 30, _sortKey = 'score', _sortDir = 1, _groupFilter = null, _isLight = false;
  var _allGroups = [], _selectedGroups = [], _driverGroups = {};
  var MAX_RULES = 6;
  var _schedEmails = [], _schedEnabled = false;
  var _dmap = {};
  var _isFirstRun = false;

  /* ── Helpers ── */
  function setDateRange() {
    var now = new Date(), from = new Date(now);
    from.setDate(from.getDate() - _days);
    var fmt = function (d) { return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }); };
    document.getElementById('dateRange').textContent = fmt(from) + ' – ' + fmt(now);
  }

  function showErr(m)  { document.getElementById('errBox').innerHTML = '<div class="err-box">&#9888; ' + m + '</div>'; }
  function clearMsg()  { document.getElementById('errBox').innerHTML = ''; document.getElementById('warnBox').innerHTML = ''; }

  function showBox(label, pct) {
    document.getElementById('tbl').innerHTML =
      '<div class="box"><div class="spinner"></div><div class="msg-txt">' + label + '</div>' +
      (pct !== undefined ? '<div class="pbar-bg"><div class="pbar-fill" style="width:' + pct + '%"></div></div>' : '') +
      '</div>';
  }

  function resetKPIs() {
    ['k1','k2','k3','k4','k5','k6'].forEach(function (id) { document.getElementById(id).textContent = '–'; });
  }

  function toast(msg, color) {
    var t = document.getElementById('toast');
    t.style.background = color || '#10b981';
    t.textContent = msg;
    t.classList.add('show');
    setTimeout(function () { t.classList.remove('show'); }, 3000);
  }

  function guessCategory(ruleName) {
    var lo = (ruleName || '').toLowerCase();
    var KEYWORDS = { speeding: ['speed'], braking: ['brak'], acceleration: ['accel'], cornering: ['corner'], seatbelt: ['seat','belt'] };
    var ks = Object.keys(KEYWORDS);
    for (var i = 0; i < ks.length; i++) {
      var kws = KEYWORDS[ks[i]];
      for (var j = 0; j < kws.length; j++) { if (lo.indexOf(kws[j]) !== -1) return ks[i]; }
    }
    return 'other';
  }

  function buildDriverName(u) {
    if (!u) return null;
    var fn = (u.firstName || '').trim(), ln = (u.lastName || '').trim();
    if (fn || ln) return (fn + ' ' + ln).trim();
    if (u.name && u.name.trim()) return u.name.trim();
    return null;
  }

  function scoreClass(s) { return s >= 80 ? 'score-hi' : s >= 60 ? 'score-med' : 'score-low'; }

  function gradeFromScore(s) {
    if (s >= 90) return { l: 'A', c: 'grade-a' };
    if (s >= 80) return { l: 'B', c: 'grade-b' };
    if (s >= 70) return { l: 'C', c: 'grade-c' };
    if (s >= 60) return { l: 'D', c: 'grade-d' };
    return { l: 'F', c: 'grade-f' };
  }

  function evClass(n) { return n === 0 ? 'ev-0' : n <= 5 ? 'ev-few' : n <= 20 ? 'ev-mid' : 'ev-many'; }

  var CATS = ['speeding','braking','acceleration','cornering','seatbelt','other'];
  var CAT_LABELS  = { speeding:'Speeding', braking:'Hard Braking', acceleration:'Hard Acceleration', cornering:'Harsh Cornering', seatbelt:'Seatbelt', other:'Uncategorized / Custom' };
  var CAT_COLORS  = { speeding:'#ef4444', braking:'#FF7B01', acceleration:'#fb923c', cornering:'#3a6bb5', seatbelt:'#00e5ff', other:'#4d6d96' };

  /* ── Weight System ── */
  function getEnabledRules() {
    return Object.keys(_ruleConfig).filter(function (rid) { return _ruleConfig[rid] && _ruleConfig[rid].enabled; });
  }
  function getWeight(rid) { return (_ruleWeights[rid] !== undefined) ? _ruleWeights[rid] : 0; }
  function getTotalWeight() { return getEnabledRules().reduce(function (s, rid) { return s + getWeight(rid); }, 0); }

  function _sidToRid(sid) {
    var enabled = getEnabledRules();
    for (var i = 0; i < enabled.length; i++) {
      if (enabled[i].replace(/[^a-zA-Z0-9]/g, '_') === sid) return enabled[i];
    }
    return null;
  }

  function equalizeWeightsData() {
    var enabled = getEnabledRules();
    if (!enabled.length) return;
    var base = Math.floor(100 / enabled.length);
    var rem  = 100 - (base * enabled.length);
    enabled.forEach(function (rid, i) { _ruleWeights[rid] = base + (i < rem ? 1 : 0); });
  }

  function redistributeAfterEdit(changedRid, newVal) {
    var enabled = getEnabledRules();
    if (enabled.length <= 1) { _ruleWeights[changedRid] = 100; return; }
    newVal = Math.max(0, Math.min(100, newVal));
    _ruleWeights[changedRid] = newVal;
    var rest = enabled.filter(function (r) { return r !== changedRid; });
    var remaining = Math.max(0, 100 - newVal);
    var oldOthersTotal = rest.reduce(function (s, r) { return s + getWeight(r); }, 0);
    if (oldOthersTotal === 0) {
      var share = Math.floor(remaining / rest.length), extra = remaining - (share * rest.length);
      rest.forEach(function (r, i) { _ruleWeights[r] = share + (i < extra ? 1 : 0); });
    } else {
      var assigned = 0;
      rest.forEach(function (r, i) {
        if (i === rest.length - 1) { _ruleWeights[r] = remaining - assigned; }
        else { var prop = Math.round((getWeight(r) / oldOthersTotal) * remaining); _ruleWeights[r] = Math.max(0, prop); assigned += _ruleWeights[r]; }
      });
    }
    var total = getTotalWeight();
    if (total !== 100) {
      var diff = 100 - total;
      for (var i = 0; i < rest.length; i++) {
        if (_ruleWeights[rest[i]] + diff >= 0) { _ruleWeights[rest[i]] += diff; break; }
      }
    }
  }

  function calcScore(ebr, norm) {
    if (!norm) return 100;
    var enabled = getEnabledRules();
    if (!enabled.length) return 100;
    var totalDeduction = 0;
    enabled.forEach(function (rid) {
      var cnt      = ebr[rid] || 0;
      var wt       = getWeight(rid);
      var perEvent = (_ruleConfig[rid] && _ruleConfig[rid].perEvent) || 3;
      var rate     = (cnt / norm) * 100;
      totalDeduction += Math.min(wt, rate * perEvent);
    });
    return Math.max(0, Math.round(100 - totalDeduction));
  }

  /* ── Rule Config ── */
  function initRuleConfig(rules, savedCfg, savedWts) {
    _allRules = rules;
    _ruleConfig = {};
    rules.forEach(function (rule) {
      var cat = guessCategory(rule.name);
      if (savedCfg && savedCfg[rule.id]) {
        _ruleConfig[rule.id] = { enabled: !!savedCfg[rule.id].enabled, category: savedCfg[rule.id].category || cat, perEvent: savedCfg[rule.id].perEvent || 3 };
      } else {
        _ruleConfig[rule.id] = { enabled: false, category: cat, perEvent: 3 };
      }
    });
    if (savedWts) {
      _ruleWeights = savedWts;
      var enabled = getEnabledRules();
      var savedTotal = enabled.reduce(function (s, r) { return s + (savedWts[r] || 0); }, 0);
      if (Math.abs(savedTotal - 100) > 2) equalizeWeightsData();
    } else {
      _ruleWeights = {};
      equalizeWeightsData();
    }
  }

  function onRuleToggled() {
    equalizeWeightsData();
    fsSave({ ruleConfig: _ruleConfig });
    if (document.getElementById('panelWeights').classList.contains('open')) renderWeightsPanel();
  }

  /* ── Rebuild / Render ── */
  function rebuildRows() {
    _rows = _rawData.map(function (d) {
      var score = calcScore(d.ebr, d.norm);
      var row   = { dname: d.dname, score: score, norm: d.norm, displayVal: d.displayVal, tripCount: d.tripCount, ebr: d.ebr, did: d.did };
      CATS.forEach(function (cat) {
        row['cat_' + cat] = Object.keys(_ruleConfig)
          .filter(function (rid) { return _ruleConfig[rid].enabled && _ruleConfig[rid].category === cat; })
          .reduce(function (s, rid) { return s + (d.ebr[rid] || 0); }, 0);
      });
      Object.keys(d.ebr).forEach(function (rid) { row['rule_' + rid] = d.ebr[rid]; });
      return row;
    });
    _rows.sort(function (a, b) { return a.score - b.score; });
    _sortKey = 'score'; _sortDir = 1;
    renderKPIs(_rows);
    updateBadge();
    renderTable(getDisplayRows());
    renderRulesPanel();
    renderWeightsPanel();
  }

  function renderKPIs(rows) {
    var avg = rows.length ? Math.round(rows.reduce(function (s, r) { return s + r.score; }, 0) / rows.length) : 0;
    var spd = 0, brk = 0;
    rows.forEach(function (r) {
      Object.keys(_ruleConfig).forEach(function (rid) {
        var cfg = _ruleConfig[rid]; if (!cfg.enabled) return;
        if (cfg.category === 'speeding') spd += (r.ebr[rid] || 0);
        if (cfg.category === 'braking')  brk += (r.ebr[rid] || 0);
      });
    });
    document.getElementById('k1').textContent = avg;
    document.getElementById('k2').textContent = rows.filter(function (r) { return r.score >= 80; }).length;
    document.getElementById('k3').textContent = rows.filter(function (r) { return r.score < 60; }).length;
    document.getElementById('k4').textContent = spd.toLocaleString();
    document.getElementById('k5').textContent = brk.toLocaleString();
    document.getElementById('k6').textContent = rows.length;
    var totalMiles = rows.reduce(function (s, r) { return s + (r.displayVal || 0); }, 0);
    document.getElementById('k4sub').textContent = totalMiles.toLocaleString() + ' mi total';
    document.getElementById('foot').textContent = 'Last refreshed: ' + new Date().toLocaleString() + '  ·  Last ' + _days + ' days  ·  Score = 100 minus weighted deductions (per 100 mi)';
  }

  function updateBadge() {
    var badge    = document.getElementById('filterBadge');
    var topCard  = document.getElementById('kpi-top');
    var riskCard = document.getElementById('kpi-risk');
    topCard.classList.remove('kpi-active','kpi-active-green');
    riskCard.classList.remove('kpi-active','kpi-active-red');
    if (!_groupFilter) { badge.style.display = 'none'; return; }
    var isTop  = _groupFilter === 'top';
    var color  = isTop ? '#10b981' : '#ef4444';
    var label  = isTop ? 'Top Performers (score 80+)' : 'At-Risk Drivers (score below 60)';
    if (isTop) topCard.classList.add('kpi-active','kpi-active-green');
    else       riskCard.classList.add('kpi-active','kpi-active-red');
    badge.style.display = 'inline-flex';
    badge.innerHTML = '<span class="filter-badge"><span class="filter-dot" style="background:' + color + '"></span>Filtered: ' + label + '<span class="filter-x" onclick="safetyDash.clearFilter()">&#x2715;</span></span>';
  }

  function getDisplayRows() {
    var rows = _rows;
    if (_groupFilter === 'top')  rows = rows.filter(function (r) { return r.score >= 80; });
    if (_groupFilter === 'risk') rows = rows.filter(function (r) { return r.score < 60; });
    if (_selectedGroups.length) {
      rows = rows.filter(function (r) {
        var dg = _driverGroups[r.did] || [];
        return _selectedGroups.some(function (gid) { return dg.indexOf(gid) !== -1; });
      });
    }
    var q = document.getElementById('srch').value.toLowerCase();
    if (q) rows = rows.filter(function (r) { return r.dname.toLowerCase().indexOf(q) !== -1; });
    return rows;
  }

  function renderTable(rows) {
    if (!rows || !rows.length) {
      document.getElementById('tbl').innerHTML = '<div class="box"><div class="msg-txt">' + (_groupFilter ? 'No drivers match this filter.' : 'No driver data found.') + '</div></div>';
      return;
    }
    var enabled = Object.keys(_ruleConfig).filter(function (rid) { return _ruleConfig[rid] && _ruleConfig[rid].enabled; });
    var useCat  = enabled.length > 8;
    var colCats = CATS.filter(function (c) { return enabled.some(function (rid) { return _ruleConfig[rid].category === c; }); });

    function thSort(k, label, cls) {
      var arrow = '';
      if (_sortKey === k) arrow = ' <em class="sort-arrow">' + (_sortDir === 1 ? '&#9650;' : '&#9660;') + '</em>';
      var c = cls ? ' class="' + cls + '"' : '';
      return '<th' + c + ' onclick="safetyDash.sort(\'' + k + '\')">' + label + arrow + '</th>';
    }

    var h = '<table><thead><tr>' +
      thSort('dname',      'Driver', 'th-driver') +
      thSort('score',      'Score',  'th-score') +
      '<th class="th-grade">Grade</th>' +
      thSort('displayVal', 'Miles',  'th-trips');

    if (useCat) {
      colCats.forEach(function (cat) { h += thSort('cat_' + cat, CAT_LABELS[cat], 'th-rule'); });
    } else {
      enabled.forEach(function (rid) {
        var rule = null;
        for (var i = 0; i < _allRules.length; i++) { if (_allRules[i].id === rid) { rule = _allRules[i]; break; } }
        var nm = rule ? rule.name : rid;
        var short = nm.length > 18 ? nm.substring(0, 16) + '..' : nm;
        h += thSort('rule_' + rid, short, 'th-rule');
      });
    }
    h += '</tr></thead><tbody>';

    rows.forEach(function (r) {
      var sc = scoreClass(r.score), gr = gradeFromScore(r.score);
      h += '<tr><td class="dname td-driver">' + r.dname + '</td>' +
        '<td><div class="score-wrap ' + sc + '"><span class="score-num">' + r.score + '</span>' +
        '<div class="score-bar-bg"><div class="score-bar-fill" style="width:' + r.score + '%"></div></div></div></td>' +
        '<td><span class="grade ' + gr.c + '">' + gr.l + '</span></td>' +
        '<td class="trips-cell">' + (r.displayVal || 0).toLocaleString() + '</td>';
      if (useCat) {
        colCats.forEach(function (cat) {
          var total = enabled.filter(function (rid) { return _ruleConfig[rid].category === cat; })
            .reduce(function (s, rid) { return s + (r.ebr[rid] || 0); }, 0);
          h += '<td><div class="ev-cell ' + evClass(total) + '"><span class="ev-dot"></span>' + total + '</div></td>';
        });
      } else {
        enabled.forEach(function (rid) {
          var n = r.ebr[rid] || 0;
          h += '<td><div class="ev-cell ' + evClass(n) + '"><span class="ev-dot"></span>' + n + '</div></td>';
        });
      }
      h += '</tr>';
    });
    document.getElementById('tbl').innerHTML = h + '</tbody></table>';
  }

  /* ── Weights Panel ── */
  function updateWeightTotalBar() {
    var total  = getTotalWeight();
    var numEl  = document.getElementById('wtTotalNum');
    var fillEl = document.getElementById('wtTotalFill');
    var hintEl = document.getElementById('wtTotalHint');
    if (!numEl) return;
    numEl.textContent = total;
    numEl.className   = 'wt-total-num ' + (total === 100 ? 'wt-total-ok' : total > 100 ? 'wt-total-over' : 'wt-total-under');
    fillEl.style.width      = Math.min(100, total) + '%';
    fillEl.style.background = total === 100 ? '#10b981' : total > 100 ? '#ef4444' : '#FF7B01';
    hintEl.textContent = total === 100 ? '✓ Balanced' : total > 100 ? '⚠ Over by ' + (total - 100) + '%' : '⚠ Under by ' + (100 - total) + '%';
    hintEl.style.color = total === 100 ? '#10b981' : total > 100 ? '#ef4444' : '#FF7B01';
  }

  function syncWeightUI(rid, sid, newVal) {
    var slEl = document.getElementById('wsl_' + sid);
    var piEl = document.getElementById('wpi_' + sid);
    var sbEl = document.getElementById('wsb_' + sid);
    var perE = (_ruleConfig[rid] && _ruleConfig[rid].perEvent) || 3;
    if (slEl) slEl.value = newVal;
    if (piEl) { piEl.value = newVal; piEl.classList.toggle('input-err', newVal < 0 || newVal > 100); }
    if (sbEl) sbEl.textContent = newVal + '% weight · ' + perE + ' pts/incident';
    getEnabledRules().forEach(function (r) {
      if (r === rid) return;
      var s2 = r.replace(/[^a-zA-Z0-9]/g, '_'), v = getWeight(r), pe2 = (_ruleConfig[r] && _ruleConfig[r].perEvent) || 3;
      var sl2 = document.getElementById('wsl_' + s2), pi2 = document.getElementById('wpi_' + s2), sb2 = document.getElementById('wsb_' + s2);
      if (sl2) sl2.value = v;
      if (pi2) pi2.value = v;
      if (sb2) sb2.textContent = v + '% weight · ' + pe2 + ' pts/incident';
    });
    updateWeightTotalBar();
  }

  function renderWeightsPanel() {
    var enabled = getEnabledRules();
    var wc = document.getElementById('weightsCount');
    if (wc) wc.textContent = enabled.length + ' rule' + (enabled.length === 1 ? '' : 's');
    updateWeightTotalBar();
    var grid = document.getElementById('weightGrid');
    if (!grid) return;
    if (!enabled.length) { grid.innerHTML = '<div class="weight-empty">No rules enabled. Enable rules in the Rules panel first.</div>'; return; }
    var h = '';
    enabled.forEach(function (rid) {
      var rule = null;
      for (var i = 0; i < _allRules.length; i++) { if (_allRules[i].id === rid) { rule = _allRules[i]; break; } }
      var rname = rule ? rule.name : rid, wt = getWeight(rid), perE = (_ruleConfig[rid] && _ruleConfig[rid].perEvent) || 3;
      var sid   = rid.replace(/[^a-zA-Z0-9]/g, '_');
      h += '<div class="weight-item">' +
        '<div class="weight-label" title="' + rname + '">' + rname + '</div>' +
        '<div class="weight-input-wrap">' +
          '<input type="range" class="weight-slider" id="wsl_' + sid + '" data-rid="' + sid + '" min="0" max="100" step="1" value="' + wt + '" oninput="safetyDash.onWeightSlide(this)"/>' +
          '<input type="number" class="weight-pct-input" id="wpi_' + sid + '" data-rid="' + sid + '" min="0" max="100" value="' + wt + '" oninput="safetyDash.onWeightInput(this)"/>' +
          '<span style="font-size:.75rem;color:var(--text3);">%</span>' +
        '</div>' +
        '<div class="weight-sub" id="wsb_' + sid + '">' + wt + '% weight · ' + perE + ' pts/incident</div>' +
      '</div>';
    });
    grid.innerHTML = h;
  }

  /* ── Rules Panel ── */
  function renderRulesPanel() {
    if (!_allRules.length) return;
    var enabledCount = Object.keys(_ruleConfig).filter(function (rid) { return _ruleConfig[rid] && _ruleConfig[rid].enabled; }).length;
    document.getElementById('ruleCount').textContent    = _allRules.length + ' rules available';
    document.getElementById('rulesEnabled').textContent = enabledCount + ' / ' + MAX_RULES + ' rules enabled';
    var bycat = {};
    CATS.forEach(function (c) { bycat[c] = []; });
    _allRules.forEach(function (rule) {
      var cat = _ruleConfig[rule.id] ? _ruleConfig[rule.id].category : guessCategory(rule.name);
      if (!bycat[cat]) bycat[cat] = [];
      bycat[cat].push(rule);
    });
    var h = '<div class="rules-cats">';
    CATS.forEach(function (cat) {
      var rules = bycat[cat] || [];
      if (!rules.length) return;
      h += '<div class="rule-cat-block">' +
        '<div class="rule-cat-title"><span class="rule-cat-dot" style="background:' + CAT_COLORS[cat] + '"></span>' + CAT_LABELS[cat] + ' (' + rules.length + ')</div>' +
        '<div class="rules-grid">';
      rules.forEach(function (rule) {
        var cfg = _ruleConfig[rule.id] || { enabled: false, category: cat, perEvent: 3 };
        var on  = cfg.enabled;
        var ptOpts = [1,2,3,4,5,7,10].map(function (v) { return '<option value="' + v + '"' + (cfg.perEvent === v ? ' selected' : '') + '>' + v + 'pt</option>'; }).join('');
        h += '<div class="rule-row' + (on ? ' rule-on' : '') + '" id="rr_' + rule.id + '" onclick="safetyDash.toggleRule(\'' + rule.id + '\')">' +
          '<div class="rule-chk"><svg viewBox="0 0 12 10" fill="none" stroke="#fff" stroke-width="2.5"><polyline points="1,5 4.5,8.5 11,1"/></svg></div>' +
          '<div class="rule-info"><div class="rule-nm" title="' + rule.name + '">' + rule.name + '</div><div class="rule-meta">' + CAT_LABELS[cat] + '</div></div>' +
          '<select class="rule-wt" onclick="event.stopPropagation()" onchange="safetyDash.setRuleWt(\'' + rule.id + '\',this.value)">' + ptOpts + '</select>' +
        '</div>';
      });
      h += '</div></div>';
    });
    h += '</div>';
    document.getElementById('rulesList').innerHTML = h;
  }

  /* ── Groups ── */
  function initGroups(groups) {
    _allGroups = groups.filter(function (g) {
      return g.name && g.name !== 'GroupCompanyId' && g.name !== 'GroupNothingId' && g.name !== 'GroupRootId';
    });
    renderGroupDropdown();
  }

  function renderGroupDropdown() {
    var list = document.getElementById('grpList');
    if (!list) return;
    if (!_allGroups.length) {
      list.innerHTML = '<div style="padding:8px 14px;font-size:.75rem;color:var(--text3);">No groups found</div>';
    } else {
      var items = _allGroups.map(function (g) {
        var sel = _selectedGroups.indexOf(g.id) !== -1;
        return '<div class="grp-item' + (sel ? ' selected' : '') + '" data-gid="' + g.id + '" onclick="safetyDash.toggleGroup(this.dataset.gid)">' +
          '<div class="grp-chk"><svg viewBox="0 0 12 10" fill="none" stroke="#fff" stroke-width="2.5"><polyline points="1,5 4.5,8.5 11,1"/></svg></div>' +
          '<span>' + g.name + '</span></div>';
      });
      list.innerHTML = items.join('');
    }
    var badge = document.getElementById('grpBadge');
    var btn   = document.getElementById('grpBtn');
    if (badge) { badge.textContent = _selectedGroups.length; badge.style.display = _selectedGroups.length ? '' : 'none'; }
    if (btn)   btn.classList.toggle('active-btn', _selectedGroups.length > 0);
    var allItem = document.getElementById('grpAll');
    if (allItem) allItem.classList.toggle('selected', _selectedGroups.length === 0);
  }

  function loadDriverGroups(dids) {
    if (!_api || !dids.length) return;
    _api.call('Get', { typeName: 'User', search: { userIds: dids }, resultsLimit: 5000 }, function (users) {
      if (!users) return;
      users.forEach(function (u) {
        if (u.id && u.companyGroups) _driverGroups[u.id] = u.companyGroups.map(function (g) { return g.id; });
      });
    }, function () {});
  }

  /* ── Schedule ── */
  function loadScheduleUI(schedData) {
    if (!schedData) return;
    _schedEmails  = schedData.emails  || [];
    _schedEnabled = schedData.enabled || false;
    if (schedData.freq)      document.getElementById('schedFreq').value      = schedData.freq;
    if (schedData.dataRange) document.getElementById('schedDataRange').value = schedData.dataRange;
    if (schedData.time)      document.getElementById('schedTime').value      = schedData.time;
    if (schedData.start)     document.getElementById('schedStart').value     = schedData.start;
    var wrap = document.getElementById('schedChips');
    wrap.querySelectorAll('.schip').forEach(function (c) { c.remove(); });
    _schedEmails.forEach(function (e) { addSchedChipEl(e); });
    updateSchedStatus();
  }

  function updateSchedStatus() {
    var el     = document.getElementById('schedStatus');
    var disBtn = document.getElementById('btnSchedDisable');
    if (!el) return;
    if (_schedEnabled && _schedEmails.length) {
      var freqEl  = document.getElementById('schedFreq');
      var freqTxt = freqEl ? freqEl.options[freqEl.selectedIndex].text : 'Weekly';
      var rangeEl = document.getElementById('schedDataRange');
      var rangeTxt= rangeEl ? rangeEl.options[rangeEl.selectedIndex].text : '';
      el.className   = 'sched-status on';
      el.textContent = '✓ Active — sends ' + freqTxt.toLowerCase() + (rangeTxt ? ' · data: ' + rangeTxt : '') + ' · ' + _schedEmails.length + ' recipient' + (_schedEmails.length > 1 ? 's' : '');
      if (disBtn) disBtn.style.display = '';
    } else {
      el.className   = 'sched-status off';
      el.textContent = 'Schedule is inactive. Add recipients and save to activate.';
      if (disBtn) disBtn.style.display = 'none';
    }
  }

  function addSchedChipEl(email) {
    var wrap  = document.getElementById('schedChips');
    var input = document.getElementById('schedChipInput');
    var chip  = document.createElement('span');
    chip.className = 'schip';
    chip.setAttribute('data-email', email);
    chip.innerHTML = email + '<span class="schip-x" onclick="safetyDash.removeSchedChip(this)">&#x2715;</span>';
    wrap.insertBefore(chip, input);
  }

  /* ── First-run Setup Gate ── */
  function showSetupGate() {
    _isFirstRun = true;
    document.getElementById('tbl').innerHTML =
      '<div class="box setup-gate">' +
      '<div class="setup-gate-icon">&#9881;</div>' +
      '<div class="setup-gate-title">Welcome — First-Time Setup Required</div>' +
      '<div class="setup-gate-body">This Geotab database doesn\'t have a scoring configuration yet. ' +
      'Please select your active rules in the <strong>Rules</strong> panel and configure weights in the <strong>Weights</strong> panel, ' +
      'then save both before data will load.</div>' +
      '<div class="setup-gate-actions">' +
      '<button class="btn btn-apply" onclick="safetyDash.togglePanel(\'rules\')">&#10003; Open Rules</button>' +
      '<button class="btn btn-apply" onclick="safetyDash.togglePanel(\'weights\')" style="margin-left:8px;">&#9881; Open Weights</button>' +
      '</div>' +
      '</div>';
    renderRulesPanel();
    renderWeightsPanel();
  }

  function completeSetupIfReady() {
    if (!_isFirstRun) return true;
    var enabledRules = getEnabledRules();
    if (!enabledRules.length) { toast('Select at least one rule before continuing', '#ef4444'); return false; }
    if (getTotalWeight() !== 100) equalizeWeightsData();
    return true;
  }

  /* ── Geotab Data Fetch ── */
  function fetchData() {
    if (!_api) { showBox('API not ready.'); return; }
    clearMsg(); resetKPIs(); _groupFilter = null; updateBadge();
    showBox('FETCHING RULES…', 5);
    var now = new Date(), from = new Date(now);
    from.setDate(from.getDate() - _days);
    var fromStr = from.toISOString(), toStr = now.toISOString();

    _api.multiCall([
      ['Get', { typeName: 'Rule',  resultsLimit: 1000 }],
      ['Get', { typeName: 'User',  search: { isDriver: true }, resultsLimit: 1000 }],
      ['Get', { typeName: 'Group', resultsLimit: 500 }]
    ], function (res) {
      var rules  = (res && res[0]) || [];
      var users  = (res && res[1]) || [];
      var groups = (res && res[2]) || [];
      initGroups(groups);
      var dmap = {};
      users.forEach(function (u) { var dn = buildDriverName(u); if (dn && u.id) dmap[u.id] = dn; });
      _dmap = dmap;

      fsLoad(function (data) {
        var savedCfg = data.ruleConfig  || null;
        var savedWts = data.ruleWeights || null;
        initRuleConfig(rules, savedCfg, savedWts);
        // Treat an empty ruleConfig ({} with no enabled rules) the same as
        // no config — the doc is freshly created but never actually configured.
        var hasConfig = savedCfg && Object.keys(savedCfg).some(function (k) { return savedCfg[k] && savedCfg[k].enabled; });
        if (!hasConfig) {
          showSetupGate();
        } else {
          doFetch(dmap, fromStr, toStr);
        }
      });
    }, function () {
      // multiCall fallback — try with just rules
      _api.call('Get', { typeName: 'Rule', resultsLimit: 1000 }, function (rules) {
        fsLoad(function (data) {
          var savedCfg = data.ruleConfig  || null;
          var savedWts = data.ruleWeights || null;
          initRuleConfig(rules || [], savedCfg, savedWts);
          var hasConfig = savedCfg && Object.keys(savedCfg).some(function (k) { return savedCfg[k] && savedCfg[k].enabled; });
          if (!hasConfig) { showSetupGate(); } else { doFetch(_dmap, fromStr, toStr); }
        });
      }, function (e) { showBox(''); showErr('Error: ' + (e && e.message ? e.message : JSON.stringify(e))); });
    });
  }

  function doFetch(dmap, fromStr, toStr) {
    showBox('FETCHING DEVICES…', 20);
    _api.call('Get', { typeName: 'Device', resultsLimit: 1000 }, function (devices) {
      if (!devices || !devices.length) { showBox(''); showErr('No devices found.'); return; }
      showBox('FETCHING TRIPS FOR ' + devices.length + ' DEVICES…', 30);
      var cap = Math.min(devices.length, 150);
      var tripCalls = devices.slice(0, cap).map(function (d) {
        return ['Get', { typeName: 'Trip', search: { deviceSearch: { id: d.id }, fromDate: fromStr, toDate: toStr }, resultsLimit: 5000 }];
      });
      var excCall = ['Get', { typeName: 'ExceptionEvent', search: { fromDate: fromStr, toDate: toStr }, resultsLimit: 100000 }];
      _api.multiCall(tripCalls.concat([excCall]), function (results) {
        showBox('BUILDING SCORECARDS…', 80);
        var excs        = (results && results[results.length - 1]) || [];
        var tripResults = results ? results.slice(0, tripCalls.length) : [];
        var milesMap = {}, tripCount = {};
        devices.slice(0, cap).forEach(function (d, i) {
          var trips = (tripResults[i]) || [];
          trips.forEach(function (t) {
            var did = t.driver && t.driver.id;
            if (!did || did === 'UnknownDriverId') return;
            milesMap[did] = (milesMap[did] || 0) + ((parseFloat(t.distance) || 0) * 0.621371);
            tripCount[did] = (tripCount[did] || 0) + 1;
            if (!dmap[did]) { var dn = buildDriverName(t.driver); if (dn) dmap[did] = dn; }
          });
        });
        var ebrd = {};
        excs.forEach(function (e) {
          var did = e.driver && e.driver.id; if (!did || did === 'UnknownDriverId') return;
          var rid = e.rule   && e.rule.id;   if (!rid) return;
          if (!ebrd[did]) ebrd[did] = {};
          ebrd[did][rid] = (ebrd[did][rid] || 0) + 1;
          if (!dmap[did] && e.driver) { var dn = buildDriverName(e.driver); if (dn) dmap[did] = dn; }
        });
        var seen = {};
        Object.keys(milesMap).forEach(function (k) { seen[k] = 1; });
        Object.keys(ebrd).forEach(function (k) { seen[k] = 1; });
        _rawData = [];
        Object.keys(seen).forEach(function (did) {
          var mi = Math.round((milesMap[did] || 0) * 10) / 10;
          _rawData.push({ dname: dmap[did] || ('Driver ' + did), norm: mi || 1, displayVal: Math.round(mi), tripCount: tripCount[did] || 0, ebr: ebrd[did] || {}, did: did });
        });
        if (!_rawData.length) { showBox(''); showErr('No driver data found.'); return; }
        rebuildRows();
        var dids = _rawData.map(function (d) { return d.did; }).filter(Boolean);
        if (dids.length) loadDriverGroups(dids);
      }, function (err) { showBox(''); showErr('Error: ' + (err && err.message ? err.message : JSON.stringify(err))); });
    }, function (err) { showBox(''); showErr('Device fetch failed: ' + (err && err.message ? err.message : JSON.stringify(err))); });
  }

  /* ── Public API ── */
  return {

    /**
     * Called once we have both the api reference and the confirmed database
     * name from the Geotab session (done inside the focus lifecycle method).
     */
    init: function (api, databaseName) {
      _api = api;
      setDateRange();
      document.getElementById('btnRefresh').onclick = fetchData;
      document.getElementById('schedStart').value = new Date().toISOString().split('T')[0];

      // Initialise Firestore (auth + ensure doc), then load theme/schedule,
      // then kick off the data fetch — all in sequence so _docRef is set
      // before fsLoad is ever called.
      initFirestore(databaseName, function (err) {
        if (err) {
          showBox('');
          showErr('Could not connect to configuration service: ' + err.message);
          return;
        }
        fsLoad(function (data) {
          if (data.preferences && data.preferences.theme === 'light') {
            _isLight = true;
            document.body.classList.add('light');
            document.getElementById('themeLbl').textContent = 'LIGHT';
            updateLogos(true);
          }
          loadScheduleUI(data.schedule || null);
          fetchData();
        });
      });
    },

    fetch: fetchData,

    setRange: function (days) {
      _days = days;
      document.querySelectorAll('.range-btn').forEach(function (b) { b.classList.toggle('active', b.textContent === days + 'D'); });
      setDateRange(); fetchData();
    },

    toggleTheme: function () {
      _isLight = !_isLight;
      document.body.classList.toggle('light', _isLight);
      document.getElementById('themeLbl').textContent = _isLight ? 'LIGHT' : 'DARK';
      updateLogos(_isLight);
      fsSave({ preferences: { theme: _isLight ? 'light' : 'dark' } });
    },

    toggleInfo: function () { document.getElementById('infoModal').classList.toggle('open'); },

    togglePanel: function (which) {
      var ids  = { weights: 'panelWeights', rules: 'panelRules', schedule: 'panelSchedule' };
      var btns = { weights: 'btnWeights',   rules: 'btnRules',   schedule: 'btnSchedule' };
      Object.keys(ids).forEach(function (k) {
        if (k === which) {
          var isOpen = document.getElementById(ids[k]).classList.toggle('open');
          if (btns[k] && document.getElementById(btns[k])) document.getElementById(btns[k]).classList.toggle('active-btn', isOpen);
          if (isOpen && k === 'rules'    && _allRules.length) renderRulesPanel();
          if (isOpen && k === 'weights') renderWeightsPanel();
          if (isOpen && k === 'schedule') {
            fsLoad(function (data) { loadScheduleUI(data.schedule || null); });
          }
        } else {
          document.getElementById(ids[k]).classList.remove('open');
          if (btns[k] && document.getElementById(btns[k])) document.getElementById(btns[k]).classList.remove('active-btn');
        }
      });
    },

    toggleRule: function (rid) {
      if (!_ruleConfig[rid]) return;
      var currentlyEnabled = _ruleConfig[rid].enabled;
      if (!currentlyEnabled && getEnabledRules().length >= MAX_RULES) {
        var warn = document.getElementById('ruleLimitWarn');
        if (warn) { warn.classList.add('show'); setTimeout(function () { warn.classList.remove('show'); }, 3000); }
        return;
      }
      _ruleConfig[rid].enabled = !currentlyEnabled;
      var row = document.getElementById('rr_' + rid);
      if (row) row.classList.toggle('rule-on', _ruleConfig[rid].enabled);
      var cnt = getEnabledRules().length;
      var el  = document.getElementById('rulesEnabled');
      if (el) el.textContent = cnt + ' / ' + MAX_RULES + ' rules enabled';
      var atLimit = cnt >= MAX_RULES;
      document.querySelectorAll('.rule-row').forEach(function (r) {
        r.classList.toggle('rule-disabled', atLimit && !r.classList.contains('rule-on'));
      });
      onRuleToggled();
    },

    setRuleWt: function (rid, val) { if (_ruleConfig[rid]) _ruleConfig[rid].perEvent = parseInt(val, 10) || 1; },

    saveRules: function () {
      if (!completeSetupIfReady()) return;
      fsSave({ ruleConfig: _ruleConfig, ruleWeights: _ruleWeights }, function (err) {
        if (err) { toast('Save failed: ' + err.message, '#ef4444'); return; }
        if (_isFirstRun) {
          _isFirstRun = false;
          toast('Setup complete — loading data…', '#10b981');
          var now = new Date(), from = new Date(now);
          from.setDate(from.getDate() - _days);
          doFetch(_dmap, from.toISOString(), now.toISOString());
        } else {
          if (_rawData.length) rebuildRows();
          toast('Rules saved and applied', '#10b981');
        }
      });
    },

    resetRules: function () {
      fsSave({ ruleConfig: {} }, function () {
        if (_allRules.length) { initRuleConfig(_allRules, null, null); renderRulesPanel(); }
        if (_rawData.length) rebuildRows();
        toast('Rules reset to defaults', '#64748b');
      });
    },

    saveWeights: function () {
      if (getTotalWeight() !== 100) { equalizeWeightsData(); renderWeightsPanel(); }
      if (_isFirstRun && !getEnabledRules().length) { toast('Enable at least one rule in the Rules panel first', '#ef4444'); return; }
      fsSave({ ruleConfig: _ruleConfig, ruleWeights: _ruleWeights }, function (err) {
        if (err) { toast('Could not save: ' + err.message, '#ef4444'); return; }
        if (_isFirstRun) {
          _isFirstRun = false;
          toast('Setup complete — loading data…', '#10b981');
          var now = new Date(), from = new Date(now);
          from.setDate(from.getDate() - _days);
          doFetch(_dmap, from.toISOString(), now.toISOString());
        } else {
          toast('Weights saved (total: ' + getTotalWeight() + '%)', '#10b981');
          if (_rawData.length) rebuildRows();
        }
      });
    },

    resetWeights: function () {
      _ruleWeights = {};
      equalizeWeightsData();
      fsSave({ ruleWeights: _ruleWeights });
      renderWeightsPanel();
      if (_rawData.length) rebuildRows();
      toast('Reset to equal distribution', '#64748b');
    },

    sort: function (k) {
      if (_sortKey === k) { _sortDir *= -1; } else { _sortKey = k; _sortDir = (k === 'dname' || k === 'score') ? 1 : -1; }
      _rows.sort(function (a, b) {
        var av = a[k] !== undefined ? a[k] : 0, bv = b[k] !== undefined ? b[k] : 0;
        if (typeof av === 'string') return _sortDir * av.localeCompare(bv);
        return _sortDir * (av - bv);
      });
      renderTable(getDisplayRows());
    },

    filter: function () { renderTable(getDisplayRows()); },
    filterGroup: function (g) { _groupFilter = (_groupFilter === g) ? null : g; updateBadge(); renderTable(getDisplayRows()); },
    clearFilter:  function () { _groupFilter = null; updateBadge(); renderTable(getDisplayRows()); },

    toggleGroupDrop: function () {
      var dd = document.getElementById('grpDropdown');
      if (dd) dd.classList.toggle('open');
      if (dd && dd.classList.contains('open')) {
        var close = function (e) {
          if (!document.getElementById('grpWrap').contains(e.target)) { dd.classList.remove('open'); document.removeEventListener('click', close); }
        };
        setTimeout(function () { document.addEventListener('click', close); }, 0);
      }
    },

    toggleGroup:      function (gid) { var idx = _selectedGroups.indexOf(gid); if (idx === -1) _selectedGroups.push(gid); else _selectedGroups.splice(idx, 1); renderGroupDropdown(); renderTable(getDisplayRows()); renderKPIs(getDisplayRows()); },
    selectAllGroups:  function ()    { _selectedGroups = []; renderGroupDropdown(); renderTable(getDisplayRows()); renderKPIs(getDisplayRows()); },
    clearGroupFilter: function ()    { _selectedGroups = []; renderGroupDropdown(); renderTable(getDisplayRows()); renderKPIs(getDisplayRows()); document.getElementById('grpDropdown').classList.remove('open'); },

    onWeightSlide: function (el) {
      var sid = el.getAttribute('data-rid'), rid = _sidToRid(sid);
      if (!rid) return;
      redistributeAfterEdit(rid, parseInt(el.value, 10));
      syncWeightUI(rid, sid, getWeight(rid));
    },

    onWeightInput: function (el) {
      var sid = el.getAttribute('data-rid'), rid = _sidToRid(sid);
      if (!rid) return;
      var v = parseInt(el.value, 10);
      if (isNaN(v)) return;
      redistributeAfterEdit(rid, v);
      syncWeightUI(rid, sid, getWeight(rid));
    },

    equalizeWeights: function () {
      equalizeWeightsData(); renderWeightsPanel();
      toast('Weights equalized to ' + Math.round(100 / Math.max(1, getEnabledRules().length)) + '% each', '#FF7B01');
    },

    schedChipKey: function (e) {
      if (e.key === 'Enter' || e.key === ',') {
        e.preventDefault();
        var val = document.getElementById('schedChipInput').value.trim().replace(/,$/, '');
        if (val && val.indexOf('@') !== -1 && _schedEmails.indexOf(val) === -1) {
          _schedEmails.push(val);
          addSchedChipEl(val);
          document.getElementById('schedChipInput').value = '';
        }
      }
      if (e.key === 'Backspace' && !document.getElementById('schedChipInput').value && _schedEmails.length) {
        var chips = document.getElementById('schedChips').querySelectorAll('.schip');
        if (chips.length) {
          var last  = chips[chips.length - 1];
          var email = last.getAttribute('data-email');
          _schedEmails = _schedEmails.filter(function (x) { return x !== email; });
          last.remove();
        }
      }
    },

    removeSchedChip: function (xEl) {
      var chip  = xEl.parentElement;
      var email = chip.getAttribute('data-email');
      _schedEmails = _schedEmails.filter(function (x) { return x !== email; });
      chip.remove();
    },

    saveSchedule: function () {
      if (!_schedEmails.length) { toast('Add at least one recipient email first', '#ef4444'); return; }
      _schedEnabled = true;
      fsSave({
        schedule: {
          emails:        _schedEmails,
          enabled:       true,
          freq:          document.getElementById('schedFreq').value,
          dataRange:     document.getElementById('schedDataRange').value,
          time:          document.getElementById('schedTime').value,
          start:         document.getElementById('schedStart').value,
          database_name: _currentDatabase
        }
      }, function (err) {
        if (!err) {
          updateSchedStatus();
          var freqEl  = document.getElementById('schedFreq');
          var freqTxt = freqEl ? freqEl.options[freqEl.selectedIndex].text.toLowerCase() : 'weekly';
          toast('Schedule saved — report will send ' + freqTxt, '#10b981');
        } else {
          toast('Save failed: ' + err.message, '#ef4444');
        }
      });
    },

    disableSchedule: function () {
      _schedEnabled = false;
      // Pass the full schedule object with enabled:false — Firestore dot-notation
      // path strings don't work reliably with the compat SDK's update().
      fsLoad(function (data) {
        var existing = data.schedule || {};
        fsSave({ schedule: Object.assign({}, existing, { enabled: false }) }, function () {
          updateSchedStatus();
          toast('Schedule disabled', '#64748b');
        });
      });
    },

    exportCSV: function () {
      if (!_rows.length) return;
      var now = new Date(), from = new Date(now);
      from.setDate(from.getDate() - _days);
      var fmt        = function (d) { return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }); };
      var exportRows = getDisplayRows();
      var enabled    = getEnabledRules();
      var ruleNames  = enabled.map(function (rid) {
        for (var i = 0; i < _allRules.length; i++) { if (_allRules[i].id === rid) return _allRules[i].name; }
        return rid;
      });
      var note  = _groupFilter ? (' - ' + (_groupFilter === 'top' ? 'Top Performers' : 'At-Risk')) : '';
      var lines = ['Driver Safety Scorecard' + note, 'Period: ' + fmt(from) + ' to ' + fmt(now), '', 'Driver,Score,Grade,Miles,' + ruleNames.join(',')];
      exportRows.forEach(function (r) {
        var gr = gradeFromScore(r.score);
        lines.push('"' + r.dname.replace(/"/g, '""') + '",' + r.score + ',' + gr.l + ',' + (r.displayVal || 0) + ',' + enabled.map(function (rid) { return r.ebr[rid] || 0; }).join(','));
      });
      var blob = new Blob([lines.join('\r\n')], { type: 'text/csv' });
      var url  = URL.createObjectURL(blob);
      var a    = document.createElement('a');
      a.href = url; a.download = 'safety-scorecard-' + _days + 'd.csv';
      document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(url);
    }
  };
})();

/* ── Logo swap ──────────────────────────────────────────── */
function updateLogos(isLight) {
  var darkLogo  = 'https://cdn.baseline.is/static/content/logos/mDtvptChQ3XwBLuxuu5hip-lala.svg';
  var lightLogo = 'https://cdn.baseline.is/static/content/logos/VpXJH6gs4Xy6HhRNbzgkmM-lala.svg';
  document.querySelectorAll('.hdr-logo').forEach(function (img) { img.src = isLight ? lightLogo : darkLogo; });
}

/* ── Geotab add-in entry point ───────────────────────────── */
geotab.addin = geotab.addin || {};
geotab.addin.safetyscorecard = function () {
  var _api   = null;
  var _state = null;

  return {
    /**
     * initialize — store references, call the callback immediately.
     * Do NOT call api.getSession() here; the session isn't guaranteed
     * to be fully ready until focus() fires.  Matches the HOS alerter
     * pattern exactly.
     */
    initialize: function (freshApi, freshState, initializeCallback) {
      _api   = freshApi;
      _state = freshState;
      if (typeof initializeCallback === 'function') initializeCallback();
    },

    /**
     * focus — called every time the user navigates to this add-in.
     * This is the correct place to call api.getSession() and kick off
     * Firestore + data loading, mirroring the HOS alerter's focus().
     */
    focus: function (freshApi, freshState) {
      _api   = freshApi;
      _state = freshState;

      // getSession provides the real authenticated database name.
      // fetch() is triggered inside init() after Firestore is ready —
      // do NOT call it here or it races against initFirestore completing.
      _api.getSession(function (session) {
        safetyDash.init(_api, session.database);
      });
    },

    blur: function () {}
  };
};
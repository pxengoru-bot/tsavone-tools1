/* Tsavone Consulting – Training site
   Plain JavaScript, no build step. Settings the owner may want to change are at the top. */
(function () {
  'use strict';

  /* ============ SETTINGS ============ */
  var SETTINGS = {
    // Live course list (Google Sheet published as CSV)
    sheetCsvUrl: 'https://docs.google.com/spreadsheets/d/10V8_jXPTFwRl2y940v_AtmMDPU8RGQBI6yH5qDy6N4I/gviz/tq?tqx=out:csv&gid=0',
    // Saved copy used if the sheet cannot be reached
    fallbackUrl: 'data/courses-fallback.json',
    // Show the saved list after this many milliseconds if the sheet is slow, then swap in live data when it arrives
    showSavedAfterMs: 2500,
    // Give up waiting for the sheet after this many milliseconds
    sheetTimeoutMs: 12000,

    // Fee estimator rules (see README "Changing fees and discounts")
    onlineDiscount: 0.40,          // online about 40% less
    groupSmallMin: 3,              // 3–4 people from one organisation ...
    groupSmallDiscount: 0.10,      // ... get 10%
    groupLargeMin: 5,              // 5 or more ...
    groupLargeDiscount: 0.15,      // ... get 15%
    earlyDiscount: 0.10,           // booking at least 30 days ahead
    selfSponsoredDiscount: 0.20,   // self-sponsored individuals
    // Used only when a course row has no fee: fee = base + perDay x days
    feeUgxBase: 150000, feeUgxPerDay: 300000,
    feeUsdBase: 50, feeUsdPerDay: 100,
    // In-house comparison shown in the estimator
    inHouseUgxPerDayMin: 3500000, inHouseUgxPerDayMax: 5000000, inHouseMaxPeople: 25,
    ugxPerUsd: 3000,

    email: 'info@tsavone.com'
  };

  // Track colours by track number. New tracks beyond 8 use the extra colours in turn.
  var TRACK_COLOURS = { 1: '#193591', 2: '#B04E00', 3: '#2B7A3A', 4: '#0B7373', 5: '#F2B705', 6: '#9B2F5E', 7: '#1565A8', 8: '#5B3FA0' };
  var EXTRA_COLOURS = ['#7A4E1D', '#2F6F8F', '#8A2D2D', '#3E6B1F', '#6B2F8A'];
  var LIGHT_COLOURS = { '#F2B705': true };   // colours that need dark text on top

  // Filterable formats: label shown, and words in the sheet that mean it
  var FORMATS = [
    { key: 'Open', words: ['open', 'open enrolment', 'open enrollment'] },
    { key: 'In-house', words: ['in-house', 'inhouse', 'in house'] },
    { key: 'Blended', words: ['blended'] },
    { key: 'Field', words: ['field', 'field-based', 'field based'] },
    { key: 'ToT', words: ['tot', 'training of trainers'] }
  ];

  /* ============ HELPERS ============ */
  var $ = function (sel, root) { return (root || document).querySelector(sel); };
  var $$ = function (sel, root) { return Array.prototype.slice.call((root || document).querySelectorAll(sel)); };
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function num(v) { var n = parseFloat(String(v == null ? '' : v).replace(/[^0-9.]/g, '')); return isFinite(n) ? n : 0; }
  function fmtUGX(n) { return 'UGX ' + Math.round(n).toLocaleString('en-GB'); }
  function fmtUSD(n) { return 'USD ' + Math.round(n).toLocaleString('en-GB'); }
  function plural(n, one, many) { return n + ' ' + (n === 1 ? one : many); }

  /* ============ CSV PARSER ============
     Handles quoted fields containing commas, quotes ("") and line breaks, and both \n and \r\n endings. */
  function parseCSV(text) {
    var rows = [], row = [], field = '', i = 0, inQuotes = false, c;
    text = String(text || '').replace(/^﻿/, '');
    while (i < text.length) {
      c = text[i];
      if (inQuotes) {
        if (c === '"') {
          if (text[i + 1] === '"') { field += '"'; i += 2; continue; }
          inQuotes = false; i++; continue;
        }
        field += c; i++; continue;
      }
      if (c === '"') { inQuotes = true; i++; continue; }
      if (c === ',') { row.push(field); field = ''; i++; continue; }
      if (c === '\r') { i++; continue; }
      if (c === '\n') { row.push(field); rows.push(row); row = []; field = ''; i++; continue; }
      field += c; i++;
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    if (!rows.length) return [];
    var headers = rows[0].map(function (h) { return h.trim().toLowerCase().replace(/\s+/g, '_'); });
    return rows.slice(1)
      .filter(function (r) { return r.some(function (v) { return v.trim() !== ''; }); })
      .map(function (r) {
        var o = {};
        headers.forEach(function (h, idx) { if (h) o[h] = (r[idx] || '').trim(); });
        return o;
      });
  }

  /* ============ DATA ============ */
  function normaliseFormat(word) {
    var w = word.trim().toLowerCase();
    if (!w) return null;
    for (var i = 0; i < FORMATS.length; i++) if (FORMATS[i].words.indexOf(w) !== -1) return FORMATS[i].key;
    return word.trim().charAt(0).toUpperCase() + word.trim().slice(1);
  }
  function trackColour(no) {
    if (TRACK_COLOURS[no]) return TRACK_COLOURS[no];
    return EXTRA_COLOURS[(Math.abs(no) || 0) % EXTRA_COLOURS.length];
  }

  // Turn sheet rows into courses. Only Active, New and Full are shown.
  function toCourses(rows) {
    var out = [];
    rows.forEach(function (r) {
      var status = String(r.status || '').trim().toLowerCase();
      if (status !== 'active' && status !== 'new' && status !== 'full') return;
      if (!r.title) return;
      var days = num(r.days);
      var trackNo = parseInt(r.track_no, 10);
      if (!isFinite(trackNo)) trackNo = 99;
      var formats = String(r.formats || '').split(/[,;\/]/).map(normaliseFormat).filter(Boolean);
      var feeUgx = num(r.fee_ugx) || (days ? SETTINGS.feeUgxBase + SETTINGS.feeUgxPerDay * days : 0);
      var feeUsd = num(r.fee_usd) || (days ? SETTINGS.feeUsdBase + SETTINGS.feeUsdPerDay * days : 0);
      out.push({
        code: (r.code || '').trim() || ('C' + (out.length + 1)),
        trackNo: trackNo,
        trackName: (r.track_name || '').trim() || ('Track ' + trackNo),
        title: r.title.trim(),
        days: days,
        formats: formats,
        audience: (r.audience || '').trim(),
        topics: String(r.topics || '').split('|').map(function (t) { return t.trim(); }).filter(Boolean),
        takesHome: (r.takes_home || '').trim(),
        feeUgx: feeUgx,
        feeUsd: feeUsd,
        status: status,
        colour: trackColour(trackNo)
      });
    });
    out.sort(function (a, b) {
      return a.trackNo - b.trackNo || a.code.localeCompare(b.code, 'en', { numeric: true });
    });
    out.forEach(function (c) {
      c.search = [c.code, c.title, c.trackName, c.audience, c.topics.join(' '), c.takesHome, c.formats.join(' ')].join(' ').toLowerCase();
    });
    return out;
  }

  function buildTracks(courses) {
    var map = {};
    courses.forEach(function (c) {
      if (!map[c.trackNo]) map[c.trackNo] = { no: c.trackNo, name: c.trackName, colour: c.colour, count: 0 };
      map[c.trackNo].count++;
    });
    return Object.keys(map).map(function (k) { return map[k]; }).sort(function (a, b) { return a.no - b.no; });
  }

  function fetchWithTimeout(url, ms) {
    var ctrl = typeof AbortController !== 'undefined' ? new AbortController() : null;
    var timer = setTimeout(function () { if (ctrl) ctrl.abort(); }, ms);
    return fetch(url, { signal: ctrl ? ctrl.signal : undefined, cache: 'no-store' })
      .then(function (res) { clearTimeout(timer); if (!res.ok) throw new Error('HTTP ' + res.status); return res.text(); })
      .catch(function (e) { clearTimeout(timer); throw e; });
  }

  function loadSheet() {
    return fetchWithTimeout(SETTINGS.sheetCsvUrl, SETTINGS.sheetTimeoutMs).then(function (text) {
      if (/^\s*</.test(text)) throw new Error('Sheet did not return CSV');
      var rows = parseCSV(text);
      if (!rows.length || !('title' in rows[0]) || !('status' in rows[0])) throw new Error('Unexpected sheet columns');
      return toCourses(rows);
    });
  }
  function loadFallback() {
    return fetchWithTimeout(SETTINGS.fallbackUrl, 15000).then(function (text) {
      var data = JSON.parse(text);
      return toCourses(Array.isArray(data) ? data : data.courses || []);
    });
  }

  /* ============ STATE ============ */
  var state = { courses: [], tracks: [], track: 'all', format: 'all', query: '', source: null };
  var byCode = {};

  /* ============ COURSE LIST ============ */
  function chipInk(colour) { return LIGHT_COLOURS[colour] ? '#1b1b1b' : '#fff'; }

  function renderChips() {
    var tc = $('#track-chips');
    var html = '<button type="button" class="chip" data-track="all" aria-pressed="' + (state.track === 'all') + '">All tracks <span class="count">' + state.courses.length + '</span></button>';
    state.tracks.forEach(function (t) {
      html += '<button type="button" class="chip" data-track="' + t.no + '" aria-pressed="' + (String(state.track) === String(t.no)) +
        '" style="--c:' + t.colour + ';--chip-ink:' + chipInk(t.colour) + '"><span class="dot" aria-hidden="true"></span>' + esc(t.name) +
        ' <span class="count">' + t.count + '</span></button>';
    });
    tc.innerHTML = html;
    tc.classList.add('scroll-x');

    var fc = $('#format-chips');
    html = '<button type="button" class="chip" data-format="all" aria-pressed="' + (state.format === 'all') + '">All formats</button>';
    FORMATS.forEach(function (f) {
      html += '<button type="button" class="chip" data-format="' + f.key + '" aria-pressed="' + (state.format === f.key) + '">' + f.key + '</button>';
    });
    fc.innerHTML = html;
  }

  function cardHTML(c) {
    var badge = c.status === 'new' ? '<span class="badge badge-new">New</span>' : c.status === 'full' ? '<span class="badge badge-full">Fully booked</span>' : '';
    var meta = '<li class="days">' + plural(c.days, 'day', 'days') + '</li>' + c.formats.map(function (f) { return '<li>' + esc(f) + '</li>'; }).join('');
    return '<button type="button" class="course-card' + (c.status === 'full' ? ' is-full' : '') + '" data-code="' + esc(c.code) +
      '" style="--c:' + c.colour + ';--chip-ink:' + chipInk(c.colour) + '" aria-haspopup="dialog">' +
      '<span class="card-band" aria-hidden="true"></span><span class="card-body">' +
      '<span class="card-top"><span class="code-chip">' + esc(c.code) + '</span>' + badge + '</span>' +
      '<span class="track-name">' + esc(c.trackName) + '</span>' +
      '<h3>' + esc(c.title) + '</h3>' +
      '<ul class="card-meta" aria-label="Duration and formats">' + meta + '</ul>' +
      '<span class="card-foot"><span class="card-fee">' + fmtUGX(c.feeUgx) + '<small>' + fmtUSD(c.feeUsd) + ' per participant</small></span>' +
      '<span class="card-more" aria-hidden="true">Details &rarr;</span></span>' +
      '</span></button>';
  }

  function filtered() {
    var words = state.query.toLowerCase().split(/\s+/).filter(Boolean);
    return state.courses.filter(function (c) {
      if (state.track !== 'all' && String(c.trackNo) !== String(state.track)) return false;
      if (state.format !== 'all' && c.formats.indexOf(state.format) === -1) return false;
      for (var i = 0; i < words.length; i++) if (c.search.indexOf(words[i]) === -1) return false;
      return true;
    });
  }

  function renderGrid() {
    var list = filtered();
    var grid = $('#course-grid');
    grid.innerHTML = list.map(cardHTML).join('');
    grid.setAttribute('aria-busy', 'false');
    $('#empty-state').hidden = list.length > 0;
    var total = state.courses.length;
    $('#results-count').textContent = list.length === total
      ? 'Showing all ' + plural(total, 'course', 'courses')
      : 'Showing ' + list.length + ' of ' + plural(total, 'course', 'courses');
    $('#clear-filters').hidden = state.track === 'all' && state.format === 'all' && !state.query;
  }

  function renderStats() {
    $('#stat-courses').textContent = state.courses.length;
    $('#stat-tracks').textContent = state.tracks.length;
  }

  /* ============ COURSE DETAIL ============ */
  var dialog = $('#course-dialog');
  var lastCard = null;

  function openCourse(code, fromEl) {
    var c = byCode[code];
    if (!c) return;
    lastCard = fromEl || null;
    dialog.style.setProperty('--c', c.colour);
    dialog.style.setProperty('--chip-ink', chipInk(c.colour));
    $('#dlg-code').textContent = c.code;
    $('#dlg-track').textContent = c.trackName;
    $('#dlg-title').textContent = c.title;
    var st = $('#dlg-status');
    st.hidden = c.status === 'active';
    st.innerHTML = c.status === 'new' ? '<span class="badge badge-new">New course</span>' : c.status === 'full' ? '<span class="badge badge-full">Fully booked</span>' : '';
    $('#dlg-days').textContent = plural(c.days, 'day', 'days');
    $('#dlg-formats').textContent = c.formats.join(', ') || 'On request';
    $('#dlg-audience').textContent = c.audience || 'Open to all interested professionals.';
    $('#dlg-topics').innerHTML = c.topics.map(function (t) { return '<li>' + esc(t) + '</li>'; }).join('');
    $('#dlg-takes').textContent = c.takesHome || 'Course workbook, templates and an action plan.';
    $('#dlg-ugx').textContent = fmtUGX(c.feeUgx);
    $('#dlg-usd').textContent = '/ ' + fmtUSD(c.feeUsd);
    var book = $('#dlg-book');
    if (c.status === 'full') {
      book.disabled = true; book.textContent = 'Fully booked';
    } else {
      book.disabled = false; book.textContent = 'Book this course';
    }
    book.dataset.code = c.code;
    $('#dlg-estimate').dataset.code = c.code;
    if (typeof dialog.showModal === 'function') dialog.showModal(); else dialog.setAttribute('open', '');
    document.body.classList.add('dialog-open');
    dialog.scrollTop = 0;
    try { history.replaceState(null, '', '#course-' + encodeURIComponent(c.code)); } catch (e) { /* ignore */ }
  }
  function closeDialog() { if (dialog.open) { if (dialog.close) dialog.close(); else dialog.removeAttribute('open'); } }
  dialog.addEventListener('close', function () {
    document.body.classList.remove('dialog-open');
    if (/^#course-/.test(location.hash)) { try { history.replaceState(null, '', '#courses'); } catch (e) { /* ignore */ } }
    if (lastCard && document.body.contains(lastCard)) lastCard.focus();
  });
  dialog.addEventListener('click', function (e) { if (e.target === dialog) closeDialog(); }); // click on backdrop

  /* ============ BOOKING FORM ============ */
  function courseLabel(c) { return c.code + ' – ' + c.title; }

  function fillSelects() {
    var groups = state.tracks.map(function (t) {
      var opts = state.courses.filter(function (c) { return c.trackNo === t.no; }).map(function (c) {
        var label = courseLabel(c) + (c.status === 'full' ? ' (fully booked)' : c.status === 'new' ? ' (new)' : '');
        return '<option value="' + esc(courseLabel(c)) + '"' + (c.status === 'full' ? ' disabled' : '') + '>' + esc(label) + '</option>';
      }).join('');
      return '<optgroup label="' + esc(t.no + '. ' + t.name) + '">' + opts + '</optgroup>';
    }).join('');

    var fc = $('#f-course'), keep = fc.value;
    fc.innerHTML = '<option value="">Choose a course</option><option value="Custom or in-house course">Custom or in-house course</option>' + groups;
    if (keep) fc.value = keep;

    var ec = $('#est-course'), keepE = ec.value;
    ec.innerHTML = state.tracks.map(function (t) {
      return '<optgroup label="' + esc(t.no + '. ' + t.name) + '">' + state.courses.filter(function (c) { return c.trackNo === t.no; }).map(function (c) {
        return '<option value="' + esc(c.code) + '">' + esc(c.code + ' – ' + c.title + ' (' + plural(c.days, 'day', 'days') + ')') + '</option>';
      }).join('') + '</optgroup>';
    }).join('');
    if (keepE && byCode[keepE]) ec.value = keepE;
  }

  function goToBooking(opts) {
    opts = opts || {};
    if (opts.course) $('#f-course').value = opts.course;
    if (opts.format) $('#f-format').value = opts.format;
    if (opts.people) $('#f-people').value = opts.people;
    if (opts.message) $('#f-message').value = opts.message;
    updateMailto();
    var target = $('#book');
    target.scrollIntoView({ behavior: matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth' });
    var form = $('#book-form');
    form.classList.remove('flash'); void form.offsetWidth; form.classList.add('flash');
    setTimeout(function () { $('#f-name').focus({ preventScroll: true }); }, 450);
  }

  function updateMailto() {
    var f = $('#book-form');
    var get = function (n) { var el = f.elements[n]; return el ? el.value.trim() : ''; };
    var subject = 'Training enquiry' + (get('course') ? ': ' + get('course') : '');
    var lines = [
      'Name: ' + get('name'), 'Organisation: ' + get('organisation'), 'Email: ' + get('email'), 'Phone: ' + get('phone'),
      'Course: ' + get('course'), 'Format: ' + get('format'), 'Participants: ' + get('participants'),
      'Preferred dates: ' + get('preferred_dates'), 'Location: ' + get('location'), '', get('message')
    ];
    $('#mailto-link').href = 'mailto:' + SETTINGS.email + '?subject=' + encodeURIComponent(subject) + '&body=' + encodeURIComponent(lines.join('\n'));
  }

  function setupForm() {
    var form = $('#book-form');
    // Send people back to our own thank-you page after FormSubmit receives the form
    if (/^https?:$/.test(location.protocol)) {
      $('#f-next').value = new URL('thank-you.html', location.href).href;
    }
    form.addEventListener('invalid', function () {
      var err = $('#form-error');
      err.textContent = 'Please fill in your name, a valid email address and choose a course.';
      err.hidden = false;
    }, true);
    form.addEventListener('input', updateMailto);
    form.addEventListener('change', updateMailto);
    form.addEventListener('submit', function (e) {
      var err = $('#form-error');
      if ($('#f-honey').value) { e.preventDefault(); return; } // robot
      if (!form.checkValidity()) {
        e.preventDefault();
        err.textContent = 'Please fill in your name, a valid email address and choose a course.';
        err.hidden = false;
        var bad = form.querySelector(':invalid'); if (bad) bad.focus();
        return;
      }
      err.hidden = true;
      var who = $('#f-org').value.trim() || $('#f-name').value.trim();
      $('#f-subject').value = 'Training enquiry: ' + ($('#f-course').value || 'General') + (who ? ' – ' + who : '');
      var btn = $('#f-submit'); btn.disabled = true; btn.textContent = 'Sending…';
      // Re-enable if the person comes back with the browser Back button
      setTimeout(function () { btn.disabled = false; btn.textContent = 'Send enquiry'; }, 8000);
    });
    updateMailto();
  }

  /* ============ FEE ESTIMATOR ============ */
  function estimate() {
    var c = byCode[$('#est-course').value];
    var out = $('#est-total'), per = $('#est-per'), lines = $('#est-lines'), hint = $('#est-hint');
    if (!c) { out.textContent = '–'; per.textContent = ''; lines.innerHTML = ''; hint.hidden = true; return null; }
    var f = $('#est-form').elements;
    var people = Math.max(1, Math.min(500, parseInt($('#est-people').value, 10) || 1));
    var intl = f.region.value === 'intl';
    var online = f.mode.value === 'online';
    var early = f.early.value === 'yes';
    var self = f.self.value === 'yes';
    var fee = intl ? c.feeUsd : c.feeUgx;
    var fmt = intl ? fmtUSD : fmtUGX;

    var total = fee * people;
    var rows = [{ label: plural(people, 'participant', 'participants') + ' × ' + fmt(fee), value: fmt(total) }];
    function apply(rate, label) {
      var cut = total * rate;
      total -= cut;
      rows.push({ label: label, value: '− ' + fmt(cut), disc: true });
    }
    if (online) apply(SETTINGS.onlineDiscount, 'Online (about ' + Math.round(SETTINGS.onlineDiscount * 100) + '% less)');
    // Group discount: only the larger one applies, and only once. Not combined with the self-sponsored rate.
    var group = 0;
    if (!self) {
      if (people >= SETTINGS.groupLargeMin) group = SETTINGS.groupLargeDiscount;
      else if (people >= SETTINGS.groupSmallMin) group = SETTINGS.groupSmallDiscount;
    }
    if (group) apply(group, 'Group of ' + people + ' (' + Math.round(group * 100) + '%)');
    if (early) apply(SETTINGS.earlyDiscount, 'Early booking (' + Math.round(SETTINGS.earlyDiscount * 100) + '%)');
    if (self) apply(SETTINGS.selfSponsoredDiscount, 'Self-sponsored (' + Math.round(SETTINGS.selfSponsoredDiscount * 100) + '%)');

    total = intl ? Math.round(total) : Math.round(total / 1000) * 1000;
    out.textContent = fmt(total);
    per.textContent = people > 1 ? 'About ' + fmt(total / people) + ' per participant' : 'For one participant';
    lines.innerHTML = rows.map(function (r) {
      return '<li' + (r.disc ? ' class="disc"' : '') + '><span>' + esc(r.label) + '</span><span>' + esc(r.value) + '</span></li>';
    }).join('');

    var notes = [];
    if (self && people >= SETTINGS.groupSmallMin) notes.push('The self-sponsored rate replaces the group discount.');
    if (people >= 10 && !online && c.days) {
      var lo = SETTINGS.inHouseUgxPerDayMin * c.days, hi = SETTINGS.inHouseUgxPerDayMax * c.days;
      var loCmp = intl ? lo / SETTINGS.ugxPerUsd : lo;
      if (loCmp < total) {
        var mil = function (n) { return (Math.round(n / 1e5) / 10).toLocaleString('en-GB'); };
        notes.push('For a group from one organisation, in-house delivery may cost less: about UGX ' + mil(lo) + '–' + mil(hi) +
          ' million for ' + plural(c.days, 'day', 'days') + ' (up to ' + SETTINGS.inHouseMaxPeople + ' people)' +
          (intl ? ', roughly ' + fmtUSD(lo / SETTINGS.ugxPerUsd) + '–' + Math.round(hi / SETTINGS.ugxPerUsd).toLocaleString('en-GB') : '') + '.');
      }
    }
    hint.hidden = !notes.length;
    hint.textContent = notes.join(' ');
    return { course: c, people: people, total: fmt(total), online: online, intl: intl, early: early, self: self };
  }

  function setupEstimator() {
    $('#est-form').addEventListener('input', estimate);
    $('#est-form').addEventListener('change', estimate);
    $('#est-form').addEventListener('submit', function (e) { e.preventDefault(); });
    $('#est-book').addEventListener('click', function () {
      var r = estimate();
      if (!r) { goToBooking(); return; }
      goToBooking({
        course: r.course.status === 'full' ? 'Custom or in-house course' : courseLabel(r.course),
        format: r.online ? 'Online' : '',
        people: r.people,
        message: 'Please send a quote. Website estimate: ' + r.total + ' for ' + plural(r.people, 'participant', 'participants') +
          ' (' + (r.intl ? 'international' : 'Uganda') + ', ' + (r.online ? 'online' : 'in-person') +
          (r.early ? ', early booking' : '') + (r.self ? ', self-sponsored' : '') + ').'
      });
    });
  }

  /* ============ STRUCTURED DATA (schema.org Course) ============ */
  function updateSchema() {
    var modes = { Open: 'Onsite', 'In-house': 'Onsite', Blended: 'Blended', Field: 'Onsite', ToT: 'Onsite' };
    var data = {
      '@context': 'https://schema.org', '@type': 'ItemList',
      itemListElement: state.courses.map(function (c, i) {
        var courseModes = [];
        c.formats.forEach(function (f) { var m = modes[f] || 'Onsite'; if (courseModes.indexOf(m) === -1) courseModes.push(m); });
        return {
          '@type': 'ListItem', position: i + 1,
          item: {
            '@type': 'Course', name: c.title, courseCode: c.code,
            description: (c.audience ? 'For ' + c.audience.charAt(0).toLowerCase() + c.audience.slice(1) + '. ' : '') + 'Covers: ' + c.topics.join('; ') + '.',
            url: location.href.split('#')[0] + '#course-' + encodeURIComponent(c.code),
            provider: { '@type': 'Organization', name: 'Tsavone Consulting', sameAs: 'https://www.tsavone.com' },
            about: c.trackName,
            inLanguage: 'en',
            offers: [{ '@type': 'Offer', category: 'Paid', price: c.feeUgx, priceCurrency: 'UGX' }, { '@type': 'Offer', category: 'Paid', price: c.feeUsd, priceCurrency: 'USD' }],
            hasCourseInstance: { '@type': 'CourseInstance', courseMode: courseModes.length ? courseModes : ['Onsite'], courseWorkload: 'P' + (c.days || 1) + 'D', location: 'Kampala, Uganda' }
          }
        };
      })
    };
    $('#course-schema').textContent = JSON.stringify(data);
  }

  /* ============ APPLY DATA ============ */
  function applyCourses(courses, source) {
    state.courses = courses;
    state.source = source;
    byCode = {};
    courses.forEach(function (c) { byCode[c.code] = c; });
    state.tracks = buildTracks(courses);
    // A link ending in #track-3 opens the list filtered to that track
    var tm = location.hash.match(/^#track-(\d+)$/);
    var jumpToTrack = false;
    if (tm && !state.trackFromLink && state.tracks.some(function (t) { return String(t.no) === tm[1]; })) {
      state.track = tm[1]; state.trackFromLink = true; jumpToTrack = true;
    }
    if (state.track !== 'all' && !state.tracks.some(function (t) { return String(t.no) === String(state.track); })) state.track = 'all';
    $('#data-note').hidden = source === 'live';
    renderChips(); renderGrid(); renderStats(); fillSelects(); estimate(); updateSchema();
    var m = location.hash.match(/^#course-(.+)$/);
    if (m && !dialog.open && byCode[decodeURIComponent(m[1])]) openCourse(decodeURIComponent(m[1]));
    if (jumpToTrack) $('#courses').scrollIntoView();
  }

  function showLoadError() {
    $('#course-grid').innerHTML = '';
    $('#course-grid').setAttribute('aria-busy', 'false');
    $('#results-count').innerHTML = 'We could not load the course list right now. Please refresh the page, or email <a href="mailto:' + SETTINGS.email + '">' + SETTINGS.email + '</a> for the catalogue.';
  }

  function loadData() {
    var liveDone = false, savedShown = false;
    var fallbackPromise = null;
    function getFallback() { return fallbackPromise || (fallbackPromise = loadFallback()); }

    var slowTimer = setTimeout(function () {
      if (liveDone) return;
      getFallback().then(function (courses) {
        if (liveDone || !courses.length) return;
        savedShown = true; applyCourses(courses, 'saved');
      }).catch(function () { /* wait for the sheet */ });
    }, SETTINGS.showSavedAfterMs);

    loadSheet().then(function (courses) {
      liveDone = true; clearTimeout(slowTimer);
      applyCourses(courses, 'live');
    }).catch(function (err) {
      liveDone = true; clearTimeout(slowTimer);
      if (window.console) console.warn('Course sheet unavailable, using saved list:', err && err.message);
      if (savedShown) return;
      getFallback().then(function (courses) {
        if (!courses.length) throw new Error('empty');
        applyCourses(courses, 'saved');
      }).catch(showLoadError);
    });
  }

  /* ============ NAVIGATION ============ */
  function setupNav() {
    var toggle = $('.nav-toggle'), nav = $('#site-nav'), header = $('.site-header');
    function setOpen(open) { toggle.setAttribute('aria-expanded', String(open)); nav.classList.toggle('open', open); }
    toggle.addEventListener('click', function () { setOpen(toggle.getAttribute('aria-expanded') !== 'true'); });
    nav.addEventListener('click', function (e) { if (e.target.closest('a')) setOpen(false); });
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && nav.classList.contains('open')) { setOpen(false); toggle.focus(); } });

    var onScroll = function () { header.classList.toggle('scrolled', window.scrollY > 8); };
    window.addEventListener('scroll', onScroll, { passive: true }); onScroll();

    if ('IntersectionObserver' in window) {
      var links = {};
      $$('.site-nav a[href^="#"]').forEach(function (a) { links[a.getAttribute('href').slice(1)] = a; });
      var io = new IntersectionObserver(function (entries) {
        entries.forEach(function (en) {
          var a = links[en.target.id]; if (!a) return;
          if (en.isIntersecting) { $$('.site-nav a.active').forEach(function (x) { x.classList.remove('active'); x.removeAttribute('aria-current'); }); a.classList.add('active'); a.setAttribute('aria-current', 'true'); }
        });
      }, { rootMargin: '-45% 0px -50% 0px' });
      Object.keys(links).forEach(function (id) { var s = document.getElementById(id); if (s) io.observe(s); });
    }

    // Links that pre-fill the booking form
    document.addEventListener('click', function (e) {
      var a = e.target.closest('[data-prefill-format],[data-prefill-course]');
      if (!a) return;
      e.preventDefault();
      goToBooking({
        format: a.getAttribute('data-prefill-format') || '',
        course: a.getAttribute('data-prefill-course') === 'custom' || a.getAttribute('data-prefill-format') === 'In-house' ? 'Custom or in-house course' : ''
      });
    });
  }

  /* ============ FILTER EVENTS ============ */
  function setupFilters() {
    $('#track-chips').addEventListener('click', function (e) {
      var b = e.target.closest('[data-track]'); if (!b) return;
      state.track = b.getAttribute('data-track'); renderChips(); renderGrid();
      var sel = $('#track-chips [aria-pressed="true"]'); if (sel) sel.focus();
    });
    $('#format-chips').addEventListener('click', function (e) {
      var b = e.target.closest('[data-format]'); if (!b) return;
      state.format = b.getAttribute('data-format'); renderChips(); renderGrid();
      var sel = $('#format-chips [aria-pressed="true"]'); if (sel) sel.focus();
    });
    var t;
    $('#course-search').addEventListener('input', function (e) {
      clearTimeout(t);
      t = setTimeout(function () { state.query = e.target.value.trim(); renderGrid(); }, 120);
    });
    $('#clear-filters').addEventListener('click', function () {
      state.track = 'all'; state.format = 'all'; state.query = ''; $('#course-search').value = '';
      renderChips(); renderGrid(); $('#course-search').focus();
    });
    $('#course-grid').addEventListener('click', function (e) {
      var card = e.target.closest('.course-card'); if (card) openCourse(card.getAttribute('data-code'), card);
    });
    $('#dlg-book').addEventListener('click', function () {
      var c = byCode[this.dataset.code]; if (!c || c.status === 'full') return;
      lastCard = null; closeDialog();
      goToBooking({ course: courseLabel(c) });
    });
    $('#dlg-estimate').addEventListener('click', function () {
      $('#est-course').value = this.dataset.code; estimate();
      lastCard = null; closeDialog();
      $('#estimator').scrollIntoView({ behavior: 'smooth' });
      setTimeout(function () { $('#est-people').focus({ preventScroll: true }); }, 450);
    });
  }

  // Respond to #track-N and #course-CODE links clicked while the page is open
  window.addEventListener('hashchange', function () {
    if (!state.courses.length) return;
    var t = location.hash.match(/^#track-(\d+)$/);
    if (t && state.tracks.some(function (x) { return String(x.no) === t[1]; })) {
      state.track = t[1]; renderChips(); renderGrid(); $('#courses').scrollIntoView();
    }
    var c = location.hash.match(/^#course-(.+)$/);
    if (c && !dialog.open && byCode[decodeURIComponent(c[1])]) openCourse(decodeURIComponent(c[1]));
  });

  /* ============ START ============ */
  setupNav();
  setupFilters();
  setupForm();
  setupEstimator();
  loadData();

  // Expose the parser for testing
  window.TsavoneTraining = { parseCSV: parseCSV, toCourses: toCourses, settings: SETTINGS };
})();

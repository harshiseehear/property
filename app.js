var __flatAppFn = function() {
  'use strict';

  // ======================== STATE ========================
  var parsedData = null;       // { flats, pairs, excelSeed }
  var allocationResult = null; // { allocations, auditTrail, validation, seed, pairs }

  // ======================== UTILITIES ========================
  function escapeHtml(str) {
    var d = document.createElement('div');
    d.appendChild(document.createTextNode(str));
    return d.innerHTML;
  }

  function pad2(n) { return n < 10 ? '0' + n : '' + n; }

  function showStatus(el, msg, type) {
    el.textContent = msg;
    el.className = 'status-msg visible ' + type;
  }

  function hideStatus(el) {
    el.className = 'status-msg';
  }

  // ======================== SEEDED PRNG (mulberry32) ========================
  function mulberry32(a) {
    return function() {
      a |= 0;
      a = a + 0x6D2B79F5 | 0;
      var t = Math.imul(a ^ a >>> 15, 1 | a);
      t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t;
      return ((t ^ t >>> 14) >>> 0) / 4294967296;
    };
  }

  function createRNG(seed) {
    var next = mulberry32(seed);
    return {
      next: next,
      randomInt: function(min, max) {
        return min + Math.floor(next() * (max - min + 1));
      },
      shuffle: function(arr) {
        var a = arr.slice();
        for (var i = a.length - 1; i > 0; i--) {
          var j = Math.floor(next() * (i + 1));
          var tmp = a[i]; a[i] = a[j]; a[j] = tmp;
        }
        return a;
      },
      pick: function(arr) {
        return arr[Math.floor(next() * arr.length)];
      }
    };
  }

  // ======================== DOM REFS ========================
  var dropZone      = document.getElementById('drop-zone');
  var fileInput     = document.getElementById('file-input');
  var fileInfo      = document.getElementById('file-info');
  var parseStatus   = document.getElementById('parse-status');
  var runBtn        = document.getElementById('run-btn');
  var rerunBtn      = document.getElementById('rerun-btn');
  var runStatus     = document.getElementById('run-status');
  var resultsSection = document.getElementById('results-section');
  var seedValueEl   = document.getElementById('seed-value');
  var validationList = document.getElementById('validation-list');
  var allocTbody    = document.querySelector('#allocation-table tbody');
  var layoutTbody   = document.querySelector('#layout-table tbody');
  var downloadBtn   = document.getElementById('download-btn');

  // ======================== FILE UPLOAD ========================
  dropZone.addEventListener('click', function() { fileInput.click(); });

  dropZone.addEventListener('dragover', function(e) {
    e.preventDefault();
    dropZone.classList.add('drag-over');
  });

  dropZone.addEventListener('dragleave', function() {
    dropZone.classList.remove('drag-over');
  });

  dropZone.addEventListener('drop', function(e) {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
  });

  fileInput.addEventListener('change', function() {
    if (fileInput.files.length) handleFile(fileInput.files[0]);
  });

  function handleFile(file) {
    if (!/\.xlsx?$/i.test(file.name)) {
      showStatus(parseStatus, 'Please upload an .xlsx file.', 'error');
      return;
    }

    var reader = new FileReader();
    reader.onload = function(e) {
      try {
        var data = new Uint8Array(e.target.result);
        var wb = XLSX.read(data, { type: 'array' });
        parsedData = parseWorkbook(wb);

        fileInfo.innerHTML =
          '<strong>' + escapeHtml(file.name) + '</strong> - ' +
          parsedData.flats.length + ' flats loaded, ' +
          parsedData.pairs.length + ' pair(s) detected' +
          (parsedData.excelSeed != null ? ', seed: ' + parsedData.excelSeed : ', no seed in file');
        fileInfo.classList.add('visible');

        showStatus(parseStatus, 'File parsed successfully. Ready to run allocation.', 'success');
        runBtn.disabled = false;
      } catch (err) {
        showStatus(parseStatus, 'Error: ' + err.message, 'error');
        runBtn.disabled = true;
        parsedData = null;
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // ======================== PARSE WORKBOOK ========================
  function parseWorkbook(wb) {
    // --- Flat Owners ---
    var ownerSheet = wb.Sheets['Flat Owners'];
    if (!ownerSheet) throw new Error('Sheet "Flat Owners" not found.');

    var rows = XLSX.utils.sheet_to_json(ownerSheet, { header: 1 });
    if (rows.length < 3) throw new Error('"Flat Owners" sheet has too few rows.');

    var flats = [];
    for (var i = 2; i < rows.length; i++) {
      var row = rows[i];
      if (!row || row.length === 0 || row[0] == null || row[0] === '') continue;
      var flatNo = parseInt(row[0], 10);
      if (isNaN(flatNo)) continue;
      flats.push({
        oldFlatNo: flatNo,
        ownerName: (row[1] || '').toString().trim(),
        contact:   (row[2] || '').toString().trim(),
        constraintGroup: (row[3] || 'Individual').toString().trim(),
        wingPreference:  (row[4] || 'Any').toString().trim()
      });
    }

    if (flats.length !== 30) {
      throw new Error('Expected 30 flats, found ' + flats.length + '.');
    }

    // --- Parse constraint groups dynamically ---
    var pairMap = {};
    var pairRe = /^Pair\s+(\d+)\s+\(with\s+#(\d+)\)$/i;

    for (var j = 0; j < flats.length; j++) {
      var m = pairRe.exec(flats[j].constraintGroup);
      if (m) {
        var pid = parseInt(m[1], 10);
        var partner = parseInt(m[2], 10);
        if (!pairMap[pid]) pairMap[pid] = [];
        pairMap[pid].push({ flatNo: flats[j].oldFlatNo, partner: partner });
      }
    }

    var pairs = [];
    var pairIds = Object.keys(pairMap).sort(function(a, b) { return a - b; });
    for (var k = 0; k < pairIds.length; k++) {
      var id = pairIds[k];
      var members = pairMap[id];
      if (members.length !== 2) {
        throw new Error('Pair ' + id + ' must have exactly 2 flats (found ' + members.length + ').');
      }
      if (members[0].partner !== members[1].flatNo || members[1].partner !== members[0].flatNo) {
        throw new Error('Pair ' + id + ' flats do not cross-reference each other.');
      }
      members.sort(function(a, b) { return a.flatNo - b.flatNo; });
      pairs.push({ pairId: parseInt(id, 10), flat1: members[0].flatNo, flat2: members[1].flatNo });
    }

    // --- Randomisation Seed ---
    var excelSeed = null;
    var seedSheet = wb.Sheets['Randomisation Seed'];
    if (seedSheet) {
      var cell = seedSheet['B5'];
      if (cell && cell.v != null && cell.v !== '') {
        var sv = parseInt(cell.v, 10);
        if (!isNaN(sv)) excelSeed = sv;
      }
    }

    return { flats: flats, pairs: pairs, excelSeed: excelSeed };
  }

  // ======================== ALLOCATION ALGORITHM ========================
  function runAllocationAlgorithm(data, seed) {
    var rng = createRNG(seed);
    var auditTrail = [];
    var allocations = {};   // oldFlatNo -> { newFlatCode, wing, floor, unit, type }

    // Step 1 — Build list of all 30 new flats
    var available = [];
    for (var f = 1; f <= 10; f++) {
      available.push('A-' + pad2(f) + '-01');
      available.push('B-' + pad2(f) + '-01');
      available.push('B-' + pad2(f) + '-02');
    }

    var usedPairFloors = [];

    // Steps 2 & 3 — Allocate each pair
    for (var p = 0; p < data.pairs.length; p++) {
      var pair = data.pairs[p];

      // Find Wing B floors where both units are still available
      var eligible = [];
      for (var fl = 1; fl <= 10; fl++) {
        if (usedPairFloors.indexOf(fl) >= 0) continue;
        var u1 = 'B-' + pad2(fl) + '-01';
        var u2 = 'B-' + pad2(fl) + '-02';
        if (available.indexOf(u1) >= 0 && available.indexOf(u2) >= 0) {
          eligible.push(fl);
        }
      }

      if (eligible.length === 0) {
        throw new Error('No eligible Wing B floor for Pair ' + pair.pairId + '.');
      }

      var chosenFloor = rng.pick(eligible);
      usedPairFloors.push(chosenFloor);

      var code1 = 'B-' + pad2(chosenFloor) + '-01';
      var code2 = 'B-' + pad2(chosenFloor) + '-02';

      allocations[pair.flat1] = { newFlatCode: code1, wing: 'B', floor: chosenFloor, unit: 1, type: 'Paired' };
      allocations[pair.flat2] = { newFlatCode: code2, wing: 'B', floor: chosenFloor, unit: 2, type: 'Paired' };

      available.splice(available.indexOf(code1), 1);
      available.splice(available.indexOf(code2), 1);

      var stepBase = p * 2 + 1;
      auditTrail.push({
        step: stepBase, type: 'Paired', oldFlatNo: pair.flat1, newFlatCode: code1,
        notes: 'Pair ' + pair.pairId + ': Floor ' + chosenFloor + ' Unit 01 (from ' + eligible.length + ' eligible floors)'
      });
      auditTrail.push({
        step: stepBase + 1, type: 'Paired', oldFlatNo: pair.flat2, newFlatCode: code2,
        notes: 'Pair ' + pair.pairId + ': Floor ' + chosenFloor + ' Unit 02'
      });
    }

    // Step 4 — Remaining flats
    var assignedSet = {};
    for (var key in allocations) assignedSet[key] = true;

    var remaining = [];
    for (var r = 0; r < data.flats.length; r++) {
      if (!assignedSet[data.flats[r].oldFlatNo]) {
        remaining.push(data.flats[r].oldFlatNo);
      }
    }

    remaining = rng.shuffle(remaining);

    for (var s = 0; s < remaining.length; s++) {
      var oldNo = remaining[s];
      var idx = rng.randomInt(0, available.length - 1);
      var picked = available[idx];
      available.splice(idx, 1);

      var parts = picked.split('-');
      allocations[oldNo] = {
        newFlatCode: picked,
        wing: parts[0],
        floor: parseInt(parts[1], 10),
        unit: parseInt(parts[2], 10),
        type: 'Random'
      };

      auditTrail.push({
        step: data.pairs.length * 2 + s + 1,
        type: 'Random', oldFlatNo: oldNo, newFlatCode: picked,
        notes: 'Random allocation (' + (available.length + 1) + ' flats were available)'
      });
    }

    // Step 5 — Validate
    var validation = validateAllocation(allocations, data.pairs);

    return {
      allocations: allocations,
      auditTrail: auditTrail,
      validation: validation,
      seed: seed,
      pairs: data.pairs
    };
  }

  // ======================== VALIDATION ========================
  function validateAllocation(allocations, pairs) {
    var checks = [];
    var allNew = [];
    for (var key in allocations) allNew.push(allocations[key].newFlatCode);

    var pairFloors = [];

    for (var p = 0; p < pairs.length; p++) {
      var pair = pairs[p];
      var a1 = allocations[pair.flat1];
      var a2 = allocations[pair.flat2];
      var ok = a1 && a2 &&
        a1.wing === 'B' && a2.wing === 'B' &&
        a1.floor === a2.floor &&
        a1.unit === 1 && a2.unit === 2;

      pairFloors.push(a1 ? a1.floor : null);

      checks.push({
        constraint: 'Pair ' + pair.pairId + ' (Flats ' + pair.flat1 + ' & ' + pair.flat2 + '): same Wing B floor, units 01 & 02',
        status: ok,
        details: ok
          ? 'Both on Wing B Floor ' + a1.floor + ' (units 01 & 02)'
          : 'FAILED - not on same Wing B floor or wrong units'
      });
    }

    // Pair floors must differ
    if (pairs.length >= 2) {
      var unique = true;
      for (var i = 0; i < pairFloors.length; i++) {
        for (var j = i + 1; j < pairFloors.length; j++) {
          if (pairFloors[i] === pairFloors[j]) unique = false;
        }
      }
      checks.push({
        constraint: 'Paired flats on different floors',
        status: unique,
        details: unique
          ? 'Pair floors: ' + pairFloors.join(', ')
          : 'FAILED - two pairs share the same floor'
      });
    }

    // No duplicate new flats
    var newSet = new Set(allNew);
    var noDups = newSet.size === allNew.length;
    checks.push({
      constraint: 'No duplicate new flat assignments',
      status: noDups,
      details: noDups
        ? 'All ' + allNew.length + ' assignments unique'
        : 'FAILED - ' + (allNew.length - newSet.size) + ' duplicate(s)'
    });

    // All 30 old flats
    var oldCount = Object.keys(allocations).length;
    checks.push({
      constraint: 'All 30 old flats assigned',
      status: oldCount === 30,
      details: oldCount === 30
        ? '30/30 old flats assigned'
        : 'FAILED - ' + oldCount + '/30 assigned'
    });

    // All 30 new flats
    checks.push({
      constraint: 'All 30 new flats used',
      status: newSet.size === 30,
      details: newSet.size === 30
        ? '30/30 new flats allocated'
        : 'FAILED - ' + newSet.size + '/30 used'
    });

    return checks;
  }

  // ======================== RENDER RESULTS ========================
  function findOldFlatByCode(allocations, code) {
    for (var key in allocations) {
      if (allocations[key].newFlatCode === code) return parseInt(key, 10);
    }
    return null;
  }

  function showResults(result) {
    resultsSection.classList.remove('hidden');
    seedValueEl.textContent = result.seed;

    // --- Validation list ---
    validationList.innerHTML = '';
    var allPass = true;
    for (var i = 0; i < result.validation.length; i++) {
      var v = result.validation[i];
      if (!v.status) allPass = false;
      var li = document.createElement('li');
      var icon = document.createElement('span');
      icon.className = v.status ? 'check-pass' : 'check-fail';
      icon.textContent = v.status ? 'PASS' : 'FAIL';
      var text = document.createElement('span');
      text.textContent = v.constraint + ' - ' + v.details;
      li.appendChild(icon);
      li.appendChild(text);
      validationList.appendChild(li);
    }

    // --- Pair lookup ---
    var pairFlatMap = {};
    for (var p = 0; p < result.pairs.length; p++) {
      pairFlatMap[result.pairs[p].flat1] = result.pairs[p].pairId;
      pairFlatMap[result.pairs[p].flat2] = result.pairs[p].pairId;
    }

    var flatLookup = {};
    for (var j = 0; j < parsedData.flats.length; j++) {
      flatLookup[parsedData.flats[j].oldFlatNo] = parsedData.flats[j];
    }

    // --- Allocation table ---
    allocTbody.innerHTML = '';
    var sorted = Object.keys(result.allocations).map(Number).sort(function(a, b) { return a - b; });
    for (var k = 0; k < sorted.length; k++) {
      var oldNo = sorted[k];
      var a = result.allocations[oldNo];
      var flat = flatLookup[oldNo] || {};
      var tr = document.createElement('tr');
      if (pairFlatMap[oldNo]) tr.className = 'pair-' + pairFlatMap[oldNo];
      tr.innerHTML =
        '<td>' + oldNo + '</td>' +
        '<td>' + escapeHtml(flat.ownerName || '') + '</td>' +
        '<td><strong>' + a.newFlatCode + '</strong></td>' +
        '<td>' + a.wing + '</td>' +
        '<td>' + a.floor + '</td>' +
        '<td>' + pad2(a.unit) + '</td>' +
        '<td>' + a.type + '</td>';
      allocTbody.appendChild(tr);
    }

    // --- Building layout (floor 10 → 1) ---
    layoutTbody.innerHTML = '';
    for (var fl = 10; fl >= 1; fl--) {
      var wingA  = findOldFlatByCode(result.allocations, 'A-' + pad2(fl) + '-01');
      var wingB1 = findOldFlatByCode(result.allocations, 'B-' + pad2(fl) + '-01');
      var wingB2 = findOldFlatByCode(result.allocations, 'B-' + pad2(fl) + '-02');

      var clsA  = wingA  && pairFlatMap[wingA]  ? 'pair-' + pairFlatMap[wingA]  : '';
      var clsB1 = wingB1 && pairFlatMap[wingB1] ? 'pair-' + pairFlatMap[wingB1] : '';
      var clsB2 = wingB2 && pairFlatMap[wingB2] ? 'pair-' + pairFlatMap[wingB2] : '';

      var row = document.createElement('tr');
      row.innerHTML =
        '<td class="floor-num">' + fl + '</td>' +
        '<td class="' + clsA  + '">' + (wingA  != null ? 'Flat ' + wingA  : '-') + '</td>' +
        '<td class="' + clsB1 + '">' + (wingB1 != null ? 'Flat ' + wingB1 : '-') + '</td>' +
        '<td class="' + clsB2 + '">' + (wingB2 != null ? 'Flat ' + wingB2 : '-') + '</td>';
      layoutTbody.appendChild(row);
    }

    showStatus(runStatus,
      allPass
        ? 'Allocation complete - all checks passed.'
        : 'Allocation complete - some checks FAILED.',
      allPass ? 'success' : 'error');

    // Scroll results into view
    document.getElementById('seed-display').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  // ======================== EXCEL OUTPUT ========================
  function generateOutput() {
    if (!allocationResult || !parsedData) return;
    var r = allocationResult;
    var wb = XLSX.utils.book_new();

    var flatLookup = {};
    for (var i = 0; i < parsedData.flats.length; i++) {
      flatLookup[parsedData.flats[i].oldFlatNo] = parsedData.flats[i];
    }

    var sorted = Object.keys(r.allocations).map(Number).sort(function(a, b) { return a - b; });

    // Sheet 1 — Allocation
    var s1 = [['Old Flat No.', 'Owner Name', 'New Flat Code', 'Wing', 'Floor', 'Unit', 'Allocation Type', 'Seed']];
    for (var j = 0; j < sorted.length; j++) {
      var o = sorted[j], a = r.allocations[o], fl = flatLookup[o] || {};
      s1.push([o, fl.ownerName || '', a.newFlatCode, a.wing, a.floor, a.unit, a.type, r.seed]);
    }
    var ws1 = XLSX.utils.aoa_to_sheet(s1);
    ws1['!cols'] = [{wch:12},{wch:22},{wch:15},{wch:8},{wch:8},{wch:8},{wch:16},{wch:14}];
    XLSX.utils.book_append_sheet(wb, ws1, 'Allocation');

    // Sheet 2 — Validation
    var s2 = [['Constraint', 'Status', 'Details']];
    for (var k = 0; k < r.validation.length; k++) {
      var v = r.validation[k];
      s2.push([v.constraint, v.status ? 'PASS' : 'FAIL', v.details]);
    }
    var ws2 = XLSX.utils.aoa_to_sheet(s2);
    ws2['!cols'] = [{wch:55},{wch:10},{wch:50}];
    XLSX.utils.book_append_sheet(wb, ws2, 'Validation');

    // Sheet 3 — Building Layout
    var s3 = [['Floor', 'Wing A (Old Flat No.)', 'Wing B Unit 1 (Old Flat No.)', 'Wing B Unit 2 (Old Flat No.)']];
    for (var f = 10; f >= 1; f--) {
      var wA  = findOldFlatByCode(r.allocations, 'A-' + pad2(f) + '-01');
      var wB1 = findOldFlatByCode(r.allocations, 'B-' + pad2(f) + '-01');
      var wB2 = findOldFlatByCode(r.allocations, 'B-' + pad2(f) + '-02');
      s3.push([f, wA != null ? wA : '', wB1 != null ? wB1 : '', wB2 != null ? wB2 : '']);
    }
    var ws3 = XLSX.utils.aoa_to_sheet(s3);
    ws3['!cols'] = [{wch:8},{wch:24},{wch:28},{wch:28}];
    XLSX.utils.book_append_sheet(wb, ws3, 'Building Layout');

    // Sheet 4 — Audit Trail
    var s4 = [['Step', 'Allocation Type', 'Old Flat No.', 'New Flat Code', 'Notes']];
    for (var t = 0; t < r.auditTrail.length; t++) {
      var at = r.auditTrail[t];
      s4.push([at.step, at.type, at.oldFlatNo, at.newFlatCode, at.notes]);
    }
    var ws4 = XLSX.utils.aoa_to_sheet(s4);
    ws4['!cols'] = [{wch:8},{wch:16},{wch:14},{wch:15},{wch:55}];
    XLSX.utils.book_append_sheet(wb, ws4, 'Audit Trail');

    // Sheet 5 — Metadata
    var allPass = r.validation.every(function(v) { return v.status; });
    var s5 = [
      ['Key', 'Value'],
      ['Generation Timestamp', new Date().toISOString()],
      ['Random Seed', r.seed],
      ['Total Flats', 30],
      ['Paired Allocations', r.pairs.length * 2],
      ['Random Allocations', 30 - r.pairs.length * 2],
      ['Allocation Status', allPass ? 'All checks passed' : 'Some checks failed']
    ];
    var ws5 = XLSX.utils.aoa_to_sheet(s5);
    ws5['!cols'] = [{wch:25},{wch:40}];
    XLSX.utils.book_append_sheet(wb, ws5, 'Metadata');

    XLSX.writeFile(wb, 'Flat_Allocation_RESULTS.xlsx');
  }

  // ======================== EVENT HANDLERS ========================
  runBtn.addEventListener('click', function() {
    if (!parsedData) return;

    var seed;
    if (parsedData.excelSeed != null) {
      seed = parsedData.excelSeed;
    } else {
      seed = Math.floor(Math.random() * 2147483647) + 1;
    }

    try {
      allocationResult = runAllocationAlgorithm(parsedData, seed);
      showResults(allocationResult);
      rerunBtn.classList.remove('hidden');
    } catch (err) {
      showStatus(runStatus, 'Allocation error: ' + err.message, 'error');
    }
  });

  rerunBtn.addEventListener('click', function() {
    if (!parsedData) return;
    var newSeed = Math.floor(Math.random() * 2147483647) + 1;
    try {
      allocationResult = runAllocationAlgorithm(parsedData, newSeed);
      showResults(allocationResult);
    } catch (err) {
      showStatus(runStatus, 'Allocation error: ' + err.message, 'error');
    }
  });

  downloadBtn.addEventListener('click', generateOutput);

};
window.__flatAppSource = __flatAppFn.toString();
__flatAppFn();

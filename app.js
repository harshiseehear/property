var __flatAppFn = function() {
  'use strict';

  // ======================== STATE ========================
  var parsedData = null;
  var allocationResult = null;

  // ======================== UTILITIES ========================
  function escapeHtml(str) {
    var d = document.createElement('div');
    d.appendChild(document.createTextNode(str));
    return d.innerHTML;
  }

  function pad2(n) { return n < 10 ? '0' + n : '' + n; }

  function makeFlatCode(wing, floor, unit) {
    return wing + '-' + pad2(floor) + '-' + pad2(unit);
  }

  function parseCode(code) {
    var parts = code.split('-');
    return { wing: parts[0], floor: parseInt(parts[1], 10), unit: parseInt(parts[2], 10) };
  }

  function showStatus(el, msg, type) {
    el.textContent = msg;
    el.className = 'status-msg visible ' + type;
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
  var dropZone       = document.getElementById('drop-zone');
  var fileInput      = document.getElementById('file-input');
  var fileInfo       = document.getElementById('file-info');
  var parseStatus    = document.getElementById('parse-status');
  var runBtn         = document.getElementById('run-btn');
  var rerunBtn       = document.getElementById('rerun-btn');
  var runStatus      = document.getElementById('run-status');
  var resultsSection = document.getElementById('results-section');
  var seedValueEl    = document.getElementById('seed-value');
  var validationList = document.getElementById('validation-list');
  var allocTbody     = document.querySelector('#allocation-table tbody');
  var layoutThead    = document.getElementById('layout-thead');
  var layoutTbody    = document.querySelector('#layout-table tbody');
  var downloadBtn    = document.getElementById('download-btn');

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

        var counts = { pair: 0, 'same-wing': 0, 'floor-pref': 0 };
        for (var c = 0; c < parsedData.constraints.length; c++) {
          counts[parsedData.constraints[c].type]++;
        }
        var parts = [];
        if (counts.pair > 0) parts.push(counts.pair + ' pair(s)');
        if (counts['same-wing'] > 0) parts.push(counts['same-wing'] + ' same-wing');
        if (counts['floor-pref'] > 0) parts.push(counts['floor-pref'] + ' floor-pref');

        fileInfo.innerHTML =
          '<strong>' + escapeHtml(file.name) + '</strong><br>' +
          'Old: ' + parsedData.totalOldFlats + ' flats | ' +
          'New: ' + parsedData.totalNewFlats + ' flats (' + parsedData.wingNames.length + ' wing(s): ' + parsedData.wingNames.join(', ') + ')<br>' +
          'Constraints: ' + (parts.length > 0 ? parts.join(', ') : 'none') +
          (parsedData.excelSeed != null ? ' | Seed: ' + parsedData.excelSeed : '');
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

    // ---- Old Building ----
    var oldSheet = wb.Sheets['Old Building'];
    if (!oldSheet) throw new Error('Sheet "Old Building" not found.');
    var oldRows = XLSX.utils.sheet_to_json(oldSheet, { header: 1 });
    if (oldRows.length < 2) throw new Error('"Old Building" needs a header row and data.');

    var oldBuilding = [];
    var totalOldFlats = 0;
    for (var oi = 1; oi < oldRows.length; oi++) {
      var orow = oldRows[oi];
      if (!orow || orow[0] == null || orow[0] === '') continue;
      var oWing = orow[0].toString().trim();
      var oFloor = parseInt(orow[1], 10);
      var oUnits = parseInt(orow[2], 10);
      if (isNaN(oFloor) || isNaN(oUnits) || oUnits <= 0) continue;
      oldBuilding.push({ wing: oWing, floor: oFloor, units: oUnits });
      totalOldFlats += oUnits;
    }
    if (totalOldFlats === 0) throw new Error('Old building has 0 flats.');

    // ---- New Building ----
    var newSheet = wb.Sheets['New Building'];
    if (!newSheet) throw new Error('Sheet "New Building" not found.');
    var newRows = XLSX.utils.sheet_to_json(newSheet, { header: 1 });
    if (newRows.length < 2) throw new Error('"New Building" needs a header row and data.');

    var newBuilding = [];
    var newWingFloorUnits = {};   // wing -> floor -> unitCount
    var allNewFlats = [];
    var wingMaxUnits = {};        // wing -> max units on any floor
    var allFloorsSet = {};
    var wingNames = [];
    var wingNameSet = {};

    for (var ni = 1; ni < newRows.length; ni++) {
      var nrow = newRows[ni];
      if (!nrow || nrow[0] == null || nrow[0] === '') continue;
      var nWing = nrow[0].toString().trim();
      var nFloor = parseInt(nrow[1], 10);
      var nUnits = parseInt(nrow[2], 10);
      if (isNaN(nFloor) || isNaN(nUnits) || nUnits <= 0) continue;
      newBuilding.push({ wing: nWing, floor: nFloor, units: nUnits });

      if (!newWingFloorUnits[nWing]) newWingFloorUnits[nWing] = {};
      newWingFloorUnits[nWing][nFloor] = nUnits;

      if (!wingMaxUnits[nWing] || nUnits > wingMaxUnits[nWing]) {
        wingMaxUnits[nWing] = nUnits;
      }
      allFloorsSet[nFloor] = true;
      if (!wingNameSet[nWing]) {
        wingNameSet[nWing] = true;
        wingNames.push(nWing);
      }
      for (var u = 1; u <= nUnits; u++) {
        allNewFlats.push(makeFlatCode(nWing, nFloor, u));
      }
    }
    wingNames.sort();

    var totalNewFlats = allNewFlats.length;
    if (totalNewFlats === 0) throw new Error('New building has 0 flats.');
    if (totalNewFlats < totalOldFlats) {
      throw new Error('New building (' + totalNewFlats + ') has fewer flats than old building (' + totalOldFlats + ').');
    }

    // ---- Flat Owners ----
    var ownerSheet = wb.Sheets['Flat Owners'];
    if (!ownerSheet) throw new Error('Sheet "Flat Owners" not found.');
    var ownerRows = XLSX.utils.sheet_to_json(ownerSheet, { header: 1 });
    if (ownerRows.length < 3) throw new Error('"Flat Owners" needs header rows + data.');

    var flats = [];
    for (var fi = 2; fi < ownerRows.length; fi++) {
      var frow = ownerRows[fi];
      if (!frow || frow.length === 0 || frow[0] == null || frow[0] === '') continue;
      var flatNo = parseInt(frow[0], 10);
      if (isNaN(flatNo)) continue;
      flats.push({
        oldFlatNo: flatNo,
        ownerName: (frow[1] || '').toString().trim(),
        contact:   (frow[2] || '').toString().trim(),
        constraintId: (frow[3] || '').toString().trim()
      });
    }
    if (flats.length !== totalOldFlats) {
      throw new Error('Expected ' + totalOldFlats + ' flats in "Flat Owners" (matching old building), found ' + flats.length + '.');
    }

    // ---- Constraints ----
    var constraints = {};
    var constraintSheet = wb.Sheets['Constraints'];
    if (constraintSheet) {
      var cRows = XLSX.utils.sheet_to_json(constraintSheet, { header: 1 });
      for (var ci = 1; ci < cRows.length; ci++) {
        var crow = cRows[ci];
        if (!crow || crow[0] == null || crow[0] === '') continue;
        var cId = crow[0].toString().trim();
        var cType = (crow[1] || '').toString().trim().toLowerCase();
        var cWing = (crow[2] || '').toString().trim();
        var cFloor = (crow[3] != null && crow[3] !== '') ? parseInt(crow[3], 10) : null;

        if (['pair', 'same-wing', 'floor-pref'].indexOf(cType) < 0) {
          throw new Error('Constraint "' + cId + '": unknown type "' + cType + '". Valid: pair, same-wing, floor-pref.');
        }
        if (cWing && !wingNameSet[cWing]) {
          throw new Error('Constraint "' + cId + '" references wing "' + cWing + '" not in new building.');
        }
        if (cFloor != null && !allFloorsSet[cFloor]) {
          throw new Error('Constraint "' + cId + '" references floor ' + cFloor + ' not in new building.');
        }

        constraints[cId] = { id: cId, type: cType, wing: cWing, floor: cFloor, flats: [] };
      }
    }

    // Link flats → constraints
    for (var li = 0; li < flats.length; li++) {
      var cid = flats[li].constraintId;
      if (cid && cid !== '') {
        if (!constraints[cid]) {
          throw new Error('Flat ' + flats[li].oldFlatNo + ' references constraint "' + cid + '" not in Constraints sheet.');
        }
        constraints[cid].flats.push(flats[li].oldFlatNo);
      }
    }

    // Validate membership counts
    var constraintList = [];
    var cKeys = Object.keys(constraints);
    for (var ck = 0; ck < cKeys.length; ck++) {
      var co = constraints[cKeys[ck]];
      if (co.type === 'pair' && co.flats.length !== 2) {
        throw new Error('"' + co.id + '" (pair) needs exactly 2 flats, found ' + co.flats.length + '.');
      }
      if (co.type === 'same-wing' && co.flats.length < 2) {
        throw new Error('"' + co.id + '" (same-wing) needs ≥2 flats, found ' + co.flats.length + '.');
      }
      if (co.type === 'floor-pref' && co.flats.length < 1) {
        throw new Error('"' + co.id + '" (floor-pref) needs ≥1 flat, found ' + co.flats.length + '.');
      }
      if (co.type === 'floor-pref' && (co.floor == null || isNaN(co.floor))) {
        throw new Error('"' + co.id + '" (floor-pref) must specify a floor number.');
      }
      co.flats.sort(function(a, b) { return a - b; });
      constraintList.push(co);
    }

    // ---- Randomisation Seed ----
    var excelSeed = null;
    var seedSheet = wb.Sheets['Randomisation Seed'];
    if (seedSheet) {
      var cell = seedSheet['B5'];
      if (cell && cell.v != null && cell.v !== '') {
        var sv = parseInt(cell.v, 10);
        if (!isNaN(sv)) excelSeed = sv;
      }
    }

    return {
      oldBuilding: oldBuilding,
      newBuilding: newBuilding,
      newWingFloorUnits: newWingFloorUnits,
      wingMaxUnits: wingMaxUnits,
      wingNames: wingNames,
      allFloors: Object.keys(allFloorsSet).map(Number).sort(function(a, b) { return a - b; }),
      allNewFlats: allNewFlats,
      totalOldFlats: totalOldFlats,
      totalNewFlats: totalNewFlats,
      flats: flats,
      constraints: constraintList,
      excelSeed: excelSeed
    };
  }

  // ======================== ALLOCATION ALGORITHM ========================
  function runAllocationAlgorithm(data, seed) {
    var rng = createRNG(seed);
    var auditTrail = [];
    var allocations = {};
    var stepNum = 0;
    var available = data.allNewFlats.slice();

    function removeAvail(code) {
      var idx = available.indexOf(code);
      if (idx >= 0) available.splice(idx, 1);
    }

    // Separate by type
    var pairs = [], sameWings = [], floorPrefs = [];
    for (var ci = 0; ci < data.constraints.length; ci++) {
      var c = data.constraints[ci];
      if (c.type === 'pair') pairs.push(c);
      else if (c.type === 'same-wing') sameWings.push(c);
      else if (c.type === 'floor-pref') floorPrefs.push(c);
    }

    // ---- PHASE 1: Pairs ----
    var usedPairFloorKeys = [];

    for (var pi = 0; pi < pairs.length; pi++) {
      var pair = pairs[pi];

      // Eligible wings
      var eligWings;
      if (pair.wing && pair.wing !== '') {
        eligWings = [pair.wing];
      } else {
        eligWings = [];
        for (var wn = 0; wn < data.wingNames.length; wn++) {
          if (data.wingMaxUnits[data.wingNames[wn]] >= 2) eligWings.push(data.wingNames[wn]);
        }
      }
      if (eligWings.length === 0) {
        throw new Error('No wing with ≥2 units/floor for "' + pair.id + '".');
      }

      // Find (wing, floor) slots with 2 consecutive available units
      var slots = [];
      for (var ew = 0; ew < eligWings.length; ew++) {
        var ewing = eligWings[ew];
        var wfMap = data.newWingFloorUnits[ewing];
        if (!wfMap) continue;
        var wfFloors = Object.keys(wfMap).map(Number);
        for (var ef = 0; ef < wfFloors.length; ef++) {
          var efl = wfFloors[ef];
          if (wfMap[efl] < 2) continue;
          var floorKey = ewing + '-' + efl;
          if (usedPairFloorKeys.indexOf(floorKey) >= 0) continue;

          for (var eu = 1; eu < wfMap[efl]; eu++) {
            var ec1 = makeFlatCode(ewing, efl, eu);
            var ec2 = makeFlatCode(ewing, efl, eu + 1);
            if (available.indexOf(ec1) >= 0 && available.indexOf(ec2) >= 0) {
              slots.push({ wing: ewing, floor: efl, code1: ec1, code2: ec2, unit1: eu, unit2: eu + 1 });
            }
          }
        }
      }

      if (slots.length === 0) {
        throw new Error('No floor with 2 consecutive available units for "' + pair.id + '".');
      }

      var chosen = rng.pick(slots);
      usedPairFloorKeys.push(chosen.wing + '-' + chosen.floor);

      allocations[pair.flats[0]] = {
        newFlatCode: chosen.code1, wing: chosen.wing, floor: chosen.floor,
        unit: chosen.unit1, type: 'Paired (' + pair.id + ')'
      };
      allocations[pair.flats[1]] = {
        newFlatCode: chosen.code2, wing: chosen.wing, floor: chosen.floor,
        unit: chosen.unit2, type: 'Paired (' + pair.id + ')'
      };
      removeAvail(chosen.code1);
      removeAvail(chosen.code2);

      stepNum++;
      auditTrail.push({
        step: stepNum, type: 'Paired', oldFlatNo: pair.flats[0], newFlatCode: chosen.code1,
        notes: pair.id + ': Wing ' + chosen.wing + ' Floor ' + chosen.floor +
               ' Unit ' + pad2(chosen.unit1) + ' (from ' + slots.length + ' eligible slot(s))'
      });
      stepNum++;
      auditTrail.push({
        step: stepNum, type: 'Paired', oldFlatNo: pair.flats[1], newFlatCode: chosen.code2,
        notes: pair.id + ': Wing ' + chosen.wing + ' Floor ' + chosen.floor + ' Unit ' + pad2(chosen.unit2)
      });
    }

    // ---- PHASE 2: Same-wing ----
    for (var si = 0; si < sameWings.length; si++) {
      var swc = sameWings[si];
      var swFlats = [];
      for (var swi = 0; swi < swc.flats.length; swi++) {
        if (!allocations[swc.flats[swi]]) swFlats.push(swc.flats[swi]);
      }
      if (swFlats.length === 0) continue;

      var targetWing;
      if (swc.wing && swc.wing !== '') {
        targetWing = swc.wing;
      } else {
        var wcounts = {};
        for (var aw = 0; aw < available.length; aw++) {
          var awW = parseCode(available[aw]).wing;
          wcounts[awW] = (wcounts[awW] || 0) + 1;
        }
        var elig = [];
        var wks = Object.keys(wcounts);
        for (var wk = 0; wk < wks.length; wk++) {
          if (wcounts[wks[wk]] >= swFlats.length) elig.push(wks[wk]);
        }
        if (elig.length === 0) {
          throw new Error('No wing has ' + swFlats.length + ' available flats for "' + swc.id + '".');
        }
        targetWing = rng.pick(elig);
      }

      var wingAvail = [];
      for (var wa = 0; wa < available.length; wa++) {
        if (parseCode(available[wa]).wing === targetWing) wingAvail.push(available[wa]);
      }
      if (wingAvail.length < swFlats.length) {
        throw new Error('Wing ' + targetWing + ' has ' + wingAvail.length + ' available but "' + swc.id + '" needs ' + swFlats.length + '.');
      }

      wingAvail = rng.shuffle(wingAvail);
      for (var sf = 0; sf < swFlats.length; sf++) {
        var swCode = wingAvail[sf];
        var swP = parseCode(swCode);
        allocations[swFlats[sf]] = {
          newFlatCode: swCode, wing: swP.wing, floor: swP.floor,
          unit: swP.unit, type: 'Same-wing (' + swc.id + ')'
        };
        removeAvail(swCode);
        stepNum++;
        auditTrail.push({
          step: stepNum, type: 'Same-wing', oldFlatNo: swFlats[sf], newFlatCode: swCode,
          notes: swc.id + ': Wing ' + targetWing
        });
      }
    }

    // ---- PHASE 3: Floor preferences ----
    for (var fp = 0; fp < floorPrefs.length; fp++) {
      var fpc = floorPrefs[fp];
      for (var ff = 0; ff < fpc.flats.length; ff++) {
        if (allocations[fpc.flats[ff]]) continue;

        var flAvail = [];
        for (var fa = 0; fa < available.length; fa++) {
          var faP = parseCode(available[fa]);
          if (faP.floor !== fpc.floor) continue;
          if (fpc.wing && fpc.wing !== '' && faP.wing !== fpc.wing) continue;
          flAvail.push(available[fa]);
        }
        if (flAvail.length === 0) {
          throw new Error('No available flat on floor ' + fpc.floor +
            (fpc.wing ? ' wing ' + fpc.wing : '') + ' for "' + fpc.id + '" (flat ' + fpc.flats[ff] + ').');
        }

        var fpPick = rng.pick(flAvail);
        var fpP = parseCode(fpPick);
        allocations[fpc.flats[ff]] = {
          newFlatCode: fpPick, wing: fpP.wing, floor: fpP.floor,
          unit: fpP.unit, type: 'Floor-pref (' + fpc.id + ')'
        };
        removeAvail(fpPick);
        stepNum++;
        auditTrail.push({
          step: stepNum, type: 'Floor-pref', oldFlatNo: fpc.flats[ff], newFlatCode: fpPick,
          notes: fpc.id + ': Floor ' + fpc.floor + ' (from ' + flAvail.length + ' available)'
        });
      }
    }

    // ---- PHASE 4: Random ----
    var assignedSet = {};
    for (var ak in allocations) assignedSet[ak] = true;
    var remaining = [];
    for (var ri = 0; ri < data.flats.length; ri++) {
      if (!assignedSet[data.flats[ri].oldFlatNo]) remaining.push(data.flats[ri].oldFlatNo);
    }

    remaining = rng.shuffle(remaining);
    for (var rm = 0; rm < remaining.length; rm++) {
      if (available.length === 0) {
        throw new Error('No available new flats for old flat ' + remaining[rm] + '.');
      }
      var rmIdx = rng.randomInt(0, available.length - 1);
      var rmPick = available[rmIdx];
      available.splice(rmIdx, 1);
      var rmP = parseCode(rmPick);
      allocations[remaining[rm]] = {
        newFlatCode: rmPick, wing: rmP.wing, floor: rmP.floor,
        unit: rmP.unit, type: 'Random'
      };
      stepNum++;
      auditTrail.push({
        step: stepNum, type: 'Random', oldFlatNo: remaining[rm], newFlatCode: rmPick,
        notes: 'Random (' + (available.length + 1) + ' were available)'
      });
    }

    var unoccupied = available.slice();
    var validation = validateAllocation(data, allocations, pairs, sameWings, floorPrefs);

    return {
      allocations: allocations,
      auditTrail: auditTrail,
      validation: validation,
      seed: seed,
      constraints: data.constraints,
      unoccupied: unoccupied
    };
  }

  // ======================== VALIDATION ========================
  function validateAllocation(data, allocations, pairs, sameWings, floorPrefs) {
    var checks = [];
    var allNew = [];
    for (var key in allocations) allNew.push(allocations[key].newFlatCode);

    // Pair checks
    var pairFloorKeys = [];
    for (var pi = 0; pi < pairs.length; pi++) {
      var pair = pairs[pi];
      var a1 = allocations[pair.flats[0]];
      var a2 = allocations[pair.flats[1]];
      var ok = a1 && a2 && a1.wing === a2.wing && a1.floor === a2.floor &&
               Math.abs(a1.unit - a2.unit) === 1;
      if (pair.wing && pair.wing !== '') ok = ok && a1 && a1.wing === pair.wing;

      pairFloorKeys.push(a1 ? a1.wing + '-' + a1.floor : 'N/A');
      checks.push({
        constraint: pair.id + ' (Flats ' + pair.flats.join(' & ') + '): same floor, adjacent' +
          (pair.wing ? ' in Wing ' + pair.wing : ''),
        status: !!ok,
        details: ok
          ? 'Wing ' + a1.wing + ' Floor ' + a1.floor + ' Units ' + pad2(a1.unit) + ' & ' + pad2(a2.unit)
          : 'FAILED'
      });
    }

    if (pairs.length >= 2) {
      var uniq = true;
      for (var i = 0; i < pairFloorKeys.length; i++) {
        for (var j = i + 1; j < pairFloorKeys.length; j++) {
          if (pairFloorKeys[i] === pairFloorKeys[j]) uniq = false;
        }
      }
      checks.push({
        constraint: 'All pairs on different wing-floor combinations',
        status: uniq,
        details: uniq ? 'Keys: ' + pairFloorKeys.join(', ') : 'FAILED — two pairs share a wing-floor'
      });
    }

    // Same-wing checks
    for (var si = 0; si < sameWings.length; si++) {
      var swc = sameWings[si];
      var swWings = [];
      for (var sf = 0; sf < swc.flats.length; sf++) {
        var swa = allocations[swc.flats[sf]];
        if (swa) swWings.push(swa.wing);
      }
      var allSame = swWings.length > 0 && swWings.every(function(w) { return w === swWings[0]; });
      var swOk = allSame && (!swc.wing || swc.wing === '' || swWings[0] === swc.wing);
      checks.push({
        constraint: swc.id + ' (Flats ' + swc.flats.join(', ') + '): same wing' +
          (swc.wing ? ' (' + swc.wing + ')' : ''),
        status: !!swOk,
        details: swOk ? 'All in Wing ' + swWings[0] : 'FAILED'
      });
    }

    // Floor-pref checks
    for (var fp = 0; fp < floorPrefs.length; fp++) {
      var fpc = floorPrefs[fp];
      for (var ff = 0; ff < fpc.flats.length; ff++) {
        var fpa = allocations[fpc.flats[ff]];
        var fpOk = fpa && fpa.floor === fpc.floor;
        if (fpc.wing && fpc.wing !== '') fpOk = fpOk && fpa && fpa.wing === fpc.wing;
        checks.push({
          constraint: fpc.id + ' (Flat ' + fpc.flats[ff] + '): floor ' + fpc.floor +
            (fpc.wing ? ', Wing ' + fpc.wing : ''),
          status: !!fpOk,
          details: fpOk ? 'Wing ' + fpa.wing + ' Floor ' + fpa.floor : 'FAILED'
        });
      }
    }

    // Global checks
    var newSet = new Set(allNew);
    checks.push({
      constraint: 'No duplicate new flat assignments',
      status: newSet.size === allNew.length,
      details: newSet.size === allNew.length
        ? 'All ' + allNew.length + ' unique'
        : 'FAILED — ' + (allNew.length - newSet.size) + ' duplicate(s)'
    });

    var oldCount = Object.keys(allocations).length;
    checks.push({
      constraint: 'All ' + data.totalOldFlats + ' old flats assigned',
      status: oldCount === data.totalOldFlats,
      details: oldCount + '/' + data.totalOldFlats + ' assigned'
    });

    var unoccCount = data.totalNewFlats - oldCount;
    checks.push({
      constraint: 'New flat occupancy',
      status: true,
      details: oldCount + ' occupied, ' + unoccCount + ' unoccupied (of ' + data.totalNewFlats + ' total)'
    });

    return checks;
  }

  // ======================== RENDER RESULTS ========================
  function buildConstraintMap() {
    var map = {};
    for (var c = 0; c < parsedData.constraints.length; c++) {
      var con = parsedData.constraints[c];
      for (var f = 0; f < con.flats.length; f++) {
        map[con.flats[f]] = con.id;
      }
    }
    return map;
  }

  function showResults(result) {
    resultsSection.classList.remove('hidden');
    seedValueEl.textContent = result.seed;
    var constraintMap = buildConstraintMap();

    var flatLookup = {};
    for (var f = 0; f < parsedData.flats.length; f++) {
      flatLookup[parsedData.flats[f].oldFlatNo] = parsedData.flats[f];
    }

    // Validation list
    validationList.innerHTML = '';
    var allPass = true;
    for (var vi = 0; vi < result.validation.length; vi++) {
      var v = result.validation[vi];
      if (!v.status) allPass = false;
      var li = document.createElement('li');
      var icon = document.createElement('span');
      icon.className = v.status ? 'check-pass' : 'check-fail';
      icon.textContent = v.status ? 'PASS' : 'FAIL';
      var text = document.createElement('span');
      text.textContent = v.constraint + ' — ' + v.details;
      li.appendChild(icon);
      li.appendChild(text);
      validationList.appendChild(li);
    }

    // Allocation table
    allocTbody.innerHTML = '';
    var sorted = Object.keys(result.allocations).map(Number).sort(function(a, b) { return a - b; });
    for (var ai = 0; ai < sorted.length; ai++) {
      var oldNo = sorted[ai];
      var a = result.allocations[oldNo];
      var flat = flatLookup[oldNo] || {};
      var tr = document.createElement('tr');
      if (constraintMap[oldNo]) tr.className = 'constrained';
      tr.innerHTML =
        '<td>' + oldNo + '</td>' +
        '<td>' + escapeHtml(flat.ownerName || '') + '</td>' +
        '<td><strong>' + a.newFlatCode + '</strong></td>' +
        '<td>' + a.wing + '</td>' +
        '<td>' + a.floor + '</td>' +
        '<td>' + pad2(a.unit) + '</td>' +
        '<td>' + escapeHtml(a.type) + '</td>';
      allocTbody.appendChild(tr);
    }

    // Dynamic building layout
    var wings = parsedData.wingNames;
    var floors = parsedData.allFloors.slice().sort(function(a, b) { return b - a; });

    // Build header
    layoutThead.innerHTML = '';
    var headerRow = document.createElement('tr');
    var thFloor = document.createElement('th');
    thFloor.textContent = 'Floor';
    headerRow.appendChild(thFloor);

    var layoutColumns = [];
    for (var wi = 0; wi < wings.length; wi++) {
      var w = wings[wi];
      var maxU = parsedData.wingMaxUnits[w];
      for (var ui = 1; ui <= maxU; ui++) {
        layoutColumns.push({ wing: w, unit: ui });
        var th = document.createElement('th');
        th.textContent = 'Wing ' + w + (maxU > 1 ? ', Unit ' + pad2(ui) : '');
        headerRow.appendChild(th);
      }
    }
    layoutThead.appendChild(headerRow);

    // Reverse lookup
    var codeToOld = {};
    for (var rk in result.allocations) {
      codeToOld[result.allocations[rk].newFlatCode] = parseInt(rk, 10);
    }

    // Build body
    layoutTbody.innerHTML = '';
    for (var fi = 0; fi < floors.length; fi++) {
      var fl = floors[fi];
      var row = document.createElement('tr');
      var tdFloor = document.createElement('td');
      tdFloor.className = 'floor-num';
      tdFloor.textContent = fl;
      row.appendChild(tdFloor);

      for (var lc = 0; lc < layoutColumns.length; lc++) {
        var col = layoutColumns[lc];
        var td = document.createElement('td');
        var wfUnits = parsedData.newWingFloorUnits[col.wing];
        var floorUnits = wfUnits ? wfUnits[fl] : 0;

        if (!floorUnits || col.unit > floorUnits) {
          td.textContent = '-';
          td.className = 'no-unit';
        } else {
          var code = makeFlatCode(col.wing, fl, col.unit);
          var oldFlat = codeToOld[code];
          if (oldFlat != null) {
            td.textContent = 'Flat ' + oldFlat;
            if (constraintMap[oldFlat]) td.className = 'constrained';
          } else {
            td.textContent = 'Unoccupied';
            td.className = 'unoccupied';
          }
        }
        row.appendChild(td);
      }
      layoutTbody.appendChild(row);
    }

    showStatus(runStatus,
      allPass
        ? 'Allocation complete — all checks passed.'
        : 'Allocation complete — some checks FAILED.',
      allPass ? 'success' : 'error');

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
    var s1 = [['Old Flat No.', 'Owner Name', 'New Flat Code', 'Wing', 'Floor', 'Unit', 'Allocation Type', 'Constraint', 'Seed']];
    for (var j = 0; j < sorted.length; j++) {
      var o = sorted[j], a = r.allocations[o], fl = flatLookup[o] || {};
      s1.push([o, fl.ownerName || '', a.newFlatCode, a.wing, a.floor, a.unit, a.type, fl.constraintId || '', r.seed]);
    }
    var ws1 = XLSX.utils.aoa_to_sheet(s1);
    ws1['!cols'] = [{wch:12},{wch:22},{wch:15},{wch:8},{wch:8},{wch:8},{wch:24},{wch:14},{wch:14}];
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

    // Sheet 3 — Building Layout (dynamic)
    var wings = parsedData.wingNames;
    var floors = parsedData.allFloors.slice().sort(function(a, b) { return b - a; });
    var lCols = [];
    var lHeader = ['Floor'];
    for (var lw = 0; lw < wings.length; lw++) {
      var w = wings[lw];
      var maxU = parsedData.wingMaxUnits[w];
      for (var lu = 1; lu <= maxU; lu++) {
        lCols.push({ wing: w, unit: lu });
        lHeader.push('Wing ' + w + (maxU > 1 ? ' Unit ' + pad2(lu) : '') + ' (Old Flat)');
      }
    }
    var s3 = [lHeader];
    var codeToOld = {};
    for (var rk in r.allocations) {
      codeToOld[r.allocations[rk].newFlatCode] = parseInt(rk, 10);
    }
    for (var lf = 0; lf < floors.length; lf++) {
      var lfn = floors[lf];
      var lrow = [lfn];
      for (var lci = 0; lci < lCols.length; lci++) {
        var lc = lCols[lci];
        var wfU = parsedData.newWingFloorUnits[lc.wing];
        var fUnits = wfU ? wfU[lfn] : 0;
        if (!fUnits || lc.unit > fUnits) {
          lrow.push('');
        } else {
          var code = makeFlatCode(lc.wing, lfn, lc.unit);
          var oldF = codeToOld[code];
          lrow.push(oldF != null ? oldF : 'Unoccupied');
        }
      }
      s3.push(lrow);
    }
    var ws3 = XLSX.utils.aoa_to_sheet(s3);
    var s3cols = [{wch:8}];
    for (var sc = 0; sc < lCols.length; sc++) s3cols.push({wch:24});
    ws3['!cols'] = s3cols;
    XLSX.utils.book_append_sheet(wb, ws3, 'Building Layout');

    // Sheet 4 — Audit Trail
    var s4 = [['Step', 'Type', 'Old Flat No.', 'New Flat Code', 'Notes']];
    for (var t = 0; t < r.auditTrail.length; t++) {
      var at = r.auditTrail[t];
      s4.push([at.step, at.type, at.oldFlatNo, at.newFlatCode, at.notes]);
    }
    var ws4 = XLSX.utils.aoa_to_sheet(s4);
    ws4['!cols'] = [{wch:8},{wch:16},{wch:14},{wch:15},{wch:60}];
    XLSX.utils.book_append_sheet(wb, ws4, 'Audit Trail');

    // Sheet 5 — Metadata
    var allPass = r.validation.every(function(v) { return v.status; });
    var cCounts = { pair: 0, 'same-wing': 0, 'floor-pref': 0 };
    for (var mc = 0; mc < parsedData.constraints.length; mc++) {
      cCounts[parsedData.constraints[mc].type]++;
    }
    var s5 = [
      ['Key', 'Value'],
      ['Generation Timestamp', new Date().toISOString()],
      ['Random Seed', r.seed],
      ['Old Building Flats', parsedData.totalOldFlats],
      ['New Building Flats', parsedData.totalNewFlats],
      ['Wings', parsedData.wingNames.join(', ')],
      ['Pair Constraints', cCounts.pair],
      ['Same-wing Constraints', cCounts['same-wing']],
      ['Floor-pref Constraints', cCounts['floor-pref']],
      ['Unoccupied New Flats', r.unoccupied.length],
      ['Status', allPass ? 'All checks passed' : 'Some checks failed']
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

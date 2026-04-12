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

        var counts = { group: 0, pref: 0 };
        for (var c = 0; c < parsedData.constraints.length; c++) {
          counts[parsedData.constraints[c].type]++;
        }
        var parts = [];
        if (counts.group > 0) parts.push(counts.group + ' group(s)');
        if (counts.pref > 0) parts.push(counts.pref + ' pref(s)');

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
        var cUnit = (crow[4] != null && crow[4] !== '') ? parseInt(crow[4], 10) : null;

        if (['group', 'pref'].indexOf(cType) < 0) {
          throw new Error('Constraint "' + cId + '": unknown type "' + cType + '". Valid: group, pref.');
        }
        if (constraints[cId]) {
          throw new Error('Duplicate constraint ID "' + cId + '" found at row ' + (ci + 1) + '. Each ID must appear only once.');
        }
        if (cWing && !wingNameSet[cWing]) {
          throw new Error('Constraint "' + cId + '" references wing "' + cWing + '" not in new building.');
        }
        if (cFloor != null && !allFloorsSet[cFloor]) {
          throw new Error('Constraint "' + cId + '" references floor ' + cFloor + ' not in new building.');
        }

        constraints[cId] = { id: cId, type: cType, wing: cWing, floor: cFloor, unit: cUnit, flats: [] };
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
      if (co.type === 'group' && (co.flats.length < 2 || co.flats.length > 3)) {
        throw new Error('"' + co.id + '" (group) needs 2 or 3 flats, found ' + co.flats.length + '.');
      }
      if (co.type === 'pref' && co.flats.length < 1) {
        throw new Error('"' + co.id + '" (pref) needs ≥1 flat, found ' + co.flats.length + '.');
      }
      if (co.type === 'pref' && !co.wing && co.floor == null && co.unit == null) {
        throw new Error('"' + co.id + '" (pref) must specify at least a wing, floor, or unit.');
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
    var groups = [], prefs = [];
    for (var ci = 0; ci < data.constraints.length; ci++) {
      var c = data.constraints[ci];
      if (c.type === 'group') groups.push(c);
      else if (c.type === 'pref') prefs.push(c);
    }

    // Split groups by size: 3-flat groups first (more constrained), then 2-flat groups
    var groups3 = [], groups2 = [];
    for (var gi = 0; gi < groups.length; gi++) {
      if (groups[gi].flats.length === 3) groups3.push(groups[gi]);
      else groups2.push(groups[gi]);
    }

    // Sort prefs by specificity: most specific first
    prefs.sort(function(a, b) {
      var sa = (a.wing ? 1 : 0) + (a.floor != null ? 1 : 0) + (a.unit != null ? 1 : 0);
      var sb = (b.wing ? 1 : 0) + (b.floor != null ? 1 : 0) + (b.unit != null ? 1 : 0);
      return sb - sa;
    });

    // ---- PHASE 1A: Groups of 3 (2 adjacent on floor X, 1 on floor X+1) ----
    var usedGroupFloorKeys = [];

    for (var g3i = 0; g3i < groups3.length; g3i++) {
      var grp3 = groups3[g3i];

      // Eligible wings: need a floor with ≥2 units AND a floor+1 with ≥1 unit
      var eligWings3;
      if (grp3.wing && grp3.wing !== '') {
        eligWings3 = [grp3.wing];
      } else {
        eligWings3 = [];
        for (var wn3 = 0; wn3 < data.wingNames.length; wn3++) {
          if (data.wingMaxUnits[data.wingNames[wn3]] >= 2) eligWings3.push(data.wingNames[wn3]);
        }
      }
      if (eligWings3.length === 0) {
        throw new Error('No wing with ≥2 units/floor for "' + grp3.id + '".');
      }

      // Find slots: (wing, floorX) with 2 consecutive available units + (wing, floorX+1) with ≥1 available unit
      var slots3 = [];
      for (var ew3 = 0; ew3 < eligWings3.length; ew3++) {
        var ewing3 = eligWings3[ew3];
        var wfMap3 = data.newWingFloorUnits[ewing3];
        if (!wfMap3) continue;
        var wfFloors3 = Object.keys(wfMap3).map(Number).sort(function(a, b) { return a - b; });
        for (var ef3 = 0; ef3 < wfFloors3.length; ef3++) {
          var floorX = wfFloors3[ef3];
          if (wfMap3[floorX] < 2) continue;
          if (grp3.floor != null && floorX !== grp3.floor) continue;
          var floorXp1 = floorX + 1;
          if (!wfMap3[floorXp1] || wfMap3[floorXp1] < 1) continue;
          var keyX = ewing3 + '-' + floorX;
          var keyXp1 = ewing3 + '-' + floorXp1;
          if (usedGroupFloorKeys.indexOf(keyX) >= 0 || usedGroupFloorKeys.indexOf(keyXp1) >= 0) continue;

          // Find consecutive available units on floorX
          for (var eu3 = 1; eu3 < wfMap3[floorX]; eu3++) {
            var ec3a = makeFlatCode(ewing3, floorX, eu3);
            var ec3b = makeFlatCode(ewing3, floorX, eu3 + 1);
            if (available.indexOf(ec3a) < 0 || available.indexOf(ec3b) < 0) continue;

            // Find available units on floorX+1
            for (var eu3c = 1; eu3c <= wfMap3[floorXp1]; eu3c++) {
              var ec3c = makeFlatCode(ewing3, floorXp1, eu3c);
              if (available.indexOf(ec3c) >= 0) {
                slots3.push({
                  wing: ewing3, floorA: floorX, floorB: floorXp1,
                  code1: ec3a, code2: ec3b, code3: ec3c,
                  unit1: eu3, unit2: eu3 + 1, unit3: eu3c
                });
              }
            }
          }
        }
      }

      if (slots3.length === 0) {
        throw new Error('No consecutive floors with enough available units for "' + grp3.id + '" (needs 2 adjacent on floor X + 1 on floor X+1).');
      }

      var chosen3 = rng.pick(slots3);
      usedGroupFloorKeys.push(chosen3.wing + '-' + chosen3.floorA);
      usedGroupFloorKeys.push(chosen3.wing + '-' + chosen3.floorB);

      allocations[grp3.flats[0]] = {
        newFlatCode: chosen3.code1, wing: chosen3.wing, floor: chosen3.floorA,
        unit: chosen3.unit1, type: 'Group (' + grp3.id + ')'
      };
      allocations[grp3.flats[1]] = {
        newFlatCode: chosen3.code2, wing: chosen3.wing, floor: chosen3.floorA,
        unit: chosen3.unit2, type: 'Group (' + grp3.id + ')'
      };
      allocations[grp3.flats[2]] = {
        newFlatCode: chosen3.code3, wing: chosen3.wing, floor: chosen3.floorB,
        unit: chosen3.unit3, type: 'Group (' + grp3.id + ')'
      };
      removeAvail(chosen3.code1);
      removeAvail(chosen3.code2);
      removeAvail(chosen3.code3);

      stepNum++;
      auditTrail.push({
        step: stepNum, type: 'Grouped', oldFlatNo: grp3.flats[0], newFlatCode: chosen3.code1,
        notes: grp3.id + ': Wing ' + chosen3.wing + ' Floor ' + chosen3.floorA +
               ' Unit ' + pad2(chosen3.unit1) + ' (from ' + slots3.length + ' eligible slot(s))'
      });
      stepNum++;
      auditTrail.push({
        step: stepNum, type: 'Grouped', oldFlatNo: grp3.flats[1], newFlatCode: chosen3.code2,
        notes: grp3.id + ': Wing ' + chosen3.wing + ' Floor ' + chosen3.floorA + ' Unit ' + pad2(chosen3.unit2)
      });
      stepNum++;
      auditTrail.push({
        step: stepNum, type: 'Grouped', oldFlatNo: grp3.flats[2], newFlatCode: chosen3.code3,
        notes: grp3.id + ': Wing ' + chosen3.wing + ' Floor ' + chosen3.floorB + ' Unit ' + pad2(chosen3.unit3)
      });
    }

    // ---- PHASE 1B: Groups of 2 (adjacent on same floor) ----
    for (var g2i = 0; g2i < groups2.length; g2i++) {
      var grp2 = groups2[g2i];

      // Eligible wings
      var eligWings2;
      if (grp2.wing && grp2.wing !== '') {
        eligWings2 = [grp2.wing];
      } else {
        eligWings2 = [];
        for (var wn2 = 0; wn2 < data.wingNames.length; wn2++) {
          if (data.wingMaxUnits[data.wingNames[wn2]] >= 2) eligWings2.push(data.wingNames[wn2]);
        }
      }
      if (eligWings2.length === 0) {
        throw new Error('No wing with ≥2 units/floor for "' + grp2.id + '".');
      }

      // Find (wing, floor) slots with 2 consecutive available units
      var slots2 = [];
      for (var ew2 = 0; ew2 < eligWings2.length; ew2++) {
        var ewing2 = eligWings2[ew2];
        var wfMap2 = data.newWingFloorUnits[ewing2];
        if (!wfMap2) continue;
        var wfFloors2 = Object.keys(wfMap2).map(Number);
        for (var ef2 = 0; ef2 < wfFloors2.length; ef2++) {
          var efl2 = wfFloors2[ef2];
          if (wfMap2[efl2] < 2) continue;
          if (grp2.floor != null && efl2 !== grp2.floor) continue;
          var floorKey2 = ewing2 + '-' + efl2;
          if (usedGroupFloorKeys.indexOf(floorKey2) >= 0) continue;

          for (var eu2 = 1; eu2 < wfMap2[efl2]; eu2++) {
            var ec2a = makeFlatCode(ewing2, efl2, eu2);
            var ec2b = makeFlatCode(ewing2, efl2, eu2 + 1);
            if (available.indexOf(ec2a) >= 0 && available.indexOf(ec2b) >= 0) {
              slots2.push({ wing: ewing2, floor: efl2, code1: ec2a, code2: ec2b, unit1: eu2, unit2: eu2 + 1 });
            }
          }
        }
      }

      if (slots2.length === 0) {
        throw new Error('No floor with 2 consecutive available units for "' + grp2.id + '".');
      }

      var chosen2 = rng.pick(slots2);
      usedGroupFloorKeys.push(chosen2.wing + '-' + chosen2.floor);

      allocations[grp2.flats[0]] = {
        newFlatCode: chosen2.code1, wing: chosen2.wing, floor: chosen2.floor,
        unit: chosen2.unit1, type: 'Group (' + grp2.id + ')'
      };
      allocations[grp2.flats[1]] = {
        newFlatCode: chosen2.code2, wing: chosen2.wing, floor: chosen2.floor,
        unit: chosen2.unit2, type: 'Group (' + grp2.id + ')'
      };
      removeAvail(chosen2.code1);
      removeAvail(chosen2.code2);

      stepNum++;
      auditTrail.push({
        step: stepNum, type: 'Grouped', oldFlatNo: grp2.flats[0], newFlatCode: chosen2.code1,
        notes: grp2.id + ': Wing ' + chosen2.wing + ' Floor ' + chosen2.floor +
               ' Unit ' + pad2(chosen2.unit1) + ' (from ' + slots2.length + ' eligible slot(s))'
      });
      stepNum++;
      auditTrail.push({
        step: stepNum, type: 'Grouped', oldFlatNo: grp2.flats[1], newFlatCode: chosen2.code2,
        notes: grp2.id + ': Wing ' + chosen2.wing + ' Floor ' + chosen2.floor + ' Unit ' + pad2(chosen2.unit2)
      });
    }

    // ---- PHASE 2: Preferences (wing and/or floor) ----
    for (var pr = 0; pr < prefs.length; pr++) {
      var prc = prefs[pr];
      for (var pf = 0; pf < prc.flats.length; pf++) {
        if (allocations[prc.flats[pf]]) continue;

        var prAvail = [];
        for (var pa = 0; pa < available.length; pa++) {
          var paP = parseCode(available[pa]);
          if (prc.wing && prc.wing !== '' && paP.wing !== prc.wing) continue;
          if (prc.floor != null && paP.floor !== prc.floor) continue;
          if (prc.unit != null && paP.unit !== prc.unit) continue;
          prAvail.push(available[pa]);
        }
        if (prAvail.length === 0) {
          var prefDesc = [];
          if (prc.wing) prefDesc.push('Wing ' + prc.wing);
          if (prc.floor != null) prefDesc.push('Floor ' + prc.floor);
          if (prc.unit != null) prefDesc.push('Unit ' + pad2(prc.unit));
          throw new Error('No available flat matching ' + prefDesc.join(', ') +
            ' for "' + prc.id + '" (flat ' + prc.flats[pf] + ').');
        }

        var prPick = rng.pick(prAvail);
        var prP = parseCode(prPick);
        allocations[prc.flats[pf]] = {
          newFlatCode: prPick, wing: prP.wing, floor: prP.floor,
          unit: prP.unit, type: 'Pref (' + prc.id + ')'
        };
        removeAvail(prPick);
        stepNum++;
        var prefNotes = prc.id + ':';
        if (prc.wing) prefNotes += ' Wing ' + prc.wing;
        if (prc.floor != null) prefNotes += ' Floor ' + prc.floor;
        if (prc.unit != null) prefNotes += ' Unit ' + pad2(prc.unit);
        prefNotes += ' (from ' + prAvail.length + ' available)';
        auditTrail.push({
          step: stepNum, type: 'Pref', oldFlatNo: prc.flats[pf], newFlatCode: prPick,
          notes: prefNotes
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
    var validation = validateAllocation(data, allocations, groups, prefs);

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
  function validateAllocation(data, allocations, groups, prefs) {
    var checks = [];
    var allNew = [];
    for (var key in allocations) allNew.push(allocations[key].newFlatCode);

    // Group checks
    var groupFloorKeys = [];
    for (var gi = 0; gi < groups.length; gi++) {
      var grp = groups[gi];

      if (grp.flats.length === 2) {
        // Group of 2: same wing, same floor, adjacent units
        var a1 = allocations[grp.flats[0]];
        var a2 = allocations[grp.flats[1]];
        var ok2 = a1 && a2 && a1.wing === a2.wing && a1.floor === a2.floor &&
                  Math.abs(a1.unit - a2.unit) === 1;
        if (grp.wing && grp.wing !== '') ok2 = ok2 && a1 && a1.wing === grp.wing;
        if (grp.floor != null) ok2 = ok2 && a1 && a1.floor === grp.floor;

        groupFloorKeys.push(a1 ? a1.wing + '-' + a1.floor : 'N/A');
        var grpDesc2 = grp.id + ' (Flats ' + grp.flats.join(' & ') + '): same floor, adjacent';
        if (grp.wing) grpDesc2 += ', Wing ' + grp.wing;
        if (grp.floor != null) grpDesc2 += ', Floor ' + grp.floor;
        checks.push({
          constraint: grpDesc2,
          status: !!ok2,
          details: ok2
            ? 'Wing ' + a1.wing + ' Floor ' + a1.floor + ' Units ' + pad2(a1.unit) + ' & ' + pad2(a2.unit)
            : 'FAILED'
        });
      } else if (grp.flats.length === 3) {
        // Group of 3: flats[0] & flats[1] same wing, same floor, adjacent; flats[2] same wing, floor+1
        var b1 = allocations[grp.flats[0]];
        var b2 = allocations[grp.flats[1]];
        var b3 = allocations[grp.flats[2]];
        var ok3 = b1 && b2 && b3 &&
                  b1.wing === b2.wing && b1.wing === b3.wing &&
                  b1.floor === b2.floor &&
                  Math.abs(b1.unit - b2.unit) === 1 &&
                  b3.floor === b1.floor + 1;
        if (grp.wing && grp.wing !== '') ok3 = ok3 && b1 && b1.wing === grp.wing;
        if (grp.floor != null) ok3 = ok3 && b1 && b1.floor === grp.floor;

        groupFloorKeys.push(b1 ? b1.wing + '-' + b1.floor : 'N/A');
        groupFloorKeys.push(b3 ? b3.wing + '-' + b3.floor : 'N/A');
        var grpDesc3 = grp.id + ' (Flats ' + grp.flats.join(', ') + '): 2 adjacent on floor X, 1 on floor X+1';
        if (grp.wing) grpDesc3 += ', Wing ' + grp.wing;
        if (grp.floor != null) grpDesc3 += ', Floor ' + grp.floor;
        checks.push({
          constraint: grpDesc3,
          status: !!ok3,
          details: ok3
            ? 'Wing ' + b1.wing + ' Floors ' + b1.floor + ' & ' + b3.floor +
              ' Units ' + pad2(b1.unit) + ' & ' + pad2(b2.unit) + ', ' + pad2(b3.unit)
            : 'FAILED'
        });
      }
    }

    if (groupFloorKeys.length >= 2) {
      var uniq = true;
      for (var i = 0; i < groupFloorKeys.length; i++) {
        for (var j = i + 1; j < groupFloorKeys.length; j++) {
          if (groupFloorKeys[i] !== 'N/A' && groupFloorKeys[i] === groupFloorKeys[j]) uniq = false;
        }
      }
      checks.push({
        constraint: 'All groups on different wing-floor combinations',
        status: uniq,
        details: uniq ? 'Keys: ' + groupFloorKeys.join(', ') : 'FAILED — two groups share a wing-floor'
      });
    }

    // Pref checks
    for (var pr = 0; pr < prefs.length; pr++) {
      var prc = prefs[pr];
      for (var pf = 0; pf < prc.flats.length; pf++) {
        var pra = allocations[prc.flats[pf]];
        var prOk = !!pra;
        if (prOk && prc.wing && prc.wing !== '') prOk = pra.wing === prc.wing;
        if (prOk && prc.floor != null) prOk = pra.floor === prc.floor;
        if (prOk && prc.unit != null) prOk = pra.unit === prc.unit;
        var prDesc = prc.id + ' (Flat ' + prc.flats[pf] + '):';
        if (prc.wing) prDesc += ' Wing ' + prc.wing;
        if (prc.floor != null) prDesc += ' Floor ' + prc.floor;
        if (prc.unit != null) prDesc += ' Unit ' + pad2(prc.unit);
        checks.push({
          constraint: prDesc,
          status: !!prOk,
          details: prOk ? 'Wing ' + pra.wing + ' Floor ' + pra.floor + ' Unit ' + pad2(pra.unit) : 'FAILED'
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
    var cCounts = { group: 0, pref: 0 };
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
      ['Group Constraints', cCounts.group],
      ['Pref Constraints', cCounts.pref],
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

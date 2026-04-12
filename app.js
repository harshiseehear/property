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
      if (co.type === 'group' && co.flats.length < 2) {
        throw new Error('"' + co.id + '" (group) needs ≥2 flats, found ' + co.flats.length + '.');
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

    // Sort groups by size descending (larger groups first = more constrained)
    groups.sort(function(a, b) { return b.flats.length - a.flats.length; });

    // Sort prefs by specificity: most specific first
    prefs.sort(function(a, b) {
      var sa = (a.wing ? 1 : 0) + (a.floor != null ? 1 : 0) + (a.unit != null ? 1 : 0);
      var sb = (b.wing ? 1 : 0) + (b.floor != null ? 1 : 0) + (b.unit != null ? 1 : 0);
      return sb - sa;
    });

    // ---- PHASE 1: Groups of N (consecutive adjacent units, filling floors then overflowing up) ----
    var usedGroupFloorKeys = [];

    for (var gri = 0; gri < groups.length; gri++) {
      var grp = groups[gri];
      var grpSize = grp.flats.length;

      // Eligible wings
      var eligWings;
      if (grp.wing && grp.wing !== '') {
        eligWings = [grp.wing];
      } else {
        eligWings = [];
        for (var wni = 0; wni < data.wingNames.length; wni++) {
          eligWings.push(data.wingNames[wni]);
        }
      }

      // Find all valid placement slots using a stack-based search
      var groupSlots = [];
      for (var ewi = 0; ewi < eligWings.length; ewi++) {
        var ewing = eligWings[ewi];
        var wfMap = data.newWingFloorUnits[ewing];
        if (!wfMap) continue;
        var wfFloors = Object.keys(wfMap).map(Number).sort(function(a, b) { return a - b; });

        for (var efi = 0; efi < wfFloors.length; efi++) {
          var startFloor = wfFloors[efi];
          if (grp.floor != null && startFloor !== grp.floor) continue;

          // Stack-based search: each state = { remaining, floor, partial }
          var stack = [{ remaining: grpSize, floor: startFloor, partial: [] }];

          while (stack.length > 0) {
            var state = stack.pop();

            if (state.remaining <= 0) {
              groupSlots.push({ wing: ewing, placements: state.partial });
              continue;
            }

            if (!wfMap[state.floor]) continue;
            var floorKey = ewing + '-' + state.floor;
            if (usedGroupFloorKeys.indexOf(floorKey) >= 0) continue;

            // Find consecutive runs of available units on this floor
            var maxUnits = wfMap[state.floor];
            var runs = [];
            var runStart = -1;
            for (var u = 1; u <= maxUnits; u++) {
              var uCode = makeFlatCode(ewing, state.floor, u);
              if (available.indexOf(uCode) >= 0) {
                if (runStart === -1) runStart = u;
              } else {
                if (runStart !== -1) {
                  runs.push({ start: runStart, len: u - runStart });
                  runStart = -1;
                }
              }
            }
            if (runStart !== -1) {
              runs.push({ start: runStart, len: maxUnits - runStart + 1 });
            }

            for (var ri = 0; ri < runs.length; ri++) {
              var run = runs[ri];
              var take = Math.min(run.len, state.remaining);
              var codes = [];
              var units = [];
              for (var ti = 0; ti < take; ti++) {
                units.push(run.start + ti);
                codes.push(makeFlatCode(ewing, state.floor, run.start + ti));
              }

              var newPartial = state.partial.concat([{ floor: state.floor, units: units, codes: codes }]);
              var newRemaining = state.remaining - take;

              if (newRemaining === 0) {
                groupSlots.push({ wing: ewing, placements: newPartial });
              } else {
                // Overflow to next consecutive floor
                stack.push({ remaining: newRemaining, floor: state.floor + 1, partial: newPartial });
              }
            }
          }
        }
      }

      if (groupSlots.length === 0) {
        throw new Error('No valid placement found for group "' + grp.id + '" (' + grpSize +
          ' flats). Need consecutive adjacent units across consecutive floors in the same wing.');
      }

      var chosenSlot = rng.pick(groupSlots);

      // Mark all used floor-keys
      for (var pi = 0; pi < chosenSlot.placements.length; pi++) {
        usedGroupFloorKeys.push(chosenSlot.wing + '-' + chosenSlot.placements[pi].floor);
      }

      // Flatten placement into ordered codes/units/floors
      var allCodes = [];
      var allUnits = [];
      var allFloors = [];
      for (var pli = 0; pli < chosenSlot.placements.length; pli++) {
        var pl = chosenSlot.placements[pli];
        for (var pci = 0; pci < pl.codes.length; pci++) {
          allCodes.push(pl.codes[pci]);
          allUnits.push(pl.units[pci]);
          allFloors.push(pl.floor);
        }
      }

      // Assign flats to group members
      for (var mi = 0; mi < grp.flats.length; mi++) {
        allocations[grp.flats[mi]] = {
          newFlatCode: allCodes[mi], wing: chosenSlot.wing, floor: allFloors[mi],
          unit: allUnits[mi], type: 'Group (' + grp.id + ')'
        };
        removeAvail(allCodes[mi]);

        stepNum++;
        var gNotes = grp.id + ': Wing ' + chosenSlot.wing + ' Floor ' + allFloors[mi] + ' Unit ' + pad2(allUnits[mi]);
        if (mi === 0) gNotes += ' (from ' + groupSlots.length + ' eligible slot(s))';
        auditTrail.push({
          step: stepNum, type: 'Grouped', oldFlatNo: grp.flats[mi], newFlatCode: allCodes[mi],
          notes: gNotes
        });
      }
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

      // Collect all allocations for this group's members
      var memberAllocs = [];
      var allMemValid = true;
      for (var gmi = 0; gmi < grp.flats.length; gmi++) {
        var gma = allocations[grp.flats[gmi]];
        if (!gma) { allMemValid = false; break; }
        memberAllocs.push(gma);
      }

      var okN = allMemValid;

      // All must be in the same wing
      if (okN) {
        for (var gwi = 1; gwi < memberAllocs.length; gwi++) {
          if (memberAllocs[gwi].wing !== memberAllocs[0].wing) { okN = false; break; }
        }
      }

      // Group members by floor, check floors are consecutive, and units are consecutive on each floor
      var byFloor = {};
      if (okN) {
        for (var gfi = 0; gfi < memberAllocs.length; gfi++) {
          var gfl = memberAllocs[gfi].floor;
          if (!byFloor[gfl]) byFloor[gfl] = [];
          byFloor[gfl].push(memberAllocs[gfi].unit);
        }

        var floorNums = Object.keys(byFloor).map(Number).sort(function(a, b) { return a - b; });
        for (var gci = 1; gci < floorNums.length; gci++) {
          if (floorNums[gci] !== floorNums[gci - 1] + 1) { okN = false; break; }
        }

        if (okN) {
          for (var gui = 0; gui < floorNums.length; gui++) {
            var flUnits = byFloor[floorNums[gui]].sort(function(a, b) { return a - b; });
            for (var guj = 1; guj < flUnits.length; guj++) {
              if (flUnits[guj] !== flUnits[guj - 1] + 1) { okN = false; break; }
            }
            if (!okN) break;
          }
        }
      }

      // Check wing/floor constraints
      if (okN && grp.wing && grp.wing !== '') {
        okN = memberAllocs[0].wing === grp.wing;
      }
      if (okN && grp.floor != null) {
        var minGrpFloor = memberAllocs[0].floor;
        for (var gmfi = 1; gmfi < memberAllocs.length; gmfi++) {
          if (memberAllocs[gmfi].floor < minGrpFloor) minGrpFloor = memberAllocs[gmfi].floor;
        }
        okN = minGrpFloor === grp.floor;
      }

      // Collect floor keys for uniqueness check
      var memberFloorSet = {};
      for (var gki = 0; gki < memberAllocs.length; gki++) {
        memberFloorSet[memberAllocs[gki].wing + '-' + memberAllocs[gki].floor] = true;
      }
      var mfKeys = Object.keys(memberFloorSet);
      for (var gmk = 0; gmk < mfKeys.length; gmk++) {
        groupFloorKeys.push(mfKeys[gmk]);
      }

      var grpDescN = grp.id + ' (Flats ' + grp.flats.join(', ') + '): ' + grp.flats.length + ' members, consecutive adjacent';
      if (grp.wing) grpDescN += ', Wing ' + grp.wing;
      if (grp.floor != null) grpDescN += ', Floor ' + grp.floor;

      var detailStr = 'FAILED';
      if (okN && memberAllocs.length > 0) {
        var usedFloors = Object.keys(byFloor).map(Number).sort(function(a, b) { return a - b; });
        var unitStrs = [];
        for (var gdi = 0; gdi < memberAllocs.length; gdi++) {
          unitStrs.push(pad2(memberAllocs[gdi].unit));
        }
        detailStr = 'Wing ' + memberAllocs[0].wing + ' Floor(s) ' + usedFloors.join(', ') + ' Units ' + unitStrs.join(', ');
      }

      checks.push({
        constraint: grpDescN,
        status: !!okN,
        details: detailStr
      });
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

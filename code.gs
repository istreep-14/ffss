/**
 * Fantasy Football Value Based Drafting (VBD) Calculator
 * Calculates VOR (Value Over Replacement) and VOLS (Value Over Last Starter)
 * Properly handles SF before FLEX allocation and supports decimal bench values
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Fantasy VBD')
    .addItem('Setup Input Sheet', 'setupInputSheet')
    .addItem('Calculate VBD', 'calculateVBD')
    .addToUi();
}

function setupInputSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create or get the Input sheet
  let inputSheet = ss.getSheetByName('Input');
  if (!inputSheet) {
    inputSheet = ss.insertSheet('Input');
  }
  
  // Clear existing content
  inputSheet.clear();
  
  // Set up the input table
  const headers = [
    ['Position Settings', '', ''],
    ['Position', 'Starters', 'Bench (can use decimals)'],
    ['QB', 1, 1],
    ['RB', 2, 2.5],
    ['WR', 2, 3],
    ['TE', 1, 1],
    ['FLEX', 1, 0],
    ['SF', 1, 0],
    ['DST', 1, 0.5],
    ['K', 1, 0.5],
    ['', '', ''],
    ['League Settings', '', ''],
    ['Number of Teams', 12, ''],
    ['', '', ''],
    ['Results will appear below when you run Calculate VBD', '', '']
  ];
  
  inputSheet.getRange(1, 1, headers.length, 3).setValues(headers);
  
  // Format the sheet
  inputSheet.getRange('A1:C1').merge().setFontWeight('bold').setHorizontalAlignment('center');
  inputSheet.getRange('A2:C2').setFontWeight('bold');
  inputSheet.getRange('A12:C12').merge().setFontWeight('bold').setHorizontalAlignment('center');
  inputSheet.getRange('A13:B13').setFontWeight('bold');
  
  // Auto-resize columns
  inputSheet.autoResizeColumns(1, 3);
  
  SpreadsheetApp.getUi().alert('Input sheet has been set up! You can use decimals for bench (e.g., 2.5 RBs). Adjust settings as needed, then run "Calculate VBD".');
}

function calculateVBD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playersSheet = ss.getSheetByName('Players');
  const inputSheet = ss.getSheetByName('Input');
  
  if (!playersSheet) {
    SpreadsheetApp.getUi().alert('Players sheet not found! Please create a sheet named "Players" with columns: Rank, Name, Team, Pos, FPTS');
    return;
  }
  
  if (!inputSheet) {
    SpreadsheetApp.getUi().alert('Input sheet not found! Please run "Setup Input Sheet" first.');
    return;
  }
  
  try {
    // Get input data
    const inputData = getInputData(inputSheet);
    const playerData = getPlayerData(playersSheet);
    
    // Calculate starters needed and replacement levels
    const calculations = calculateReplacementLevels(playerData, inputData);
    
    // Add VOR and VOLS columns to Players sheet
    addVBDColumns(playersSheet, playerData, calculations);
    
    // Display results on Input sheet
    displayResults(inputSheet, calculations, inputData);
    
    SpreadsheetApp.getUi().alert('VBD calculations complete! Check the Players sheet for VOR and VOLS columns, and the Input sheet for detailed position analysis.');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
    console.error(error);
  }
}

function getInputData(inputSheet) {
  const data = inputSheet.getDataRange().getValues();
  const positionSettings = {};
  let numTeams = 12;
  
  // Parse position settings (rows 3-10, accounting for 0-based indexing)
  for (let i = 2; i < 10; i++) {
    if (data[i] && data[i][0]) {
      const pos = data[i][0];
      const starters = parseFloat(data[i][1]) || 0;
      const bench = parseFloat(data[i][2]) || 0;
      
      positionSettings[pos] = {
        starters: starters,
        bench: bench
      };
    }
  }
  
  // Get number of teams
  if (data[12] && data[12][1]) {
    numTeams = parseInt(data[12][1]) || 12;
  }
  
  return { positionSettings, numTeams };
}

function getPlayerData(playersSheet) {
  const data = playersSheet.getDataRange().getValues();
  const headers = data[0];
  const players = [];
  
  // Find column indices
  const rankCol = headers.indexOf('Rank');
  const nameCol = headers.indexOf('Name');
  const teamCol = headers.indexOf('Team');
  const posCol = headers.indexOf('Pos');
  const fptsCol = headers.indexOf('FPTS');
  
  if (nameCol === -1 || posCol === -1 || fptsCol === -1) {
    throw new Error('Required columns not found. Please ensure you have: Name, Pos, FPTS (Rank and Team are optional)');
  }
  
  // Process player data
  for (let i = 1; i < data.length; i++) {
    if (data[i][nameCol] && data[i][posCol] && data[i][fptsCol]) {
      players.push({
        rank: data[i][rankCol] || i,
        name: data[i][nameCol],
        team: data[i][teamCol] || '',
        position: data[i][posCol],
        fpts: parseFloat(data[i][fptsCol]) || 0,
        row: i + 1
      });
    }
  }
  
  return players;
}

function calculateReplacementLevels(players, inputData) {
  const { positionSettings, numTeams } = inputData;
  
  // Group players by their TRUE position (not roster position)
  const playersByPosition = {};
  players.forEach(player => {
    if (!playersByPosition[player.position]) {
      playersByPosition[player.position] = [];
    }
    playersByPosition[player.position].push(player);
  });
  
  // Sort players by FPTS within each position (descending)
  Object.keys(playersByPosition).forEach(pos => {
    playersByPosition[pos].sort((a, b) => b.fpts - a.fpts);
  });
  
  // Calculate direct starter needs first
  const directStarters = {};
  const corePositions = ['QB', 'RB', 'WR', 'TE', 'DST', 'K'];
  
  corePositions.forEach(pos => {
    directStarters[pos] = Math.round((positionSettings[pos]?.starters || 0) * numTeams);
  });
  
  // Track total starters needed (will include SF and FLEX allocations)
  const totalStarters = { ...directStarters };
  
  // STEP 1: Calculate SF allocation FIRST (since QBs can only go here)
  const sfEligible = ['QB', 'RB', 'WR', 'TE'];
  const sfCount = Math.round((positionSettings.SF?.starters || 0) * numTeams);
  const sfAllocation = allocateSuperflexPositions(playersByPosition, sfEligible, directStarters, sfCount);
  
  // Add SF allocations to total starters
  Object.keys(sfAllocation).forEach(pos => {
    totalStarters[pos] = (totalStarters[pos] || 0) + sfAllocation[pos];
  });
  
  // STEP 2: Calculate FLEX allocation AFTER SF (RB/WR/TE only)
  const flexEligible = ['RB', 'WR', 'TE'];
  const flexCount = Math.round((positionSettings.FLEX?.starters || 0) * numTeams);
  const flexAllocation = allocateFlexPositions(playersByPosition, flexEligible, totalStarters, flexCount);
  
  // Add FLEX allocations to total starters
  Object.keys(flexAllocation).forEach(pos => {
    totalStarters[pos] = (totalStarters[pos] || 0) + flexAllocation[pos];
  });
  
  // Calculate bench needs (supporting decimals)
  const totalBench = {};
  corePositions.forEach(pos => {
    const benchDecimal = positionSettings[pos]?.bench || 0;
    totalBench[pos] = Math.round(benchDecimal * numTeams);
  });
  
  // Calculate replacement and last starter levels for TRUE positions only
  const replacementLevels = {};
  const lastStarterLevels = {};
  
  corePositions.forEach(pos => {
    const posPlayers = playersByPosition[pos] || [];
    const startersNeeded = totalStarters[pos] || 0;
    const benchNeeded = totalBench[pos] || 0;
    
    // Last starter level (VOLS baseline)
    if (startersNeeded > 0 && startersNeeded <= posPlayers.length) {
      lastStarterLevels[pos] = posPlayers[startersNeeded - 1].fpts;
    } else {
      lastStarterLevels[pos] = posPlayers.length > 0 ? posPlayers[posPlayers.length - 1].fpts : 0;
    }
    
    // Replacement level (VOR baseline) - after starters + bench
    const replacementIndex = startersNeeded + benchNeeded;
    if (replacementIndex < posPlayers.length) {
      replacementLevels[pos] = posPlayers[replacementIndex].fpts;
    } else {
      // If we don't have enough players, use the last available player
      replacementLevels[pos] = posPlayers.length > 0 ? posPlayers[posPlayers.length - 1].fpts : 0;
    }
  });
  
  return {
    directStarters,
    totalStarters,
    totalBench,
    replacementLevels,
    lastStarterLevels,
    sfAllocation,
    flexAllocation,
    playersByPosition
  };
}

function allocateSuperflexPositions(playersByPosition, eligiblePositions, currentStarters, sfCount) {
  const allocation = {};
  eligiblePositions.forEach(pos => allocation[pos] = 0);
  
  if (sfCount === 0) return allocation;
  
  // Create pool of remaining players for SF, prioritizing QBs
  const remainingPlayers = [];
  
  eligiblePositions.forEach(pos => {
    if (playersByPosition[pos]) {
      const startIndex = currentStarters[pos] || 0;
      for (let i = startIndex; i < playersByPosition[pos].length; i++) {
        remainingPlayers.push({
          ...playersByPosition[pos][i],
          originalPosition: pos
        });
      }
    }
  });
  
  // Sort by FPTS (descending) - this naturally prioritizes top QBs
  remainingPlayers.sort((a, b) => b.fpts - a.fpts);
  
  // Allocate top players to SF
  for (let i = 0; i < Math.min(sfCount, remainingPlayers.length); i++) {
    allocation[remainingPlayers[i].originalPosition]++;
  }
  
  return allocation;
}

function allocateFlexPositions(playersByPosition, eligiblePositions, currentStarters, flexCount) {
  const allocation = {};
  eligiblePositions.forEach(pos => allocation[pos] = 0);
  
  if (flexCount === 0) return allocation;
  
  // Create pool of remaining players for FLEX (after SF allocation)
  const remainingPlayers = [];
  
  eligiblePositions.forEach(pos => {
    if (playersByPosition[pos]) {
      const startIndex = currentStarters[pos] || 0;
      for (let i = startIndex; i < playersByPosition[pos].length; i++) {
        remainingPlayers.push({
          ...playersByPosition[pos][i],
          originalPosition: pos
        });
      }
    }
  });
  
  // Sort by FPTS (descending)
  remainingPlayers.sort((a, b) => b.fpts - a.fpts);
  
  // Allocate top remaining players to FLEX
  for (let i = 0; i < Math.min(flexCount, remainingPlayers.length); i++) {
    allocation[remainingPlayers[i].originalPosition]++;
  }
  
  return allocation;
}

function addVBDColumns(playersSheet, players, calculations) {
  const data = playersSheet.getDataRange().getValues();
  const headers = data[0];
  
  // Check if VOR and VOLS columns exist, if not add them
  let vorCol = headers.indexOf('VOR');
  let volsCol = headers.indexOf('VOLS');
  
  if (vorCol === -1) {
    headers.push('VOR');
    vorCol = headers.length - 1;
  }
  
  if (volsCol === -1) {
    headers.push('VOLS');
    volsCol = headers.length - 1;
  }
  
  // Update headers
  playersSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Calculate and add VOR and VOLS for each player based on their TRUE position
  players.forEach(player => {
    const truePosition = player.position;
    const replacementLevel = calculations.replacementLevels[truePosition] || 0;
    const lastStarterLevel = calculations.lastStarterLevels[truePosition] || 0;
    
    const vor = player.fpts - replacementLevel;
    const vols = player.fpts - lastStarterLevel;
    
    // Update the row with VOR and VOLS
    playersSheet.getRange(player.row, vorCol + 1).setValue(Math.round(vor * 10) / 10);
    playersSheet.getRange(player.row, volsCol + 1).setValue(Math.round(vols * 10) / 10);
  });
  
  // Format VOR and VOLS columns
  const lastRow = playersSheet.getLastRow();
  if (lastRow > 1) {
    playersSheet.getRange(2, vorCol + 1, lastRow - 1, 1).setNumberFormat('0.0');
    playersSheet.getRange(2, volsCol + 1, lastRow - 1, 1).setNumberFormat('0.0');
  }
}

function displayResults(inputSheet, calculations, inputData) {
  const { directStarters, totalStarters, totalBench, replacementLevels, lastStarterLevels, sfAllocation, flexAllocation } = calculations;
  const { numTeams } = inputData;
  
  // Clear previous results
  const lastRow = inputSheet.getLastRow();
  if (lastRow > 15) {
    inputSheet.getRange(16, 1, lastRow - 15, 20).clear();
  }
  
  const results = [
    [''],
    ['CALCULATION RESULTS'],
    [''],
    ['Position Allocation Analysis:'],
    ['Position', 'Direct', 'SF', 'FLEX', 'Total Starters', 'Bench', 'Replacement Level', 'Last Starter Level']
  ];
  
  const corePositions = ['QB', 'RB', 'WR', 'TE', 'DST', 'K'];
  
  corePositions.forEach(pos => {
    const direct = directStarters[pos] || 0;
    const sf = sfAllocation[pos] || 0;
    const flex = flexAllocation[pos] || 0;
    const total = totalStarters[pos] || 0;
    const bench = totalBench[pos] || 0;
    const replacement = replacementLevels[pos] || 0;
    const lastStarter = lastStarterLevels[pos] || 0;
    
    results.push([
      pos, 
      direct, 
      sf, 
      flex, 
      total, 
      bench,
      replacement.toFixed(1), 
      lastStarter.toFixed(1)
    ]);
  });
  
  results.push(['']);
  results.push(['Allocation Details:']);
  results.push(['']);
  
  // SF Allocation Details
  results.push(['SF Positions (' + Math.round((inputData.positionSettings.SF?.starters || 0) * numTeams) + ' total):']);
  let sfTotal = 0;
  Object.keys(sfAllocation).forEach(pos => {
    if (sfAllocation[pos] > 0) {
      results.push([`  ${pos}:`, sfAllocation[pos]]);
      sfTotal += sfAllocation[pos];
    }
  });
  results.push(['  SF Total:', sfTotal]);
  
  results.push(['']);
  
  // FLEX Allocation Details  
  results.push(['FLEX Positions (' + Math.round((inputData.positionSettings.FLEX?.starters || 0) * numTeams) + ' total):']);
  let flexTotal = 0;
  Object.keys(flexAllocation).forEach(pos => {
    if (flexAllocation[pos] > 0) {
      results.push([`  ${pos}:`, flexAllocation[pos]]);
      flexTotal += flexAllocation[pos];
    }
  });
  results.push(['  FLEX Total:', flexTotal]);
  
  results.push(['']);
  results.push(['Bench Calculation (decimals rounded to nearest whole):']);
  corePositions.forEach(pos => {
    const benchDecimal = inputData.positionSettings[pos]?.bench || 0;
    const benchTotal = totalBench[pos] || 0;
    if (benchDecimal > 0) {
      results.push([`${pos}:`, `${benchDecimal} × ${numTeams} = ${benchTotal}`]);
    }
  });
  
  // Write results to sheet
  const startRow = 16;
  const maxCols = Math.max(...results.map(row => row.length));
  
  // Pad shorter rows with empty strings
  const paddedResults = results.map(row => {
    while (row.length < maxCols) {
      row.push('');
    }
    return row;
  });
  
  inputSheet.getRange(startRow, 1, paddedResults.length, maxCols).setValues(paddedResults);
  
  // Format headers and important rows
  inputSheet.getRange(startRow + 1, 1, 1, maxCols).setFontWeight('bold').setHorizontalAlignment('center');
  inputSheet.getRange(startRow + 4, 1, 1, maxCols).setFontWeight('bold');
  inputSheet.getRange(startRow + 11, 1, 1, maxCols).setFontWeight('bold');
  inputSheet.getRange(startRow + 14, 1, 1, maxCols).setFontWeight('bold');
  inputSheet.getRange(startRow + 21, 1, 1, maxCols).setFontWeight('bold');
  
  // Auto-resize columns
  inputSheet.autoResizeColumns(1, maxCols);
}

/**
 * Helper function to parse decimal/fraction input
 */
function parseDecimalInput(value) {
  if (typeof value === 'number') {
    return value;
  }
  
  if (typeof value === 'string') {
    // Handle fractions like "2.5" or "2 1/2"
    if (value.includes('/')) {
      // Simple fraction parsing
      const parts = value.split('/');
      if (parts.length === 2) {
        return parseFloat(parts[0]) / parseFloat(parts[1]);
      }
    }
    return parseFloat(value) || 0;
  }
  
  return 0;
}

/**
 * Function to manually set position requirements with decimal support
 * Usage: setPositionRequirements({QB: {starters: 1, bench: 1.5}, RB: {starters: 2, bench: 2.5}, ...})
 */
function setPositionRequirements(settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let inputSheet = ss.getSheetByName('Input');
  
  if (!inputSheet) {
    setupInputSheet();
    inputSheet = ss.getSheetByName('Input');
  }
  
  // Update the position settings
  const positions = ['QB', 'RB', 'WR', 'TE', 'FLEX', 'SF', 'DST', 'K'];
  
  positions.forEach((pos, index) => {
    const row = 3 + index;
    if (settings[pos]) {
      inputSheet.getRange(row, 2).setValue(settings[pos].starters);
      inputSheet.getRange(row, 3).setValue(settings[pos].bench);
    }
  });
}

/**
 * Debug function to test calculations
 */
function debugCalculations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Input');
  const playersSheet = ss.getSheetByName('Players');
  
  if (!inputSheet || !playersSheet) {
    console.log('Missing required sheets');
    return;
  }
  
  const inputData = getInputData(inputSheet);
  const playerData = getPlayerData(playersSheet);
  const calculations = calculateReplacementLevels(playerData, inputData);
  
  console.log('Input Data:', inputData);
  console.log('Total Starters:', calculations.totalStarters);
  console.log('SF Allocation:', calculations.sfAllocation);
  console.log('FLEX Allocation:', calculations.flexAllocation);
  console.log('Replacement Levels:', calculations.replacementLevels);
  console.log('Last Starter Levels:', calculations.lastStarterLevels);
}

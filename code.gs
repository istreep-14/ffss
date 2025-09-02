/**
 * Fantasy Football Value Based Drafting (VBD) Calculator
 * Calculates VOR (Value Over Replacement) and VOLS (Value Over Last Starter)
 * Properly handles SF before FLEX allocation and supports decimal bench values
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Fantasy VBD menu
  ui.createMenu('Fantasy VBD')
    .addItem('Setup Input Sheet', 'setupInputSheet')
    .addItem('Calculate VBD', 'calculateVBD')
    .addItem('Calculate VONA', 'showVONADialog')
    .addToUi();
    
  // Draft Board menu
  ui.createMenu('Draft Board')
    .addItem('Setup Helper Tab', 'setupHelperTab')
    .addItem('Setup Draft Board', 'setupDraftBoard')
    .addItem('Refresh Rosters', 'refreshRosters')
    .addSeparator()
    .addItem('Test Roster Size Calculation', 'testRosterSizeCalculation')
    .addItem('Help', 'showHelp')
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
  
  // Minimal formatting - just headers
  inputSheet.getRange('A2:C2').setFontWeight('bold');
  inputSheet.getRange('A13:B13').setFontWeight('bold');
  
  // Set fixed column widths instead of auto-resize
  inputSheet.setColumnWidth(1, 150); // Position column
  inputSheet.setColumnWidth(2, 80);  // Starters column
  inputSheet.setColumnWidth(3, 180); // Bench column
  
  SpreadsheetApp.getUi().alert('Input sheet has been set up! You can use decimals for bench (e.g., 2.5 RBs). Adjust settings as needed, then run "Calculate VBD".');
}

function setupHelperTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create or get the Helper sheet
  let helperSheet = ss.getSheetByName('Helper');
  if (!helperSheet) {
    helperSheet = ss.insertSheet('Helper');
  }
  
  // Clear existing content
  helperSheet.clear();
  
  // Set up the configuration table
  const config = [
    ['Roster Size', 16],
    ['Teams', 12],
    ['Players', '=B1*B2']
  ];
  
  // Set values
  helperSheet.getRange(1, 1, config.length, 2).setValues(config);
  
  // Format headers
  helperSheet.getRange('A1:A3').setFontWeight('bold');
  
  // Add formula for players calculation
  helperSheet.getRange('B3').setFormula('=B1*B2');
  
  // Set column widths
  helperSheet.setColumnWidth(1, 120); // Label column
  helperSheet.setColumnWidth(2, 80);  // Value column
  
  // Add border around the configuration area
  helperSheet.getRange('A1:B3').setBorder(true, true, true, true, true, true);
  
  SpreadsheetApp.getUi().alert('Helper tab has been set up! Adjust roster size and teams as needed.');
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

/**
 * Shows a dialog to input pick number for VONA calculation
 */
function showVONADialog() {
  const ui = SpreadsheetApp.getUi();
  
  // Get current pick number from draft board if available
  const defaultPick = getCurrentPickNumber();
  
  const result = ui.prompt(
    'Calculate VONA (Value Over Next Available)',
    'Enter the pick number to calculate VONA for:\n' +
    '(This should be your next pick or any future pick)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (result.getSelectedButton() === ui.Button.OK) {
    const pickNumber = parseInt(result.getResponseText());
    
    if (isNaN(pickNumber) || pickNumber < 1) {
      ui.alert('Invalid pick number. Please enter a positive number.');
      return;
    }
    
    calculateVONA(pickNumber);
  }
}

/**
 * Gets the current pick number from draft board if available
 */
function getCurrentPickNumber() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const draftSheet = ss.getSheetByName('draft');
  
  if (!draftSheet) return 1;
  
  try {
    const data = draftSheet.getDataRange().getValues();
    let pickCount = 0;
    
    // Count non-empty cells in the draft board
    for (let row = 1; row < data.length; row++) {
      for (let col = 1; col < data[row].length; col++) {
        if (data[row][col] && data[row][col].toString().trim()) {
          pickCount++;
        }
      }
    }
    
    return pickCount + 1;
  } catch (e) {
    return 1;
  }
}

/**
 * Main VONA calculation function
 */
function calculateVONA(pickNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playersSheet = ss.getSheetByName('Players');
  const inputSheet = ss.getSheetByName('Input');
  const draftSheet = ss.getSheetByName('draft');
  
  if (!playersSheet || !inputSheet) {
    SpreadsheetApp.getUi().alert('Required sheets not found! Please ensure you have "Players" and "Input" sheets set up.');
    return;
  }
  
  try {
    // Get all necessary data
    const inputData = getInputData(inputSheet);
    const playerData = getPlayerData(playersSheet);
    const draftedPlayers = getDraftedPlayers(draftSheet);
    
    // Filter out drafted players
    const availablePlayers = playerData.filter(player => 
      !draftedPlayers.some(drafted => 
        drafted.toLowerCase() === player.name.toLowerCase()
      )
    );
    
    // Calculate VONA for each available player
    const vonaResults = calculateVONAForPlayers(
      availablePlayers, 
      pickNumber, 
      inputData,
      draftedPlayers.length
    );
    
    // Display results
    displayVONAResults(vonaResults, pickNumber);
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error calculating VONA: ' + error.message);
    console.error(error);
  }
}

/**
 * Get list of drafted players from draft board
 */
function getDraftedPlayers(draftSheet) {
  const draftedPlayers = [];
  
  if (!draftSheet) return draftedPlayers;
  
  const data = draftSheet.getDataRange().getValues();
  
  // Skip header row and collect all non-empty cells
  for (let row = 1; row < data.length; row++) {
    for (let col = 1; col < data[row].length; col++) {
      const playerName = data[row][col];
      if (playerName && playerName.toString().trim()) {
        draftedPlayers.push(playerName.toString().trim());
      }
    }
  }
  
  return draftedPlayers;
}

/**
 * Calculate VONA for all available players
 */
function calculateVONAForPlayers(availablePlayers, targetPick, inputData, currentPick) {
  const { numTeams } = inputData;
  const picksUntilTarget = targetPick - currentPick - 1;
  
  // Group players by position
  const playersByPosition = {};
  availablePlayers.forEach(player => {
    if (!playersByPosition[player.position]) {
      playersByPosition[player.position] = [];
    }
    playersByPosition[player.position].push(player);
  });
  
  // Sort each position by FPTS descending
  Object.keys(playersByPosition).forEach(pos => {
    playersByPosition[pos].sort((a, b) => b.fpts - a.fpts);
  });
  
  // Calculate expected available value at target pick for each position
  const expectedValueByPosition = {};
  const positions = ['QB', 'RB', 'WR', 'TE', 'DST', 'K'];
  
  positions.forEach(pos => {
    const posPlayers = playersByPosition[pos] || [];
    
    // Estimate how many players of this position will be drafted
    const draftRate = getPositionDraftRate(pos, currentPick, numTeams);
    const expectedDrafted = Math.floor(picksUntilTarget * draftRate);
    
    // Find the expected available player at target pick
    const expectedIndex = Math.min(expectedDrafted, posPlayers.length - 1);
    
    if (expectedIndex >= 0 && expectedIndex < posPlayers.length) {
      // Average the value around the expected index for more stability
      const startIdx = Math.max(0, expectedIndex - 1);
      const endIdx = Math.min(posPlayers.length - 1, expectedIndex + 1);
      let totalValue = 0;
      let count = 0;
      
      for (let i = startIdx; i <= endIdx; i++) {
        totalValue += posPlayers[i].fpts;
        count++;
      }
      
      expectedValueByPosition[pos] = totalValue / count;
    } else {
      expectedValueByPosition[pos] = 0;
    }
  });
  
  // Calculate VONA for each player
  const results = availablePlayers.map(player => {
    const expectedValue = expectedValueByPosition[player.position] || 0;
    const vona = player.fpts - expectedValue;
    
    return {
      ...player,
      vona: vona,
      expectedValue: expectedValue,
      positionScarcity: calculatePositionScarcity(
        player.position, 
        playersByPosition[player.position], 
        picksUntilTarget
      )
    };
  });
  
  // Sort by VONA descending
  results.sort((a, b) => b.vona - a.vona);
  
  // Calculate position gaps (value of waiting)
  const positionGaps = calculatePositionGaps(
    playersByPosition, 
    expectedValueByPosition,
    picksUntilTarget
  );
  
  return {
    players: results,
    expectedValues: expectedValueByPosition,
    positionGaps: positionGaps
  };
}

/**
 * Estimate draft rate for each position based on historical data
 */
function getPositionDraftRate(position, currentPick, numTeams) {
  // These are approximate draft rates based on typical fantasy drafts
  // Can be adjusted based on league tendencies
  const baseRates = {
    'RB': 0.30,  // 30% of picks are RBs
    'WR': 0.28,  // 28% of picks are WRs
    'QB': 0.15,  // 15% of picks are QBs
    'TE': 0.12,  // 12% of picks are TEs
    'DST': 0.08, // 8% of picks are DSTs
    'K': 0.07    // 7% of picks are Ks
  };
  
  // Adjust rates based on draft stage
  const round = Math.ceil(currentPick / numTeams);
  let rate = baseRates[position] || 0.1;
  
  // Early rounds favor RB/WR
  if (round <= 3) {
    if (position === 'RB' || position === 'WR') {
      rate *= 1.3;
    } else if (position === 'K' || position === 'DST') {
      rate *= 0.2;
    }
  }
  // Middle rounds see more QBs and TEs
  else if (round <= 8) {
    if (position === 'QB' || position === 'TE') {
      rate *= 1.2;
    }
  }
  // Late rounds see more K/DST
  else {
    if (position === 'K' || position === 'DST') {
      rate *= 1.5;
    }
  }
  
  return Math.min(rate, 1.0);
}

/**
 * Calculate scarcity factor for a position
 */
function calculatePositionScarcity(position, positionPlayers, picksUntilTarget) {
  if (!positionPlayers || positionPlayers.length === 0) return 0;
  
  // Calculate the drop-off in value
  const topTier = positionPlayers.slice(0, 3).reduce((sum, p) => sum + p.fpts, 0) / Math.min(3, positionPlayers.length);
  const nextTier = positionPlayers.slice(3, 8).reduce((sum, p) => sum + p.fpts, 0) / Math.min(5, positionPlayers.length - 3);
  
  const dropOff = topTier - nextTier;
  const playersRemaining = positionPlayers.length;
  
  // Higher scarcity if big drop-off and few players remaining
  const scarcity = (dropOff / topTier) * (10 / (playersRemaining + 10));
  
  return scarcity;
}

/**
 * Calculate the value gap for waiting on each position
 */
function calculatePositionGaps(playersByPosition, expectedValues, picksUntilTarget) {
  const gaps = {};
  
  Object.keys(playersByPosition).forEach(pos => {
    const players = playersByPosition[pos] || [];
    if (players.length === 0) {
      gaps[pos] = { gap: 0, recommendation: 'No players available' };
      return;
    }
    
    // Current best available
    const currentBest = players[0].fpts;
    const expectedValue = expectedValues[pos] || 0;
    const gap = currentBest - expectedValue;
    
    // Calculate tier breaks
    let tierBreak = 'Gradual decline';
    if (players.length >= 3) {
      const top3Avg = players.slice(0, 3).reduce((sum, p) => sum + p.fpts, 0) / 3;
      const next3Avg = players.slice(3, 6).reduce((sum, p) => sum + p.fpts, 0) / Math.min(3, players.length - 3);
      const tierDrop = ((top3Avg - next3Avg) / top3Avg) * 100;
      
      if (tierDrop > 15) {
        tierBreak = 'Major tier drop';
      } else if (tierDrop > 8) {
        tierBreak = 'Moderate tier drop';
      }
    }
    
    // Generate recommendation
    let recommendation;
    if (gap > 20) {
      recommendation = 'PRIORITY - Significant value loss if you wait';
    } else if (gap > 10) {
      recommendation = 'Consider drafting - Notable value loss expected';
    } else if (gap > 5) {
      recommendation = 'Can wait - Minimal value loss';
    } else {
      recommendation = 'Safe to wait - Plenty of value remaining';
    }
    
    gaps[pos] = {
      gap: gap,
      currentBest: currentBest,
      expectedAtPick: expectedValue,
      tierBreak: tierBreak,
      recommendation: recommendation,
      playersRemaining: players.length
    };
  });
  
  return gaps;
}

/**
 * Display VONA results in a new sheet
 */
function displayVONAResults(results, pickNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create or get VONA sheet
  let vonaSheet = ss.getSheetByName('VONA Analysis');
  if (!vonaSheet) {
    vonaSheet = ss.insertSheet('VONA Analysis');
  } else {
    vonaSheet.clear();
  }
  
  // Set up headers
  const headers = [
    ['VONA Analysis for Pick #' + pickNumber, '', '', '', '', '', ''],
    ['Generated: ' + new Date().toLocaleString(), '', '', '', '', '', ''],
    ['', '', '', '', '', '', ''],
    ['Top Players by VONA', '', '', '', '', '', ''],
    ['Rank', 'Name', 'Team', 'Pos', 'FPTS', 'Expected at Pick', 'VONA']
  ];
  
  vonaSheet.getRange(1, 1, headers.length, headers[0].length).setValues(headers);
  
  // Format headers
  vonaSheet.getRange(1, 1, 1, 7).merge().setFontSize(16).setFontWeight('bold');
  vonaSheet.getRange(2, 1, 1, 7).merge().setFontStyle('italic');
  vonaSheet.getRange(4, 1, 1, 7).merge().setFontSize(14).setFontWeight('bold');
  vonaSheet.getRange(5, 1, 1, 7).setFontWeight('bold').setBackground('#f0f0f0');
  
  // Add top 30 players by VONA
  const topPlayers = results.players.slice(0, 30);
  const playerData = topPlayers.map((player, index) => [
    index + 1,
    player.name,
    player.team || '',
    player.position,
    player.fpts.toFixed(1),
    player.expectedValue.toFixed(1),
    player.vona.toFixed(1)
  ]);
  
  if (playerData.length > 0) {
    vonaSheet.getRange(6, 1, playerData.length, 7).setValues(playerData);
  }
  
  // Add position analysis section
  const positionRow = 6 + Math.max(playerData.length, 1) + 2;
  const positionHeaders = [
    ['Position Analysis - Value of Waiting', '', '', '', '', '', ''],
    ['Position', 'Current Best', 'Expected at Pick', 'Value Gap', 'Tier Status', 'Players Left', 'Recommendation']
  ];
  
  vonaSheet.getRange(positionRow, 1, positionHeaders.length, positionHeaders[0].length).setValues(positionHeaders);
  vonaSheet.getRange(positionRow, 1, 1, 7).merge().setFontSize(14).setFontWeight('bold');
  vonaSheet.getRange(positionRow + 1, 1, 1, 7).setFontWeight('bold').setBackground('#f0f0f0');
  
  // Add position data
  const positions = ['QB', 'RB', 'WR', 'TE', 'DST', 'K'];
  const positionData = positions.map(pos => {
    const gap = results.positionGaps[pos];
    if (!gap) {
      return [pos, 'N/A', 'N/A', 'N/A', 'N/A', 0, 'No data'];
    }
    return [
      pos,
      gap.currentBest.toFixed(1),
      gap.expectedAtPick.toFixed(1),
      gap.gap.toFixed(1),
      gap.tierBreak,
      gap.playersRemaining,
      gap.recommendation
    ];
  });
  
  vonaSheet.getRange(positionRow + 2, 1, positionData.length, 7).setValues(positionData);
  
  // Color code recommendations
  const recommendationRange = vonaSheet.getRange(positionRow + 2, 7, positionData.length, 1);
  for (let i = 0; i < positionData.length; i++) {
    const cell = recommendationRange.getCell(i + 1, 1);
    const recommendation = positionData[i][6];
    
    if (recommendation.includes('PRIORITY')) {
      cell.setBackground('#ffcccc').setFontWeight('bold');
    } else if (recommendation.includes('Consider')) {
      cell.setBackground('#ffe6cc');
    } else if (recommendation.includes('Can wait')) {
      cell.setBackground('#ffffcc');
    } else if (recommendation.includes('Safe')) {
      cell.setBackground('#ccffcc');
    }
  }
  
  // Add key insights
  const insightsRow = positionRow + 2 + positions.length + 2;
  const insights = generateVONAInsights(results, pickNumber);
  
  vonaSheet.getRange(insightsRow, 1).setValue('Key Insights').setFontSize(14).setFontWeight('bold');
  vonaSheet.getRange(insightsRow + 1, 1, insights.length, 1).setValues(insights.map(i => [i]));
  
  // Format columns
  vonaSheet.setColumnWidth(1, 60);
  vonaSheet.setColumnWidth(2, 180);
  vonaSheet.setColumnWidth(3, 60);
  vonaSheet.setColumnWidth(4, 60);
  vonaSheet.setColumnWidth(5, 80);
  vonaSheet.setColumnWidth(6, 120);
  vonaSheet.setColumnWidth(7, 200);
  
  // Apply borders
  const lastRow = vonaSheet.getLastRow();
  vonaSheet.getRange(1, 1, lastRow, 7).setBorder(true, true, true, true, true, true);
  
  SpreadsheetApp.getUi().alert(
    'VONA analysis complete!\n' +
    'Check the "VONA Analysis" sheet for detailed results.\n' +
    'The analysis shows which positions you should prioritize based on expected value decline.'
  );
}

/**
 * Generate key insights from VONA analysis
 */
function generateVONAInsights(results, pickNumber) {
  const insights = [];
  
  // Find positions with biggest gaps
  const sortedGaps = Object.entries(results.positionGaps)
    .filter(([pos, gap]) => gap.gap !== undefined)
    .sort((a, b) => b[1].gap - a[1].gap);
  
  if (sortedGaps.length > 0) {
    const [topPos, topGap] = sortedGaps[0];
    insights.push(`• ${topPos} has the highest value gap (${topGap.gap.toFixed(1)} points) - consider prioritizing`);
  }
  
  // Find scarce positions
  sortedGaps.forEach(([pos, gap]) => {
    if (gap.playersRemaining < 10 && gap.gap > 5) {
      insights.push(`• ${pos} position is becoming scarce with only ${gap.playersRemaining} quality players remaining`);
    }
  });
  
  // Best overall VONA players
  const topVONA = results.players.slice(0, 3);
  if (topVONA.length > 0) {
    insights.push(`• Top VONA players: ${topVONA.map(p => `${p.name} (${p.vona.toFixed(1)})`).join(', ')}`);
  }
  
  // Positions safe to wait on
  const safePositions = sortedGaps
    .filter(([pos, gap]) => gap.gap < 5 && gap.playersRemaining > 15)
    .map(([pos]) => pos);
  
  if (safePositions.length > 0) {
    insights.push(`• Safe to wait on: ${safePositions.join(', ')} - minimal value loss expected`);
  }
  
  return insights;
}

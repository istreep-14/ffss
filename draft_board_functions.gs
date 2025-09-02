/**
 * Main function to set up the draft board system
 * Creates team roster sheet and team configuration sheet based on draft board setup
 */
function setupDraftBoard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const draftSheet = ss.getSheetByName('draft');
  
  if (!draftSheet) {
    throw new Error('Sheet named "draft" not found. Please ensure the draft board sheet exists.');
  }
  
  // Analyze the draft board to determine team count and roster size
  const draftData = analyzeDraftBoard(draftSheet);
  const { teamCount, rosterSize, totalPlayers } = draftData;
  
  // Create or update the team roster sheet
  const rosterSheet = createOrUpdateSheet(ss, 'Team Rosters');
  setupRosterSheet(rosterSheet, teamCount, rosterSize);
  
  // Create or update the team configuration sheet
  const configSheet = createOrUpdateSheet(ss, 'Team Config');
  setupTeamConfigSheet(configSheet, teamCount);
  
  // Apply formulas to populate players automatically
  applyRosterFormulas(rosterSheet, teamCount, rosterSize);
  
  SpreadsheetApp.getUi().alert(
    `Draft board setup complete!\n` +
    `Teams: ${teamCount}\n` +
    `Roster Size: ${rosterSize}\n` +
    `Total Players: ${totalPlayers}`
  );
}

/**
 * Calculates roster size from Input sheet based on starter positions
 */
function calculateRosterSizeFromInput() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Input');
  
  if (!inputSheet) {
    throw new Error('Input sheet not found! Please run "Setup Input Sheet" first.');
  }
  
  // Get the starter counts from B3:B10 (rows 3-10, column 2)
  let totalStarters = 0;
  for (let row = 3; row <= 10; row++) {
    const starterCount = inputSheet.getRange(row, 2).getValue();
    if (starterCount && !isNaN(starterCount)) {
      totalStarters += Number(starterCount);
    }
  }
  
  return totalStarters;
}

/**
 * Analyzes the draft board to determine team count and roster size
 */
function analyzeDraftBoard(draftSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const helperSheet = ss.getSheetByName('Helper');
  
  // Get values from Helper tab if it exists
  if (helperSheet) {
    const rosterSize = helperSheet.getRange('B1').getValue();
    const teamCount = helperSheet.getRange('B2').getValue();
    const totalPlayers = helperSheet.getRange('B3').getValue();
    
    // Validate the values
    if (rosterSize && teamCount && !isNaN(rosterSize) && !isNaN(teamCount)) {
      return { 
        teamCount: Number(teamCount), 
        rosterSize: Number(rosterSize), 
        totalPlayers: Number(totalPlayers) 
      };
    }
  }
  
  // Fall back to original logic if Helper tab doesn't exist or has invalid values
  // Find the last column with data (team count)
  let lastCol = 2; // Start from column B
  const maxCol = draftSheet.getMaxColumns();
  
  // Check for data in row 2 (first round picks)
  for (let col = 2; col <= maxCol; col++) {
    const cellValue = draftSheet.getRange(2, col).getValue();
    if (cellValue === '') {
      lastCol = col - 1;
      break;
    }
    if (col === maxCol) {
      lastCol = col;
    }
  }
  
  const teamCount = lastCol - 1; // Subtract 1 because we started from column B (2)
  
  // Get roster size from Input sheet
  let rosterSize;
  try {
    rosterSize = calculateRosterSizeFromInput();
  } catch (error) {
    // Fall back to analyzing draft board if Input sheet is not available
    let lastRow = 2;
    const maxRow = draftSheet.getMaxRows();
    
    // Check every 3rd row starting from row 2 (rounds are in rows 2, 5, 8, 11, etc.)
    for (let row = 2; row <= maxRow; row += 3) {
      let hasData = false;
      for (let col = 2; col <= lastCol; col++) {
        if (draftSheet.getRange(row, col).getValue() !== '') {
          hasData = true;
          break;
        }
      }
      if (!hasData) {
        break;
      }
      lastRow = row;
    }
    
    rosterSize = Math.floor((lastRow - 2) / 3) + 1;
  }
  
  const totalPlayers = teamCount * rosterSize;
  
  return { teamCount, rosterSize, totalPlayers };
}

/**
 * Creates or updates a sheet with the given name
 */
function createOrUpdateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  } else {
    // Clear existing content
    sheet.clear();
  }
  
  return sheet;
}

/**
 * Sets up the roster sheet with headers and formatting
 */
function setupRosterSheet(sheet, teamCount, rosterSize) {
  // Set headers
  sheet.getRange(1, 1).setValue('Pick #');
  sheet.getRange(1, 2).setValue('Round');
  sheet.getRange(1, 3).setValue('Team #');
  sheet.getRange(1, 4).setValue('Team Name');
  sheet.getRange(1, 5).setValue('Player');
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, 5);
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setBackground('#E0E0E0');
  
  // Generate pick data
  const pickData = generateSnakeDraftOrder(teamCount, rosterSize);
  
  // Populate pick numbers, rounds, and team numbers
  for (let i = 0; i < pickData.length; i++) {
    const row = i + 2;
    sheet.getRange(row, 1).setValue(i + 1); // Pick number
    sheet.getRange(row, 2).setValue(pickData[i].round); // Round
    sheet.getRange(row, 3).setValue(pickData[i].team); // Team number
  }
  
  // Add borders
  const dataRange = sheet.getRange(1, 1, pickData.length + 1, 5);
  dataRange.setBorder(true, true, true, true, true, true);
  
  // Auto-resize columns
  for (let col = 1; col <= 5; col++) {
    sheet.autoResizeColumn(col);
  }
}

/**
 * Sets up the team configuration sheet
 */
function setupTeamConfigSheet(sheet, teamCount) {
  // Set headers
  sheet.getRange(1, 1).setValue('Team #');
  sheet.getRange(1, 2).setValue('Team Name');
  sheet.getRange(1, 3).setValue('Controlled');
  sheet.getRange(1, 4).setValue('Notes');
  
  // Format headers
  const headerRange = sheet.getRange(1, 1, 1, 4);
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setBackground('#E0E0E0');
  
  // Add team numbers and default names
  for (let i = 1; i <= teamCount; i++) {
    const row = i + 1;
    sheet.getRange(row, 1).setValue(i);
    sheet.getRange(row, 2).setValue(`Team ${i}`);
    
    // Add checkbox for controlled teams
    const checkbox = sheet.getRange(row, 3);
    checkbox.insertCheckboxes();
    checkbox.setValue(false);
  }
  
  // Add borders
  const dataRange = sheet.getRange(1, 1, teamCount + 1, 4);
  dataRange.setBorder(true, true, true, true, true, true);
  
  // Auto-resize columns
  for (let col = 1; col <= 4; col++) {
    sheet.autoResizeColumn(col);
  }
}

/**
 * Generates snake draft order
 */
function generateSnakeDraftOrder(teamCount, rosterSize) {
  const pickData = [];
  
  for (let round = 1; round <= rosterSize; round++) {
    if (round % 2 === 1) {
      // Odd rounds: 1, 2, 3, ..., teamCount
      for (let team = 1; team <= teamCount; team++) {
        pickData.push({ round: round, team: team });
      }
    } else {
      // Even rounds: teamCount, teamCount-1, ..., 1
      for (let team = teamCount; team >= 1; team--) {
        pickData.push({ round: round, team: team });
      }
    }
  }
  
  return pickData;
}

/**
 * Applies formulas to automatically populate players from the draft sheet
 */
function applyRosterFormulas(rosterSheet, teamCount, rosterSize) {
  const totalPicks = teamCount * rosterSize;
  
  // Apply team name formula
  for (let row = 2; row <= totalPicks + 1; row++) {
    const teamNumCell = `C${row}`;
    const formula = `=IF(ISBLANK(${teamNumCell}),"",IFERROR(VLOOKUP(${teamNumCell},'Team Config'!A:B,2,FALSE),"Team " & ${teamNumCell}))`;
    rosterSheet.getRange(row, 4).setFormula(formula);
  }
  
  // Apply player name formula
  for (let row = 2; row <= totalPicks + 1; row++) {
    const pickData = getPickDataForRow(row - 1, teamCount, rosterSize);
    const draftRow = 2 + (pickData.round - 1) * 3; // Rounds are in rows 2, 5, 8, 11, etc.
    const draftCol = pickData.team + 1; // Add 1 because column A is 1, B is 2, etc.
    
    const formula = `=IFERROR(draft!${columnToLetter(draftCol)}${draftRow},"")`;
    rosterSheet.getRange(row, 5).setFormula(formula);
  }
}

/**
 * Gets pick data for a specific pick number
 */
function getPickDataForRow(pickNumber, teamCount, rosterSize) {
  const round = Math.ceil(pickNumber / teamCount);
  let teamInRound;
  
  if (round % 2 === 1) {
    // Odd round
    teamInRound = ((pickNumber - 1) % teamCount) + 1;
  } else {
    // Even round (snake back)
    teamInRound = teamCount - ((pickNumber - 1) % teamCount);
  }
  
  return { round: round, team: teamInRound };
}

/**
 * Converts column number to letter
 */
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// Note: onOpen function has been moved to code.gs to avoid conflicts

/**
 * Refreshes the roster sheet formulas
 */
function refreshRosters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const draftSheet = ss.getSheetByName('draft');
  const rosterSheet = ss.getSheetByName('Team Rosters');
  
  if (!draftSheet || !rosterSheet) {
    SpreadsheetApp.getUi().alert('Required sheets not found. Please run Setup Draft Board first.');
    return;
  }
  
  const draftData = analyzeDraftBoard(draftSheet);
  applyRosterFormulas(rosterSheet, draftData.teamCount, draftData.rosterSize);
  
  SpreadsheetApp.getUi().alert('Roster formulas refreshed!');
}

/**
 * Shows help information
 */
function showHelp() {
  const helpText = `
Draft Board Help:

1. IMPORTANT: Configure your league settings in the Helper tab first!
   - Set Roster Size in cell B1 (default: 16)
   - Set Teams in cell B2 (default: 12) 
   - Total Players is automatically calculated in B3

2. Set up your draft board in the 'draft' sheet:
   - Enter player names in columns B through Q (or less for fewer teams)
   - Use rows 2, 5, 8, 11, etc. for rounds (with helper text in between)

3. Run 'Setup Draft Board' from the Draft Board menu to:
   - Create 'Team Rosters' sheet with automatic player population
   - Create 'Team Config' sheet for team names and control settings
   - The number of teams and rounds will be based on your Helper tab settings

4. In 'Team Config' sheet:
   - Enter custom team names
   - Check boxes for teams you control

5. The 'Team Rosters' sheet will automatically:
   - Show pick order for snake draft
   - Display team names from config
   - Pull player names from the draft board

6. Use 'Test Roster Size Calculation' to verify the configuration.
7. Use 'Refresh Rosters' if you need to update formulas after changes.
  `;
  
  SpreadsheetApp.getUi().alert('Draft Board Help', helpText, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Test function to verify roster size calculation from Helper tab
 */
function testRosterSizeCalculation() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const helperSheet = ss.getSheetByName('Helper');
    
    if (helperSheet) {
      const rosterSize = helperSheet.getRange('B1').getValue();
      const teamCount = helperSheet.getRange('B2').getValue();
      const totalPlayers = helperSheet.getRange('B3').getValue();
      
      SpreadsheetApp.getUi().alert(
        'Draft Configuration Test',
        `Helper tab configuration:\n` +
        `- Roster Size: ${rosterSize} rounds\n` +
        `- Teams: ${teamCount}\n` +
        `- Total Players: ${totalPlayers}\n\n` +
        'These values will be used for the draft board setup.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      // Fall back to Input sheet calculation if Helper tab doesn't exist
      const rosterSize = calculateRosterSizeFromInput();
      SpreadsheetApp.getUi().alert(
        'Roster Size Test',
        `Calculated roster size from Input sheet (B3:B10): ${rosterSize} rounds\n\n` +
        'This is based on the total number of starter positions.\n' +
        'Consider setting up the Helper tab for easier configuration.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
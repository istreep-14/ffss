// Google Apps Script for Fantasy Draft Sheet Setup
// Copy this code into Tools > Script Editor in your Google Sheet

function setupDraftSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Configuration - Modify these values as needed
  const teamNames = ['Team 1', 'Team 2', 'Team 3', 'Team 4', 'Team 5', 'Team 6', 'Team 7', 'Team 8', 'Team 9', 'Team 10'];
  const positions = ['QB', 'RB', 'RB', 'WR', 'WR', 'TE', 'FLEX', 'SUPERFLEX', 'BN1', 'BN2', 'BN3', 'BN4', 'BN5', 'BN6'];
  
  // Create or get sheets
  let mainSheet = ss.getSheetByName('Draft Board');
  if (!mainSheet) {
    mainSheet = ss.insertSheet('Draft Board');
  }
  
  let draftPicksSheet = ss.getSheetByName('Draft Picks');
  if (!draftPicksSheet) {
    draftPicksSheet = ss.insertSheet('Draft Picks');
  }
  
  // Setup Main Draft Board
  setupMainSheet(mainSheet, teamNames, positions);
  
  // Setup Draft Picks Sheet
  setupDraftPicksSheet(draftPicksSheet);
  
  // Add formulas to main sheet
  addFormulasToMainSheet(mainSheet, teamNames.length, positions.length);
  
  // Format sheets
  formatSheets(mainSheet, draftPicksSheet);
  
  SpreadsheetApp.getUi().alert('Draft sheet setup complete!');
}

function setupMainSheet(sheet, teamNames, positions) {
  sheet.clear();
  
  // Set headers
  sheet.getRange(1, 1).setValue('Position');
  for (let i = 0; i < teamNames.length; i++) {
    sheet.getRange(1, i + 2).setValue(teamNames[i]);
  }
  
  // Set positions
  for (let i = 0; i < positions.length; i++) {
    sheet.getRange(i + 2, 1).setValue(positions[i]);
  }
  
  // Add summary section headers
  const summaryStartRow = positions.length + 4;
  sheet.getRange(summaryStartRow, 1).setValue('POSITION SUMMARY');
  sheet.getRange(summaryStartRow + 1, 1).setValue('Position');
  
  const uniquePositions = ['QB', 'RB', 'WR', 'TE', 'Total'];
  for (let i = 0; i < uniquePositions.length; i++) {
    sheet.getRange(summaryStartRow + 2 + i, 1).setValue(uniquePositions[i]);
  }
  
  // Copy team names to summary section
  for (let i = 0; i < teamNames.length; i++) {
    sheet.getRange(summaryStartRow + 1, i + 2).setValue(teamNames[i]);
  }
}

function setupDraftPicksSheet(sheet) {
  sheet.clear();
  
  // Set headers
  const headers = ['Pick #', 'Round', 'Team', 'Player', 'Position'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Add sample data
  const sampleData = [
    [1, 1, 'Team 1', 'Sample Player 1', 'RB'],
    [2, 1, 'Team 2', 'Sample Player 2', 'QB']
  ];
  sheet.getRange(2, 1, sampleData.length, 5).setValues(sampleData);
}

function addFormulasToMainSheet(sheet, numTeams, numPositions) {
  const starters = 9; // Positions before bench
  
  for (let col = 2; col <= numTeams + 1; col++) {
    const teamName = sheet.getRange(1, col).getValue();
    
    for (let row = 2; row <= numPositions + 1; row++) {
      const position = sheet.getRange(row, 1).getValue();
      let formula = '';
      
      if (row <= starters + 1) {
        // Starter positions formula
        formula = `=IFERROR(
          IF(LEFT($A${row},2)="BN","",
            IF($A${row}="FLEX",
              IF(COUNTIFS('Draft Picks'!$C:$C,$${String.fromCharCode(64 + col)}$1,'Draft Picks'!$E:$E,"RB")>COUNTIF($A$2:$A$${starters + 1},"RB"),
                INDEX('Draft Picks'!$D:$D,SMALL(IF(('Draft Picks'!$C:$C=$${String.fromCharCode(64 + col)}$1)*('Draft Picks'!$E:$E="RB"),ROW('Draft Picks'!$D:$D)),COUNTIF($A$2:$A$${starters + 1},"RB")+1)),
                IF(COUNTIFS('Draft Picks'!$C:$C,$${String.fromCharCode(64 + col)}$1,'Draft Picks'!$E:$E,"WR")>COUNTIF($A$2:$A$${starters + 1},"WR"),
                  INDEX('Draft Picks'!$D:$D,SMALL(IF(('Draft Picks'!$C:$C=$${String.fromCharCode(64 + col)}$1)*('Draft Picks'!$E:$E="WR"),ROW('Draft Picks'!$D:$D)),COUNTIF($A$2:$A$${starters + 1},"WR")+1)),
                  IF(COUNTIFS('Draft Picks'!$C:$C,$${String.fromCharCode(64 + col)}$1,'Draft Picks'!$E:$E,"TE")>COUNTIF($A$2:$A$${starters + 1},"TE"),
                    INDEX('Draft Picks'!$D:$D,SMALL(IF(('Draft Picks'!$C:$C=$${String.fromCharCode(64 + col)}$1)*('Draft Picks'!$E:$E="TE"),ROW('Draft Picks'!$D:$D)),COUNTIF($A$2:$A$${starters + 1},"TE")+1)),
                    ""))),
              IF($A${row}="SUPERFLEX",
                IF(COUNTIFS('Draft Picks'!$C:$C,$${String.fromCharCode(64 + col)}$1,'Draft Picks'!$E:$E,"QB")>COUNTIF($A$2:$A$${starters + 1},"QB"),
                  INDEX('Draft Picks'!$D:$D,SMALL(IF(('Draft Picks'!$C:$C=$${String.fromCharCode(64 + col)}$1)*('Draft Picks'!$E:$E="QB"),ROW('Draft Picks'!$D:$D)),COUNTIF($A$2:$A$${starters + 1},"QB")+1)),
                  IF(COUNTIFS('Draft Picks'!$C:$C,$${String.fromCharCode(64 + col)}$1,'Draft Picks'!$E:$E,"RB")>COUNTIF($A$2:$A$${starters + 1},"RB"),
                    INDEX('Draft Picks'!$D:$D,SMALL(IF(('Draft Picks'!$C:$C=$${String.fromCharCode(64 + col)}$1)*('Draft Picks'!$E:$E="RB"),ROW('Draft Picks'!$D:$D)),COUNTIF($A$2:$A$${starters + 1},"RB")+1)),
                    IF(COUNTIFS('Draft Picks'!$C:$C,$${String.fromCharCode(64 + col)}$1,'Draft Picks'!$E:$E,"WR")>COUNTIF($A$2:$A$${starters + 1},"WR"),
                      INDEX('Draft Picks'!$D:$D,SMALL(IF(('Draft Picks'!$C:$C=$${String.fromCharCode(64 + col)}$1)*('Draft Picks'!$E:$E="WR"),ROW('Draft Picks'!$D:$D)),COUNTIF($A$2:$A$${starters + 1},"WR")+1)),
                      IF(COUNTIFS('Draft Picks'!$C:$C,$${String.fromCharCode(64 + col)}$1,'Draft Picks'!$E:$E,"TE")>COUNTIF($A$2:$A$${starters + 1},"TE"),
                        INDEX('Draft Picks'!$D:$D,SMALL(IF(('Draft Picks'!$C:$C=$${String.fromCharCode(64 + col)}$1)*('Draft Picks'!$E:$E="TE"),ROW('Draft Picks'!$D:$D)),COUNTIF($A$2:$A$${starters + 1},"TE")+1)),
                        "")))),
                IF(COUNTIF($${String.fromCharCode(64 + col)}$2:$${String.fromCharCode(64 + col)}${row - 1},$A${row})<COUNTIF($A$2:$A$${starters + 1},$A${row}),
                  INDEX('Draft Picks'!$D:$D,SMALL(IF(('Draft Picks'!$C:$C=$${String.fromCharCode(64 + col)}$1)*('Draft Picks'!$E:$E=$A${row}),ROW('Draft Picks'!$D:$D)),COUNTIF($${String.fromCharCode(64 + col)}$2:$${String.fromCharCode(64 + col)}${row - 1},$A${row})+1)),
                  "")))),"")`;
      } else {
        // Bench positions formula
        formula = `=IFERROR(INDEX('Draft Picks'!$D:$D,SMALL(IF(('Draft Picks'!$C:$C=$${String.fromCharCode(64 + col)}$1)*(COUNTIF($${String.fromCharCode(64 + col)}$2:$${String.fromCharCode(64 + col)}$${starters + 1},'Draft Picks'!$D:$D)=0),ROW('Draft Picks'!$D:$D)),ROW()-${starters + 1})),"")`;
      }
      
      sheet.getRange(row, col).setFormula(formula);
    }
    
    // Add summary formulas
    const summaryStartRow = numPositions + 4;
    const summaryPositions = ['QB', 'RB', 'WR', 'TE'];
    
    for (let i = 0; i < summaryPositions.length; i++) {
      const posFormula = `=COUNTIFS('Draft Picks'!$C:$C,$${String.fromCharCode(64 + col)}$1,'Draft Picks'!$E:$E,"${summaryPositions[i]}")`;
      sheet.getRange(summaryStartRow + 2 + i, col).setFormula(posFormula);
    }
    
    // Total formula
    const totalFormula = `=COUNTIF('Draft Picks'!$C:$C,$${String.fromCharCode(64 + col)}$1)`;
    sheet.getRange(summaryStartRow + 6, col).setFormula(totalFormula);
  }
}

function formatSheets(mainSheet, draftPicksSheet) {
  // Format main sheet
  mainSheet.setFrozenRows(1);
  mainSheet.setFrozenColumns(1);
  
  // Header formatting
  mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn())
    .setBackground('#2196F3')
    .setFontColor('white')
    .setFontWeight('bold');
  
  // Position column formatting
  mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, 1)
    .setBackground('#E3F2FD')
    .setFontWeight('bold');
  
  // Conditional formatting for filled positions
  const numTeams = mainSheet.getLastColumn() - 1;
  const numPositions = mainSheet.getLastRow();
  
  for (let col = 2; col <= numTeams + 1; col++) {
    const range = mainSheet.getRange(2, col, numPositions - 1, 1);
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenCellNotEmpty()
      .setBackground('#C8E6C9')
      .setRanges([range])
      .build();
    mainSheet.addConditionalFormatRule(rule);
  }
  
  // Format draft picks sheet
  draftPicksSheet.setFrozenRows(1);
  draftPicksSheet.getRange(1, 1, 1, 5)
    .setBackground('#4CAF50')
    .setFontColor('white')
    .setFontWeight('bold');
  
  // Auto-resize columns
  mainSheet.autoResizeColumns(1, mainSheet.getLastColumn());
  draftPicksSheet.autoResizeColumns(1, 5);
}

// Menu creation
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Draft Tools')
    .addItem('Setup Draft Sheet', 'setupDraftSheet')
    .addItem('Clear Draft Data', 'clearDraftData')
    .addItem('Generate Mock Draft', 'generateMockDraft')
    .addSeparator()
    .addItem('About', 'showAbout')
    .addToUi();
}

function clearDraftData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Clear Draft Data',
    'This will clear all draft picks. Are you sure?',
    ui.ButtonSet.YES_NO
  );
  
  if (response == ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const draftPicksSheet = ss.getSheetByName('Draft Picks');
    if (draftPicksSheet) {
      draftPicksSheet.getRange(2, 1, draftPicksSheet.getLastRow() - 1, 5).clearContent();
    }
    ui.alert('Draft data cleared!');
  }
}

function generateMockDraft() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const draftPicksSheet = ss.getSheetByName('Draft Picks');
  const mainSheet = ss.getSheetByName('Draft Board');
  
  if (!draftPicksSheet || !mainSheet) {
    SpreadsheetApp.getUi().alert('Please run Setup Draft Sheet first!');
    return;
  }
  
  // Sample player pool
  const players = [
    {name: 'Patrick Mahomes', position: 'QB'},
    {name: 'Josh Allen', position: 'QB'},
    {name: 'Lamar Jackson', position: 'QB'},
    {name: 'Christian McCaffrey', position: 'RB'},
    {name: 'Austin Ekeler', position: 'RB'},
    {name: 'Derrick Henry', position: 'RB'},
    {name: 'Nick Chubb', position: 'RB'},
    {name: 'Jonathan Taylor', position: 'RB'},
    {name: 'Saquon Barkley', position: 'RB'},
    {name: 'Tony Pollard', position: 'RB'},
    {name: 'Justin Jefferson', position: 'WR'},
    {name: 'Ja\'Marr Chase', position: 'WR'},
    {name: 'Tyreek Hill', position: 'WR'},
    {name: 'Stefon Diggs', position: 'WR'},
    {name: 'CeeDee Lamb', position: 'WR'},
    {name: 'A.J. Brown', position: 'WR'},
    {name: 'Travis Kelce', position: 'TE'},
    {name: 'Mark Andrews', position: 'TE'},
    {name: 'T.J. Hockenson', position: 'TE'},
    {name: 'George Kittle', position: 'TE'}
  ];
  
  // Get team names
  const teamNames = [];
  for (let col = 2; col <= mainSheet.getLastColumn(); col++) {
    const teamName = mainSheet.getRange(1, col).getValue();
    if (teamName) teamNames.push(teamName);
  }
  
  // Generate mock draft (snake draft)
  const draftData = [];
  let pickNum = 1;
  const rounds = 3;
  
  for (let round = 1; round <= rounds; round++) {
    const order = round % 2 === 1 ? teamNames : teamNames.slice().reverse();
    
    for (let team of order) {
      if (players.length > 0) {
        const randomIndex = Math.floor(Math.random() * players.length);
        const player = players.splice(randomIndex, 1)[0];
        draftData.push([pickNum, round, team, player.name, player.position]);
        pickNum++;
      }
    }
  }
  
  // Clear existing data and add new
  draftPicksSheet.getRange(2, 1, draftPicksSheet.getLastRow() - 1, 5).clearContent();
  if (draftData.length > 0) {
    draftPicksSheet.getRange(2, 1, draftData.length, 5).setValues(draftData);
  }
  
  SpreadsheetApp.getUi().alert('Mock draft generated!');
}

function showAbout() {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2>Fantasy Football Draft Tracker</h2>
      <p>This tool helps you track your fantasy football draft with automatic roster filling.</p>
      <h3>Features:</h3>
      <ul>
        <li>Automatic position filling (starters → flex → bench)</li>
        <li>Real-time draft tracking</li>
        <li>Position counters</li>
        <li>Snake draft support</li>
      </ul>
      <h3>How to use:</h3>
      <ol>
        <li>Run "Setup Draft Sheet" from the Draft Tools menu</li>
        <li>Enter picks in the "Draft Picks" sheet as they happen</li>
        <li>Watch rosters fill automatically on the "Draft Board" sheet</li>
      </ol>
    </div>
  `)
  .setWidth(400)
  .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'About Draft Tracker');
}